"""Entry point.

    python -m tracker.main run                 # check every product due in this run
    python -m tracker.main run --only my-id    # one product, ignores the schedule
    python -m tracker.main run --dry-run       # print results, save nothing
    python -m tracker.main plan                # (GitHub Actions) is anything due? is a browser needed?
"""
from __future__ import annotations

import argparse
import logging
import os
import random
import sys
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path

import yaml

from .extractors import Context, extract_price
from .fetcher import BrowserRenderer, Fetcher
from .storage import PriceRecord, build_storage

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "products.yaml"
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "track-prices.yml"

DEFAULT_SETTINGS = {
    "runs_per_day": 1,
    "proxy_mode": "auto",
    "render_fallback": True,
    "default_currency": None,
    "delay_seconds": [2, 5],
    "timeout_seconds": 30,
}

log = logging.getLogger("tracker")


# ------------------------------------------------------------------ config
def load_config(path: Path = CONFIG_PATH) -> tuple[dict, list[dict]]:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    settings = {**DEFAULT_SETTINGS, **(data.get("settings") or {})}
    products = [p for p in (data.get("products") or []) if p.get("enabled", True)]
    seen = set()
    for p in products:
        if not p.get("id") or not p.get("url"):
            raise ValueError(f"every product needs an id and a url: {p}")
        if p["id"] in seen:
            raise ValueError(f"duplicate product id: {p['id']}")
        seen.add(p["id"])
        p.setdefault("proxy_mode", settings["proxy_mode"])
        p.setdefault("render", settings["render_fallback"])
    return settings, products


# ---------------------------------------------------------------- schedule
def workflow_slots(path: Path = WORKFLOW_PATH) -> list[str]:
    """The cron lines of the workflow, in file order. Each one is a 'slot'."""
    if not path.exists():
        return []
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    on = data.get("on", data.get(True)) or {}  # PyYAML reads the key `on` as True
    return [s["cron"] for s in on.get("schedule", []) if "cron" in s]


def active_slots(n_slots: int, runs_per_day: int) -> set[int]:
    """Spread `runs_per_day` runs evenly over the available slots.
    3 slots: 1/day -> {0}, 2/day -> {0, 2}, 3/day -> {0, 1, 2}"""
    if n_slots <= 0:
        return set()
    runs = max(1, min(int(runs_per_day), n_slots))
    if runs == 1:
        return {0}
    return {round(i * (n_slots - 1) / (runs - 1)) for i in range(runs)}


def due_products(settings: dict, products: list[dict], schedule: str | None) -> list[dict]:
    """Manual / local runs (no schedule) check everything."""
    slots = workflow_slots()
    if not schedule or schedule not in slots:
        return products
    slot = slots.index(schedule)
    return [p for p in products
            if slot in active_slots(len(slots), p.get("runs_per_day", settings["runs_per_day"]))]


# --------------------------------------------------------------------- run
def check_product(product: dict, fetcher: Fetcher, renderer: BrowserRenderer | None,
                  settings: dict, run_id: str) -> PriceRecord:
    ctx = Context(url=product["url"], product=product, fetcher=fetcher)
    result = extract_price(ctx, renderer)
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    base = dict(checked_at=now, run_id=run_id, product_id=product["id"],
                product_name=product.get("name") or (result.title if result else "") or "",
                url=product["url"], used_proxy=ctx.used_proxy, rendered=ctx.rendered)
    if result:
        currency = result.currency or product.get("currency") or settings.get("default_currency")
        return PriceRecord(**base, status="ok", price=result.price.quantize(type(result.price)("0.01")),
                           currency=currency, in_stock=result.in_stock, method=result.method)
    status = "not_found" if ctx.not_found else "blocked" if ctx.blocked else "failed"
    return PriceRecord(**base, status=status,
                       error=" | ".join(ctx.attempts)[:2000])


def write_summary(records: list[PriceRecord], proxy_bytes: int) -> None:
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    lines = ["| Product | Status | Price | Method | Proxy |", "|---|---|---|---|---|"]
    for r in records:
        price = f"{r.price} {r.currency or ''}" if r.price is not None else "-"
        lines.append(f"| {r.product_name or r.product_id} | {r.status} | {price} | {r.method or '-'} | "
                     f"{'yes' if r.used_proxy else 'no'} |")
    lines.append(f"\nProxy traffic this run (HTTP only): {proxy_bytes / 1024:.0f} KB")
    text = "\n".join(lines)
    if path:
        with open(path, "a", encoding="utf-8") as fh:
            fh.write(text + "\n")
    else:
        print(text)


def cmd_run(args) -> int:
    settings, products = load_config()
    if args.only:
        products = [p for p in products if p["id"] == args.only]
        if not products:
            log.error("no enabled product with id %s", args.only)
            return 2
    else:
        products = due_products(settings, products, os.environ.get("SCHEDULE"))
    if not products:
        log.info("nothing due in this run")
        return 0

    proxy_url = os.environ.get("PROXY_URL")
    if not proxy_url:
        log.warning("PROXY_URL is not set, blocked sites will fail")
    fetcher = Fetcher(proxy_url, timeout=settings["timeout_seconds"])
    renderer = BrowserRenderer(proxy_url) if settings["render_fallback"] else None
    run_id = uuid.uuid4().hex[:12]
    records: list[PriceRecord] = []

    try:
        for i, product in enumerate(products):
            if i:
                time.sleep(random.uniform(*settings["delay_seconds"]))
            record = check_product(product, fetcher, renderer, settings, run_id)
            records.append(record)
            if record.status == "ok":
                log.info("%-30s %s %s via %s", product["id"], record.price, record.currency or "", record.method)
            else:
                log.warning("%-30s %s: %s", product["id"], record.status, record.error)
    finally:
        if renderer:
            renderer.close()

    write_summary(records, fetcher.bytes_via_proxy)
    if not args.dry_run:
        build_storage().write(records)

    ok = sum(r.status == "ok" for r in records)
    log.info("done: %d/%d prices found", ok, len(records))
    return 0 if ok else 1


def cmd_plan(_args) -> int:
    settings, products = load_config()
    due = due_products(settings, products, os.environ.get("SCHEDULE"))
    out = {
        "should_run": str(bool(due)).lower(),
        "needs_browser": str(bool(settings["render_fallback"]) and any(p.get("render") for p in due)).lower(),
        "due_count": str(len(due)),
    }
    target = os.environ.get("GITHUB_OUTPUT")
    lines = "\n".join(f"{k}={v}" for k, v in out.items())
    if target:
        with open(target, "a", encoding="utf-8") as fh:
            fh.write(lines + "\n")
    print(lines)
    return 0


def load_dotenv(path: Path = ROOT / ".env") -> None:
    """Local runs: read .env without overriding real environment variables."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            key, value = line.split("=", 1)
            if value.strip():
                os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def main(argv=None) -> int:
    load_dotenv()
    logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO"),
                        format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    parser = argparse.ArgumentParser(prog="tracker")
    sub = parser.add_subparsers(dest="cmd", required=True)
    run = sub.add_parser("run")
    run.add_argument("--only", help="check a single product id (ignores the schedule)")
    run.add_argument("--dry-run", action="store_true", help="do not save results")
    sub.add_parser("plan")
    args = parser.parse_args(argv)
    return {"run": cmd_run, "plan": cmd_plan}[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
