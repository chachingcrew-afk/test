"""Find same-EAN competitor pages in the Netherlands for each KidsDeco product.

    SERPAPI_KEY=... python scripts/find_competitors.py KidsDeco_100_termek_versenytarsak.xlsx [--limit 5]

1. Asks Google Shopping (gl=nl) for each EAN through SerpApi.
2. Opens every merchant page and marks it verified only if the EAN is on the page.
3. Writes data/competitor_candidates.csv for review.

Nothing is added to the tracker automatically. After reviewing the CSV, add the
verified rows to config/products/kidsdeco.yaml (or re-run the import) so the
price tracker picks them up.
"""
import argparse
import csv
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
import requests

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tracker.fetcher import Fetcher, FetchError  # noqa: E402

SERP_URL = "https://serpapi.com/search.json"
OUT = ROOT / "data" / "competitor_candidates.csv"
OWN_SHEET = "100 kiválasztott termék"
FIELDS = ["kidsdeco_id", "ean", "shop", "title", "url", "price_text", "verified", "checked_at"]


def own_products(xlsx: Path) -> list[tuple[int, str]]:
    ws = openpyxl.load_workbook(xlsx, read_only=True)[OWN_SHEET]
    out = []
    for row in ws.iter_rows(min_row=11, values_only=True):
        if isinstance(row[0], int) and row[4]:
            out.append((row[0], str(row[4]).strip()))
    return out


def ean_variants(ean: str) -> set[str]:
    """The same code can be written as EAN-13, EAN-12 (UPC) or with leading zeros."""
    digits = re.sub(r"\D", "", ean)
    return {digits, digits.zfill(13), digits.lstrip("0")} - {""}


def ean_on_page(html: str, ean: str) -> bool:
    return any(re.search(rf"(?<!\d){re.escape(v)}(?!\d)", html) for v in ean_variants(ean))


def serp_shopping(ean: str, key: str) -> list[dict]:
    resp = requests.get(SERP_URL, params={
        "engine": "google_shopping", "q": ean, "gl": "nl", "hl": "nl", "api_key": key,
    }, timeout=60)
    resp.raise_for_status()
    return resp.json().get("shopping_results", [])


def verify(fetcher: Fetcher, url: str, ean: str) -> str:
    try:
        resp = fetcher.get(url)
    except FetchError as err:
        return "blokkolt" if err.blocked else "nem elérhető"
    return "igen" if ean_on_page(resp.text, ean) else "nem"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx", type=Path)
    parser.add_argument("--limit", type=int, default=None, help="only the first N products (test run)")
    args = parser.parse_args()

    key = os.environ.get("SERPAPI_KEY")
    if not key:
        sys.exit("SERPAPI_KEY is not set")

    proxy = os.environ.get("PROXY_URL") or None
    fetcher = Fetcher(proxy)
    products = own_products(args.xlsx)[: args.limit]
    checked_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M")
    rows, seen = [], set()

    for kd_id, ean in products:
        try:
            results = serp_shopping(ean, key)
        except requests.RequestException as err:
            print(f"{kd_id} {ean}: search failed ({err})")
            continue
        found = 0
        for item in results:
            url = item.get("product_link") or item.get("link")
            if not url or (kd_id, url) in seen:
                continue
            seen.add((kd_id, url))
            verified = verify(fetcher, url, ean)
            rows.append({
                "kidsdeco_id": kd_id, "ean": ean, "shop": item.get("source", ""),
                "title": item.get("title", ""), "url": url,
                "price_text": item.get("price", ""), "verified": verified, "checked_at": checked_at,
            })
            found += verified == "igen"
            time.sleep(1)
        print(f"{kd_id} {ean}: {len(results)} results, {found} verified")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(rows)
    verified = sum(r["verified"] == "igen" for r in rows)
    print(f"wrote {OUT.relative_to(ROOT)}: {len(rows)} candidates, {verified} verified")


if __name__ == "__main__":
    main()
