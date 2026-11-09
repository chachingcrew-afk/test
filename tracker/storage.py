"""Storage backends. The scraper only talks to the Storage interface, so switching
from CSV to Supabase is configuration, not code:

    STORAGE_BACKENDS=csv            (default, testing)
    STORAGE_BACKENDS=csv,supabase   (transition: write both)
    STORAGE_BACKENDS=supabase       (production)
"""
from __future__ import annotations

import csv
import os
from dataclasses import asdict, dataclass, fields
from decimal import Decimal
from pathlib import Path

import requests


@dataclass
class PriceRecord:
    checked_at: str            # ISO 8601 UTC
    run_id: str
    product_id: str
    product_name: str
    url: str
    status: str                # ok | failed | blocked
    price: Decimal | None = None
    currency: str | None = None
    in_stock: bool | None = None
    method: str | None = None  # which extractor found the price
    used_proxy: bool = False
    rendered: bool = False
    error: str | None = None   # attempts log when the price was not found


FIELDNAMES = [f.name for f in fields(PriceRecord)]


class Storage:
    def write(self, records: list[PriceRecord]) -> None:
        raise NotImplementedError


class CsvStorage(Storage):
    def __init__(self, path: str | Path = "data/prices.csv"):
        self.path = Path(path)

    def write(self, records: list[PriceRecord]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        is_new = not self.path.exists() or self.path.stat().st_size == 0
        with self.path.open("a", newline="", encoding="utf-8") as fh:
            writer = csv.DictWriter(fh, fieldnames=FIELDNAMES)
            if is_new:
                writer.writeheader()
            for r in records:
                row = asdict(r)
                row["price"] = str(r.price) if r.price is not None else ""
                writer.writerow({k: ("" if v is None else v) for k, v in row.items()})


class SupabaseStorage(Storage):
    """Inserts through the Supabase REST API (PostgREST). No extra SDK needed.

    Needs SUPABASE_URL and SUPABASE_SERVICE_KEY (service role key, keep it in GitHub Secrets).
    The (product_id, checked_at) unique key makes re-imports and retries safe.
    """

    def __init__(self, url: str | None = None, key: str | None = None, table: str = "price_checks"):
        self.url = (url or os.environ.get("SUPABASE_URL", "")).rstrip("/")
        self.key = key or os.environ.get("SUPABASE_SERVICE_KEY", "")
        self.table = table
        if not self.url or not self.key:
            raise RuntimeError("Supabase storage needs SUPABASE_URL and SUPABASE_SERVICE_KEY")

    def write(self, records: list[PriceRecord], batch_size: int = 500) -> None:
        rows = [to_db_row(r) for r in records]
        for i in range(0, len(rows), batch_size):
            resp = requests.post(
                f"{self.url}/rest/v1/{self.table}?on_conflict=product_id,checked_at",
                json=rows[i:i + batch_size],
                headers={
                    "apikey": self.key,
                    "Authorization": f"Bearer {self.key}",
                    "Content-Type": "application/json",
                    "Prefer": "resolution=ignore-duplicates,return=minimal",
                },
                timeout=30,
            )
            if resp.status_code >= 300:
                raise RuntimeError(f"Supabase insert failed: HTTP {resp.status_code} {resp.text[:300]}")


def to_db_row(r: PriceRecord) -> dict:
    row = asdict(r)
    row["price"] = float(r.price) if r.price is not None else None
    return row


def from_csv_row(row: dict) -> PriceRecord:
    def boolean(v):
        return None if v in ("", None) else str(v).lower() == "true"

    return PriceRecord(
        checked_at=row["checked_at"], run_id=row.get("run_id", ""), product_id=row["product_id"],
        product_name=row.get("product_name", ""), url=row["url"], status=row["status"],
        price=Decimal(row["price"]) if row.get("price") else None, currency=row.get("currency") or None,
        in_stock=boolean(row.get("in_stock")), method=row.get("method") or None,
        used_proxy=bool(boolean(row.get("used_proxy"))), rendered=bool(boolean(row.get("rendered"))),
        error=row.get("error") or None,
    )


class MultiStorage(Storage):
    def __init__(self, backends: list[Storage]):
        self.backends = backends

    def write(self, records: list[PriceRecord]) -> None:
        errors = []
        for b in self.backends:
            try:
                b.write(records)
            except Exception as err:  # one failing backend must not lose data in the other
                errors.append(f"{type(b).__name__}: {err}")
        if errors:
            raise RuntimeError("; ".join(errors))


def build_storage(spec: str | None = None, csv_path: str = "data/prices.csv") -> Storage:
    names = [n.strip().lower() for n in (spec or os.environ.get("STORAGE_BACKENDS") or "csv").split(",") if n.strip()]
    backends: list[Storage] = []
    for name in names:
        if name == "csv":
            backends.append(CsvStorage(csv_path))
        elif name == "supabase":
            backends.append(SupabaseStorage())
        else:
            raise ValueError(f"unknown storage backend: {name}")
    return backends[0] if len(backends) == 1 else MultiStorage(backends)
