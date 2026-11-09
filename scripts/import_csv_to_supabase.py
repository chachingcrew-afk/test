"""One-off: copy the CSV history into Supabase when you switch over.

    SUPABASE_URL=... SUPABASE_SERVICE_KEY=... python scripts/import_csv_to_supabase.py [data/prices.csv]

Safe to run more than once: duplicates (same product_id + checked_at) are skipped.
"""
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tracker.storage import SupabaseStorage, from_csv_row  # noqa: E402


def main() -> None:
    path = Path(sys.argv[1] if len(sys.argv) > 1 else "data/prices.csv")
    with path.open(encoding="utf-8") as fh:
        records = [from_csv_row(row) for row in csv.DictReader(fh)]
    SupabaseStorage().write(records)
    print(f"imported {len(records)} rows from {path}")


if __name__ == "__main__":
    main()
