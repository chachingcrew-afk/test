"""Write the latest tracked prices back into the KidsDeco competitor Excel.

    python scripts/export_to_excel.py INPUT.xlsx data/prices.csv OUTPUT.xlsx

Adds price columns to both sheets (keeping the existing tables, styles and formulas):
- Versenytárstalálatok: competitor price, currency, stock, fetch status, method,
  fetch time, KidsDeco price and the difference in %
- 100 kiválasztott termék: KidsDeco price and stock, number of priced EUR competitors,
  cheapest / average competitor price, cheapest same-EAN price, KidsDeco vs cheapest
Comparisons only use EUR prices; other currencies are shown but not compared.
"""
import csv
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.worksheet.table import TableColumn

sys.path.insert(0, str(Path(__file__).resolve().parent))
from import_excel import COMPETITOR_SHEET, OWN_SHEET, cell_url, header_index, slug  # noqa: E402

STATUS_HU = {"ok": "OK", "blocked": "Blokkolva", "not_found": "Nem található (404)", "failed": "Nem találta az árat"}
OWN = f"'{OWN_SHEET}'"
COMP = f"'{COMPETITOR_SHEET}'"


def latest_by_product(csv_path: Path) -> dict[str, dict]:
    latest: dict[str, dict] = {}
    with csv_path.open(encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            prev = latest.get(row["product_id"])
            if prev is None or row["checked_at"] >= prev["checked_at"]:
                latest[row["product_id"]] = row
    return latest


def add_columns(ws, header_row: int, last_row: int, specs: list[tuple[str, float]]) -> int:
    """Append header cells styled like the existing ones, extend the table. Returns first new column index."""
    first = ws.max_column + 1
    template_header = ws.cell(header_row, ws.max_column)
    template_body = ws.cell(header_row + 1, 2)  # a plain text column
    for offset, (name, width) in enumerate(specs):
        col = first + offset
        h = ws.cell(header_row, col, name)
        for attr in ("font", "fill", "alignment", "border"):
            setattr(h, attr, copy(getattr(template_header, attr)))
        for r in range(header_row + 1, last_row + 1):
            c = ws.cell(r, col)
            for attr in ("font", "alignment", "border"):
                setattr(c, attr, copy(getattr(template_body, attr)))
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    table = next(iter(ws.tables.values()))
    start = table.ref.split(":")[0]
    table.ref = f"{start}:{openpyxl.utils.get_column_letter(first + len(specs) - 1)}{last_row}"
    next_id = max(c.id for c in table.tableColumns) + 1
    for offset, (name, _) in enumerate(specs):
        table.tableColumns.append(TableColumn(id=next_id + offset, name=name))
    return first


def to_number(value: str):
    try:
        return float(value) if value not in ("", None) else None
    except ValueError:
        return None


def stock_hu(value: str) -> str:
    return {"True": "Igen", "False": "Nem"}.get(value, "")


def main() -> None:
    src, prices_csv, out = Path(sys.argv[1]), Path(sys.argv[2]), Path(sys.argv[3])
    latest = latest_by_product(prices_csv)
    wb = openpyxl.load_workbook(src)
    L = openpyxl.utils.get_column_letter

    # ---------------------------------------------------------- competitors
    ws = wb[COMPETITOR_SHEET]
    hdr, col = header_index(ws, None)
    last = hdr
    for r in range(hdr + 1, ws.max_row + 1):
        if ws.cell(r, 1).value is not None:
            last = r
    c0 = add_columns(ws, hdr, last, [
        ("Konkurens ár", 14), ("Pénznem", 11), ("Készleten", 11), ("Ár lekérés státusza", 22),
        ("Ár forrása", 20), ("Ár lekérés ideje", 18), ("KidsDeco ár (EUR)", 16), ("Eltérés a KidsDeco ártól", 18),
    ])
    cp, cc, cs, cst, cm, ct, ckd, cdiff = (L(c0 + i) for i in range(8))
    comp_first, comp_last = hdr + 1, last

    # own sheet columns are needed by the lookup formula; they are added below at fixed positions
    own_ws = wb[OWN_SHEET]
    own_hdr, own_col = header_index(own_ws, None)
    own_last = max(r for r in range(own_hdr + 1, own_ws.max_row + 1) if own_ws.cell(r, 1).value is not None)
    own_price_col = L(own_ws.max_column + 1)

    for r in range(comp_first, comp_last + 1):
        row = ws[r]
        kd_id, url = row[col["ID"]].value, cell_url(row[col["Konkurens URL"]])
        if url:
            shop = row[col["Konkurens webshop"]].value or ""
            n = row[col["Találat sorszáma"]].value
            rec = latest.get(f"kd-{kd_id:03d}-c{int(n):02d}-{slug(shop)}")
            if rec:
                ws[f"{cp}{r}"] = to_number(rec["price"])
                ws[f"{cc}{r}"] = rec["currency"] or None
                ws[f"{cs}{r}"] = stock_hu(rec["in_stock"]) or None
                ws[f"{cst}{r}"] = STATUS_HU.get(rec["status"], rec["status"])
                ws[f"{cm}{r}"] = rec["method"] or None
                ws[f"{ct}{r}"] = datetime.fromisoformat(rec["checked_at"]).replace(tzinfo=None)
        url_col = L(col["Konkurens URL"] + 1)
        ws[f"{ckd}{r}"] = (f"=IF({url_col}{r}=\"\",\"\",IFERROR(INDEX({OWN}!${own_price_col}${own_hdr + 1}:${own_price_col}${own_last},"
                           f"MATCH(A{r},{OWN}!$A${own_hdr + 1}:$A${own_last},0)),\"\"))")
        ws[f"{cdiff}{r}"] = (f"=IF(AND(ISNUMBER({cp}{r}),ISNUMBER({ckd}{r}),{cc}{r}=\"EUR\"),"
                             f"IF({ckd}{r}>0,{cp}{r}/{ckd}{r}-1,\"\"),\"\")")
        ws[f"{cp}{r}"].number_format = "#,##0.00"
        ws[f"{ckd}{r}"].number_format = "#,##0.00"
        ws[f"{ct}{r}"].number_format = "yyyy-mm-dd hh:mm"
        ws[f"{cdiff}{r}"].number_format = "+0.0%;-0.0%;0.0%"

    # ----------------------------------------------------------------- own
    ws = own_ws
    o0 = add_columns(ws, own_hdr, own_last, [
        ("KidsDeco ár (EUR)", 16), ("KidsDeco készlet", 14), ("Árazott EUR konkurensek", 16),
        ("Legolcsóbb konkurens (EUR)", 18), ("Konkurens átlagár (EUR)", 18),
        ("Legolcsóbb azonos EAN (EUR)", 18), ("KidsDeco vs legolcsóbb", 16),
    ])
    op, ostock, ocnt, omin, oavg, oean, odiff = (L(o0 + i) for i in range(7))
    assert op == own_price_col
    rng = lambda c: f"{COMP}!${c}${comp_first}:${c}${comp_last}"  # noqa: E731
    match_col = L(col["Egyezés minősége"] + 1)
    for r in range(own_hdr + 1, own_last + 1):
        kd_id = ws.cell(r, 1).value
        rec = latest.get(f"kd-{kd_id:03d}-kidsdeco") if isinstance(kd_id, int) else None
        if rec and rec["status"] == "ok":
            ws[f"{op}{r}"] = to_number(rec["price"])
            ws[f"{ostock}{r}"] = stock_hu(rec["in_stock"]) or None
        elif rec:
            ws[f"{ostock}{r}"] = STATUS_HU.get(rec["status"], rec["status"])
        cond = f"{rng('A')},$A{r},{rng(cc)},\"EUR\",{rng(cp)},\">0\""
        ws[f"{ocnt}{r}"] = f"=COUNTIFS({cond})"
        ws[f"{omin}{r}"] = f"=IF({ocnt}{r}>0,_xlfn.MINIFS({rng(cp)},{cond}),\"\")"
        ws[f"{oavg}{r}"] = f"=IF({ocnt}{r}>0,AVERAGEIFS({rng(cp)},{cond}),\"\")"
        ean = f"{cond},{rng(match_col)},\"Azonos EAN\""
        ws[f"{oean}{r}"] = f"=IF(COUNTIFS({ean})>0,_xlfn.MINIFS({rng(cp)},{ean}),\"\")"
        ws[f"{odiff}{r}"] = (f"=IF(AND(ISNUMBER({op}{r}),ISNUMBER({omin}{r})),"
                             f"IF({omin}{r}>0,{op}{r}/{omin}{r}-1,\"\"),\"\")")
        for c in (op, omin, oavg, oean):
            ws[f"{c}{r}"].number_format = "#,##0.00"
        ws[f"{odiff}{r}"].number_format = "+0.0%;-0.0%;0.0%"

    # notes so the reader knows where the numbers come from
    run_time = max((v["checked_at"] for v in latest.values()), default="")[:16].replace("T", " ")
    note = (f"Árak: automatikus lekérés a price-tracker programmal, {run_time} UTC. "
            "Az összehasonlítás csak EUR árakat használ; a 'Hasonló funkció / változat' találatok ára eltérő terméket "
            "jelenthet, ezért az 'azonos EAN' oszlop a legmegbízhatóbb.")
    for sheet, row in ((wb[OWN_SHEET], own_hdr - 1), (wb[COMPETITOR_SHEET], hdr - 1)):
        cell = sheet.cell(row, 1, note)
        cell.font = copy(sheet.cell(row - 1 if sheet.cell(row - 1, 1).value else 3, 1).font)

    wb.save(out)
    print(f"saved {out}")


if __name__ == "__main__":
    main()
