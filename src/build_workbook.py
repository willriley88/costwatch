#!/usr/bin/env python3
"""
build_workbook.py — generate an input-cost review workbook from public data.

The deliverable a cost accountant would actually keep: one workbook, refreshed
on demand, showing what each purchased input class did over 3 and 12 months,
plus a bill-of-materials sheet that converts those index moves into a dollar
impact on THEIR product cost.

That last sheet is the point. Everybody can find a PPI index; nobody wants an
index. They want "your material cost is up $2.14 a unit and here is which input
did it."

Usage
-----
    python3 build_workbook.py                       # all catalog series
    python3 build_workbook.py --category metals energy
    python3 build_workbook.py --bom my_bom.csv --out review.xlsx

BOM csv format (header required):
    component,input_series,qty_per_unit,unit_cost_usd,baseline_date
    housing,WPU072,0.85,3.20,2025-09-01
"""

from __future__ import annotations

import argparse
import csv
import sys
from datetime import date, datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import series as S

NAVY = "FF1F2937"
ACCENT = "FFB45309"
LIGHT = "FFF3F4F6"
RED = "FFB91C1C"
GREEN = "FF15803D"


def _hdr(ws, row: int, labels: list[str], widths: list[int] | None = None) -> None:
    for i, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=text)
        c.font = Font(bold=True, color="FFFFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def sheet_summary(wb: Workbook, rows: list[dict]) -> None:
    ws = wb.create_sheet("Cost Summary")
    ws["A1"] = "Input cost review"
    ws["A1"].font = Font(bold=True, size=15, color=NAVY)
    ws["A2"] = f"Generated {datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC · public US series (BLS PPI/CPI, EIA via FRED) · no proprietary data"
    ws["A2"].font = Font(size=9, italic=True, color="FF6B7280")

    _hdr(ws, 4, ["Input", "Category", "Series", "As of", "Level", "Unit", "3-mo %", "12-mo %", "Obs"],
         [34, 13, 15, 12, 12, 22, 10, 10, 8])

    r = 5
    for row in sorted(rows, key=lambda x: (x["category"], x["label"])):
        ws.cell(row=r, column=1, value=row["label"])
        ws.cell(row=r, column=2, value=row["category"])
        ws.cell(row=r, column=3, value=row["sid"]).font = Font(name="Consolas", size=9)
        if row["error"]:
            c = ws.cell(row=r, column=4, value=f"FETCH FAILED: {row['error'][:60]}")
            c.font = Font(color=RED, size=9)
            r += 1
            continue
        ws.cell(row=r, column=4, value=row["as_of"])
        ws.cell(row=r, column=5, value=round(row["value"], 3)).number_format = "#,##0.000"
        ws.cell(row=r, column=6, value=row["unit"]).font = Font(size=9)
        for col, key in ((7, "chg_3m_pct"), (8, "chg_12m_pct")):
            v = row[key]
            c = ws.cell(row=r, column=col, value=(round(v, 2) if v is not None else "n/a"))
            c.number_format = "+0.00%;-0.00%"
            if v is not None:
                c.value = v / 100.0
                c.font = Font(color=RED if v > 0 else GREEN, bold=abs(v) >= 5)
        ws.cell(row=r, column=9, value=row["obs_count"])
        if r % 2 == 0:
            for col in range(1, 10):
                ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=LIGHT)
        r += 1

    ws.cell(row=r + 1, column=1,
            value="Positive = input cost rose. Indices are not dollars; use the BOM Impact sheet to convert.").font = \
        Font(italic=True, size=9, color="FF6B7280")


def sheet_history(wb: Workbook, sids: list[str], months: int = 60) -> None:
    """Aligned monthly history + a chart. This is what gets pasted into a deck."""
    ws = wb.create_sheet("History")

    data: dict[str, dict[str, float]] = {}
    for sid in sids:
        try:
            obs = S.load(sid)
        except S.FetchError:
            continue
        monthly: dict[str, float] = {}
        for d, v in obs:
            monthly[d[:7]] = v  # last observation in each month wins
        data[sid] = monthly

    if not data:
        ws["A1"] = "no series fetched"
        return

    keys = sorted({k for m in data.values() for k in m})[-months:]
    sids_ok = [s for s in sids if s in data]
    _hdr(ws, 1, ["Month"] + [S.CATALOG[s].label if s in S.CATALOG else s for s in sids_ok],
         [10] + [18] * len(sids_ok))

    for i, mk in enumerate(keys, start=2):
        ws.cell(row=i, column=1, value=mk)
        for j, sid in enumerate(sids_ok, start=2):
            v = data[sid].get(mk)
            if v is not None:
                ws.cell(row=i, column=j, value=round(v, 3)).number_format = "#,##0.0"

    chart = LineChart()
    chart.title = "Input cost history (rebased view: read the trend, not the level)"
    chart.height, chart.width = 10, 26
    chart.y_axis.title = "index / price"
    refs = Reference(ws, min_col=2, max_col=1 + len(sids_ok), min_row=1, max_row=1 + len(keys))
    chart.add_data(refs, titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=1 + len(keys)))
    ws.add_chart(chart, f"{get_column_letter(len(sids_ok) + 3)}2")


def sheet_bom(wb: Workbook, bom: list[dict], rows: list[dict]) -> None:
    """Translate index moves into dollars per unit — the part that sells."""
    ws = wb.create_sheet("BOM Impact")
    ws["A1"] = "Bill-of-materials cost impact"
    ws["A1"].font = Font(bold=True, size=15, color=NAVY)
    ws["A2"] = "Applies each input's published index change to your own baseline unit cost."
    ws["A2"].font = Font(size=9, italic=True, color="FF6B7280")

    by_sid = {r["sid"]: r for r in rows}
    _hdr(ws, 4, ["Component", "Input series", "Qty/unit", "Baseline $/unit-of-input",
                 "Baseline cost/unit", "Index 12-mo %", "Implied $/unit now", "Delta $/unit"],
         [22, 28, 10, 22, 16, 13, 18, 14])

    r = 5
    first_data_row = r
    for item in bom:
        sid = item["input_series"]
        meta = S.CATALOG.get(sid)
        chg = (by_sid.get(sid) or {}).get("chg_12m_pct")
        qty = float(item["qty_per_unit"])
        unit_cost = float(item["unit_cost_usd"])
        base = qty * unit_cost

        ws.cell(row=r, column=1, value=item["component"])
        ws.cell(row=r, column=2, value=f"{meta.label} ({sid})" if meta else sid).font = Font(size=9)
        ws.cell(row=r, column=3, value=qty).number_format = "#,##0.000"
        ws.cell(row=r, column=4, value=unit_cost).number_format = '"$"#,##0.00'
        ws.cell(row=r, column=5, value=base).number_format = '"$"#,##0.00'
        if chg is None:
            ws.cell(row=r, column=6, value="n/a").font = Font(color=RED, size=9)
        else:
            c = ws.cell(row=r, column=6, value=chg / 100.0)
            c.number_format = "+0.0%;-0.0%"
            c.font = Font(color=RED if chg > 0 else GREEN)
            now = base * (1 + chg / 100.0)
            ws.cell(row=r, column=7, value=now).number_format = '"$"#,##0.00'
            d = ws.cell(row=r, column=8, value=now - base)
            d.number_format = '"+$"#,##0.00;"-$"#,##0.00'
            d.font = Font(bold=True, color=RED if now > base else GREEN)
        r += 1

    if r > first_data_row:
        ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
        for col in (5, 7, 8):
            L = get_column_letter(col)
            c = ws.cell(row=r, column=col, value=f"=SUM({L}{first_data_row}:{L}{r - 1})")
            c.font = Font(bold=True)
            c.number_format = '"$"#,##0.00'
            c.fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=r + 2, column=1,
                value="Assumes your input price tracks its published index 1:1. Contracted or hedged "
                      "inputs will move less; treat this as an exposure estimate, not an invoice.").font = \
            Font(italic=True, size=9, color="FF6B7280")


def sheet_sources(wb: Workbook, rows: list[dict]) -> None:
    ws = wb.create_sheet("Sources")
    ws["A1"] = "Provenance"
    ws["A1"].font = Font(bold=True, size=15, color=NAVY)
    _hdr(ws, 3, ["Series", "Input", "Publisher", "Access", "URL"], [15, 32, 26, 20, 62])
    r = 4
    for row in sorted(rows, key=lambda x: x["sid"]):
        pub = "BLS (via FRED)" if row["sid"].startswith(("WPU", "WPS", "PPIACO", "CU", "APU")) else "EIA/market (via FRED)"
        ws.cell(row=r, column=1, value=row["sid"]).font = Font(name="Consolas", size=9)
        ws.cell(row=r, column=2, value=row["label"])
        ws.cell(row=r, column=3, value=pub)
        ws.cell(row=r, column=4, value="public, no API key")
        ws.cell(row=r, column=5, value=f"https://fred.stlouisfed.org/series/{row['sid']}").font = Font(size=9, color="FF1D4ED8")
        r += 1
    ws.cell(row=r + 1, column=1, value=(
        "All series are US federal statistical output redistributed by FRED. Verified keyless "
        "2026-09-27. EIA API v2, USDA NASS QuickStats and USDA AMS all require registration and "
        "are deliberately not used."
    )).font = Font(italic=True, size=9, color="FF6B7280")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--category", nargs="*", help="limit to these categories")
    ap.add_argument("--bom", help="bill-of-materials CSV")
    ap.add_argument("--out", default="input-cost-review.xlsx")
    ap.add_argument("--refresh", action="store_true", help="bypass the 12h cache")
    ap.add_argument("--months", type=int, default=60, help="history months (default 60)")
    args = ap.parse_args()

    sids = list(S.CATALOG)
    if args.category:
        want = {c.lower() for c in args.category}
        sids = [s for s in sids if S.CATALOG[s].category.lower() in want]
        if not sids:
            raise SystemExit(f"no series in categories {sorted(want)}; "
                             f"available: {sorted({v.category for v in S.CATALOG.values()})}")

    print(f"fetching {len(sids)} series...")
    rows = S.series_table(sids, refresh=args.refresh)
    ok = [r for r in rows if not r["error"]]
    print(f"  {len(ok)}/{len(rows)} ok")
    for r in rows:
        if r["error"]:
            print(f"  FAILED {r['sid']}: {r['error']}")

    bom = []
    if args.bom:
        with open(args.bom, newline="") as fh:
            bom = list(csv.DictReader(fh))
        missing = [b["input_series"] for b in bom if b["input_series"] not in S.CATALOG]
        if missing:
            print(f"  note: BOM references series outside the catalog: {missing}")

    wb = Workbook()
    default = wb.active
    if default is not None:
        wb.remove(default)
    sheet_summary(wb, rows)
    sheet_history(wb, [r["sid"] for r in ok], months=args.months)
    if bom:
        sheet_bom(wb, bom, rows)
    sheet_sources(wb, rows)
    wb.save(args.out)

    print(f"\nwrote {args.out}")
    print(f"  sheets: {', '.join(wb.sheetnames)}")
    movers = sorted((r for r in ok if r["chg_12m_pct"] is not None),
                    key=lambda r: abs(r["chg_12m_pct"]), reverse=True)[:5]
    print("\nlargest 12-month moves:")
    for m in movers:
        print(f"  {m['chg_12m_pct']:+7.2f}%  {m['label']} ({m['sid']}, as of {m['as_of']})")


if __name__ == "__main__":
    main()
