"""
Hotel P&L Builder
=================

Reads a folder of Operating Statement .xls exports and builds one formatted
Excel P&L workbook per hotel, showing Actual vs Budget vs Projected for every
year you have a statement for, with large variances highlighted.

    python hotel_pl_tool.py --input "C:\\path\\to\\statements" --output "C:\\path\\out"

One source file = one hotel, one year. The year and hotel name are read out of
the report header, so files can be named anything.

Notes on the source format (fixed layout, column J holds the line-item label):
    A  PTD actual        G  PTD budget
    K  YTD actual        Q  YTD budget        N  YTD last year
The YTD-last-year column is unreliable in these exports (it repeats the PTD
value), so history is built from separate annual files instead.
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import OrderedDict, defaultdict
from pathlib import Path

import xlrd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as gcl

# --------------------------------------------------------------------------
# configuration
# --------------------------------------------------------------------------
VAR_PCT = 0.10                  # flag when |actual-budget| / budget exceeds this
VAR_MIN = 5000                  # ...and the dollar swing exceeds this

COL_LABEL = 9                   # J
COL_PTD_ACT, COL_PTD_BUD = 0, 6         # A, G
COL_YTD_ACT, COL_YTD_BUD, COL_YTD_LY = 10, 16, 13   # K, Q, N

# page title (as it appears in the statement)  ->  output tab
TAB_MAP = OrderedDict([
    ("Rooms Department",      "Rooms"),
    ("Food Department",       "Food"),
    ("Beverage Department",   "Beverage"),
    ("Spa",                   "Miscellaneous"),
    ("Miscellaneous Revenue", "Miscellaneous"),
    ("Rental Income",         "Miscellaneous"),
])

# lines pulled onto the Fixed Expenses tab (matched case-insensitively)
FIXED_LINES = ["Real Estates Taxes", "Real Estate Taxes", "Insurance"]

# Summary tab, grouped. Every line comes off the statement's own Summary page.
SUMMARY_GROUPS = [
    ("REVENUE", ["Room", "Food & Beverage", "Spa Revenue", "Miscellaneous",
                 "Rental Income", "Total Revenue"]),
    ("EXPENSES", ["Total Departmental Expenses", "Total Undistributed Expenses",
                  "Non-Operating Expenses", "Total Non-OperatingExpenses"]),
    ("PROFIT", ["Total Department Profit or Loss", "Operating Profit or Loss",
                "Net Income or Loss"]),
    ("OPERATING STATISTICS", ["Rooms Available", "Rooms Sold", "A.D.R.",
                              "Occupancy", "REV PAR"]),
]
SUMMARY_LINES = [l for _, g in SUMMARY_GROUPS for l in g]
RATIO_LINES = {"a.d.r.", "occupancy", "rev par"}   # not dollars

# --------------------------------------------------------------------------
# styling
# --------------------------------------------------------------------------
FONT = "Segoe UI"
INK, GREEN, MUTED = "1C1C1C", "14532D", "6B6B6B"
F_TITLE = Font(name=FONT, size=14, bold=True, color=INK)
F_SUB = Font(name=FONT, size=9, italic=True, color=MUTED)
F_YEAR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_HDR = Font(name=FONT, size=9, bold=True, color=INK)
F_SECT = Font(name=FONT, size=10, bold=True, color=GREEN)
F_LBL = Font(name=FONT, size=10, color=INK)
F_TOT = Font(name=FONT, size=10, bold=True, color=INK)
F_INP = Font(name=FONT, size=10, color="0F52A8")            # Projected = you type it
F_INPB = Font(name=FONT, size=10, bold=True, color="0F52A8")
FILL_YEAR = PatternFill("solid", fgColor="14532D")
FILL_ALT = PatternFill("solid", fgColor="F4F6F5")
FILL_BAD = PatternFill("solid", fgColor="F8D2D2")
FILL_GOOD = PatternFill("solid", fgColor="D8EDDF")
thin = Side(style="thin", color="4A4A4A")
hair = Side(style="hair", color="C9C9C9")
MONEY = '#,##0;(#,##0)'
RATIO = '#,##0.00'


# --------------------------------------------------------------------------
# parsing
# --------------------------------------------------------------------------
class Line:
    __slots__ = ("page", "section", "label", "act", "bud", "is_total")

    def __init__(self, page, section, label, act, bud, is_total):
        self.page, self.section, self.label = page, section, label
        self.act, self.bud, self.is_total = act, bud, is_total

    @property
    def key(self):
        return (self.page, self.section, self.label)


def _txt(sheet, r, c):
    try:
        v = sheet.cell_value(r, c)
    except IndexError:
        return ""
    return v.strip() if isinstance(v, str) else v


def _num(sheet, r, c):
    try:
        v = sheet.cell_value(r, c)
    except IndexError:
        return None
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def parse_statement(path=None, *, data: bytes = None, name: str = ""):
    """-> (hotel, year, [Line, ...])  from a file path or raw .xls bytes."""
    if data is not None:
        book = xlrd.open_workbook(file_contents=data)
        path = Path(name or "uploaded.xls")
    else:
        book = xlrd.open_workbook(str(path), on_demand=False)
    sh = book.sheet_by_index(0)

    hotel, year = None, None
    for r in range(min(sh.nrows, 12)):
        g = str(_txt(sh, r, 6))
        if "Property:" in g and "Operating" in g and hotel is None:
            hotel = g.split("Property:")[-1].strip()
        m = re.search(r"As of\s+\d{1,2}/\d{1,2}/(\d{4})", g)
        if m:
            year = int(m.group(1))
    if hotel is None:
        for r in range(min(sh.nrows, 12)):
            g = str(_txt(sh, r, 6))
            if "Property:" in g:
                hotel = g.split("Property:")[-1].strip()
                break
    if hotel is None or year is None:
        raise ValueError(f"{path.name}: could not read hotel name / year from header")

    lines, page_title, section = [], None, None
    seen_page_header = False
    for r in range(sh.nrows):
        if str(_txt(sh, r, 0)) == "PTD":          # column-header row = new page
            page_title, section, seen_page_header = None, None, True
            continue
        if not seen_page_header:
            continue
        label = _txt(sh, r, COL_LABEL)
        if not isinstance(label, str) or not label:
            continue
        act = _num(sh, r, COL_YTD_ACT)
        bud = _num(sh, r, COL_YTD_BUD)
        if act is None and bud is None:           # a heading, not a data row
            if page_title is None:
                page_title = label
            section = label
            continue
        if page_title is None:
            page_title = "Summary"
        low = label.lower()
        lines.append(Line(page_title, section or page_title, label,
                          act or 0.0, bud or 0.0,
                          low.startswith("total") or "profit or loss" in low
                          or low.startswith("net income")))
    return hotel, year, lines


# --------------------------------------------------------------------------
# workbook writing
# --------------------------------------------------------------------------
def _year_header(ws, years, start_row, first_col=2):
    """Two-row banner: merged year over Actual / Budget / Projected / Variance."""
    for i, yr in enumerate(years):
        c0 = first_col + i * 4
        ws.merge_cells(start_row=start_row, start_column=c0,
                       end_row=start_row, end_column=c0 + 3)
        cell = ws.cell(row=start_row, column=c0, value=yr)
        cell.font, cell.fill = F_YEAR, FILL_YEAR
        cell.alignment = Alignment(horizontal="center")
        cell.number_format = '0'
        for j, h in enumerate(["Actual", "Budget", "Projected", "Var vs Bud"]):
            hc = ws.cell(row=start_row + 1, column=c0 + j, value=h)
            hc.font, hc.alignment = F_HDR, Alignment(horizontal="right", wrap_text=True)
            hc.border = Border(bottom=thin)
            ws.column_dimensions[gcl(c0 + j)].width = 13
        for rr in (start_row, start_row + 1):
            ws.cell(row=rr, column=c0).border = Border(left=thin, bottom=thin)


def _write_grid(ws, rows, years, data, start_row, first_col=2, at=None):
    """rows: list of (key, label, is_total). data: {year: {key: (act, bud)}}
    `at` (optional dict) is filled with label.lower() -> row for chart wiring."""
    r = start_row
    band = False
    for key, label, is_total in rows:
        if key is None:                                   # section heading
            ws.cell(row=r, column=1, value=label).font = F_SECT
            ws.cell(row=r, column=1).border = Border(bottom=thin)
            r += 1
            band = False
            continue
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = F_TOT if is_total else F_LBL
        lc.alignment = Alignment(indent=0 if is_total else 1)
        ratio = label.strip().lower() in RATIO_LINES
        fmt = RATIO if ratio else MONEY
        for i, yr in enumerate(years):
            c0 = first_col + i * 4
            act, bud = data.get(yr, {}).get(key, (0.0, 0.0))
            a = ws.cell(row=r, column=c0, value=act)
            b = ws.cell(row=r, column=c0 + 1, value=bud)
            p = ws.cell(row=r, column=c0 + 2, value=None)      # user-supplied
            v = ws.cell(row=r, column=c0 + 3,
                        value=f"={gcl(c0)}{r}-{gcl(c0+1)}{r}")
            for cell in (a, b, p, v):
                cell.number_format = fmt
                cell.font = F_TOT if is_total else F_LBL
                if band:
                    cell.fill = FILL_ALT
            p.font = F_INPB if is_total else F_INP
            if is_total:
                for cell in (a, b, p, v):
                    cell.border = Border(top=hair)
        if band:
            lc.fill = FILL_ALT
        if at is not None:
            at.setdefault(label.strip().lower(), r)
        band = not band
        r += 1
    return r


def _apply_variance_rules(ws, years, first_data_row, last_row, first_col=2):
    for i in range(len(years)):
        c0 = first_col + i * 4
        var, bud = gcl(c0 + 3), gcl(c0 + 1)
        rng = f"{var}{first_data_row}:{var}{last_row}"
        over = (f'=AND({bud}{first_data_row}<>0,'
                f'ABS({var}{first_data_row})>{VAR_MIN},'
                f'ABS({var}{first_data_row}/{bud}{first_data_row})>{VAR_PCT},'
                f'{var}{first_data_row}<0)')
        under = over.replace(f'{var}{first_data_row}<0)', f'{var}{first_data_row}>0)')
        ws.conditional_formatting.add(rng, FormulaRule(formula=[over[1:]], fill=FILL_BAD))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[under[1:]], fill=FILL_GOOD))


def _add_charts(ws, years, at, start_row):
    """A small year-by-year block under the grid, plus three charts off it."""
    from openpyxl.chart import BarChart, LineChart, Reference

    METRICS = [("Total Revenue", "total revenue"),
               ("Room", "room"),
               ("Food & Beverage", "food & beverage"),
               ("Miscellaneous", "miscellaneous"),
               ("Rental Income", "rental income"),
               ("Operating Profit", "operating profit or loss"),
               ("Net Income", "net income or loss"),
               ("Total Revenue Budget", "total revenue")]

    hdr = start_row
    ws.cell(row=hdr, column=1, value="YEAR-OVER-YEAR DATA  (feeds the charts)").font = F_SECT
    ws.cell(row=hdr, column=1).border = Border(bottom=thin)
    ws.cell(row=hdr + 1, column=1, value="Year").font = F_HDR
    for j, (title, _) in enumerate(METRICS):
        c = ws.cell(row=hdr + 1, column=2 + j, value=title)
        c.font, c.alignment = F_HDR, Alignment(horizontal="right", wrap_text=True)
        c.border = Border(bottom=thin)
        ws.column_dimensions[gcl(2 + j)].width = 15

    for i, yr in enumerate(years):
        r = hdr + 2 + i
        ws.cell(row=r, column=1, value=yr).number_format = '0'
        ws.cell(row=r, column=1).font = F_LBL
        src_a = gcl(2 + i * 4)          # Actual column for this year in the grid above
        src_b = gcl(3 + i * 4)          # Budget column
        for j, (title, key) in enumerate(METRICS):
            row_in_grid = at.get(key)
            col = src_b if title.endswith("Budget") else src_a
            cell = ws.cell(row=r, column=2 + j,
                           value=f"={col}{row_in_grid}" if row_in_grid else 0)
            cell.number_format, cell.font = MONEY, F_LBL

    first, last = hdr + 2, hdr + 1 + len(years)
    cats = Reference(ws, min_col=1, min_row=first, max_row=last)
    anchor_row = last + 2

    def place(chart, title, col_idxs, anchor, kind="bar", stacked=False):
        chart.title = title
        chart.height, chart.width = 8.5, 17
        chart.y_axis.numFmt = '#,##0'
        chart.legend.position = "b"
        for ci in col_idxs:
            chart.add_data(Reference(ws, min_col=ci, min_row=first - 1, max_row=last),
                           titles_from_data=True)
        chart.set_categories(cats)
        if kind == "bar":
            chart.type = "col"
            if stacked:
                chart.grouping, chart.overlap = "stacked", 100
        ws.add_chart(chart, anchor)

    place(BarChart(), "TOTAL REVENUE  —  ACTUAL vs BUDGET", [2, 9], f"A{anchor_row}")
    place(BarChart(), "REVENUE MIX BY DEPARTMENT", [3, 4, 5, 6],
          f"K{anchor_row}", stacked=True)
    place(LineChart(), "OPERATING PROFIT & NET INCOME", [7, 8], f"A{anchor_row + 18}")


def _sheet_header(ws, hotel, title, note):
    ws["A1"] = f"{hotel}  —  {title}"
    ws["A1"].font = F_TITLE
    ws["A2"] = note
    ws["A2"].font = F_SUB
    ws.column_dimensions["A"].width = 44
    ws.sheet_view.showGridLines = False


def build_workbook(hotel, per_year, out_path: Path, sources=None):
    # only the years we actually have a statement for - never zero-filled
    years = span = sorted(per_year)

    data = defaultdict(dict)
    order = OrderedDict()          # key -> (page, section, label, is_total)
    for yr in years:
        for ln in per_year[yr]:
            data[yr][ln.key] = (ln.act, ln.bud)
            order.setdefault(ln.key, ln)

    wb = Workbook()
    wb.remove(wb.active)

    # ---- Summary ---------------------------------------------------------
    ws = wb.create_sheet("Summary")
    _sheet_header(ws, hotel, "SUMMARY",
                  f"One column block per year you have a statement for. "
                  f"Variance highlights when the gap tops {VAR_PCT:.0%} and ${VAR_MIN:,}. Charts are below.")
    _year_header(ws, span, 4)
    rows, used = [], set()
    for heading, labels in SUMMARY_GROUPS:
        rows.append((None, heading, False))
        for name in labels:
            for key, ln in order.items():
                if (ln.page == "Summary" and key not in used
                        and ln.label.lower() == name.lower()):
                    rows.append((key, ln.label, ln.is_total))
                    used.add(key)
                    break
    at = {}
    end = _write_grid(ws, rows, span, data, 6, at=at)
    _apply_variance_rules(ws, span, 6, end - 1)

    # margins, computed off the rows just written
    ws.cell(row=end + 1, column=1, value="MARGINS").font = F_SECT
    ws.cell(row=end + 1, column=1).border = Border(bottom=thin)
    rev, op, ni = at.get("total revenue"), at.get("operating profit or loss"),         at.get("net income or loss")
    for k, (lab, src) in enumerate([("Operating Profit Margin", op),
                                    ("Net Income Margin", ni)]):
        r = end + 2 + k
        ws.cell(row=r, column=1, value=lab).font = F_LBL
        if not (src and rev):
            continue
        for i in range(len(span)):
            for j in (0, 1):                       # actual and budget columns
                c = gcl(2 + i * 4 + j)
                cell = ws.cell(row=r, column=2 + i * 4 + j,
                               value=f'=IF({c}{rev}=0,"",{c}{src}/{c}{rev})')
                cell.number_format, cell.font = '0.0%', F_LBL
    ws.freeze_panes = "B6"
    _add_charts(ws, span, at, end + 5)

    # ---- department tabs -------------------------------------------------
    tabs = OrderedDict()
    for key, ln in order.items():
        tab = TAB_MAP.get(ln.page)
        if tab:
            tabs.setdefault(tab, []).append((key, ln))
    for tab in ["Rooms", "Food", "Beverage", "Miscellaneous"]:
        items = tabs.get(tab, [])
        ws = wb.create_sheet(tab)
        _sheet_header(ws, hotel, tab.upper(),
                      "Actual and Budget come straight from the operating statements. "
                      "Projected is yours to fill in. Var = Actual - Budget.")
        _year_header(ws, span, 4)
        rows, cur = [], None
        for key, ln in items:
            if ln.section != cur:
                rows.append((None, ln.section, False))
                cur = ln.section
            rows.append((key, ln.label, ln.is_total))
        end = _write_grid(ws, rows, span, data, 6)
        _apply_variance_rules(ws, span, 6, max(end - 1, 6))
        ws.freeze_panes = "B6"

    # ---- Fixed Expenses --------------------------------------------------
    ws = wb.create_sheet("Fixed Expenses")
    _sheet_header(ws, hotel, "FIXED EXPENSES",
                  "Real estate taxes and insurance only — corporate and payroll taxes "
                  "are deliberately excluded.")
    _year_header(ws, span, 4)
    rows = [(None, "FIXED EXPENSES", False)]
    fx = [f.lower() for f in FIXED_LINES]
    for key, ln in order.items():
        if ln.label.lower() in fx and ln.page == "Non-Operating Expenses":
            rows.append((key, ln.label, False))
    end = _write_grid(ws, rows, span, data, 6)
    tot = end
    ws.cell(row=tot, column=1, value="TOTAL FIXED EXPENSES").font = F_TOT
    for i in range(len(span)):
        c0 = 2 + i * 4
        for j in range(4):
            col = gcl(c0 + j)
            c = ws.cell(row=tot, column=c0 + j,
                        value=f"=SUM({col}7:{col}{tot-1})")
            c.number_format, c.font = MONEY, F_TOT
            c.border = Border(top=thin, bottom=Side(style="double", color=INK))
    _apply_variance_rules(ws, span, 6, tot)
    ws.freeze_panes = "B6"

    for s in wb.worksheets:
        s.sheet_properties.tabColor = GREEN if s.title != "Summary" else "1C1C1C"
    # the chart-feed cells are formulas with no cached result; without this the
    # charts open empty until something forces a recalculation
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return span


def main():
    ap = argparse.ArgumentParser(description="Build hotel P&L workbooks from operating statements.")
    ap.add_argument("--input", required=True, help="folder containing .xls operating statements")
    ap.add_argument("--output", default=None, help="output folder (default: alongside input)")
    a = ap.parse_args()

    src = Path(a.input)
    out = Path(a.output) if a.output else src
    out.mkdir(parents=True, exist_ok=True)

    files = sorted(p for p in src.glob("*.xls") if not p.name.startswith("~$"))
    if not files:
        sys.exit(f"no .xls files found in {src}")

    hotels = defaultdict(dict)
    srcmap = defaultdict(dict)
    for p in files:
        try:
            hotel, year, lines = parse_statement(p)
        except Exception as e:                                  # noqa: BLE001
            print(f"  SKIPPED {p.name}: {e}")
            continue
        hotels[hotel][year] = lines
        srcmap[hotel][year] = p.name
        print(f"  read {p.name:52} -> {hotel} {year} ({len(lines)} lines)")

    for hotel, per_year in hotels.items():
        safe = re.sub(r"[^\w\- ]", "", hotel).strip() or "Hotel"
        dest = out / f"{safe} - P&L.xlsx"
        span = build_workbook(hotel, per_year, dest, srcmap[hotel])
        print(f"\n{hotel}: {len(span)} year(s)  ->  {dest.name}")
        print(f"  years included : {span}")


if __name__ == "__main__":
    main()
