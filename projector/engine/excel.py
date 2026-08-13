"""Build the downloadable Excel workbook as a live model, not a snapshot.

The Assumptions sheet holds real input cells behind defined names. Every
forecast sheet is built from Excel formulas that read those names, so changing
a growth rate, the capacity or the basis in Excel recalculates the day-by-day
projection, the monthly roll-up, the segment tabs and the variances.

Only two kinds of cell are hard numbers:

* the per-day **base** rooms and ADR, which come out of the uploaded history and
  cannot depend on an assumption, and
* the **history** and **growth** sheets, which are actuals.

Every formula is written with its Python-computed result cached alongside, so
the workbook shows the right numbers even in a viewer that never recalculates.
"""

from __future__ import annotations

import io

import numpy as np
import pandas as pd
from xlsxwriter.utility import xl_range_abs, xl_rowcol_to_cell

from . import growth
from . import history as hist
from .forecast import Assumptions, ForecastModel, compare_to_history, expand

DAILY_SHEET = "Forecast Day by Day"
MONTHLY_SHEET = "Forecast Monthly"
GROWTH_SHEET = "Growth"

_NUMBER_HINTS = (
    ("OCC %", "pct"),
    ("Mix %", "pct"),
    ("pts", "pct"),
    ("Y/Y %", "pct"),
    ("vs LY %", "pct"),
    ("ADR", "money"),
    ("RevPAR", "money"),
    ("Revenue", "money0"),
    ("Rooms", "int"),
    ("Days", "int"),
)


def _is_percent(column: str) -> bool:
    """Columns measured on a percentage scale, by name."""
    name = str(column).rstrip()
    return name.endswith("%") or " pts" in name or name.endswith("pts")


def _fmt_for(column: str) -> str | None:
    # Anything on a percentage scale is a rate, whatever else the name contains --
    # "Rooms historical %" must not pick up the "Rooms" integer format.
    if _is_percent(column):
        return "pct"
    for hint, style in _NUMBER_HINTS:
        if hint in str(column):
            return style
    return None


def _to_fractions(frame: pd.DataFrame, percent_cols) -> tuple[pd.DataFrame, set]:
    """Rescale percentage columns to fractions so Excel's % format can show them.

    The engine works in percentage points because that is what the app displays;
    Excel wants 0.5185 behind a 0.00% format, not 51.85.
    """
    cols = set(percent_cols) if percent_cols is not None else {
        c for c in frame.columns if _is_percent(c)
    }
    cols = {c for c in cols if c in frame.columns}
    if not cols:
        return frame, cols
    out = frame.copy()
    for column in cols:
        out[column] = pd.to_numeric(out[column], errors="coerce") / 100.0
    return out, cols


def _formats(book) -> dict:
    return {
        "header": book.add_format(
            {"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1,
             "text_wrap": True, "valign": "vcenter"}
        ),
        "helper_header": book.add_format(
            {"bold": True, "bg_color": "#8496B0", "font_color": "white", "border": 1,
             "text_wrap": True, "valign": "vcenter"}
        ),
        "title": book.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864"}),
        "section": book.add_format(
            {"bold": True, "bg_color": "#D6DCE4", "border": 1, "font_color": "#1F3864"}
        ),
        "label": book.add_format({"align": "left"}),
        "note": book.add_format({"italic": True, "font_color": "#666666", "text_wrap": True}),
        "input": book.add_format(
            {"bg_color": "#FFF2CC", "border": 1, "num_format": "#,##0.00"}
        ),
        "input_pct": book.add_format({"bg_color": "#FFF2CC", "border": 1, "num_format": "0.00%"}),
        "input_int": book.add_format({"bg_color": "#FFF2CC", "border": 1, "num_format": "#,##0"}),
        "input_text": book.add_format({"bg_color": "#FFF2CC", "border": 1}),
        "derived_pct": book.add_format({"num_format": "0.00%", "font_color": "#1F3864"}),
        # Whole numbers everywhere, percentages as real percentages to 2dp.
        "money": book.add_format({"num_format": "$#,##0"}),
        "money0": book.add_format({"num_format": "$#,##0"}),
        "pct": book.add_format({"num_format": "0.00%"}),
        "int": book.add_format({"num_format": "#,##0"}),
        "date": book.add_format({"num_format": "yyyy-mm-dd"}),
    }


# --------------------------------------------------------------------------- #
# Value sheets (history and growth: actuals, nothing to recalculate)
# --------------------------------------------------------------------------- #

def _write_values(writer, name: str, frame: pd.DataFrame, formats: dict, freeze_col: int = 1,
                  percent_cols=None):
    sheet = name[:31]
    frame, pct_cols = _to_fractions(frame, percent_cols)
    frame.to_excel(writer, sheet_name=sheet, index=False, startrow=1, header=False)
    ws = writer.sheets[sheet]
    for idx, column in enumerate(frame.columns):
        ws.write(0, idx, str(column), formats["header"])
        style = "pct" if column in pct_cols else _fmt_for(str(column))
        ws.set_column(idx, idx, max(10, min(28, len(str(column)) + 3)),
                      formats.get(style) if style else None)
    ws.freeze_panes(1, freeze_col)
    ws.autofilter(0, 0, max(len(frame), 1), max(len(frame.columns) - 1, 0))
    return ws


def block_rows(blocks: list[tuple[str, pd.DataFrame, object]]) -> list[tuple[int, int, int]]:
    """Row positions each block will occupy: (header row, first data row, last).

    Pure arithmetic mirroring `_write_blocks`, so the Assumptions sheet can point
    formulas at a block that has not been written yet and the tab order can still
    put Assumptions first.
    """
    positions, row = [], 0
    for _, frame, _ in blocks:
        positions.append((row + 1, row + 2, row + 1 + len(frame)))
        row += 2 + len(frame) + 2
    return positions


def _write_blocks(book, formats: dict, name: str,
                  blocks: list[tuple[str, pd.DataFrame, object]]) -> None:
    """Stack several titled tables down one sheet instead of spreading them over tabs."""
    ws = book.add_worksheet(name[:31])
    ws.hide_gridlines(2)
    widths: dict[int, int] = {}
    row = 0
    for title, frame, percent_cols in blocks:
        frame, pct_cols = _to_fractions(frame, percent_cols)
        ws.write(row, 0, title, formats["section"])
        for c in range(1, max(len(frame.columns), 1)):
            ws.write(row, c, "", formats["section"])
        row += 1
        for c, column in enumerate(frame.columns):
            ws.write(row, c, str(column), formats["header"])
            widths[c] = max(widths.get(c, 10), min(28, len(str(column)) + 3))
        row += 1
        for _, record in frame.iterrows():
            for c, column in enumerate(frame.columns):
                value = record[column]
                style = "pct" if column in pct_cols else _fmt_for(str(column))
                fmt = formats.get(style) if style else None
                if isinstance(value, (bool, np.bool_)):
                    ws.write_boolean(row, c, bool(value))
                elif pd.isna(value):
                    ws.write_blank(row, c, None, fmt)
                elif isinstance(value, (int, float, np.integer, np.floating)):
                    ws.write_number(row, c, float(value), fmt)
                else:
                    ws.write_string(row, c, str(value))
            row += 1
        row += 2  # a blank line between blocks

    for c, width in widths.items():
        ws.set_column(c, c, width)
    ws.freeze_panes(0, 1)


# --------------------------------------------------------------------------- #
# Assumptions: the input sheet everything else points at
# --------------------------------------------------------------------------- #

def _write_assumptions(
    book, formats: dict, a: Assumptions, capacity: int, segments: list[str],
    weights: np.ndarray, derived: pd.DataFrame | None, bases_at: tuple[int, int, int],
    notes: list[str],
) -> None:
    ws = book.add_worksheet("Assumptions")
    ws.set_column(0, 0, 34)
    ws.set_column(1, 6, 15)
    ws.hide_gridlines(2)

    ws.write(0, 0, "BUDGET ASSUMPTIONS", formats["title"])
    ws.write(1, 0, "Every shaded cell is an input. All Forecast sheets recalculate from them.",
             formats["note"])

    def name(label: str, row: int, col: int = 1) -> None:
        book.define_name(label, f"=Assumptions!{xl_rowcol_to_cell(row, col, True, True)}")

    row = 3
    ws.write(row, 0, "HOTEL", formats["section"])
    ws.write(row, 1, "", formats["section"])
    row += 1
    ws.write(row, 0, "Rooms available (capacity)", formats["label"])
    ws.write_number(row, 1, capacity, formats["input_int"])
    name("Capacity", row)
    row += 1
    ws.write(row, 0, "Occupancy ceiling", formats["label"])
    ws.write_number(row, 1, a.occ_ceiling, formats["input_pct"])
    name("OccCeiling", row)
    row += 1
    ws.write(row, 0, "Round rooms to whole numbers (1 = yes)", formats["label"])
    ws.write_number(row, 1, 1 if a.round_rooms else 0, formats["input_int"])
    name("RoundRooms", row)

    row += 2
    ws.write(row, 0, "GROWTH", formats["section"])
    ws.write(row, 1, "", formats["section"])
    row += 1
    ws.write(row, 0, "Derived from", formats["label"])
    ws.write_string(row, 1, a.growth_basis or growth.DEFAULT_BASIS, formats["input_text"])
    ws.data_validation(row, 1, row, 1, {"validate": "list", "source": growth.BASES})
    name("Basis", row)
    row += 1
    ws.write(row, 0, "Share of the historical rate applied", formats["label"])
    ws.write_number(row, 1, a.growth_damping, formats["input_pct"])
    name("Damping", row)
    row += 1
    ws.write(row, 0, "Cap derived growth at +/- (0 = no cap)", formats["label"])
    ws.write_number(row, 1, a.growth_cap or 0.0, formats["input_pct"])
    name("GrowthCap", row)
    row += 1
    ws.write(row, 0, "ADR adjustment (on top of derived)", formats["label"])
    ws.write_number(row, 1, a.adr_growth, formats["input_pct"])
    name("AdrAdj", row)
    row += 1
    ws.write(row, 0, "Occupancy adjustment (on top of derived)", formats["label"])
    ws.write_number(row, 1, a.occ_growth, formats["input_pct"])
    name("OccAdj", row)

    # --- segment growth -------------------------------------------------- #
    row += 2
    ws.write(row, 0, "SEGMENT GROWTH", formats["section"])
    for c in range(1, 7):
        ws.write(row, c, "", formats["section"])
    row += 1
    ws.write(row, 0, "Leave an override blank to keep the rate derived from the basis above.",
             formats["note"])
    row += 1
    headers = ["Segment", "ADR derived", "ADR override", "ADR in use",
               "Occ derived", "Occ override", "Occ in use"]
    for c, head in enumerate(headers):
        ws.write(row, c, head, formats["header"])
    seg_first = row + 1

    # The basis table is a block partway down the Growth sheet, not its own tab.
    bases_header, bases_first, bases_last = bases_at
    key_range = xl_range_abs(bases_first, 0, bases_last, 0)
    value_range = xl_range_abs(bases_first, 3, bases_last, 6)
    basis_header = xl_range_abs(bases_header, 3, bases_header, 6)

    for i, seg in enumerate(segments):
        r = seg_first + i
        seg_cell = xl_rowcol_to_cell(r, 0, True, False)
        ws.write_string(r, 0, seg)
        for metric, derived_col, override_col, inuse_col in (
            ("ADR", 1, 2, 3), ("Rooms", 4, 5, 6),
        ):
            # The Growth sheet holds these as fractions already -- no /100 here.
            raw = (
                f"Damping*IFERROR(INDEX('{GROWTH_SHEET}'!{value_range},"
                f"MATCH({seg_cell}&\"|{metric}\",'{GROWTH_SHEET}'!{key_range},0),"
                f"MATCH(Basis,'{GROWTH_SHEET}'!{basis_header},0)),0)"
            )
            # ROUND to 4dp mirrors engine.growth.derive, so a recalculation in
            # Excel lands on the same rate the app used.
            formula = f"=ROUND(IF(GrowthCap<=0,{raw},MAX(-GrowthCap,MIN(GrowthCap,{raw}))),4)"
            cached = 0.0
            if derived is not None:
                match = derived.loc[derived["Segment"] == seg]
                if not match.empty:
                    cached = float(match[f"{metric} applied %"].iat[0]) / 100
            ws.write_formula(r, derived_col, formula, formats["derived_pct"], cached)

            # An edit made in the app travels as an override, otherwise the
            # workbook would silently revert it to the derived rate.
            in_use = (a.segment_adr_growth if metric == "ADR" else a.segment_occ_growth).get(
                seg, cached
            )
            if abs(in_use - cached) > 1e-9:
                ws.write_number(r, override_col, in_use, formats["input_pct"])
            else:
                ws.write_blank(r, override_col, None, formats["input_pct"])

            d_cell = xl_rowcol_to_cell(r, derived_col)
            o_cell = xl_rowcol_to_cell(r, override_col)
            ws.write_formula(
                r, inuse_col, f'=IF({o_cell}="",{d_cell},{o_cell})',
                formats["derived_pct"], in_use,
            )

    seg_last = seg_first + len(segments) - 1
    book.define_name("SegAdr", f"=Assumptions!{xl_range_abs(seg_first, 3, seg_last, 3)}")
    book.define_name("SegOcc", f"=Assumptions!{xl_range_abs(seg_first, 6, seg_last, 6)}")

    # --- month overrides -------------------------------------------------- #
    row = seg_last + 2
    ws.write(row, 0, "MONTH OVERRIDES", formats["section"])
    for c in range(1, 3):
        ws.write(row, c, "", formats["section"])
    row += 1
    for c, head in enumerate(["Month", "ADR", "Occupancy"]):
        ws.write(row, c, head, formats["header"])
    month_first = row + 1
    for m in range(12):
        r = month_first + m
        ws.write_string(r, 0, pd.Timestamp(2000, m + 1, 1).strftime("%B"))
        ws.write_number(r, 1, a.month_adr_growth.get(m + 1, 0.0), formats["input_pct"])
        ws.write_number(r, 2, a.month_occ_growth.get(m + 1, 0.0), formats["input_pct"])
    month_last = month_first + 11
    book.define_name("MonAdr", f"=Assumptions!{xl_range_abs(month_first, 1, month_last, 1)}")
    book.define_name("MonOcc", f"=Assumptions!{xl_range_abs(month_first, 2, month_last, 2)}")

    # --- blend weights ----------------------------------------------------- #
    row = month_last + 2
    ws.write(row, 0, "HISTORY BLEND WEIGHTS", formats["section"])
    ws.write(row, 1, "", formats["section"])
    row += 1
    for c, head in enumerate(["Years back", "Weight"]):
        ws.write(row, c, head, formats["header"])
    weight_first = row + 1
    for j, w in enumerate(weights):
        ws.write_number(weight_first + j, 0, j + 1)
        ws.write_number(weight_first + j, 1, float(w), formats["input_pct"])
    weight_last = weight_first + len(weights) - 1
    book.define_name("YrWeight", f"=Assumptions!{xl_range_abs(weight_first, 1, weight_last, 1)}")

    # --- fixed settings and notes ------------------------------------------ #
    row = weight_last + 2
    ws.write(row, 0, "BAKED INTO THE BASE (change in the app, not here)", formats["section"])
    ws.write(row, 1, "", formats["section"])
    # Each row carries its own format. Sniffing the Python type instead put the
    # money format on the smoothing weights, so 0.75 rendered as "$1".
    fixed: list[tuple[str, object, str | None]] = [
        ("Forecast start", a.start.strftime("%Y-%m-%d"), None),
        ("Forecast end", a.end.strftime("%Y-%m-%d"), None),
        ("Segments projected", ", ".join(a.segments), None),
        ("Demand window (days, centred)", 2 * a.smoothing_days + 1, "int"),
        *(
            (
                f"Day smoothing - {seg} (0% = last year's day, 100% = smoothed)",
                a.smoothing_for(seg),
                "pct",
            )
            for seg in a.segments
        ),
    ]
    for i, (label, value, style) in enumerate(fixed, start=1):
        ws.write(row + i, 0, label, formats["label"])
        # Numbers go in as numbers -- writing "0.75" as text makes Excel flag
        # every one of these cells as a number stored as text.
        if style is None:
            ws.write_string(row + i, 1, str(value))
        else:
            ws.write_number(row + i, 1, float(value), formats[style])
    row += len(fixed) + 2
    ws.write(row, 0, "Smoothing applies to rooms only, and blends each night as "
                     "(1 - s) x last year's actual + s x the smoothed estimate. ADR is never "
                     "smoothed. The base rooms and ADR on the day-by-day sheet come from the "
                     "history and do not move with the inputs above.", formats["note"])

    if notes:
        row += 2
        ws.write(row, 0, "DATA NOTES FROM THE IMPORT", formats["section"])
        ws.write(row, 1, "", formats["section"])
        for i, note in enumerate(notes, start=1):
            ws.write(row + i, 0, note, formats["note"])


# --------------------------------------------------------------------------- #
# Day by day: the live model
# --------------------------------------------------------------------------- #

def _daily_layout(segments: list[str], n_years: int) -> tuple[list[str], set[str]]:
    """Column order for the day-by-day sheet, plus which columns are helpers."""
    columns = ["Date", "Day", "Month", "MonthNum", "YearsBack"]
    helpers: set[str] = {"MonthNum", "YearsBack"}
    for seg in segments:
        for j in range(n_years):
            columns.append(f"{seg} Base Rooms Y{j + 1}")
        for j in range(n_years):
            columns.append(f"{seg} Base ADR Y{j + 1}")
        columns += [f"{seg} Occ factor", f"{seg} ADR factor", f"{seg} Raw Rooms"]
        helpers.update(columns[-(2 * n_years + 3):])
    columns += ["Raw Total Rooms", "Capacity Scale"]
    helpers.update({"Raw Total Rooms", "Capacity Scale"})
    for seg in segments:
        columns.append(f"{seg} Cum Rooms")
        helpers.add(f"{seg} Cum Rooms")
    for seg in segments:
        columns.append(f"{seg} Rooms")
    for seg in segments:
        columns.append(f"{seg} ADR")
    for seg in segments:
        columns.append(f"{seg} Revenue")
    columns += ["Total Rooms", "Total Revenue", "Total ADR", "Rooms Available", "OCC %", "RevPAR"]
    return columns, helpers


def _write_daily(book, formats: dict, model: ForecastModel, a: Assumptions,
                 values: dict[str, np.ndarray]) -> list[str]:
    segments = model.segments
    n_years = len(model.weights)
    columns, helpers = _daily_layout(segments, n_years)
    index = {name: i for i, name in enumerate(columns)}

    ws = book.add_worksheet(DAILY_SHEET)
    for i, column in enumerate(columns):
        ws.write(0, i, column, formats["helper_header"] if column in helpers else formats["header"])
        style = _fmt_for(column)
        options = {"level": 1, "hidden": True} if column in helpers else {}
        ws.set_column(i, i, max(10, min(26, len(column) + 3)),
                      formats.get(style) if style else None, options)
    ws.set_column(index["Date"], index["Date"], 12, formats["date"])
    ws.freeze_panes(1, 3)
    ws.outline_settings(True, False, True, False)

    def cell(column: str, r: int, abs_col: bool = False) -> str:
        return xl_rowcol_to_cell(r + 1, index[column], False, abs_col)

    months = model.dates.month.to_numpy()
    month_labels = model.dates.to_period("M").astype(str)

    for r in range(len(model.dates)):
        ws.write_datetime(r + 1, index["Date"], model.dates[r].to_pydatetime(), formats["date"])
        ws.write_string(r + 1, index["Day"], model.dates[r].strftime("%a"))
        ws.write_string(r + 1, index["Month"], month_labels[r])
        ws.write_number(r + 1, index["MonthNum"], int(months[r]))
        ws.write_number(r + 1, index["YearsBack"], int(model.years_back[r]))

        for seg in segments:
            for j in range(n_years):
                ws.write_number(r + 1, index[f"{seg} Base Rooms Y{j + 1}"],
                                float(model.base_rooms[seg][j][r]))
                ws.write_number(r + 1, index[f"{seg} Base ADR Y{j + 1}"],
                                float(model.base_adr[seg][j][r]))

        month_cell = cell("MonthNum", r, abs_col=True)
        years_cell = cell("YearsBack", r, abs_col=True)

        for i, seg in enumerate(segments, start=1):
            ws.write_formula(
                r + 1, index[f"{seg} Occ factor"],
                f"=(1+INDEX(SegOcc,{i}))*(1+OccAdj)*(1+INDEX(MonOcc,{month_cell}))",
                None, float(values[f"{seg} Occ factor"][r]),
            )
            ws.write_formula(
                r + 1, index[f"{seg} ADR factor"],
                f"=(1+INDEX(SegAdr,{i}))*(1+AdrAdj)*(1+INDEX(MonAdr,{month_cell}))",
                None, float(values[f"{seg} ADR factor"][r]),
            )

            occ_f = cell(f"{seg} Occ factor", r)
            adr_f = cell(f"{seg} ADR factor", r)
            rooms_terms = "+".join(
                f"INDEX(YrWeight,{j + 1})*{cell(f'{seg} Base Rooms Y{j + 1}', r)}"
                f"*{occ_f}^({years_cell}+{j})"
                for j in range(n_years)
            )
            adr_terms = "+".join(
                f"INDEX(YrWeight,{j + 1})*{cell(f'{seg} Base ADR Y{j + 1}', r)}"
                f"*{adr_f}^({years_cell}+{j})"
                for j in range(n_years)
            )
            ws.write_formula(
                r + 1, index[f"{seg} Raw Rooms"], f"=({rooms_terms})/SUM(YrWeight)",
                None, float(values[f"{seg} Raw Rooms"][r]),
            )
            ws.write_formula(
                r + 1, index[f"{seg} ADR"], f"=MAX(0,({adr_terms})/SUM(YrWeight))",
                formats["money"], float(values[f"{seg} ADR"][r]),
            )

        raw_sum = "+".join(cell(f"{seg} Raw Rooms", r) for seg in segments)
        ws.write_formula(r + 1, index["Raw Total Rooms"], f"={raw_sum}",
                         None, float(values["Raw Total Rooms"][r]))
        raw_total = cell("Raw Total Rooms", r)
        ws.write_formula(
            r + 1, index["Capacity Scale"],
            f"=IF(AND(Capacity*OccCeiling>0,{raw_total}>Capacity*OccCeiling),"
            f"Capacity*OccCeiling/{raw_total},1)",
            None, float(values["Capacity Scale"][r]),
        )

        scale = cell("Capacity Scale", r)
        for seg in segments:
            raw = cell(f"{seg} Raw Rooms", r)
            # Running total of unrounded rooms, so the rounding below can carry
            # its remainder forward instead of leaking it.
            running = f"{raw}*{scale}"
            if r > 0:
                running = f"{cell(f'{seg} Cum Rooms', r - 1)}+{running}"
            ws.write_formula(
                r + 1, index[f"{seg} Cum Rooms"], f"={running}",
                None, float(values[f"{seg} Cum Rooms"][r]),
            )

        for seg in segments:
            raw = cell(f"{seg} Raw Rooms", r)
            cum = cell(f"{seg} Cum Rooms", r)
            # Difference of the rounded running totals. Rounding each day alone
            # would drop every day holding under half a room and lose volume.
            rounded = (
                f"ROUND({cum},0)-ROUND({cell(f'{seg} Cum Rooms', r - 1)},0)"
                if r > 0
                else f"ROUND({cum},0)"
            )
            ws.write_formula(
                r + 1, index[f"{seg} Rooms"],
                f"=IF(RoundRooms=1,{rounded},{raw}*{scale})",
                formats["int"], float(values[f"{seg} Rooms"][r]),
            )
        for seg in segments:
            ws.write_formula(
                r + 1, index[f"{seg} Revenue"],
                f"={cell(f'{seg} Rooms', r)}*{cell(f'{seg} ADR', r)}",
                formats["money0"], float(values[f"{seg} Revenue"][r]),
            )

        rooms_range = f"{cell(f'{segments[0]} Rooms', r)}:{cell(f'{segments[-1]} Rooms', r)}"
        rev_range = f"{cell(f'{segments[0]} Revenue', r)}:{cell(f'{segments[-1]} Revenue', r)}"
        ws.write_formula(r + 1, index["Total Rooms"], f"=SUM({rooms_range})",
                         formats["int"], float(values["Total Rooms"][r]))
        ws.write_formula(r + 1, index["Total Revenue"], f"=SUM({rev_range})",
                         formats["money0"], float(values["Total Revenue"][r]))
        total_rooms, total_rev = cell("Total Rooms", r), cell("Total Revenue", r)
        ws.write_formula(r + 1, index["Total ADR"],
                         f"=IF({total_rooms}>0,{total_rev}/{total_rooms},0)",
                         formats["money"], float(values["Total ADR"][r]))
        ws.write_formula(r + 1, index["Rooms Available"], "=Capacity",
                         formats["int"], float(values["Rooms Available"][r]))
        ws.write_formula(r + 1, index["OCC %"],
                         f"=IF(Capacity>0,{total_rooms}/Capacity,0)",
                         formats["pct"], float(values["OCC %"][r]) / 100)
        ws.write_formula(r + 1, index["RevPAR"],
                         f"=IF(Capacity>0,{total_rev}/Capacity,0)",
                         formats["money"], float(values["RevPAR"][r]))

    return columns


# --------------------------------------------------------------------------- #
# Monthly roll-up: SUMIF over the day-by-day sheet
# --------------------------------------------------------------------------- #

def _write_monthly(book, formats: dict, model: ForecastModel, a: Assumptions,
                   daily_columns: list[str], values: dict[str, np.ndarray],
                   comparison: pd.DataFrame) -> tuple[list[str], list[str]]:
    segments = model.segments
    n_rows = len(model.dates)
    index = {name: i for i, name in enumerate(daily_columns)}

    def daily_range(column: str) -> str:
        col = index[column]
        return f"'{DAILY_SHEET}'!{xl_range_abs(1, col, n_rows, col)}"

    frame = pd.DataFrame({"Month": model.dates.to_period("M").astype(str)})
    for column in [f"{s} Rooms" for s in segments] + [f"{s} Revenue" for s in segments] + [
        "Total Rooms", "Total Revenue"
    ]:
        frame[column] = values[column]
    grouped = frame.groupby("Month", sort=True)
    months = list(grouped.groups.keys())
    sums = grouped.sum(numeric_only=True)
    days = grouped.size()

    columns = ["Month", "Days", "Rooms Available"]
    segment_block: list[str] = []
    for seg in segments:
        segment_block += [f"{seg} Rooms", f"{seg} Revenue", f"{seg} ADR",
                          f"{seg} OCC %", f"{seg} Mix %"]
    columns += segment_block
    columns += ["Total Rooms", "Total Revenue", "Total ADR", "OCC %", "RevPAR"]
    # The LY comparison lives here rather than on its own sheet -- it is the same
    # twelve rows, and splitting it made the reader hold two tabs in their head.
    columns += ["LY Month", "LY Rooms", "LY ADR", "LY Revenue", "LY OCC %",
                "Rooms vs LY %", "ADR vs LY %", "Revenue vs LY %", "OCC vs LY"]

    ws = book.add_worksheet(MONTHLY_SHEET)
    col_of = {name: i for i, name in enumerate(columns)}
    for i, column in enumerate(columns):
        ws.write(0, i, column, formats["header"])
        style = _fmt_for(column)
        # Per-segment detail is collapsed by default so the sheet opens on totals.
        options = {"level": 1, "hidden": True} if column in segment_block else {}
        ws.set_column(i, i, max(11, min(24, len(column) + 3)),
                      formats.get(style) if style else None, options)
    ws.freeze_panes(1, 1)
    ws.outline_settings(True, False, True, False)

    month_col = daily_range("Month")
    for r, month in enumerate(months):
        row = r + 1
        month_cell = xl_rowcol_to_cell(row, 0, False, True)
        ws.write_string(row, 0, month)
        ws.write_formula(row, col_of["Days"], f"=COUNTIF({month_col},{month_cell})",
                         formats["int"], int(days[month]))
        days_cell = xl_rowcol_to_cell(row, col_of["Days"])
        ws.write_formula(row, col_of["Rooms Available"], f"={days_cell}*Capacity",
                         formats["int"], int(days[month]) * a.capacity)
        avail_cell = xl_rowcol_to_cell(row, col_of["Rooms Available"])

        total_rooms = float(sums.loc[month, "Total Rooms"])
        total_rev = float(sums.loc[month, "Total Revenue"])

        for seg in segments:
            rooms = float(sums.loc[month, f"{seg} Rooms"])
            revenue = float(sums.loc[month, f"{seg} Revenue"])
            ws.write_formula(
                row, col_of[f"{seg} Rooms"],
                f"=SUMIF({month_col},{month_cell},{daily_range(f'{seg} Rooms')})",
                formats["int"], rooms,
            )
            ws.write_formula(
                row, col_of[f"{seg} Revenue"],
                f"=SUMIF({month_col},{month_cell},{daily_range(f'{seg} Revenue')})",
                formats["money0"], revenue,
            )
            rooms_cell = xl_rowcol_to_cell(row, col_of[f"{seg} Rooms"])
            rev_cell = xl_rowcol_to_cell(row, col_of[f"{seg} Revenue"])
            ws.write_formula(row, col_of[f"{seg} ADR"],
                             f"=IF({rooms_cell}>0,{rev_cell}/{rooms_cell},0)",
                             formats["money"], revenue / rooms if rooms else 0.0)
            ws.write_formula(row, col_of[f"{seg} OCC %"],
                             f"=IF({avail_cell}>0,{rooms_cell}/{avail_cell},0)",
                             formats["pct"],
                             rooms / (days[month] * a.capacity) if a.capacity else 0.0)
            total_cell = xl_rowcol_to_cell(row, col_of["Total Rooms"])
            ws.write_formula(row, col_of[f"{seg} Mix %"],
                             f"=IF({total_cell}>0,{rooms_cell}/{total_cell},0)",
                             formats["pct"], rooms / total_rooms if total_rooms else 0.0)

        ws.write_formula(row, col_of["Total Rooms"],
                         f"=SUMIF({month_col},{month_cell},{daily_range('Total Rooms')})",
                         formats["int"], total_rooms)
        ws.write_formula(row, col_of["Total Revenue"],
                         f"=SUMIF({month_col},{month_cell},{daily_range('Total Revenue')})",
                         formats["money0"], total_rev)
        tr = xl_rowcol_to_cell(row, col_of["Total Rooms"])
        tv = xl_rowcol_to_cell(row, col_of["Total Revenue"])
        ws.write_formula(row, col_of["Total ADR"], f"=IF({tr}>0,{tv}/{tr},0)",
                         formats["money"], total_rev / total_rooms if total_rooms else 0.0)
        ws.write_formula(row, col_of["OCC %"], f"=IF({avail_cell}>0,{tr}/{avail_cell},0)",
                         formats["pct"],
                         total_rooms / (days[month] * a.capacity) if a.capacity else 0.0)
        ws.write_formula(row, col_of["RevPAR"], f"=IF({avail_cell}>0,{tv}/{avail_cell},0)",
                         formats["money"],
                         total_rev / (days[month] * a.capacity) if a.capacity else 0.0)

        # --- last year, and the variances against it ------------------------- #
        prior = comparison.iloc[r]
        ws.write_string(row, col_of["LY Month"], str(prior["LY Month"]))
        for head, source, fmt in (("LY Rooms", "LY Rooms", "int"),
                                  ("LY ADR", "LY ADR", "money"),
                                  ("LY Revenue", "LY Revenue", "money0"),
                                  ("LY OCC %", "LY OCC %", "pct")):
            value = prior[source]
            if pd.isna(value):
                ws.write_blank(row, col_of[head], None, formats[fmt])
            else:
                ws.write_number(row, col_of[head],
                                float(value) / 100 if fmt == "pct" else float(value),
                                formats[fmt])

        for head, current_col, prior_col in (("Rooms vs LY %", "Total Rooms", "LY Rooms"),
                                             ("ADR vs LY %", "Total ADR", "LY ADR"),
                                             ("Revenue vs LY %", "Total Revenue", "LY Revenue")):
            cur = xl_rowcol_to_cell(row, col_of[current_col])
            pri = xl_rowcol_to_cell(row, col_of[prior_col])
            variance = prior[head]
            ws.write_formula(row, col_of[head], f'=IF({pri}>0,{cur}/{pri}-1,"")',
                             formats["pct"],
                             0.0 if pd.isna(variance) else float(variance) / 100)
        occ_cell = xl_rowcol_to_cell(row, col_of["OCC %"])
        ly_occ_cell = xl_rowcol_to_cell(row, col_of["LY OCC %"])
        occ_gap = prior["OCC pts vs LY"]
        ws.write_formula(row, col_of["OCC vs LY"], f"={occ_cell}-{ly_occ_cell}",
                         formats["pct"], 0.0 if pd.isna(occ_gap) else float(occ_gap) / 100)

    return columns, months


# --------------------------------------------------------------------------- #

def build_workbook(
    daily: pd.DataFrame,
    model: ForecastModel,
    assumptions: Assumptions,
    capacity: int,
    fiscal_start_month: int,
    all_segments: list[str],
    notes: list[str] | None = None,
    derived: pd.DataFrame | None = None,
) -> bytes:
    """Return the finished .xlsx as bytes, ready for st.download_button."""
    buffer = io.BytesIO()
    writer = pd.ExcelWriter(buffer, engine="xlsxwriter", datetime_format="yyyy-mm-dd")
    book = writer.book
    formats = _formats(book)

    hist_annual = hist.annual(daily, capacity, fiscal_start_month, all_segments)
    bases = growth.all_bases(hist_annual, all_segments)
    values = expand(model, assumptions)

    forecast = pd.DataFrame({"Date": model.dates})
    forecast["Month"] = model.dates.to_period("M").astype(str)
    for column in values:
        forecast[column] = values[column]
    comparison = compare_to_history(forecast, daily, capacity, model.segments)

    blocks: list[tuple[str, pd.DataFrame, object]] = []
    if derived is not None:
        blocks.append(("Rates in use, derived from the basis on Assumptions", derived, None))
    bases_index = len(blocks)
    blocks.append(("Every basis compared (% a year)", bases, growth.BASES))
    for metric in ("Rooms", "ADR"):
        block = growth.yearly_growth(hist_annual, all_segments, metric)
        blocks.append(
            (f"{metric}: year-over-year % by segment", block,
             [c for c in block.columns if c != "Segment"])
        )

    _write_assumptions(
        book, formats, assumptions, capacity, model.segments, model.weights,
        derived, block_rows(blocks)[bases_index], notes or [],
    )
    daily_columns = _write_daily(book, formats, model, assumptions, values)
    _write_monthly(book, formats, model, assumptions, daily_columns, values, comparison)

    history_daily = daily.copy()
    history_daily["Date"] = pd.to_datetime(history_daily["Date"]).dt.date
    _write_values(writer, "History Day by Day", history_daily, formats)
    _write_values(writer, "History Monthly", hist.monthly(daily, capacity, all_segments), formats)
    _write_values(writer, "History Annual", hist_annual, formats)

    # One growth sheet rather than four. Each block was a different arrangement
    # of the same year-over-year numbers, and a reader comparing them was paging
    # between tabs to do it.
    _write_blocks(book, formats, GROWTH_SHEET, blocks)

    writer.close()
    buffer.seek(0)
    return buffer.getvalue()
