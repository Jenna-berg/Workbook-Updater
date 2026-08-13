import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
import io
import csv
import re
import collections
import zipfile
import datetime
import hashlib
import json
import bcrypt
from xml.sax.saxutils import escape
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# ── CSV parsing ───────────────────────────────────────────────────────────────

MONTH_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
DAILY_RE = re.compile(r"^\d{1,2}[/-]\d{1,2}[/-]\d{4}")


def is_formula(value) -> bool:
    return isinstance(value, str) and value.strip().startswith("=")


def is_datelike(value) -> bool:
    return isinstance(value, (datetime.datetime, datetime.date))


def parse_csv(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str, encoding="utf-8-sig")
    df = raw.iloc[2:].reset_index(drop=True)
    df.columns = range(df.shape[1])
    return df


def classify_row(date_str: str):
    """Return ('daily', date) | ('monthly', (year, month)) | (None, None)"""
    if not isinstance(date_str, str):
        return None, None
    date_str = date_str.strip()
    if DAILY_RE.match(date_str):
        raw = date_str[:10].replace("/", "-")
        try:
            d = datetime.datetime.strptime(raw, "%m-%d-%Y").date()
            return "daily", d
        except ValueError:
            return None, None
    parts = date_str.split()
    if len(parts) == 2 and parts[0][:3].lower() in MONTH_ABBR:
        try:
            month = MONTH_ABBR[parts[0][:3].lower()]
            year = int(parts[1])
            return "monthly", (year, month)
        except ValueError:
            pass
    return None, None


def safe_float(val):
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


# Margaritaville's PMS exports an "Occupancy Statistics" .xlsx instead of the
# standard "Business on the Books" CSV every other hotel uses. Confirmed the
# SAME export (same "DATE PRINTED" timestamp, same values) is used for both
# SR and Forecast, matching how one CSV already feeds ROB/SR/Forecast
# together for every other hotel — so this is ONE parser covering every
# field either flow needs, not a separate one per workbook type. It
# normalizes the export into a DataFrame with the exact same column
# positions as parse_csv() (0=date, 1=Rms Sold, 4=OOO, 5=Room Revenue,
# 6=ADR, 7=Grp PU TY, 8=Grp N/PU TY, 9=Grp Rev TY, 15=Trans count,
# 16=Trans Rev) — mapping confirmed against real exports — so
# STRATEGY_CSV_COLS / build_strategy_change_plan / build_forecast_change_plan
# need no changes at all.
MARGARITAVILLE_SOURCE_FIELDS = {
    "rms sold":     1,   # -> Forecast Rooms Sold (both future & actual)
    "ooo rms":      4,   # -> SR ooo_rms
    "room revenue": 5,   # -> Forecast Revenue (actual/past dates)
    "adr":          6,   # -> Forecast ADR OTB (future dates)
    "grp pkup rms": 7,   # -> SR grp_pu_ty
    "grp rem":      8,   # -> SR grp_npu_ty ("remaining" = not yet picked up)
    "grp rm rev":   9,   # -> SR grp_rev_ty
    "trans rms":    15,  # -> SR otb_trans
    "trans rm rev": 16,  # -> SR trans_rev_ty
}


def parse_margaritaville_source(file_bytes: bytes) -> pd.DataFrame:
    """Parse Margaritaville's 'Occupancy Statistics' PMS export (feeds ROB*/
    SR/Forecast — see MARGARITAVILLE_SOURCE_FIELDS). Detects the header row
    and field columns by their text labels — never by color; the source
    file's color-coding was only for human reference while this mapping was
    being worked out, not something to parse at runtime (this app never uses
    cell color to find targets). Skips 'History Total' / 'Forecasted Total' /
    'Total' summary rows and the trailing filter/timestamp/hotel-name rows at
    the bottom of the sheet (any row whose date column doesn't parse as a
    real date).
    * ROB mapping not wired up yet — pending a small tweak to be confirmed.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.worksheets[0]

    header_row = None
    for r in range(1, min(ws.max_row, 30) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and "history/forecasted" in v.strip().lower():
                header_row = r
                break
        if header_row:
            break
    if header_row is None:
        raise ValueError("Could not find the 'History/Forecasted' header row in the source file.")

    col_for_field = {}
    for c in range(1, ws.max_column + 1):
        label = str(ws.cell(header_row, c).value or "").strip().lower()
        for field_label, dest_col in MARGARITAVILLE_SOURCE_FIELDS.items():
            if label == field_label:
                col_for_field[dest_col] = c
    missing = [label for label, dest_col in MARGARITAVILLE_SOURCE_FIELDS.items() if dest_col not in col_for_field]
    if missing:
        raise ValueError(f"Could not find expected column(s) in source file: {', '.join(missing)}.")

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "").strip()
        if not label or "total" in label.lower():
            continue
        date_val = ws.cell(r, 2).value
        # The top data row is sometimes frozen/pinned in Margaritaville's
        # source (confirmed real case) and can retain a native Excel date
        # value there while every other row's date is plain text — a
        # str-only check silently dropped that one row every time.
        if isinstance(date_val, (datetime.datetime, datetime.date)):
            date_str = date_val.strftime("%m/%d/%Y")
        elif isinstance(date_val, str) and DAILY_RE.match(date_val.strip()):
            date_str = date_val.strip()
        else:
            continue
        row_data = {0: date_str}
        for dest_col, src_col in col_for_field.items():
            row_data[dest_col] = safe_float(ws.cell(r, src_col).value)
        rows.append(row_data)

    if not rows:
        raise ValueError("No daily rows found in source file.")

    max_col = max(max(r.keys()) for r in rows)
    df = pd.DataFrame(rows).reindex(columns=range(max_col + 1))
    return _add_margaritaville_monthly_totals(df)


# Columns build_rob_change_plan reads for a "monthly" row: Revenue, Room
# Nights, Grp PU, Grp N/PU, Grp Rev. Same column positions the standard
# Business on the Books CSV already provides monthly totals for directly —
# Margaritaville's source has no such totals, so they're synthesized here by
# summing the daily rows for each calendar month present in the data.
ROB_MONTHLY_SUM_COLS = [1, 5, 7, 8, 9]


def _add_margaritaville_monthly_totals(df: pd.DataFrame) -> pd.DataFrame:
    """Append one synthetic 'monthly' row (e.g. 'Jul 2026') per calendar
    month present in the daily rows, summing ROB_MONTHLY_SUM_COLS — so
    build_rob_change_plan (which only reads rows classify_row calls
    'monthly') works unchanged, the same way it already does for every other
    hotel's CSV, which provides these totals directly."""
    sums = {}  # (year, month) -> {col: running sum}
    for _, row in df.iterrows():
        date_str = str(row[0]).strip() if row[0] else ""
        kind, d = classify_row(date_str)
        if kind != "daily":
            continue
        key = (d.year, d.month)
        bucket = sums.setdefault(key, {c: 0.0 for c in ROB_MONTHLY_SUM_COLS})
        for c in ROB_MONTHLY_SUM_COLS:
            v = row.get(c)
            if v is not None and not pd.isna(v):
                bucket[c] += v

    if not sums:
        return df

    monthly_rows = []
    for (year, month), bucket in sums.items():
        month_name = datetime.date(year, month, 1).strftime("%b")
        row_data = {0: f"{month_name} {year}"}
        row_data.update(bucket)
        monthly_rows.append(row_data)

    monthly_df = pd.DataFrame(monthly_rows).reindex(columns=df.columns)
    return pd.concat([df, monthly_df], ignore_index=True)


def parse_bob_source(uploaded_file) -> pd.DataFrame:
    """Dispatch on file extension: .csv is the standard Business on the
    Books export every hotel uses; .xlsx is Margaritaville's differently-
    formatted PMS export (SR + Forecast wired up so far — ROB needs a small
    additional tweak once that's confirmed)."""
    file_bytes = uploaded_file.read()
    if uploaded_file.name.lower().endswith(".xlsx"):
        return parse_margaritaville_source(file_bytes)
    return parse_csv(file_bytes)


# ── Hilton portfolio source files ────────────────────────────────────────────
# Two exports feed a Hilton run, and neither can produce the ROB on its own:
#
#   SRP Activity      one sheet covering every Hilton property, stay-level,
#                     identified by 'Property - InnCode'
#   Group Wash        one file PER hotel, group-block level, per occupancy date
#
# Group has to come from the Wash report, not from SRP's own 'convention' SRP
# Type — that flag badly undercounts group at some properties (Kansas City
# September: 270 rooms by SRP against 1,017 by Wash). So the ROB totals are
# assembled as SRP transient + Wash pick-up rather than taken from SRP whole.

WASH_PERM_SEGMENT = "PERM"   # Market Segment marking the airline/crew blocks


def _find_header_row(df_raw, first_col_name, limit=40):
    """Row index of the header inside a raw (header=None) export.

    These exports print a filter block above the table and the number of
    filters varies between pulls — the same report has landed on row 11 and
    row 13 — so the header position must be found, never assumed.
    """
    col0 = df_raw[0].astype(str).str.strip()
    hits = df_raw.index[col0 == first_col_name]
    if len(hits) == 0:
        raise ValueError(f"Could not find a '{first_col_name}' header row in the export")
    return int(hits[0])


def _spread_stay(arrival, nights):
    """Yield each occupancy date of a stay. A stay is booked once but occupies
    a room on every night of its span, and month totals are by occupancy — a
    stay arriving 30 Aug for five nights puts two nights in August and three
    in September, not five in August."""
    for i in range(int(nights)):
        yield (arrival + datetime.timedelta(days=i)).date()


def parse_srp_activity(file_like):
    """SRP Activity export → {inncode: {(year, month): {seg: [nights, revenue]}}}

    seg is 'TOT', 'GRP', 'PRM' or 'TRN' (transient = everything that is
    neither group nor permanent). Only TRN is ultimately used for the ROB,
    since group and permanent come from the Wash report, but the rest is kept
    for reconciliation display.
    """
    raw = pd.read_excel(file_like, sheet_name=0, header=None)
    hr = _find_header_row(raw, "Stay ID")
    df = pd.read_excel(file_like, sheet_name=0, header=hr)
    df = df[df["Stay ID"].notna()]

    out = collections.defaultdict(
        lambda: collections.defaultdict(
            lambda: collections.defaultdict(lambda: [0, 0.0])))
    for rec in df.to_dict("records"):
        nights = int(safe_float(rec.get("Room Nights")) or 0)
        if nights <= 0:
            continue
        arrival = rec.get("Arrival Date")
        if not isinstance(arrival, (datetime.datetime, datetime.date)):
            continue
        per = (safe_float(rec.get("Room Revenue *")) or 0.0) / nights
        is_perm = str(rec.get("MCAT", "")).strip().lower() == "permanent"
        is_grp = str(rec.get("SRP Type", "")).strip().lower() == "convention" and not is_perm
        kind = "PRM" if is_perm else ("GRP" if is_grp else "TRN")
        inn = str(rec.get("Property - InnCode", "")).strip().upper()
        for d in _spread_stay(arrival, nights):
            b = out[inn][(d.year, d.month)]
            b["TOT"][0] += 1;   b["TOT"][1] += per
            b[kind][0] += 1;    b[kind][1] += per
    return out


def parse_group_wash(file_like):
    """Individual Group Wash export → {(year, month): {seg: {...}}}

    seg is 'GRP' (Market Segment != PERM) or 'PRM' (== PERM); each holds
    pu_rooms / pu_rev / av_rooms / av_rev.

    'Pick Up' is what has actually been reserved out of the block and belongs
    in the ROB's column E. 'Available Block' is the unpicked remainder and
    belongs in column G ('not p/u'). Market Segment is the discriminator, not
    Forecast Group — one real property has a block named 'Group_PERM_SMRF'
    whose segment is SMRF, i.e. ordinary group business despite the name.
    """
    raw = pd.read_excel(file_like, sheet_name=0, header=None)
    hr = _find_header_row(raw, "Group Code")
    df = pd.read_excel(file_like, sheet_name=0, header=hr)
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]

    out = collections.defaultdict(
        lambda: collections.defaultdict(
            lambda: {"pu_rooms": 0.0, "pu_rev": 0.0, "av_rooms": 0.0, "av_rev": 0.0}))
    for rec in df.to_dict("records"):
        occ = rec.get("Occupancy Date")
        if not isinstance(occ, (datetime.datetime, datetime.date)):
            continue
        pu = safe_float(rec.get("Pick Up")) or 0.0
        av = safe_float(rec.get("Available Block")) or 0.0
        rate = safe_float(rec.get("Rate")) or 0.0
        seg = "PRM" if str(rec.get("Market Segment", "")).strip().upper() == WASH_PERM_SEGMENT else "GRP"
        b = out[(occ.year, occ.month)][seg]
        b["pu_rooms"] += pu
        b["pu_rev"] += pu * rate
        b["av_rooms"] += av
        b["av_rev"] += av * rate
    return out


# ── IHG portfolio source file ────────────────────────────────────────────────
# "History and Forecast Business Block" — a PDF covering the current month
# end to end, one row per date plus two Subtotal rows and a Total. The rows
# above the first Subtotal are days that have already happened; the rows below
# it are still on the books. That single split drives everything: the ROB
# takes the Total row, the Forecast takes the daily rows on both sides of it,
# and the Strategy Report takes the first Subtotal.
#
# Columns, left to right:
#   Date, Total Occ., Arr. Rms.,
#   Individual: Deduct Rms., Room Revenue, Non-D Rms., Non-Deduct Room Revenue
#   Blocks:     Deduct Rms., Room Revenue, Non-D Rms., Non-Deduct Room Revenue
#   Total Hotel: Occ %, Room Revenue, Average Rate

_IHG_NUM = r"([\d,]+\.?\d*)"
_IHG_ROW = re.compile(
    r"^(?P<label>\d{2}-\d{2}-\d{2}|Subtotal|Total)\s+"
    + r"\s+".join([_IHG_NUM] * 10)
    + r"\s+(?P<occpct>[\d.]+)%\s+" + _IHG_NUM + r"\s+" + _IHG_NUM + r"\s*$"
)
_IHG_FIELDS = ["total_occ", "arr_rms", "ind_rms", "ind_rev", "ind_nond_rms",
               "ind_nond_rev", "blk_rms", "blk_rev", "blk_nond_rms", "blk_nond_rev"]


def _ihg_num(s):
    try:
        return float(str(s).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0


def parse_ihg_history_forecast(file_like):
    """History and Forecast Business Block PDF → parsed rows.

    Returns {report_date, days: [...], subtotals: [...], total: {...}} where
    each day carries date, total_occ, blk_rms, blk_rev, total_rev, avg_rate
    and an `is_past` flag set from the position of the first Subtotal — the
    report prints completed days above it and on-the-books days below.
    """
    import pdfplumber

    text = ""
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"

    report_date = None
    m = re.search(r"(\d{2})-(\d{2})-(\d{2})\s*$", text.split("\n")[0].strip())
    if m:
        mm, dd, yy = (int(x) for x in m.groups())
        report_date = datetime.date(2000 + yy, mm, dd)

    days, subtotals, total = [], [], None
    seen_subtotal = False
    for line in text.split("\n"):
        mt = _IHG_ROW.match(line.strip())
        if not mt:
            continue
        nums = [_ihg_num(g) for g in mt.groups()[1:11]]
        rec = dict(zip(_IHG_FIELDS, nums))
        rec["occ_pct"] = float(mt.group("occpct"))
        rec["total_rev"] = _ihg_num(mt.groups()[-2])
        rec["avg_rate"] = _ihg_num(mt.groups()[-1])
        label = mt.group("label")
        if label == "Subtotal":
            subtotals.append(rec)
            seen_subtotal = True
        elif label == "Total":
            total = rec
        else:
            mm, dd, yy = (int(x) for x in label.split("-"))
            rec["date"] = datetime.date(2000 + yy, mm, dd)
            # Above the first Subtotal = already happened.
            rec["is_past"] = not seen_subtotal
            days.append(rec)

    if total is None:
        raise ValueError(
            "No 'Total' row found — is this the History and Forecast Business "
            "Block report?")
    return {"report_date": report_date, "days": days,
            "subtotals": subtotals, "total": total}


_BOB_MONTHS = {m.upper(): i + 1 for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
     "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"])}


def parse_ihg_business_on_books(file_like):
    """Business on the Books PDF → {report_date, months: {(year, month): {...}}}

    One page per month, running twelve-odd months out from the report date,
    each ending in a Total row. Per month it yields rooms, group_rooms,
    group_rev and total_rev.

    Group rooms is Blk-R + D-Blk. That pair together equals the single
    'Blocks Deduct Rms.' column of the History and Forecast report over the
    same span, which is what lets the two reports be used side by side.

    Note the first month is partial — this report starts at the report date,
    not the 1st — so the current month must come from History and Forecast
    instead, which covers it end to end.
    """
    import pdfplumber

    months, report_date = {}, None
    with pdfplumber.open(file_like) as pdf:
        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            if report_date is None and lines:
                m = re.search(r"(\d{2})-(\d{2})-(\d{2})\s*$", lines[0].strip())
                if m:
                    mm, dd, yy = (int(x) for x in m.groups())
                    report_date = datetime.date(2000 + yy, mm, dd)

            hdr = next((l for l in lines if "Ind-R" in l and "Blk-R" in l), None)
            tot = next((l for l in lines if l.strip().startswith("Total ")), None)
            if not hdr or not tot:
                continue
            mh = re.match(r"^([A-Za-z]{3})\s+(\d{2})\b", hdr.strip())
            if not mh:
                continue
            mon = _BOB_MONTHS.get(mh.group(1).upper())
            if not mon:
                continue
            year = 2000 + int(mh.group(2))

            nums = re.findall(r"[\d,]+\.?\d*", tot)
            if len(nums) < 16:
                continue
            v = [_ihg_num(x) for x in nums]
            # Ind-R Blk-R D-Blk N-Blk T-Rms OOO OOS Turn Avbl Occ%
            # Ind-Revenue Ind-Avg Blk-Revenue Blk-Avg Total-Revenue Total-Avg
            entry = {
                "rooms":       v[4],
                "group_rooms": v[1] + v[2],
                "group_rev":   v[12],
                "total_rev":   v[14],
                # D-Blk on its own is the blocked-but-not-picked-up remainder,
                # which is what the ROB's pink 'Group rooms NPU' column wants.
                # It is also counted inside group_rooms above, where the ROB
                # asks for the whole block.
                "npu_rooms":   v[2],
                "blk_avg":     v[13],
                "days":        [],
            }
            # Per-date rows, used to drive a Forecast for a month other than
            # the current one. Every day here is still on the books — this
            # report starts at the report date, so nothing in it has happened.
            for line in lines:
                dm = re.match(r"^(\d{2})-(\d{2})-(\d{2})\s+[A-Za-z]{3}\s+(.*)$",
                              line.strip())
                if not dm:
                    continue
                dnums = re.findall(r"[\d,]+\.?\d*", dm.group(4))
                if len(dnums) < 16:
                    continue
                dv = [_ihg_num(x) for x in dnums]
                d_mm, d_dd, d_yy = (int(dm.group(i)) for i in (1, 2, 3))
                entry["days"].append({
                    "date":      datetime.date(2000 + d_yy, d_mm, d_dd),
                    "ind_rms":   dv[0],
                    "blk_rms":   dv[1],
                    "npu_rms":   dv[2],
                    "rooms":     dv[4],
                    "ooo":       dv[5],
                    "turn":      dv[7],
                    "ind_rev":   dv[10],
                    "blk_rev":   dv[12],
                    "total_rev": dv[14],
                    "avg_rate":  dv[15],
                })
            months[(year, mon)] = entry
    if not months:
        raise ValueError(
            "No month totals found — is this the Business on the Books report?")
    return {"report_date": report_date, "months": months}


def build_ihg_rob_plan(parsed, ws, as_of=None, bob=None):
    """ROB changes for one IHG hotel from the two IHG PDFs.

    History and Forecast covers the current month end to end and supplies it;
    Business on the Books starts at the report date, so its own current month
    is partial and is deliberately skipped in favour of that. Every later
    month comes from Business on the Books.

    Months already closed are never touched, and neither is any cell holding a
    formula.
    """
    as_of = as_of or parsed.get("report_date") or datetime.date.today()
    blocks = rob_month_blocks(ws)
    changes = [{"row": 4, "col": 5, "label": "As-of date", "month": None,
                "new_value": as_of, "skip_reason": None}]

    def put(row, col, name, value, month, source):
        if row is None:
            return
        changes.append({
            "row": row, "col": col, "label": f"{name} ({source})", "month": month,
            "new_value": value,
            "skip_reason": "formula" if is_formula(ws.cell(row, col).value) else None,
        })

    def emit(labels, month, rooms, rev, grp_rooms, grp_rev, source):
        put(labels.get("revenue"),        5, "Revenue",        round(rev, 2),         month, source)
        put(labels.get("room nights"),    5, "Room Nights",    int(round(rooms)),     month, source)
        put(labels.get("group rms sold"), 5, "Group Rms sold", int(round(grp_rooms)), month, source)
        put(labels.get("group rm rev"),   5, "Group Rm Rev",   round(grp_rev, 2),     month, source)

    def emit_npu(labels, month, m):
        """Column G — the pink 'Group rooms NPU' pair.

        Rooms are D-Blk, the part of the block still unpicked; revenue is that
        at the block's own average rate, which is how the column has been kept
        by hand (Nov: 55 x 112.16 = 6,168.80 to the cent).

        This comes from Business on the Books even for the current month:
        History and Forecast has no unpicked figure, and only the days still
        ahead can have one anyway, which is exactly the span Business on the
        Books starts from.
        """
        npu = m.get("npu_rooms", 0.0)
        put(labels.get("group rms sold"), 7, "Group NPU rms",
            int(round(npu)), month, "BoB")
        put(labels.get("group rm rev"), 7, "Group NPU rev",
            round(npu * m.get("blk_avg", 0.0), 2), month, "BoB")

    cur = blocks.get(as_of.month - 1)
    if cur:
        t = parsed["total"]
        emit(cur, as_of.month, t["total_occ"], t["total_rev"],
             t["blk_rms"], t["blk_rev"], "H&F")
        if bob:
            m = bob["months"].get((as_of.year, as_of.month))
            if m:
                emit_npu(cur, as_of.month, m)

    if bob:
        for (year, month), m in sorted(bob["months"].items()):
            if year != as_of.year or month <= as_of.month:
                continue          # current month comes from H&F; past is closed
            labels = blocks.get(month - 1)
            if not labels:
                continue
            emit(labels, month, m["rooms"], m["total_rev"],
                 m["group_rooms"], m["group_rev"], "BoB")
            emit_npu(labels, month, m)
    return changes


def build_ihg_forecast_plan(parsed, ws, wb=None):
    """Forecast changes for one IHG hotel from the same PDF.

    A day that has already happened is an actual and goes to the ACTUAL block
    as rooms + revenue; a day still ahead is on the books and goes to the OTB
    block as rooms + rate. Which side a day falls on comes from the report's
    own Subtotal split rather than from today's date, so a report pulled for
    an earlier as-of still lands correctly.

    `wb` is needed because an untouched week tab carries its dates as formulas
    pointing at WK1 rather than as literal dates; without the workbook there is
    nothing to resolve them against and no column would match.
    """
    rows = locate_forecast_rows(ws)
    if not rows:
        return []
    otb_row       = rows["otb_rooms_row"]
    adr_row       = rows["adr_otb_row"]
    act_rooms_row = rows["actual_rooms_row"]
    act_rev_row   = rows["actual_revenue_row"]

    col_of = {}
    for date_val, col in build_forecast_date_col_map(
            ws, wb=wb, date_row=rows["date_row"]).items():
        col_of[date_val.date() if isinstance(date_val, datetime.datetime) else date_val] = col

    changes = []

    def put(row, col, label, value):
        if row is None or col is None:
            return
        changes.append({
            "row": row, "col": col, "label": label, "month": None,
            "new_value": value,
            "skip_reason": "formula" if is_formula(ws.cell(row, col).value) else None,
        })

    for d in parsed["days"]:
        col = col_of.get(d["date"])
        if col is None:
            continue
        stamp = d["date"].strftime("%b %d").replace(" 0", " ")
        if d["is_past"]:
            put(act_rooms_row, col, f"{stamp} actual rooms",   int(round(d["total_occ"])))
            put(act_rev_row,   col, f"{stamp} actual revenue", round(d["total_rev"], 2))
        else:
            put(otb_row, col, f"{stamp} OTB rooms", int(round(d["total_occ"])))
            put(adr_row, col, f"{stamp} OTB ADR",   round(d["avg_rate"], 2))

    snap = find_forecast_snapshot_row(ws)
    if snap:
        for d in parsed["days"]:
            col = col_of.get(d["date"])
            if col is None:
                continue
            stamp = d["date"].strftime("%b %d").replace(" 0", " ")
            put(snap, col, f"{stamp} wk/wk rooms", int(round(d["total_occ"])))
    return changes


def detect_ihg_sr_columns(ws):
    """{field: column} for an IHG Strategy Report, read off the header.

    Headings are split across rows 3 and 4 and only the pair is unique — 'TRAN'
    alone is both the LY transient rooms column and the start of the transient
    revenue heading — so the two rows are joined before matching.
    """
    wanted = {
        ("trans", "sold"):     "trans_rms",
        ("group", "sold"):     "group_rms",
        ("prop.", "regrets"):  "regrets",
        ("grp rms", "n/pu"):   "group_npu",
        ("ooo", "rooms"):      "ooo",
        ("tran", "rev ty"):    "trans_rev",
        ("grp", "rev ty"):     "group_rev",
    }
    out = {}
    for c in range(1, 70):
        h3 = str(ws.cell(3, c).value or "").strip().lower()
        h4 = str(ws.cell(4, c).value or "").strip().lower()
        key = wanted.get((h3, h4))
        if key and key not in out:
            out[key] = c
    return out


def build_ihg_sr_plan(hf, bob, ws):
    """Strategy Report changes for one IHG hotel from the two PDFs.

    Each dated row takes its figures from whichever report covers that day:
    History and Forecast for days already behind the Business on the Books
    start, Business on the Books from that date on. The boundary is read from
    the report itself rather than assumed, since it moves with the pull date.

    History and Forecast carries no OOO or turnaway columns, so for the earlier
    days those two are left as they are rather than zeroed.

    Monthly total rows never match a date and so are skipped, which is what
    keeps them from being written over.
    """
    cols = detect_ihg_sr_columns(ws)
    if "trans_rms" not in cols:
        return []

    hf_by_date = {d["date"]: d for d in (hf or {}).get("days", [])}
    bob_by_date = {}
    for m in (bob or {}).get("months", {}).values():
        for d in m.get("days", []):
            bob_by_date[d["date"]] = d
    if not bob_by_date and not hf_by_date:
        return []
    bob_start = min(bob_by_date) if bob_by_date else None

    changes = []

    def put(row, col, label, value, when):
        if col is None or value is None:
            return
        changes.append({
            "row": row, "col": col, "label": f"{when:%b %d} {label}".replace(" 0", " "),
            "month": when.month, "new_value": value,
            "skip_reason": "formula" if is_formula(ws.cell(row, col).value) else None,
        })

    for r in range(5, ws.max_row + 1):
        v = ws.cell(r, 3).value
        if isinstance(v, datetime.datetime):
            day = v.date()
        elif isinstance(v, datetime.date):
            day = v
        else:
            continue                     # blank rows and monthly totals

        if bob_start is not None and day >= bob_start:
            d = bob_by_date.get(day)
            if not d:
                continue
            put(r, cols.get("trans_rms"), "trans sold",  int(round(d["ind_rms"])), day)
            put(r, cols.get("group_rms"), "group sold",
                int(round(d["blk_rms"] + d["npu_rms"])), day)
            put(r, cols.get("group_npu"), "grp n/pu",    int(round(d["npu_rms"])), day)
            put(r, cols.get("ooo"),       "OOO",         int(round(d["ooo"])), day)
            put(r, cols.get("regrets"),   "regrets",     int(round(d["turn"])), day)
            put(r, cols.get("trans_rev"), "trans rev",   round(d["ind_rev"], 2), day)
            put(r, cols.get("group_rev"), "group rev",   round(d["blk_rev"], 2), day)
        else:
            d = hf_by_date.get(day)
            if not d:
                continue
            put(r, cols.get("trans_rms"), "trans sold", int(round(d["ind_rms"])), day)
            put(r, cols.get("group_rms"), "group sold", int(round(d["blk_rms"])), day)
            put(r, cols.get("trans_rev"), "trans rev",  round(d["ind_rev"], 2), day)
            put(r, cols.get("group_rev"), "group rev",  round(d["blk_rev"], 2), day)
            # A day that has already happened can have nothing left unpicked.
            put(r, cols.get("group_npu"), "grp n/pu", 0, day)
            # History and Forecast carries no OOO or turnaway columns, so these
            # keep whatever earlier weeks put there — unless it's text, which
            # would feed a #VALUE! into the row's demand and occupancy formulas.
            for key in ("ooo", "regrets"):
                col = cols.get(key)
                if col and isinstance(ws.cell(r, col).value, str) \
                        and not is_formula(ws.cell(r, col).value):
                    put(r, col, f"{key} (cleared non-numeric)", 0, day)
    return changes


def find_forecast_snapshot_row(ws, search_from=45, search_to=90):
    """Row in the week-over-week pickup grid that belongs to THIS week tab.

    The grid stacks one row per weekly snapshot with a 'Pick UP' row between
    each pair, computed as this snapshot minus the one above. Every row except
    one is a formula reaching back into an earlier week's tab; the row for the
    tab you're on is the one whose column A is '=A2' — a self-reference to its
    own as-of date. That is the only row that should ever be written, and it
    moves down by two with each successive week (WK2 row 57, WK3 59, WK4 61).

    Returns None if no such row is found, so a template that doesn't carry
    this grid is skipped rather than written to at a guessed row.
    """
    for r in range(search_from, search_to):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().replace("+", "").replace(" ", "").upper() == "=A2":
            return r
    return None


def build_ihg_next_month_forecast_plan(bob, ws, target_month, wb=None):
    """Forecast changes for a month other than the current one, from the
    Business on the Books daily rows.

    Everything in that report is still on the books — it starts at the report
    date — so every day goes to the OTB rows and none to the actual rows.
    """
    key = (target_month.year, target_month.month)
    month = (bob or {}).get("months", {}).get(key)
    if not month or not month.get("days"):
        return []

    rows = locate_forecast_rows(ws)
    if not rows:
        return []

    col_of = {}
    for date_val, col in build_forecast_date_col_map(
            ws, wb=wb, date_row=rows["date_row"]).items():
        col_of[date_val.date() if isinstance(date_val, datetime.datetime) else date_val] = col

    changes = []
    for d in month["days"]:
        col = col_of.get(d["date"])
        if col is None:
            continue
        stamp = d["date"].strftime("%b %d").replace(" 0", " ")
        for row, label, value in [
            (rows["otb_rooms_row"], f"{stamp} OTB rooms", int(round(d["rooms"]))),
            (rows["adr_otb_row"],   f"{stamp} OTB ADR",   round(d["avg_rate"], 2)),
        ]:
            changes.append({
                "row": row, "col": col, "label": label, "month": target_month.month,
                "new_value": value,
                "skip_reason": "formula" if is_formula(ws.cell(row, col).value) else None,
            })
    return changes


# ── ROB Update ───────────────────────────────────────────────────────────────

ROB_SHEETS = ["wk one", "wk two", "wk three", "wk four", "wk five", "wk six"]


def find_secondary_col(ws, block_start):
    candidates = []
    for cell in ws[block_start]:
        if cell.column <= 5:
            continue
        if isinstance(cell.value, str) and "variance" in cell.value.strip().lower():
            candidates.append(cell.column)
    return min(candidates) if candidates else None


def build_hilton_rob_plan(srp_months, wash_months, ws, as_of=None):
    """ROB changes for one Hilton hotel from the two Hilton exports.

    srp_months  — parse_srp_activity()[inncode]
    wash_months — parse_group_wash() for that same hotel

    Column E takes what is on the books, column G ('not p/u') the unpicked
    remainder of the group blocks. Totals are assembled rather than lifted
    from SRP: transient comes from SRP, group and permanent from the Wash
    report's pick-up, because SRP's own 'convention' flag undercounts group.

    Rows are located by label, and any cell already holding a formula is left
    alone — weeks get reconciled by hand into expressions like '=294767+55017'
    and those must not be flattened.
    """
    as_of = as_of or datetime.date.today()
    blocks = rob_month_blocks(ws)
    changes = []

    def put(row, col, label, value, month):
        if row is None or value is None:
            return
        if is_formula(ws.cell(row, col).value):
            changes.append({"row": row, "col": col, "label": label, "month": month,
                            "new_value": value, "skip_reason": "formula"})
            return
        changes.append({"row": row, "col": col, "label": label, "month": month,
                        "new_value": value, "skip_reason": None})

    changes.append({"row": 4, "col": 5, "label": "As-of date", "month": None,
                    "new_value": as_of, "skip_reason": None})

    for mi, labels in sorted(blocks.items()):
        month = mi + 1
        if month < as_of.month:
            continue                      # never rewrite a closed month
        key = (as_of.year, month)
        s = srp_months.get(key) or {}
        w = wash_months.get(key) or {}
        trn = s.get("TRN", [0, 0.0])
        g = w.get("GRP", {})
        p = w.get("PRM", {})

        g_pu_r, g_pu_v = g.get("pu_rooms", 0.0), g.get("pu_rev", 0.0)
        p_pu_r, p_pu_v = p.get("pu_rooms", 0.0), p.get("pu_rev", 0.0)

        tot_rooms = trn[0] + g_pu_r + p_pu_r
        tot_rev = trn[1] + g_pu_v + p_pu_v
        if not (tot_rooms or g_pu_r or p_pu_r):
            continue                      # nothing on the books for this month

        L = labels.get
        put(L("revenue"),        5, "Revenue",        round(tot_rev, 2), month)
        put(L("room nights"),    5, "Room Nights",    int(round(tot_rooms)), month)
        put(L("group rms sold"), 5, "Group Rms sold", int(round(g_pu_r)), month)
        put(L("group rm rev"),   5, "Group Rm Rev",   round(g_pu_v, 2), month)
        put(L("group rms sold"), 7, "Group not p/u rms", int(round(g.get("av_rooms", 0.0))), month)
        put(L("group rm rev"),   7, "Group not p/u rev", round(g.get("av_rev", 0.0), 2), month)
        if L("perm rms sold"):
            put(L("perm rms sold"), 5, "Perm Rms Sold", int(round(p_pu_r)), month)
            put(L("perm rm rev"),   5, "Perm Rm Rev",   round(p_pu_v, 2), month)
            put(L("perm rms sold"), 7, "Perm not p/u rms", int(round(p.get("av_rooms", 0.0))), month)
            put(L("perm rm rev"),   7, "Perm not p/u rev", round(p.get("av_rev", 0.0), 2), month)
    return changes


def build_rob_change_plan(df, ws, grp_npu_rev_override: dict = None):
    """grp_npu_rev_override: optional {(year, month): dollar_value} — when
    present for a given month, writes that literal value into the 'Group Not
    P/U rev' secondary-column cell instead of the standard count*ADR formula.
    Used for Margaritaville, whose source data doesn't include a reliable
    Not-P/U room count to build that formula from — instead the value is
    computed elsewhere as the difference between two comparable PMS exports
    (one including not-yet-picked-up group revenue, one excluding it)."""
    today = datetime.date.today()
    current_month = today.month
    current_year = today.year
    changes = []
    # Rows per month block, read off this sheet rather than assumed — hotels
    # with a Permanent-rooms section use 11-row blocks, not 8, and every row
    # below is relative to it.
    block_step = rob_block_step(ws)
    # Last writable row = the final row of the Dec block. For the usual 8-row
    # layout this is 99, so `>= 100` behaves exactly as the old hardcoded
    # guard did; an 11-row sheet correctly allows through row 135.
    max_data_row = 4 + block_step * 12

    # E4 = as-of date
    changes.append({
        "row": 4, "col": 5, "label": "As-of date", "month": None,
        "new_value": today, "skip_reason": None,
    })

    for _, row in df.iterrows():
        date_str = str(row[0]).strip() if row[0] else ""
        kind, info = classify_row(date_str)
        if kind != "monthly":
            continue
        year, month = info
        prev_month = current_month - 1 if current_month > 1 else 12
        prev_year  = current_year if current_month > 1 else current_year - 1
        if year == prev_year and month == prev_month:
            pass  # allow previous month (final numbers come in on the 1st)
        elif year != current_year or month < current_month:
            continue

        month_index = month - 1
        block_start = 4 + block_step * month_index

        rev     = safe_float(row[5])
        rms     = safe_float(row[1])
        grp_pu  = safe_float(row[7])
        grp_npu = safe_float(row[8])
        grp_rvn = safe_float(row[9])

        grp_sold = (grp_pu or 0) + (grp_npu or 0) if grp_pu is not None and grp_npu is not None else None
        sec_col = find_secondary_col(ws, block_start)

        entries = [
            (block_start + 1, 5, "Revenue",        rev,       False),
            (block_start + 2, 5, "Room Nights",     rms,       False),
            (block_start + 4, 5, "Group Rms Sold",  grp_sold,  False),
            (block_start + 5, 5, "Group Rm Rev",    grp_rvn,   False),
        ]
        if sec_col:
            npu_row    = block_start + 4
            entries.append((npu_row, sec_col, "Group Not P/U rooms", grp_npu, False))

            override_val = grp_npu_rev_override.get((year, month)) if grp_npu_rev_override else None
            if override_val is not None:
                entries.append((npu_row + 1, sec_col, "Group Not P/U rev (computed)", override_val, False))
            else:
                from openpyxl.utils import get_column_letter
                sec_letter = get_column_letter(sec_col)
                adr_row    = block_start + 6
                npu_formula = f"={sec_letter}{npu_row}*E{adr_row}"
                entries.append((npu_row + 1, sec_col, "Group Not P/U rev (formula)", npu_formula, True))

        for r, c, label, val, is_formula_write in entries:
            skip = None
            if r >= max_data_row:
                skip = f"row≥{max_data_row}"
            elif not is_formula_write and is_formula(ws.cell(r, c).value):
                skip = "formula"
            changes.append({"row": r, "col": c, "label": label, "month": month,
                             "new_value": val, "skip_reason": skip})

    return changes


def compute_grp_npu_rev_override(df, npu_compare_df):
    """Margaritaville ROB only: npu_compare_df is the same source format but
    from a second PMS export that includes not-yet-picked-up group revenue
    (df itself is the export that excludes it — confirmed the smaller of the
    two is used everywhere else). The difference per month is the dollar
    value for the "Group Not P/U rev" bright-green box on the ROB. Returns
    None if either input is missing (i.e. every hotel except Margaritaville)."""
    if npu_compare_df is None or df is None:
        return None

    def _monthly_col5(source_df):
        out = {}
        for _, row in source_df.iterrows():
            kind, info = classify_row(str(row[0]).strip() if row[0] else "")
            if kind == "monthly":
                out[info] = safe_float(row[5])
        return out

    main_sums    = _monthly_col5(df)
    compare_sums = _monthly_col5(npu_compare_df)
    return {
        key: compare_sums[key] - main_sums[key]
        for key in main_sums
        if key in compare_sums and main_sums[key] is not None and compare_sums[key] is not None
    }


def apply_rob_changes(wb, sheet_name, changes):
    ws = wb[sheet_name]
    for ch in changes:
        if ch["skip_reason"]:
            continue
        ws.cell(ch["row"], ch["col"]).value = ch["new_value"]


DONE_TAB_RGB = "FF00B050"  # green — set by color_tab_done() when a week is genuinely complete
DONE_TAB_HEX = "00B050"    # same green, without the alpha channel


def _is_done_color(rgb_value) -> bool:
    """True if rgb_value is our green 'done' marker. Matches on the trailing
    6 hex digits (case-insensitive) so it recognizes a tab colored directly in
    Excel's own Tab Color picker (which often stores 6-digit RGB with no alpha
    prefix) as well as ones this app set (8-digit ARGB) — a real week that a
    hotel had already filled in and marked green by hand was being treated as
    'not done' by a strict 8-char match, so the app kept re-picking it.
    """
    return isinstance(rgb_value, str) and rgb_value[-6:].upper() == DONE_TAB_HEX


_MONTH_LABELS = {m.lower(): i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}
_MONTH_LABELS.update({m.lower(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"])})
ROB_DEFAULT_BLOCK_STEP = 8


def rob_block_step(ws, default=ROB_DEFAULT_BLOCK_STEP):
    """Rows per month block on a ROB week tab, read off the sheet.

    Most ROBs use 8 rows per month (Revenue/Room Nights/ADR/Group x3/Pickup),
    but hotels with a Permanent-rooms section (airline crew contracts) use 11
    — they carry three extra rows (Perm Rms Sold / Perm Rm Rev / Perm ADR).
    Assuming 8 everywhere makes every row calculation land inside the wrong
    month on those sheets: on an 11-row ROB the "August" block computes to
    row 60, which is actually inside June.

    Derived from the Jan/Feb month labels in column A and cross-checked
    against Mar, so a sheet that doesn't look like a ROB falls back to the
    historical constant rather than guessing.
    """
    found = {}
    for r in range(1, 200):
        v = ws.cell(r, 1).value
        if not isinstance(v, str):
            continue
        idx = _MONTH_LABELS.get(v.strip().lower())
        if idx is not None and idx not in found:
            found[idx] = r
    if 0 in found and 1 in found:
        step = found[1] - found[0]
        if 2 in found and found[2] - found[1] != step:
            return default            # spacing isn't uniform — don't trust it
        if 4 <= step <= 20:
            return step
    return default


def _rob_week_taken_reason(ws, block_start):
    """Why this week tab is unavailable, or None if it's free to write to.

    Two things this has to get right, both confirmed against real files:

    A zero is not data. These templates render empty currency cells as "$ -"
    (a literal 0), so testing `isinstance(v, (int, float))` alone counts an
    untouched month as filled and silently skips the tab.

    A formula usually IS data. Weeks are routinely reconciled by hand into an
    expression — real examples are '=294767+55017+2028' and
    '=320153-11611+6135' sitting in the Revenue cell of an already-completed
    week. Treating only literal numbers as data marks those weeks available
    and overwrites the reconciliation. The exception is a bare cross-sheet
    reference like ="wk one"!E82, which is the template's own placeholder
    mirroring week one rather than this week's figures.
    """
    tc = ws.sheet_properties.tabColor
    if tc is not None and _is_done_color(getattr(tc, "rgb", None)):
        return "tab marked done (green)"
    rev = ws.cell(block_start + 1, 5).value
    rms = ws.cell(block_start + 2, 5).value

    def filled(v):
        if isinstance(v, bool) or v is None:
            return False
        if isinstance(v, (int, float)):
            return v != 0
        if is_formula(str(v)):
            return "!" not in str(v)   # cross-sheet mirror = still a placeholder
        return False

    if filled(rev) or filled(rms):
        return (f"already has this month's data "
                f"(E{block_start+1}={rev!r}, E{block_start+2}={rms!r})")
    return None


def rob_month_blocks(ws):
    """{month_index: {row_label_lower: row_number}} for a ROB week tab.

    Locates every row by its column-A label instead of by a fixed offset from
    the block start. Offsets are not safe to assume even within one workbook:
    a real file has 'wk one' on 11-row blocks with Perm at offsets 7/8 while
    'wk two'..'wk six' are 12-row with Perm at 8/9, so an offset that is right
    on one tab writes into the wrong row on the next.

    Labels are lower-cased and whitespace-collapsed, e.g. 'revenue',
    'room nights', 'group rms sold', 'group rm rev', 'perm rms sold',
    'perm rm rev', 'adr', 'pickup wow'.
    """
    starts = {}
    for r in range(1, 260):
        v = ws.cell(r, 1).value
        if not isinstance(v, str):
            continue
        s = v.strip()
        # month header cells hold just the month name — guard against a long
        # sentence that merely begins with one
        if len(s) <= 12 and s[:3].lower() in _MONTH_LABELS:
            starts.setdefault(_MONTH_LABELS[s[:3].lower()], r)
    ordered = sorted(starts.items(), key=lambda kv: kv[1])
    if not ordered:
        return {}
    span = (ordered[1][1] - ordered[0][1]) if len(ordered) > 1 else ROB_DEFAULT_BLOCK_STEP
    out = {}
    for n, (mi, r0) in enumerate(ordered):
        end = ordered[n + 1][1] if n + 1 < len(ordered) else r0 + span
        labels = {}
        for r in range(r0 + 1, end):
            v = ws.cell(r, 1).value
            if isinstance(v, str) and v.strip():
                labels.setdefault(re.sub(r"\s+", " ", v.strip().lower()), r)
        out[mi] = labels
    return out


def rob_week_status(wb, sheet_names):
    """[(sheet_name, reason_or_None), ...] — why each week tab was passed over.
    Purely diagnostic; drives the 'why did it skip a week' caption in the UI.
    """
    month = datetime.date.today().month
    out = []
    for name in sheet_names:
        ws = wb[name]
        block_start = 4 + rob_block_step(ws) * (month - 1)
        out.append((name, _rob_week_taken_reason(ws, block_start)))
    return out


def first_uncolored_sheet(wb, sheet_names):
    """Return the first ROB week tab that is neither marked done (our green)
    nor already holding real data in this month's block.

    Color alone isn't enough in either direction: a master template can carry
    its own unrelated baked-in tab color on a sheet that's never been touched
    (so "any color = done" produces false positives), while a real, already-
    filled week can fail to carry our exact green (so "only our green = done"
    produces false negatives — confirmed on a real hotel: week one already had
    this month's Revenue/Room Nights filled in but got silently re-picked and
    overwritten because its tab color didn't match). Checking actual data in
    the cells this month's update is about to write closes that gap.
    """
    month = datetime.date.today().month
    for name in sheet_names:
        ws = wb[name]
        block_start = 4 + rob_block_step(ws) * (month - 1)
        if _rob_week_taken_reason(ws, block_start) is None:
            return name
    return sheet_names[-1]  # fallback: last sheet


def first_unhighlighted_forecast_sheet(wb, sheet_names):
    """Return the first Forecast week tab that is neither marked done (our
    green) nor already holding real OTB Rooms Sold data.

    Tab color alone can't be trusted in either direction here: hotels color-
    code Forecast tabs by hand with whatever color they like, but the SAME
    color can also be a stray artifact baked into a never-used master
    template — confirmed on real files: Hampton's untouched master has ALL 9
    week tabs pre-colored, and Provincetown Brass's untouched master has
    FCST-WK1 pre-colored magenta despite being completely blank. Checking the
    OTB Rooms Sold row for actual filled-in numbers is the reliable signal;
    our own exact green is kept as a secondary check for weeks we marked done
    ourselves but that ended up with no literal numbers written (e.g. every
    cell in that row happened to be formula-protected).
    """
    for name in sheet_names:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        tc = ws.sheet_properties.tabColor
        if tc is not None and _is_done_color(getattr(tc, "rgb", None)):
            continue
        rows = locate_forecast_rows(ws)
        if not rows:
            return name  # can't verify — treat as available rather than guess
        otb_row = rows["otb_rooms_row"]
        # Check the actual date-mapped columns for this sheet rather than a
        # hardcoded 2-9 range — confirmed real case (Tybee): a filled week
        # got silently treated as available and overwritten, most likely
        # because its real data sat outside that fixed range. Falls back to
        # the old range if the date row itself can't be read. safe_float
        # (not a strict isinstance check) also catches numbers stored as
        # text, which a strict int/float check would miss entirely.
        col_map = build_forecast_date_col_map(ws, ws.parent, date_row=rows["date_row"])
        data_cols = col_map.values() if col_map else range(2, 10)
        has_data = any(safe_float(ws.cell(otb_row, c).value) is not None for c in data_cols)
        if has_data:
            continue
        return name
    return sheet_names[-1]  # fallback: last sheet


def first_undone_strategy_sheet(wb, sheet_names):
    """Return the first Strategy Report week tab that is neither marked done
    (our green) nor already holding real OTB TY Trans data.

    Same reasoning as first_unhighlighted_forecast_sheet: tab color alone
    can't be trusted — confirmed on real files that the identical purple
    (FF9900FF) marks a genuinely completed Surfside week AND sits untouched
    on Wolfboro's never-used master template. Check actual filled data as the
    reliable signal; our own exact green is a secondary check for weeks we
    marked done ourselves but that ended up with no literal numbers written.
    """
    for name in sheet_names:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        tc = ws.sheet_properties.tabColor
        if tc is not None and _is_done_color(getattr(tc, "rgb", None)):
            continue
        col_map = detect_strategy_columns(ws)
        otb_col = col_map.get("otb_trans")
        if not otb_col:
            return name  # can't verify — treat as available rather than guess
        has_data = any(isinstance(ws.cell(r, otb_col).value, (int, float)) for r in range(5, 15))
        if has_data:
            continue
        return name
    return sheet_names[-1]  # fallback: last sheet


def color_tab_done(wb, sheet_name):
    """Mark a sheet tab green to indicate it has been completed."""
    from openpyxl.styles.colors import Color
    wb[sheet_name].sheet_properties.tabColor = Color(rgb=DONE_TAB_RGB)


def clear_tab_colors(wb, sheet_names):
    """Reset tab color to none for every listed sheet. New-month setup copies
    the master/prior file, which can carry over a stale 'done' tab color —
    without this, a week that hasn't been touched yet can look completed and
    get skipped by the 'first uncolored sheet' auto-detect."""
    for name in sheet_names:
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = None


def strip_tables(wb):
    """Remove Excel table definitions to prevent openpyxl save corruption."""
    for ws in wb.worksheets:
        ws.tables.clear()


def apply_portfolio_plans(svc, jobs, undo_key):
    """Write, mark done, upload — shared by the Hilton and IHG runs.

    Each job is {key, file_id, file_name, wb_bytes, sheet, changes}. The
    pre-write value of every cell is snapshotted first so the run can be undone
    in one step, and the clean pre-write bytes are kept alongside so the undo
    rebuilds from the original rather than from what was just written.

    A macro-enabled workbook is reopened with keep_vba so uploading it back
    doesn't strip the macros.
    """
    saved, errors, snapshot = [], [], {}
    for job in jobs:
        try:
            keep_vba = str(job.get("file_name", "")).lower().endswith(".xlsm")
            wb = openpyxl.load_workbook(io.BytesIO(job["wb_bytes"]),
                                        data_only=False, keep_vba=keep_vba)
            ws = wb[job["sheet"]]
            writes = [c for c in job["changes"] if not c.get("skip_reason")]
            prev_tab = ws.sheet_properties.tabColor
            snapshot[job["key"]] = {
                "file_id":   job["file_id"],
                "file_name": job["file_name"],
                "wb_bytes":  job["wb_bytes"],
                "sheet":     job["sheet"],
                "keep_vba":  keep_vba,
                # The tab colour is part of what a run changes — undoing the
                # cells but leaving the tab green would leave the week looking
                # complete and get it skipped on the next run.
                "tab_rgb":   getattr(prev_tab, "rgb", None) if prev_tab is not None else None,
                "cells":     {(job["sheet"], c["row"], c["col"]):
                              ws.cell(c["row"], c["col"]).value for c in writes},
            }
            for c in writes:
                ws.cell(c["row"], c["col"]).value = c["new_value"]
                if isinstance(c["new_value"], (datetime.date, datetime.datetime)):
                    ws.cell(c["row"], c["col"]).number_format = "m/d/yyyy"
            color_tab_done(wb, job["sheet"])
            strip_tables(wb)
            out = io.BytesIO()
            wb.save(out)
            drive_upload(svc, job["file_id"], out.getvalue(), job["file_name"])
            saved.append(f'{job["file_name"]} → {job["sheet"]} ({len(writes)} cells)')
        except Exception as e:
            errors.append(f'{job["key"]}: {e}')
    if snapshot:
        st.session_state[undo_key] = snapshot
    return saved, errors


def undo_portfolio_plans(svc, undo_key):
    """Put every snapshotted cell back and re-upload."""
    snapshot = st.session_state.get(undo_key) or {}
    if not snapshot:
        return [], ["Nothing to undo — no snapshot from this session."]
    saved, errors = [], []
    for key, info in snapshot.items():
        try:
            wb = openpyxl.load_workbook(io.BytesIO(info["wb_bytes"]),
                                        data_only=False,
                                        keep_vba=info.get("keep_vba", False))
            ws = wb[info["sheet"]]
            for (_sheet, row, col), original in info["cells"].items():
                ws.cell(row, col).value = original
            rgb = info.get("tab_rgb")
            if rgb:
                from openpyxl.styles.colors import Color
                ws.sheet_properties.tabColor = Color(rgb=rgb)
            else:
                ws.sheet_properties.tabColor = None
            strip_tables(wb)
            out = io.BytesIO()
            wb.save(out)
            drive_upload(svc, info["file_id"], out.getvalue(), info["file_name"])
            saved.append(info["file_name"])
        except Exception as e:
            errors.append(f"{key}: {e}")
    if not errors:
        st.session_state.pop(undo_key, None)
    return saved, errors


# ── Strategy Report ───────────────────────────────────────────────────────────

STRATEGY_SHEETS = ["WKONE", "WKTWO", "WKTHREE", "WKFOUR", "WKFIVE"]
FORECAST_SHEETS = ["FCST-WK1", "FCST-WK2", "FCST-WK3", "FCST-WK4",
                   "FCST-WK5", "FCST-WK6", "FCST-WK7", "FCST-WK8", "FCST-WK9"]

# CSV column index for each field (0-based) — source data never changes
STRATEGY_CSV_COLS = {
    "otb_trans":    (15, "OTB TY Trans (Indiv Count)"),
    "grp_pu_ty":    ( 7, "GRP PU TY"),
    "grp_npu_ty":   ( 8, "GRP N/PU TY"),
    "ooo_rms":      ( 4, "OOO RMS"),
    "trans_rev_ty": (16, "Trans Rev TY"),
    "grp_rev_ty":   ( 9, "Grp Rev TY"),
}

# Each field: list of (row3_keyword, row4_keyword) pairs to try in order.
# Match = both keywords found (case-insensitive) in their respective rows of that column.
# A None keyword means "don't check that row."
# "!WORD" suffix(es) on a keyword mean the combined headers must NOT contain WORD.
STRATEGY_FIELD_PATTERNS = {
    # ── TY columns (written from CSV) ──────────────────────────────────────────
    "otb_trans":       [("OTB TY", "TRANS"),            ("TRANS!LY", "SOLD!LY")],
    "grp_pu_ty":       [("GRP PU", "TY!LY"),            ("GROUP!LY", "SOLD!LY")],
    "grp_npu_ty":      [("GRP N/PU", "TY!LY"),          ("GRP RMS", "N/PU"),       ("N/PU!LY", None)],
    "ooo_rms":         [("OOO", None)],
    "trans_rev_ty":    [("TY TRANS", "REV"),             ("TRAN!LY", "REV TY")],
    "grp_rev_ty":      [("GRP TY", "REV"),               ("GRP!LY!N/PU", "REV TY")],
    "otb_lst_wk":      [("OTB", "LST WEK"),               ("OTB", "LST WK"),         ("OTB", "LAST WK"), ("OTB LST", None)],
    "casino_ballroom": [("CASINO", "BALLROOM!LY"),       ("CASINO BALLROOM", None)],
    # ── LY columns (written from last year's SR) ───────────────────────────────
    "otb_ly_trans":    [("LY", "TRAN"),                  ("OTB LY", "TRANS"),       ("TRANS!TY", "SOLD!TY"), ("LY", "TRANS!TY")],
    "grp_pu_ly":       [("LY", "GRP"),                   ("GRP PU", "LY"),          ("GROUP!TY", "LY"),        ("GRP PU LY", None)],
    "grp_npu_ly":      [("GRP N/PU", "LY"),              ("N/PU LY", None),         ("GRP RMS", "LY")],
    "trans_rev_ly":    [("LY TRANS", "REV"),             ("TRAN!TY", "REV LY"),     ("LY", "TRANS REV")],
    "grp_rev_ly":      [("GRP LY", "REV"),               ("GRP!TY!N/PU", "REV LY")],
    "grp_npu_rev_ly":  [("GRP N/PU", "REV LY"),         ("N/PU LY", "REV"),        ("N/PU", "REV LY")],
    "casino_ballroom_ly": [("CASINO", "BALLROOM!TY"),    ("CASINO", "LY")],
}

# Maps LY destination field → TY source field in last year's SR
LY_FROM_TY = {
    "otb_ly_trans":    "otb_trans",
    "grp_pu_ly":       "grp_pu_ty",
    "grp_npu_ly":      "grp_npu_ty",
    "trans_rev_ly":    "trans_rev_ty",
    "grp_rev_ly":      "grp_rev_ty",
    "grp_npu_rev_ly":  "grp_npu_ty",  # source is GRP N/PU TY
    "casino_ballroom_ly": "casino_ballroom",
}

def _kw_matches(cell_val, keyword, r3_val, r4_val):
    """Check if keyword matches cell_val.
    Supports !WORD suffixes — the combined headers must NOT contain those words.
    e.g. 'TRAN!LY!ADR' matches if cell contains 'TRAN' and neither header contains 'LY' or 'ADR'.
    """
    parts = keyword.split("!")
    kw = parts[0].strip()
    excludes = [p.strip().upper() for p in parts[1:]]
    if kw and kw.upper() not in str(cell_val or "").upper():
        return False
    combined = (str(r3_val or "") + " " + str(r4_val or "")).upper()
    for excl in excludes:
        if excl in combined:
            return False
    return True


def detect_strategy_columns(ws):
    """Scan rows 3+4 of THIS sheet and return {field_key: col_index} for each
    field. Every sheet is re-scanned independently — never assume two sheets
    (even in the same workbook/hotel) share column positions. Week-1 vs
    week-2+ tabs can differ (e.g. an extra pickup-tracking column shifts
    everything after it), so a value pinned from one sheet can silently be
    wrong on another.

    If row 4 is blank under a TY/LY column, infer it from context (e.g., blank
    under "OTB TY" → assume "TRANS", blank under "GRP PU TY" → assume "TY").
    """
    max_col = ws.max_column
    # Build lookup: col → (r3_text, r4_text)
    headers = {}
    for c in range(1, max_col + 1):
        r3_val = str(ws.cell(3, c).value or "").strip()
        r4_val = str(ws.cell(4, c).value or "").strip()

        # If row 4 is blank or "None", infer from row 3 context
        if (not r4_val or r4_val == "None") and r3_val:
            if "OTB TY" in r3_val:
                r4_val = "TRANS"  # OTB TY column always has TRANS
            elif "TY" in r3_val and "LY" not in r3_val:
                r4_val = "TY"  # Other TY columns infer "TY"
            elif "LY" in r3_val:
                r4_val = "LY"  # LY columns infer "LY"

        headers[c] = (r3_val, r4_val)

    col_map = {}
    for field, patterns in STRATEGY_FIELD_PATTERNS.items():
        found = None
        for r3_kw, r4_kw in patterns:
            for c, (r3v, r4v) in headers.items():
                # Match each keyword against the combined row3+row4 text,
                # not just its "own" row — confirmed real case: a header
                # like "OTB TY TRANS" doesn't always split the same way
                # across the two rows on every sheet (e.g. "OTB" / "TY
                # TRANS" instead of "OTB TY" / "TRANS"), and a strict
                # own-row-only check missed it entirely.
                combined = f"{r3v} {r4v}"
                r3_ok = r3_kw is None or _kw_matches(combined, r3_kw, r3v, r4v)
                r4_ok = r4_kw is None or _kw_matches(combined, r4_kw, r3v, r4v)
                if r3_ok and r4_ok and (r3_kw or r4_kw):
                    found = c
                    break
            if found:
                break
        if found:
            col_map[field] = found
        else:
            col_map[field] = None  # will surface as a warning, not a crash

    return col_map


_CROSS_SHEET_REF_RE = re.compile(r"^=(?:'([^']+)'|([A-Za-z0-9_]+))!\$?([A-Z]+)\$?(\d+)$")
_LOCAL_OFFSET_RE    = re.compile(r"^=\$?([A-Z]+)\$?(\d+)([+-]\d+)$")


def _resolve_formula_date(wb, sheet_name, formula, _depth=0):
    """Resolve a date-continuation formula to an actual date without a real
    spreadsheet engine — handles the two real patterns this app's SR
    templates use: a cross-sheet single-cell reference (e.g. '=WKONE!C5',
    used by every non-WKONE week tab to mirror WKONE's own calendar
    column — the same "other weeks reference wk one" pattern ROB uses) and
    a local same-column +N day increment (e.g. '=C5+1'). Recurses through
    a chain (a local offset can itself point at a cross-sheet ref and vice
    versa) with a depth cap as a cycle guard. Confirmed real case:
    Provincetown Crowne Pointe's WKTHREE date column is 100% cross-sheet
    references with no literal cell of its own anywhere — a resolver that
    only trusted local '=prevRow+1' chains, requiring a literal anchor
    first, could never resolve a single date from it and always came back
    empty, blocking all CSV data from being written.
    """
    if _depth > 20 or not isinstance(formula, str) or not formula.startswith("="):
        return None
    formula = formula.strip()

    m = _CROSS_SHEET_REF_RE.match(formula)
    if m:
        target_sheet = m.group(1) or m.group(2)
        if target_sheet not in wb.sheetnames:
            return None
        col_num = column_index_from_string(m.group(3))
        row_num = int(m.group(4))
        val = wb[target_sheet].cell(row_num, col_num).value
        if isinstance(val, datetime.datetime):
            return val.date()
        if isinstance(val, datetime.date):
            return val
        return _resolve_formula_date(wb, target_sheet, val, _depth + 1)

    m = _LOCAL_OFFSET_RE.match(formula)
    if m:
        col_num = column_index_from_string(m.group(1))
        row_num = int(m.group(2))
        offset  = int(m.group(3))
        val = wb[sheet_name].cell(row_num, col_num).value
        if isinstance(val, datetime.datetime):
            base = val.date()
        elif isinstance(val, datetime.date):
            base = val
        else:
            base = _resolve_formula_date(wb, sheet_name, val, _depth + 1)
            if base is None:
                return None
        return base + datetime.timedelta(days=offset)

    return None


def detect_date_column(ws, wb=None):
    """Find the column whose data rows (5+) contain the earliest daily dates —
    i.e. the column that maps to each row's actual calendar date.
    Scans cols 1-10 only (dates are always on the left side).

    A sheet commonly has a Last Year date column right next to the This Year
    one (e.g. col 1 = LY dates, col 3 = TY dates), both starting on the 1st
    of the same month in different years — under a pure "most consecutive,
    prefer earliest day-of-month" score they tie exactly, and confirmed on a
    real file (Anchor In) the tie silently kept the LY column, so the whole
    date-to-row map was built a year off and almost nothing matched the CSV.
    Breaking ties by which column's dates start closest to today reliably
    picks the current/forward-looking column instead.

    That "closest to today" tie-break only works when candidates are
    genuinely a year or more apart — it breaks down when two candidates are
    just a day or two apart (which happens when a sheet has more than one
    plausible calendar column, e.g. one mislabeled/off by a day), because
    whichever one starts *later* will almost always look "closer to today"
    and win even when it's the wrong one. Confirmed real case: Provincetown
    Crowne Pointe's WKTHREE mirrors two of WKONE's columns via cross-sheet
    formula — one starting the 2nd (genuinely wrong data in WKONE itself),
    one starting the 1st (correct) — and pure proximity picked the 2nd on
    almost every day of the month. This app's own convention
    (restructure_sr_dates) always starts a TY column on day 1 of the
    target month, so preferring a candidate that starts on the 1st is a
    more reliable tie-break than raw proximity, checked before falling
    back to proximity for the remaining genuine (same-day-of-month) ties.

    Formula cells are resolved via _resolve_formula_date (handles both
    local +N increments and cross-sheet references like '=WKONE!C5') when
    wb is provided, instead of just being ignored or blindly assumed to
    continue by exactly 1 day.

    Scans 60 rows (not just the first 10) — a 10-row window is too easily
    fooled by a sheet whose real calendar column doesn't fully establish
    itself that close to the top, or where some other left-side column
    coincidentally looks date-like in just that narrow band. A wider
    window gives every genuine candidate room to prove real consecutive
    length before scoring, at negligible extra cost.
    """
    today = datetime.date.today()
    sheet_name = ws.title
    best_col, best_key = 3, None  # fallback to col 3
    for c in range(1, 11):
        anchor_date = None   # first date of the longest verified-consecutive run
        last_date   = None   # most recently confirmed date in the current run
        count = 0
        best_run = 0
        run_anchor = None
        for r in range(5, min(ws.max_row + 1, 65)):
            v = ws.cell(r, c).value
            if isinstance(v, datetime.datetime):
                d = v.date()
            elif isinstance(v, datetime.date):
                d = v
            elif isinstance(v, str) and v.startswith("=") and wb is not None:
                d = _resolve_formula_date(wb, sheet_name, v)
                if d is None:
                    last_date = None
                    continue
            else:
                last_date = None  # non-date cell — any run in progress is broken
                continue

            if last_date is not None and (d - last_date).days == 1:
                count += 1
            else:
                run_anchor = d  # gap or scattered date — start a fresh run here
                count = 1
            last_date = d
            if count > best_run:
                best_run = count
                anchor_date = run_anchor
        if best_run < 3 or anchor_date is None:
            continue
        proximity  = abs((anchor_date - today).days)
        starts_1st = 1 if anchor_date.day == 1 else 0
        key = (best_run, starts_1st, -proximity)
        if best_key is None or key > best_key:
            best_key = key
            best_col = c
    return best_col


def detect_comp_set_columns(ws, col_map):
    """Find the comp set chart's TY (far-left) and LY (far-right) columns.
    The chart sits between Restrictions and TY TRANS REV.
    LY col  = last non-empty text header to the LEFT of TY TRANS REV.
    TY col  = first non-empty text header to the RIGHT of Restrictions.
    Scans rows 2-6 for headers (tolerates slight layout shifts).
    Returns (ty_col, ly_col) or (None, None) if not found.
    """
    trans_rev_col = col_map.get("trans_rev_ty")
    if not trans_rev_col:
        return None, None

    restrict_col = find_restrictions_col(ws, upto_col=trans_rev_col - 1)

    left_bound  = (restrict_col + 1) if restrict_col else max(1, trans_rev_col - 30)
    right_bound = trans_rev_col - 1

    def _has_text_header(c):
        for scan_r in range(2, 7):
            v = str(ws.cell(scan_r, c).value or "").strip()
            if v and not v.replace(".", "").isdigit():
                return True
        return False

    # LY col: scan left from just before TY TRANS REV
    ly_col = None
    for c in range(right_bound, left_bound - 1, -1):
        if _has_text_header(c):
            ly_col = c
            break

    # TY col: scan right from just after Restrictions
    ty_col = None
    for c in range(left_bound, right_bound + 1):
        if _has_text_header(c):
            ty_col = c
            break

    if ty_col == ly_col:
        return None, None  # same column — chart not found

    return ty_col, ly_col


def get_ly_sr_data(service, hotel_id, hotel_name, current_month, sheet_name):
    """Fetch LY data from last year's SR (same month, same week tab).
    Returns {date: {ly_field: value}} where dates are mapped to this year.
    Also returns (comp_ty_col_in_ly_ws, comp_ly_col_in_current_ws_placeholder)
    via a separate 'comp_set' key: {this_year_date: value}.
    """
    ly_month = current_month.replace(year=current_month.year - 1)
    result, err = resolve_drive_workbook(service, hotel_id, hotel_name,
                                         "Strategy Report", month_date=ly_month)
    if err or not result:
        return {}

    file_id, _ = result
    try:
        file_bytes = drive_download(service, file_id)
        ly_wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception:
        return {}

    if sheet_name not in ly_wb.sheetnames:
        return {}

    ly_ws = ly_wb[sheet_name]
    ly_col_map  = detect_strategy_columns(ly_ws)
    ly_date_col = detect_date_column(ly_ws, wb=ly_wb)

    # Build date→row for last year's sheet
    ly_date_row = {}
    for r in range(5, ly_ws.max_row + 1):
        v = ly_ws.cell(r, ly_date_col).value
        if isinstance(v, datetime.datetime):
            ly_date_row[v.date()] = r
        elif isinstance(v, datetime.date):
            ly_date_row[v] = r

    # Comp set: far-left TY col in LY sheet
    comp_ty_col, _ = detect_comp_set_columns(ly_ws, ly_col_map)

    out = {}  # {this_year_date: {field: value}}
    for ly_date, r in ly_date_row.items():
        this_year_date = ly_date.replace(year=ly_date.year + 1)
        row_data = {}

        # Pull each TY source field from last year
        for ly_dest_field, ty_src_field in LY_FROM_TY.items():
            src_col = ly_col_map.get(ty_src_field)
            if src_col:
                v = ly_ws.cell(r, src_col).value
                if v is not None and not is_formula(v):
                    row_data[ly_dest_field] = safe_float(v)

        # Comp set TY value (far-left hotel col in LY sheet) — keep text as-is (e.g. "Sold out", "LOS2")
        if comp_ty_col:
            v = ly_ws.cell(r, comp_ty_col).value
            if v is not None and not is_formula(v):
                row_data["comp_set_ly"] = v

        if row_data:
            out[this_year_date] = row_data

    return out


def build_date_row_map(wb, prefer_sheet=None, fallback_to_wkone=True):
    """Build {date: row_number} using auto-detected date column, trying
    prefer_sheet (the tab actually being updated) first and falling back to
    WKONE. When subsequent rows contain formulas (=C5+1 style), extrapolates
    from the first real date so the full year is mapped correctly.

    Confirmed real case: Provincetown Harbor Hotel's WKONE date column was
    stuck a year behind (2024/2025) while WKFOUR — the sheet actually being
    updated — had its own correct, independently populated 2025/2026
    column. Always reading WKONE broke every week's update, even weeks
    whose own calendar was perfectly fine, because their own date column
    was never even looked at.

    Set fallback_to_wkone=False when the caller is about to WRITE
    current-cycle data and prefer_sheet's own column must be the sole
    source of truth — no cross-sheet fallback allowed, since writing based
    on a different sheet's calendar than the one being edited risks
    silently landing data on the wrong row if the two ever drift apart.
    """
    candidates = [prefer_sheet] if prefer_sheet and prefer_sheet in wb.sheetnames else []
    if fallback_to_wkone and "WKONE" in wb.sheetnames:
        candidates.append("WKONE")

    for sheet_name in candidates:
        ws = wb[sheet_name]
        date_col = detect_date_column(ws, wb=wb)
        mapping = {}
        anchor_date = None
        anchor_row  = None
        # Scan at least 400 rows regardless of this sheet's own max_row —
        # confirmed real case: a week tab whose date column is entirely
        # cross-sheet formulas referencing WKONE can have its own "used
        # range" end far short of WKONE's actual full-year extent (e.g. if
        # the formulas were only ever dragged down ~33 rows), truncating
        # every month after that even though WKONE and the CSV both cover
        # the full year. Reading past a sheet's real content just returns
        # blank cells (openpyxl doesn't error), so this is always safe.
        scan_end = max(ws.max_row, 400)
        for row_num in range(5, scan_end + 1):
            val = ws.cell(row_num, date_col).value
            if isinstance(val, datetime.datetime):
                d = val.date()
            elif isinstance(val, datetime.date):
                d = val
            elif isinstance(val, str) and val.startswith("="):
                # Resolves both local +N increments and cross-sheet
                # references (e.g. '=WKONE!C5') — non-WKONE week tabs
                # mirror WKONE's calendar this way by design, so this must
                # actually follow the reference, not just assume +1 day.
                d = _resolve_formula_date(wb, sheet_name, val)
                if d is None and anchor_row is not None:
                    # Unreadable cell, but the row-to-date relationship is a
                    # fixed, already-proven pattern (row N is always N minus
                    # anchor_row days after the anchor) — confirmed real
                    # case: Crowne Pointe's date formula chain was only ever
                    # filled down partway, leaving later rows genuinely
                    # blank even though the CSV and WKONE both have a full
                    # year of real data. Extrapolate instead of stopping.
                    d = anchor_date + datetime.timedelta(days=row_num - anchor_row)
                elif d is None:
                    continue
            elif anchor_row is not None:
                # Truly blank cell past the established anchor — same
                # extrapolation as above.
                d = anchor_date + datetime.timedelta(days=row_num - anchor_row)
            else:
                continue
            if anchor_date is None:
                anchor_date = d
                anchor_row  = row_num
            mapping[d] = row_num
        if mapping:
            # Auto-correct a stale year instead of requiring the sheet's
            # own cells to be manually fixed every cycle — confirmed real
            # case: Harbor Hotel's WKONE date column was genuinely stuck at
            # 2025 (never advanced when the year rolled to 2026). If the
            # detected column's anchor year doesn't match the current year,
            # shift every mapped date forward by that many years. A
            # correctly-dated column's anchor is already close to today
            # (that's what detect_date_column's proximity scoring picks
            # for), so this is a no-op in the normal case.
            year_shift = datetime.date.today().year - anchor_date.year
            if year_shift != 0:
                shifted = {}
                for d, r in mapping.items():
                    try:
                        shifted[d.replace(year=d.year + year_shift)] = r
                    except ValueError:
                        # Feb 29 with no Feb 29 in the shifted year
                        shifted[d.replace(month=2, day=28, year=d.year + year_shift)] = r
                mapping = shifted
            return mapping
    return {}


def find_otb_date_cell(ws):
    """Return (row, col) of the as-of date cell — one row above the OTB/TRANS header.
    Tries both Plymouth-style (OTB TY / TRANS) and Long Beach-style (TRANS / SOLD).
    """
    col_map = detect_strategy_columns(ws)
    otb_col = col_map.get("otb_trans")
    if otb_col:
        # Find which of rows 3 or 4 has the header, then go one above
        for r in range(2, 6):
            v = str(ws.cell(r, otb_col).value or "").strip().upper()
            if "OTB" in v or "TRANS" in v or "SOLD" in v:
                return r - 1, otb_col
    return 2, 4  # fallback


def _extract_otb_trans_by_date(wb, sheet_name, from_date):
    """Read OTB TY Trans keyed by date from an in-memory workbook.
    Uses build_date_row_map for date→row resolution so formula-based date
    cells (=C5+1 style) are handled via anchor extrapolation.
    Only returns rows where date >= from_date.
    """
    if sheet_name not in wb.sheetnames:
        return {}
    ws = wb[sheet_name]
    col_map = detect_strategy_columns(ws)
    otb_col = col_map.get("otb_trans")
    if not otb_col:
        return {}
    date_row_map = build_date_row_map(wb, prefer_sheet=sheet_name)
    out = {}
    for d, r in date_row_map.items():
        if d < from_date:
            continue
        val = ws.cell(r, otb_col).value
        if val is not None and not is_formula(val):
            out[d] = safe_float(val)
    return out


def _extract_ly_data_from_wb(ly_wb, sheet_name, ty_wb=None):
    """Read all LY source fields + comp set TY col from an in-memory workbook.
    Returns {this_year_date: {field: value}} (dates shifted to match TY sheet's year).

    Robust to year mismatches in LY file (e.g., AUG2025 file has 2024 dates due to
    copy/paste error). Detects target year from TY sheet and shifts LY dates accordingly.
    """
    if sheet_name not in ly_wb.sheetnames:
        return {}

    # Detect target year from TY sheet if provided, otherwise infer from LY year + 1
    # For fiscal year files (Aug-Jul), use the LAST date to get the correct year,
    # not the first date (e.g., AUG2026 starts 2026-08-01 but ends 2027-07-31, so target_year=2027)
    target_year = None
    if ty_wb and sheet_name in ty_wb.sheetnames:
        ty_ws = ty_wb[sheet_name]
        ty_date_col = detect_date_column(ty_ws, wb=ty_wb)
        # Scan to find last date in column
        last_date_year = None
        for r in range(ty_ws.max_row, 4, -1):
            v = ty_ws.cell(r, ty_date_col).value
            if isinstance(v, datetime.datetime):
                last_date_year = v.year
                break
            elif isinstance(v, datetime.date):
                last_date_year = v.year
                break
        if last_date_year and last_date_year >= 2025:
            target_year = last_date_year

    # If target_year not detected, infer from LY source file's max year
    if not target_year:
        ws = ly_wb[sheet_name]
        max_year = None
        for col in range(1, ws.max_column + 1):
            for row in range(5, min(50, ws.max_row + 1)):
                v = ws.cell(row, col).value
                if isinstance(v, datetime.datetime):
                    if max_year is None or v.year > max_year:
                        max_year = v.year
                elif isinstance(v, datetime.date):
                    if max_year is None or v.year > max_year:
                        max_year = v.year
        if max_year:
            target_year = max_year + 1

    ws = ly_wb[sheet_name]
    col_map  = detect_strategy_columns(ws)

    # For LY extraction in multi-year files, use the HIGHEST year-labeled date column.
    # AUG2025 has: col 1 labeled "2024" (old), col 3 labeled "2025" (current)
    # We want the "2025" dates (col 3) for extraction.
    # Find the highest year label in row 4 and use its date column.
    date_col = detect_date_column(ws, wb=ly_wb)  # Default: column 3 (2025 TY)

    # Look for year labels in row 4 and find the highest year
    max_year_in_file = None
    max_year_col = None
    for col in range(1, ws.max_column + 1):
        cell_val = ws.cell(4, col).value
        if isinstance(cell_val, (int, float)):
            year_val = int(cell_val)
            if year_val >= 2020 and (max_year_in_file is None or year_val > max_year_in_file):
                max_year_in_file = year_val
                max_year_col = col

    # If we found a year label, use the date column for that year
    if max_year_col:
        for test_col in [max_year_col, max_year_col + 1]:  # Year label might be in col, date in col+1
            if test_col <= ws.max_column:
                test_val = ws.cell(5, test_col).value
                if isinstance(test_val, (datetime.datetime, datetime.date)):
                    date_col = test_col
                    break

    comp_ty_col, _ = detect_comp_set_columns(ws, col_map)

    out = {}
    for r in range(5, ws.max_row + 1):
        v = ws.cell(r, date_col).value
        if isinstance(v, datetime.datetime): d = v.date()
        elif isinstance(v, datetime.date):   d = v
        else: continue

        # Shift to target year, applying DOW adjustment (-1 day for non-leap year).
        # CRITICAL: Only subtract the -1 day if it doesn't cross a month boundary.
        # At year/month boundaries (e.g., Jan 1), -1 day crosses to previous month/year
        # and breaks the date matching. In those cases, use the date as-is.
        try:
            # Shift to target year and apply -1 day for DOW alignment.
            # This matches day-of-week across years (Mon 2025 → Mon 2026).
            # Example: 2025-08-02 (Fri) → 2026-08-01 (Fri)
            base_date = datetime.date(d.year + 1, d.month, d.day)
            this_year = base_date - datetime.timedelta(days=1)
        except ValueError:
            # Feb 29 in non-leap year: use Feb 28 instead
            if d.month == 2 and d.day == 29:
                base_date = datetime.date(d.year + 1, 2, 28)
                this_year = base_date - datetime.timedelta(days=1)
            else:
                # Other date error: skip this row
                continue

        row_data = {}
        # Read from TY columns — for SR LY extraction, this year's TY becomes next year's LY
        for ly_dest, ty_src in LY_FROM_TY.items():
            src_col = col_map.get(ty_src)
            if src_col:
                val = ws.cell(r, src_col).value
                if val is not None and not is_formula(val):
                    row_data[ly_dest] = safe_float(val)
        if comp_ty_col:
            val = ws.cell(r, comp_ty_col).value
            if val is not None and not is_formula(val):
                row_data["comp_set_ly"] = val  # preserve text values like "Sold out", "LOS2"
        if row_data:
            out[this_year] = row_data
    return out


def build_strategy_change_plan(df, wb, sheet_name, prev_month_wb=None, ly_wb=None,
                               scope_start=None, scope_end=None):
    """Build strategy changes.
    prev_month_wb: in-memory previous month's SR workbook (for OTB Lst Wek on WKONE)
    ly_wb:         in-memory last year's SR workbook (for all LY columns, every week)
    """
    today = datetime.date.today()
    if scope_start is None:
        scope_start = today.replace(day=1)
    if scope_end is None:
        scope_end = datetime.date(today.year + 1, 12, 31)

    date_row_map = build_date_row_map(wb, prefer_sheet=sheet_name)
    ws = wb[sheet_name]

    # Detect actual column positions from headers — no guessing
    col_map = detect_strategy_columns(ws)
    ly_only_fields = {"otb_lst_wk", "otb_ly_trans", "grp_pu_ly", "grp_npu_ly",
                      "trans_rev_ly", "grp_rev_ly", "grp_npu_rev_ly", "casino_ballroom_ly"}
    # Casino ballroom is optional — only some hotels (e.g., Hampton Beach) have it
    optional_fields = {"casino_ballroom", "casino_ballroom_ly"}
    missing = [f for f, c in col_map.items() if c is None and f not in ly_only_fields and f not in optional_fields]
    if missing:
        st.warning(f"Strategy: could not locate columns for: {', '.join(missing)}")

    # Comp set columns in current sheet
    comp_ty_col_cur, comp_ly_col_cur = detect_comp_set_columns(ws, col_map)

    # OTB Lst Wek — WKONE only, from previous month's SR (already in memory)
    prev_otb_map = {}
    src_sheet = None
    if sheet_name == "WKONE" and col_map.get("otb_lst_wk") and prev_month_wb:
        # Use last FILLED tab (opposite of first_undone) — checked by actual
        # OTB data, not tab color (see first_undone_strategy_sheet: the same
        # color can mark a genuinely completed week OR sit untouched on a
        # never-used master template, so color alone can't tell "filled").
        last_filled = None
        for s in STRATEGY_SHEETS:
            if s not in prev_month_wb.sheetnames:
                continue
            pcol_map = detect_strategy_columns(prev_month_wb[s])
            potb_col = pcol_map.get("otb_trans")
            if potb_col and any(isinstance(prev_month_wb[s].cell(r, potb_col).value, (int, float)) for r in range(5, 15)):
                last_filled = s
        src_sheet = last_filled or (STRATEGY_SHEETS[-1] if STRATEGY_SHEETS[-1] in prev_month_wb.sheetnames else None)
        if src_sheet:
            prev_otb_map = _extract_otb_trans_by_date(prev_month_wb, src_sheet, scope_start)

    # LY data — every week, from last year's same month/week tab (already in memory)
    ly_data = {}
    if ly_wb:
        ly_data = _extract_ly_data_from_wb(ly_wb, sheet_name, ty_wb=wb)

    changes = []

    # Today's date above the OTB TY TRANS header
    date_row, date_col = find_otb_date_cell(ws)
    if date_row >= 1:
        changes.append({
            "date": today, "row": date_row, "col": date_col,
            "label": "As-of date", "new_value": today,
            "skip_reason": "formula" if is_formula(ws.cell(date_row, date_col).value) else None,
        })

    # As-of date above OTB Lst Wek header (row above row 3, i.e. row 2)
    lst_wk_col = col_map.get("otb_lst_wk")
    if lst_wk_col and prev_month_wb and src_sheet:
        prev_ws = prev_month_wb[src_sheet]
        prev_date_row, _ = find_otb_date_cell(prev_ws)
        src_date = prev_ws.cell(prev_date_row, date_col).value if prev_date_row >= 1 else None
        if src_date is None:
            # fallback: look in row above the lst_wk header in source
            for hr in range(2, 6):
                if str(prev_ws.cell(hr, lst_wk_col).value or "").strip():
                    src_date = prev_ws.cell(hr - 1, lst_wk_col).value
                    break
        if src_date:
            hdr_row = next((r for r in range(2, 6) if str(ws.cell(r, lst_wk_col).value or "").strip()), 3)
            label_row = hdr_row - 1
            if label_row >= 1:
                changes.append({
                    "date": None, "row": label_row, "col": lst_wk_col,
                    "label": "OTB Lst Wek as-of date",
                    "new_value": src_date,
                    "skip_reason": "formula" if is_formula(ws.cell(label_row, lst_wk_col).value) else None,
                })

    # As-of date above OTB LY TRANS header
    ly_trans_col = col_map.get("otb_ly_trans")
    if ly_trans_col and ly_wb and sheet_name in ly_wb.sheetnames:
        ly_ws_src = ly_wb[sheet_name]
        ly_date_row, _ = find_otb_date_cell(ly_ws_src)
        ly_src_date = ly_ws_src.cell(ly_date_row, date_col).value if ly_date_row >= 1 else None
        if ly_src_date:
            if isinstance(ly_src_date, datetime.datetime):
                ly_src_date = ly_src_date.date()
            hdr_row = next((r for r in range(2, 6) if str(ws.cell(r, ly_trans_col).value or "").strip()), 3)
            label_row = hdr_row - 1
            if label_row >= 1:
                changes.append({
                    "date": None, "row": label_row, "col": ly_trans_col,
                    "label": "OTB LY Trans as-of date",
                    "new_value": ly_src_date,
                    "skip_reason": None,
                })

    # ── CSV-sourced TY columns (only when BOB uploaded) ──────────────────────
    # Uses ONLY this sheet's own date column (no WKONE fallback) — this is
    # current-cycle data being written, and it must always be placed
    # against the calendar actually printed on the sheet being edited.
    own_date_row_map = build_date_row_map(wb, prefer_sheet=sheet_name, fallback_to_wkone=False)
    for _, row in (df.iterrows() if df is not None else []):
        date_str = str(row[0]).strip() if row[0] else ""
        kind, info = classify_row(date_str)
        if kind != "daily":
            continue
        d = info
        if d < scope_start or d > scope_end:
            continue
        if d not in own_date_row_map:
            continue
        excel_row = own_date_row_map[d]
        for field, (csv_col, label) in STRATEGY_CSV_COLS.items():
            excel_col = col_map.get(field)
            if excel_col is None:
                continue
            val = safe_float(row[csv_col])
            # BOB data is always authoritative — write even if cell has a template formula
            changes.append({
                "date": d, "row": excel_row, "col": excel_col,
                "label": label, "new_value": val, "skip_reason": None,
            })

    # ── Drive-sourced columns (run every time, no CSV needed) ─────────────────
    all_dates = set(date_row_map.keys()) & (set(prev_otb_map) | set(ly_data))
    for d in sorted(all_dates):
        if d < scope_start or d > scope_end:
            continue
        excel_row = date_row_map[d]

        # OTB Lst Wek — WKONE only
        lst_wk_col = col_map.get("otb_lst_wk")
        if lst_wk_col and d in prev_otb_map:
            skip = "formula" if is_formula(ws.cell(excel_row, lst_wk_col).value) else None
            changes.append({
                "date": d, "row": excel_row, "col": lst_wk_col,
                "label": "OTB Lst Wek", "new_value": prev_otb_map[d], "skip_reason": skip,
            })

        # LY columns
        if d in ly_data:
            row_ly = ly_data[d]
            for ly_field, ly_label in [
                ("otb_ly_trans",   "OTB LY Trans"),
                ("grp_pu_ly",      "GRP PU LY"),
                ("grp_npu_ly",     "GRP N/PU LY"),
                ("trans_rev_ly",   "LY Trans Rev"),
                ("grp_rev_ly",     "GRP LY Rev"),
                ("grp_npu_rev_ly", "GRP N/PU LY Rev"),
            ]:
                dest_col = col_map.get(ly_field)
                if dest_col and ly_field in row_ly:
                    changes.append({
                        "date": d, "row": excel_row, "col": dest_col,
                        "label": ly_label, "new_value": row_ly[ly_field], "skip_reason": None,
                    })

            # Comp set far-right (LY) ← far-left (TY) from last year
            if comp_ly_col_cur and "comp_set_ly" in row_ly:
                changes.append({
                    "date": d, "row": excel_row, "col": comp_ly_col_cur,
                    "label": "Comp Set LY", "new_value": row_ly["comp_set_ly"], "skip_reason": None,
                })

    # ── Blank LY cells with no confirmed LY data — don't leave stale leftovers
    # from whatever the template last held (e.g. a day with no LY rates). Only
    # runs when last year's SR actually loaded AND we extracted LY data (not
    # just a regular weekly upload). Skip if ly_data is empty (extraction failed).
    if ly_wb and ly_data:
        ly_field_labels = [
            ("otb_ly_trans",   "OTB LY Trans"),
            ("grp_pu_ly",      "GRP PU LY"),
            ("grp_npu_ly",     "GRP N/PU LY"),
            ("trans_rev_ly",   "LY Trans Rev"),
            ("grp_rev_ly",     "GRP LY Rev"),
            ("grp_npu_rev_ly", "GRP N/PU LY Rev"),
        ]
        for d, excel_row in date_row_map.items():
            if d < scope_start or d > scope_end:
                continue
            row_ly = ly_data.get(d, {})
            for ly_field, ly_label in ly_field_labels:
                dest_col = col_map.get(ly_field)
                if not dest_col or ly_field in row_ly:
                    continue
                cur_val = ws.cell(excel_row, dest_col).value
                if cur_val is None or is_formula(cur_val):
                    continue
                changes.append({
                    "date": d, "row": excel_row, "col": dest_col,
                    "label": f"{ly_label} (no LY data — cleared)",
                    "new_value": None, "skip_reason": None,
                })
            if comp_ly_col_cur and "comp_set_ly" not in row_ly:
                cur_val = ws.cell(excel_row, comp_ly_col_cur).value
                if cur_val is not None and not is_formula(cur_val):
                    changes.append({
                        "date": d, "row": excel_row, "col": comp_ly_col_cur,
                        "label": "Comp Set LY (no LY data — cleared)",
                        "new_value": None, "skip_reason": None,
                    })

    return changes


def apply_strategy_changes(wb, sheet_name, changes):
    ws = wb[sheet_name]
    for ch in changes:
        if ch["skip_reason"]:
            continue
        ws.cell(ch["row"], ch["col"]).value = ch["new_value"]


def parse_rate_csv(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes), dtype=str, encoding="utf-8-sig")
    df.columns = [c.strip() for c in df.columns]
    return df


def find_header_col(ws, keyword, header_rows=(2, 3, 4)):
    """Find the leftmost column whose concatenated header text (rows 2-4) contains keyword."""
    keyword = keyword.strip().lower()
    col_texts = {}
    for row_num in header_rows:
        for cell in ws[row_num]:
            col = cell.column
            if col not in col_texts:
                col_texts[col] = ""
            if isinstance(cell.value, str):
                col_texts[col] += cell.value.strip()
    matches = [col for col, text in col_texts.items() if keyword in text.lower()]
    return min(matches) if matches else None


def find_restrictions_col(ws, upto_col=None):
    """Find the 'Restrictions' header column, scanning rows 3-4.
    The header can be split across those two rows AND hyphenated for
    word-wrap (e.g. 'Restric'+'tions', or 'Restri-'+'ctions' — confirmed on
    a real hotel's sheet, ALLEGRIA/'Long Beach'). Stripping spaces and
    hyphens before matching handles both forms; a plain substring search on
    the raw concatenation misses the hyphenated one and returns nothing.

    Also confirmed a master template with the header simply misspelled —
    'Restictions' (missing the second 'r') on Hampton/Ashworth By The Sea, in
    a single cell, no split at all. Rather than chase every possible typo,
    match on 'REST' + 'TION' both present, which covers correct spelling,
    the hyphenated/split form, and this typo alike.
    """
    upto_col = upto_col or ws.max_column
    for c in range(1, upto_col + 1):
        combined_hdr = (str(ws.cell(3, c).value or "") + str(ws.cell(4, c).value or "")).upper().replace(" ", "").replace("-", "")
        if "REST" in combined_hdr and "TION" in combined_hdr:
            return c
    return None


def build_rates_change_plan(rate_df, wb, sheet_name):
    today = datetime.date.today()
    # include previous month — final numbers arrive on the 1st of the following month
    prev_month_start = (today.replace(day=1) - datetime.timedelta(days=1)).replace(day=1)
    scope_start = prev_month_start

    date_row_map = build_date_row_map(wb, prefer_sheet=sheet_name)
    # Cover the sheet's actual date range, not a hardcoded Dec 31 — the SR's
    # rolling 12 months usually crosses into the next calendar year (e.g. a
    # July setup runs into next June), and a hardcoded year-end cutoff was
    # silently dropping every month past December, regardless of what real
    # data the CSV had for those dates.
    scope_end = max(date_row_map.keys()) if date_row_map else datetime.date(today.year, 12, 31)
    ws = wb[sheet_name]

    restric_col = find_restrictions_col(ws)

    # Find hotel rate column by scanning right from Restrictions looking for "hotel" or "rate"
    # Don't assume it's restric_col+1 — there may be intermediate columns (e.g. casino ballroom).
    # Always read actual column headers, never assume positions.
    hotel_col = None
    if restric_col:
        for c in range(restric_col + 1, min(restric_col + 10, ws.max_column + 1)):
            hdr3 = str(ws.cell(3, c).value or "").upper()
            hdr4 = str(ws.cell(4, c).value or "").upper()
            combined = hdr3 + " " + hdr4
            if ("HOTEL" in combined or "RATE" in combined) and "BALLROOM" not in combined:
                hotel_col = c
                break

    changes = []
    warnings = []
    if not restric_col:
        warnings.append("Could not find Restrictions column in sheet headers.")
    if not hotel_col and restric_col:
        warnings.append("Could not find Hotel Rate column after Restrictions.")

    for _, row in rate_df.iterrows():
        date_str = str(row.get("Date", "")).strip()
        d = None
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                d = datetime.datetime.strptime(date_str, fmt).date()
                break
            except ValueError:
                continue
        if d is None:
            continue
        if d < scope_start or d > scope_end:
            continue
        if d not in date_row_map:
            continue

        excel_row  = date_row_map[d]
        double_val = safe_float(row.get("Double", ""))
        mlos_val   = safe_float(row.get("Min Length of Stay", ""))

        if hotel_col:
            skip = "formula" if is_formula(ws.cell(excel_row, hotel_col).value) else None
            changes.append({"date": d, "row": excel_row, "col": hotel_col,
                            "label": "Hotel Rate", "new_value": double_val, "skip_reason": skip})
        if restric_col:
            skip = "formula" if is_formula(ws.cell(excel_row, restric_col).value) else None
            changes.append({"date": d, "row": excel_row, "col": restric_col,
                            "label": "Restrictions (MLOS)", "new_value": mlos_val, "skip_reason": skip})

    return changes, warnings


# ── Forecast helpers ─────────────────────────────────────────────────────────

DATE_FORMATS = [
    "%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%y",
    "%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y",
    "%d-%b-%Y", "%d/%b/%Y",
]

def parse_any_date(val):
    """Parse a date from a datetime object, int serial, or string in many formats."""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.date() if isinstance(val, datetime.datetime) else val
    if isinstance(val, (int, float)):
        # Excel serial date (days since 1900-01-01, with Lotus bug offset)
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val)))
    if isinstance(val, str):
        val = val.strip()
        for fmt in DATE_FORMATS:
            try:
                return datetime.datetime.strptime(val, fmt).date()
            except ValueError:
                continue
    return None


def locate_forecast_rows(ws):
    """Find the As-of date / OTB Rooms Sold / ADR OTB / actual Rooms Sold /
    actual Revenue rows by reading column-A titles, instead of assuming fixed
    row numbers. Confirmed on a real workbook that this drifts within a single
    file: WK1 has an extra 'Occupancy' row that WK4/WK8/WK9 don't, shifting
    every row below it by one — a hardcoded row 14/16/19 lands on a blank row
    or the wrong label on most week tabs.

    'Rooms Sold' and 'Revenue' each appear multiple times in this template, so
    rows are resolved in reading order relative to the unique 'ADR OTB' label:
    the first 'Rooms Sold' above it is the OTB (future) entry row, the first
    'Rooms Sold' below it is the actual entry row, and the first 'Revenue'
    below that is the actual revenue row.

    Returns None if the expected labels aren't all found (caller should warn
    and skip rather than guess a row number).
    """
    labels = []
    for r in range(1, 31):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip():
            labels.append((r, v.strip().lower()))

    def find_after(text, after_row=0):
        for r, v in labels:
            if r > after_row and text in v:
                return r
        return None

    dow_row            = find_after("day of week")
    otb_rooms_row      = find_after("rooms sold")
    adr_otb_row        = find_after("adr otb")
    actual_rooms_row   = find_after("rooms sold", adr_otb_row) if adr_otb_row else None
    actual_revenue_row = find_after("revenue", actual_rooms_row) if actual_rooms_row else None

    if not all([dow_row, otb_rooms_row, adr_otb_row, actual_rooms_row, actual_revenue_row]):
        return None

    return {
        "as_of_row":          dow_row - 1,
        "date_row":           dow_row + 1,
        "otb_rooms_row":      otb_rooms_row,
        "adr_otb_row":        adr_otb_row,
        "actual_rooms_row":   actual_rooms_row,
        "actual_revenue_row": actual_revenue_row,
    }


def build_forecast_date_col_map(ws, wb=None, date_row=4):
    """Return {date: col_index} from date_row. Falls back to WK1 for formula-only sheets."""
    month_start = parse_any_date(ws.cell(date_row, 2).value)

    # If this sheet's col B is a formula, find the start date from any WK sheet with a literal
    if month_start is None and wb is not None:
        for sname in wb.sheetnames:
            if "glance" in sname.lower():
                continue
            candidate = parse_any_date(wb[sname].cell(date_row, 2).value)
            if candidate is not None:
                month_start = candidate
                break

    if month_start is None:
        return {}

    col_map = {}
    col = 2
    while col <= ws.max_column:
        cell = ws.cell(date_row, col)
        if isinstance(cell.value, str) and "total" in cell.value.lower():
            break
        if cell.value is None and col > 2:
            break
        col_map[month_start + datetime.timedelta(days=col - 2)] = col
        col += 1
    return col_map


def row_is_filled(ws, r):
    """Return True if cols B-AF contain actual numbers OR cross-sheet formula refs."""
    for col in range(2, 33):
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)):
            return True
        if isinstance(v, str) and v.startswith("='"):
            return True
    return False


def find_next_pickup_data_row(ws):
    """Find the next available row in the pick-up tracking chart.

    Strategy:
    1. Locate the 'Day of Week' section header (last one before 'Total Pick UP').
    2. Within the section, find the last row that has actual numbers in B-AF
       (cross-sheet formula refs look empty to openpyxl, so only direct numbers count).
    3. Scan forward from there for the first row with no numbers and no skip keyword
       — works for WK1/WK2 (back-to-back entries) and WK3+ (Pick UP rows between).
    """
    skip_keywords = {"pick", "day", "total", "forecast", "budget", "last year"}

    # Find Total Pick UP boundary
    total_row = None
    for r in range(40, 150):
        if "total pick" in str(ws.cell(r, 1).value or "").lower():
            total_row = r
            break
    search_end = total_row if total_row else 150

    # Find last 'Day of Week' header before the boundary
    section_start = None
    for r in range(40, search_end):
        if "day of week" in str(ws.cell(r, 1).value or "").lower():
            section_start = r
    if section_start is None:
        return None

    # Find the last row in the section that has actual numbers in B-AF
    last_filled = None
    for r in range(section_start + 2, search_end):
        if row_is_filled(ws, r):
            last_filled = r

    if last_filled is None:
        # Nothing filled yet — return first non-header row in section
        last_filled = section_start + 1

    # Scan forward from last_filled+1 for the first row with no numbers
    for r in range(last_filled + 1, search_end):
        a_val = str(ws.cell(r, 1).value or "").strip().lower()
        if any(k in a_val for k in skip_keywords):
            continue
        if not row_is_filled(ws, r):
            return r

    return None


def extract_rob_month_end_data(rob_wb, target_month):
    """Pull Budget and Last Year Room Nights + Revenue for target_month from any ROB sheet."""
    for sheet_name in rob_wb.sheetnames:
        ws = rob_wb[sheet_name]
        # Scan for the month block header row (col A = month abbreviation, e.g. "Jul")
        month_abbr = target_month.strftime("%b").lower()
        for r in range(1, ws.max_row + 1):
            cell_val = str(ws.cell(r, 1).value or "").strip().lower()
            if cell_val != month_abbr:
                continue
            # Found the block header row — find Budget and Last Year columns
            budget_col = ly_col = None
            for c in range(1, 25):
                hdr = str(ws.cell(r, c).value or "").strip()
                if not budget_col and "budget" in hdr.lower():
                    budget_col = c
                elif not ly_col and ("month end" in hdr.lower() or ("last" in hdr.lower() and "year" in hdr.lower())):
                    ly_col = c
                if budget_col and ly_col:
                    break
            # Also check row 4 (global header) if not found on block row
            if not budget_col or not ly_col:
                for c in range(1, 25):
                    hdr = str(ws.cell(4, c).value or "").strip()
                    if not budget_col and "budget" in hdr.lower():
                        budget_col = c
                    elif not ly_col and "month end" in hdr.lower():
                        ly_col = c
                    if budget_col and ly_col:
                        break
            if not budget_col or not ly_col:
                continue
            # Scan the next ~8 rows for Revenue and Room Nights labels in col A
            rev_row = rms_row = None
            for dr in range(1, 9):
                label = str(ws.cell(r + dr, 1).value or "").strip().lower()
                if "revenue" in label and rev_row is None:
                    rev_row = r + dr
                elif ("room" in label or "night" in label) and rms_row is None:
                    rms_row = r + dr
                if rev_row and rms_row:
                    break
            if not rev_row or not rms_row:
                continue
            return {
                "budget_rev": safe_float(ws.cell(rev_row, budget_col).value),
                "budget_rms": safe_float(ws.cell(rms_row, budget_col).value),
                "ly_rev":     safe_float(ws.cell(rev_row, ly_col).value),
                "ly_rms":     safe_float(ws.cell(rms_row, ly_col).value),
            }
    return None


def find_month_ending_forecast_cells(fcst_ws):
    """Return cell coords for Budget/LY Room Nts and Revenue in the Month Ending Forecast table."""
    for r in range(1, fcst_ws.max_row + 1):
        for c in range(1, fcst_ws.max_column + 1):
            if "month ending forecast" in str(fcst_ws.cell(r, c).value or "").lower():
                # Header found — scan next ~6 rows for Budget and Last Year rows
                # and the header row for column positions
                hdr_row = col_rms = col_rev = None
                budget_row = ly_row = None
                for dr in range(1, 8):
                    row_vals = {col: str(fcst_ws.cell(r + dr, col).value or "").strip()
                                for col in range(1, 8)}
                    row_text = " ".join(row_vals.values()).lower()
                    if "room" in row_text and hdr_row is None:
                        hdr_row = r + dr
                        for col, v in row_vals.items():
                            if "room" in v.lower():
                                col_rms = col
                            elif "revenue" in v.lower() or "rev" in v.lower():
                                col_rev = col
                    if "budget" in row_text and budget_row is None:
                        budget_row = r + dr
                    if "last year" in row_text and ly_row is None:
                        ly_row = r + dr
                if col_rms and col_rev and budget_row and ly_row:
                    return col_rms, col_rev, budget_row, ly_row
    return None, None, None, None


def build_forecast_change_plan(df, ws, rob_wb=None, is_wk1=False):
    """Build list of cell writes for the Forecast sheet."""
    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)
    month_start = today.replace(day=1)

    rows = locate_forecast_rows(ws)
    if not rows:
        return [], ["Could not read row titles (As-of date / Rooms Sold / ADR OTB / Revenue) from forecast sheet."]

    col_map = build_forecast_date_col_map(ws, ws.parent, date_row=rows["date_row"])
    if not col_map:
        return [], ["Could not read date row from forecast sheet."]

    changes = []
    warnings = []

    # As-of date (row directly above "Day of Week")
    changes.append({
        "label": "As-of date", "row": rows["as_of_row"], "col": 1,
        "new_value": today, "skip_reason": "formula" if is_formula(ws.cell(rows["as_of_row"], 1).value) else None,
    })

    # Build lookup from CSV: date -> row dict
    # Uses the same flexible date matching as ROB/SR (classify_row) — the BOB
    # CSV's date format (1-2 digit month/day, "/" or "-") wasn't matching the
    # old hyphen-only, zero-padded regex here, so Forecast silently got no data.
    daily_rows = {}
    for _, row in df.iterrows():
        date_str = str(row.iloc[0]).strip()
        kind, d = classify_row(date_str)
        if kind != "daily":
            continue
        daily_rows[d] = row

    for d, col in col_map.items():
        if d not in daily_rows:
            continue
        csv_row = daily_rows[d]
        rms = safe_float(csv_row.iloc[1])
        adr = safe_float(csv_row.iloc[6])
        rev = safe_float(csv_row.iloc[5])

        is_future = d >= today
        is_past   = d <= yesterday and d >= month_start

        # Rooms Sold (future / OTB)
        if is_future:
            skip = "formula" if is_formula(ws.cell(rows["otb_rooms_row"], col).value) else None
            changes.append({"label": f"Rooms Sold (future) {d}", "row": rows["otb_rooms_row"], "col": col,
                            "new_value": rms, "skip_reason": skip})
            # ADR OTB (future)
            skip = "formula" if is_formula(ws.cell(rows["adr_otb_row"], col).value) else None
            changes.append({"label": f"ADR OTB {d}", "row": rows["adr_otb_row"], "col": col,
                            "new_value": adr, "skip_reason": skip})

        # Rooms Sold (actuals)
        if is_past:
            skip = "formula" if is_formula(ws.cell(rows["actual_rooms_row"], col).value) else None
            changes.append({"label": f"Rooms Sold (actual) {d}", "row": rows["actual_rooms_row"], "col": col,
                            "new_value": rms, "skip_reason": skip})
            # Revenue (actuals)
            skip = "formula" if is_formula(ws.cell(rows["actual_revenue_row"], col).value) else None
            changes.append({"label": f"Revenue (actual) {d}", "row": rows["actual_revenue_row"], "col": col,
                            "new_value": rev, "skip_reason": skip})

    # Pick-up tracking row: write full month rooms sold to next available row
    target_row = find_next_pickup_data_row(ws)

    if target_row:
        changes.append({"label": "Pickup tracking: date", "row": target_row, "col": 1,
                        "new_value": today, "skip_reason": None})
        for d, col in col_map.items():
            if d not in daily_rows:
                continue
            rms = safe_float(daily_rows[d].iloc[1])
            skip = "formula" if is_formula(ws.cell(target_row, col).value) else None
            changes.append({"label": f"Pickup tracking: Rooms Sold {d}",
                            "row": target_row, "col": col,
                            "new_value": rms, "skip_reason": skip})
    else:
        warnings.append("No available row found in pick-up tracking chart.")

    # Month Ending Forecast table — only on WK1, only when ROB workbook provided
    if is_wk1 and rob_wb is not None:
        target_month = today.replace(day=1)
        rob_data = extract_rob_month_end_data(rob_wb, target_month)
        if rob_data:
            col_rms, col_rev, budget_row, ly_row = find_month_ending_forecast_cells(ws)
            if col_rms and col_rev and budget_row and ly_row:
                entries = [
                    (budget_row, col_rms, "Month End Forecast: Budget Room Nts", rob_data["budget_rms"]),
                    (budget_row, col_rev, "Month End Forecast: Budget Revenue",  rob_data["budget_rev"]),
                    (ly_row,     col_rms, "Month End Forecast: LY Room Nts",     rob_data["ly_rms"]),
                    (ly_row,     col_rev, "Month End Forecast: LY Revenue",       rob_data["ly_rev"]),
                ]
                for r, c, label, val in entries:
                    skip = "formula" if is_formula(ws.cell(r, c).value) else None
                    changes.append({"label": label, "row": r, "col": c,
                                    "new_value": val, "skip_reason": skip})
            else:
                warnings.append("Could not locate Month Ending Forecast table in forecast sheet.")
        else:
            warnings.append("Could not find month data in ROB workbook for Month Ending Forecast.")

    return changes, warnings


def build_next_month_forecast_plan(df, ws):
    """For weeks 3 & 4: write Rooms Sold and ADR OTB (rows located by title,
    not a fixed offset) for ALL dates in the next month (everything in the CSV
    beyond the current month). Also writes the As-of date and pick-up tracking row.
    """
    today = datetime.date.today()
    current_month_end = (today.replace(day=1) + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)

    rows = locate_forecast_rows(ws)
    if not rows:
        return [], ["Could not read row titles (As-of date / Rooms Sold / ADR OTB / Revenue) from next-month forecast sheet."]

    col_map = build_forecast_date_col_map(ws, ws.parent, date_row=rows["date_row"])
    if not col_map:
        return [], ["Could not read date row from next-month forecast sheet."]

    changes = []
    warnings = []

    # As-of date (row directly above "Day of Week")
    changes.append({
        "label": "As-of date", "row": rows["as_of_row"], "col": 1,
        "new_value": today,
        "skip_reason": "formula" if is_formula(ws.cell(rows["as_of_row"], 1).value) else None,
    })

    daily_rows = {}
    for _, row in df.iterrows():
        date_str = str(row.iloc[0]).strip()
        kind, d = classify_row(date_str)
        if kind != "daily":
            continue
        daily_rows[d] = row

    for d, col in col_map.items():
        # Only next month dates (after current month end)
        if d <= current_month_end:
            continue
        if d not in daily_rows:
            continue
        csv_row = daily_rows[d]
        rms = safe_float(csv_row.iloc[1])
        adr = safe_float(csv_row.iloc[6])

        skip_rms = "formula" if is_formula(ws.cell(rows["otb_rooms_row"], col).value) else None
        skip_adr = "formula" if is_formula(ws.cell(rows["adr_otb_row"], col).value) else None
        changes.append({"label": f"Rooms Sold (future) {d}", "row": rows["otb_rooms_row"], "col": col, "new_value": rms, "skip_reason": skip_rms})
        changes.append({"label": f"ADR OTB {d}",             "row": rows["adr_otb_row"],   "col": col, "new_value": adr, "skip_reason": skip_adr})

    # Pick-up tracking row (same logic — full month of next month dates)
    target_row = find_next_pickup_data_row(ws)
    if target_row:
        changes.append({"label": "Pickup tracking: date", "row": target_row, "col": 1,
                        "new_value": today, "skip_reason": None})
        for d, col in col_map.items():
            if d not in daily_rows:
                continue
            rms = safe_float(daily_rows[d].iloc[1])
            skip = "formula" if is_formula(ws.cell(target_row, col).value) else None
            changes.append({"label": f"Pickup tracking: Rooms Sold {d}",
                            "row": target_row, "col": col, "new_value": rms, "skip_reason": skip})
    else:
        warnings.append("No available row found in next-month pick-up tracking chart.")

    return changes, warnings


def apply_forecast_changes(wb, sheet_name, changes):
    ws = wb[sheet_name]
    for ch in changes:
        if ch["skip_reason"]:
            continue
        ws.cell(ch["row"], ch["col"]).value = ch["new_value"]


# ── Google Drive ──────────────────────────────────────────────────────────────

MULTI_ID_PREFIX = "MULTI:"

# Hotels whose Drive folders had to be shared per-month/year directly (no
# common parent folder to share instead) get grouped by a known keyword
# instead of trying to generically parse the date prefix out of the folder
# name — that parsing kept breaking on real naming inconsistencies (mixed
# case, 2- vs 4-digit years, colon/period/no separator, even "Anchor In" vs
# "Hyannis Anchor In" for the same hotel). Matching on a fixed keyword that's
# guaranteed present in every one of that hotel's folder names is far more
# reliable. Add an entry here for each hotel using this sharing pattern.
KNOWN_MULTI_FOLDER_HOTELS = {
    "Hyannis Anchor In":     ["ANCHOR"],
    "Provincetown Surfside": ["SURFSIDE"],
    "Hotel 1620":            ["1620", "PLYMOUTH"],
    "Wolfeboro":             ["WOLF"],
    "Provincetown Harbor Hotel": ["HARBOR"],
    "Provincetown Brass Key":    ["BRASS"],
    "Provincetown Crowne Pointe": ["CROWN"],
}


def _strip_dedup_suffix(name):
    """Strip a trailing ' (1)', ' (2)', etc. — Drive's auto-added suffix when
    a folder name collides with an existing one (confirmed real case:
    Surfside's hotel folder is literally named 'SURFSIDE (1)' in Drive).
    Purely cosmetic for the dropdown; the real folder_id is unaffected.
    """
    return re.sub(r'\s*\(\d+\)\s*$', '', name).strip()


def _is_rev_reports_name(name):
    """True if a folder name denotes a 'Revenue Reports' folder. Confirmed
    real naming variants include the full phrase and the abbreviation
    'REV RPTS' (Hotel 1620's convention, e.g. 'G. JUL2018 REV RPTS HOTEL
    1620') — a plain 'REVENUE REPORTS' in name_upper substring check misses
    the abbreviated form entirely, which broke both hotel grouping and
    workbook resolution for that hotel.
    """
    return bool(re.search(r'REVENUE REPORTS|REV RPTS', name.upper()))


def _extract_hotel_name_from_rev_folder(name):
    """Strip date-ish prefixes and the phrase 'Revenue Reports' from a folder
    name, leaving just the hotel name. Confirmed real naming is inconsistent
    even for a single hotel (Hyannis Anchor In, sharing per-month/year folders
    since there's no common parent to share instead): 'A: APR2025 Revenue
    Reports Anchor In', 'A: JUN2025 Revenue Reports Hyannis Anchor In',
    'Jul2024 Revenue Reports Hyannis Anchor In', 'A: MAY25 Revenue Reports
    Hyannis Anchor In' — mixing a leading single-letter marker (with a colon,
    period, or dash — not just a period), 2- or 4-digit years, and even
    inconsistent inclusion of "Hyannis" in the hotel name itself.

    Only strips a leading single letter when followed by a clear separator
    (":", ".", "-") — a bare leading letter with no separator is the start of
    a month name (e.g. "Jul2024") and must be left alone, not eaten.
    """
    s = name.strip()
    s = re.sub(r'^[A-Za-z]\s*[\.\:\-]\s*', '', s)
    s = re.sub(r'^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s*\d{2,4}\s+', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^\d{4}\s+', '', s)
    s = re.sub(r'revenue\s*reports|rev\s+rpts', '', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+', ' ', s).strip()
    return s.strip(' -:.,')


@st.cache_data(ttl=300)
def get_hotels_from_drive():
    """Return list of (display_name, folder_id) for every top-level folder
    that contains a 'REVENUE REPORTS' subfolder — i.e. each hotel folder.
    Cached for 5 minutes so it doesn't hit Drive on every rerender.

    Some hotels are shared with the service account directly at the REVENUE
    REPORTS folder level, not its parent — this happens on Shared Drives
    where the person granting access can only share folders they themselves
    have permission on, and Drive permissions don't propagate upward to a
    parent. Some of those hotels additionally have a SEPARATE REVENUE
    REPORTS folder per year (confirmed real case: Hyannis Anchor In), each
    shared individually since there's no common parent to share instead —
    those get grouped into a single hotel entry by name, with folder_id set
    to 'MULTI:<id>,<id>,...' listing every year's folder. resolve_drive_workbook
    and _find_rev_reports_folder_for_year both know how to unpack this.
    """
    try:
        svc = get_drive_service()
        q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        result = svc.files().list(
            q=q, fields="files(id, name)", pageSize=200,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        folders = result.get("files", [])

        hotels = []
        known_groups = {}  # hotel display name -> [folder_id, ...]
        rev_groups = []     # list of {"display": str, "ids": [folder_id, ...]} for unregistered hotels
        for folder in folders:
            name = folder["name"]
            name_upper = name.upper()

            # "Ancillary" folders are a different report category entirely
            # (ancillary revenue / front desk reports, not the ROB/SR/Forecast
            # per-month structure) — confirmed real case: "Ancillary
            # Provincetown Surfside" showing up as its own bogus hotel entry.
            # Exclude from grouping regardless of which hotel keyword matches.
            if "ANCILLARY" in name_upper:
                continue

            known_match = next((hn for hn, kws in KNOWN_MULTI_FOLDER_HOTELS.items()
                                 if any(kw in name_upper for kw in kws)), None)
            if known_match:
                known_groups.setdefault(known_match, []).append(folder["id"])
                continue

            if _is_rev_reports_name(name):
                extracted = _extract_hotel_name_from_rev_folder(name)
                if not extracted:
                    # No hotel name left after stripping the date prefix and
                    # "Revenue Reports" — an orphaned/mis-named folder with no
                    # usable hotel name (confirmed real case: "I. SEPT2021
                    # Revenue Reports", missing the hotel name entirely).
                    # Skip it rather than showing a meaningless entry.
                    continue
                norm = extracted.upper()
                # Merge by substring containment, not exact match — the same
                # hotel's own folders don't always spell its name the same way
                # (confirmed real case: some of Hyannis Anchor In's folders
                # say just "Anchor In", others "Hyannis Anchor In"). Keep the
                # longest/most complete variant seen as the display name.
                match = next((g for g in rev_groups
                              if norm in g["display"].upper() or g["display"].upper() in norm), None)
                if match:
                    match["ids"].append(folder["id"])
                    if len(extracted) > len(match["display"]):
                        match["display"] = extracted
                else:
                    rev_groups.append({"display": extracted, "ids": [folder["id"]]})
                continue
            if re.search(r'\b20\d{2}\b', name):
                continue
            child_q = ("'%s' in parents and trashed = false and "
                       "mimeType = 'application/vnd.google-apps.folder'") % folder["id"]
            children = svc.files().list(
                q=child_q, fields="files(name)", pageSize=20,
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute()
            has_rev = any(_is_rev_reports_name(c["name"]) for c in children.get("files", []))
            if has_rev:
                hotels.append((_strip_dedup_suffix(name), folder["id"]))

        for hotel_name, ids in known_groups.items():
            hotels.append((_strip_dedup_suffix(hotel_name), ids[0] if len(ids) == 1 else MULTI_ID_PREFIX + ",".join(ids)))

        for info in rev_groups:
            display = _strip_dedup_suffix(info["display"])
            if len(info["ids"]) == 1:
                hotels.append((display, info["ids"][0]))
            else:
                hotels.append((display, MULTI_ID_PREFIX + ",".join(info["ids"])))

        return sorted(hotels, key=lambda x: x[0])
    except Exception:
        return []

WORKBOOK_TYPES = ["ROB", "Strategy Report", "Forecast"]

# ── Portfolios ───────────────────────────────────────────────────────────────
# Hotels are discovered from Drive folder names, which carry no notion of
# grouping, so membership is declared here. Each portfolio's data arrives in a
# different export, which is why they get separate tabs rather than one list.
#
# Matching is on keywords, not exact names: Drive folder naming is
# inconsistent for the same hotel and two properties are filed under a
# different name entirely (Long Beach as ALLEGRIA, Westerly as Pleasant View
# Inn). Keep keywords distinctive enough not to collide.
PORTFOLIOS = ["Stay In Touch", "Hilton", "IHG"]

PORTFOLIO_HOTELS = {
    "Stay In Touch": {
        "Wolfeboro":      ["WOLF"],
        "Harbor Hotel":   ["HARBOR"],
        "Foxberry":       ["FOXBER", "FOXBUR"],
        "Tybee":          ["TYBEE"],
        "Provincetown Inn": ["PROVINCETOWN INN", "PTOWN INN"],
        "Middletown":     ["MIDDLETOWN"],
        "Brass Key":      ["BRASS"],
        "Long Beach":     ["LONG BEACH", "ALLEGRIA"],
        "Westerly":       ["WESTERLY", "PLEASANT VIEW"],
        "Crown Point":    ["CROWN", "CROWNE"],
        "Ashworth":       ["ASHWORTH", "HAMPTON BEACH"],
        "Anchor Inn":     ["ANCHOR"],
        "Surfside":       ["SURFSIDE"],
        "1620":           ["1620", "PLYMOUTH"],
    },
    "Hilton": {
        "Northbrook":     ["NORTHBROOK"],
        "Andover":        ["ANDOVER"],
        "Ann Arbor":      ["ANN ARBOR"],
        "Silver Spring":  ["SILVER SPRING"],
        "Mesa":           ["MESA"],
        "Littleton":      ["LITTLETON"],
        "Nashua":         ["NASHUA"],
        "Kansas City":    ["KANSAS"],
        "Memphis":        ["MEMPHIS"],
    },
    "IHG": {
        "Salem":          ["SALEM"],
        "Manchester":     ["MANCHESTER"],
        "Weirton":        ["WEIRTON"],
    },
}

# Property codes as they appear in the SRP Activity export's 'Property -
# InnCode' column, which is how one shared export gets split per hotel.
# Confirmed from real exports; the three blanks have not appeared in one yet.
HILTON_INNCODES = {
    "Littleton":     "LTNNH",
    "Nashua":        "ASHSS",
    "Ann Arbor":     "ARBAA",
    "Mesa":          "MESWH",
    "Kansas City":   "MCIAP",
    "Memphis":       "MEMPH",
    # "Northbrook":  "?",   not yet seen in an export
    # "Andover":     "?",
    # "Silver Spring": "?",
}

# Hilton properties do not run a Strategy Report.
PORTFOLIO_WORKBOOKS = {
    "Stay In Touch": WORKBOOK_TYPES,
    "Hilton":        ["ROB", "Forecast"],
    "IHG":           WORKBOOK_TYPES,
}


def portfolio_of(hotel_display_name):
    """Which portfolio a Drive-discovered hotel belongs to, or None."""
    up = (hotel_display_name or "").upper()
    for pf, members in PORTFOLIO_HOTELS.items():
        for kws in members.values():
            if any(k in up for k in kws):
                return pf
    return None


def hotels_in_portfolio(portfolio, discovered):
    """[(display_name, folder_id)] from `discovered` belonging to `portfolio`,
    ordered as declared above so the checkbox list is stable between runs.
    Hotels that are declared but have no matching Drive folder are dropped —
    surface those separately rather than showing an entry that cannot run."""
    members = PORTFOLIO_HOTELS.get(portfolio, {})
    out = []
    for kws in members.values():
        for name, fid in discovered:
            if any(k in name.upper() for k in kws) and (name, fid) not in out:
                out.append((name, fid))
    return out


def portfolio_hotels_missing(portfolio, discovered):
    """Declared hotels with no matching Drive folder."""
    members = PORTFOLIO_HOTELS.get(portfolio, {})
    names = [n.upper() for n, _ in discovered]
    return [label for label, kws in members.items()
            if not any(any(k in n for k in kws) for n in names)]

# Maps workbook type → partial filename keyword to search for in Drive
WORKBOOK_KEYWORDS = {
    "ROB":             "ROB",
    "Strategy Report": "STRATEGY",
    "Forecast":        "FORECAST",
}

CREDS_PATH = "credentials.json.json"
SCOPES     = ["https://www.googleapis.com/auth/drive"]


@st.cache_resource
def get_drive_service():
    # Production (Streamlit Cloud): credentials stored in st.secrets["google_credentials"]
    if "google_credentials" in st.secrets:
        import json
        creds = service_account.Credentials.from_service_account_info(
            dict(st.secrets["google_credentials"]), scopes=SCOPES
        )
    else:
        # Local dev: read from file
        creds = service_account.Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


@st.cache_data(ttl=3600)
def has_edit_access(folder_id, email="workbook-updater-live@linchris-rob-updater.iam.gserviceaccount.com"):
    """Check if the given email has editor access to a folder.
    Returns True if editor, False if viewer-only or no access.
    Cached for 1 hour to avoid repeated API calls.
    """
    try:
        svc = get_drive_service()
        perms = svc.permissions().list(
            fileId=folder_id, fields="permissions(emailAddress,role)",
            supportsAllDrives=True
        ).execute()
        for perm in perms.get("permissions", []):
            if perm.get("emailAddress") == email:
                return perm.get("role") == "owner" or perm.get("role") == "organizer" or perm.get("role") == "editor"
        return False
    except Exception:
        return False


@st.cache_data(ttl=3600)
def get_hotels_with_edit_access():
    """Return list of hotels with edit access. Cached for 1 hour."""
    all_hotels = get_hotels_from_drive()
    filtered = []
    for hotel_name, folder_id in all_hotels:
        # Handle MULTI_ID format
        if folder_id.startswith(MULTI_ID_PREFIX):
            ids = folder_id[len(MULTI_ID_PREFIX):].split(",")
            if any(has_edit_access(fid) for fid in ids):
                filtered.append((hotel_name, folder_id))
        else:
            if has_edit_access(folder_id):
                filtered.append((hotel_name, folder_id))
    return filtered


def drive_find_folder_by_keyword(service, keyword, parent_id=None):
    """Return the first folder whose name contains keyword (case-insensitive)."""
    q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    if parent_id:
        q += " and '%s' in parents" % parent_id
    result = service.files().list(
        q=q, fields="files(id, name)", pageSize=100,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    for f in result.get("files", []):
        if keyword.lower() in f["name"].lower():
            return f["id"], f["name"]
    return None, None


def _pick_rev_reports_candidate(candidates, year_kw, month_kw):
    """Rank REVENUE REPORTS folder candidates for a target month/year.
    CRITICAL: Prefer year-only folders (2026 REVENUE REPORTS) over month
    folders, even for the target month. Month folders should only be used
    if they're the ONLY folders available (flat per-month layout like Harbor).
    Order of preference:
      1. name contains the year but NO month abbreviation — a true year
         folder (month subfolders live inside it) [NESTED LAYOUT];
      2. name contains target month AND no other month abbr [FLAT LAYOUT];
      3. name contains year but NO month abbr (if step 2 didn't find anything);
      4. name contains the year at all (fallback);
      5. first candidate.
    """
    if not candidates:
        return None
    month_kw_2digit = month_kw[:3] + month_kw[-2:] if month_kw else None

    # Step 1: Prefer year-only folders (nested layout)
    year_only = []
    for f in candidates:
        name_upper = f["name"].upper()
        if year_kw in f["name"] and not any(m.upper() in name_upper for m in MONTH_ABBR):
            year_only.append(f)
    if year_only:
        return year_only[0]

    # Step 2: If no year-only folders, try exact month match (flat per-month layout)
    for f in candidates:
        name_upper = f["name"].upper()
        if month_kw and (month_kw in name_upper or (month_kw_2digit and month_kw_2digit in name_upper)):
            return f

    # Step 3-5: Fallback rankings
    for f in candidates:
        if year_kw in f["name"]:
            return f
    return candidates[0]


def _find_rev_reports_folder_for_year(service, hotel_id, year_kw, month_kw=None):
    """Find the REVENUE REPORTS folder to use for a given year (and,
    preferably, the specific target month — see _pick_rev_reports_candidate
    for the ranking; the old first-year-substring-match behavior silently
    picked whichever same-year sibling Drive happened to list first for
    hotels with one folder per month).
    Some hotels have ONE 'REVENUE REPORTS' folder for years directly (year/month
    subfolders inside it); others have a SEPARATE '<year> REVENUE REPORTS <hotel>'
    folder per year (or per month), sitting side by side.
    """
    # Hotel has separate REVENUE REPORTS folders shared per year/month (each
    # individually shared, since there's no common parent to share instead).
    if hotel_id.startswith(MULTI_ID_PREFIX):
        candidate_ids = hotel_id[len(MULTI_ID_PREFIX):].split(",")
        candidates = []
        for cid in candidate_ids:
            try:
                info = service.files().get(fileId=cid, fields="name", supportsAllDrives=True).execute()
                candidates.append({"id": cid, "name": info["name"]})
            except Exception:
                continue
        best = _pick_rev_reports_candidate(candidates, year_kw, month_kw)
        if best:
            return best["id"], best["name"]
        return None, None

    q = ("mimeType = 'application/vnd.google-apps.folder' and trashed = false "
         "and '%s' in parents") % hotel_id
    children = service.files().list(
        q=q, fields="files(id, name)", pageSize=100,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute().get("files", [])
    # Match "revenue reports", "rev reports", or just "revenue" in folder name
    candidates = [f for f in children
                  if any(kw in f["name"].lower() for kw in ["revenue reports", "rev reports", "revenue"])]
    best = _pick_rev_reports_candidate(candidates, year_kw, month_kw)
    if best:
        return best["id"], best["name"]

    # Some hotels are shared directly at the REVENUE REPORTS folder level
    # (Shared Drive permissions don't propagate to a parent folder) — hotel_id
    # IS that folder already in that case, not its parent, so there's no
    # child to find. Check hotel_id's own name before giving up.
    try:
        self_info = service.files().get(
            fileId=hotel_id, fields="name", supportsAllDrives=True
        ).execute()
        if "revenue reports" in self_info.get("name", "").lower():
            return hotel_id, self_info["name"]
    except Exception:
        pass

    return None, None


def _find_month_folder_under_rev(service, rev_id, year_kw, month_kw, target_month, hotel_name):
    """Locate the month folder for a new-month setup, handling both layouts:
    month folders directly inside the REVENUE REPORTS folder (common when
    there's one REVENUE REPORTS folder per year), or nested under a year
    subfolder. NEVER creates a folder — the month folder is expected to
    already exist (created ahead of time as part of the standard process).
    Returns (None, None) if it can't be found, so the caller can error out
    with a clear message instead of the app silently creating a new one in
    the wrong place.

    Finding the "year subfolder" is trickier than a plain keyword search:
    a sibling MONTH folder (e.g. "JUL2026 REVENUE REPORTS X") also contains
    the year digits as a substring, so a naive year_kw match can pick a
    sibling month folder instead of a real year folder. Excluding any
    candidate whose name contains a month abbreviation avoids that ambiguity.
    """
    # Flat per-month layout (confirmed real case: Provincetown Harbor Hotel,
    # 'H: AUG2026 REVENUE REPORTS ...'): the rev folder's own name already
    # names the target month — it IS the month folder; there is no month
    # subfolder inside it to find.
    try:
        rev_info = service.files().get(fileId=rev_id, fields="name", supportsAllDrives=True).execute()
        rev_name = rev_info.get("name", "")
        month_kw_2digit = month_kw[:3] + month_kw[-2:]
        if month_kw in rev_name.upper() or month_kw_2digit in rev_name.upper():
            return rev_id, rev_name
    except Exception:
        pass

    month_id, month_name = drive_find_folder_by_keyword(service, month_kw, parent_id=rev_id)
    if month_id:
        return month_id, month_name

    q = ("mimeType = 'application/vnd.google-apps.folder' and trashed = false "
         "and '%s' in parents") % rev_id
    siblings = service.files().list(
        q=q, fields="files(id, name)", pageSize=100,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute().get("files", [])
    year_id = None
    for f in siblings:
        name_upper = f["name"].upper()
        if year_kw in f["name"] and not any(m.upper() in name_upper for m in MONTH_ABBR):
            year_id = f["id"]
            break
    if year_id:
        return drive_find_month_folder(service, year_id, month_kw)

    # Fallback: search recursively through nested folders (handles deep nesting like SALEM)
    # Some hotels have multiple intermediate folders before reaching the month folder
    def search_recursive(parent_id, depth=0):
        if depth > 3:  # Prevent infinite recursion
            return None, None
        try:
            q = ("mimeType = 'application/vnd.google-apps.folder' and trashed = false "
                 "and '%s' in parents") % parent_id
            children = service.files().list(
                q=q, fields="files(id, name)", pageSize=100,
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute().get("files", [])

            best_match = None
            for child in children:
                child_name_upper = child["name"].upper()
                # Prefer folders with both month_kw and "REVENUE" (more specific)
                # to avoid stopping at intermediate folders
                if month_kw in child_name_upper and "REVENUE" in child_name_upper:
                    return child["id"], child["name"]
                # Also check for "REPORTS" to catch "AUG2026 REVENEU REPORTS SALEM"
                if month_kw in child_name_upper and "REPORTS" in child_name_upper:
                    best_match = (child["id"], child["name"])
                # Recurse into subfolders to find deeper matches
                result_id, result_name = search_recursive(child["id"], depth + 1)
                if result_id:
                    return result_id, result_name
            if best_match:
                return best_match
        except Exception:
            pass
        return None, None

    return search_recursive(rev_id)


def drive_find_file(service, keyword, parent_id):
    """Return (file_id, file_name) for first xlsx whose name contains keyword,
    excluding files whose name also contains 'copy' (to skip backup copies).

    Also matches native Google Sheets (mimeType 'application/vnd.google-apps.
    spreadsheet') — confirmed real case: Hotel 1620's Forecast workbook was
    created directly as a Google Sheet rather than uploaded as .xlsx, so an
    xlsx/xlsm-only filter silently missed it even though the name matched.
    drive_download handles exporting these to xlsx bytes on read.
    """
    q = ("'%s' in parents and trashed = false "
         "and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
         "or mimeType = 'application/vnd.ms-excel.sheet.macroenabled.12' "
         "or mimeType = 'application/vnd.google-apps.spreadsheet')") % parent_id
    result = service.files().list(
        q=q, fields="files(id, name)",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = result.get("files", [])
    # prefer non-copy files; fall back to copy if nothing else found
    files.sort(key=lambda f: (1 if "copy" in f["name"].lower() else 0))
    for f in files:
        name_lower = f["name"].lower()
        if keyword.lower() in name_lower and "master" not in name_lower:
            return f["id"], f["name"]
    return None, None


def drive_download(service, file_id) -> bytes:
    """Download a file's bytes. Native Google Sheets (created directly in
    Drive rather than uploaded as .xlsx — confirmed real case: Hotel 1620's
    Forecast workbook) can't be read via get_media like a normal blob file;
    they must be exported to xlsx format instead."""
    meta = service.files().get(fileId=file_id, fields="mimeType", supportsAllDrives=True).execute()
    buf = io.BytesIO()
    if meta.get("mimeType") == "application/vnd.google-apps.spreadsheet":
        req = service.files().export_media(
            fileId=file_id,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    dl  = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()


def drive_upload(service, file_id, file_bytes: bytes, file_name: str):
    """Overwrite an existing Drive file with new bytes."""
    buf   = io.BytesIO(file_bytes)
    media = MediaIoBaseUpload(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    # supportsAllDrives is required here even though drive_download's
    # get_media() works without it — confirmed real case: writing back to a
    # Shared Drive file (Hyannis Anchor In) 404'd on update() alone, for
    # every workbook type, despite the same file_id having just been read
    # successfully moments earlier.
    service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()


def drive_find_month_folder(service, parent_id: str, month_kw: str):
    """Return (folder_id, folder_name) for the month folder under parent_id,
    or (None, None) if it isn't there. Never creates a folder — the month
    folder is expected to already exist as part of the standard process;
    the caller is responsible for erroring out clearly when it's missing
    instead of the app silently creating one (which is how folders ended
    up nested in the wrong place before).
    """
    q = ("mimeType='application/vnd.google-apps.folder' and trashed=false "
         "and '%s' in parents") % parent_id
    result = service.files().list(
        q=q, fields="files(id,name)", pageSize=100,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    for f in result.get("files", []):
        if month_kw in f["name"].upper():
            return f["id"], f["name"]
    return None, None


def drive_copy_file(service, source_file_id: str, new_name: str, parent_folder_id: str):
    """Copy a Drive file to a new name in the given folder. Returns (new_file_id, new_name).
    supportsAllDrives=True is required when the destination is a shared drive —
    service accounts have no personal storage quota.
    """
    body = {"name": new_name, "parents": [parent_folder_id]}
    copied = service.files().copy(
        fileId=source_file_id, body=body, fields="id,name",
        supportsAllDrives=True,
    ).execute()
    return copied["id"], copied["name"]


def _hotel_search_scope_ids(service, hotel_id):
    """Return every folder id that could plausibly hold a hotel's MASTER
    template files: each of the hotel's own root candidate folder(s)
    (unwrapping a MULTI:<id>,<id>,... group) plus their direct children.

    Confirmed real case: master-file lookups used to search ALL of Drive
    with no folder scoping at all — Wolfeboro's 'Set Up New ROB' found and
    copied Hotel 1620's ROB master (Drive's unordered global search just
    happened to return it first), silently mislabeling the result 'JUL2026
    ROB PLYMOUTH.xlsx' inside Wolfeboro's own folder.
    """
    if hotel_id.startswith(MULTI_ID_PREFIX):
        root_ids = hotel_id[len(MULTI_ID_PREFIX):].split(",")
    else:
        root_ids = [hotel_id]

    scope_ids = list(root_ids)
    for rid in root_ids:
        q = ("mimeType = 'application/vnd.google-apps.folder' and trashed = false "
             "and '%s' in parents") % rid
        try:
            children = service.files().list(
                q=q, fields="files(id)", pageSize=100,
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute().get("files", [])
        except Exception:
            children = []
        scope_ids.extend(c["id"] for c in children)
    return scope_ids


def find_rob_master(service, hotel_id: str):
    """Search the hotel's own Drive tree for the ROB master file."""
    scope_ids = _hotel_search_scope_ids(service, hotel_id)
    if not scope_ids:
        return None, "Could not resolve hotel folder to search."
    parent_clause = " or ".join("'%s' in parents" % sid for sid in scope_ids)
    q = ("trashed=false and (%s) "
         "and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
         "or mimeType='application/vnd.ms-excel.sheet.macroenabled.12') "
         "and name contains 'MASTER' and name contains 'ROB'") % parent_clause
    result = service.files().list(
        q=q, fields="files(id,name,parents)", pageSize=50,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    for f in result.get("files", []):
        return f["id"], f["name"]
    return None, "No ROB master file found in Drive."


def _is_rob_month_blank(ws, block_start):
    """Return True if cols 2,3,4 of the Revenue row are all empty (None, '', or formula)."""
    rev_row = block_start + 1
    return all(
        ws.cell(rev_row, c).value is None
        or ws.cell(rev_row, c).value == ""
        or is_formula(str(ws.cell(rev_row, c).value))
        for c in [2, 3, 4]
    )


def _rob_as_of_date(year, month):
    """Return the 1st of the month for a given year, advanced past Sat/Sun to Monday."""
    d = datetime.date(year, month, 1)
    if d.weekday() == 5:    # Saturday → Monday
        d += datetime.timedelta(days=2)
    elif d.weekday() == 6:  # Sunday → Monday
        d += datetime.timedelta(days=1)
    return d


def _resolve_cell(prev_wb_data, prev_wb_formulas, sheet_name, row, col):
    """Read a cell value, following simple cross-sheet formula references if needed.
    When data_only=True returns None (no cached value), reads the formula string
    from prev_wb_formulas and follows the reference (e.g. ='wk one'!B45)."""
    import re
    from openpyxl.utils import column_index_from_string

    def _read_data(wb, sname, r, c):
        if wb and sname in wb.sheetnames:
            return wb[sname].cell(r, c).value
        return None

    v = _read_data(prev_wb_data, sheet_name, row, col)
    if v is not None:
        return v

    # No cached value — try to follow the formula reference
    if prev_wb_formulas and sheet_name in prev_wb_formulas.sheetnames:
        formula = prev_wb_formulas[sheet_name].cell(row, col).value
        if formula and str(formula).startswith("="):
            m = re.match(r"^='?([^'!]+)'?!([A-Za-z]+)(\d+)$", str(formula).strip())
            if m:
                ref_sheet = m.group(1)
                ref_col   = column_index_from_string(m.group(2))
                ref_row   = int(m.group(3))
                v = _read_data(prev_wb_data, ref_sheet, ref_row, ref_col)
                if v is not None:
                    return v
                # Referenced sheet might itself have a formula — follow one more level
                v2 = _resolve_cell(prev_wb_data, prev_wb_formulas, ref_sheet, ref_row, ref_col)
                return v2
    return None


def _fill_rob_prev_table(wk1_ws, prev_wb, prev_wb_formulas, target_month):
    """Fill the 'Week 1 Previous Sheet - CALCULATION ONLY' table in wk1.
    Scans dynamically for header, year columns, and month rows.
    Pulls Revenue from last completed week tab of prev ROB."""

    month_abbrs = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"]

    def _as_year(v):
        """Try to extract a 4-digit year from a cell value (int, float, str, date)."""
        if isinstance(v, (datetime.datetime, datetime.date)):
            return v.year
        if isinstance(v, (int, float)):
            iv = int(v)
            if 2000 <= iv <= 2100:
                return iv
            # Excel date serial — convert
            if 40000 <= iv <= 60000:
                try:
                    d = datetime.date(1899, 12, 30) + datetime.timedelta(days=iv)
                    return d.year
                except Exception:
                    pass
        if isinstance(v, str):
            s = v.strip()
            if s.isdigit() and 2000 <= int(s) <= 2100:
                return int(s)
        return None

    # ── 1. Find header cell ───────────────────────────────────────────────────
    hdr_row = hdr_col = None
    for r in range(1, min(wk1_ws.max_row + 1, 60)):
        for c in range(1, wk1_ws.max_column + 1):
            val = str(wk1_ws.cell(r, c).value or "").strip().lower()
            if "week 1 previous" in val or ("calculation only" in val and "week" in val):
                hdr_row, hdr_col = r, c
                break
        if hdr_row:
            break
    if not hdr_row:
        return "Week 1 Previous Sheet table not found in wk one"

    years_in_order = [target_month.year - 3, target_month.year - 2,
                      target_month.year - 1, target_month.year]

    # ── 2. Find month label column first (Jan/Feb… to the LEFT of year cols) ─
    month_label_col = None
    for r in range(hdr_row + 1, hdr_row + 35):
        for c in range(max(1, hdr_col - 5), hdr_col + 5):
            v = str(wk1_ws.cell(r, c).value or "").strip().lower()
            if v in month_abbrs:
                month_label_col = c
                break
        if month_label_col:
            break
    if not month_label_col:
        return "Could not find month label column in Week 1 Previous Sheet table"

    # ── 3. Collect all month rows using that exact column ─────────────────────
    dest_month_row = {}  # month_idx (0-based) → row
    for r in range(hdr_row + 1, hdr_row + 35):
        v = str(wk1_ws.cell(r, month_label_col).value or "").strip().lower()
        if v in month_abbrs:
            dest_month_row[month_abbrs.index(v)] = r
    if not dest_month_row:
        return "Could not find month rows in Week 1 Previous Sheet table"

    # ── 4. Find year columns — scan to the RIGHT of month label col ───────────
    year_start_col = month_label_col + 1
    dest_year_col = {}  # year (int) → col in new wk1

    for dr in range(1, 8):
        # Try parsing actual year values
        parsed = {}
        for c in range(year_start_col, year_start_col + 20):
            yr = _as_year(wk1_ws.cell(hdr_row + dr, c).value)
            if yr:
                parsed[yr] = c
        if parsed:
            dest_year_col = parsed
            break
        # Fallback: find non-empty cells positionally (formulas count as non-empty)
        non_empty = []
        for c in range(year_start_col, year_start_col + 20):
            v = wk1_ws.cell(hdr_row + dr, c).value
            if v is not None and str(v).strip():
                non_empty.append(c)
        if len(non_empty) >= 2:
            for i, c in enumerate(non_empty[:4]):
                dest_year_col[years_in_order[i]] = c
            break

    if not dest_year_col:
        return "Could not find year columns in Week 1 Previous Sheet table"

    # ── 4. Build ordered list of week sheet names (most recent first) ────────
    wk_order = ["wk six", "wk five", "wk four", "wk three", "wk two", "wk one"]
    wk_sheet_names = []  # sheet names in prev ROB, most-recent first
    for wk_try in wk_order:
        matches = [s for s in prev_wb.sheetnames if wk_try in s.lower()]
        if matches:
            wk_sheet_names.append(matches[0])
    if not wk_sheet_names:
        return "No week tabs found in previous ROB"

    # ── 5. Year → source col: scan wk one row 4 for dates/years ─────────────
    base_year = target_month.year
    src_year_col = {base_year - 3: 2, base_year - 2: 3, base_year - 1: 4, base_year: 5}

    wk1_sheet_name = wk_sheet_names[-1]  # wk one is last in most-recent-first list
    ref_ws = prev_wb[wk1_sheet_name]
    detected = {}
    for c in range(1, min(ref_ws.max_column + 1, 20)):
        yr = _as_year(ref_ws.cell(4, c).value)
        if yr and 2000 <= yr <= 2100:
            detected[yr] = c
    if len(detected) >= 3:
        src_year_col = detected

    # ── 6. Write Revenue values — check each cell individually ───────────────
    # For each (month, year) cell: walk week tabs most-recent→oldest.
    # Each cell is checked on its own — some are hardcoded numbers, some are
    # formulas pointing elsewhere. _resolve_cell handles both cases.
    for month_idx, dest_row in dest_month_row.items():
        rev_row = 4 + 8 * month_idx + 1
        for year, dest_col in dest_year_col.items():
            src_col = src_year_col.get(year)
            if not src_col:
                continue
            v = None
            for sheet_name in wk_sheet_names:
                v = _resolve_cell(prev_wb, prev_wb_formulas, sheet_name, rev_row, src_col)
                if v is not None and not is_formula(str(v)):
                    break  # found a real value in this tab — use it
                v = None   # reset if None or formula string slipped through
            if v is not None:
                wk1_ws.cell(dest_row, dest_col).value = v

    return None


def _fill_rob_sheet(new_ws, prev_ws, ly_ws, target_month, is_wk_one, wk_one_sheet_name):
    """Fill one ROB sheet tab with historical data.
    Preserves formulas from master template — only overwrites cells with values.
    """
    from openpyxl.utils import get_column_letter

    target_idx  = target_month.month - 1   # 0-based (Jul = 6)
    prev_idx    = target_idx - 1           # most recently completed month (Jun = 5)
    # LY col → new col shift: LY has [2022,2023,2024,2025], new needs [2023,2024,2025,2026]
    ly_to_new    = {3: 2, 4: 3, 5: 4}
    # Rows per month block, read off each sheet rather than assumed (8 for most
    # hotels, 11 where there's a Permanent-rooms section). Last year's workbook
    # is measured separately so a template change between years can't shift the
    # rows we read from.
    step         = rob_block_step(new_ws)
    ly_step      = rob_block_step(ly_ws) if ly_ws is not None else step
    data_offsets = list(range(1, step))    # offset 0 (date header) handled separately

    # ── As-of dates for this week tab ──────────────────────────────────────
    # Taken from row 4 of THIS sheet's counterpart in last year's ROB, so
    # AUG2026 'wk two' gets its 2023/2024/2025 reporting dates from AUG2025
    # 'wk two'. Each week was reported roughly seven days apart, so these must
    # never be shared between tabs.
    #
    # Row 4 specifically: it holds the literal dates, while every month block
    # below it chains upward (=B4, =B52, ...). Reading a lower block instead
    # picks up the chain formula itself on most sheets — and on some it picks
    # up whatever else is parked in that cell (one real file has a revenue
    # figure sitting in the Aug block's date position on 'wk two'). Requiring
    # a real date type keeps that kind of stray number out of the header.
    #
    # Col 5 (current year) is a placeholder; the weekly update stamps the real
    # as-of date when it runs.
    as_of_dates = {}
    if ly_ws:
        for ly_col, new_col in ly_to_new.items():
            v = ly_ws.cell(4, ly_col).value
            # A never-used week tab carries Excel's zero date (1899-12-30),
            # which is a real datetime and would otherwise be copied across as
            # though it were a reporting date.
            if isinstance(v, (datetime.datetime, datetime.date)) and v.year >= 2000:
                as_of_dates[new_col] = v
    as_of_dates[5] = datetime.datetime(target_month.year, target_month.month, target_month.day)

    for month_idx in range(12):
        block_start = 4 + step * month_idx

        # ── Date header row (offset 0) ─────────────────────────────────────
        # Weeks 2+ used to get ='wk one'!<col><row> here, which forced every
        # week tab to display wk one's reporting dates.
        #
        # Row 4 always gets the real date. Month blocks below it keep a
        # same-sheet chain (=B4, =SUM(B12)) because that correctly carries
        # THIS tab's date down the sheet — but a cross-sheet reference is
        # overwritten, since pointing at another week's tab reintroduces the
        # very problem being fixed (one real template chains its lower blocks
        # to ='wk one'!B11).
        for col, date_val in as_of_dates.items():
            cell = new_ws.cell(block_start, col)
            cur = str(cell.value) if cell.value is not None else ""
            same_sheet_chain = is_formula(cur) and "!" not in cur
            if block_start != 4 and same_sheet_chain:
                continue
            cell.value = date_val

        # ── Data rows (offsets 1–7) ───────────────────────────────────────────
        if month_idx < prev_idx:
            # Past months (Jan–May when target=Jul): copy all 4 cols from prev ROB
            # For new hotels with no historical data (e.g., first year), prev_ws will be None —
            # just skip and leave those months with master template values (no error).
            if not _is_rob_month_blank(new_ws, block_start):
                continue
            if is_wk_one:
                if prev_ws is None:
                    continue
                for dr in data_offsets:
                    r = block_start + dr
                    for c in [2, 3, 4, 5]:
                        # Skip if master template has a formula in this cell
                        if is_formula(str(new_ws.cell(r, c).value)):
                            continue
                        v = prev_ws.cell(r, c).value
                        if v is not None and not is_formula(str(v)) and not is_datelike(v):
                            new_ws.cell(r, c).value = v
            else:
                for dr in data_offsets:
                    r = block_start + dr
                    for c in [2, 3, 4, 5]:
                        # Skip if master template has a formula in this cell
                        if is_formula(str(new_ws.cell(r, c).value)):
                            continue
                        col_ltr = get_column_letter(c)
                        new_ws.cell(r, c).value = f"='{wk_one_sheet_name}'!{col_ltr}{r}"

        elif month_idx == prev_idx:
            # Prev month (Jun when target=Jul):
            # Cols 2,3,4 = historical years from LY ROB (same source as Jul+)
            # Col 5     = actual current-year data from prev ROB (built up weekly)
            if not _is_rob_month_blank(new_ws, block_start):
                continue
            if is_wk_one:
                if ly_ws:
                    for dr in data_offsets:
                        r = block_start + dr
                        for ly_col, new_col in ly_to_new.items():
                            # Skip if master template has a formula in this cell
                            if is_formula(str(new_ws.cell(r, new_col).value)):
                                continue
                            v = ly_ws.cell(r, ly_col).value
                            if v is not None and not is_formula(str(v)) and not is_datelike(v):
                                new_ws.cell(r, new_col).value = v
                    ly_sec_col = find_secondary_col(ly_ws, block_start) or 7
                    for dr in [4, 5, 6]:
                        r = block_start + dr
                        # Skip if master template has a formula in this cell
                        if is_formula(str(new_ws.cell(r, 8).value)):
                            continue
                        v = ly_ws.cell(r, ly_sec_col).value
                        if v is not None and not is_formula(str(v)) and not is_datelike(v):
                            new_ws.cell(r, 8).value = v
            else:
                for dr in data_offsets:
                    r = block_start + dr
                    for c in [2, 3, 4, 5]:
                        # Skip if master template has a formula in this cell
                        if is_formula(str(new_ws.cell(r, c).value)):
                            continue
                        col_ltr = get_column_letter(c)
                        new_ws.cell(r, c).value = f"='{wk_one_sheet_name}'!{col_ltr}{r}"

        else:
            # Current month and future months (Jul+): cols 2,3,4 from LY ROB
            # For new hotels with no historical data (e.g., Foxberry), ly_ws will be None —
            # just skip and leave those months with master template values (no error).
            if ly_ws is None:
                continue
            for dr in data_offsets:
                r = block_start + dr
                for ly_col, new_col in ly_to_new.items():
                    # Skip if master template has a formula in this cell
                    if is_formula(str(new_ws.cell(r, new_col).value)):
                        continue
                    v = ly_ws.cell(r, ly_col).value
                    if v is not None and not is_formula(str(v)) and not is_datelike(v):
                        new_ws.cell(r, new_col).value = v
            ly_sec_col = find_secondary_col(ly_ws, block_start) or 7
            for dr in [4, 5, 6]:
                r = block_start + dr
                # Skip if master template has a formula in this cell
                if is_formula(str(new_ws.cell(r, 8).value)):
                    continue
                v = ly_ws.cell(r, ly_sec_col).value
                if v is not None and not is_formula(str(v)) and not is_datelike(v):
                    new_ws.cell(r, 8).value = v


def setup_new_rob_month(service, hotel_id: str, hotel_name: str, target_month: datetime.date):
    """Full ROB new-month setup.
    Returns (new_file_name, error_or_warn_str, new_file_id, original_bytes).
    `original_bytes` is the file's content exactly as it was before this
    run's changes (None on any early-exit failure) — the caller can use it
    to offer an "undo"/reset-to-original for this setup, the same way the
    Strategy Report setup already does.
    """
    year_kw  = str(target_month.year)
    month_kw = target_month.strftime("%b%Y").upper()

    # ── Locate month folder (never created — must already exist) ────────────
    rev_id, rev_name = _find_rev_reports_folder_for_year(service, hotel_id, year_kw, month_kw)
    if not rev_id:
        return None, "No REVENUE REPORTS folder.", None, None
    month_id, month_name = _find_month_folder_under_rev(service, rev_id, year_kw, month_kw, target_month, hotel_name)
    if not month_id:
        return None, f"Could not find the {month_kw} folder for {hotel_name} — it should already exist.", None, None

    # Diagnostic: which folders are being used
    folder_diagnostic = f"Revenue Reports folder: {rev_name}; Month folder: {month_name}"

    # ── Find or copy the file ─────────────────────────────────────────────────
    # Diagnostic: list every file in month_id whose name contains "ROB" —
    # confirmed real risk this session (Wolfeboro's duplicate-named folders)
    # that two files with the same display name could sit in the same
    # folder, with drive_find_file silently picking one while a human
    # browsing to the file by name lands on the other.
    dup_check_warnings = []
    try:
        dup_q = ("'%s' in parents and trashed = false and name contains 'ROB'") % month_id
        dup_files = service.files().list(
            q=dup_q, fields="files(id,name)", pageSize=20,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute().get("files", [])
        if len(dup_files) > 1:
            dup_check_warnings.append(
                f"Multiple files matching 'ROB' found in the target month folder: "
                + ", ".join(f"'{f['name']}' (id: {f['id']})" for f in dup_files)
            )
    except Exception:
        pass

    existing_id, existing_name = drive_find_file(service, "ROB", month_id)
    is_fresh_copy = False
    if existing_id and "master" not in existing_name.lower():
        new_file_id, new_file_name = existing_id, existing_name
    else:
        is_fresh_copy = True
        master_id, master_name = find_rob_master(service, hotel_id)
        if not master_id:
            return None, master_name, None, None
        hotel_suffix = hotel_name.upper()
        name_upper   = master_name.upper()
        if "ROB" in name_upper:
            after = master_name[name_upper.find("ROB") + 3:].strip()
            after = after.replace(".xlsx","").replace(".xlsm","").replace(".XLSX","").replace(".XLSM","").strip()
            if after:
                hotel_suffix = after
        ext = ".xlsm" if master_name.lower().endswith(".xlsm") else ".xlsx"
        new_file_name = f"{month_kw} ROB {hotel_suffix}{ext}"
        try:
            new_file_id, new_file_name = drive_copy_file(service, master_id, new_file_name, month_id)
        except Exception as e:
            return None, str(e), None, None

    # ── Load all three workbooks ──────────────────────────────────────────────
    new_wb_bytes = drive_download(service, new_file_id)
    new_wb = openpyxl.load_workbook(io.BytesIO(new_wb_bytes), data_only=False)
    if is_fresh_copy:
        clear_tab_colors(new_wb, ROB_SHEETS)

    # Resolution/load failures here used to be swallowed silently — the
    # ROB workbook would just come back with July onward blank and no
    # indication why. Both lookup failures (resolve_drive_workbook returning
    # an error) and load failures (download/openpyxl exceptions) are now
    # captured into `warnings` and surfaced to the caller alongside the
    # success message, matching the diagnostic the SR flow already shows.
    warnings = [
        f"Target file id: {new_file_id} (name: {new_file_name})",
        f"Folder placement: {folder_diagnostic}",
    ] + dup_check_warnings

    prev_month_dt = (target_month - datetime.timedelta(days=1)).replace(day=1)
    prev_result, prev_err = resolve_drive_workbook(service, hotel_id, hotel_name, "ROB",
                                                     month_date=prev_month_dt)
    prev_wb = None
    prev_wb_formulas = None
    if prev_result:
        warnings.append(f"Prev month ({prev_month_dt.strftime('%b %Y')}) resolved to: {prev_result[1]}")
        try:
            prev_bytes = drive_download(service, prev_result[0])
            prev_wb = openpyxl.load_workbook(io.BytesIO(prev_bytes), data_only=True)
            prev_wb_formulas = openpyxl.load_workbook(io.BytesIO(prev_bytes), data_only=False)
        except Exception as e:
            warnings.append(f"Prev month ({prev_month_dt.strftime('%b %Y')}) workbook found but failed to load: {e}")
    else:
        warnings.append(f"Prev month ({prev_month_dt.strftime('%b %Y')}) not found: {prev_err}")

    ly_month_dt = target_month.replace(year=target_month.year - 1)
    ly_result, ly_err = resolve_drive_workbook(service, hotel_id, hotel_name, "ROB",
                                                month_date=ly_month_dt)
    ly_wb = None
    if ly_result:
        warnings.append(f"Last year ({ly_month_dt.strftime('%b %Y')}) resolved to: {ly_result[1]}")
        try:
            ly_wb = openpyxl.load_workbook(
                io.BytesIO(drive_download(service, ly_result[0])), data_only=True)
        except Exception as e:
            warnings.append(f"Last year ({ly_month_dt.strftime('%b %Y')}) workbook found but failed to load: {e}")
    else:
        warnings.append(f"Last year ({ly_month_dt.strftime('%b %Y')}) not found — future months' historical "
                         f"columns will be blank: {ly_err}")

    # Sheet lookups below are exact, case-sensitive matches against
    # ROB_SHEETS ("wk one", "wk two", ...) — a workbook can load successfully
    # above and still contribute nothing if its own tab names don't match
    # that exact casing/spacing (e.g. "Wk One", "WK1"). Checking only "wk
    # one" previously missed this: a source workbook can have a correctly
    # named "wk one" (so wk one fills fine) while missing/mismatching "wk
    # two" through "wk six" — confirmed real case where wk one populated but
    # wk two silently didn't. Check every ROB_SHEETS name, not just the first.
    for label, wb_obj in [("Prev month", prev_wb), ("Last year", ly_wb)]:
        if wb_obj is None:
            continue
        missing = [s for s in ROB_SHEETS if s not in wb_obj.sheetnames]
        if missing:
            warnings.append(
                f"{label} workbook is missing tabs: {missing} — "
                f"actual tabs: {wb_obj.sheetnames}"
            )

    # ── Fill each sheet ───────────────────────────────────────────────────────
    wk_one_name = ROB_SHEETS[0]
    for sheet_name in ROB_SHEETS:
        if sheet_name not in new_wb.sheetnames:
            continue
        new_ws  = new_wb[sheet_name]
        prev_ws = prev_wb[sheet_name] if prev_wb and sheet_name in prev_wb.sheetnames else None
        ly_ws   = ly_wb[sheet_name]   if ly_wb   and sheet_name in ly_wb.sheetnames   else None
        is_wk_one = (sheet_name == wk_one_name)
        _fill_rob_sheet(new_ws, prev_ws, ly_ws, target_month, is_wk_one, wk_one_name)

    # Direct readback of the target month's own block, right after the fill
    # loop and before save/upload — confirms whether the in-memory write
    # actually happened per sheet, isolating "fill logic didn't write it
    # for this tab" from "something after this point (save/upload) lost
    # it". Checked across all six week tabs, not just WKONE — confirmed
    # real case where wk one's fill succeeded but wk two's silently didn't.
    # Also check Jan/Feb/Mar/Dec to verify backward/future months are filled.
    target_idx = target_month.month - 1
    target_block_start = 4 + 8 * target_idx
    readback = {
        s: new_wb[s].cell(target_block_start + 1, 2).value
        for s in ROB_SHEETS if s in new_wb.sheetnames
    }
    wk_one_sheet = ROB_SHEETS[0]
    if wk_one_sheet in new_wb.sheetnames:
        ws = new_wb[wk_one_sheet]
        month_check = {}
        for m, mlabel in [(1, "Jan"), (2, "Feb"), (3, "Mar"), (12, "Dec")]:
            block = 4 + 8 * (m - 1)
            month_check[mlabel] = ws.cell(block + 1, 2).value
        warnings.append(
            f"Months check (wk one only) — Jan/Feb/Mar/Dec Revenue (col B): {month_check!r}"
        )
    warnings.append(
        f"Readback check — {target_month.strftime('%b %Y')} Revenue (row "
        f"{target_block_start + 1}, col B) per tab after fill: {readback!r}"
    )

    # ── Fill Week 1 Previous Sheet table in wk one ───────────────────────────
    if prev_wb and wk_one_name in new_wb.sheetnames:
        err = _fill_rob_prev_table(new_wb[wk_one_name], prev_wb, prev_wb_formulas, target_month)
        if err:
            warnings.append(f"Prev table: {err}")

    strip_tables(new_wb)
    out = io.BytesIO()
    new_wb.save(out)
    drive_upload(service, new_file_id, out.getvalue(), new_file_name)
    warn_str = "; ".join(warnings) if warnings else None
    return new_file_name, warn_str, new_file_id, new_wb_bytes


def find_forecast_master(service, hotel_id: str):
    """Search the hotel's own Drive tree for the Forecast master file (.xlsx or .xlsm)."""
    scope_ids = _hotel_search_scope_ids(service, hotel_id)
    if not scope_ids:
        return None, "Could not resolve hotel folder to search."
    parent_clause = " or ".join("'%s' in parents" % sid for sid in scope_ids)
    q = ("trashed=false and (%s) "
         "and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
         "or mimeType='application/vnd.ms-excel.sheet.macroenabled.12') "
         "and name contains 'MASTER' and name contains 'FORECAST'") % parent_clause
    result = service.files().list(
        q=q, fields="files(id,name,parents)", pageSize=50,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    for f in result.get("files", []):
        return f["id"], f["name"]
    return None, "No FORECAST master file found in Drive."


def setup_new_forecast_month(service, hotel_id: str, hotel_name: str, target_month: datetime.date):
    """
    Copy Forecast master → rename for target_month → place in month folder → set B4 date.
    Returns (new_file_name, error_str).
    """
    year_kw  = str(target_month.year)
    month_kw = target_month.strftime("%b%Y").upper()

    rev_id, _ = _find_rev_reports_folder_for_year(service, hotel_id, year_kw, month_kw)
    if not rev_id:
        return None, "No REVENUE REPORTS folder."

    month_id, _ = _find_month_folder_under_rev(service, rev_id, year_kw, month_kw, target_month, hotel_name)
    if not month_id:
        return None, f"Could not find the {month_kw} folder for {hotel_name} — it should already exist."

    # Check if Forecast already exists
    existing_id, existing_name = drive_find_file(service, "FORECAST", month_id)
    if existing_id and "master" not in existing_name.lower():
        return existing_name, None

    master_id, master_name = find_forecast_master(service, hotel_id)
    if not master_id:
        return None, master_name

    # Infer hotel suffix from master name
    hotel_suffix = hotel_name.upper()
    name_upper = master_name.upper()
    ext = ".xlsm" if master_name.lower().endswith(".xlsm") else ".xlsx"
    for kw in ("FORECAST",):
        if kw in name_upper:
            after = master_name[name_upper.find(kw) + len(kw):].strip()
            after = after.replace(".xlsx","").replace(".xlsm","").replace(".XLSX","").replace(".XLSM","").strip()
            if after:
                hotel_suffix = after
            break

    new_file_name = f"{month_kw} FORECAST {hotel_suffix}{ext}"
    try:
        new_file_id, created_name = drive_copy_file(service, master_id, new_file_name, month_id)
    except Exception as e:
        return None, str(e)

    # Set B4 = first day of target_month in FCST-WK1
    try:
        wb_bytes  = drive_download(service, new_file_id)
        wb        = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=False)
        clear_tab_colors(wb, FORECAST_SHEETS)
        sheet     = FORECAST_SHEETS[0] if FORECAST_SHEETS[0] in wb.sheetnames else wb.sheetnames[1] if len(wb.sheetnames) > 1 else wb.sheetnames[0]
        ws        = wb[sheet]
        # Find "Day of Week" cell → one right + one down = start date cell
        date_cell_row = date_cell_col = None
        for r in range(1, 15):
            for c in range(1, 10):
                if "day of week" in str(ws.cell(r, c).value or "").lower():
                    date_cell_row = r + 1
                    date_cell_col = c + 1
                    break
            if date_cell_row:
                break
        if date_cell_row and date_cell_col:
            ws.cell(date_cell_row, date_cell_col).value = datetime.datetime(
                target_month.year, target_month.month, 1)
        strip_tables(wb)
        out = io.BytesIO()
        wb.save(out)
        drive_upload(service, new_file_id, out.getvalue(), created_name)
    except Exception as e:
        return created_name, f"Copied OK but could not set start date: {e}"

    return created_name, None


def find_sr_master(service, hotel_id: str):
    """Search the hotel's own Drive tree for the SR master file.
    Returns (file_id, file_name) or (None, error_str).
    """
    scope_ids = _hotel_search_scope_ids(service, hotel_id)
    if not scope_ids:
        return None, "Could not resolve hotel folder to search."
    parent_clause = " or ".join("'%s' in parents" % sid for sid in scope_ids)
    q = ("trashed=false and (%s) "
         "and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
         "and name contains 'MASTER' and name contains 'STRATEGY'") % parent_clause
    result = service.files().list(
        q=q, fields="files(id,name,parents)", pageSize=50,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    for f in result.get("files", []):
        return f["id"], f["name"]
    return None, "No STRATEGY master file found in Drive."


def setup_new_sr_month(service, hotel_id: str, hotel_name: str, target_month: datetime.date):
    """
    New month SR setup:
      1. Find SR master, copy it, rename to [MON][YEAR] STRATEGY [HOTEL].xlsx
      2. Find or create month folder under the year folder
      3. Move the copy into that folder
    Returns (new_file_name, error_str).
    """
    year_kw = str(target_month.year)
    month_kw = target_month.strftime("%b%Y").upper()

    # Resolve REVENUE REPORTS and month folder
    rev_id, rev_name = _find_rev_reports_folder_for_year(service, hotel_id, year_kw, month_kw)
    if not rev_id:
        return None, "No REVENUE REPORTS folder."

    # Find month folder — never created, must already exist
    month_id, month_name = _find_month_folder_under_rev(service, rev_id, year_kw, month_kw, target_month, hotel_name)
    if not month_id:
        return None, f"Could not find the {month_kw} folder for {hotel_name} — it should already exist. (Revenue Reports folder: {rev_name})"

    # Check if SR already exists in that folder
    existing_id, existing_name = drive_find_file(service, "STRATEGY", month_id)
    if existing_id and "master" not in existing_name.lower():
        return existing_name, None  # already set up

    # Find master
    master_id, master_name = find_sr_master(service, hotel_id)
    if not master_id:
        return None, master_name  # error string

    # Infer hotel suffix from master file name for the new file name
    # e.g. "MASTER 2026 STRATEGY PLYMOUTH.xlsx" → "PLYMOUTH"
    hotel_suffix = hotel_name.upper()
    name_upper = master_name.upper().replace(".XLSX", "")
    if "STRATEGY" in name_upper:
        after = name_upper[name_upper.find("STRATEGY") + len("STRATEGY"):].strip()
        if after:
            hotel_suffix = master_name[master_name.upper().find("STRATEGY") + len("STRATEGY"):].strip().replace(".xlsx", "").replace(".XLSX", "").strip()

    new_file_name = f"{month_kw} STRATEGY {hotel_suffix}.xlsx"
    try:
        _, created_name = drive_copy_file(service, master_id, new_file_name, month_id)
    except Exception as e:
        return None, str(e)
    return created_name, None


def get_prev_month_otb_trans(service, hotel_id: str, hotel_name: str, current_month: datetime.date):
    """Pull OTB TY Trans values for current_month dates from previous month's SR last filled tab.
    Returns {date: value} or {} on any failure.
    """
    prev_month = (current_month.replace(day=1) - datetime.timedelta(days=1)).replace(day=1)
    result, err = resolve_drive_workbook(service, hotel_id, hotel_name, "Strategy Report", month_date=prev_month)
    if err or not result:
        return {}

    file_id, _ = result
    try:
        file_bytes = drive_download(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception:
        return {}

    # Find last filled (colored) tab
    last_filled_sheet = None
    for sheet_name in STRATEGY_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        tab_color = ws.sheet_properties.tabColor
        if tab_color is not None:
            last_filled_sheet = sheet_name

    if not last_filled_sheet:
        return {}

    ws = wb[last_filled_sheet]
    col_map = detect_strategy_columns(ws)
    otb_col = col_map.get("otb_trans")
    date_col = detect_date_column(ws, wb=wb)
    if not otb_col:
        return {}

    result_map = {}
    for r in range(5, ws.max_row + 1):
        date_val = ws.cell(r, date_col).value
        if isinstance(date_val, datetime.datetime):
            d = date_val.date()
        elif isinstance(date_val, datetime.date):
            d = date_val
        else:
            continue
        if d < current_month:
            continue
        cell_val = ws.cell(r, otb_col).value
        if cell_val is not None and not is_formula(cell_val):
            result_map[d] = safe_float(cell_val)

    return result_map


def resolve_drive_workbook(service, hotel_id: str, hotel_name: str, workbook_type: str, month_date: datetime.date = None):
    """
    Walk Drive to find the target workbook. Handles two folder structures:
      A) Hotel > MMMYYYY REVENUE REPORTS HOTEL > files  (month in folder name, files direct)
      B) Hotel > REVENUE REPORTS > Year > Month > files (nested year/month subfolders)
    Returns ((file_id, file_name), None) or (None, error_message).
    Never touches files whose name contains 'master'.

    A "hotel" can also be a MULTI:<id>,<id>,... group — several candidate
    root folders sharing one dropdown entry, either because they're each a
    flat per-year/month folder shared directly (Hyannis Anchor In) or full
    duplicate top-level hotel folders from historical typos/copies
    (confirmed real case: "Provinceetown Surfside", "Provincertown
    Surfside", "Surfside (1)" all being the same hotel). Each candidate is
    resolved fully (structures A/B/C, recursing into its own REVENUE REPORTS
    child if it has one) rather than guessed from its name alone — whichever
    candidate actually contains the target file wins.
    """
    if month_date is None:
        month_date = datetime.date.today()

    month_kw        = month_date.strftime("%b%Y").upper()
    month_kw_2digit = month_date.strftime("%b%y").upper()
    year_kw         = str(month_date.year)
    wb_keyword      = WORKBOOK_KEYWORDS[workbook_type]

    def _list_subfolders(parent_id):
        q = (f"'{parent_id}' in parents and trashed = false and "
             f"mimeType = 'application/vnd.google-apps.folder'")
        return service.files().list(
            q=q, fields="files(id, name)", pageSize=100,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute().get("files", [])

    def _find_file_in(folder_id, folder_name):
        fid, fname = drive_find_file(service, wb_keyword, folder_id)
        if not fid:
            return None, f"No '{wb_keyword}' workbook found in '{folder_name}'."
        if "master" in fname.lower():
            return None, f"Resolved file '{fname}' looks like a master doc — aborting."
        return (fid, fname), None

    def _resolve_single(single_id, single_name):
        """Resolve within ONE candidate root — either a full hotel-parent
        folder (with its own REVENUE REPORTS > year > month nesting) or a
        folder that IS already the REVENUE REPORTS level directly (detected
        from its own name). Returns ((file_id, file_name), None) or (None, err).
        """
        self_is_rev = _is_rev_reports_name(single_name)

        # A0: candidate IS already the month-level folder — its own name
        # contains both "REVENUE REPORTS" and the target month, so the file
        # is directly inside it, not further down. Confirmed real case:
        # Tybee's month folders are individually named "G: JUL2026 REVENUE
        # REPORTS TYBEE"; because that name also contains "REVENUE REPORTS"
        # it gets swept into the same MULTI: candidate group as Tybee's root
        # and year-level folders (get_hotels_from_drive merges by extracted
        # hotel name, not by structural level). Without this check,
        # self_is_rev below always assumes further year/month descent is
        # needed, finds no subfolders (children are files here), and fails
        # even though this candidate is the correct one.
        if self_is_rev and month_kw in single_name.upper():
            result = _find_file_in(single_id, single_name)
            if result[0]:
                return result

        children = _list_subfolders(single_id)

        # A: Hotel > MMMYYYY REVENUE REPORTS HOTEL > file
        if not self_is_rev:
            a = next((f for f in children
                       if _is_rev_reports_name(f["name"]) and month_kw in f["name"].upper()), None)
            if a:
                result = _find_file_in(a["id"], a["name"])
                if result[0]:
                    return result

        # B: Hotel > REVENUE REPORTS > MMMYYYY ... > file
        if self_is_rev:
            rev = {"id": single_id, "name": single_name}
        else:
            rev = next((f for f in children
                        if _is_rev_reports_name(f["name"]) and year_kw in f["name"]), None)
            if not rev:
                rev = next((f for f in children if _is_rev_reports_name(f["name"])), None)
        if rev:
            rev_children = children if self_is_rev else _list_subfolders(rev["id"])
            b1 = next((f for f in rev_children if month_kw in f["name"].upper()), None)
            if b1:
                result = _find_file_in(b1["id"], b1["name"])
                if result[0]:
                    return result
            b2_year = next((f for f in rev_children if year_kw in f["name"].upper()), None)
            if b2_year:
                b2_month_id, b2_month_name = drive_find_folder_by_keyword(
                    service, month_kw, parent_id=b2_year["id"])
                if b2_month_id:
                    result = _find_file_in(b2_month_id, b2_month_name)
                    if result[0]:
                        return result

        # C: Hotel > Year > Month > file  (no REVENUE REPORTS wrapper)
        c_year = next((f for f in children
                       if year_kw in f["name"].upper()
                       and not _is_rev_reports_name(f["name"])), None)
        if c_year:
            c_month_id, c_month_name = drive_find_folder_by_keyword(
                service, month_kw, parent_id=c_year["id"])
            if c_month_id:
                result = _find_file_in(c_month_id, c_month_name)
                if result[0]:
                    return result

        return None, f"Could not find '{month_kw}' workbook under '{single_name}'."

    if hotel_id.startswith(MULTI_ID_PREFIX):
        candidate_ids = hotel_id[len(MULTI_ID_PREFIX):].split(",")
        candidates = []
        for cid in candidate_ids:
            try:
                info = service.files().get(fileId=cid, fields="name", supportsAllDrives=True).execute()
                candidates.append({"id": cid, "name": info["name"]})
            except Exception:
                continue
        if not candidates:
            return None, f"Could not read any of the shared folders for '{hotel_name}'."

        # Try the most likely-named candidates first (pure ordering
        # optimization), but fall through to the next candidate if one
        # doesn't actually contain the file — e.g. a stale duplicate folder
        # from a typo — instead of giving up after the first name match.
        def _sort_key(f):
            name_up = f["name"].upper()
            if month_kw in name_up or month_kw_2digit in name_up:
                return 0
            if year_kw in name_up:
                return 1
            return 2
        ordered = sorted(candidates, key=_sort_key)

        last_err = None
        for cand in ordered:
            result, err = _resolve_single(cand["id"], cand["name"])
            if result:
                return result, None
            last_err = err
        return None, last_err or f"Could not find '{month_kw}' workbook for '{hotel_name}'."

    return _resolve_single(hotel_id, hotel_name)


DOW_ABBREVS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]


def _count_sheet_data_rows(ws):
    """Count how many data rows exist from row 5 downward by finding the last
    row that has any non-empty content across the first 10 columns."""
    last_row = 4
    for r in range(5, ws.max_row + 1):
        if any(ws.cell(r, c).value is not None for c in range(1, 11)):
            last_row = r
    return max(0, last_row - 4)  # number of rows starting from row 5


def restructure_sr_dates(wb, target_month):
    """Restructure the three date columns in every strategy sheet for a new month.
    Row 4 (header):    col 1 = LY year, col 3 = TY year — else stale years carry
                        over from whatever the master template last had.
    Col 1 (LY date):   starts at day 2 of target_month, last year
    Col 2 (day of wk): abbreviation matching the TY date in col 3
    Col 3 (TY date):   starts at day 1 of target_month, this year
    Row count matches the master — hotels open only part of the year have fewer rows.
    """
    ty_start = target_month
    ly_start = datetime.date(ty_start.year - 1, ty_start.month, 2)

    for sheet_name in STRATEGY_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        # Skip merged cells (can't set value on merged cells)
        try:
            ws.cell(4, 1).value = ly_start.year
        except (AttributeError, ValueError):
            pass
        try:
            ws.cell(4, 3).value = ty_start.year
        except (AttributeError, ValueError):
            pass
        num_rows = min(_count_sheet_data_rows(ws), 365)
        if num_rows == 0:
            continue
        for i in range(num_rows):
            row     = 5 + i
            ty_date = ty_start + datetime.timedelta(days=i)
            ly_date = ly_start + datetime.timedelta(days=i)
            dow     = DOW_ABBREVS[ty_date.weekday()]
            try:
                ws.cell(row, 1).value = datetime.datetime(ly_date.year, ly_date.month, ly_date.day)
            except (AttributeError, ValueError):
                pass
            try:
                ws.cell(row, 2).value = dow
            except (AttributeError, ValueError):
                pass
            try:
                ws.cell(row, 3).value = datetime.datetime(ty_date.year, ty_date.month, ty_date.day)
            except (AttributeError, ValueError):
                pass


def _load_wb_from_drive(svc, hotel_id, hotel_name, wb_type, month_date, data_only=True):
    """Download and parse a workbook from Drive. Returns openpyxl.Workbook or None.
    data_only=True (default) returns cached cell values — use for reference workbooks.
    data_only=False preserves formulas — use for workbooks we intend to write back.
    """
    result, err = resolve_drive_workbook(svc, hotel_id, hotel_name, wb_type, month_date=month_date)
    if err or not result:
        return None
    try:
        return openpyxl.load_workbook(io.BytesIO(drive_download(svc, result[0])), data_only=data_only)
    except Exception:
        return None


# ── Ancillary Revenue (Plymouth/Hotel 1620 only, for now) ────────────────────
# Different shape from ROB/SR/Forecast: one workbook with a tab per month
# (not a file per hotel per month), and within a month's tab, up to 5 "weeks"
# sit side by side as column-pairs (WK1=cols B/C ... WK5=cols J/K) rather than
# as separate sheets. The folder/file are hardcoded rather than resolved
# through the general multi-hotel machinery since there's currently exactly
# one of these across all of Drive.
ANCILLARY_REVENUE_FOLDER_NAME = "M - Ancillary Revenue files PLYMOUTH - TEST"
ANCILLARY_UPSELL_HEADER_ROW   = 3   # "Name" / "Total number" / "Total Revenue" ...
ANCILLARY_UPSELL_FIRST_ROW    = 4
ANCILLARY_WEEK_COLS = {1: (2, 3), 2: (4, 5), 3: (6, 7), 4: (8, 9), 5: (10, 11)}  # week -> (count_col, revenue_col)
_YELLOW_FILL_RGB = "FFFFFF00"


def find_ancillary_revenue_file(service):
    """Locate Plymouth/Hotel 1620's Ancillary Revenue workbook.
    Returns (file_id, file_name) or (None, error_str)."""
    q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    result = service.files().list(
        q=q, fields="files(id, name)", pageSize=1000,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    target = ANCILLARY_REVENUE_FOLDER_NAME.strip().lower()
    folder_id = next((f["id"] for f in result.get("files", [])
                       if f["name"].strip().lower() == target), None)
    if not folder_id:
        return None, (f"Drive folder '{ANCILLARY_REVENUE_FOLDER_NAME}' not found. "
                       f"Create it and share it with the service account, then try again.")

    q2 = ("'%s' in parents and trashed = false and "
          "(mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
          "or mimeType='application/vnd.ms-excel.sheet.macroenabled.12')") % folder_id
    result2 = service.files().list(
        q=q2, fields="files(id, name)", pageSize=20,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = result2.get("files", [])
    if not files:
        return None, f"No workbook found in '{ANCILLARY_REVENUE_FOLDER_NAME}'."
    return files[0]["id"], files[0]["name"]


def parse_addon_production_csv(file_bytes: bytes) -> dict:
    """Parse the wide-format 'Add-on Production' export: one row per
    upsell item, columns are Name, Total, Revenue, Avg, then a
    (count, revenue) pair per day, in date order starting from the 1st of
    whatever month the export covers.

    Returns {normalized_name: [(count, revenue), ...]} — the per-day list
    is in calendar order starting day 1, so summing any 7-item slice gives
    that week's totals regardless of which week is being filled.

    Uses Python's built-in csv module rather than pandas — confirmed real
    case: this export's rows don't all have the same number of trailing
    empty columns (the two summary rows at the top are shorter than the
    per-item rows below), and pd.read_csv errors out ("Expected N fields
    ... saw M") the moment row lengths aren't uniform. The stdlib reader
    has no such expectation; each row is just whatever fields it has.
    """
    text = file_bytes.decode("utf-8-sig")
    reader = list(csv.reader(io.StringIO(text)))
    out = {}
    for row in reader[3:]:   # row 0 = date-range summary, row 1 = blank col header, row 2 = day-header row
        if not row:
            continue
        name = str(row[0] or "").strip()
        if not name:
            continue
        daily = []
        # Columns 4+ (0-indexed) are (count, revenue) pairs, one per day.
        c = 4
        while c + 1 < len(row):
            count_raw = row[c] if c < len(row) else ""
            rev_raw   = row[c + 1] if c + 1 < len(row) else ""
            count = safe_float(count_raw)
            rev   = safe_float(rev_raw)
            if count is None and rev is None:
                break
            daily.append((count or 0, rev or 0))
            c += 2
        out[_normalize_ancillary_name(name)] = daily
    return out


def _normalize_ancillary_name(name: str) -> str:
    """Loosen an item name for matching between the CSV export and the
    sheet's own row labels — confirmed real case: the CSV has trailing
    spaces on several names ('Bottle of House Wine ') that the sheet
    doesn't, and casing isn't always identical either."""
    return re.sub(r"\s+", " ", name).strip().upper()


def find_next_available_week(ws, header_row=ANCILLARY_UPSELL_HEADER_ROW,
                              first_data_row=ANCILLARY_UPSELL_FIRST_ROW):
    """Return the week number (1-5) whose column-pair is marked yellow
    anywhere in the Upsell table's data rows — i.e. whichever week the
    sheet itself is flagging as next to fill in. Scans from the header row
    down to the first row whose col-A label is 'TOTALS'. Returns None if
    no yellow cells are found (nothing left to fill, or already done)."""
    last_row = first_data_row
    for r in range(first_data_row, first_data_row + 30):
        label = str(ws.cell(r, 1).value or "").strip().upper()
        if label == "TOTALS":
            break
        last_row = r
    for week, (count_col, rev_col) in ANCILLARY_WEEK_COLS.items():
        for r in range(first_data_row, last_row + 1):
            for c in (count_col, rev_col):
                fill = ws.cell(r, c).fill
                rgb = fill.fgColor.rgb if fill and fill.fgColor and fill.fgColor.type == "rgb" else None
                if rgb == _YELLOW_FILL_RGB:
                    return week
    return None


def build_ancillary_addon_change_plan(daily_by_name: dict, ws, week: int,
                                       header_row=ANCILLARY_UPSELL_HEADER_ROW,
                                       first_data_row=ANCILLARY_UPSELL_FIRST_ROW):
    """Build changes for one week's (count, revenue) columns in the Upsell
    table, summing the Add-on Production CSV's daily figures for days
    (week-1)*7+1 .. week*7 of the month against each sheet row's own label
    (matched via _normalize_ancillary_name — row order isn't assumed to
    match the CSV's row order).
    """
    count_col, rev_col = ANCILLARY_WEEK_COLS[week]
    day_start = (week - 1) * 7
    day_end   = day_start + 7  # exclusive

    changes = []
    r = first_data_row
    while True:
        label = str(ws.cell(r, 1).value or "").strip()
        if not label or label.upper() == "TOTALS":
            break
        key = _normalize_ancillary_name(label)
        daily = daily_by_name.get(key)
        if daily is not None:
            slice_ = daily[day_start:day_end]
            total_count = sum(d[0] for d in slice_)
            total_rev   = sum(d[1] for d in slice_)
            changes.append({"row": r, "col": count_col, "label": f"{label} — count", "new_value": total_count})
            changes.append({"row": r, "col": rev_col,   "label": f"{label} — revenue", "new_value": total_rev})
        r += 1
    return changes


def apply_ancillary_changes(ws, changes):
    for ch in changes:
        ws.cell(ch["row"], ch["col"]).value = ch["new_value"]


# ── Monthly OOO Report (Sell-Out Efficiency Report .xlsm) ────────────────────
# This workbook has a macro-driven "Save Report" shape button on every one of
# its 170+ dated tabs. Confirmed by direct round-trip test: loading it with
# openpyxl and saving it back — even completely untouched — silently drops
# every one of those shape/button objects (openpyxl's drawing writer doesn't
# preserve non-image/non-chart shapes). So this feature never opens the file
# through openpyxl for writing. It only uses openpyxl (read-only) to read
# values off the existing dated tabs, then adds the new summary tab via
# direct zip/XML surgery — copying every original zip entry byte-for-byte
# and touching only the 3 small manifest parts that must change to register
# a new sheet (workbook.xml, workbook.xml.rels, [Content_Types].xml). This
# was verified to leave xl/vbaProject.bin and every xl/drawings/*.xml
# identical to the source file.
OOO_TAB_RGB = "FF00FF00"  # bright green, per explicit request — distinct from
                          # the app's other "done" green (FF00B050) used elsewhere
OOO_REPORT_FILE_KEYWORD = "SELL-OUT EFFICIENCY REPORT"


def find_ooo_report_file(service):
    """Locate the company-wide Sell-Out Efficiency Report workbook (currently
    shared as 'COPY Sell-Out Efficiency Report.xlsm - TEST'). Searched by
    filename across all of Drive rather than a fixed folder, since there's
    exactly one of these (unlike the per-hotel workbooks). Returns
    (file_id, file_name) or (None, error_str) — errors out on zero or
    multiple matches rather than guessing which one is right."""
    q = ("trashed = false and "
         "(mimeType='application/vnd.ms-excel.sheet.macroenabled.12' "
         "or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')")
    result = service.files().list(
        q=q, fields="files(id, name)", pageSize=1000, corpora="allDrives",
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    candidates = [f for f in result.get("files", [])
                  if OOO_REPORT_FILE_KEYWORD in f["name"].strip().upper()]
    if not candidates:
        return None, (f"No workbook matching '{OOO_REPORT_FILE_KEYWORD.title()}' found anywhere "
                       f"in Drive — make sure it's been shared with the service account.")
    if len(candidates) > 1:
        names = ", ".join(f["name"] for f in candidates)
        return None, f"Found {len(candidates)} matching workbooks, expected exactly 1: {names}"
    return candidates[0]["id"], candidates[0]["name"]


def _parse_ooo_sheet_date(name):
    """Dated tabs are named like '07-28-2026' (M-D-YYYY, not always zero-
    padded). Anything that doesn't match — e.g. a stray old '10-22-22'
    2-digit-year tab found in a real file — is skipped rather than guessed."""
    m = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{4})$', name.strip())
    if not m:
        return None
    month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return datetime.date(year, month, day)
    except ValueError:
        return None


def list_ooo_available_months(wb):
    """Returns [(year, month, [sheet_names_in_date_order]), ...] sorted with
    the most recent month first. Derived from the tabs actually present in
    the uploaded file, never assumed from today's date."""
    months = {}
    for name in wb.sheetnames:
        d = _parse_ooo_sheet_date(name)
        if d:
            months.setdefault((d.year, d.month), []).append((d, name))
    out = []
    for (y, m), pairs in months.items():
        pairs.sort()
        out.append((y, m, [n for _, n in pairs]))
    out.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return out


def build_ooo_monthly_totals(wb, year, month, sheet_names):
    """Sums column J ('End. OOO Rooms', per the row-7 header) per property
    (column C) across every daily tab in the given month. Row range is read
    per-day (row 8 down until column C is blank) rather than assumed fixed,
    since not every hotel appears on every day. Headers are checked on each
    sheet before trusting the columns — a sheet that doesn't match the
    expected 'Property' / 'End...OOO' layout is skipped, not guessed at."""
    totals = {}
    order = []
    days_included = 0
    skipped = []
    for name in sheet_names:
        ws = wb[name]
        prop_header = str(ws.cell(row=7, column=3).value or "").strip().lower()
        ooo_header = str(ws.cell(row=7, column=10).value or "").strip().lower().replace("\n", " ")
        if "propert" not in prop_header or "end" not in ooo_header or "ooo" not in ooo_header:
            skipped.append(name)
            continue
        days_included += 1
        row = 8
        while True:
            prop = ws.cell(row=row, column=3).value
            if prop is None or str(prop).strip() == "":
                break
            prop = str(prop).strip()
            val = ws.cell(row=row, column=10).value
            try:
                val = float(val) if val is not None else 0.0
            except (TypeError, ValueError):
                val = 0.0
            if prop not in totals:
                totals[prop] = 0.0
                order.append(prop)
            totals[prop] += val
            row += 1
    return order, totals, days_included, skipped


# ── OOO Report: per-property ADR pulled from each hotel's own ROB ────────────
# Each hotel's ROB is a full-year "trailing" sheet: every WK tab (wk one..wk
# six) contains one 8-row block per calendar month (Jan..Dec), laid out as
# Month-header / Revenue / Room Nights / ADR / Group Rms sold / Group Rm Rev
# / Group Rm ADR / PICKUP WoW, starting at row 4 and repeating every 8 rows —
# confirmed against a real file: block_start = 4 + 8*(month-1), Revenue =
# block_start+1, Room Nights = block_start+2, ADR = block_start+3 (verified
# July 2026 block landing at rows 52-59, ADR at row 55, matching a real
# screenshot of Hotel 1620's ROB with July 2026 ADR highlighted at that exact
# position). Year columns (2023/2024/2025/2026...) are consistent across
# every month block in a sheet, but how many trailing years — and therefore
# which column is "this year" — varies per hotel, so it's detected from the
# one header row that reliably holds literal (non-formula) dates: Jan's.
_OOO_HOTEL_MATCH_EXCLUDE = ("TEST",)


def _find_rob_year_column(ws, target_year):
    """Jan's header row (row 4) holds literal dates; every other month's
    header is a formula chain off the previous month and its cached value is
    often stale/blank on files this app has re-saved with openpyxl (which
    never recalculates formulas) — so only Jan's row is trusted here."""
    for c in range(2, 12):
        v = ws.cell(4, c).value
        if isinstance(v, (datetime.datetime, datetime.date)) and v.year == target_year:
            return c
        if isinstance(v, (int, float)) and int(v) == target_year:
            return c
    return None


def read_rob_adr(wb_bytes, sheet_name, target_year, target_month):
    """ADR is computed here as Revenue / Room Nights rather than read from
    the sheet's own ADR formula cell — Revenue and Room Nights are written
    as literal values by this app (always reliable), while the ADR formula
    cell's cached result is frequently blank on files this app has re-saved,
    since openpyxl never recalculates formulas on save. Returns (adr_or_None,
    error_str_or_None)."""
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=True)
    if sheet_name not in wb.sheetnames:
        return None, f"Sheet '{sheet_name}' not found."
    ws = wb[sheet_name]
    year_col = _find_rob_year_column(ws, target_year)
    if year_col is None:
        year_col = 5  # fall back to column E, the typical "current year" column
    block_start = 4 + 8 * (target_month - 1)
    rev = safe_float(ws.cell(block_start + 1, year_col).value)
    rms = safe_float(ws.cell(block_start + 2, year_col).value)
    if rev is None or rms is None or rms == 0:
        month_label = datetime.date(target_year, target_month, 1).strftime("%b %Y")
        return None, f"No {month_label} Revenue/Room Nights on '{sheet_name}' (col {get_column_letter(year_col)})."
    return rev / rms, None


def match_ooo_property_to_hotel(prop_name, hotels):
    """hotels: [(name, id), ...] from get_hotels_from_drive(). Matches by
    substring in either direction after normalizing (uppercase, strip
    punctuation) — the same 'read the actual name, don't assume a fixed
    mapping' approach used for every other keyword match in this app — and
    falls back to the KNOWN_MULTI_FOLDER_HOTELS keyword registry, which
    catches display-name mismatches like 'Plymouth 1620' (OOO report) vs
    'Hotel 1620' (Drive) that share no substring but share a registered
    keyword. Excludes '-TEST' sandbox copies. Returns (hotel_name, hotel_id,
    error_note) — error_note is None on a clean single match, otherwise
    explains why nothing was matched (no match / ambiguous / no hotels
    visible at all) rather than silently guessing."""
    if not hotels:
        return None, None, ("the app can see 0 hotel folders in Drive — nothing is shared "
                            "with this environment's service account (or the Drive listing failed)")

    def norm(s):
        return re.sub(r'[^A-Z0-9 ]', '', s.upper()).strip()
    prop_norm = norm(prop_name)
    candidates = []
    for name, hid in hotels:
        if any(x in name.upper() for x in _OOO_HOTEL_MATCH_EXCLUDE):
            continue
        hotel_norm = norm(name)
        if prop_norm and (prop_norm in hotel_norm or hotel_norm in prop_norm):
            candidates.append((name, hid))
            continue
        # Keyword-registry fallback: the property name and the hotel's Drive
        # display name may each contain a registered keyword without either
        # being a substring of the other.
        for reg_name, kws in KNOWN_MULTI_FOLDER_HOTELS.items():
            if norm(reg_name) == hotel_norm or reg_name.upper() in name.upper():
                if any(kw in prop_norm for kw in kws):
                    candidates.append((name, hid))
                break
    if not candidates:
        return None, None, "no matching hotel found in Drive"
    if len(candidates) > 1:
        return None, None, f"ambiguous — matched {len(candidates)} hotels: {', '.join(c[0] for c in candidates)}"
    return candidates[0][0], candidates[0][1], None


def _find_ooo_adr_source_workbook(service, hotel_id, hotel_name):
    """Finds the most recent week with data in the hotel's ROB file.
    Scans current month first (WK6 -> WK1), then previous month if current is empty.
    Returns (wb_bytes, sheet_name, file_name, error_str)."""
    today = datetime.date.today()
    cur_month_dt  = today.replace(day=1)
    prev_month_dt = (cur_month_dt - datetime.timedelta(days=1)).replace(day=1)
    tried = []
    for month_dt in (cur_month_dt, prev_month_dt):
        result, err = resolve_drive_workbook(service, hotel_id, hotel_name, "ROB", month_dt)
        if not result:
            tried.append(f"{month_dt.strftime('%b %Y')} ROB: {err}")
            continue
        fid, fname = result
        wb_bytes = drive_download(service, fid)
        wb = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=False)
        # First pass: look for tab with any color (marked/highlighted)
        for sheet_name in reversed(ROB_SHEETS):
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            tc = ws.sheet_properties.tabColor
            # Any color means it's been marked/highlighted
            if tc is not None and getattr(tc, "rgb", None):
                return wb_bytes, sheet_name, fname, None
        # Second pass: if no tab marked done, use the first week with any data
        for sheet_name in reversed(ROB_SHEETS):
            if sheet_name not in wb.sheetnames:
                continue
            ws = wb[sheet_name]
            # Check if sheet has any OTB data (column B, rows 5-14)
            if any(ws.cell(r, 2).value for r in range(5, 15)):
                return wb_bytes, sheet_name, fname, None
        tried.append(f"{fname}: no week with data found")
    return None, None, None, "; ".join(tried) or "No ROB file found for the current or previous month."


def build_ooo_adr_lookup(service, order, hotels, year, month):
    """For each property in `order`, resolves it to a Drive hotel and reads
    its most-recently-available ADR for (year, month) off that hotel's ROB.
    Returns {property: (adr_or_None, note_or_None)} — note explains why a
    property has no ADR (unmatched hotel, ambiguous match, no filled ROB
    week found, etc.) so gaps are visible rather than silently blank."""
    out = {}
    for prop in order:
        hotel_name, hotel_id, match_note = match_ooo_property_to_hotel(prop, hotels)
        if not hotel_name:
            out[prop] = (None, match_note)
            continue
        wb_bytes, sheet_name, file_name, src_err = _find_ooo_adr_source_workbook(service, hotel_id, hotel_name)
        if not wb_bytes:
            out[prop] = (None, src_err)
            continue
        adr, adr_err = read_rob_adr(wb_bytes, sheet_name, year, month)
        out[prop] = (adr, adr_err)
    return out


def _ooo_extract_style_ids(zin, wb_xml, rels_xml):
    """Sample the newest dated tab's cell styles so the summary tab visually
    matches the daily report — navy title bar, bordered bold headers, bordered
    property names, blue bordered numbers, blue $-formatted rates. Style
    indices are read from the actual sheet XML at build time (never
    hardcoded): cellXfs order isn't guaranteed stable across future edits to
    the workbook, same reason column positions are re-detected per sheet
    everywhere else in this app. Returns a dict of style ids (values may be
    None if a sample cell wasn't found — those cells just render unstyled)."""
    sample_path = None
    for m in re.finditer(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb_xml):
        if _parse_ooo_sheet_date(m.group(1)):
            rel = re.search(r'Id="%s"[^>]*Target="(worksheets/sheet\d+\.xml)"' % m.group(2), rels_xml)
            if rel:
                sample_path = 'xl/' + rel.group(1)
                break
    ids = {"title": None, "title_bar": None, "header_left": None,
           "header": None, "prop": None, "num": None, "money": None}
    if not sample_path:
        return ids
    try:
        sx = zin.read(sample_path).decode('utf-8')
    except KeyError:
        return ids

    def s_of(ref):
        m = re.search(r'<c [^>]*r="%s"[^>]*?s="(\d+)"' % ref, sx)
        if not m:  # s= can precede r= depending on the writer
            m = re.search(r'<c [^>]*s="(\d+)"[^>]*?r="%s"' % ref, sx)
        return m.group(1) if m else None

    ids["title"]       = s_of("C5")   # navy bar, bold white text
    ids["title_bar"]   = s_of("D5")   # navy bar continuation
    ids["header_left"] = s_of("C7")   # bold, bordered, left ("Property")
    ids["header"]      = s_of("D7")   # bold, bordered, centered
    ids["prop"]        = s_of("C8")   # property name, bordered
    ids["num"]         = s_of("F8")   # blue centered #,##0, bordered
    ids["money"]       = s_of("L8")   # blue centered $#,##0.00, bordered
    return ids


def _ooo_build_new_sheet_xml(month_label, order, totals, days_included,
                             adr_by_property=None, style_ids=None):
    adr_by_property = adr_by_property or {}
    sid = style_ids or {}

    def _s(kind):
        v = sid.get(kind)
        return f' s="{v}"' if v else ''

    def cell_str(ref, text, kind=None):
        return (f'<c r="{ref}"{_s(kind)} t="inlineStr">'
                f'<is><t xml:space="preserve">{escape(str(text))}</t></is></c>')

    def cell_num(ref, val, kind=None):
        return f'<c r="{ref}"{_s(kind)}><v>{val}</v></c>'

    def cell_empty(ref, kind):
        return f'<c r="{ref}"{_s(kind)}/>'

    # Row 1: navy title bar across A-D, matching the daily report's banner.
    rows_xml = [('<row r="1">'
                 + cell_str("A1", f"Monthly OOO Rooms Summary - {month_label}", "title")
                 + cell_empty("B1", "title_bar") + cell_empty("C1", "title_bar")
                 + cell_empty("D1", "title_bar") + '</row>')]
    r = 3
    rows_xml.append(f'<row r="{r}">'
                    + cell_str(f"A{r}", "Property", "header_left")
                    + cell_str(f"B{r}", "Total End. OOO Rooms", "header")
                    + cell_str(f"C{r}", "ADR", "header")
                    + cell_str(f"D{r}", "Revenue", "header") + '</row>')
    r += 1
    grand_total = 0.0
    grand_revenue = 0.0
    for prop in order:
        v = totals[prop]
        grand_total += v
        adr, _note = adr_by_property.get(prop, (None, None))
        adr_cell = cell_num(f"C{r}", round(adr, 2), "money") if adr is not None else cell_empty(f"C{r}", "prop")
        if adr is not None:
            revenue = adr * v
            grand_revenue += revenue
            rev_cell = cell_num(f"D{r}", round(revenue, 2), "money")
        else:
            rev_cell = cell_empty(f"D{r}", "prop")
        rows_xml.append(f'<row r="{r}">'
                        + cell_str(f"A{r}", prop, "prop")
                        + cell_num(f"B{r}", v, "num")
                        + adr_cell + rev_cell + '</row>')
        r += 1
    rows_xml.append(f'<row r="{r}">'
                    + cell_str(f"A{r}", "TOTAL", "header_left")
                    + cell_num(f"B{r}", grand_total, "num")
                    + cell_empty(f"C{r}", "prop")
                    + cell_num(f"D{r}", round(grand_revenue, 2), "money") + '</row>')
    r += 1
    rows_xml.append(f'<row r="{r}">{cell_str(f"A{r}", f"({days_included} daily reports included)")}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheetPr><tabColor rgb="{OOO_TAB_RGB}"/></sheetPr>'
        f'<dimension ref="A1:D{r}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        '<cols><col min="1" max="1" width="30" customWidth="1"/><col min="2" max="2" width="22" customWidth="1"/>'
        '<col min="3" max="3" width="14" customWidth="1"/><col min="4" max="4" width="16" customWidth="1"/></cols>'
        '<sheetData>' + "".join(rows_xml) + '</sheetData>'
        '</worksheet>'
    )


def inject_ooo_monthly_sheet(original_bytes, year, month, order, totals, days_included, adr_by_property=None):
    """Adds the bright-green monthly summary tab to the OOO report, placed
    immediately to the right of the 'Report' tab (i.e. 2nd tab on the sheet),
    via raw zip/XML surgery. Never opens the file through openpyxl for
    writing — see the module note above for why. Returns (new_bytes,
    tab_name, error_str). error_str is set (and new_bytes is None) if a tab
    with that name already exists, rather than silently duplicating it."""
    month_label = datetime.date(year, month, 1).strftime("%B %Y")
    tab_name = datetime.date(year, month, 1).strftime("%b %Y").upper()  # e.g. "JUL 2026"
    tab_name = re.sub(r'[:\\/?*\[\]]', '', tab_name)[:31]

    zin = zipfile.ZipFile(io.BytesIO(original_bytes))

    wb_xml = zin.read('xl/workbook.xml').decode('utf-8')
    if f'name="{escape(tab_name)}"' in wb_xml:
        return None, tab_name, (f"A tab named '{tab_name}' already exists in this workbook — "
                                 f"delete it manually first if you want to regenerate it.")

    sheet_nums = [int(m.group(1)) for m in
                  (re.match(r'xl/worksheets/sheet(\d+)\.xml$', n) for n in zin.namelist()) if m]
    new_sheet_num = max(sheet_nums) + 1 if sheet_nums else 1
    new_sheet_path = f"xl/worksheets/sheet{new_sheet_num}.xml"

    existing_sheet_ids = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb_xml)]
    new_sheet_id = max(existing_sheet_ids) + 1 if existing_sheet_ids else 1

    rels_xml = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    existing_rids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
    new_rid = f"rId{max(existing_rids) + 1 if existing_rids else 1}"

    style_ids = _ooo_extract_style_ids(zin, wb_xml, rels_xml)
    new_sheet_xml = _ooo_build_new_sheet_xml(month_label, order, totals, days_included,
                                             adr_by_property, style_ids)

    new_sheet_el = f'<sheet state="visible" name="{escape(tab_name)}" sheetId="{new_sheet_id}" r:id="{new_rid}"/>'
    # Goes right after the "Report" tab (2nd position), not at the very front —
    # falls back to right after whichever tab is first if no "Report" tab is found,
    # rather than guessing at a name match.
    report_match = re.search(r'<sheet[^>]+name="Report"[^>]*/>', wb_xml)
    if not report_match:
        report_match = re.search(r'<sheet[^>]*/>', wb_xml)
    if report_match:
        insert_pos = report_match.end()
        wb_xml_new = wb_xml[:insert_pos] + new_sheet_el + wb_xml[insert_pos:]
    else:
        wb_xml_new = wb_xml.replace('<sheets>', '<sheets>' + new_sheet_el, 1)

    new_rel_el = (f'<Relationship Id="{new_rid}" '
                  f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                  f'Target="worksheets/sheet{new_sheet_num}.xml"/>')
    rels_xml_new = rels_xml.replace('</Relationships>', new_rel_el + '</Relationships>', 1)

    ct_xml = zin.read('[Content_Types].xml').decode('utf-8')
    new_override = (f'<Override ContentType="application/vnd.openxmlformats-officedocument'
                    f'.spreadsheetml.worksheet+xml" PartName="/{new_sheet_path}"/>')
    ct_xml_new = ct_xml.replace('</Types>', new_override + '</Types>', 1)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/workbook.xml':
                data = wb_xml_new.encode('utf-8')
            elif item.filename == 'xl/_rels/workbook.xml.rels':
                data = rels_xml_new.encode('utf-8')
            elif item.filename == '[Content_Types].xml':
                data = ct_xml_new.encode('utf-8')
            zout.writestr(item, data)
        zout.writestr(new_sheet_path, new_sheet_xml)
    return out_buf.getvalue(), tab_name, None


# ── User accounts (self-serve requests + admin approval) ─────────────────────
# Streamlit Cloud's filesystem is ephemeral and st.secrets is read-only at
# runtime, so per-person accounts created through the app can't live in either
# — store them in a small JSON file in Drive instead, using the exact same
# download/upload plumbing already used for every workbook in this app.

APP_DATA_FOLDER_NAME = "Workbook Updater App Data"
USERS_FILE_NAME      = "users.json"


def _find_app_data_folder(service):
    """Find the shared App Data folder — matched case-insensitively against
    APP_DATA_FOLDER_NAME (confirmed real case: exact-match search missed a
    folder created as 'workbook updater app data', different casing than the
    original name). Must be created once and shared with the service
    account, the same way every hotel's Drive folder is — can live anywhere
    inside a Shared Drive, including nested inside a hotel's folder."""
    q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    result = service.files().list(
        q=q, fields="files(id, name)", pageSize=1000,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    target = APP_DATA_FOLDER_NAME.strip().lower()
    for f in result.get("files", []):
        if f["name"].strip().lower() == target:
            return f["id"]
    return None


def _find_or_create_users_file(service):
    """Return (file_id, error). error is a user-facing string when the App
    Data folder itself hasn't been created/shared yet — creates users.json
    inside it (empty) on first use otherwise."""
    folder_id = _find_app_data_folder(service)
    if not folder_id:
        return None, (f"Drive folder '{APP_DATA_FOLDER_NAME}' not found. Create it in "
                       f"Google Drive and share it with the service account (same as a "
                       f"hotel folder), then try again.")
    q = "'%s' in parents and trashed = false and name = '%s'" % (folder_id, USERS_FILE_NAME)
    result = service.files().list(
        q=q, fields="files(id, name)", pageSize=5,
        supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute()
    files = result.get("files", [])
    if files:
        return files[0]["id"], None
    empty = json.dumps({"users": []}).encode("utf-8")
    media = MediaIoBaseUpload(io.BytesIO(empty), mimetype="application/json", resumable=True)
    created = service.files().create(
        body={"name": USERS_FILE_NAME, "parents": [folder_id]},
        media_body=media, fields="id", supportsAllDrives=True,
    ).execute()
    return created["id"], None


def _load_users(service, file_id):
    """Return the list of user dicts from users.json (empty list if unreadable)."""
    try:
        raw = drive_download(service, file_id)
        return json.loads(raw.decode("utf-8")).get("users", [])
    except Exception:
        return []


def _save_users(service, file_id, users):
    payload = json.dumps({"users": users}, indent=2).encode("utf-8")
    drive_upload(service, file_id, payload, USERS_FILE_NAME)


# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Linchris Weekly Tools", layout="wide")


# ── Login gate ────────────────────────────────────────────────────────────────
def check_login(username: str, password: str):
    """Return (ok, is_admin). Checks the single admin account (Streamlit
    secrets) first, then falls back to approved entries in users.json."""
    admin_user = st.secrets["auth"]["username"]
    admin_hash = st.secrets["auth"]["password_hash"].encode()
    if username == admin_user and bcrypt.checkpw(password.encode(), admin_hash):
        return True, True

    try:
        svc = get_drive_service()
        file_id, err = _find_or_create_users_file(svc)
        if err:
            return False, False
        users = _load_users(svc, file_id)
    except Exception:
        return False, False

    for u in users:
        if (u.get("username") == username and u.get("status") == "approved"
                and bcrypt.checkpw(password.encode(), u.get("password_hash", "").encode())):
            return True, u.get("role") == "admin"
    return False, False


LOGIN_ENABLED = True
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["is_admin"] = False
    st.session_state["username"] = None

if LOGIN_ENABLED and not st.session_state["authenticated"]:
    st.title("Linchris Hotel Corporation")
    st.subheader("Please log in to continue")

    login_tab, request_tab = st.tabs(["Log In", "Request Access"])

    with login_tab:
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log In")
            if submitted:
                ok, is_admin = check_login(username, password)
                if ok:
                    st.session_state["authenticated"] = True
                    st.session_state["is_admin"] = is_admin
                    st.session_state["username"] = username
                    st.rerun()
                else:
                    st.error("Incorrect username or password — or your account is still pending admin approval.")

    with request_tab:
        st.caption("Create your own login. An admin needs to approve it before you can log in.")
        with st.form("request_access_form"):
            req_username     = st.text_input("Choose a username", key="req_username")
            req_display_name = st.text_input("Your name", key="req_display_name")
            req_password     = st.text_input("Choose a password", type="password", key="req_password")
            req_confirm      = st.text_input("Confirm password", type="password", key="req_confirm")
            req_submitted    = st.form_submit_button("Request Access")
            if req_submitted:
                admin_user = st.secrets["auth"]["username"]
                uname = req_username.strip()
                if not uname or not req_password:
                    st.error("Username and password are required.")
                elif req_password != req_confirm:
                    st.error("Passwords don't match.")
                elif uname.lower() == admin_user.lower():
                    st.error("That username is taken.")
                else:
                    try:
                        svc = get_drive_service()
                        file_id, err = _find_or_create_users_file(svc)
                        if err:
                            st.error(err)
                        else:
                            existing_users = _load_users(svc, file_id)
                            if any(u.get("username", "").lower() == uname.lower() for u in existing_users):
                                st.error("That username is already taken or pending approval.")
                            else:
                                existing_users.append({
                                    "username":      uname,
                                    "display_name":  req_display_name.strip(),
                                    "password_hash": bcrypt.hashpw(req_password.encode(), bcrypt.gensalt()).decode(),
                                    "status":        "pending",
                                    "role":          "editor",
                                    "requested_at":  datetime.datetime.now().isoformat(),
                                    "decided_at":    None,
                                })
                                _save_users(svc, file_id, existing_users)
                                st.success("Request submitted — an admin needs to approve it before you can log in.")
                    except Exception as e:
                        st.error(f"Could not submit request: {e}")

    st.stop()

st.markdown("""
<style>
  .block-container { max-width: 100% !important; padding-left: 2rem !important; padding-right: 2rem !important; }

  .stTabs [data-baseweb="tab-list"] { gap: 8px; border-bottom: 1px solid #E5E7EB; }
  .stTabs [data-baseweb="tab"] {
    background: #F1F3F5; color: #1E293B;
    border-radius: 6px 6px 0 0; padding: 8px 20px; font-weight: 600;
  }
  .stTabs [aria-selected="true"] {
    background: #2563EB !important; color: #FFFFFF !important;
    box-shadow: inset 0 -3px 0 #C9A84C;
  }
  div[data-testid="metric-container"] {
    background: #F8F9FA; border: 1px solid #E5E7EB; border-left: 3px solid #C9A84C;
    border-radius: 8px; padding: 12px;
  }
  button[data-testid="stBaseButton-pills"],
  button[data-testid="stBaseButton-pillsActive"] {
    border-radius: 8px !important;
    border: 1.5px solid #94A3B8 !important;
    font-size: 1.5rem !important;
    padding: 1rem 1.8rem !important;
  }
  button[data-testid="stBaseButton-pillsActive"] {
    background-color: #2563EB !important;
    border-color: #2563EB !important;
    color: #FFFFFF !important;
  }

  /* Larger, consistent widget labels + selectbox text app-wide */
  label[data-testid="stWidgetLabel"] p { font-size: 1.2rem !important; }
  div[data-testid="stSelectbox"] div[data-baseweb="select"] div { font-size: 1.25rem !important; }
  [data-testid="stFileUploaderDropzone"] { font-size: 1.15rem !important; }
  [data-testid="stFileUploaderDropzoneInstructions"] span { font-size: 1.15rem !important; }
  ul[data-testid="stSelectboxVirtualDropdown"] li {
    font-size: 1.25rem !important;
    -webkit-font-smoothing: antialiased;
    text-rendering: optimizeLegibility;
  }
</style>
""", unsafe_allow_html=True)

if "view" not in st.session_state:
    st.session_state["view"] = "main"


def render_admin_settings(svc, users_file_id, users_err):
    st.title("Admin Settings")
    if st.button("← Back to app"):
        st.session_state["view"] = "main"
        st.rerun()

    if users_err:
        st.warning(f"Account requests unavailable: {users_err}")
        return

    all_users  = _load_users(svc, users_file_id)
    admin_user = st.secrets["auth"]["username"]

    st.subheader("Pending Requests")
    pending = [u for u in all_users if u.get("status") == "pending"]
    if not pending:
        st.caption("No pending requests.")
    for u in pending:
        c1, c2, c3, c4 = st.columns([3, 3, 1, 1])
        c1.write(f"**{u.get('username')}**")
        c2.write(u.get("display_name") or "—")
        if c3.button("Approve", key=f"approve_{u.get('username')}"):
            for uu in all_users:
                if uu.get("username") == u.get("username"):
                    uu["status"] = "approved"
                    uu["decided_at"] = datetime.datetime.now().isoformat()
            _save_users(svc, users_file_id, all_users)
            st.rerun()
        if c4.button("Reject", key=f"reject_{u.get('username')}"):
            all_users = [uu for uu in all_users if uu.get("username") != u.get("username")]
            _save_users(svc, users_file_id, all_users)
            st.rerun()

    st.divider()

    st.subheader("All Users")
    rows = [{"Username": admin_user, "Name": "—", "Role": "Admin", "Status": "approved"}]
    for u in all_users:
        rows.append({
            "Username": u.get("username"),
            "Name":     u.get("display_name") or "—",
            "Role":     "Admin" if u.get("role") == "admin" else "Editor",
            "Status":   u.get("status"),
        })
    st.dataframe(rows, use_container_width=True)

    approved = [u for u in all_users if u.get("status") == "approved"]
    if approved:
        st.caption("Change role / revoke access:")
        for u in approved:
            rc1, rc2, rc3 = st.columns([4, 2, 1])
            rc1.write(f"{u.get('username')} ({u.get('display_name') or '—'})")
            current_role = "Admin" if u.get("role") == "admin" else "Editor"
            new_role = rc2.selectbox("Role", ["Editor", "Admin"],
                                      index=["Editor", "Admin"].index(current_role),
                                      key=f"role_{u.get('username')}", label_visibility="collapsed")
            if new_role.lower() != u.get("role", "editor"):
                for uu in all_users:
                    if uu.get("username") == u.get("username"):
                        uu["role"] = new_role.lower()
                _save_users(svc, users_file_id, all_users)
                st.rerun()
            if rc3.button("Revoke", key=f"revoke_{u.get('username')}"):
                all_users = [uu for uu in all_users if uu.get("username") != u.get("username")]
                _save_users(svc, users_file_id, all_users)
                st.rerun()

    st.divider()

    st.subheader("Add New User")
    st.caption("Creates and approves an account directly — skips the request/approval step.")
    with st.form("admin_add_user_form"):
        new_username = st.text_input("Username", key="admin_new_username")
        new_display  = st.text_input("Name", key="admin_new_display")
        new_password = st.text_input("Password", type="password", key="admin_new_password")
        new_role     = st.selectbox("Role", ["Editor", "Admin"], key="admin_new_role",
                                     help="Editor: normal app access, no Admin Settings. Admin: full access including this page.")
        submitted = st.form_submit_button("Add User")
        if submitted:
            if not new_username or not new_password:
                st.error("Username and password are required.")
            elif new_username == admin_user or any(u.get("username") == new_username for u in all_users):
                st.error("That username is already taken.")
            else:
                all_users.append({
                    "username":      new_username,
                    "password_hash": bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode(),
                    "display_name":  new_display,
                    "role":          new_role.lower(),
                    "status":        "approved",
                    "requested_at":  datetime.datetime.now().isoformat(),
                    "decided_at":    datetime.datetime.now().isoformat(),
                })
                _save_users(svc, users_file_id, all_users)
                st.success(f"User '{new_username}' added and approved as {new_role}.")
                st.rerun()


title_col, toggle_col, profile_col = st.columns([5, 1, 1])
with title_col:
    st.title("Linchris Hotel Corporation — Weekly Update Tools")
with toggle_col:
    st.write("")
    test_mode = st.toggle("Test Mode", value=False, key="test_mode")
with profile_col:
    st.write("")
    with st.popover(f"👤 {st.session_state.get('username') or 'Account'}"):
        st.write(f"**{st.session_state.get('username')}**")
        st.caption("Admin" if st.session_state.get("is_admin") else "Editor")
        if st.session_state.get("is_admin"):
            if st.button("⚙️ Admin Settings", key="open_admin_settings"):
                st.session_state["view"] = "admin_settings"
                st.rerun()
        if st.button("Log Out", key="logout_btn"):
            st.session_state["authenticated"] = False
            st.session_state["is_admin"] = False
            st.session_state["username"] = None
            st.session_state["view"] = "main"
            st.rerun()

if st.session_state.get("view") == "admin_settings" and st.session_state.get("is_admin"):
    _admin_svc = None
    try:
        _admin_svc = get_drive_service()
        _users_file_id, _users_err = _find_or_create_users_file(_admin_svc)
    except Exception as e:
        _users_file_id, _users_err = None, str(e)
    render_admin_settings(_admin_svc, _users_file_id, _users_err)
    st.stop()

def render_hilton_update(hotels):
    """Hilton portfolio run.

    Differs from the other portfolios in two ways. Several properties are run
    at once, because a single SRP Activity export covers all of them and gets
    split by 'Property - InnCode'. And group figures come from a per-hotel
    Group Wash export rather than from the booking data, so each selected
    hotel needs its own wash file.

    Hilton properties do not run a Strategy Report.
    """
    if not hotels:
        st.info("No Hilton properties found in Drive.")
        return

    st.caption(
        "One SRP Activity export covers every property; each hotel also needs its "
        "own Group Wash export. Group and Permanent figures come from the Wash "
        "report — the booking export supplies only the transient remainder."
    )

    with st.container(border=True):
        st.markdown("**Properties to run**")
        cols = st.columns(3)
        selected = []
        for i, (name, fid) in enumerate(hotels):
            with cols[i % 3]:
                if st.checkbox(name, key=f"hil_sel_{fid}"):
                    selected.append((name, fid))

        wb_sels = st.pills(
            "Workbooks to update",
            PORTFOLIO_WORKBOOKS["Hilton"],
            selection_mode="multi",
            default=PORTFOLIO_WORKBOOKS["Hilton"],
            key="hil_wb",
        ) or []

        srp_file = st.file_uploader(
            "SRP Activity — all Hilton properties (one file)",
            type=["xlsx"], key="hil_srp")

        wash_files = {}
        if selected:
            st.markdown("**Group Wash report — one per property**")
            wcols = st.columns(2)
            for i, (name, fid) in enumerate(selected):
                with wcols[i % 2]:
                    wash_files[name] = st.file_uploader(
                        name, type=["xlsx"], key=f"hil_wash_{fid}")

    if not selected:
        st.info("Select at least one property.")
        return
    if not srp_file:
        st.info("Upload the SRP Activity export to continue.")
        return

    try:
        srp = parse_srp_activity(srp_file)
    except Exception as e:
        st.error(f"Could not read the SRP Activity export: {e}")
        return

    st.success(f"SRP Activity: {len(srp)} properties — {', '.join(sorted(srp))}")

    missing_wash = [n for n, _ in selected if not wash_files.get(n)]
    if missing_wash:
        st.info(f"Waiting on Group Wash report for: {', '.join(missing_wash)}")
        return

    if st.button("Preview changes", key="hil_preview", type="primary"):
        svc = get_drive_service()
        jobs, problems = [], []
        for name, fid in selected:
            try:
                wash = parse_group_wash(wash_files[name])
            except Exception as e:
                problems.append(f"{name}: could not read the Group Wash report — {e}")
                continue

            inn = _match_inncode(name, srp)
            if not inn:
                problems.append(
                    f"{name}: no matching 'Property - InnCode' in the SRP export. "
                    f"Codes present: {', '.join(sorted(srp))}. Add it to "
                    f"HILTON_INNCODES so the split knows which rows belong here.")
                continue

            for wb_type in wb_sels:
                result, err = resolve_drive_workbook(svc, fid, name, wb_type)
                if err or not result:
                    problems.append(f"{name} — {wb_type}: {err}")
                    continue
                file_id, file_name = result
                if wb_type != "ROB":
                    problems.append(
                        f"{name} — Forecast ({file_name}): ingest not wired up yet, "
                        f"skipped.")
                    continue
                raw = drive_download(svc, file_id)
                wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=False)
                avail = [s for s in ROB_SHEETS if s in wb.sheetnames]
                sheet = first_uncolored_sheet(wb, avail)
                changes = build_hilton_rob_plan(srp.get(inn, {}), wash, wb[sheet])
                passed = [f"{n} ({w})" for n, w in rob_week_status(wb, avail)
                          if w and n != sheet]
                jobs.append({
                    "key": f"{name} — {wb_type}",
                    "file_id": file_id, "file_name": file_name,
                    "wb_bytes": raw, "sheet": sheet, "changes": changes,
                    "note": f"  ·  InnCode {inn}"
                            + ("  ·  skipped " + "; ".join(passed) if passed else ""),
                })
        st.session_state["hil_jobs"] = jobs
        st.session_state["hil_problems"] = problems

    jobs = st.session_state.get("hil_jobs") or []
    for msg in st.session_state.get("hil_problems") or []:
        st.warning(msg)
    if not jobs:
        return

    for job in jobs:
        st.divider()
        st.subheader(job["key"] + f" — {job['sheet']}")
        st.caption(job["file_name"] + job.get("note", ""))
        _show_ihg_plan(job["changes"])

    st.divider()
    total = sum(len([c for c in j["changes"] if not c["skip_reason"]]) for j in jobs)
    st.write(f"**{total} cells across {len(jobs)} workbook(s).** "
             f"Applying writes to Drive and marks each tab done (green).")
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Confirm and apply to Drive", key="hil_apply", type="primary"):
            saved, errors = apply_portfolio_plans(
                get_drive_service(), jobs, "hil_undo")
            for s in saved:
                st.success(f"Saved {s}")
            for e in errors:
                st.error(e)
            if saved:
                st.session_state.pop("hil_jobs", None)
    with col_b:
        if st.session_state.get("hil_undo"):
            if st.button("Undo last apply", key="hil_undo_btn"):
                saved, errors = undo_portfolio_plans(get_drive_service(), "hil_undo")
                for s in saved:
                    st.success(f"Reverted {s}")
                for e in errors:
                    st.error(e)


def render_ihg_update(hotels):
    """IHG portfolio run.

    One hotel at a time, like Stay In Touch, but fed by the History and
    Forecast Business Block PDF rather than a CSV. That single report drives
    both workbooks: the ROB takes its Total row for the current month, and the
    Forecast takes the daily rows — completed days as actuals, on-the-books
    days as OTB rooms and rate.
    """
    if not hotels:
        st.info("No IHG properties found in Drive.")
        return

    hotel_names = [h[0] for h in hotels]
    id_map = {h[0]: h[1] for h in hotels}

    st.caption(
        "Two PDFs. History and Forecast covers the current month end to end — "
        "it fills that month of the ROB and the whole Forecast, split into "
        "actuals and on-the-books at its own Subtotal line. Business on the "
        "Books starts at the report date and runs a year out, filling every "
        "later month of the ROB."
    )

    with st.container(border=True):
        col_h, col_w = st.columns([3, 3])
        with col_h:
            hotel_sel = st.selectbox("Hotel", hotel_names, key="ihg_hotel")
        with col_w:
            wb_sels = st.pills(
                "Workbooks to update",
                WORKBOOK_TYPES,
                selection_mode="multi",
                default=WORKBOOK_TYPES,
                key="ihg_wb",
            ) or []
        c1, c2 = st.columns(2)
        with c1:
            pdf_file = st.file_uploader(
                "History and Forecast Business Block (PDF)",
                type=["pdf"], key=f"ihg_pdf_{hotel_sel}")
        with c2:
            bob_file = st.file_uploader(
                "Business on the Books (PDF)",
                type=["pdf"], key=f"ihg_bob_{hotel_sel}")
        ihg_next_month = st.checkbox(
            "Include next month's Forecast",
            key="ihg_fcst_next",
            help="Fills next month's Forecast workbook from the Business on the "
                 "Books daily rows. Needs that PDF.")

    if not pdf_file:
        st.info("Upload the History and Forecast PDF to continue.")
        return

    try:
        parsed = parse_ihg_history_forecast(pdf_file)
    except Exception as e:
        st.error(f"Could not read the History and Forecast PDF: {e}")
        return

    bob = None
    if bob_file:
        try:
            bob = parse_ihg_business_on_books(bob_file)
        except Exception as e:
            st.error(f"Could not read the Business on the Books PDF: {e}")
            return

    past = [d for d in parsed["days"] if d["is_past"]]
    future = [d for d in parsed["days"] if not d["is_past"]]
    t = parsed["total"]
    st.success(
        f"Report date {parsed['report_date']} — {len(parsed['days'])} days "
        f"({len(past)} completed, {len(future)} on the books). "
        f"Month total {t['total_occ']:,.0f} rooms / ${t['total_rev']:,.2f}."
    )

    if bob:
        # The two reports overlap from the report date to month end; if they
        # disagree there they were pulled at different moments and the ROB
        # would mix them.
        cur = bob["months"].get((parsed["report_date"].year, parsed["report_date"].month))
        sub = parsed["subtotals"][1] if len(parsed["subtotals"]) > 1 else None
        if cur and sub and abs(cur["total_rev"] - sub["total_rev"]) > 1:
            st.warning(
                f"The two reports disagree over their overlapping dates "
                f"(Business on the Books ${cur['total_rev']:,.2f} vs History and "
                f"Forecast ${sub['total_rev']:,.2f}). They were probably pulled at "
                f"different times — re-pull both together for a clean run."
            )
        st.success(
            f"Business on the Books — {len(bob['months'])} months, "
            f"{sum(1 for m in bob['months'].values() if m['rooms'])} with business on them."
        )
    else:
        st.info("No Business on the Books PDF — only the current month of the ROB will be filled.")

    if st.button("Preview changes", key="ihg_preview", type="primary"):
        svc = get_drive_service()
        hotel_id = id_map[hotel_sel]
        jobs, problems = [], []
        for wb_type in wb_sels:
            result, err = resolve_drive_workbook(svc, hotel_id, hotel_sel, wb_type)
            if err or not result:
                problems.append(f"{wb_type}: {err}")
                continue
            file_id, file_name = result
            raw = drive_download(svc, file_id)
            wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=False)
            if wb_type == "ROB":
                avail = [s for s in ROB_SHEETS if s in wb.sheetnames]
                sheet = first_uncolored_sheet(wb, avail)
                changes = build_ihg_rob_plan(parsed, wb[sheet], bob=bob)
                passed = [f"{n} ({w})" for n, w in rob_week_status(wb, avail)
                          if w and n != sheet]
                note = "  ·  skipped " + "; ".join(passed) if passed else ""
            elif wb_type == "Strategy Report":
                avail = [s for s in STRATEGY_SHEETS if s in wb.sheetnames]
                sheet = first_undone_strategy_sheet(wb, avail)
                changes = build_ihg_sr_plan(parsed, bob, wb[sheet])
                note = ""
                if not bob:
                    problems.append(
                        "Strategy Report: without the Business on the Books PDF "
                        "only the days before it starts can be filled.")
                if not changes:
                    problems.append(
                        f"{file_name}: could not match the Strategy Report's "
                        f"column headings on '{sheet}'.")
                    continue
                # The as-of date sits above the grid rather than in it.
                changes.insert(0, {
                    "row": 2, "col": 5, "label": "As-of date", "month": None,
                    "new_value": parsed["report_date"],
                    "skip_reason": "formula" if is_formula(wb[sheet].cell(2, 5).value) else None,
                })
            else:
                avail = [s for s in FORECAST_SHEETS if s in wb.sheetnames]
                sheet = first_unhighlighted_forecast_sheet(wb, avail)
                changes = build_ihg_forecast_plan(parsed, wb[sheet], wb=wb)
                note = ""
                if not changes:
                    problems.append(
                        f"{file_name}: could not locate the OTB / actual rows on "
                        f"'{sheet}' — its column-A labels don't match the template.")
                    continue
            jobs.append({"key": f"{wb_type} ({parsed['report_date']:%b %Y})",
                         "file_id": file_id, "file_name": file_name,
                         "wb_bytes": raw, "sheet": sheet,
                         "changes": changes, "note": note})

        if ihg_next_month and "Forecast" in wb_sels:
            ref = parsed["report_date"] or datetime.date.today()
            nxt = (ref.replace(day=1) + datetime.timedelta(days=32)).replace(day=1)
            if not bob:
                problems.append("Next month's Forecast needs the Business on the "
                                "Books PDF — that's where its daily rows come from.")
            else:
                nm_result, nm_err = resolve_drive_workbook(
                    svc, hotel_id, hotel_sel, "Forecast", month_date=nxt)
                if nm_err or not nm_result:
                    problems.append(f"{nxt:%b %Y} Forecast: {nm_err}")
                else:
                    nm_id, nm_name = nm_result
                    nm_raw = drive_download(svc, nm_id)
                    nm_wb = openpyxl.load_workbook(io.BytesIO(nm_raw), data_only=False)
                    nm_avail = [s for s in FORECAST_SHEETS if s in nm_wb.sheetnames]
                    nm_sheet = first_unhighlighted_forecast_sheet(nm_wb, nm_avail)
                    nm_changes = build_ihg_next_month_forecast_plan(
                        bob, nm_wb[nm_sheet], nxt, wb=nm_wb)
                    if not nm_changes:
                        problems.append(
                            f"Nothing to write for {nxt:%B %Y} — no daily rows in "
                            f"Business on the Books, or '{nm_sheet}' doesn't match "
                            f"the template.")
                    else:
                        jobs.append({"key": f"Forecast ({nxt:%b %Y})",
                                     "file_id": nm_id, "file_name": nm_name,
                                     "wb_bytes": nm_raw, "sheet": nm_sheet,
                                     "changes": nm_changes, "note": ""})

        st.session_state["ihg_jobs"] = jobs
        st.session_state["ihg_problems"] = problems

    jobs = st.session_state.get("ihg_jobs") or []
    for msg in st.session_state.get("ihg_problems") or []:
        st.error(msg)
    if not jobs:
        return

    for job in jobs:
        st.divider()
        st.subheader(f"{job['key']} — {job['sheet']}")
        st.caption(job["file_name"] + job.get("note", ""))
        _show_ihg_plan(job["changes"])

    st.divider()
    total = sum(len([c for c in j["changes"] if not c["skip_reason"]]) for j in jobs)
    st.write(f"**{total} cells across {len(jobs)} workbook(s).** "
             f"Applying writes to Drive and marks each tab done (green).")
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Confirm and apply to Drive", key="ihg_apply", type="primary"):
            saved, errors = apply_portfolio_plans(
                get_drive_service(), jobs, "ihg_undo")
            for s in saved:
                st.success(f"Saved {s}")
            for e in errors:
                st.error(e)
            if saved:
                st.session_state.pop("ihg_jobs", None)
    with col_b:
        if st.session_state.get("ihg_undo"):
            if st.button("Undo last apply", key="ihg_undo_btn"):
                saved, errors = undo_portfolio_plans(get_drive_service(), "ihg_undo")
                for s in saved:
                    st.success(f"Reverted {s}")
                for e in errors:
                    st.error(e)


def _show_ihg_plan(changes):
    writes = [c for c in changes if not c["skip_reason"]]
    skips = [c for c in changes if c["skip_reason"]]
    st.dataframe(pd.DataFrame([
        {"row": c["row"], "col": get_column_letter(c["col"]),
         "field": c["label"], "value": c["new_value"]}
        for c in writes]), use_container_width=True, hide_index=True)
    if skips:
        st.caption(f"{len(skips)} cell(s) left alone because they hold a formula.")


def _match_inncode(hotel_name, srp):
    """Map a Drive hotel name to the InnCode used in the SRP export."""
    code = HILTON_INNCODES.get(hotel_name)
    if code and code in srp:
        return code
    up = hotel_name.upper()
    for label, kws in PORTFOLIO_HOTELS["Hilton"].items():
        if any(k in up for k in kws):
            c = HILTON_INNCODES.get(label)
            if c and c in srp:
                return c
    return None


# ── Manual upload (test mode only) ───────────────────────────────────────────
if test_mode:
 with st.expander("Manual Upload", expanded=False):
    with st.expander("ROB Update"):
        st.header("ROB Master Workbook Update")
        csv_file = st.file_uploader("Upload CSV (Business on the Books)", type=["csv", "xlsx"], key="rob_csv")
        xl_file  = st.file_uploader("Upload ROB Master Workbook (.xlsx)", type=["xlsx"], key="rob_xl")
        npu_compare_file = st.file_uploader(
            "Occupancy Statistics — with unpicked group revenue included (Margaritaville only, optional)",
            type=["xlsx"], key="rob_npu_compare")

        if csv_file and xl_file:
            xl_bytes  = xl_file.read()

            df = parse_bob_source(csv_file)
            npu_compare_df = parse_bob_source(npu_compare_file) if npu_compare_file else None
            wb = openpyxl.load_workbook(io.BytesIO(xl_bytes), data_only=False)

            auto_sheet = first_uncolored_sheet(wb, ROB_SHEETS)

            # A selectbox's `index=` is only honored the first time its `key` is
            # created — once rob_sheet exists in session_state, Streamlit ignores
            # index= on every rerun and keeps the old selection. Force a reset
            # whenever the uploaded file's bytes change (see the identical fix
            # applied to the Forecast manual-upload tab).
            rob_xl_hash = hashlib.md5(xl_bytes).hexdigest()
            if st.session_state.get("rob_xl_hash") != rob_xl_hash:
                st.session_state["rob_xl_hash"] = rob_xl_hash
                st.session_state["rob_sheet"] = auto_sheet

            sheet_choice = st.selectbox("Week tab", ROB_SHEETS, key="rob_sheet")
            st.caption(f"Auto-detected next tab: **{auto_sheet}**")
            _passed = [f"{n} — {why}" for n, why in rob_week_status(wb, ROB_SHEETS)
                       if why and n != auto_sheet]
            if _passed:
                st.caption("Skipped: " + "; ".join(_passed))
    
            if st.button("Preview Changes", key="rob_preview"):
                ws = wb[sheet_choice]
                grp_npu_rev_override = compute_grp_npu_rev_override(df, npu_compare_df)
                changes = build_rob_change_plan(df, ws, grp_npu_rev_override=grp_npu_rev_override)
                st.session_state["rob_changes"]   = changes
                st.session_state["rob_wb_bytes"]  = xl_bytes
                st.session_state["rob_sheet_sel"] = sheet_choice
    
            if "rob_changes" in st.session_state:
                changes    = st.session_state["rob_changes"]
                will_write = [c for c in changes if not c["skip_reason"]]
                skipped    = [c for c in changes if c["skip_reason"]]
    
                c1, c2 = st.columns(2)
                c1.metric("Cells to update", len(will_write))
                c2.metric("Skipped",         len(skipped))
    
                preview_rows = []
                for c in changes:
                    preview_rows.append({
                        "Month":  c["month"] or "—",
                        "Label":  c["label"],
                        "Row":    c["row"],
                        "Col":    c["col"],
                        "Value":  c["new_value"],
                        "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                    })
                st.dataframe(preview_rows, use_container_width=True)
    
                if st.button("Confirm and Apply Changes", key="rob_apply"):
                    wb2 = openpyxl.load_workbook(io.BytesIO(st.session_state["rob_wb_bytes"]), data_only=False)
                    apply_rob_changes(wb2, st.session_state["rob_sheet_sel"], changes)
                    color_tab_done(wb2, st.session_state["rob_sheet_sel"])
                    strip_tables(wb2)
                    out = io.BytesIO()
                    wb2.save(out)
                    st.download_button(
                        "Download Updated ROB Workbook",
                        data=out.getvalue(),
                        file_name="ROB_Master_updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    
    with st.expander("Strategy Report"):
        st.header("Strategy Report Update")

        col_a, col_b = st.columns(2)
        with col_a:
            csv_file2 = st.file_uploader("Upload CSV (Business on the Books)", type=["csv", "xlsx"], key="str_csv")
        with col_b:
            rate_file2 = st.file_uploader("Upload Rates & Restrictions CSV", type=["csv"], key="str_rate")
    
        col_c, col_d = st.columns(2)
        with col_c:
            xl_file2 = st.file_uploader("Upload Strategy Report Workbook (.xlsx)", type=["xlsx"], key="str_xl")
    
        if xl_file2:
            xl_bytes2 = xl_file2.read()
            wb2_peek  = openpyxl.load_workbook(io.BytesIO(xl_bytes2), data_only=False)

            auto_sheet2 = first_undone_strategy_sheet(wb2_peek, STRATEGY_SHEETS)

            # Same stale-selection issue as ROB/Forecast: force a reset whenever
            # the uploaded file's bytes change, since index= is ignored once
            # str_sheet already exists in session_state.
            str_xl_hash = hashlib.md5(xl_bytes2).hexdigest()
            if st.session_state.get("str_xl_hash") != str_xl_hash:
                st.session_state["str_xl_hash"] = str_xl_hash
                st.session_state["str_sheet"] = auto_sheet2

            sheet_choice2 = st.selectbox("Week tab", STRATEGY_SHEETS, key="str_sheet")
            st.caption(f"Auto-detected next tab: **{auto_sheet2}**")
    
            if (csv_file2 or rate_file2) and st.button("Preview Changes", key="str_preview"):
                wb2_full = openpyxl.load_workbook(io.BytesIO(xl_bytes2), data_only=False)
                all_changes = []
    
                if csv_file2:
                    df2 = parse_bob_source(csv_file2)
                    all_changes += build_strategy_change_plan(df2, wb2_full, sheet_choice2)
    
                rate_warnings = []
                if rate_file2:
                    rate_df = parse_rate_csv(rate_file2.read())
                    rate_changes, rate_warnings = build_rates_change_plan(
                        rate_df, wb2_full, sheet_choice2)
                    all_changes += rate_changes
    
                st.session_state["str_changes"]   = all_changes
                st.session_state["str_wb_bytes"]  = xl_bytes2
                st.session_state["str_sheet_sel"] = sheet_choice2
                st.session_state["str_warnings"]  = rate_warnings
    
            if "str_changes" in st.session_state:
                for w in st.session_state.get("str_warnings", []):
                    st.warning(w)
    
                changes2    = st.session_state["str_changes"]
                will_write2 = [c for c in changes2 if not c["skip_reason"]]
                skipped2    = [c for c in changes2 if c["skip_reason"]]
    
                c1, c2, c3 = st.columns(3)
                c1.metric("Cells to update", len(will_write2))
                c2.metric("Skipped",         len(skipped2))
                c3.metric("Days in scope",   len({c["date"] for c in changes2}))
    
                preview_rows2 = []
                for c in changes2:
                    preview_rows2.append({
                        "Date":   str(c["date"]),
                        "Label":  c["label"],
                        "Row":    c["row"],
                        "Col":    c["col"],
                        "Value":  c["new_value"],
                        "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                    })
                st.dataframe(preview_rows2, use_container_width=True)
    
                if st.button("Confirm and Apply Changes", key="str_apply"):
                    wb3 = openpyxl.load_workbook(io.BytesIO(st.session_state["str_wb_bytes"]), data_only=False)
                    apply_strategy_changes(wb3, st.session_state["str_sheet_sel"], changes2)
                    color_tab_done(wb3, st.session_state["str_sheet_sel"])
                    strip_tables(wb3)
                    out2 = io.BytesIO()
                    wb3.save(out2)
                    st.download_button(
                        "Download Updated Strategy Workbook",
                        data=out2.getvalue(),
                        file_name="Strategy_Report_updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    
    with st.expander("Forecast"):
        st.header("Forecast Update")
    
        fcst_csv     = st.file_uploader("Upload CSV (Business on the Books)", type=["csv", "xlsx"], key="fcst_csv")
        fcst_xl      = st.file_uploader("Upload Current Month Forecast Workbook (.xlsx/.xlsm)", type=["xlsx", "xlsm"], key="fcst_xl")
        st.caption("Weeks 3 & 4 only: also upload next month's forecast workbook.")
        fcst_xl_next = st.file_uploader("Upload Next Month Forecast Workbook (.xlsx/.xlsm)", type=["xlsx", "xlsm"], key="fcst_xl_next")
    
        if fcst_xl:
            fcst_bytes = fcst_xl.read()
            wb_fcst_peek = openpyxl.load_workbook(io.BytesIO(fcst_bytes), data_only=False)
    
            avail_fcst = [s for s in FORECAST_SHEETS if s in wb_fcst_peek.sheetnames]
            if not avail_fcst:
                st.error(f"None of the expected week tabs (FCST-WK1 – FCST-WK9) were found in "
                         f"this workbook. It has: {', '.join(wb_fcst_peek.sheetnames)}. Make sure "
                         f"you uploaded the destination Forecast workbook here, not the source "
                         f"data file.")
            auto_fcst  = first_unhighlighted_forecast_sheet(wb_fcst_peek, avail_fcst) if avail_fcst else None

            # A selectbox's `index=` is only honored the first time its `key` is
            # created — once fcst_sheet exists in session_state, Streamlit ignores
            # index= on every rerun and keeps the old selection. Re-uploading a new
            # file (this week's, with more tabs now done) would silently keep
            # writing to whatever tab was last picked instead of the newly
            # auto-detected one. Force a reset whenever the uploaded bytes change.
            fcst_hash = hashlib.md5(fcst_bytes).hexdigest()
            if st.session_state.get("fcst_xl_hash") != fcst_hash:
                st.session_state["fcst_xl_hash"] = fcst_hash
                if auto_fcst:
                    st.session_state["fcst_sheet"] = auto_fcst

            fcst_sheet = st.selectbox("Week tab", avail_fcst, key="fcst_sheet")
            if auto_fcst:
                st.caption(f"Auto-detected next tab: **{auto_fcst}**")
    
            if fcst_xl_next:
                fcst_next_bytes = fcst_xl_next.read()
                wb_next_peek    = openpyxl.load_workbook(io.BytesIO(fcst_next_bytes), data_only=False)
                avail_next      = [s for s in FORECAST_SHEETS if s in wb_next_peek.sheetnames]
                auto_next       = first_unhighlighted_forecast_sheet(wb_next_peek, avail_next) if avail_next else None

                fcst_next_hash = hashlib.md5(fcst_next_bytes).hexdigest()
                if st.session_state.get("fcst_xl_next_hash") != fcst_next_hash:
                    st.session_state["fcst_xl_next_hash"] = fcst_next_hash
                    if auto_next:
                        st.session_state["fcst_sheet_next"] = auto_next

                fcst_sheet_next = st.selectbox("Next month week tab", avail_next, key="fcst_sheet_next")
                if auto_next:
                    st.caption(f"Auto-detected next month tab: **{auto_next}**")
    
            if fcst_csv and fcst_sheet and st.button("Preview Changes", key="fcst_preview"):
                # Current month workbook
                wb_fcst_full = openpyxl.load_workbook(io.BytesIO(fcst_bytes), data_only=False)
                df_fcst = parse_bob_source(fcst_csv)
                ws_fcst = wb_fcst_full[fcst_sheet]
                fcst_changes, fcst_warnings = build_forecast_change_plan(df_fcst, ws_fcst)
    
                st.session_state["fcst_changes"]   = fcst_changes
                st.session_state["fcst_wb_bytes"]  = fcst_bytes
                st.session_state["fcst_sheet_sel"] = fcst_sheet
                st.session_state["fcst_warnings"]  = fcst_warnings
    
                # Next month workbook (weeks 3 & 4)
                if fcst_xl_next:
                    wb_next_full = openpyxl.load_workbook(io.BytesIO(fcst_next_bytes), data_only=False)
                    ws_next      = wb_next_full[fcst_sheet_next]
                    next_changes, next_warnings = build_next_month_forecast_plan(df_fcst, ws_next)
                    st.session_state["fcst_next_changes"]   = next_changes
                    st.session_state["fcst_next_wb_bytes"]  = fcst_next_bytes
                    st.session_state["fcst_next_sheet_sel"] = fcst_sheet_next
                    st.session_state["fcst_next_warnings"]  = next_warnings
                else:
                    st.session_state.pop("fcst_next_changes", None)
    
            if "fcst_changes" in st.session_state:
                for w in st.session_state.get("fcst_warnings", []):
                    st.warning(w)
    
                fcst_ch   = st.session_state["fcst_changes"]
                will_fcst = [c for c in fcst_ch if not c["skip_reason"]]
                skip_fcst = [c for c in fcst_ch if c["skip_reason"]]
    
                st.subheader("Current month changes")
                c1, c2 = st.columns(2)
                c1.metric("Cells to update", len(will_fcst))
                c2.metric("Skipped",         len(skip_fcst))
    
                preview_fcst = []
                for c in fcst_ch:
                    preview_fcst.append({
                        "Label":  c["label"],
                        "Row":    c["row"],
                        "Col":    c["col"],
                        "Value":  c["new_value"],
                        "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                    })
                st.dataframe(preview_fcst, use_container_width=True)
    
                if "fcst_next_changes" in st.session_state:
                    next_ch = st.session_state["fcst_next_changes"]
                    for w in st.session_state.get("fcst_next_warnings", []):
                        st.warning(w)
                    will_next = [c for c in next_ch if not c["skip_reason"]]
                    skip_next = [c for c in next_ch if c["skip_reason"]]
                    st.subheader("Next month changes")
                    n1, n2 = st.columns(2)
                    n1.metric("Cells to update", len(will_next))
                    n2.metric("Skipped",         len(skip_next))
                    preview_next = []
                    for c in next_ch:
                        preview_next.append({
                            "Label":  c["label"],
                            "Row":    c["row"],
                            "Col":    c["col"],
                            "Value":  c["new_value"],
                            "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                        })
                    st.dataframe(preview_next, use_container_width=True)
    
                if st.button("Confirm and Apply Changes", key="fcst_apply"):
                    # Apply current month
                    wb_out = openpyxl.load_workbook(io.BytesIO(st.session_state["fcst_wb_bytes"]), data_only=False)
                    apply_forecast_changes(wb_out, st.session_state["fcst_sheet_sel"], fcst_ch)
                    color_tab_done(wb_out, st.session_state["fcst_sheet_sel"])
                    strip_tables(wb_out)
                    out3 = io.BytesIO()
                    wb_out.save(out3)
                    st.download_button(
                        "Download Updated Current Month Forecast",
                        data=out3.getvalue(),
                        file_name="Forecast_current_updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    
                    # Apply next month (if uploaded)
                    if "fcst_next_changes" in st.session_state:
                        wb_next_out = openpyxl.load_workbook(io.BytesIO(st.session_state["fcst_next_wb_bytes"]), data_only=False)
                        apply_forecast_changes(wb_next_out, st.session_state["fcst_next_sheet_sel"],
                                               st.session_state["fcst_next_changes"])
                        color_tab_done(wb_next_out, st.session_state["fcst_next_sheet_sel"])
                        strip_tables(wb_next_out)
                        out4 = io.BytesIO()
                        wb_next_out.save(out4)
                        st.download_button(
                            "Download Updated Next Month Forecast",
                            data=out4.getvalue(),
                            file_name="Forecast_next_month_updated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
    
# ── Main menu ────────────────────────────────────────────────────────────────
# Two sections rather than one flat row of tools. The tab objects are created
# inside their section but used further down at module level — a Streamlit tab
# is bound to its position on the page, not to the block that made it, so the
# tool bodies below stay exactly where they are.
main_revenue, main_financial = st.tabs(["📈  Revenue Management", "💵  Financial Tools"])

with main_revenue:
    tab_weekly, tab_ancillary, tab_ooo = st.tabs(
        ["Weekly Workbook Update", "Ancillary Revenue", "Monthly OOO Report"])

with main_financial:
    tab_pl, tab_projection = st.tabs(["P&L Spreadsheet", "1-Year Projection"])

with tab_weekly:
    st.divider()
    # ── Google Drive Update ───────────────────────────────────────────────────────
    col_title, col_month = st.columns([5, 2])
    with col_title:
        st.header("Weekly Workbook Update")
    with col_month:
        st.write("")
        st.write("")
        st.caption(f"📅 Current month: **{datetime.date.today().strftime('%B %Y')}**")

    all_discovered = get_hotels_from_drive()

    # Each portfolio's data arrives in a different export, so they get their own
    # tab rather than sharing one hotel list.
    portfolio = st.radio("Portfolio", PORTFOLIOS, horizontal=True, key="drive_portfolio")

    hotels = hotels_in_portfolio(portfolio, all_discovered)
    hotel_names = [h[0] for h in hotels]
    hotel_id_map = {h[0]: h[1] for h in hotels}
    allowed_wbs = PORTFOLIO_WORKBOOKS[portfolio]

    missing = portfolio_hotels_missing(portfolio, all_discovered)
    if missing:
        st.warning(
            f"No Drive folder found for: {', '.join(missing)} — "
            f"they won't appear below. A hotel is only discovered if its folder "
            f"contains a 'REVENUE REPORTS' subfolder shared with the service account."
        )

    if portfolio == "Hilton":
        render_hilton_update(hotels)
        st.stop()
    if portfolio == "IHG":
        render_ihg_update(hotels)
        st.stop()

    start_new_month = False
    with st.container(border=True):
        col_h, col_ref, col_w = st.columns([3, 1, 3])
        with col_h:
            hotel_sel = st.selectbox("Hotel", hotel_names if hotel_names else ["(no hotels found)"], key="drive_hotel")
        with col_ref:
            st.write("")
            if st.button("↺", key="refresh_hotels", help="Refresh hotel list"):
                get_hotels_from_drive.clear()
                st.rerun()
        with col_w:
            wb_sels = st.pills(
                "Workbooks to update",
                allowed_wbs,
                selection_mode="multi",
                default=allowed_wbs,
                key="drive_wb",
            ) or []

        # Keying these to the selected hotel clears any uploaded file the moment you
        # switch hotels — one hotel's BOB/R&R CSV should never carry over and get
        # applied to a different hotel.
        drive_csv = st.file_uploader("CSV — Business on the Books", type=["csv", "xlsx"], key=f"drive_csv_{hotel_sel}")
        drive_rate_csv = None
        if "Strategy Report" in (wb_sels or []):
            drive_rate_csv = st.file_uploader("CSV — Rates & Restrictions", type=["csv"], key=f"drive_rate_csv_{hotel_sel}")
        drive_npu_compare_csv = None
        if "ROB" in (wb_sels or []) and "margaritaville" in hotel_sel.lower():
            drive_npu_compare_csv = st.file_uploader(
                "Occupancy Statistics — with unpicked group revenue included",
                type=["xlsx"], key=f"drive_npu_compare_{hotel_sel}", width=500)

        opt_col1, opt_col2 = st.columns(2)
        forecast_next_month = False
        if "Forecast" in (wb_sels or []):
            with opt_col1:
                forecast_next_month = st.checkbox("Include next month's Forecast", key="drive_fcst_next")
        with opt_col2:
            start_new_month = st.checkbox("Set up new month", key="drive_new_month")
    if start_new_month:
        with st.container(border=True):
            today         = datetime.date.today()
            cur_month_dt  = today.replace(day=1)
            prev_month_dt = (cur_month_dt - datetime.timedelta(days=1)).replace(day=1)
            next_month_dt = (cur_month_dt + datetime.timedelta(days=32)).replace(day=1)
            month_options = {
                prev_month_dt.strftime("%B %Y"): prev_month_dt,
                cur_month_dt.strftime("%B %Y"):  cur_month_dt,
                next_month_dt.strftime("%B %Y"): next_month_dt,
            }
            month_labels = list(month_options.keys())
            # Week 4+ of the current month → default to next month; week 1 → default
            # to the current month. Weeks 2-3 keep the plain current-month default.
            # All three options remain selectable regardless of the default.
            default_dt = next_month_dt if today.day >= 22 else cur_month_dt
            sel_month_label = st.selectbox("Month to set up", month_labels,
                                            index=month_labels.index(default_dt.strftime("%B %Y")),
                                            key="setup_month_sel")
            setup_month_dt  = month_options[sel_month_label]
            month_kw        = setup_month_dt.strftime("%b%Y").upper()

            rob_col, sr_col = st.columns(2)

            # ── ROB setup ──────────────────────────────────────────────────────────
            with rob_col:
                st.markdown("**ROB**")
                if st.button("Set Up New ROB", key="btn_setup_rob", type="primary", use_container_width=True):
                    try:
                        svc         = get_drive_service()
                        hotel_id_nm = hotel_id_map.get(hotel_sel, "")
                        with st.spinner("Setting up ROB — this may take a moment..."):
                            rob_name, rob_err, rob_file_id, rob_orig_bytes = setup_new_rob_month(
                                svc, hotel_id_nm, hotel_sel, setup_month_dt)
                        if rob_err and not rob_name:
                            if "storageQuotaExceeded" in str(rob_err):
                                _, master_name = find_rob_master(svc, hotel_id_nm)
                                rob_suffix = hotel_sel.upper()
                                if master_name and "ROB" in master_name.upper():
                                    after = master_name[master_name.upper().find("ROB") + 3:].strip()
                                    after = after.replace(".xlsx","").replace(".xlsm","").replace(".XLSX","").replace(".XLSM","").strip()
                                    if after:
                                        rob_suffix = after
                                ext = ".xlsm" if master_name and master_name.lower().endswith(".xlsm") else ".xlsx"
                                suggested_name = f"{month_kw} ROB {rob_suffix}{ext}"
                                st.warning(
                                    f"Auto-copy requires a Shared Drive. Do this in Google Drive first:\n\n"
                                    f"1. Right-click **{master_name or 'the ROB master'}** → *Make a copy*\n"
                                    f"2. Rename to: **`{suggested_name}`**\n"
                                    f"3. Move into the **{month_kw}** folder\n\n"
                                    f"Then click **Set Up New ROB** again."
                                )
                            else:
                                st.error(f"ROB setup error: {rob_err}")
                        else:
                            if rob_err:
                                st.warning(rob_err)
                            if rob_file_id and rob_orig_bytes is not None:
                                st.session_state["setup_undo_rob"] = {
                                    "file_id":   rob_file_id,
                                    "file_name": rob_name,
                                    "bytes":     rob_orig_bytes,
                                }
                            st.success(f"**{rob_name}** ready for {setup_month_dt.strftime('%B %Y')}.")
                    except Exception as e:
                        st.error(f"ROB setup error: {e}")

                # Reset button — shown after a successful ROB setup
                if "setup_undo_rob" in st.session_state:
                    if st.button("↩ Reset ROB to Original", key="setup_reset_rob", type="secondary", use_container_width=True):
                        try:
                            info = st.session_state["setup_undo_rob"]
                            with st.spinner("Restoring original ROB workbook..."):
                                drive_upload(get_drive_service(), info["file_id"], info["bytes"], info["file_name"])
                            del st.session_state["setup_undo_rob"]
                            st.success(f"**{info['file_name']}** restored to original state.")
                        except Exception as e:
                            st.error(f"Reset error: {e}")

            # ── Strategy Report setup ──────────────────────────────────────────────
            with sr_col:
                st.markdown("**Strategy Report**")
                if st.button("Set Up New SR", key="btn_setup_new_wb", type="primary", use_container_width=True):
                    try:
                        svc         = get_drive_service()
                        hotel_id_nm = hotel_id_map.get(hotel_sel, "")

                        # Step 1 — ensure the file exists; skip copy if it's already there
                        is_fresh_copy = False
                        with st.spinner("Step 1 / 3 — locating or creating workbook..."):
                            existing, find_err = resolve_drive_workbook(svc, hotel_id_nm, hotel_sel,
                                                                  "Strategy Report", month_date=setup_month_dt)
                            if existing:
                                st.info(f"Found existing file: **{existing[1]}** — skipping copy.")
                            else:
                                is_fresh_copy = True
                                created_name, create_err = setup_new_sr_month(svc, hotel_id_nm, hotel_sel, setup_month_dt)
                                if create_err:
                                    master_id, master_name = find_sr_master(svc, hotel_id_nm)
                                    hotel_suffix = ""
                                    if master_name and "STRATEGY" in master_name.upper():
                                        hotel_suffix = master_name[master_name.upper().find("STRATEGY") + len("STRATEGY"):].strip().replace(".xlsx","").replace(".XLSX","").strip()
                                    suggested_name = f"{month_kw} STRATEGY {hotel_suffix}.xlsx".strip()
                                    if "storageQuotaExceeded" in str(create_err):
                                        st.warning(
                                            f"Auto-copy requires a Shared Drive. Do this in Google Drive first:\n\n"
                                            f"1. Right-click **{master_name or 'the SR master'}** → *Make a copy*\n"
                                            f"2. Rename to: **`{suggested_name}`**\n"
                                            f"3. Move into the **{month_kw}** folder\n\n"
                                            f"Then click **Set Up New SR** again."
                                        )
                                    else:
                                        st.error(f"Could not create workbook: {create_err}")
                                    st.stop()

                        # Step 2 — load reference workbooks into memory
                        with st.spinner("Step 2 / 3 — loading reference workbooks..."):
                            prev_month_dt    = (setup_month_dt - datetime.timedelta(days=1)).replace(day=1)
                            ly_month_dt      = setup_month_dt.replace(year=setup_month_dt.year - 1)
                            prev_month_sr_wb = _load_wb_from_drive(svc, hotel_id_nm, hotel_sel, "Strategy Report", prev_month_dt, data_only=False)
                            ly_sr_wb         = _load_wb_from_drive(svc, hotel_id_nm, hotel_sel, "Strategy Report", ly_month_dt)
                        st.info(f"Prev month ({prev_month_dt.strftime('%b %Y')}): {'✓' if prev_month_sr_wb else '✗ not found'}")
                        st.info(f"Last year  ({ly_month_dt.strftime('%b %Y')}): {'✓' if ly_sr_wb else '✗ not found'}")

                        # Step 3 — populate all 5 weeks
                        with st.spinner("Step 3 / 3 — populating all weeks..."):
                            result, err = resolve_drive_workbook(svc, hotel_id_nm, hotel_sel,
                                                                 "Strategy Report", month_date=setup_month_dt)
                            if err:
                                st.error(f"Cannot open new workbook: {err}")
                                st.stop()
                            file_id, file_name = result
                            wb_bytes = drive_download(svc, file_id)
                            original_bytes = wb_bytes
                            wb       = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=False)
                            if is_fresh_copy:
                                clear_tab_colors(wb, STRATEGY_SHEETS)
                            restructure_sr_dates(wb, setup_month_dt)
                            first_ws = wb[STRATEGY_SHEETS[0]] if STRATEGY_SHEETS[0] in wb.sheetnames else None
                            num_rows = _count_sheet_data_rows(first_ws) if first_ws else 365
                            full_scope_start = setup_month_dt
                            full_scope_end   = setup_month_dt + datetime.timedelta(days=max(0, num_rows - 1))
                            total_written = 0
                            for sheet_name in STRATEGY_SHEETS:
                                if sheet_name not in wb.sheetnames:
                                    continue
                                changes = build_strategy_change_plan(None, wb, sheet_name,
                                                                      prev_month_wb=prev_month_sr_wb,
                                                                      ly_wb=ly_sr_wb,
                                                                      scope_start=full_scope_start,
                                                                      scope_end=full_scope_end)
                                apply_strategy_changes(wb, sheet_name, changes)
                                total_written += len([c for c in changes if not c.get("skip_reason")])
                            strip_tables(wb)
                            out = io.BytesIO()
                            wb.save(out)
                            drive_upload(svc, file_id, out.getvalue(), file_name)
                            st.session_state["setup_undo_sr"] = {
                                "file_id":   file_id,
                                "file_name": file_name,
                                "bytes":     original_bytes,
                            }

                        st.success(
                            f"**{file_name}** is set up for {setup_month_dt.strftime('%B %Y')}. "
                            f"Populated **{total_written}** cells across all weeks."
                        )
                    except Exception as e:
                        st.error(f"Setup error: {e}")

                # Reset button — shown after a successful SR setup
                if "setup_undo_sr" in st.session_state:
                    if st.button("↩ Reset SR to Original", key="setup_reset_sr", type="secondary", use_container_width=True):
                        try:
                            info = st.session_state["setup_undo_sr"]
                            with st.spinner("Restoring original SR workbook..."):
                                drive_upload(get_drive_service(), info["file_id"], info["bytes"], info["file_name"])
                            del st.session_state["setup_undo_sr"]
                            st.success(f"**{info['file_name']}** restored to original state.")
                        except Exception as e:
                            st.error(f"Reset error: {e}")




    def build_all_plans(svc, hotel_sel, hotel_id, wb_sels, df, rate_df, forecast_next_month=False, npu_compare_df=None):
        today = datetime.date.today()
        current_month = today.replace(day=1)
        all_plans = {}

        grp_npu_rev_override = compute_grp_npu_rev_override(df, npu_compare_df)

        # Pre-load reference workbooks into memory once — used for cross-sheet lookups
        prev_month_sr_wb = None
        ly_sr_wb         = None
        if "Strategy Report" in wb_sels:
            prev_month_dt = (current_month - datetime.timedelta(days=1)).replace(day=1)
            ly_month_dt   = current_month.replace(year=current_month.year - 1)
            prev_month_sr_wb = _load_wb_from_drive(svc, hotel_id, hotel_sel, "Strategy Report", prev_month_dt, data_only=False)
            ly_sr_wb         = _load_wb_from_drive(svc, hotel_id, hotel_sel, "Strategy Report", ly_month_dt)
            # Comp Set LY / OTB LY Trans / GRP LY etc. all come from ly_sr_wb — if it's
            # not found, those fields silently produce nothing (no warning previously),
            # which looked like "dates transferred but no text" with no explanation why.
            st.info(f"SR reference workbooks — Prev month ({prev_month_dt.strftime('%b %Y')}): "
                    f"{'✓ found' if prev_month_sr_wb else '✗ NOT FOUND — OTB Lst Wek will be blank'} | "
                    f"Last year ({ly_month_dt.strftime('%b %Y')}): "
                    f"{'✓ found' if ly_sr_wb else '✗ NOT FOUND — all LY columns (incl. Comp Set LY text) will be blank'}")

        for wb_type in wb_sels:
            result, err = resolve_drive_workbook(svc, hotel_id, hotel_sel, wb_type)
            if err:
                st.error(f"{wb_type}: {err}")
                continue
            file_id, file_name = result
            wb_bytes = drive_download(svc, file_id)
            wb       = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=False)
            if wb_type == "ROB":
                avail    = [s for s in ROB_SHEETS if s in wb.sheetnames]
                auto     = first_uncolored_sheet(wb, avail)
                sheet    = auto or avail[0]
                changes  = build_rob_change_plan(df, wb[sheet], grp_npu_rev_override=grp_npu_rev_override)
                warnings = []
                # Say which weeks were passed over and why — a silently skipped
                # week tab is otherwise invisible until someone spots the gap.
                _passed = [f"{n} ({why})" for n, why in rob_week_status(wb, avail)
                           if why and n != sheet]
                if _passed:
                    warnings.append(f"Writing to '{sheet}'. Skipped: " + "; ".join(_passed))
            elif wb_type == "Strategy Report":
                avail    = [s for s in STRATEGY_SHEETS if s in wb.sheetnames]
                auto     = first_undone_strategy_sheet(wb, avail)
                sheet    = auto or avail[0]
                date_row_map_debug = build_date_row_map(wb, prefer_sheet=sheet)
                st.info(f"SR: **{file_name}** (id: `{file_id}`) → sheet **{sheet}** | "
                        f"date rows mapped: {len(date_row_map_debug)} | "
                        f"date range: {min(date_row_map_debug) if date_row_map_debug else 'none'} – {max(date_row_map_debug) if date_row_map_debug else 'none'}")
                if "WKONE" in wb.sheetnames:
                    from openpyxl.utils import get_column_letter
                    wkone_ws = wb["WKONE"]
                    wkone_col = detect_date_column(wkone_ws, wb=wb)
                    raw_r5  = wkone_ws.cell(5, wkone_col).value
                    raw_r10 = wkone_ws.cell(10, wkone_col).value
                    st.info(f"WKONE date column detected: **{get_column_letter(wkone_col)}** | "
                            f"raw value at row 5: `{raw_r5!r}` | raw value at row 10: `{raw_r10!r}`")
                    # Column C specifically — the active sheet's date formulas
                    # reference WKONE!C directly, which may not be the same
                    # column detect_date_column just picked as WKONE's own
                    # "best" column above. Need this exact value to settle
                    # whether the remaining offset originates in WKONE's own
                    # column C data, not wherever else WKONE's calendar lives.
                    raw_c5  = wkone_ws.cell(5, 3).value
                    raw_c6  = wkone_ws.cell(6, 3).value
                    st.info(f"WKONE column C specifically — raw row 5: `{raw_c5!r}` | raw row 6: `{raw_c6!r}`")
                # Strict, no-fallback view of the ACTIVE sheet's own date column —
                # this is what actually gates whether CSV data can be written at
                # all (own_date_row_map inside build_strategy_change_plan uses
                # the exact same fallback_to_wkone=False call). Always shown so
                # one run's output settles both "wrong day" and "no data" at once
                # instead of needing another round of screenshots.
                own_debug = build_date_row_map(wb, prefer_sheet=sheet, fallback_to_wkone=False)
                own_col = detect_date_column(wb[sheet], wb=wb)
                from openpyxl.utils import get_column_letter as _gcl
                st.info(f"**{sheet}**'s own date column (strict, no WKONE fallback): "
                        f"col **{_gcl(own_col)}** | rows mapped: {len(own_debug)} | "
                        f"range: {min(own_debug) if own_debug else 'NONE — this is why no CSV data would write'} "
                        f"– {max(own_debug) if own_debug else ''} | "
                        f"raw row5: `{wb[sheet].cell(5, own_col).value!r}` | "
                        f"raw row10: `{wb[sheet].cell(10, own_col).value!r}`")
                if df is not None:
                    sample_dates = [str(df.iloc[i, 0]) for i in range(min(5, len(df)))]
                    bob_daily = sum(1 for _, r in df.iterrows() if classify_row(str(r[0]).strip())[0] == "daily")
                    st.info(f"BOB CSV: {len(df)} rows | daily rows matched: {bob_daily} | first 5 col-0 values: {sample_dates}")
                else:
                    st.warning("BOB CSV: df is None — no CSV uploaded or parse failed")
                # Only extract LY data during month setup, not on regular CSV uploads
                # (ly_sr_wb is already cleared of blanking logic if ly_data is empty)
                changes  = build_strategy_change_plan(df, wb, sheet,
                                                       prev_month_wb=prev_month_sr_wb,
                                                       ly_wb=None)
                warnings = []
                if rate_df is not None:
                    rate_changes, rate_warnings = build_rates_change_plan(rate_df, wb, sheet)
                    changes  += rate_changes
                    warnings += rate_warnings
            else:  # Forecast — current month (no Month Ending Forecast fill here)
                avail    = [s for s in FORECAST_SHEETS if s in wb.sheetnames]
                auto     = first_unhighlighted_forecast_sheet(wb, avail)
                sheet    = auto or avail[0]
                changes, warnings = build_forecast_change_plan(df, wb[sheet])
            all_plans[wb_type] = {
                "file_id":   file_id,
                "file_name": file_name,
                "wb_bytes":  wb_bytes,
                "sheet":     sheet,
                "changes":   changes,
                "warnings":  warnings,
            }

        # Next-month Forecast: only when checkbox is ticked
        if "Forecast" in wb_sels and forecast_next_month:
            next_month_dt = (datetime.date.today().replace(day=1) + datetime.timedelta(days=32)).replace(day=1)
            nm_result, nm_err = resolve_drive_workbook(svc, hotel_id, hotel_sel, "Forecast", month_date=next_month_dt)
            if nm_err:
                # Workbook not found — auto-create from master
                st.info(f"Next month Forecast not found — creating from master...")
                created_name, setup_err = setup_new_forecast_month(svc, hotel_id, hotel_sel, next_month_dt)
                if setup_err and not created_name:
                    st.warning(f"Next month Forecast: {setup_err}")
                    nm_result = None
                else:
                    if setup_err:
                        st.warning(setup_err)
                    else:
                        st.success(f"Created **{created_name}** for {next_month_dt.strftime('%B %Y')}.")
                    nm_result, nm_err = resolve_drive_workbook(svc, hotel_id, hotel_sel, "Forecast", month_date=next_month_dt)
                    if nm_err:
                        st.warning(f"Still could not find next month Forecast after creation: {nm_err}")
                        nm_result = None

            # Populate week 3/4's data into next month's WK1 regardless of
            # whether the file already existed or was just created above —
            # this used to live in the sibling `else` of `if nm_err`, so on
            # the very run that created the file, population never ran at
            # all (the new copy was created with dates only, no OTB data).
            if nm_result:
                nm_file_id, nm_file_name = nm_result
                nm_bytes = drive_download(svc, nm_file_id)
                nm_wb    = openpyxl.load_workbook(io.BytesIO(nm_bytes), data_only=False)
                nm_avail = [s for s in FORECAST_SHEETS if s in nm_wb.sheetnames]
                nm_auto  = first_unhighlighted_forecast_sheet(nm_wb, nm_avail)
                nm_sheet = nm_auto or nm_avail[0]
                nm_changes, nm_warnings = build_next_month_forecast_plan(df, nm_wb[nm_sheet])
                # Month Ending Forecast table — fill Budget + LY from next month's ROB
                nm_is_wk1 = (nm_sheet == FORECAST_SHEETS[0])
                if nm_is_wk1:
                    nm_rob_result, _ = resolve_drive_workbook(svc, hotel_id, hotel_sel, "ROB",
                                                               month_date=next_month_dt)
                    if nm_rob_result:
                        nm_rob_wb = openpyxl.load_workbook(
                            io.BytesIO(drive_download(svc, nm_rob_result[0])), data_only=True)
                        extra, extra_warn = build_forecast_change_plan(
                            df, nm_wb[nm_sheet], rob_wb=nm_rob_wb, is_wk1=True)
                        # Only keep the Month Ending Forecast entries from extra
                        nm_changes += [c for c in extra if "Month End Forecast" in c.get("label", "")]
                        nm_warnings += extra_warn
                all_plans["Forecast (next month)"] = {
                    "file_id":   nm_file_id,
                    "file_name": nm_file_name,
                    "wb_bytes":  nm_bytes,
                    "sheet":     nm_sheet,
                    "changes":   nm_changes,
                    "warnings":  nm_warnings,
                }

        return all_plans


    def _snapshot_changes(wb, sheet_name, changes):
        """Return {(sheet, row, col): original_value} for every cell in the change plan."""
        ws = wb[sheet_name]
        return {
            (sheet_name, ch["row"], ch["col"]): ws.cell(ch["row"], ch["col"]).value
            for ch in changes
            if not ch.get("skip_reason")
        }


    def apply_and_upload(svc, all_plans):
        saved, errors = [], []
        undo_snapshot = {}  # cumulative snapshot across all workbooks
        for wb_type, plan in all_plans.items():
            try:
                wb_apply = openpyxl.load_workbook(io.BytesIO(plan["wb_bytes"]), data_only=False)
                # Snapshot originals BEFORE writing
                snap = _snapshot_changes(wb_apply, plan["sheet"], plan["changes"])
                undo_snapshot[wb_type] = {
                    "file_id":   plan["file_id"],
                    "file_name": plan["file_name"],
                    "wb_bytes":  plan["wb_bytes"],   # clean pre-write bytes
                    "sheet":     plan["sheet"],
                    "cells":     snap,
                }
                if wb_type == "ROB":
                    apply_rob_changes(wb_apply, plan["sheet"], plan["changes"])
                elif wb_type == "Strategy Report":
                    apply_strategy_changes(wb_apply, plan["sheet"], plan["changes"])
                else:
                    apply_forecast_changes(wb_apply, plan["sheet"], plan["changes"])
                color_tab_done(wb_apply, plan["sheet"])
                strip_tables(wb_apply)
                out = io.BytesIO()
                wb_apply.save(out)
                drive_upload(svc, plan["file_id"], out.getvalue(), plan["file_name"])
                saved.append(plan["file_name"])
            except Exception as e:
                errors.append(f"{wb_type}: {e}")
        if saved:
            st.session_state["undo_snapshot"] = undo_snapshot
        return saved, errors


    def undo_all_changes(svc):
        """Restore every snapshotted cell to its original value and re-upload."""
        snapshot = st.session_state.get("undo_snapshot", {})
        if not snapshot:
            return [], ["No snapshot found — nothing to undo."]
        saved, errors = [], []
        for wb_type, info in snapshot.items():
            try:
                wb = openpyxl.load_workbook(io.BytesIO(info["wb_bytes"]), data_only=False)
                ws = wb[info["sheet"]]
                for (sheet, row, col), orig_val in info["cells"].items():
                    ws.cell(row, col).value = orig_val
                strip_tables(wb)
                out = io.BytesIO()
                wb.save(out)
                drive_upload(svc, info["file_id"], out.getvalue(), info["file_name"])
                saved.append(info["file_name"])
            except Exception as e:
                errors.append(f"{wb_type}: {e}")
        if saved:
            del st.session_state["undo_snapshot"]
        return saved, errors


    ready = drive_csv and wb_sels

    if test_mode:
        # ── Test mode: preview first, then confirm ────────────────────────────────
        if ready and st.button("Preview Changes", key="drive_preview"):
            try:
                svc = get_drive_service()
                df      = parse_bob_source(drive_csv) if drive_csv else None
                rate_df = parse_rate_csv(drive_rate_csv.read()) if drive_rate_csv else None
                npu_compare_df = parse_bob_source(drive_npu_compare_csv) if drive_npu_compare_csv else None
                st.session_state["drive_plans"]     = build_all_plans(svc, hotel_sel, hotel_id_map.get(hotel_sel, ""), wb_sels, df, rate_df, forecast_next_month, npu_compare_df)
                st.session_state["drive_hotel_sel"] = hotel_sel
            except Exception as e:
                st.error(f"Drive error: {e}")

        if "drive_plans" in st.session_state:
            all_plans = st.session_state["drive_plans"]
            for wb_type, plan in all_plans.items():
                st.subheader(wb_type)
                st.caption(f"File: **{plan['file_name']}** — Tab: **{plan['sheet']}**")
                for w in plan["warnings"]:
                    st.warning(w)
                ch = plan["changes"]
                will_write = [c for c in ch if not c["skip_reason"]]
                skipped    = [c for c in ch if c["skip_reason"]]
                c1, c2 = st.columns(2)
                c1.metric("Cells to update", len(will_write))
                c2.metric("Skipped",         len(skipped))
                st.dataframe([{
                    "Label":  c["label"],
                    "Row":    c["row"],
                    "Col":    c["col"],
                    "Value":  c["new_value"],
                    "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                } for c in ch], use_container_width=True)

            if st.button("Confirm and Save All to Google Drive", key="drive_apply"):
                try:
                    saved, errors = apply_and_upload(get_drive_service(), all_plans)
                    for name in saved:
                        st.success(f"Saved **{name}** to Google Drive.")
                    for err in errors:
                        st.error(err)
                except Exception as e:
                    st.error(f"Drive error: {e}")
    else:
        # ── Normal mode: one click ────────────────────────────────────────────────
        if ready and st.button("Upload Data to Workbooks", key="drive_go", type="primary"):
            try:
                svc = get_drive_service()
                df      = parse_bob_source(drive_csv) if drive_csv else None
                rate_df = parse_rate_csv(drive_rate_csv.read()) if drive_rate_csv else None
                npu_compare_df = parse_bob_source(drive_npu_compare_csv) if drive_npu_compare_csv else None
                with st.spinner("Updating workbooks in Google Drive..."):
                    all_plans       = build_all_plans(svc, hotel_sel, hotel_id_map.get(hotel_sel, ""), wb_sels, df, rate_df, forecast_next_month, npu_compare_df)
                    saved, errors   = apply_and_upload(svc, all_plans)
                for name in saved:
                    st.success(f"Saved **{name}** to Google Drive.")
                for err in errors:
                    st.error(err)
            except Exception as e:
                st.error(f"Drive error: {e}")

    # ── Undo button (shown whenever a snapshot exists) ────────────────────────────
    if "undo_snapshot" in st.session_state:
        st.divider()
        undo_col, _ = st.columns([2, 5])
        with undo_col:
            if st.button("↩ Undo Last Upload", key="undo_all", type="secondary", use_container_width=True):
                try:
                    with st.spinner("Restoring original values..."):
                        saved, errors = undo_all_changes(get_drive_service())
                    for name in saved:
                        st.success(f"Restored **{name}** to original state.")
                    for err in errors:
                        st.error(err)
                except Exception as e:
                    st.error(f"Undo error: {e}")


with tab_ancillary:
    st.caption("Fills in the next available week (auto-detected from the sheet's own "
               "yellow highlighting) in the Upsell Overview table.")

    ar_month = st.selectbox(
        "Month tab", ["FEB", "MAR", "APR", "MAY", "JUN", "JUL"],
        index=["FEB", "MAR", "APR", "MAY", "JUN", "JUL"].index(
            datetime.date.today().strftime("%b").upper()
        ) if datetime.date.today().strftime("%b").upper() in
             ["FEB", "MAR", "APR", "MAY", "JUN", "JUL"] else 5,
        key="ar_month",
    )
    ar_addon_file = st.file_uploader(
        "Add-on Production report (CSV)", type=["csv"], key="ar_addon_csv")
    ar_upsell_file = st.file_uploader(
        "Upsell Report (format not wired up yet — feeds the room-upgrade rows, "
        "coming soon)", type=["csv", "xlsx"], key="ar_upsell_csv", disabled=True)

    if ar_addon_file and st.button("Preview Ancillary Revenue Changes", key="ar_preview"):
        try:
            svc = get_drive_service()
            ar_file_id, ar_file_name = find_ancillary_revenue_file(svc)
            if not ar_file_id:
                st.error(ar_file_name)  # error string in the name slot
            else:
                ar_bytes = drive_download(svc, ar_file_id)
                ar_wb = openpyxl.load_workbook(io.BytesIO(ar_bytes), data_only=False)
                if ar_month not in ar_wb.sheetnames:
                    st.error(f"No '{ar_month}' tab found in {ar_file_name}.")
                else:
                    ar_ws = ar_wb[ar_month]
                    ar_week = find_next_available_week(ar_ws)
                    if ar_week is None:
                        st.warning("No yellow-highlighted week found — every week may already be filled in.")
                    else:
                        daily_by_name = parse_addon_production_csv(ar_addon_file.read())
                        ar_changes = build_ancillary_addon_change_plan(daily_by_name, ar_ws, ar_week)
                        st.session_state["ar_file_id"]    = ar_file_id
                        st.session_state["ar_file_name"]  = ar_file_name
                        st.session_state["ar_bytes"]      = ar_bytes
                        st.session_state["ar_month_sel"]  = ar_month
                        st.session_state["ar_week_sel"]   = ar_week
                        st.session_state["ar_changes"]    = ar_changes
                        st.success(f"Detected **Week {ar_week}** as next available in **{ar_month}**.")
        except Exception as e:
            st.error(f"Preview error: {e}")

    if "ar_changes" in st.session_state:
        ar_changes = st.session_state["ar_changes"]
        st.write(f"**{len(ar_changes)}** cells will be written to "
                 f"**{st.session_state['ar_file_name']}** → {st.session_state['ar_month_sel']} "
                 f"→ Week {st.session_state['ar_week_sel']}:")
        st.dataframe(
            [{"Cell": f"{openpyxl.utils.get_column_letter(c['col'])}{c['row']}",
              "Field": c["label"], "New value": c["new_value"]} for c in ar_changes],
            use_container_width=True,
        )
        if st.button("Apply to Google Drive", key="ar_apply", type="primary"):
            try:
                svc = get_drive_service()
                ar_wb = openpyxl.load_workbook(io.BytesIO(st.session_state["ar_bytes"]), data_only=False)
                ar_ws = ar_wb[st.session_state["ar_month_sel"]]
                apply_ancillary_changes(ar_ws, st.session_state["ar_changes"])
                out = io.BytesIO()
                ar_wb.save(out)
                drive_upload(svc, st.session_state["ar_file_id"], out.getvalue(), st.session_state["ar_file_name"])
                st.success(f"**{st.session_state['ar_file_name']}** updated.")
                for key in ["ar_changes", "ar_bytes", "ar_file_id", "ar_file_name", "ar_month_sel", "ar_week_sel"]:
                    del st.session_state[key]
            except Exception as e:
                st.error(f"Apply error: {e}")

with tab_ooo:
    st.caption(
        "Pulls the Sell-Out Efficiency Report straight from Drive, adds a bright-green "
        "summary tab — right after the Report tab (2nd tab on the sheet), named e.g. "
        "'JUL 2026' — totaling each property's End. OOO Rooms across every daily tab in "
        "the selected month, with an ADR column pulled from that property's own ROB "
        "(most recently-finalized week, current month falling back to previous month), "
        "and writes it back to Drive. The file is never re-saved through openpyxl, so the "
        "macro buttons on every existing tab are left completely untouched."
    )

    today = datetime.date.today()
    ooo_cur_month_dt  = today.replace(day=1)
    ooo_prev_month_dt = (ooo_cur_month_dt - datetime.timedelta(days=1)).replace(day=1)
    ooo_next_month_dt = (ooo_cur_month_dt + datetime.timedelta(days=32)).replace(day=1)
    ooo_month_options = {
        ooo_prev_month_dt.strftime("%B %Y"): ooo_prev_month_dt,
        ooo_cur_month_dt.strftime("%B %Y"):  ooo_cur_month_dt,
        ooo_next_month_dt.strftime("%B %Y"): ooo_next_month_dt,
    }
    ooo_month_labels = list(ooo_month_options.keys())
    # Default to the previous (just-completed) month — e.g. on Aug 1st this
    # defaults to July — since current/next are only there for override.
    ooo_sel_label = st.selectbox(
        "Month to summarize", ooo_month_labels,
        index=ooo_month_labels.index(ooo_prev_month_dt.strftime("%B %Y")),
        key="ooo_month_sel",
    )
    ooo_target_dt = ooo_month_options[ooo_sel_label]

    if st.button("Preview Monthly OOO Report", key="ooo_preview", type="primary"):
        try:
            svc = get_drive_service()
            ooo_file_id, ooo_file_name = find_ooo_report_file(svc)
            if not ooo_file_id:
                st.error(ooo_file_name)
            else:
                ooo_bytes = drive_download(svc, ooo_file_id)
                ooo_wb = openpyxl.load_workbook(io.BytesIO(ooo_bytes), data_only=True, read_only=True)
                ooo_months = list_ooo_available_months(ooo_wb)
                match = next((m for m in ooo_months
                              if m[0] == ooo_target_dt.year and m[1] == ooo_target_dt.month), None)
                if not match:
                    st.error(f"No dated tabs found for {ooo_target_dt.strftime('%B %Y')} in {ooo_file_name}.")
                else:
                    ooo_year, ooo_month, ooo_sheet_names = match
                    order, totals, days_included, skipped = build_ooo_monthly_totals(
                        ooo_wb, ooo_year, ooo_month, ooo_sheet_names)
                    if not order:
                        st.error(
                            "None of that month's tabs matched the expected layout "
                            "('Property' header in row 7 col C, 'End...OOO' in row 7 col J) — "
                            "nothing was totaled, so no report was built."
                        )
                    else:
                        with st.spinner("Looking up ADR for each property from its ROB..."):
                            hotels_for_adr = get_hotels_from_drive()
                            if not hotels_for_adr:
                                # A silent [] can also be a cached transient Drive
                                # failure (5-min TTL) — clear and retry once before
                                # concluding nothing is shared.
                                get_hotels_from_drive.clear()
                                hotels_for_adr = get_hotels_from_drive()
                            adr_lookup = build_ooo_adr_lookup(svc, order, hotels_for_adr, ooo_year, ooo_month)
                        st.session_state["ooo_hotels_seen"] = [h[0] for h in hotels_for_adr]
                        st.session_state["ooo_file_id"]       = ooo_file_id
                        st.session_state["ooo_file_name"]     = ooo_file_name
                        st.session_state["ooo_bytes"]         = ooo_bytes
                        st.session_state["ooo_year"]          = ooo_year
                        st.session_state["ooo_month"]         = ooo_month
                        st.session_state["ooo_order"]         = order
                        st.session_state["ooo_totals"]        = totals
                        st.session_state["ooo_days_included"] = days_included
                        st.session_state["ooo_skipped"]       = skipped
                        st.session_state["ooo_adr_lookup"]    = adr_lookup
                        matched = sum(1 for adr, _ in adr_lookup.values() if adr is not None)
                        st.success(f"Ready to summarize **{ooo_target_dt.strftime('%B %Y')}** "
                                   f"from **{ooo_file_name}** ({days_included} daily reports found, "
                                   f"ADR matched for {matched}/{len(order)} properties).")
        except Exception as e:
            st.error(f"Preview error: {e}")

    if "ooo_order" in st.session_state:
        order      = st.session_state["ooo_order"]
        totals     = st.session_state["ooo_totals"]
        adr_lookup = st.session_state["ooo_adr_lookup"]
        hotels_seen = st.session_state.get("ooo_hotels_seen", [])
        with st.expander(f"Hotel folders this app can see in Drive ({len(hotels_seen)})"):
            st.write(", ".join(hotels_seen) if hotels_seen else
                     "None — no hotel folders are shared with this environment's service account.")
        if st.session_state.get("ooo_skipped"):
            st.caption(f"Skipped tabs that didn't match the expected layout: "
                       f"{', '.join(st.session_state['ooo_skipped'])}")
        st.dataframe(
            [{"Property": p, "Total End. OOO Rooms": totals[p],
              "ADR": adr_lookup[p][0],
              "Revenue": round(adr_lookup[p][0] * totals[p], 2) if adr_lookup[p][0] is not None else None,
              "ADR note": adr_lookup[p][1] or ""} for p in order],
            use_container_width=True,
        )
        if st.button("Apply to Google Drive", key="ooo_apply", type="primary"):
            try:
                svc = get_drive_service()
                new_bytes, tab_name, err = inject_ooo_monthly_sheet(
                    st.session_state["ooo_bytes"], st.session_state["ooo_year"],
                    st.session_state["ooo_month"], order, totals,
                    st.session_state["ooo_days_included"], adr_lookup)
                if err:
                    st.error(err)
                else:
                    drive_upload(svc, st.session_state["ooo_file_id"], new_bytes,
                                 st.session_state["ooo_file_name"])
                    st.success(f"Added **{tab_name}** to **{st.session_state['ooo_file_name']}**.")
                    for key in ["ooo_file_id", "ooo_file_name", "ooo_bytes", "ooo_year", "ooo_month",
                                "ooo_order", "ooo_totals", "ooo_days_included", "ooo_skipped",
                                "ooo_adr_lookup", "ooo_hotels_seen"]:
                        st.session_state.pop(key, None)
            except Exception as e:
                st.error(f"Apply error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# P&L Spreadsheet
# ══════════════════════════════════════════════════════════════════════════════
with tab_pl:
    st.divider()
    st.header("P&L Spreadsheet")
    st.caption(
        "Upload one Operating Statement (.xls) per hotel per year. The hotel name "
        "and year are read out of each report header, so filenames do not matter. "
        "Ten files for one hotel gives you ten years."
    )

    import hotel_pl_tool as PL
    st.caption("Builder version: **{}**".format(getattr(PL, "__version__", "unknown")))

    pl_files = st.file_uploader(
        "Operating statements (.xls)", type=["xls"],
        accept_multiple_files=True, key="pl_uploads",
    )

    if not pl_files:
        st.info("Drop your operating statements above to get started.")
    else:
        pl_parsed, pl_problems, pl_rows, pl_asof = {}, [], [], {}
        for f in pl_files:
            try:
                h, y, lines, asof = PL.parse_statement(data=f.getvalue(), name=f.name)
                used = "yes"
                prev = pl_asof.get((h, y))
                if prev is not None and prev >= asof:
                    used = "no - older as-of than another file for this year"
                else:
                    if prev is not None:
                        used = "yes - replaced an earlier as-of for this year"
                    pl_parsed.setdefault(h, {})[y] = lines
                    pl_asof[(h, y)] = asof
                pl_rows.append({"File": f.name, "Hotel": h, "Year": y,
                                "As of": "{}/{}".format(asof[0], asof[1]),
                                "Lines read": len(lines), "Used": used})
            except Exception as exc:
                pl_problems.append("{}: {}".format(f.name, exc))
                pl_rows.append({"File": f.name, "Hotel": "COULD NOT READ",
                                "Year": None, "As of": "", "Lines read": 0,
                                "Used": "no"})

        # always show what every file resolved to - this is where surprises show up
        with st.expander("Files read  ({} uploaded, {} hotel(s) found)".format(
                len(pl_files), len(pl_parsed)), expanded=len(pl_parsed) != 1):
            st.dataframe(pd.DataFrame(pl_rows), use_container_width=True, hide_index=True)
            if len(pl_parsed) > 1:
                st.warning(
                    "More than one hotel name was found, so the years are split "
                    "between them. If these are all the same property, the report "
                    "headers differ - check the Hotel column above."
                )

        for msg in pl_problems:
            st.warning("Skipped - " + msg)

        if not pl_parsed:
            st.error("No statements could be read.")
        else:
            hotel_pick = st.selectbox("Hotel", sorted(pl_parsed), key="pl_hotel")
            per_year = pl_parsed[hotel_pick]
            years = sorted(per_year)
            years = years[-PL.MAX_YEARS:]
            per_year = {y: per_year[y] for y in years}
            st.success("**{}** - {} year(s): {}".format(
                hotel_pick, len(years), ", ".join(str(y) for y in years)))
            if len(years) == 1:
                st.info(
                    "Each operating statement covers **one year**, so one file gives "
                    "one column. To get 10 years, upload 10 statements for this hotel "
                    "- one per year end. Check the 'Files read' table above to confirm "
                    "every file was picked up and resolved to the year you expect."
                )

            WANT = ["Total Revenue", "Room", "Food & Beverage", "Miscellaneous",
                    "Rental Income", "Operating Profit or Loss",
                    "Net Income or Loss", "A.D.R.", "Occupancy", "REV PAR"]
            recs = []
            for yr in years:
                row = {"Year": yr}
                for ln in per_year[yr]:
                    if ln.page == "Summary" and ln.label in WANT and ln.label not in row:
                        row[ln.label] = ln.act
                        row[ln.label + " (Budget)"] = ln.bud
                recs.append(row)
            pl_df = pd.DataFrame(recs).set_index("Year")

            latest = years[-1]
            mcols = st.columns(4)

            def _pl_metric(col, label, field, money=True):
                if field not in pl_df.columns:
                    return
                act = pl_df.loc[latest, field]
                bkey = field + " (Budget)"
                bud = pl_df.loc[latest, bkey] if bkey in pl_df.columns else None
                delta = None
                if bud not in (None, 0):
                    delta = "{:+.1%} vs budget".format((act - bud) / bud)
                shown = "${:,.0f}".format(act) if money else "{:,.2f}".format(act)
                col.metric(label, shown, delta)

            _pl_metric(mcols[0], "Total Revenue {}".format(latest), "Total Revenue")
            _pl_metric(mcols[1], "Operating Profit {}".format(latest),
                       "Operating Profit or Loss")
            _pl_metric(mcols[2], "Net Income {}".format(latest), "Net Income or Loss")
            _pl_metric(mcols[3], "ADR {}".format(latest), "A.D.R.", money=False)

            with st.expander("Show the underlying numbers"):
                st.dataframe(pl_df.style.format("{:,.2f}"), use_container_width=True)

            st.divider()
            if st.button("Build P&L Workbook", key="pl_build", type="primary"):
                try:
                    buf = io.BytesIO()
                    PL.build_workbook(hotel_pick, per_year, buf, {})
                    safe_name = re.sub(r"[^A-Za-z0-9 -]", "", hotel_pick).strip() or "Hotel"
                    st.download_button(
                        "Download P&L Workbook",
                        data=buf.getvalue(),
                        file_name=safe_name + " - P&L.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="pl_dl",
                    )
                    st.success("Built. Tabs: Summary, Rooms, Food, Beverage, "
                               "Miscellaneous, Fixed Expenses.")
                except Exception as exc:
                    st.error("Build error: {}".format(exc))


# ══════════════════════════════════════════════════════════════════════════════
# 1-Year Projection
# ══════════════════════════════════════════════════════════════════════════════
with tab_projection:
    st.divider()
    try:
        from projector import ui as projector_ui
        projector_ui.render()
    except Exception as exc:
        st.error(f"1-Year Projection failed to load: {exc}")
        st.caption(
            "It needs altair, numpy and XlsxWriter — check they installed with "
            "the rest of requirements.txt."
        )
