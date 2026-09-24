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
