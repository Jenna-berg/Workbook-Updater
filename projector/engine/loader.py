"""Parse an uploaded day-by-day segment workbook into a clean daily history frame.

Column titles are re-detected from the actual sheet every time -- never assume a
fixed position or exact spelling. The source files are pivot-table exports and
the wording drifts between hotels ("Corp - Preferred Revenue" next to
"Corp - Pref Rooms" in the same file).
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

# Canonical segment order used everywhere downstream.
SEGMENTS = ["Transient", "Group", "Government", "Corp - Preferred", "Ext Stay"]

# Every segment is projected, Group included. Group is far the lumpiest -- 48% of
# nights sell none at all -- but that is an argument for smoothing it, not for
# dropping it: the day-level detail is noise while the seasonal shape is solid,
# and heavy smoothing keeps the shape without pretending to know which night a
# block lands on.
DEFAULT_FORECAST_SEGMENTS = ["Transient", "Group", "Government", "Corp - Preferred", "Ext Stay"]

# Ordered most-specific first so "Corp - Preferred" wins before a bare "pref".
_SEGMENT_PATTERNS: list[tuple[str, list[str]]] = [
    ("Ext Stay", [r"ext\w*\s*stay", r"extended"]),
    ("Corp - Preferred", [r"corp", r"preferred", r"\bpref\b"]),
    ("Government", [r"\bgov"]),
    ("Group", [r"\bgroup"]),
    ("Transient", [r"\btrans"]),
]

_METRIC_PATTERNS: list[tuple[str, list[str]]] = [
    ("adr", [r"\badr\b", r"average\s*daily\s*rate"]),
    ("rooms", [r"\broom", r"\brns\b", r"\brmns\b"]),
    ("revenue", [r"\brev\b", r"revenue"]),
]


class LoaderError(ValueError):
    """Raised when the uploaded workbook cannot be interpreted."""


@dataclass
class LoadResult:
    daily: pd.DataFrame
    sheet_name: str
    header_row: int  # 0-indexed row of the header inside the sheet
    column_map: dict[tuple[str, str], str]  # (segment, metric) -> original header
    comp_column: str | None
    notes: list[str] = field(default_factory=list)

    @property
    def start(self) -> pd.Timestamp:
        return self.daily["Date"].min()

    @property
    def end(self) -> pd.Timestamp:
        return self.daily["Date"].max()


def _norm(text: object) -> str:
    """Lower-case, strip punctuation, collapse whitespace."""
    s = str(text if text is not None else "").lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _classify(header: object) -> tuple[str, str] | None:
    """Map one raw header to (segment, metric), or None if it is not a data column."""
    norm = _norm(header)
    if not norm:
        return None

    metric = next(
        (name for name, pats in _METRIC_PATTERNS if any(re.search(p, norm) for p in pats)),
        None,
    )
    if metric is None:
        return None

    segment = next(
        (name for name, pats in _SEGMENT_PATTERNS if any(re.search(p, norm) for p in pats)),
        None,
    )
    if segment is None:
        # "Comp Rooms" has a metric but no segment -- handled separately.
        return ("__comp__", metric) if "comp" in norm else None

    return segment, metric


def _looks_like_date_header(header: object) -> bool:
    norm = _norm(header)
    return bool(norm) and bool(re.search(r"\bdate\b|\bday\b|business day|stay date", norm))


def _score_header_row(row: pd.Series) -> tuple[int, dict[tuple[str, str], object], object]:
    """Return (match count, {(segment, metric): column label}, date column label)."""
    mapping: dict[tuple[str, str], object] = {}
    date_col = None
    for col, value in row.items():
        if date_col is None and _looks_like_date_header(value):
            date_col = col
            continue
        key = _classify(value)
        if key is not None and key not in mapping:
            mapping[key] = col
    return len(mapping), mapping, date_col


def _find_header(raw: pd.DataFrame, max_scan: int = 25) -> tuple[int, dict, object]:
    best = (-1, {}, None, -1)  # score, mapping, date col, row index
    for idx in range(min(max_scan, len(raw))):
        score, mapping, date_col = _score_header_row(raw.iloc[idx])
        if date_col is not None:
            score += 2
        if score > best[0]:
            best = (score, mapping, date_col, idx)
    return best[3], best[1], best[2]


def load_workbook(source, sheet_name: str | None = None) -> LoadResult:
    """Read `source` (path or file-like) and return a tidy daily history frame.

    The returned `daily` frame has one row per calendar date and columns
    ``Date`` plus ``<Segment> Rooms`` / ``<Segment> ADR`` / ``<Segment> Revenue``
    for every segment present, plus ``Comp Rooms``.
    """
    sheets = pd.read_excel(source, sheet_name=sheet_name, header=None)
    if isinstance(sheets, pd.DataFrame):
        sheets = {sheet_name or 0: sheets}

    best_name, best_raw, best_score = None, None, -1
    best_header, best_map, best_date = 0, {}, None
    for name, raw in sheets.items():
        if raw.empty:
            continue
        header_row, mapping, date_col = _find_header(raw)
        score = len(mapping) + (2 if date_col is not None else 0)
        if score > best_score:
            best_name, best_raw, best_score = name, raw, score
            best_header, best_map, best_date = header_row, mapping, date_col

    if best_raw is None or best_date is None or not best_map:
        raise LoaderError(
            "Could not find a day-by-day table. The sheet needs a date column and "
            "columns such as 'Transient Rooms' / 'Transient ADR'."
        )

    body = best_raw.iloc[best_header + 1 :].copy()
    notes: list[str] = []

    # Drop pivot subtotal rows and anything that is not a real date.
    dates = pd.to_datetime(body[best_date], errors="coerce")
    dropped = int(dates.isna().sum())
    if dropped:
        labels = (
            body.loc[dates.isna(), best_date].astype(str).str.strip().replace("nan", "")
        )
        labels = sorted({v for v in labels if v})
        notes.append(
            f"Ignored {dropped} non-date row(s)" + (f": {', '.join(labels[:3])}" if labels else "")
        )
    body = body.loc[dates.notna()].copy()
    body["Date"] = dates.loc[dates.notna()].dt.normalize()

    out = pd.DataFrame({"Date": body["Date"].to_numpy()})
    column_map: dict[tuple[str, str], str] = {}

    def numeric(col) -> pd.Series:
        return pd.to_numeric(body[col], errors="coerce").fillna(0.0).to_numpy()

    for segment in SEGMENTS:
        rooms_col = best_map.get((segment, "rooms"))
        adr_col = best_map.get((segment, "adr"))
        rev_col = best_map.get((segment, "revenue"))
        if rooms_col is None and adr_col is None and rev_col is None:
            continue

        rooms = pd.Series(numeric(rooms_col)) if rooms_col is not None else pd.Series(0.0, index=out.index)
        if rev_col is not None:
            revenue = pd.Series(numeric(rev_col))
        elif adr_col is not None:
            revenue = rooms * pd.Series(numeric(adr_col))
        else:
            revenue = pd.Series(0.0, index=out.index)

        # ADR is always recomputed from rooms and revenue. Some exports carry a
        # revenue figure on days with zero rooms, which makes the stored ADR
        # blank or misleading.
        adr = pd.Series(0.0, index=out.index)
        mask = rooms > 0
        adr.loc[mask] = revenue.loc[mask] / rooms.loc[mask]

        stray = int(((rooms <= 0) & (revenue.abs() > 0.005)).sum())
        if stray:
            notes.append(
                f"{segment}: {stray} day(s) had revenue with no rooms sold - revenue kept, ADR left at 0"
            )
        # Rooms against zero or negative revenue give no usable rate. Negative
        # is the common case -- refunds and adjustments posted to the night.
        unpriced = (rooms > 0) & (revenue <= 0.005)
        if unpriced.any():
            biggest = int(rooms[unpriced].max())
            negatives = int((unpriced & (revenue < -0.005)).sum())
            notes.append(
                f"{segment}: {int(unpriced.sum())} day(s) had rooms sold with no usable "
                f"rate ({negatives} of them negative revenue, largest {biggest} rooms) - "
                "the forecast prices those nights at the surrounding period's ADR"
            )

        out[f"{segment} Rooms"] = rooms.to_numpy()
        out[f"{segment} Revenue"] = revenue.to_numpy()
        out[f"{segment} ADR"] = adr.to_numpy()
        for metric, col in (("rooms", rooms_col), ("adr", adr_col), ("revenue", rev_col)):
            if col is not None:
                column_map[(segment, metric)] = str(best_raw.iloc[best_header][col])

    comp_col = best_map.get(("__comp__", "rooms"))
    out["Comp Rooms"] = numeric(comp_col) if comp_col is not None else 0.0

    out = out.groupby("Date", as_index=False).sum(numeric_only=True).sort_values("Date")
    out = out.reset_index(drop=True)

    # Re-derive ADR after any same-date aggregation.
    for segment in present_segments(out):
        rooms = out[f"{segment} Rooms"]
        out[f"{segment} ADR"] = 0.0
        mask = rooms > 0
        out.loc[mask, f"{segment} ADR"] = out.loc[mask, f"{segment} Revenue"] / rooms[mask]

    gaps = _missing_days(out["Date"])
    if gaps:
        notes.append(f"{gaps} calendar day(s) missing from the history and treated as zero")

    out = _fill_calendar(out)
    return LoadResult(
        daily=out,
        sheet_name=str(best_name),
        header_row=best_header,
        column_map=column_map,
        comp_column=str(best_raw.iloc[best_header][comp_col]) if comp_col is not None else None,
        notes=notes,
    )


def present_segments(daily: pd.DataFrame) -> list[str]:
    return [s for s in SEGMENTS if f"{s} Rooms" in daily.columns]


def _missing_days(dates: pd.Series) -> int:
    if dates.empty:
        return 0
    full = pd.date_range(dates.min(), dates.max(), freq="D")
    return len(full.difference(pd.DatetimeIndex(dates)))


def _fill_calendar(daily: pd.DataFrame) -> pd.DataFrame:
    full = pd.date_range(daily["Date"].min(), daily["Date"].max(), freq="D")
    out = daily.set_index("Date").reindex(full).fillna(0.0)
    out.index.name = "Date"
    return out.reset_index()


def add_totals(daily: pd.DataFrame, segments: list[str] | None = None) -> pd.DataFrame:
    """Append Total Rooms / Revenue / ADR columns for the given segments."""
    segments = segments or present_segments(daily)
    out = daily.copy()
    out["Total Rooms"] = out[[f"{s} Rooms" for s in segments]].sum(axis=1)
    out["Total Revenue"] = out[[f"{s} Revenue" for s in segments]].sum(axis=1)
    out["Total ADR"] = 0.0
    mask = out["Total Rooms"] > 0
    out.loc[mask, "Total ADR"] = out.loc[mask, "Total Revenue"] / out.loc[mask, "Total Rooms"]
    return out


def infer_capacity(daily: pd.DataFrame, segments: list[str] | None = None) -> int:
    """Best guess at rooms available, from the busiest night on record.

    Sold rooms plus comp rooms hit the physical room count on peak nights, so the
    observed maximum is a reliable floor for capacity.
    """
    segments = segments or present_segments(daily)
    occupied = daily[[f"{s} Rooms" for s in segments]].sum(axis=1) + daily.get(
        "Comp Rooms", pd.Series(0.0, index=daily.index)
    )
    if occupied.empty:
        return 0
    return int(round(occupied.max()))
