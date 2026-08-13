"""Historical roll-ups: monthly, annual, segment mix and year-over-year growth."""

from __future__ import annotations

import pandas as pd

from .loader import present_segments


def year_label(dates: pd.Series, fiscal_start_month: int) -> pd.Series:
    """Label each date with the year it belongs to.

    With ``fiscal_start_month = 1`` this is the calendar year. Otherwise a year
    runs from that month through the following month-minus-one and is labelled
    e.g. ``Apr21-Mar22``.
    """
    dates = pd.to_datetime(dates)
    if fiscal_start_month == 1:
        return dates.dt.year.astype(str)

    start_year = dates.dt.year.where(dates.dt.month >= fiscal_start_month, dates.dt.year - 1)
    start_abbr = pd.Timestamp(2000, fiscal_start_month, 1).strftime("%b")
    end_month = 12 if fiscal_start_month == 1 else fiscal_start_month - 1
    end_abbr = pd.Timestamp(2000, end_month, 1).strftime("%b")
    return (
        start_abbr
        + start_year.mod(100).map("{:02d}".format)
        + "-"
        + end_abbr
        + (start_year + 1).mod(100).map("{:02d}".format)
    )


def _aggregate(frame: pd.DataFrame, group_key, segments: list[str], capacity: int) -> pd.DataFrame:
    agg: dict[str, tuple[str, str]] = {"Days": ("Date", "size")}
    for seg in segments:
        agg[f"{seg} Rooms"] = (f"{seg} Rooms", "sum")
        agg[f"{seg} Revenue"] = (f"{seg} Revenue", "sum")
    if "Comp Rooms" in frame.columns:
        agg["Comp Rooms"] = ("Comp Rooms", "sum")

    out = frame.groupby(group_key, sort=True).agg(**agg)

    out["Total Rooms"] = out[[f"{s} Rooms" for s in segments]].sum(axis=1)
    out["Total Revenue"] = out[[f"{s} Revenue" for s in segments]].sum(axis=1)
    out["Rooms Available"] = out["Days"] * capacity

    for seg in segments:
        out[f"{seg} ADR"] = _safe_div(out[f"{seg} Revenue"], out[f"{seg} Rooms"])
        out[f"{seg} OCC %"] = _safe_div(out[f"{seg} Rooms"], out["Rooms Available"]) * 100
        out[f"{seg} Mix %"] = _safe_div(out[f"{seg} Rooms"], out["Total Rooms"]) * 100

    out["Total ADR"] = _safe_div(out["Total Revenue"], out["Total Rooms"])
    out["OCC %"] = _safe_div(out["Total Rooms"], out["Rooms Available"]) * 100
    out["RevPAR"] = _safe_div(out["Total Revenue"], out["Rooms Available"])
    return out


def _safe_div(num: pd.Series, den: pd.Series) -> pd.Series:
    return (num / den.replace(0, pd.NA)).astype(float).fillna(0.0)


def monthly(daily: pd.DataFrame, capacity: int, segments: list[str] | None = None) -> pd.DataFrame:
    segments = segments or present_segments(daily)
    frame = daily.copy()
    frame["Month"] = pd.to_datetime(frame["Date"]).dt.to_period("M")
    out = _aggregate(frame, "Month", segments, capacity)
    out.index = out.index.astype(str)
    out.index.name = "Month"
    return out.reset_index()


def annual(
    daily: pd.DataFrame,
    capacity: int,
    fiscal_start_month: int = 1,
    segments: list[str] | None = None,
) -> pd.DataFrame:
    """Year totals with year-over-year growth on the headline metrics."""
    segments = segments or present_segments(daily)
    frame = daily.copy()
    frame["Year"] = year_label(frame["Date"], fiscal_start_month)
    out = _aggregate(frame, "Year", segments, capacity)

    expected = _expected_days(daily["Date"], fiscal_start_month)
    out.insert(1, "Complete", [out.loc[y, "Days"] >= expected.get(y, 0) for y in out.index])

    growth_cols = ["Total Rooms", "Total Revenue", "Total ADR", "OCC %", "RevPAR"] + [
        f"{s} {m}" for s in segments for m in ("Rooms", "ADR")
    ]
    for col in growth_cols:
        out[f"{col} Y/Y %"] = out[col].pct_change() * 100
        if col == "OCC %":  # occupancy moves in points, not percent
            out["OCC pts Y/Y"] = out["OCC %"].diff()
            out = out.drop(columns=["OCC % Y/Y %"])

    out.index.name = "Year"
    return out.reset_index()


def _expected_days(dates: pd.Series, fiscal_start_month: int) -> dict[str, int]:
    """Days each labelled year would contain if it were complete."""
    dates = pd.to_datetime(dates)
    full = pd.Series(pd.date_range(dates.min().replace(day=1), dates.max(), freq="D"))
    labels = year_label(full, fiscal_start_month)
    counts: dict[str, int] = {}
    for label in labels.unique():
        block = full[labels == label]
        start = block.min()
        year_start = pd.Timestamp(start.year, fiscal_start_month, 1)
        if start < year_start:
            year_start = pd.Timestamp(start.year - 1, fiscal_start_month, 1)
        counts[label] = (year_start + pd.DateOffset(years=1) - year_start).days
    return counts


def segment_matrix(monthly_frame: pd.DataFrame, segments: list[str], metric: str) -> pd.DataFrame:
    """Month rows x segment columns for one metric, for the segment comparison tab."""
    present = [s for s in segments if f"{s} {metric}" in monthly_frame.columns]
    out = monthly_frame[["Month", *[f"{s} {metric}" for s in present]]].copy()
    # Strip the whole metric suffix -- "Group Mix %" must become "Group", not "Group Mix".
    out.columns = ["Month", *present]
    return out


def same_month_growth(monthly_frame: pd.DataFrame, column: str) -> pd.DataFrame:
    """Month-over-same-month-last-year growth for one column."""
    frame = monthly_frame[["Month", column]].copy()
    period = pd.PeriodIndex(frame["Month"], freq="M")
    frame["Cal Month"] = period.strftime("%b")
    frame["Yr"] = period.year
    wide = frame.pivot_table(index="Cal Month", columns="Yr", values=column, sort=False)
    order = [pd.Timestamp(2000, m, 1).strftime("%b") for m in range(1, 13)]
    return wide.reindex([m for m in order if m in wide.index])
