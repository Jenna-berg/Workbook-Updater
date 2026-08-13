"""Derive forward growth rates per segment from the historical data itself.

The projection's growth assumptions are read out of the history rather than
typed in. Which basis you read them on matters enormously -- on the Rutland file
Ext Stay runs at -24.9% a year over four years but +13.0% over the last one --
so the basis is an explicit, switchable assumption and every basis is shown
side by side rather than hidden behind one number.

Only complete years are used. A part year would understate every rate.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

BASES = [
    "Most recent year",
    "3-year CAGR",
    "Full-history CAGR",
    "Weighted trend (recent years count more)",
]
DEFAULT_BASIS = "3-year CAGR"
METRICS = ("Rooms", "ADR")


def _cagr(values: np.ndarray, years: int) -> float:
    """Compound annual growth over the last `years` steps of the series."""
    if len(values) < years + 1:
        return float("nan")
    start, end = values[-1 - years], values[-1]
    if start <= 0 or end <= 0:
        return float("nan")
    return (end / start) ** (1 / years) - 1


def _weighted_trend(values: np.ndarray, half_life: float = 1.5) -> float:
    """Slope of a log-linear fit with exponentially decaying weights.

    Uses every year but lets the recent ones dominate, so one odd year early in
    the series cannot set the whole forward rate.
    """
    usable = values > 0
    if usable.sum() < 2:
        return float("nan")
    idx = np.arange(len(values), dtype=float)[usable]
    logs = np.log(values[usable])
    weights = 0.5 ** ((idx.max() - idx) / half_life)
    mean_x = np.average(idx, weights=weights)
    mean_y = np.average(logs, weights=weights)
    variance = np.average((idx - mean_x) ** 2, weights=weights)
    if variance <= 0:
        return float("nan")
    slope = np.average((idx - mean_x) * (logs - mean_y), weights=weights) / variance
    return float(np.exp(slope) - 1)


def _rate(values: np.ndarray, basis: str) -> float:
    if basis == "Most recent year":
        return _cagr(values, 1)
    if basis == "3-year CAGR":
        return _cagr(values, 3)
    if basis == "Full-history CAGR":
        return _cagr(values, max(len(values) - 1, 1))
    if basis == "Weighted trend (recent years count more)":
        return _weighted_trend(values)
    raise ValueError(f"Unknown growth basis: {basis}")


def complete_years(annual: pd.DataFrame) -> pd.DataFrame:
    return annual[annual["Complete"]].reset_index(drop=True)


def years_for_basis(basis: str, year_labels: list[str]) -> list[str]:
    """Which year labels a basis actually reads.

    A 3-year CAGR spans four year-ends, not every complete year on file, and
    saying otherwise overstates how much history is behind the number.
    """
    if basis == "Most recent year":
        return year_labels[-2:]
    if basis == "3-year CAGR":
        return year_labels[-4:]
    return year_labels


def all_bases(annual: pd.DataFrame, segments: list[str]) -> pd.DataFrame:
    """Every basis for every segment and metric, as percentages, for comparison.

    The leading ``Key`` column is what the exported workbook's Assumptions sheet
    looks up with MATCH, so the basis dropdown in Excel can re-derive every
    segment rate without an array formula. Column order is part of that
    contract: Key, Segment, Metric, then one column per basis.
    """
    years = complete_years(annual)
    rows = []
    for seg in [*segments, "Total"]:
        for metric in METRICS:
            column = f"Total {metric}" if seg == "Total" else f"{seg} {metric}"
            if column not in years.columns:
                continue
            values = years[column].to_numpy(dtype=float)
            row = {"Key": f"{seg}|{metric}", "Segment": seg, "Metric": metric}
            row.update({basis: _rate(values, basis) * 100 for basis in BASES})
            rows.append(row)
    return pd.DataFrame(rows)


def yearly_growth(annual: pd.DataFrame, segments: list[str], metric: str) -> pd.DataFrame:
    """Year-over-year % per segment, one row per segment, one column per year."""
    years = complete_years(annual)
    rows = []
    for seg in [*segments, "Total"]:
        column = f"Total {metric}" if seg == "Total" else f"{seg} {metric}"
        if column not in years.columns:
            continue
        series = years[column].astype(float)
        pct = series.pct_change() * 100
        row = {"Segment": seg}
        row.update({label: pct.iat[i] for i, label in enumerate(years["Year"])})
        rows.append(row)
    out = pd.DataFrame(rows)
    return out.drop(columns=[c for c in out.columns if out[c].isna().all() and c != "Segment"])


def derive(
    annual: pd.DataFrame,
    segments: list[str],
    basis: str = DEFAULT_BASIS,
    damping: float = 1.0,
    cap: float | None = 0.15,
) -> pd.DataFrame:
    """Forward growth per segment on one basis, damped and clamped.

    `damping` scales the raw rate (0.5 = "budget half the historical trend").
    `cap` clamps the result to +/- that decimal; an unclamped -25% a year
    compounds a segment most of the way out of existence over a few years.
    Returns a frame indexed by segment with the raw and applied rates as
    percentages, so the UI can show what was trimmed and why.
    """
    years = complete_years(annual)
    rows = []
    for seg in segments:
        row: dict[str, object] = {"Segment": seg}
        for metric in METRICS:
            column = f"{seg} {metric}"
            raw = (
                _rate(years[column].to_numpy(dtype=float), basis)
                if column in years.columns
                else float("nan")
            )
            applied = 0.0 if np.isnan(raw) else raw * damping
            if cap is not None:
                applied = float(np.clip(applied, -cap, cap))
            # Two decimals as a percentage. The UI grid and the exported
            # workbook's formula both round to the same place, so a rate never
            # reads 3.93% on screen while driving 3.9271% underneath.
            applied = round(applied, 4)
            row[f"{metric} historical %"] = raw * 100
            row[f"{metric} applied %"] = applied * 100
            row[f"{metric} capped"] = (
                cap is not None and not np.isnan(raw) and abs(raw * damping) > cap + 1e-12
            )
        rows.append(row)

    out = pd.DataFrame(rows)
    out.attrs["basis"] = basis
    out.attrs["years_used"] = years_for_basis(basis, list(years["Year"]))
    out.attrs["years_available"] = list(years["Year"])
    return out


def as_assumption_dicts(derived: pd.DataFrame) -> tuple[dict[str, float], dict[str, float]]:
    """Split a derived table into the (adr_growth, occ_growth) dicts Assumptions wants."""
    adr = {r["Segment"]: r["ADR applied %"] / 100 for _, r in derived.iterrows()}
    occ = {r["Segment"]: r["Rooms applied %"] / 100 for _, r in derived.iterrows()}
    return adr, occ
