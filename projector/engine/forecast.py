"""Day-by-day projection engine.

Method, per segment:

1.  Every forecast date is anchored to historical dates exactly 364 days apart
    (52 weeks), so the anchor always lands on the same weekday and close to the
    same point in the season. Weekday pattern is the single biggest driver of a
    hotel's daily shape, so calendar-date matching is the wrong default.
2.  Up to four prior years can be blended. Each year is escalated forward to the
    forecast year by the growth assumptions before blending, so an older year
    contributes its shape without dragging the level down.
3.  A single day of history is noisy, so ROOMS are blended under a smoothing
    slider with a smoothed version of themselves: the local demand level over a
    window of calendar days, times that date's day-of-week index. 0 keeps the
    literal prior-year day, 1 uses the pure level x weekday shape. Spreading
    demand across the surrounding period while holding the weekday shape is what
    keeps a one-off spike from being copied forward without flattening the week.
    ADR is never smoothed -- a rate is a rate, and revenue falls out of rooms x
    ADR. A night that sold nothing has no rate of its own and takes the local
    revenue-over-rooms level instead.
4.  Growth is applied multiplicatively as global x per-segment x per-month, and
    compounded by the number of years the anchor sits behind the forecast date.
5.  Total rooms are capped at capacity x the occupancy ceiling, scaled back
    proportionally across segments, then rounded so each day's segment rooms are
    whole numbers.

Steps 1-3 depend only on the history, so they are split into `build_model`.
Steps 4-5 are applied by `expand`. That split is what lets the exported workbook
ship the base as constants and rebuild the growth half as live Excel formulas.
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from .loader import DEFAULT_FORECAST_SEGMENTS, present_segments

WEEK_YEAR = 364  # 52 weeks -- preserves day of week
MONTH_ABBR = [pd.Timestamp(2000, m, 1).strftime("%b") for m in range(1, 13)]


@dataclass
class Assumptions:
    capacity: int
    start: pd.Timestamp
    end: pd.Timestamp
    segments: list[str] = field(default_factory=lambda: list(DEFAULT_FORECAST_SEGMENTS))
    adr_growth: float = 0.0           # global adjustment on top of the derived rates
    occ_growth: float = 0.0           # global rooms-sold adjustment
    # Per-segment rates are derived from the history by engine.growth, then edited.
    segment_adr_growth: dict[str, float] = field(default_factory=dict)
    segment_occ_growth: dict[str, float] = field(default_factory=dict)
    month_adr_growth: dict[int, float] = field(default_factory=dict)  # 1-12 -> decimal
    month_occ_growth: dict[int, float] = field(default_factory=dict)
    year_weights: list[float] = field(default_factory=lambda: [1.0])
    smoothing: float = 0.35           # 0 = literal prior-year day, 1 = fully smoothed
    # Per-segment smoothing wins over the global value when present. Segments
    # differ enormously here: Government is nearly pure noise, Transient's
    # "spikes" are holidays that recur every year.
    segment_smoothing: dict[str, float] = field(default_factory=dict)
    smoothing_days: int = 10          # half-width of the demand window, in calendar days
    occ_ceiling: float = 1.0          # max share of capacity that can be sold
    round_rooms: bool = True
    # Recorded for the audit trail -- how the per-segment rates were derived.
    growth_basis: str = ""
    growth_damping: float = 1.0
    growth_cap: float | None = None

    def smoothing_for(self, segment: str) -> float:
        """Per-segment smoothing, falling back to the global slider."""
        return float(np.clip(self.segment_smoothing.get(segment, self.smoothing), 0.0, 1.0))

    def effective(self, segment: str, month: int) -> tuple[float, float]:
        """Combined ADR and occupancy growth factors for one segment and month."""
        adr = (
            (1 + self.adr_growth)
            * (1 + self.segment_adr_growth.get(segment, 0.0))
            * (1 + self.month_adr_growth.get(month, 0.0))
        )
        occ = (
            (1 + self.occ_growth)
            * (1 + self.segment_occ_growth.get(segment, 0.0))
            * (1 + self.month_occ_growth.get(month, 0.0))
        )
        return adr, occ


def default_horizon(history_end: pd.Timestamp) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Twelve months starting the first of the month after the last full month."""
    history_end = pd.Timestamp(history_end)
    last_full_month_end = history_end.to_period("M").to_timestamp("M")
    if history_end < last_full_month_end:  # partial trailing month -- step back one
        last_full_month_end = (history_end.to_period("M") - 1).to_timestamp("M")
    start = (last_full_month_end + pd.Timedelta(days=1)).normalize()
    end = (start + pd.DateOffset(years=1) - pd.Timedelta(days=1)).normalize()
    return start, end


def excel_round(values: np.ndarray) -> np.ndarray:
    """Round half away from zero the way Excel's ROUND does.

    Excel carries 15 significant decimal digits, so it sees a cumulative sum of
    1232.4999999999998 as 1232.5 and rounds it up, where plain IEEE arithmetic
    rounds it down. Normalising to 15 significant digits first reproduces that,
    and keeps the app and the exported workbook on the same room counts -- two
    days of the Rutland year land on exactly this knife edge.
    """
    normalised = np.array([float(f"{v:.15g}") for v in np.asarray(values, dtype=float)])
    return np.floor(normalised + 0.5)


def snap_demand_window(half_days: int) -> int:
    """Round a half-window to one whose full width is an odd multiple of 7 days.

    The level has to average out the weekly cycle before a weekday index is
    multiplied back on, and only a whole number of weeks does that. An 11-day
    window (+/- 5) still carries part of the cycle, which double-counts against
    the index and quietly loses a couple of percent of the year. Valid half
    widths are therefore 3, 10, 17, 24 ... days -- 1, 3, 5, 7 weeks wide.
    """
    weeks_back = max(round((int(half_days) - 3) / 7), 0)
    return 7 * weeks_back + 3


def _same_weekday_mean(values: np.ndarray, half_weeks: int) -> np.ndarray:
    """Mean of each weekday over +/- half_weeks weeks, ignoring missing slots."""
    n = len(values)
    offsets = [7 * i for i in range(-half_weeks, half_weeks + 1)]
    stacked = np.full((len(offsets), n), np.nan)
    for i, off in enumerate(offsets):
        # A history shorter than the window has no observation that far out.
        if abs(off) >= n:
            continue
        if off >= 0:
            stacked[i, : n - off] = values[off:]
        else:
            stacked[i, -off:] = values[: n + off]
    # An all-NaN column is a date with no usable observation of its weekday --
    # legitimate, and handled by the caller, so do not warn about it.
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", RuntimeWarning)
        return np.nanmean(stacked, axis=0)


def _centred_mean(values: np.ndarray, half_days: int) -> np.ndarray:
    """Centred calendar-day rolling mean, shrinking the window at the edges."""
    window = 2 * half_days + 1
    return (
        pd.Series(values)
        .rolling(window, center=True, min_periods=1)
        .mean()
        .to_numpy(dtype=float)
    )


def segment_reliability(
    daily: pd.DataFrame, segments: list[str], half_days: int = 10
) -> dict[str, float]:
    """How much of each segment's day-level wobble is real, recurring signal.

    Strip the local level and the day-of-week shape from the history and what is
    left is each date's own deviation. Correlating that residual against the same
    date 364 days earlier is a test-retest reliability: two independent looks at
    the same underlying date effect. Because the noise in the two years is
    independent, the correlation *is* the share of residual variance that is
    signal rather than noise.

    High means the deviations recur -- holidays, annual group business -- and
    smoothing would erase something real. Near zero means the segment is noise
    and should be smoothed hard.
    """
    profiles = _local_profiles(daily, segments, half_days)
    out: dict[str, float] = {}
    for seg in segments:
        rooms = daily[f"{seg} Rooms"].to_numpy(dtype=float)
        smooth = profiles[seg]["rooms"].to_numpy(dtype=float)
        residual = np.divide(
            rooms, smooth, out=np.full_like(rooms, np.nan), where=smooth > 0.5
        )
        if len(residual) <= WEEK_YEAR:
            out[seg] = 0.0
            continue
        this_year, last_year = residual[WEEK_YEAR:], residual[:-WEEK_YEAR]
        usable = np.isfinite(this_year) & np.isfinite(last_year)
        if usable.sum() < 30 or np.std(this_year[usable]) == 0 or np.std(last_year[usable]) == 0:
            out[seg] = 0.0
            continue
        out[seg] = float(np.corrcoef(this_year[usable], last_year[usable])[0, 1])
    return out


def suggested_smoothing(reliability: dict[str, float]) -> dict[str, float]:
    """Smoothing weight per segment: shrink a day toward the smooth estimate by
    however much of it is noise.

    With reliability ``r`` as the signal share, the variance-minimising blend
    puts weight ``r`` on the observed day and ``1 - r`` on the smoothed one, so
    the smoothing slider wants to sit at ``1 - r``. Capped at 0.9 so no segment
    loses its shape entirely.
    """
    # Rounded to two places so the value shown in the grid, written to the
    # workbook and used by the engine are all the same number.
    return {
        seg: round(float(np.clip(1.0 - value, 0.0, 0.9)), 2)
        for seg, value in reliability.items()
    }


def _local_profiles(
    daily: pd.DataFrame, segments: list[str], half_days: int = 10, dow_weeks: int = 6
) -> dict[str, pd.DataFrame]:
    """Split rooms into a local demand level and a day-of-week shape.

    A hotel's daily pattern is two things at once: how busy the *period* is, and
    how the week is shaped inside it. Smoothing has to respect both, so this
    separates them:

    * **Level** -- a centred calendar-day mean over ``+/- half_days``, snapped
      to a whole number of weeks by `snap_demand_window`. This is the "demand in
      this stretch of the season" figure. Widen it to smooth over a longer time
      frame, narrow it to track short swings.
    * **Day-of-week index** -- each day's ratio to its own local level, averaged
      across the same weekday over ``+/- dow_weeks`` weeks, then renormalised so
      the seven indices around any date average to 1.

    Smoothed rooms are level x index. Averaging raw calendar days instead would
    blend Saturdays into midweek and flatten the week; multiplying a weekday
    shape onto a local level spreads a spike across the surrounding period *and*
    keeps Saturday looking like a Saturday.

    Because the level is a centred mean and the index is normalised to 1, the
    total is preserved except where the window runs off the ends of the history.
    """
    index = pd.DatetimeIndex(pd.to_datetime(daily["Date"]))
    half_days = snap_demand_window(half_days)
    out: dict[str, pd.DataFrame] = {}

    for seg in segments:
        rooms = daily[f"{seg} Rooms"].to_numpy(dtype=float)
        revenue = daily[f"{seg} Revenue"].to_numpy(dtype=float)

        # Only rooms are smoothed. A rate is a rate -- it does not need demand
        # spread across a period the way volume does, and revenue falls out of
        # rooms x ADR downstream. `rate` here is not a smoothed ADR: it is the
        # local revenue-over-rooms level, used only as the fallback for a day
        # that sold nothing and therefore has no rate of its own.
        rooms_level = _centred_mean(rooms, half_days)
        rev_level = _centred_mean(revenue, half_days)
        rate = np.divide(
            rev_level, rooms_level, out=np.zeros_like(rev_level), where=rooms_level > 0
        )

        out[seg] = pd.DataFrame(
            {
                "rooms": np.maximum(rooms_level * _dow_index(rooms, rooms_level, dow_weeks), 0.0),
                "rate": np.maximum(rate, 0.0),
            },
            index=index,
        )

    return out


def _dow_index(values: np.ndarray, level: np.ndarray, dow_weeks: int) -> np.ndarray:
    """Weekday shape as a ratio to the local level, normalised to average 1.

    This is a ratio of averages, not an average of ratios. Averaging the daily
    ``value / level`` overweights days when the level happens to be low and
    exaggerates the shape: on Group it produced a Friday-to-Monday index of 5.4x
    against a true five-year ratio of 3.09x, which then over-peaked every
    projected Friday.
    """
    mean_values = _same_weekday_mean(values, dow_weeks)
    mean_level = _same_weekday_mean(level, dow_weeks)
    dow = np.divide(
        mean_values, mean_level, out=np.ones_like(mean_values), where=mean_level > 0
    )
    dow = np.nan_to_num(dow, nan=1.0)
    # The seven weekday indices around any date must average to 1, otherwise the
    # shape would quietly move the level.
    normaliser = _centred_mean(dow, 3)
    return np.divide(dow, normaliser, out=np.ones_like(dow), where=normaliser > 0)


def _anchor_years(dates: pd.DatetimeIndex, history_end: pd.Timestamp) -> np.ndarray:
    """For each forecast date, how many 364-day steps back reach the history."""
    deltas = (dates - pd.Timestamp(history_end)).days.to_numpy()
    steps = np.ceil(np.maximum(deltas, 1) / WEEK_YEAR).astype(int)
    return np.maximum(steps, 1)


@dataclass
class ForecastModel:
    """The growth-free half of the projection.

    Everything here comes out of the history alone -- the blended, smoothed base
    rooms and ADR for each date and each blended year. No growth assumption
    touches it. Growth is applied on top by `expand`, which is what lets the
    exported workbook carry these as constants and rebuild the rest as live
    Excel formulas off the Assumptions sheet.
    """

    dates: pd.DatetimeIndex
    segments: list[str]
    weights: np.ndarray                     # normalised blend weights, one per year
    years_back: np.ndarray                  # 364-day steps from each date to year 1
    base_rooms: dict[str, np.ndarray]       # segment -> (n_years, n_dates)
    base_adr: dict[str, np.ndarray]
    unanchored: np.ndarray                  # dates with no usable history at all


def build_model(daily: pd.DataFrame, assumptions: Assumptions) -> ForecastModel:
    """Blend and smooth the history into a per-date base. Growth-independent."""
    segments = [s for s in assumptions.segments if f"{s} Rooms" in daily.columns]
    if not segments:
        raise ValueError("No forecastable segments selected.")

    history = daily.set_index(pd.to_datetime(daily["Date"]))
    history_start, history_end = history.index.min(), history.index.max()

    dates = pd.date_range(assumptions.start, assumptions.end, freq="D")
    if len(dates) == 0:
        raise ValueError("Forecast end date must fall on or after the start date.")

    weights = np.asarray(assumptions.year_weights, dtype=float)
    weights = weights[weights > 0]
    if weights.size == 0:
        weights = np.array([1.0])
    weights = weights / weights.sum()

    years_back = _anchor_years(dates, history_end)
    profiles = _local_profiles(daily, segments, assumptions.smoothing_days)

    n_years, n_dates = len(weights), len(dates)
    anchors_by_year, valid_by_year = [], []
    for j in range(n_years):
        anchors = dates - pd.to_timedelta((years_back + j) * WEEK_YEAR, unit="D")
        anchors_by_year.append(anchors)
        valid_by_year.append((anchors >= history_start) & (anchors <= history_end))

    # Where a blended year reaches past the start of the history, fall back to
    # the nearest year that does exist rather than dropping it. Renormalising
    # per date instead would make the blend weights date-dependent, which no
    # spreadsheet formula could then reproduce.
    resolved = []
    for j in range(n_years):
        anchors, valid = anchors_by_year[j], valid_by_year[j]
        chosen = anchors.to_numpy().copy()
        for k in range(j - 1, -1, -1):
            gap = ~valid & valid_by_year[k]
            if gap.any():
                chosen[gap] = anchors_by_year[k].to_numpy()[gap]
                valid = valid | gap
        resolved.append((pd.DatetimeIndex(chosen), valid))

    unanchored = ~np.any([v for _, v in resolved], axis=0)

    base_rooms: dict[str, np.ndarray] = {}
    base_adr: dict[str, np.ndarray] = {}
    for seg in segments:
        rooms_hist = history[f"{seg} Rooms"]
        rev_hist = history[f"{seg} Revenue"]
        profile = profiles[seg]
        smoothing = assumptions.smoothing_for(seg)

        rooms_stack = np.zeros((n_years, n_dates))
        adr_stack = np.zeros((n_years, n_dates))

        for j, (anchors, valid) in enumerate(resolved):
            raw_rooms = np.nan_to_num(rooms_hist.reindex(anchors).to_numpy(dtype=float))
            raw_rev = np.nan_to_num(rev_hist.reindex(anchors).to_numpy(dtype=float))
            raw_adr = np.divide(
                raw_rev, raw_rooms, out=np.zeros_like(raw_rev), where=raw_rooms > 0
            )

            smooth = profile.reindex(anchors)
            smooth_rooms = np.nan_to_num(smooth["rooms"].to_numpy(dtype=float))
            local_rate = np.nan_to_num(smooth["rate"].to_numpy(dtype=float))

            # Calibrate the smoothed series to carry exactly the raw total over
            # the nights actually being projected. Smoothing is meant to move
            # volume around, not create or destroy it, and any difference here is
            # the smoother running off the ends of the history rather than
            # signal. With the totals equal, every blend of the two carries the
            # same total, so the year is invariant to the smoothing slider.
            raw_total = raw_rooms[valid].sum()
            smooth_total = smooth_rooms[valid].sum()
            if smooth_total > 0 and raw_total > 0:
                smooth_rooms = smooth_rooms * (raw_total / smooth_total)

            # Smoothing moves volume around, never the rate. A non-positive ADR
            # means the night carries no usable rate: nothing was sold, or the
            # export booked rooms against zero or negative revenue. The Rutland
            # history has 60 such nights -- 9 at zero revenue and 51 at negative
            # revenue from refunds, one of them 22 rooms at -$42.50. All of them
            # take the local revenue-over-rooms rate, because otherwise those
            # rooms would be budgeted at zero or at a negative rate.
            rooms_stack[j] = np.where(
                valid, (1 - smoothing) * raw_rooms + smoothing * smooth_rooms, 0.0
            )
            adr_stack[j] = np.where(valid, np.where(raw_adr > 0, raw_adr, local_rate), 0.0)

        base_rooms[seg] = np.maximum(rooms_stack, 0.0)
        base_adr[seg] = np.maximum(adr_stack, 0.0)

    return ForecastModel(
        dates=dates,
        segments=segments,
        weights=weights,
        years_back=years_back,
        base_rooms=base_rooms,
        base_adr=base_adr,
        unanchored=unanchored,
    )


def expand(model: ForecastModel, assumptions: Assumptions) -> dict[str, np.ndarray]:
    """Apply growth to a model and return every intermediate, keyed by column name.

    The exported workbook mirrors these column for column, so the values written
    as cached formula results always agree with what the app shows.
    """
    dates = model.dates
    months = dates.month.to_numpy()
    out: dict[str, np.ndarray] = {}

    raw_rooms_by_seg: dict[str, np.ndarray] = {}
    adr_by_seg: dict[str, np.ndarray] = {}

    for seg in model.segments:
        adr_factor, occ_factor = _factor_arrays(assumptions, seg, months)
        out[f"{seg} Occ factor"] = occ_factor
        out[f"{seg} ADR factor"] = adr_factor

        rooms_acc = np.zeros(len(dates))
        adr_acc = np.zeros(len(dates))
        for j, w in enumerate(model.weights):
            steps = model.years_back + j
            out[f"{seg} Base Rooms Y{j + 1}"] = model.base_rooms[seg][j]
            out[f"{seg} Base ADR Y{j + 1}"] = model.base_adr[seg][j]
            rooms_acc += w * model.base_rooms[seg][j] * occ_factor**steps
            adr_acc += w * model.base_adr[seg][j] * adr_factor**steps

        total_weight = model.weights.sum()
        raw_rooms_by_seg[seg] = np.maximum(rooms_acc / total_weight, 0.0)
        adr_by_seg[seg] = np.maximum(adr_acc / total_weight, 0.0)
        out[f"{seg} Raw Rooms"] = raw_rooms_by_seg[seg]

    raw_total = np.sum([raw_rooms_by_seg[s] for s in model.segments], axis=0)
    ceiling = max(assumptions.capacity * assumptions.occ_ceiling, 0.0)
    scale = np.ones(len(dates))
    if ceiling > 0:
        over = raw_total > ceiling
        scale[over] = ceiling / raw_total[over]
    out["Raw Total Rooms"] = raw_total
    out["Capacity Scale"] = scale

    for seg in model.segments:
        scaled = raw_rooms_by_seg[seg] * scale
        out[f"{seg} Cum Rooms"] = np.cumsum(scaled)
        if assumptions.round_rooms:
            # Round the running total, not each day on its own, and take the
            # difference. Rounding days independently leaks volume: smoothing
            # spreads rooms into many days holding a fraction under a half, and
            # each of those rounds away to nothing. Rounding the cumulative
            # series carries the remainder forward instead, so the year total is
            # the rounded exact total and cannot drift with the smoothing
            # slider. Half away from zero throughout, matching Excel's ROUND.
            rounded_cum = excel_round(out[f"{seg} Cum Rooms"])
            rooms = np.diff(np.concatenate([[0.0], rounded_cum]))
        else:
            rooms = scaled
        out[f"{seg} Rooms"] = rooms
        out[f"{seg} ADR"] = adr_by_seg[seg]
        out[f"{seg} Revenue"] = rooms * adr_by_seg[seg]

    total_rooms = np.sum([out[f"{s} Rooms"] for s in model.segments], axis=0)
    total_revenue = np.sum([out[f"{s} Revenue"] for s in model.segments], axis=0)
    out["Total Rooms"] = total_rooms
    out["Total Revenue"] = total_revenue
    out["Total ADR"] = np.divide(
        total_revenue, total_rooms, out=np.zeros_like(total_revenue), where=total_rooms > 0
    )
    capacity = float(assumptions.capacity)
    out["Rooms Available"] = np.full(len(dates), capacity)
    out["OCC %"] = total_rooms / capacity * 100 if capacity else np.zeros(len(dates))
    out["RevPAR"] = total_revenue / capacity if capacity else np.zeros(len(dates))
    return out


def public_columns(segments: list[str]) -> list[str]:
    """The columns the app shows: helpers and per-year bases stay out of the way."""
    cols = ["Date", "Day", "Month"]
    for seg in segments:
        cols += [f"{seg} Rooms", f"{seg} ADR"]
    for seg in segments:
        cols.append(f"{seg} Revenue")
    return cols + ["Total Rooms", "Total Revenue", "Total ADR", "Rooms Available", "OCC %", "RevPAR"]


def forecast_frame(model: ForecastModel, assumptions: Assumptions) -> pd.DataFrame:
    """The day-by-day frame the app displays, from an already-built model."""
    values = expand(model, assumptions)
    out = pd.DataFrame({"Date": model.dates})
    out["Day"] = model.dates.strftime("%a")
    out["Month"] = model.dates.to_period("M").astype(str)
    for column in public_columns(model.segments):
        if column not in out.columns:
            out[column] = values[column]
    return out


def build_forecast(daily: pd.DataFrame, assumptions: Assumptions) -> pd.DataFrame:
    """Return a day-by-day projection frame for the assumption set."""
    return forecast_frame(build_model(daily, assumptions), assumptions)


def _factor_arrays(assumptions: Assumptions, segment: str, months: np.ndarray):
    adr = np.ones(len(months))
    occ = np.ones(len(months))
    for month in np.unique(months):
        a, o = assumptions.effective(segment, int(month))
        mask = months == month
        adr[mask] = a
        occ[mask] = o
    return adr, occ


def compare_to_history(
    forecast: pd.DataFrame,
    daily: pd.DataFrame,
    capacity: int,
    segments: list[str],
) -> pd.DataFrame:
    """Month-by-month forecast vs the matching 364-day-back actuals."""
    from .history import monthly

    hist_monthly = monthly(daily, capacity, segments).set_index("Month")

    fc = forecast.copy()
    anchor = pd.to_datetime(fc["Date"]) - pd.Timedelta(days=WEEK_YEAR)
    fc["Anchor Month"] = anchor.dt.to_period("M").astype(str)

    rows = []
    for month, block in fc.groupby("Month", sort=True):
        anchor_month = block["Anchor Month"].mode().iat[0]
        prior = hist_monthly.loc[anchor_month] if anchor_month in hist_monthly.index else None
        rooms = block["Total Rooms"].sum()
        revenue = block["Total Revenue"].sum()
        available = len(block) * capacity
        rows.append(
            {
                "Month": month,
                "Rooms": rooms,
                "ADR": revenue / rooms if rooms else 0.0,
                "Revenue": revenue,
                "OCC %": rooms / available * 100 if available else 0.0,
                "RevPAR": revenue / available if available else 0.0,
                "LY Month": anchor_month,
                "LY Rooms": prior["Total Rooms"] if prior is not None else np.nan,
                "LY ADR": prior["Total ADR"] if prior is not None else np.nan,
                "LY Revenue": prior["Total Revenue"] if prior is not None else np.nan,
                "LY OCC %": prior["OCC %"] if prior is not None else np.nan,
            }
        )

    out = pd.DataFrame(rows)
    for col in ("Rooms", "ADR", "Revenue"):
        out[f"{col} vs LY %"] = (out[col] / out[f"LY {col}"] - 1) * 100
    out["OCC pts vs LY"] = out["OCC %"] - out["LY OCC %"]
    return out
