"""Budget Projector -- day-by-day rooms and ADR budget from 5 years of history."""

from __future__ import annotations

import altair as alt
import pandas as pd
import streamlit as st

from .engine import (
    Assumptions,
    LoaderError,
    build_model,
    compare_to_history,
    default_horizon,
    forecast_frame,
    infer_capacity,
    load_workbook,
    present_segments,
    segment_reliability,
    suggested_smoothing,
)
from .engine import excel as excel_export
from .engine import growth
from .engine import history as hist
from .engine.loader import DEFAULT_FORECAST_SEGMENTS

# set_page_config intentionally dropped — the host app owns it.

MONTHS = [pd.Timestamp(2000, m, 1).strftime("%B") for m in range(1, 13)]
# Full window widths in days. Whole numbers of weeks, so the demand level
# averages the weekly cycle out before the weekday shape is multiplied back.
DEMAND_WINDOWS = [7, 21, 35, 49, 63]
SMOOTH_GROUP_ONLY = "Group only"


@st.cache_data(show_spinner="Reading the workbook...")
def _load(file_bytes: bytes, name: str):  # `name` only widens the cache key
    result = load_workbook(pd.io.common.BytesIO(file_bytes))
    return result.daily, result.sheet_name, result.notes


def _pct(label: str, value: float, key: str, help_text: str | None = None) -> float:
    return st.number_input(label, value=value, step=0.5, format="%.2f", key=key, help=help_text) / 100


def _kpi_row(current: dict[str, float], prior: dict[str, float] | None) -> None:
    cols = st.columns(5)
    specs = [
        ("Rooms sold", "rooms", "{:,.0f}", "{:+.1f}%"),
        ("Occupancy", "occ", "{:.1f}%", "{:+.1f} pts"),
        ("ADR", "adr", "${:,.2f}", "{:+.1f}%"),
        ("RevPAR", "revpar", "${:,.2f}", "{:+.1f}%"),
        ("Total revenue", "revenue", "${:,.0f}", "{:+.1f}%"),
    ]
    for col, (label, key, fmt, delta_fmt) in zip(cols, specs):
        delta = None
        if prior and prior.get(key):
            if key == "occ":
                delta = delta_fmt.format(current[key] - prior[key])
            else:
                delta = delta_fmt.format((current[key] / prior[key] - 1) * 100)
        col.metric(label, fmt.format(current[key]), delta)


def _totals(rooms: float, revenue: float, available: float) -> dict[str, float]:
    return {
        "rooms": rooms,
        "revenue": revenue,
        "adr": revenue / rooms if rooms else 0.0,
        "occ": rooms / available * 100 if available else 0.0,
        "revpar": revenue / available if available else 0.0,
    }


# --------------------------------------------------------------------------- #
# Input
# --------------------------------------------------------------------------- #



def render():
    """Draw the Budget Projector into the current Streamlit container."""
    st.title("🏨 Budget Projector")
    st.caption(
        "Upload a day-by-day segment export, set the growth assumptions, and download "
        "a full day-by-day budget for the year ahead."
    )

    uploaded = st.file_uploader(
        "Day-by-day workbook (.xlsx)",
        type=["xlsx", "xlsm", "xls"],
        help="Needs a date column plus Rooms / ADR / Revenue columns per segment.",
    )

    if uploaded is None:
        st.info(
            "**Expected format** — one row per date, with columns like `Days in Date`, "
            "`Transient Rooms`, `Transient ADR`, `Group Rooms`, `Government Rooms`, "
            "`Corp - Pref Rooms`, `Ext Stay Rooms` and their ADR/Revenue partners. "
            "Column titles are matched by wording, not position, so exact spelling and "
            "column order can vary between hotels. Pivot subtotal rows are ignored."
        )
        return

    try:
        daily, sheet_name, notes = _load(uploaded.getvalue(), uploaded.name)
    except LoaderError as exc:
        st.error(str(exc))
        return

    all_segments = present_segments(daily)
    history_start, history_end = daily["Date"].min(), daily["Date"].max()
    detected_capacity = infer_capacity(daily)
    default_start, default_end = default_horizon(history_end)

    st.success(
        f"Read **{len(daily):,}** days from sheet **{sheet_name}** — "
        f"{history_start:%d %b %Y} to {history_end:%d %b %Y} — "
        f"segments: {', '.join(all_segments)}."
    )
    if notes:
        with st.expander(f"{len(notes)} data note(s) from the import"):
            for note in notes:
                st.write("•", note)

    # --------------------------------------------------------------------------- #
    # Assumptions (sidebar: visible from every tab)
    # --------------------------------------------------------------------------- #

    with st.sidebar:
        st.header("Assumptions")

        capacity = st.number_input(
            "Rooms available (capacity)",
            min_value=1,
            value=int(detected_capacity) or 100,
            step=1,
            help=f"Detected {detected_capacity} from the busiest night on record "
            "(rooms sold plus comp). Override if the true room count differs.",
        )

        st.subheader("History view")
        fiscal_start_month = MONTHS.index(
            st.selectbox(
                "Year starts in", MONTHS, index=int(history_start.month) - 1,
                help="Sets the year boundary for every historical roll-up and for the "
                "growth rates derived from them.",
            )
        ) + 1

        st.subheader("Forecast period")
        col_a, col_b = st.columns(2)
        start = col_a.date_input("Start", value=default_start.date())
        end = col_b.date_input("End", value=default_end.date())

        st.subheader("Growth")
        st.caption("Rates are read out of your history, per segment. Edit them on the front tab.")
        growth_basis = st.selectbox(
            "Derived from", growth.BASES, index=growth.BASES.index(growth.DEFAULT_BASIS),
            help="Which stretch of history sets each segment's forward rate. Only "
            "complete years are used.",
        )
        growth_damping = st.slider(
            "Apply % of the historical rate", 0, 150, 100, 5,
            help="100% budgets the full historical trend. Lower it to budget a "
            "fraction of it.",
        ) / 100
        growth_cap = st.number_input(
            "Cap derived growth at ± %", min_value=0.0, max_value=100.0, value=15.0, step=1.0,
            help="A raw rate like -25% a year compounds a segment most of the way out "
            "of existence. Set 0 to remove the cap.",
        ) / 100 or None

        adr_growth = _pct(
            "ADR adjustment %", 0.0, "adr_g",
            "Extra ADR growth on top of every segment's derived rate.",
        )
        occ_growth = _pct(
            "Occupancy adjustment %", 0.0, "occ_g",
            "Extra rooms-sold growth on top of every segment's derived rate. At a "
            "fixed room count this is the same as occupancy growth.",
        )

        st.subheader("Method")
        forecast_segments = st.multiselect(
            "Segments to project",
            all_segments,
            default=[s for s in DEFAULT_FORECAST_SEGMENTS if s in all_segments],
            help="Group is included and smoothed hard — its day-level detail is noise, but its seasonal shape is solid.",
        )
        blend_choice = st.selectbox(
            "History blended",
            ["Last year only", "2 years (70 / 30)", "3 years (60 / 30 / 10)", "4 years (50 / 25 / 15 / 10)"],
            index=0,
            help="Older years are escalated forward by the growth assumptions before blending.",
        )
        weights = {
            "Last year only": [1.0],
            "2 years (70 / 30)": [0.7, 0.3],
            "3 years (60 / 30 / 10)": [0.6, 0.3, 0.1],
            "4 years (50 / 25 / 15 / 10)": [0.5, 0.25, 0.15, 0.10],
        }[blend_choice]

        smoothing_mode = st.radio(
            "Day smoothing",
            [SMOOTH_GROUP_ONLY, "Auto (per segment)", "Same for all segments"],
            help="Group only smooths the lumpiest segment and leaves every other one "
            "on its literal date last year. Auto sets each segment from how much of "
            "its day-to-day movement repeats year over year — see the Growth Rates tab.",
        )
        smoothing = st.slider(
            "Smoothing", 0.0, 1.0, 0.75, 0.05,
            disabled=smoothing_mode == "Auto (per segment)",
            help="Applies to rooms only — ADR is never smoothed. 0 copies the literal "
            "same-weekday date last year, 1 uses the local demand level times that "
            "date's day-of-week shape. In Group-only mode this is the value Group gets.",
        )
        smoothing_days = st.select_slider(
            "Demand window",
            options=DEMAND_WINDOWS,
            value=21,
            format_func=lambda d: f"{d} days ({d // 7} week{'s' if d > 7 else ''})",
            help="How wide a stretch a day's demand is spread over when smoothing — "
            "the window is centred, so 21 days means the 10 days either side plus "
            "the day itself. Wider is steadier; narrower tracks short swings. "
            "Widths are whole weeks so the weekday shape stays intact.",
        )
        occ_ceiling = st.slider(
            "Occupancy ceiling %", 50, 100, 100, 1,
            help="Days that project above this share of capacity are scaled back.",
        ) / 100
        round_rooms = st.checkbox("Round rooms to whole numbers", value=True)

    if not forecast_segments:
        st.warning("Select at least one segment to project.")
        return

    assumptions = Assumptions(
        capacity=int(capacity),
        start=pd.Timestamp(start),
        end=pd.Timestamp(end),
        segments=forecast_segments,
        adr_growth=adr_growth,
        occ_growth=occ_growth,
        year_weights=weights,
        smoothing=smoothing,
        smoothing_days=(smoothing_days - 1) // 2,
        occ_ceiling=occ_ceiling,
        round_rooms=round_rooms,
        growth_basis=growth_basis,
        growth_damping=growth_damping,
        growth_cap=growth_cap,
    )

    history_annual = hist.annual(daily, assumptions.capacity, fiscal_start_month, all_segments)
    history_monthly = hist.monthly(daily, assumptions.capacity, all_segments)

    complete_year_count = int(history_annual["Complete"].sum())
    if complete_year_count < 2:
        st.error(
            f"Only {complete_year_count} complete year(s) of history — growth rates need at "
            "least two. Change the year-start month, or upload a longer history."
        )
        return

    derived = growth.derive(
        history_annual, forecast_segments, growth_basis, growth_damping, growth_cap
    )

    reliability = segment_reliability(daily, all_segments, assumptions.smoothing_days)
    auto_smoothing = suggested_smoothing(reliability)
    if smoothing_mode == SMOOTH_GROUP_ONLY:
        # Group is the lumpiest segment by a wide margin; the rest keep their literal
        # date from last year.
        default_smoothing = {
            seg: (smoothing if seg == "Group" else 0.0) for seg in forecast_segments
        }
    elif smoothing_mode == "Auto (per segment)":
        default_smoothing = {seg: auto_smoothing[seg] for seg in forecast_segments}
    else:
        default_smoothing = {seg: smoothing for seg in forecast_segments}

    if smoothing_mode == SMOOTH_GROUP_ONLY and "Group" not in forecast_segments:
        st.warning(
            "Smoothing is set to Group only, but Group is not among the projected "
            "segments — nothing is being smoothed. Add Group, or pick another mode."
        )

    tab_inputs, tab_growth, tab_daily, tab_monthly, tab_segments, tab_history = st.tabs(
        ["Inputs & Summary", "Growth Rates", "Day by Day", "Monthly Summary", "Segments", "Historical"]
    )

    # --------------------------------------------------------------------------- #
    # Front sheet: fine-grained overrides, then the headline numbers
    # --------------------------------------------------------------------------- #

    with tab_inputs:
        st.subheader("Growth assumptions")
        st.caption(
            f"Per-segment rates are read from your history on the **{growth_basis}** basis "
            f"(reading {', '.join(derived.attrs['years_used'])}) and are editable. They multiply with "
            "the sidebar adjustment and any month override: a segment derived at +6% ADR "
            "with a +1% adjustment and a July override of +2% lands at +9.2% for July."
        )

        col_seg, col_month = st.columns([1.15, 1])

        with col_seg:
            st.markdown("**By segment** — derived from history, edit to override")
            seg_frame = pd.DataFrame(
                {
                    "Segment": derived["Segment"],
                    "ADR growth %": derived["ADR applied %"].round(2),
                    "Occupancy growth %": derived["Rooms applied %"].round(2),
                    "Smoothing": [default_smoothing[s] for s in derived["Segment"]],
                }
            )
            seg_edit = st.data_editor(
                seg_frame,
                hide_index=True,
                use_container_width=True,
                disabled=["Segment"],
                column_config={
                    "Smoothing": st.column_config.NumberColumn(
                        min_value=0.0, max_value=1.0, step=0.05, format="%.2f",
                        help="Rooms only; ADR is never smoothed. 0 copies last year's "
                        "day literally, 1 fully smooths it.",
                    )
                },
                # Re-deriving must reset the grid, so every input that feeds a default
                # is part of the key.
                key=f"seg_overrides|{growth_basis}|{growth_damping}|{growth_cap}"
                f"|{forecast_segments}|{smoothing_mode}|{smoothing}|{smoothing_days}",
            )
            capped = derived[derived["Rooms capped"] | derived["ADR capped"]]["Segment"].tolist()
            if capped:
                st.caption(
                    f"⚠️ Capped at ±{growth_cap:.0%}: {', '.join(capped)}. "
                    "See the Growth Rates tab for the raw rate."
                )

        with col_month:
            st.markdown("**By month** (% on top of segment)")
            month_frame = pd.DataFrame(
                {"Month": MONTHS, "ADR growth %": 0.0, "Occupancy growth %": 0.0}
            )
            month_edit = st.data_editor(
                month_frame,
                hide_index=True,
                use_container_width=True,
                disabled=["Month"],
                height=460,
                key="month_overrides",
            )

        assumptions.segment_adr_growth = {
            row["Segment"]: row["ADR growth %"] / 100 for _, row in seg_edit.iterrows()
        }
        assumptions.segment_occ_growth = {
            row["Segment"]: row["Occupancy growth %"] / 100 for _, row in seg_edit.iterrows()
        }
        assumptions.segment_smoothing = {
            row["Segment"]: float(row["Smoothing"]) for _, row in seg_edit.iterrows()
        }
        assumptions.month_adr_growth = {
            MONTHS.index(row["Month"]) + 1: row["ADR growth %"] / 100
            for _, row in month_edit.iterrows()
        }
        assumptions.month_occ_growth = {
            MONTHS.index(row["Month"]) + 1: row["Occupancy growth %"] / 100
            for _, row in month_edit.iterrows()
        }

    try:
        model = build_model(daily, assumptions)
        forecast = forecast_frame(model, assumptions)
    except ValueError as exc:
        st.error(str(exc))
        return

    if model.unanchored.any():
        st.warning(
            f"{int(model.unanchored.sum())} forecast day(s) have no matching history 364 days "
            "back and project as zero. Move the forecast period, or shorten it."
        )

    forecast_monthly = hist.monthly(forecast, assumptions.capacity, forecast_segments)
    versus_ly = compare_to_history(forecast, daily, assumptions.capacity, forecast_segments)

    # What each segment grows at once the sidebar adjustment is folded in. Month
    # overrides stack on top of these and are shown separately.
    applied = pd.DataFrame(
        [
            {
                "Segment": seg,
                "Derived ADR %": derived.loc[derived["Segment"] == seg, "ADR applied %"].iat[0],
                "Edited ADR %": assumptions.segment_adr_growth.get(seg, 0.0) * 100,
                "ADR incl. adjustment %": (
                    (1 + assumptions.segment_adr_growth.get(seg, 0.0))
                    * (1 + assumptions.adr_growth) - 1
                ) * 100,
                "Derived Occ %": derived.loc[derived["Segment"] == seg, "Rooms applied %"].iat[0],
                "Edited Occ %": assumptions.segment_occ_growth.get(seg, 0.0) * 100,
                "Occ incl. adjustment %": (
                    (1 + assumptions.segment_occ_growth.get(seg, 0.0))
                    * (1 + assumptions.occ_growth) - 1
                ) * 100,
            }
            for seg in forecast_segments
        ]
    )

    with tab_growth:
        st.subheader("What the history says")
        st.caption(
            f"Complete years on file: {', '.join(derived.attrs['years_available'])}. "
            "Part years are excluded — they would understate every rate."
        )

        for metric, label in (("Rooms", "Rooms sold"), ("ADR", "ADR")):
            st.markdown(f"**{label} — year-over-year %**")
            st.dataframe(
                growth.yearly_growth(history_annual, all_segments, metric).round(1),
                hide_index=True,
                use_container_width=True,
            )

        st.divider()
        st.subheader("Every basis, side by side")
        st.caption(
            "The basis is the single biggest lever on the budget — pick it deliberately. "
            f"Currently using **{growth_basis}**."
        )
        bases = growth.all_bases(history_annual, all_segments)
        st.dataframe(
            bases.drop(columns=["Key"]).round(2),
            hide_index=True,
            use_container_width=True,
            column_config={
                b: st.column_config.NumberColumn(format="%.2f%%", help=b) for b in growth.BASES
            },
        )

        basis_chart = bases.melt(
            id_vars=["Segment", "Metric"], value_vars=growth.BASES,
            var_name="Basis", value_name="Growth %",
        )
        st.altair_chart(
            alt.Chart(basis_chart)
            .mark_bar()
            .encode(
                x=alt.X("Basis:N", title=None, axis=alt.Axis(labels=False)),
                y=alt.Y("Growth %:Q"),
                color=alt.Color("Basis:N"),
                column=alt.Column("Segment:N", title=None),
                row=alt.Row("Metric:N", title=None),
                tooltip=["Segment", "Metric", "Basis", alt.Tooltip("Growth %:Q", format=".2f")],
            )
            .properties(width=90, height=140),
        )

        st.divider()
        st.subheader("Rates feeding the forecast")
        st.dataframe(
            applied.round(2),
            hide_index=True,
            use_container_width=True,
            column_config={
                c: st.column_config.NumberColumn(format="%.2f%%")
                for c in applied.columns
                if c != "Segment"
            },
        )
        st.caption(
            "**Derived** is what the history gave, after damping and capping. **Edited** is "
            "what is in the front-tab grid. **Incl. adjustment** folds in the sidebar "
            "adjustment. Month overrides multiply on top of these."
        )

        st.divider()
        st.subheader("How much of each segment is signal?")
        st.caption(
            "Strip out the local demand level and the day-of-week shape, and what is left "
            "is each date's own deviation. **Repeatability** is how strongly that deviation "
            "matches the same date 364 days earlier — the share of it that is real "
            "recurring business rather than noise. Auto smoothing is `1 − repeatability`, "
            "capped at 0.90: shrink a day toward the smooth estimate by however much of it "
            "is noise."
        )
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Segment": seg,
                        "Rooms / day": daily[f"{seg} Rooms"].mean(),
                        "Repeatability": reliability[seg],
                        "Auto smoothing": auto_smoothing[seg],
                        "In use": assumptions.smoothing_for(seg) if seg in forecast_segments else None,
                        "Projected": seg in forecast_segments,
                    }
                    for seg in all_segments
                ]
            ).round(3),
            hide_index=True,
            use_container_width=True,
        )
        st.caption(
            "A spiky segment is not automatically one to smooth — Group swings hardest but "
            "its swings tend to recur, because the same events come back each year."
        )

        active_months = {
            MONTHS[m - 1]: (a, assumptions.month_occ_growth.get(m, 0.0))
            for m, a in assumptions.month_adr_growth.items()
            if a or assumptions.month_occ_growth.get(m, 0.0)
        }
        if active_months:
            st.markdown("**Month overrides in effect**")
            st.dataframe(
                pd.DataFrame(
                    [
                        {"Month": k, "ADR %": v[0] * 100, "Occupancy %": v[1] * 100}
                        for k, v in active_months.items()
                    ]
                ).round(2),
                hide_index=True,
                use_container_width=True,
            )


    with tab_inputs:
        st.divider()
        st.subheader(f"Projected {assumptions.start:%d %b %Y} – {assumptions.end:%d %b %Y}")

        available = len(forecast) * assumptions.capacity
        current = _totals(forecast["Total Rooms"].sum(), forecast["Total Revenue"].sum(), available)
        prior = _totals(
            versus_ly["LY Rooms"].sum(),
            versus_ly["LY Revenue"].sum(),
            available,
        )
        _kpi_row(current, prior)
        st.caption("Comparison is against the same weekdays 364 days earlier.")

        chart_frame = versus_ly.melt(
            id_vars="Month", value_vars=["Rooms", "LY Rooms"], var_name="Series", value_name="Rooms sold"
        )
        st.altair_chart(
            alt.Chart(chart_frame)
            .mark_bar()
            .encode(
                x=alt.X("Month:N", title=None),
                y=alt.Y("Rooms sold:Q"),
                xOffset="Series:N",
                color=alt.Color("Series:N", title=None),
                tooltip=["Month", "Series", alt.Tooltip("Rooms sold:Q", format=",.0f")],
            )
            .properties(height=280),
            use_container_width=True,
        )

        st.divider()
        st.subheader("Download")
        workbook = excel_export.build_workbook(
            daily=daily,
            model=model,
            assumptions=assumptions,
            capacity=assumptions.capacity,
            fiscal_start_month=fiscal_start_month,
            all_segments=all_segments,
            notes=notes,
            derived=derived,
        )
        st.download_button(
            "Download budget workbook (.xlsx)",
            data=workbook,
            file_name=f"budget_{assumptions.start:%Y%m%d}_{assumptions.end:%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(
            "Includes assumptions, day-by-day forecast, monthly summary, forecast vs LY, "
            "a tab per segment, and the full historical roll-ups."
        )

    # --------------------------------------------------------------------------- #

    with tab_daily:
        st.subheader("Day-by-day projection")
        display = forecast.copy()
        display["Date"] = pd.to_datetime(display["Date"]).dt.date
        st.dataframe(
            display,
            hide_index=True,
            use_container_width=True,
            height=560,
            column_config={
                **{
                    f"{s} ADR": st.column_config.NumberColumn(format="$%.2f")
                    for s in forecast_segments
                },
                **{
                    f"{s} Revenue": st.column_config.NumberColumn(format="$%.0f")
                    for s in forecast_segments
                },
                "Total ADR": st.column_config.NumberColumn(format="$%.2f"),
                "Total Revenue": st.column_config.NumberColumn(format="$%.0f"),
                "RevPAR": st.column_config.NumberColumn(format="$%.2f"),
                "OCC %": st.column_config.NumberColumn(format="%.1f%%"),
            },
        )

        metric = st.radio("Chart", ["OCC %", "Total ADR", "Total Rooms", "RevPAR"], horizontal=True)
        st.altair_chart(
            alt.Chart(forecast)
            .mark_line(size=1.2)
            .encode(
                x=alt.X("Date:T", title=None),
                y=alt.Y(f"{metric}:Q", scale=alt.Scale(zero=False)),
                tooltip=["Date:T", "Day:N", alt.Tooltip(f"{metric}:Q", format=",.2f")],
            )
            .properties(height=300),
            use_container_width=True,
        )

    # --------------------------------------------------------------------------- #

    with tab_monthly:
        st.subheader("Monthly summary vs last year")
        st.dataframe(
            versus_ly,
            hide_index=True,
            use_container_width=True,
            column_config={
                "ADR": st.column_config.NumberColumn(format="$%.2f"),
                "LY ADR": st.column_config.NumberColumn(format="$%.2f"),
                "Revenue": st.column_config.NumberColumn(format="$%.0f"),
                "LY Revenue": st.column_config.NumberColumn(format="$%.0f"),
                "RevPAR": st.column_config.NumberColumn(format="$%.2f"),
                "OCC %": st.column_config.NumberColumn(format="%.1f%%"),
                "LY OCC %": st.column_config.NumberColumn(format="%.1f%%"),
                "Rooms": st.column_config.NumberColumn(format="%,d"),
                "LY Rooms": st.column_config.NumberColumn(format="%,d"),
            },
        )

        st.subheader("Full monthly detail")
        st.dataframe(forecast_monthly, hide_index=True, use_container_width=True)

    # --------------------------------------------------------------------------- #

    with tab_segments:
        st.subheader("Segment view")
        metric = st.selectbox("Metric", ["ADR", "OCC %", "Rooms", "Mix %", "Revenue"])

        combined_hist = hist.segment_matrix(history_monthly, all_segments, metric).assign(Period="Actual")
        combined_fcst = hist.segment_matrix(forecast_monthly, forecast_segments, metric).assign(
            Period="Forecast"
        )
        combined = pd.concat([combined_hist, combined_fcst], ignore_index=True)
        long = combined.melt(
            id_vars=["Month", "Period"], var_name="Segment", value_name=metric
        ).dropna(subset=[metric])
        long["Month"] = pd.PeriodIndex(long["Month"], freq="M").to_timestamp()

        st.altair_chart(
            alt.Chart(long)
            .mark_line(point=False)
            .encode(
                x=alt.X("Month:T", title=None),
                y=alt.Y(f"{metric}:Q", scale=alt.Scale(zero=(metric in ("Rooms", "Revenue", "Mix %")))),
                color=alt.Color("Segment:N"),
                strokeDash=alt.StrokeDash("Period:N", title=None),
                tooltip=["Month:T", "Segment:N", "Period:N", alt.Tooltip(f"{metric}:Q", format=",.2f")],
            )
            .properties(height=360),
            use_container_width=True,
        )

        st.caption("Solid = actual history, dashed = projection.")
        st.dataframe(combined, hide_index=True, use_container_width=True, height=400)

    # --------------------------------------------------------------------------- #

    with tab_history:
        st.subheader("Year over year")
        headline = [
            "Year", "Complete", "Days", "Total Rooms", "Rooms Available", "OCC %",
            "Total ADR", "RevPAR", "Total Revenue",
            "Total Rooms Y/Y %", "OCC pts Y/Y", "Total ADR Y/Y %", "Total Revenue Y/Y %",
        ]
        st.dataframe(
            history_annual[[c for c in headline if c in history_annual.columns]],
            hide_index=True,
            use_container_width=True,
            column_config={
                "Total ADR": st.column_config.NumberColumn(format="$%.2f"),
                "RevPAR": st.column_config.NumberColumn(format="$%.2f"),
                "Total Revenue": st.column_config.NumberColumn(format="$%.0f"),
                "OCC %": st.column_config.NumberColumn(format="%.1f%%"),
            },
        )
        if not history_annual["Complete"].all():
            st.caption("Years marked incomplete do not cover a full 12 months — read their Y/Y with care.")

        st.subheader("By segment")
        seg_metric = st.selectbox("Segment metric", ["Rooms", "ADR", "OCC %", "Mix %", "Revenue"], key="hist_metric")
        seg_cols = ["Year"] + [f"{s} {seg_metric}" for s in all_segments if f"{s} {seg_metric}" in history_annual.columns]
        st.dataframe(history_annual[seg_cols], hide_index=True, use_container_width=True)

        growth_cols = [c for c in history_annual.columns if c.endswith("Y/Y %") and any(
            c.startswith(f"{s} ") for s in all_segments)]
        st.subheader("Segment year-over-year growth %")
        st.dataframe(history_annual[["Year", *growth_cols]], hide_index=True, use_container_width=True)

        st.subheader("Occupancy mix by segment")
        # A partial year would sit in the chart at the same width as a full one.
        complete_years = history_annual[history_annual["Complete"]]
        mix = complete_years[["Year"] + [f"{s} Mix %" for s in all_segments]].melt(
            id_vars="Year", var_name="Segment", value_name="Mix %"
        )
        mix["Segment"] = mix["Segment"].str.replace(" Mix %", "", regex=False)
        st.altair_chart(
            alt.Chart(mix)
            .mark_bar()
            .encode(
                x=alt.X("Year:N", title=None),
                y=alt.Y("Mix %:Q", stack="normalize", axis=alt.Axis(format="%")),
                color=alt.Color("Segment:N"),
                tooltip=["Year", "Segment", alt.Tooltip("Mix %:Q", format=".1f")],
            )
            .properties(height=320),
            use_container_width=True,
        )

        st.subheader("Monthly history")
        st.dataframe(history_monthly, hide_index=True, use_container_width=True, height=400)
