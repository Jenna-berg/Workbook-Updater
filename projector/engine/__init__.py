"""Budget projection engine: load history, analyse it, project it forward."""

from .forecast import (
    Assumptions,
    ForecastModel,
    build_forecast,
    build_model,
    compare_to_history,
    default_horizon,
    expand,
    forecast_frame,
    public_columns,
    segment_reliability,
    snap_demand_window,
    suggested_smoothing,
)
from .loader import (
    DEFAULT_FORECAST_SEGMENTS,
    SEGMENTS,
    LoaderError,
    LoadResult,
    add_totals,
    infer_capacity,
    load_workbook,
    present_segments,
)

__all__ = [
    "Assumptions",
    "ForecastModel",
    "build_forecast",
    "build_model",
    "compare_to_history",
    "default_horizon",
    "expand",
    "forecast_frame",
    "public_columns",
    "segment_reliability",
    "snap_demand_window",
    "suggested_smoothing",
    "DEFAULT_FORECAST_SEGMENTS",
    "SEGMENTS",
    "LoaderError",
    "LoadResult",
    "add_totals",
    "infer_capacity",
    "load_workbook",
    "present_segments",
]
