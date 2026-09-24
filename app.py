import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formula.translate import Translator
import io
import csv
import re
import collections
import zipfile
import datetime
import hashlib
import json
import bcrypt
from pathlib import Path
from copy import copy, deepcopy
from xml.sax.saxutils import escape
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import math
import calendar

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
