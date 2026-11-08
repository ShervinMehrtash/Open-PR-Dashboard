# open_prs_dashboard.py
# ---------------------------------------------------------------------------
# Procurement Intelligence — open PRs · delivery risk · material criticality ·
# inventory receipts (SAP movement 101)
#
# Run:
#   pip install streamlit pandas numpy openpyxl jdatetime plotly
#   streamlit run open_prs_dashboard.py
# ---------------------------------------------------------------------------

from __future__ import annotations

import inspect
import io
import re
from datetime import date, datetime, time as dtime
from pathlib import Path
from types import SimpleNamespace

import jdatetime
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Procurement Intelligence", page_icon="◈", layout="wide")

DEFAULT_FILE = "PR from Bahman 1404.xlsx"
STOCK_FILE = "Stock.xlsx"
SHEET_NAME = "Open PRs"

BUCKETS = ["Not overdue", "1–7 days", "8–14 days", "15–30 days", "31–60 days", "61–90 days", ">90 days", "Date missing"]
BUCKET_COLORS = {
    "Not overdue": "#2EC4B6", "1–7 days": "#8BD17C", "8–14 days": "#F5D547", "15–30 days": "#F5A524",
    "31–60 days": "#F2813C", "61–90 days": "#EE5D4F", ">90 days": "#C7283D", "Date missing": "#7C8BA1",
}
SOURCE_COLORS = {"داخلی": "#5B8DEF", "خارجی (FPG)": "#F5A524", "نامشخص": "#7C8BA1", "مختلط": "#2EC4B6"}
STATUS_LABELS = {"CRITICAL": "🔴 Critical", "Overdue": "🟠 Overdue", "Normal": "🟢 On track"}
STATUS_COLORS = {"🔴 Critical": "#EE5D4F", "🟠 Overdue": "#F5A524", "🟢 On track": "#2EC4B6"}
ACCENT = ["#5B8DEF", "#2EC4B6", "#F5A524", "#9B7BFF", "#EE5D4F", "#8BD17C", "#F2813C", "#4CC9F0"]

# ---------------------------------------------------------------------------
# Look & feel
# ---------------------------------------------------------------------------
st.markdown("""
<style>
:root { color-scheme: dark; }
.stApp { background: radial-gradient(1200px 600px at 8% -5%, #1c2c47 0%, #111a2b 45%, #0a101c 100%); }
[data-testid="stHeader"] { background: rgba(10,16,28,.70); backdrop-filter: blur(8px); }
[data-testid="stSidebar"] { background: linear-gradient(180deg,#101b2e 0%,#0b1424 100%); border-right: 1px solid #22344f; }
.block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1650px; }
h1,h2,h3,h4 { letter-spacing: -.02em; color: #eaf1fb; }

.hero { padding: 22px 26px; border-radius: 20px; margin-bottom: 14px;
  background: linear-gradient(120deg, rgba(91,141,239,.22), rgba(46,196,182,.10) 55%, rgba(155,123,255,.14));
  border: 1px solid rgba(139,169,204,.22); }
.hero h1 { margin: 0; font-size: 2rem; font-weight: 780; background: linear-gradient(90deg,#eaf1fb,#8fb4ff);
  -webkit-background-clip: text; background-clip: text; color: transparent; }
.hero p { margin: 4px 0 0 0; color: #9fb2ca; font-size: .98rem; }

.pill { display:inline-block; padding: 4px 11px; margin: 8px 8px 0 0; border-radius: 999px; font-size: .78rem;
  color:#c9d9ee; background: rgba(91,141,239,.14); border: 1px solid rgba(91,141,239,.30); }
.pill.warn { background: rgba(245,165,36,.12); border-color: rgba(245,165,36,.35); color:#ffd596; }

.kpi { padding: 15px 17px; border-radius: 16px; border: 1px solid rgba(139,169,204,.18);
  background: linear-gradient(150deg, rgba(36,55,82,.80), rgba(18,29,46,.88)); box-shadow: 0 10px 28px rgba(0,0,0,.18);
  border-left: 4px solid #5B8DEF; min-height: 108px; }
.kpi-blue { border-left-color:#5B8DEF; } .kpi-teal { border-left-color:#2EC4B6; }
.kpi-amber { border-left-color:#F5A524; } .kpi-red { border-left-color:#EE5D4F; }
.kpi-violet { border-left-color:#9B7BFF; } .kpi-green { border-left-color:#8BD17C; }
.kpi-label { color:#9fb2ca; font-size:.78rem; text-transform: uppercase; letter-spacing:.06em; }
.kpi-value { color:#f3f7ff; font-size:1.75rem; font-weight:740; line-height:1.25; margin-top:2px; }
.kpi-sub { color:#8da3bf; font-size:.80rem; margin-top:2px; }

.callout { padding: 12px 15px; border-radius: 14px; margin-bottom: 10px; color:#d6e2f3; font-size:.93rem;
  background: rgba(91,141,239,.10); border: 1px solid rgba(91,141,239,.25); }
.callout.red { background: rgba(238,93,79,.10); border-color: rgba(238,93,79,.35); }
.callout.amber { background: rgba(245,165,36,.10); border-color: rgba(245,165,36,.35); }
.callout.green { background: rgba(46,196,182,.10); border-color: rgba(46,196,182,.32); }

.section-title { margin: 18px 0 2px 0; font-size: 1.12rem; font-weight: 700; color:#eaf1fb; }
.section-sub { margin: 0 0 8px 0; color:#8da3bf; font-size:.88rem; }

[data-testid="stTabs"] button { border-radius: 12px 12px 0 0; padding: 12px 16px; }
[data-testid="stDataFrame"], [data-testid="stPlotlyChart"] { border: 1px solid rgba(139,169,204,.16);
  border-radius: 16px; overflow: hidden; background: rgba(18,29,46,.55); }
.stButton button, .stDownloadButton button { border-radius: 11px; border: 1px solid #45658d;
  background: linear-gradient(135deg,#294b73,#1c3554); color: #f4f8ff; }
.stButton button:hover, .stDownloadButton button:hover { border-color: #8bb5e6; color: #fff; }
div[data-testid="stExpander"] { border: 1px solid rgba(139,169,204,.18); border-radius: 14px; }
hr { border-color: rgba(139,169,204,.18); }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Streamlit-version-proof helpers
# ---------------------------------------------------------------------------
def _stretch_kwargs(fn):
    """`width="stretch"` on new Streamlit, `use_container_width=True` on older versions."""
    try:
        p = inspect.signature(fn).parameters.get("width")
        if p is not None and isinstance(p.default, str):
            return {"width": "stretch"}
    except (TypeError, ValueError):
        pass
    return {"use_container_width": True}


_DF_STRETCH = _stretch_kwargs(st.dataframe)
_CHART_STRETCH = _stretch_kwargs(st.plotly_chart)


def show_df(df, column_config=None, height=None):
    kwargs = dict(hide_index=True, **_DF_STRETCH)
    if column_config:
        kwargs["column_config"] = column_config
    if height:
        kwargs["height"] = height
    st.dataframe(df, **kwargs)


def show_chart(fig):
    st.plotly_chart(fig, **_CHART_STRETCH)


def style_fig(fig, height=360, legend=True):
    fig.update_layout(
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#c9d6e8"), height=height, margin=dict(l=12, r=12, t=52, b=12),
        title=dict(font=dict(size=15, color="#eaf1fb")), showlegend=legend,
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )
    fig.update_xaxes(gridcolor="rgba(139,169,204,.12)", zeroline=False)
    fig.update_yaxes(gridcolor="rgba(139,169,204,.12)", zeroline=False)
    return fig


def kpi_row(cards):
    cols = st.columns(len(cards))
    for col, (label, value, sub, tone) in zip(cols, cards):
        col.markdown(
            f'<div class="kpi kpi-{tone}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>',
            unsafe_allow_html=True,
        )


def callout(text, tone=""):
    st.markdown(f'<div class="callout {tone}">{text}</div>', unsafe_allow_html=True)


def section(title, sub=""):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)
    if sub:
        st.markdown(f'<div class="section-sub">{sub}</div>', unsafe_allow_html=True)


def pills(items):
    html = "".join(f'<span class="pill{" warn" if warn else ""}">{txt}</span>' for txt, warn in items)
    st.markdown(html, unsafe_allow_html=True)


def pct_config(*names):
    cfg = {}
    for n in names:
        cfg[n] = st.column_config.ProgressColumn(n, min_value=0, max_value=100, format="%.0f%%")
    return cfg


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------
PERSIAN_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def normalize_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).translate(PERSIAN_DIGITS).strip()


def normalize_code(value):
    """Material / plant / movement-type codes: text, no stray '.0' from Excel numbers."""
    return re.sub(r"\.0+$", "", normalize_text(value))


def normalize_column_name(value):
    return re.sub(r"\s+", " ", normalize_text(value)).strip()


def first_matching_column(columns, candidates):
    lookup = {normalize_column_name(c).casefold(): c for c in columns}
    for candidate in candidates:
        key = normalize_column_name(candidate).casefold()
        if key in lookup:
            return lookup[key]
    return None


def parse_jalali_date(value):
    """Jalali or Gregorian date → Timestamp. Accepts Excel dates/serials, Persian digits,
    dd/mm/yyyy, yyyy/mm/dd, yyyymmdd. Years < 1700 are treated as Jalali."""
    if value is None:
        return pd.NaT
    try:
        if pd.isna(value):
            return pd.NaT
    except (TypeError, ValueError):
        pass
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value).normalize()

    s = re.sub(r"\s+", " ", normalize_text(value)).strip()
    if not s:
        return pd.NaT

    # yyyymmdd (Jalali or Gregorian), e.g. 14050817 / 20261108
    if re.fullmatch(r"\d{8}", s):
        return _make_date(int(s[:4]), int(s[4:6]), int(s[6:]))

    # Excel serial number (e.g. 45678 or 45678.0)
    if re.fullmatch(r"\d+(\.\d+)?", s):
        n = float(s)
        if n > 20000:
            return (pd.Timestamp("1899-12-30") + pd.to_timedelta(n, unit="D")).normalize()
        return pd.NaT

    s = re.split(r"[ T]", s)[0].replace("-", "/").replace(".", "/")
    parts = s.split("/")
    if len(parts) != 3:
        return pd.NaT
    try:
        a, b, c = (int(p) for p in parts)
    except ValueError:
        return pd.NaT
    if a > 31:                       # yyyy/mm/dd
        return _make_date(a, b, c)
    return _make_date(c, b, a)       # dd/mm/yyyy


def _make_date(year, month, day):
    try:
        if year < 1700:
            return pd.Timestamp(jdatetime.date(year, month, day).togregorian()).normalize()
        return pd.Timestamp(year=year, month=month, day=day).normalize()
    except Exception:
        return pd.NaT


def parse_clock(value):
    """Excel time / datetime / fraction-of-day / 'HH:MM[:SS]' / SAP 'HHMMSS' → datetime.time or None."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.time()
    if isinstance(value, dtime):
        return value
    if isinstance(value, (int, float, np.integer, np.floating)):
        v = float(value)
        if 0 <= v < 1:
            secs = int(round(v * 86400)) % 86400
            return dtime(secs // 3600, (secs % 3600) // 60, secs % 60)
        value = str(int(v))
    s = normalize_text(value)
    if not s:
        return None
    if re.fullmatch(r"\d{5,6}", s):
        s = s.zfill(6)
        try:
            return dtime(int(s[:2]), int(s[2:4]), int(s[4:]))
        except ValueError:
            return None
    try:
        return pd.to_datetime(s).time()
    except Exception:
        return None


def to_jalali_str(value):
    if value is None or pd.isna(value):
        return ""
    try:
        return jdatetime.date.fromgregorian(date=pd.Timestamp(value).date()).strftime("%Y/%m/%d")
    except Exception:
        return ""


def jalali_series(series):
    return series.map(to_jalali_str)


def jalali_datetime_series(series):
    def f(v):
        if pd.isna(v):
            return ""
        return f"{to_jalali_str(v)} {pd.Timestamp(v):%H:%M}"
    return series.map(f)


def fmt_num(x):
    if x is None or pd.isna(x):
        return "—"
    x = float(x)
    if abs(x) >= 1000:
        return f"{x:,.0f}"
    return f"{x:,.2f}".rstrip("0").rstrip(".")


# ---------------------------------------------------------------------------
# File loading (cached on file bytes)
# ---------------------------------------------------------------------------
class PRDataError(Exception):
    """User-facing problem with the PR workbook."""


def read_source(uploaded, path):
    if uploaded is not None:
        return uploaded.getvalue()
    p = Path(str(path).strip().strip('"'))
    if not p.is_file():
        raise FileNotFoundError(str(p))
    return p.read_bytes()


@st.cache_data(show_spinner="Reading PR workbook…")
def load_pr_workbook(data: bytes):
    xls = pd.ExcelFile(io.BytesIO(data))
    sheet = SHEET_NAME if SHEET_NAME in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
    df.columns = [normalize_column_name(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    return df, sheet


@st.cache_data(show_spinner="Reading stock file…")
def load_stock(data: bytes):
    xls = pd.ExcelFile(io.BytesIO(data))
    sheet = "Stock" if "Stock" in xls.sheet_names else xls.sheet_names[0]
    sdf = pd.read_excel(xls, sheet_name=sheet, dtype=object)
    sdf.columns = [normalize_column_name(c) for c in sdf.columns]
    mat_col = first_matching_column(sdf.columns, ["Material"])
    qty_col = first_matching_column(sdf.columns, ["Unrestricted", "Unrestricted Stock", "Unrestricted stock"])
    if mat_col is None or qty_col is None:
        raise ValueError(f"Stock file needs 'Material' and 'Unrestricted' columns. Found: {list(sdf.columns)}")
    sdf = sdf[[mat_col, qty_col]].copy()
    sdf.columns = ["Material", "Unrestricted"]
    sdf["Material"] = sdf["Material"].map(normalize_code)
    sdf["Unrestricted"] = pd.to_numeric(
        sdf["Unrestricted"].map(normalize_text).str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0)
    return sdf[sdf["Material"].ne("")].groupby("Material", as_index=False)["Unrestricted"].sum()


MOVEMENT_REQUIRED = ["Movement type", "Material", "Plant", "Posting Date",
                     "Qty in unit of entry", "Entry Date", "Time of Entry"]


@st.cache_data(show_spinner="Reading movement workbook…")
def load_movements(data: bytes):
    xls = pd.ExcelFile(io.BytesIO(data))
    m = pd.read_excel(xls, sheet_name=xls.sheet_names[0], dtype=object)
    m.columns = [normalize_column_name(c) for c in m.columns]
    resolved = {k: first_matching_column(m.columns, [k]) for k in MOVEMENT_REQUIRED}
    missing = [k for k, v in resolved.items() if v is None]
    if missing:
        raise ValueError("Missing movement columns: " + ", ".join(missing) + f". Found: {list(m.columns)}")
    m = m[[resolved[k] for k in MOVEMENT_REQUIRED]].copy()
    m.columns = MOVEMENT_REQUIRED
    m["Movement type"] = m["Movement type"].map(normalize_code)
    m["Material"] = m["Material"].map(normalize_code)
    m["Plant"] = m["Plant"].map(normalize_code)
    m["Qty in unit of entry"] = pd.to_numeric(
        m["Qty in unit of entry"].map(normalize_text).str.replace(",", "", regex=False), errors="coerce"
    ).fillna(0)

    dates = pd.to_datetime(m["Entry Date"].map(parse_jalali_date), errors="coerce")
    clocks = m["Time of Entry"].map(parse_clock)
    offsets = clocks.map(
        lambda t: pd.Timedelta(hours=t.hour, minutes=t.minute, seconds=t.second) if t is not None else pd.Timedelta(0)
    )
    ts = dates + offsets
    posting = pd.to_datetime(m["Posting Date"].map(parse_jalali_date), errors="coerce")
    ts = ts.where(ts.notna(), posting)      # fallback: posting date at 00:00
    m["Entry timestamp"] = ts
    return m


# ---------------------------------------------------------------------------
# PR pipeline
# ---------------------------------------------------------------------------
def build_pr_table(raw, ignore_nearly_completed, done_threshold):
    """Clean the raw PR sheet → one row per open PR line. Returns (df, stats)."""
    stats = {"rows_read": len(raw), "flag_removed": 0, "nearly_done_ignored": 0, "warnings": []}

    # 1) Closed / Deletion indicator = "X"  → remove
    for names in (["Closed", "Closed indicator", "Closed ind."],
                  ["Deletion indicator", "Deletion Indicator", "Deletion ind.", "Deletion ind"]):
        col = first_matching_column(raw.columns, names)
        if col is not None:
            flagged = raw[col].map(normalize_text).str.upper().eq("X")
            stats["flag_removed"] += int(flagged.sum())
            raw = raw.loc[~flagged]

    col_map = {
        "material": first_matching_column(raw.columns, ["Material"]),
        "short_text": first_matching_column(raw.columns, ["Short Text", "Short text", "ShortText", "Description"]),
        "purchasing_group": first_matching_column(raw.columns, ["Purchasing group", "Purchasing Group"]),
        "delivery_date": first_matching_column(raw.columns, ["Delivery date", "Delivery Date"]),
        "requisition_date": first_matching_column(raw.columns, ["Requisition date", "Requisition Date"]),
        "requested": first_matching_column(raw.columns, ["Quantity requested", "Quantity Requested"]),
        "remaining": first_matching_column(raw.columns, [
            "Still to be delivered(P.O)", "Still to be delivered (P.O)", "Still to be delivered",
            "Still to be delivered(P.O.)"]),
        "processing_status": first_matching_column(raw.columns, ["Processing status", "Processing Status"]),
    }
    optional = ("short_text", "purchasing_group", "processing_status")
    missing = [k for k, v in col_map.items() if v is None and k not in optional]
    if missing:
        raise PRDataError("These required columns were not found: " + ", ".join(missing)
                          + ". Columns detected: " + ", ".join(map(str, raw.columns)))
    if col_map["short_text"] is None:
        stats["warnings"].append("'Short Text' not found — descriptions are empty.")
    if col_map["purchasing_group"] is None:
        stats["warnings"].append("'Purchasing group' not found — domestic/foreign split unavailable.")
    if col_map["processing_status"] is None:
        stats["warnings"].append("'Processing status' not found — status filter unavailable.")

    raw = raw.reset_index(drop=True)
    df = pd.DataFrame(index=raw.index)
    df["Material"] = raw[col_map["material"]].map(normalize_code)
    df["Short Text"] = raw[col_map["short_text"]].map(normalize_text) if col_map["short_text"] else ""
    if col_map["processing_status"]:
        ps = raw[col_map["processing_status"]].map(normalize_text)
        df["Processing status"] = ps.where(ps.ne(""), "(blank)")
    else:
        df["Processing status"] = "—"

    if col_map["purchasing_group"]:
        pg = raw[col_map["purchasing_group"]].map(normalize_text).str.upper()
        df["Purchasing group"] = pg
        df["Source type"] = np.where(pg.eq(""), "نامشخص", np.where(pg.eq("FPG"), "خارجی (FPG)", "داخلی"))
    else:
        df["Purchasing group"] = ""
        df["Source type"] = "نامشخص"

    df["Delivery date"] = pd.to_datetime(raw[col_map["delivery_date"]].map(parse_jalali_date), errors="coerce")
    df["Requisition date"] = pd.to_datetime(raw[col_map["requisition_date"]].map(parse_jalali_date), errors="coerce")
    for out_col, key in (("Quantity requested", "requested"), ("Still to be delivered", "remaining")):
        df[out_col] = pd.to_numeric(
            raw[col_map[key]].map(normalize_text).str.replace(",", "", regex=False), errors="coerce"
        )
    df = df[df["Material"].ne("")].copy()
    df["Quantity requested"] = df["Quantity requested"].fillna(0)
    df["Still to be delivered"] = df["Still to be delivered"].fillna(0).clip(lower=0)

    # 2) Open = still something to deliver
    df = df[df["Still to be delivered"] > 0].copy()

    # 3) Optional: ignore lines that are ≥ threshold % delivered
    if ignore_nearly_completed:
        req = df["Quantity requested"].astype(float)
        remaining_ratio = df["Still to be delivered"].astype(float) / req.where(req > 0)
        nearly_done = (remaining_ratio <= (1 - done_threshold / 100.0) + 1e-9).fillna(False)
        stats["nearly_done_ignored"] = int(nearly_done.sum())
        df = df[~nearly_done].copy()

    return df.reset_index(drop=True), stats


def add_derived(df, as_of, due_window):
    df = df.copy()
    req = df["Quantity requested"].astype(float)
    rem = df["Still to be delivered"].astype(float)
    df["Delivered quantity (estimated)"] = (req - rem).clip(lower=0)
    df["Delivery completion %"] = (df["Delivered quantity (estimated)"] / req.where(req > 0) * 100).fillna(0).clip(0, 100)
    df["Remaining %"] = (rem / req.where(req > 0) * 100).fillna(0).clip(0, 100)

    analysis_date = pd.Timestamp(as_of).normalize()
    has_date = df["Delivery date"].notna()
    df["Days past delivery"] = (analysis_date - df["Delivery date"]).dt.days.astype("float64")
    df["Days until delivery"] = -df["Days past delivery"]
    df["Overdue"] = has_date & (df["Days past delivery"] > 0)
    df["Overdue remaining"] = df["Still to be delivered"].where(df["Overdue"], 0.0)

    cut = pd.cut(df["Days past delivery"], bins=[-np.inf, 0, 7, 14, 30, 60, 90, np.inf], labels=BUCKETS[:7])
    df["Delay bucket"] = cut.astype(object).where(has_date, "Date missing")
    df["Delay bucket"] = df["Delay bucket"].fillna("Date missing")

    df["Due status"] = np.select(
        [~has_date, df["Overdue"], df["Days until delivery"] <= due_window],
        ["No date", "Overdue", f"Due ≤{due_window} d"], default="Later",
    )
    df["Upcoming delivery date"] = df["Delivery date"].where(df["Days until delivery"] >= 0)
    return df


def build_material_summary(filtered, stock_df, severe_delay_days, low_delivery_pct, min_critical_remaining):
    def first_text(s):
        return next((x for x in s if x), "")

    ms = filtered.groupby("Material", sort=False).agg(**{
        "Description": ("Short Text", first_text),
        "Source": ("Source type", "first"),
        "_n_src": ("Source type", "nunique"),
        "Open PRs": ("Material", "size"),
        "Total requested": ("Quantity requested", "sum"),
        "Estimated delivered": ("Delivered quantity (estimated)", "sum"),
        "Remaining quantity": ("Still to be delivered", "sum"),
        "Overdue PRs": ("Overdue", "sum"),
        "Overdue remaining": ("Overdue remaining", "sum"),
        "Oldest delivery date": ("Delivery date", "min"),
        "Latest delivery date": ("Delivery date", "max"),
        "Next due date": ("Upcoming delivery date", "min"),
        "Worst delay (days)": ("Days past delivery", "max"),
        "Earliest requisition": ("Requisition date", "min"),
    }).reset_index()
    ms.loc[ms["_n_src"] > 1, "Source"] = "مختلط"
    ms = ms.drop(columns="_n_src")

    total = ms["Total requested"].astype(float)
    ms["Completion %"] = (ms["Estimated delivered"] / total.where(total > 0) * 100).fillna(0).clip(0, 100)
    ms["Remaining %"] = (ms["Remaining quantity"] / total.where(total > 0) * 100).fillna(0).clip(0, 100)
    ms["Worst delay (days)"] = ms["Worst delay (days)"].fillna(0).clip(lower=0)
    ms["Overdue PRs"] = ms["Overdue PRs"].astype(int)

    if stock_df is not None:
        ms = ms.merge(stock_df.rename(columns={"Unrestricted": "Unrestricted stock"}), on="Material", how="left")
        ms["Unrestricted stock"] = ms["Unrestricted stock"].fillna(0)

    critical = (ms["Remaining quantity"] > min_critical_remaining) & (
        (ms["Worst delay (days)"] >= severe_delay_days)
        | ((ms["Completion %"] <= low_delivery_pct) & (ms["Overdue PRs"] > 0))
    )
    ms["Criticality"] = np.where(critical, "CRITICAL", np.where(ms["Overdue PRs"] > 0, "Overdue", "Normal"))
    order = {"CRITICAL": 0, "Overdue": 1, "Normal": 2}
    ms = (ms.assign(_o=ms["Criticality"].map(order))
            .sort_values(["_o", "Worst delay (days)", "Remaining quantity"], ascending=[True, False, False])
            .drop(columns="_o").reset_index(drop=True))
    ms["Status"] = ms["Criticality"].map(STATUS_LABELS)
    return ms


def with_jalali(df, cols):
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = jalali_series(out[c])
    return out


def search_mask(frame, text, columns):
    text = normalize_text(text).casefold()
    if not text:
        return pd.Series(True, index=frame.index)
    mask = pd.Series(False, index=frame.index)
    for c in columns:
        if c in frame.columns:
            mask |= frame[c].astype(str).str.casefold().str.contains(text, regex=False, na=False)
    return mask


def csv_button(label, frame, name):
    st.download_button(label, data=frame.to_csv(index=False).encode("utf-8-sig"), file_name=name, mime="text/csv")


def short_label(material, desc, n=28):
    desc = (desc or "")
    return f"{material} · {desc[:n]}{'…' if len(desc) > n else ''}" if desc else str(material)


# ---------------------------------------------------------------------------
# Sidebar: data + rules
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### ◈ Control panel")
    with st.expander("📁 Data files", expanded=True):
        uploaded_file = st.file_uploader("PR workbook (Excel)", type=["xlsx", "xls"], key="pr_upload")
        file_path = st.text_input("…or PR file path", value=DEFAULT_FILE)
        uploaded_stock = st.file_uploader("Stock file (optional)", type=["xlsx", "xls"], key="stock_upload")
        stock_path = st.text_input("…or Stock file path", value=STOCK_FILE)
        uploaded_movements = st.file_uploader(
            "SAP movement file", type=["xlsx", "xls"], key="movement_upload",
            help="Columns: Movement type, Material, Plant, Posting Date, Qty in unit of entry, Entered on, Entered at.",
        )

    with st.expander("⚙️ Analysis rules", expanded=True):
        as_of = st.date_input("Analysis date", value=date.today(), format="DD/MM/YYYY")
        ignore_nearly_completed = st.toggle(
            "Ignore PRs that are almost fully delivered", value=True,
            help="ON: open PR lines whose delivered share is at or above the threshold are removed from all PR analysis.",
        )
        done_threshold = st.slider("…delivered share threshold (%)", 50, 100, 90, disabled=not ignore_nearly_completed)
        due_window = st.number_input("“Due soon” window (days)", min_value=1, max_value=120, value=10, step=1)
        severe_delay_days = st.number_input("Critical delay threshold (days)", min_value=1, max_value=730, value=30)
        low_delivery_pct = st.slider("Low delivery threshold (%)", 0, 100, 20)
        min_critical_remaining = st.number_input("Min. remaining qty for critical flag", min_value=0.0, value=0.0, step=1.0)
        consider_stock = st.toggle(
            "Exclude materials with Unrestricted stock > 0", value=False,
            help="Materials that already have positive Unrestricted stock in the Stock file are removed from the PR analysis.",
        )
        st.caption("تاریخ‌های متنی dd/mm/yyyy شمسی و تاریخ‌های واقعی Excel میلادی در نظر گرفته می‌شوند.")


# ---------------------------------------------------------------------------
# Hero
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="hero"><h1>◈ Procurement Intelligence</h1>'
    '<p>Open PRs · delivery risk · material criticality · inventory receipts</p></div>',
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Load + build PR data (non-fatal: the movement tab still works without it)
# ---------------------------------------------------------------------------
pr_problem = None
df = pd.DataFrame()
stats = {}
stock_df = None
sheet_used = ""

try:
    raw, sheet_used = load_pr_workbook(read_source(uploaded_file, file_path))
    df, stats = build_pr_table(raw, ignore_nearly_completed, done_threshold)
    if df.empty:
        pr_problem = "No open PR rows found: 'Still to be delivered' is zero/missing, or every line was filtered out by the rules."
except FileNotFoundError:
    pr_problem = "PR workbook not found. Upload it in the sidebar or enter a correct file path. (فایل Excel را از نوار کناری آپلود کنید.)"
except PRDataError as e:
    pr_problem = str(e)
except Exception as e:
    pr_problem = f"Could not read the PR workbook: {e}"

stock_note = None
try:
    stock_df = load_stock(read_source(uploaded_stock, stock_path))
except FileNotFoundError:
    if consider_stock:
        stock_note = f"Stock file not found at '{stock_path}'. Continuing without the stock rule."
except Exception as e:
    stock_note = f"Could not read Stock file: {e}"

stock_excluded_lines = 0
stock_excluded_materials = 0
if pr_problem is None and consider_stock and stock_df is not None:
    covered = set(stock_df.loc[stock_df["Unrestricted"] > 0, "Material"])
    stock_excluded_lines = int(df["Material"].isin(covered).sum())
    stock_excluded_materials = len(covered.intersection(df["Material"].unique()))
    df = df[~df["Material"].isin(covered)].copy()
    if df.empty:
        pr_problem = "After applying the stock rule, no open PR rows remain."

pr_ready = pr_problem is None
filtered = pd.DataFrame()
base = pd.DataFrame()

if pr_ready:
    df = add_derived(df, as_of, int(due_window))

    # ---------------- Filters (sidebar) ----------------
    with st.sidebar:
        with st.expander("🔎 PR filters", expanded=True):
            present_sources = [s for s in ["داخلی", "خارجی (FPG)", "نامشخص"] if s in set(df["Source type"])]
            selected_sources = st.multiselect("Source (domestic / foreign)", present_sources, default=present_sources,
                                              help="FPG = foreign procurement (بازرگانی خارجی).")
            statuses = sorted(df["Processing status"].astype(str).unique().tolist())
            selected_statuses = st.multiselect("Processing status", statuses, default=statuses)
            selected_materials = st.multiselect("Material (empty = all)", sorted(df["Material"].unique().tolist()), default=[])
            selected_buckets = st.multiselect("Delay bucket", BUCKETS, default=BUCKETS)
            overdue_only = st.toggle("Overdue only", value=False)
            use_req_date = st.toggle("Filter by requisition date", value=False)
            req_min, req_max = df["Requisition date"].min(), df["Requisition date"].max()
            req_range = None
            if pd.notna(req_min) and pd.notna(req_max):
                req_range = st.date_input(
                    "Requisition date range", value=(req_min.date(), req_max.date()),
                    min_value=req_min.date(), max_value=req_max.date(),
                    format="DD/MM/YYYY", disabled=not use_req_date,
                )

    base = df[df["Source type"].isin(selected_sources) & df["Processing status"].isin(selected_statuses)]
    if selected_materials:
        base = base[base["Material"].isin(selected_materials)]
    if use_req_date and isinstance(req_range, (tuple, list)) and len(req_range) == 2:
        base = base[base["Requisition date"].between(pd.Timestamp(req_range[0]), pd.Timestamp(req_range[1]), inclusive="both")]

    filtered = base[base["Delay bucket"].isin(selected_buckets)]
    if overdue_only:
        filtered = filtered[filtered["Overdue"]]

# ---------------------------------------------------------------------------
# Status pills
# ---------------------------------------------------------------------------
if pr_ready:
    items = [(f"PR sheet “{sheet_used}” · {stats['rows_read']:,} rows read", False),
             (f"{stats['flag_removed']:,} closed / deleted rows removed", False)]
    if ignore_nearly_completed:
        items.append((f"{stats['nearly_done_ignored']:,} lines ≥{done_threshold}% delivered ignored", False))
    if consider_stock and stock_df is not None:
        items.append((f"{stock_excluded_lines:,} lines excluded by stock ({stock_excluded_materials:,} materials)", False))
    items.append((f"Analysis date {to_jalali_str(pd.Timestamp(as_of))} (شمسی)", False))
    for w in stats.get("warnings", []):
        items.append((w, True))
    if stock_note:
        items.append((stock_note, True))
    pills(items)

if pr_problem:
    st.warning(pr_problem)


# ---------------------------------------------------------------------------
# PR panels
# ---------------------------------------------------------------------------
def build_insights(f, ms):
    notes = []
    n = len(f)
    od = f[f["Overdue"]]
    rem_total = f["Still to be delivered"].sum()
    if n:
        share = len(od) / n * 100
        rem_share = od["Still to be delivered"].sum() / rem_total * 100 if rem_total else 0
        tone = "red" if share >= 50 else ("amber" if share >= 25 else "green")
        notes.append((f"<b>{share:.0f}% of open PR lines are overdue</b> ({len(od):,} of {n:,}), holding "
                      f"{fmt_num(od['Still to be delivered'].sum())} units — {rem_share:.0f}% of everything still to be delivered.", tone))
    if len(od):
        worst = ms.sort_values("Worst delay (days)", ascending=False).iloc[0]
        notes.append((f"<b>Longest delay:</b> {short_label(worst['Material'], worst['Description'], 40)} — "
                      f"{worst['Worst delay (days)']:.0f} days past delivery date.", "red"))
        top = ms.sort_values("Overdue remaining", ascending=False).iloc[0]
        notes.append((f"<b>Biggest overdue quantity:</b> {short_label(top['Material'], top['Description'], 40)} — "
                      f"{fmt_num(top['Overdue remaining'])} units still missing.", "amber"))
    soon = f[f["Days until delivery"].between(0, int(due_window), inclusive="both")]
    if len(soon):
        notes.append((f"<b>{len(soon):,} PR lines ({soon['Material'].nunique():,} materials)</b> are due within the next "
                      f"{int(due_window)} days — {fmt_num(soon['Still to be delivered'].sum())} units.", ""))
    src = f.groupby("Source type").agg(rem=("Still to be delivered", "sum"), od=("Overdue remaining", "sum"))
    src = src[src["rem"] > 0]
    if len(src) > 1:
        src["rate"] = src["od"] / src["rem"] * 100
        w = src["rate"].idxmax()
        notes.append((f"<b>Highest overdue rate by source:</b> {w} — {src.loc[w, 'rate']:.0f}% of its remaining quantity is overdue.", ""))
    return notes


def render_overview(f, ms):
    notes = build_insights(f, ms)
    if notes:
        section("Key insights", "Generated automatically from the current filters")
        cols = st.columns(2)
        for i, (txt, tone) in enumerate(notes):
            with cols[i % 2]:
                callout(txt, tone)

    bucket_summary = (
        f.groupby("Delay bucket")
        .agg(**{"PR lines": ("Material", "size"), "Materials": ("Material", "nunique"),
                "Remaining quantity": ("Still to be delivered", "sum")})
        .reindex(BUCKETS, fill_value=0).reset_index()
    )
    c1, c2 = st.columns([3, 2])
    with c1:
        fig = px.bar(bucket_summary, x="Delay bucket", y="Remaining quantity", color="Delay bucket",
                     color_discrete_map=BUCKET_COLORS, category_orders={"Delay bucket": BUCKETS},
                     hover_data=["PR lines", "Materials"], title="How late is the remaining quantity?")
        fig.update_layout(xaxis_title=None, yaxis_title="Remaining quantity")
        show_chart(style_fig(fig, 380, legend=False))
    with c2:
        status = f.groupby("Due status").agg(lines=("Material", "size")).reset_index()
        cmap = {"Overdue": "#EE5D4F", f"Due ≤{int(due_window)} d": "#F5A524", "Later": "#2EC4B6", "No date": "#7C8BA1"}
        fig = px.pie(status, names="Due status", values="lines", hole=0.58, color="Due status",
                     color_discrete_map=cmap, title="PR lines by delivery status")
        fig.update_traces(textinfo="percent+label")
        show_chart(style_fig(fig, 380, legend=False))

    c3, c4 = st.columns(2)
    with c3:
        top = ms[ms["Overdue remaining"] > 0].nlargest(10, "Overdue remaining").copy()
        if top.empty:
            callout("No overdue quantities under the current filters.", "green")
        else:
            top["Label"] = [short_label(m, d) for m, d in zip(top["Material"], top["Description"])]
            fig = px.bar(top.sort_values("Overdue remaining"), x="Overdue remaining", y="Label", orientation="h",
                         color="Worst delay (days)", color_continuous_scale=["#F5D547", "#F2813C", "#C7283D"],
                         title="Top 10 materials by overdue quantity (color = worst delay, days)")
            fig.update_layout(yaxis_title=None, xaxis_title="Overdue remaining quantity")
            show_chart(style_fig(fig, 420))
    with c4:
        horizon = f[f["Delivery date"].notna()].copy()
        if horizon.empty:
            callout("No delivery dates available for the horizon chart.", "amber")
        else:
            d = horizon["Days until delivery"]
            horizon["Horizon"] = np.where(horizon["Overdue"], "Overdue", np.where(d > 84, "> 12 weeks", ""))
            wk = (d // 7).clip(lower=0).fillna(0).astype(int)
            lab = wk.map(lambda w: f"{w * 7}–{w * 7 + 6} d")
            horizon["Horizon"] = horizon["Horizon"].where(horizon["Horizon"].ne(""), lab)
            order = ["Overdue"] + [f"{w * 7}–{w * 7 + 6} d" for w in range(12)] + ["> 12 weeks"]
            hz = horizon.groupby("Horizon")["Still to be delivered"].sum().reindex(order).dropna().reset_index()
            hz["Color"] = np.where(hz["Horizon"] == "Overdue", "Overdue", "Upcoming")
            fig = px.bar(hz, x="Horizon", y="Still to be delivered", color="Color",
                         color_discrete_map={"Overdue": "#EE5D4F", "Upcoming": "#5B8DEF"},
                         category_orders={"Horizon": order}, title="Delivery horizon — what arrives when (days from analysis date)")
            fig.update_layout(xaxis_title=None, yaxis_title="Remaining quantity")
            show_chart(style_fig(fig, 420, legend=False))

    c5, c6 = st.columns(2)
    with c5:
        sb = f.groupby(["Source type", "Delay bucket"])["Still to be delivered"].sum().reset_index()
        fig = px.bar(sb, x="Source type", y="Still to be delivered", color="Delay bucket", color_discrete_map=BUCKET_COLORS,
                     category_orders={"Delay bucket": BUCKETS}, title="Domestic vs foreign — remaining quantity by delay")
        fig.update_layout(xaxis_title=None, yaxis_title="Remaining quantity")
        show_chart(style_fig(fig, 380))
    with c6:
        ps = (f.groupby("Processing status")
               .agg(**{"PR lines": ("Material", "size"), "Remaining quantity": ("Still to be delivered", "sum"),
                       "Overdue remaining": ("Overdue remaining", "sum")}).reset_index()
               .sort_values("Remaining quantity"))
        fig = px.bar(ps, y="Processing status", x=["Overdue remaining", "Remaining quantity"], orientation="h",
                     barmode="overlay", title="Processing status — remaining vs overdue quantity",
                     color_discrete_sequence=["#EE5D4F", "#5B8DEF"])
        fig.update_layout(yaxis_title=None, xaxis_title="Quantity", legend_title_text="")
        show_chart(style_fig(fig, 380))

    section("Source summary")
    src = (f.groupby("Source type")
            .agg(**{"PR lines": ("Material", "size"), "Materials": ("Material", "nunique"),
                    "Remaining quantity": ("Still to be delivered", "sum"),
                    "Overdue remaining": ("Overdue remaining", "sum")}).reset_index())
    src["Overdue share %"] = (src["Overdue remaining"] / src["Remaining quantity"].where(src["Remaining quantity"] > 0) * 100).fillna(0)
    show_df(src, pct_config("Overdue share %"))


def render_material_tab(ms):
    section("Material summary", "One row per material, most critical first")
    c1, c2 = st.columns([2, 3])
    q = c1.text_input("Search material code / description", key="mat_search")
    pick = c2.radio("Status", ["All", "🔴 Critical", "🟠 Overdue", "🟢 On track"], horizontal=True, key="mat_status")
    view = ms[search_mask(ms, q, ["Material", "Description"])]
    if pick != "All":
        view = view[view["Status"] == pick]
    cols = ["Status", "Material", "Description", "Source", "Open PRs", "Remaining quantity", "Completion %",
            "Overdue PRs", "Overdue remaining", "Worst delay (days)", "Next due date", "Oldest delivery date",
            "Latest delivery date", "Earliest requisition"]
    if "Unrestricted stock" in view.columns:
        cols.insert(6, "Unrestricted stock")
    out = with_jalali(view, ["Next due date", "Oldest delivery date", "Latest delivery date", "Earliest requisition"])[cols]
    out = out.rename(columns={c: f"{c} (شمسی)" for c in ["Next due date", "Oldest delivery date",
                                                       "Latest delivery date", "Earliest requisition"]})
    st.caption(f"{len(out):,} materials")
    show_df(out, pct_config("Completion %"), height=560)
    csv_button("Download material summary (CSV)", out, "open_pr_material_summary.csv")


def render_pr_tab(f):
    section("Open PR line details")
    q = st.text_input("Search material code / description", key="pr_search")
    view = f[search_mask(f, q, ["Material", "Short Text"])]
    cols = ["Material", "Short Text", "Source type", "Purchasing group", "Processing status", "Requisition date",
            "Delivery date", "Quantity requested", "Delivered quantity (estimated)", "Still to be delivered",
            "Delivery completion %", "Days past delivery", "Delay bucket", "Overdue"]
    out = view[cols].sort_values(["Days past delivery", "Still to be delivered"], ascending=[False, False], na_position="last")
    out = with_jalali(out, ["Requisition date", "Delivery date"]).rename(columns={
        "Short Text": "Material description", "Source type": "Source", "Purchasing group": "Purch. group",
        "Requisition date": "Requisition date (شمسی)", "Delivery date": "Delivery date (شمسی)"})
    st.caption(f"{len(out):,} PR lines")
    show_df(out, pct_config("Delivery completion %"), height=560)
    csv_button("Download PR details (CSV)", out, "open_pr_filtered_details.csv")


def render_due_tab(b, stock):
    section(f"Open PRs due within {int(due_window)} days",
            "Open lines whose delivery date is between the analysis date and the window end (overdue lines are in the other tabs).")
    due = b[b["Days until delivery"].between(0, int(due_window), inclusive="both")].copy()
    if due.empty:
        callout(f"No open PRs are due within the next {int(due_window)} days.", "green")
        return
    urgent = due[due["Days until delivery"] <= 3]
    kpi_row([
        ("PR lines due", f"{len(due):,}", f"within {int(due_window)} days", "blue"),
        ("Materials", f"{due['Material'].nunique():,}", "unique", "violet"),
        ("Quantity due", fmt_num(due["Still to be delivered"].sum()), "remaining on those lines", "teal"),
        ("Urgent (≤ 3 days)", f"{len(urgent):,}", f"{urgent['Material'].nunique():,} materials", "red"),
    ])

    d = due.groupby("Days until delivery").agg(qty=("Still to be delivered", "sum"), lines=("Material", "size")).reset_index()
    d["Days until delivery"] = d["Days until delivery"].astype(int)
    d["Urgency"] = np.where(d["Days until delivery"] <= 3, "Urgent (≤3 d)", np.where(d["Days until delivery"] <= 7, "Soon (4–7 d)", "Upcoming"))
    fig = px.bar(d, x="Days until delivery", y="qty", color="Urgency", hover_data=["lines"],
                 color_discrete_map={"Urgent (≤3 d)": "#EE5D4F", "Soon (4–7 d)": "#F5A524", "Upcoming": "#5B8DEF"},
                 title="Quantity due by days remaining")
    fig.update_layout(xaxis_title="Days until delivery date", yaxis_title="Remaining quantity")
    show_chart(style_fig(fig, 320))

    mode = st.radio("View", ["PR lines", "Per material"], horizontal=True, key="due_view")
    stock_map = stock.set_index("Material")["Unrestricted"] if stock is not None else None
    if mode == "PR lines":
        out = due.sort_values(["Days until delivery", "Still to be delivered"], ascending=[True, False]).copy()
        out["Urgency"] = np.where(out["Days until delivery"] <= 3, "🔴 ≤3 d", np.where(out["Days until delivery"] <= 7, "🟠 ≤7 d", "🟡 ≤10 d+"))
        out["Delivery date (شمسی)"] = jalali_series(out["Delivery date"])
        out = out.rename(columns={"Short Text": "Material description", "Source type": "Source", "Days until delivery": "Days remaining"})
        cols = ["Urgency", "Material", "Material description", "Processing status", "Source", "Delivery date (شمسی)",
                "Days remaining", "Quantity requested", "Still to be delivered", "Remaining %"]
        if stock_map is not None:
            out["Unrestricted stock"] = out["Material"].map(stock_map).fillna(0)
            cols.append("Unrestricted stock")
        out = out[cols]
        fname = "prs_due_soon_lines.csv"
    else:
        g = due.groupby("Material").agg(**{
            "Material description": ("Short Text", lambda s: next((x for x in s if x), "")),
            "PR lines": ("Material", "size"), "Remaining quantity": ("Still to be delivered", "sum"),
            "Earliest due": ("Delivery date", "min"), "Days until first delivery": ("Days until delivery", "min"),
        }).reset_index().sort_values(["Days until first delivery", "Remaining quantity"], ascending=[True, False])
        g["Earliest due (شمسی)"] = jalali_series(g["Earliest due"])
        out = g.drop(columns="Earliest due")
        if stock_map is not None:
            out["Unrestricted stock"] = out["Material"].map(stock_map).fillna(0)
        fname = "prs_due_soon_materials.csv"
    show_df(out, pct_config("Remaining %") if "Remaining %" in out.columns else None, height=480)
    csv_button("Download (CSV)", out, fname)


def render_critical_tab(ms):
    section("Critical materials",
            f"Rule: remaining > {min_critical_remaining:g} AND (worst delay ≥ {severe_delay_days} days OR "
            f"completion ≤ {low_delivery_pct}% with at least one overdue PR).")
    crit = ms[ms["Criticality"] == "CRITICAL"]
    counts = ms["Status"].value_counts()
    kpi_row([
        ("Critical", f"{counts.get('🔴 Critical', 0):,}", "materials", "red"),
        ("Overdue (not critical)", f"{counts.get('🟠 Overdue', 0):,}", "materials", "amber"),
        ("On track", f"{counts.get('🟢 On track', 0):,}", "materials", "teal"),
        ("Critical remaining qty", fmt_num(crit["Remaining quantity"].sum()), "units", "violet"),
    ])
    plot = ms[ms["Remaining quantity"] > 0].copy()
    plot["Label"] = [short_label(m, d, 34) for m, d in zip(plot["Material"], plot["Description"])]
    fig = px.scatter(plot, x="Worst delay (days)", y="Remaining quantity", size="Open PRs", color="Status",
                     color_discrete_map=STATUS_COLORS, hover_name="Label", log_y=True,
                     hover_data={"Open PRs": True, "Completion %": ":.0f", "Worst delay (days)": ":.0f"},
                     title="Priority map — top-right = long delay AND large missing quantity")
    fig.add_vline(x=float(severe_delay_days), line_dash="dot", line_color="#EE5D4F",
                  annotation_text=f"{int(severe_delay_days)} d", annotation_position="top")
    fig.update_layout(xaxis_title="Worst delay (days)", yaxis_title="Remaining quantity (log scale)")
    show_chart(style_fig(fig, 430))

    if crit.empty:
        callout("No materials meet the current criticality rules under these filters.", "green")
        return
    cols = ["Material", "Description", "Source", "Open PRs", "Remaining quantity", "Completion %", "Overdue PRs",
            "Overdue remaining", "Worst delay (days)"]
    if "Unrestricted stock" in crit.columns:
        cols.insert(5, "Unrestricted stock")
    out = crit[cols]
    show_df(out, pct_config("Completion %"), height=420)
    csv_button("Download critical materials (CSV)", out, "critical_materials.csv")


def render_buckets_tab(f):
    section("Delay buckets", "Built from the filtered PR lines; unique materials are counted separately.")
    summary = (f.groupby("Delay bucket")
                .agg(**{"Open PR lines": ("Material", "size"), "Unique materials": ("Material", "nunique"),
                        "Remaining quantity": ("Still to be delivered", "sum")})
                .reindex(BUCKETS, fill_value=0).reset_index())
    total = summary["Remaining quantity"].sum()
    summary["Share of remaining %"] = (summary["Remaining quantity"] / total * 100) if total else 0
    c1, c2 = st.columns([2, 3])
    with c1:
        show_df(summary, pct_config("Share of remaining %"))
    with c2:
        piv = f.pivot_table(index="Source type", columns="Delay bucket", values="Still to be delivered",
                            aggfunc="sum", fill_value=0).reindex(columns=BUCKETS, fill_value=0)
        fig = px.imshow(piv, text_auto=".0f", aspect="auto", color_continuous_scale=["#14213a", "#2f5aa8", "#F5A524", "#C7283D"],
                        title="Remaining quantity — source × delay bucket")
        fig.update_layout(xaxis_title=None, yaxis_title=None, coloraxis_showscale=False)
        show_chart(style_fig(fig, 300, legend=False))


# ---------------------------------------------------------------------------
# Inventory-receipts panel
# ---------------------------------------------------------------------------
def movement_excel(by_plant, pivot, detail, rows):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        by_plant.to_excel(writer, sheet_name="By plant", index=False)
        pivot.to_excel(writer, sheet_name="Material x plant", index=False)
        detail.to_excel(writer, sheet_name="By plant and material", index=False)
        rows.to_excel(writer, sheet_name="Movement rows", index=False)
    return buf.getvalue()


def render_movements_tab(open_prs):
    section("Inventory receipts (movement 101)",
            "How much of each material entered inventory in the selected period, per plant. "
            "Movement 102 is a reversal and is never counted as a receipt.")
    if uploaded_movements is None:
        callout("Upload the SAP movement workbook in the sidebar (Data files) to activate this panel.", "amber")
        return
    try:
        mv = load_movements(uploaded_movements.getvalue())
    except Exception as e:
        st.error(f"Could not process the movement workbook: {e}")
        return

    valid_ts = mv["Entry timestamp"].dropna()
    if valid_ts.empty:
        st.error("No valid 'Entered on' / 'Entered at' (or Posting Date) values could be parsed.")
        return
    data_end = valid_ts.max()

    # ---- period controls ----
    c1, c2 = st.columns([3, 2])
    period = c1.radio("Period", ["Last 24 h", "Last 48 h", "Last 7 days", "Custom hours", "Date range"],
                      horizontal=True, key="mv_period")
    window_start = window_end = None
    if period == "Date range":
        default_to = data_end.date()
        picked = c2.date_input("From – to", value=(default_to - pd.Timedelta(days=1), default_to),
                               format="DD/MM/YYYY", key="mv_range")
        if isinstance(picked, (tuple, list)) and len(picked) == 2:
            window_start = pd.Timestamp(picked[0])
            window_end = pd.Timestamp(picked[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
        else:
            st.info("Pick both a start and an end date.")
            return
    else:
        hours = {"Last 24 h": 24, "Last 48 h": 48, "Last 7 days": 168}.get(period)
        if hours is None:
            hours = int(c2.number_input("Hours", min_value=1, max_value=8760, value=24, step=1, key="mv_hours"))
        anchor = st.radio("Window ends at", ["Now (this computer)", "Latest entry in file"], horizontal=True, key="mv_anchor")
        window_end = pd.Timestamp.now() if anchor.startswith("Now") else data_end
        window_start = window_end - pd.Timedelta(hours=hours)

    net_reversals = st.toggle("Net out reversals (subtract 102 quantity from 101)", value=False, key="mv_net",
                              help="OFF (default): 102 lines are simply excluded. ON: their quantity is subtracted per plant/material.")

    in_window = mv["Entry timestamp"].between(window_start, window_end, inclusive="both")
    base_ok = mv["Material"].ne("") & mv["Plant"].ne("")
    receipts = mv[in_window & base_ok & mv["Movement type"].eq("101") & (mv["Qty in unit of entry"] > 0)].copy()
    reversals = mv[in_window & base_ok & mv["Movement type"].eq("102")].copy()

    st.caption(f"Window: {window_start:%Y-%m-%d %H:%M} → {window_end:%Y-%m-%d %H:%M} "
               f"({to_jalali_str(window_start)} → {to_jalali_str(window_end)} شمسی) · file covers up to {data_end:%Y-%m-%d %H:%M}")

    if receipts.empty:
        callout("No 101 receipt movements were found in the selected period. Try a longer period or "
                "“Latest entry in file” as the window end.", "amber")
        return

    # ---- filters ----
    plants = sorted(receipts["Plant"].unique().tolist())
    f1, f2 = st.columns([2, 2])
    sel_plants = f1.multiselect("Plants", plants, default=plants, key="mv_plants")
    mat_q = f2.text_input("Search material", key="mv_search")
    receipts = receipts[receipts["Plant"].isin(sel_plants) & search_mask(receipts, mat_q, ["Material"])]
    reversals = reversals[reversals["Plant"].isin(sel_plants) & search_mask(reversals, mat_q, ["Material"])]
    if receipts.empty:
        callout("No receipts match these filters.", "amber")
        return

    receipts["Signed quantity"] = receipts["Qty in unit of entry"]
    work = receipts
    if net_reversals and not reversals.empty:
        rev = reversals.copy()
        rev["Signed quantity"] = -rev["Qty in unit of entry"].abs()
        work = pd.concat([receipts, rev], ignore_index=True)

    desc_map = {}
    if not open_prs.empty:
        desc_map = (open_prs[open_prs["Short Text"].ne("")].drop_duplicates("Material").set_index("Material")["Short Text"]).to_dict()
    open_rem = open_prs.groupby("Material")["Still to be delivered"].sum() if not open_prs.empty else pd.Series(dtype=float)

    by_plant = (work.groupby("Plant")
                .agg(**{"Received quantity": ("Signed quantity", "sum"), "Materials": ("Material", "nunique"),
                        "Movement lines": ("Material", "size"), "Last entry": ("Entry timestamp", "max")})
                .reset_index().sort_values("Received quantity", ascending=False))
    detail = (work.groupby(["Plant", "Material"])
              .agg(**{"Received quantity": ("Signed quantity", "sum"), "Movement lines": ("Material", "size"),
                      "First entry": ("Entry timestamp", "min"), "Last entry": ("Entry timestamp", "max")})
              .reset_index())
    detail.insert(2, "Description", detail["Material"].map(desc_map).fillna(""))
    detail["Open PR remaining"] = detail["Material"].map(open_rem).fillna(0)
    detail = detail.sort_values(["Plant", "Received quantity"], ascending=[True, False])

    total_qty = work["Signed quantity"].sum()
    rev_qty = reversals["Qty in unit of entry"].abs().sum()
    kpi_row([
        ("Total received", fmt_num(total_qty), "net of reversals" if net_reversals else "101 only", "teal"),
        ("Receipt lines", f"{len(receipts):,}", "101 movements", "blue"),
        ("Materials", f"{receipts['Material'].nunique():,}", "received", "violet"),
        ("Plants", f"{receipts['Plant'].nunique():,}", "with receipts", "amber"),
        ("Reversals (102)", f"{len(reversals):,}", f"{fmt_num(rev_qty)} units " + ("netted" if net_reversals else "excluded"), "red"),
        ("Last receipt", f"{receipts['Entry timestamp'].max():%d %b %H:%M}", to_jalali_str(receipts['Entry timestamp'].max()) + " شمسی", "green"),
    ])

    # ---- charts ----
    g1, g2 = st.columns(2)
    with g1:
        fig = px.bar(by_plant, x="Plant", y="Received quantity", color="Plant", text="Received quantity",
                     hover_data=["Materials", "Movement lines"], title="Received quantity by plant",
                     color_discrete_sequence=ACCENT)
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig.update_layout(xaxis_title=None, yaxis_title="Received quantity")
        show_chart(style_fig(fig, 380, legend=False))
    with g2:
        span_h = (window_end - window_start).total_seconds() / 3600
        step = "60min" if span_h <= 72 else "1440min"
        t = work.copy()
        t["Period"] = t["Entry timestamp"].dt.floor(step)
        tl = t.groupby(["Period", "Plant"])["Signed quantity"].sum().reset_index()
        fig = px.bar(tl, x="Period", y="Signed quantity", color="Plant", color_discrete_sequence=ACCENT,
                     title=f"Receipts over time ({'hourly' if span_h <= 72 else 'daily'})")
        fig.update_layout(xaxis_title=None, yaxis_title="Received quantity")
        show_chart(style_fig(fig, 380))

    g3, g4 = st.columns(2)
    with g3:
        topm = (work.groupby(["Material", "Plant"])["Signed quantity"].sum().reset_index())
        tot = topm.groupby("Material")["Signed quantity"].sum().nlargest(15).index
        topm = topm[topm["Material"].isin(tot)].copy()
        topm["Label"] = [short_label(m, desc_map.get(m, ""), 22) for m in topm["Material"]]
        order = [short_label(m, desc_map.get(m, ""), 22) for m in tot[::-1]]
        fig = px.bar(topm, y="Label", x="Signed quantity", color="Plant", orientation="h",
                     color_discrete_sequence=ACCENT, category_orders={"Label": order}, title="Top 15 received materials")
        fig.update_layout(yaxis_title=None, xaxis_title="Received quantity")
        show_chart(style_fig(fig, 460))
    with g4:
        tm = detail[detail["Received quantity"] > 0]
        fig = px.treemap(tm, path=[px.Constant("All plants"), "Plant", "Material"], values="Received quantity",
                         color="Plant", color_discrete_sequence=ACCENT, title="Share of receipts — plant → material")
        show_chart(style_fig(fig, 460, legend=False))

    # ---- tables ----
    section("Plant overview")
    plant_tbl = by_plant.copy()
    plant_tbl["Last entry (شمسی)"] = jalali_datetime_series(plant_tbl["Last entry"])
    show_df(plant_tbl.drop(columns="Last entry"))

    section("Material × plant matrix", "Each row is a material; plant columns show the quantity received; sorted by total.")
    piv = work.pivot_table(index="Material", columns="Plant", values="Signed quantity", aggfunc="sum", fill_value=0)
    piv["Total"] = piv.sum(axis=1)
    piv = piv.sort_values("Total", ascending=False).reset_index()
    piv.insert(1, "Description", piv["Material"].map(desc_map).fillna(""))
    piv["Open PR remaining"] = piv["Material"].map(open_rem).fillna(0)
    show_df(piv, height=480)

    section("Detail by plant and material", "“Open PR remaining” links the receipt to the PR still waiting for this material (PR filters ignored).")
    det_show = detail.copy()
    det_show["First entry (شمسی)"] = jalali_datetime_series(det_show["First entry"])
    det_show["Last entry (شمسی)"] = jalali_datetime_series(det_show["Last entry"])
    det_show = det_show.drop(columns=["First entry", "Last entry"])
    show_df(det_show, height=420)

    rows = work.copy()
    rows["Entry (شمسی)"] = jalali_datetime_series(rows["Entry timestamp"])
    st.download_button(
        "Download receipts analysis (Excel)",
        data=movement_excel(plant_tbl.drop(columns="Last entry"), piv, det_show, rows),
        file_name="inventory_receipts_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Main layout
# ---------------------------------------------------------------------------
tab_names = ["Overview", "Material summary", "PR details", f"Due ≤{int(due_window)} days", "Critical materials",
             "Delay buckets", "Inventory receipts"]
t_over, t_mat, t_pr, t_due, t_crit, t_buck, t_mov = st.tabs(tab_names)

ms = pd.DataFrame()
if pr_ready and not filtered.empty:
    ms = build_material_summary(filtered, stock_df, severe_delay_days, low_delivery_pct, min_critical_remaining)

    tot_req = filtered["Quantity requested"].sum()
    tot_rem = filtered["Still to be delivered"].sum()
    tot_del = filtered["Delivered quantity (estimated)"].sum()
    completion = tot_del / tot_req * 100 if tot_req else 0
    od = filtered[filtered["Overdue"]]
    n_crit = int((ms["Criticality"] == "CRITICAL").sum())
    soon_n = int(filtered["Days until delivery"].between(0, int(due_window), inclusive="both").sum())
    src_rem = filtered.groupby("Source type")["Still to be delivered"].sum()

    kpi_row([
        ("Open PR lines", f"{len(filtered):,}", f"{filtered['Material'].nunique():,} unique materials", "blue"),
        ("Remaining quantity", fmt_num(tot_rem), f"{completion:.0f}% est. delivered", "teal"),
        ("Overdue lines", f"{len(od):,}", f"{len(od) / len(filtered) * 100:.0f}% of open lines", "red"),
        ("Overdue remaining", fmt_num(od["Still to be delivered"].sum()), "units past delivery date", "amber"),
    ])
    st.write("")
    kpi_row([
        ("Critical materials", f"{n_crit:,}", "see Critical tab", "red"),
        (f"Due ≤{int(due_window)} days", f"{soon_n:,}", "PR lines arriving soon", "violet"),
        ("Domestic remaining", fmt_num(src_rem.get("داخلی", 0)), "داخلی", "blue"),
        ("Foreign remaining", fmt_num(src_rem.get("خارجی (FPG)", 0)), "خارجی (FPG)", "green"),
    ])

pr_empty_msg = pr_problem or "No rows match the current filters. Change or clear one or more filters."

with t_over:
    if ms.empty:
        st.warning(pr_empty_msg)
    else:
        render_overview(filtered, ms)
with t_mat:
    if ms.empty:
        st.warning(pr_empty_msg)
    else:
        render_material_tab(ms)
with t_pr:
    if ms.empty:
        st.warning(pr_empty_msg)
    else:
        render_pr_tab(filtered)
with t_due:
    if not pr_ready or base.empty:
        st.warning(pr_empty_msg)
    else:
        render_due_tab(base, stock_df)
with t_crit:
    if ms.empty:
        st.warning(pr_empty_msg)
    else:
        render_critical_tab(ms)
with t_buck:
    if ms.empty:
        st.warning(pr_empty_msg)
    else:
        render_buckets_tab(filtered)
with t_mov:
    render_movements_tab(df if pr_ready else pd.DataFrame())

# ---------------------------------------------------------------------------
# Full Excel export (sidebar)
# ---------------------------------------------------------------------------
if not ms.empty:
    pr_export = with_jalali(filtered, ["Requisition date", "Delivery date", "Upcoming delivery date"]).rename(columns={
        "Short Text": "Material description", "Source type": "Source", "Purchasing group": "Purch. group",
        "Requisition date": "Requisition date (شمسی)", "Delivery date": "Delivery date (شمسی)"})
    ms_export = with_jalali(ms, ["Next due date", "Oldest delivery date", "Latest delivery date", "Earliest requisition"])
    bucket_export = (filtered.groupby("Delay bucket")
                     .agg(**{"PR lines": ("Material", "size"), "Materials": ("Material", "nunique"),
                             "Remaining quantity": ("Still to be delivered", "sum")})
                     .reindex(BUCKETS, fill_value=0).reset_index())
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pr_export.to_excel(writer, sheet_name="Filtered PR details", index=False)
        ms_export.to_excel(writer, sheet_name="Material summary", index=False)
        bucket_export.to_excel(writer, sheet_name="Delay buckets", index=False)
        ms_export[ms_export["Criticality"] == "CRITICAL"].to_excel(writer, sheet_name="Critical materials", index=False)
    st.sidebar.divider()
    st.sidebar.download_button(
        "⬇ Complete PR report (Excel)", data=buf.getvalue(), file_name="open_prs_analytical_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
