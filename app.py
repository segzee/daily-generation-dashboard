"""
Plant Daily Operational Dashboard — Streamlit replica of the Power BI report.

Run with:
    streamlit run app.py

Upload one or more "DAILY PLANT STATUS FOR <DD MONTH YYYY>.xlsx" files
(the same format used to feed the Power BI folder pipeline) and the app
rebuilds the same KPI cards, unit comparison chart, gauge, unit status
table, trend line, and offline-unit alert.
"""

import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------
# Page setup + palette (matches the final Power BI color scheme)
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Plant Daily Operational Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed",
)

CREAM = "#F8F6F0"
BORDER = "#E1E0D9"
GREEN_BG, GREEN_TX = "#EAF3DE", "#27500A"
RED_BG, RED_TX = "#FCEBEB", "#791F1F"
AMBER_BG, AMBER_TX = "#FAEEDA", "#854F0B"
EXPECTED_COLOR = "#C3B8A3"
ACTUAL_COLOR = "#0C447C"

st.markdown(
    f"""
    <style>
    .stApp {{ background-color: #FFFFFF; }}
    .kpi-card {{
        background: {CREAM}; border: 1px solid {BORDER}; border-radius: 6px;
        min-height: 76px; padding: 9px 11px; text-align: left;
    }}
    .kpi-card.good {{ background: {GREEN_BG}; border-color: #C0DD97; }}
    .kpi-card.bad {{ background: {RED_BG}; border-color: #F09595; }}
    .kpi-label {{ font-size: 11px; line-height: 1.25; color: #6E6E6E; margin-bottom: 4px; }}
    .kpi-label.good {{ color: {GREEN_TX}; }}
    .kpi-label.bad {{ color: {RED_TX}; }}
    .kpi-value {{ font-size: 20px; line-height: 1.1; font-weight: 600; color: #0B0B0B; white-space: nowrap; }}
    .kpi-value.good {{ color: {GREEN_TX}; }}
    .kpi-value.bad {{ color: {RED_TX}; }}
    .section-box {{
        background: #FFFFFF; border: 1px solid {BORDER}; border-radius: 6px;
        padding: 8px 10px;
    }}
    .section-title {{ font-size: 11px; color: #6E6E6E; margin-bottom: 3px; }}
    .alert-box {{
        background: {RED_BG}; border-radius: 8px; padding: 14px 16px;
    }}
    .alert-heading {{ color: {RED_TX}; font-size: 14px; font-weight: 600; margin-bottom: 4px; }}
    .alert-body {{ color: #A32D2D; font-size: 13px; line-height: 1.4; }}
    [data-testid="stSidebar"] {{ width: 280px; }}
    [data-testid="stSidebarContent"] {{ padding: 1rem 1.1rem; }}
    [data-testid="stMainBlockContainer"] {{ padding: 1.25rem 1.5rem 2rem; max-width: 1400px; }}
    [data-testid="stHorizontalBlock"] {{ gap: 0.65rem; }}
    [data-testid="stDataFrame"] {{ background: #FFFFFF; }}
    [data-testid="stDataFrame"] iframe {{ background: #FFFFFF; }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------
# Extraction — mirrors the corrected Power Query cell mapping
# ----------------------------------------------------------------------
def parse_daily_file(file, filename: str):
    import openpyxl

    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=1, max_row=36, values_only=True))

    def cell(r, c):
        try:
            return rows[r - 1][c]
        except IndexError:
            return None

    m = re.search(r"FOR[_ ]?(\d{2})[_ ]([A-Za-z]+)[_ ](\d{4})", filename.upper())
    report_date = None
    if m:
        day, month_name, year = m.groups()
        report_date = datetime.strptime(f"{day} {month_name} {year}", "%d %B %Y").date()

    units = []
    for r in (14, 15, 16, 17):
        units.append(
            {
                "Unit": cell(r, 1),
                "AvailCap": cell(r, 2),
                "DeclaredCap": cell(r, 3),
                "Load0600": cell(r, 4),
                "PrevDeclaredCap": cell(r, 5),
                "DeclaredHrs": cell(r, 6),
                "RunHrs": cell(r, 7),
                "DeratedHrs": cell(r, 8),
                "ExpectedGen": cell(r, 9),
                "ActualGen": cell(r, 10),
                "EnergyDeficit": cell(r, 11),
                "Comments": cell(r, 12),
                "ReportDate": report_date,
            }
        )

    summary = {
        "ReportDate": report_date,
        "StationServiceConsumption": cell(21, 4),
        "EnergySentOut": cell(21, 9),
        "MaxLoadGenerated": cell(21, 12),
        "MaxLoadTime": cell(21, 14),
        "AvgGenerationDay": cell(22, 4),
        "DeclaredCapUtilPct": cell(22, 9),
        "PlantWaterPowerRatio": cell(22, 12),
        "StationWaterPowerRatio": cell(22, 15),
        "HeadWater": cell(26, 6),
        "TailWater": cell(26, 9),
        "GrossOperatingHead": cell(26, 11),
        "AvgInflow": cell(27, 6),
        "AvgTurbineDischarge": cell(27, 9),
        "Spillage": cell(27, 11),
        "AvgStationDischarge": cell(27, 13),
        "StorageDifferential": cell(27, 15),
        "PctLiveVolume": cell(28, 6),
    }
    return units, summary


@st.cache_data(show_spinner=False)
def load_all(uploaded_files):
    unit_rows, summary_rows = [], []
    for f in uploaded_files:
        filename = getattr(f, "name", Path(f).name)
        if filename.startswith("~$"):
            continue
        if hasattr(f, "read"):
            units, summary = parse_daily_file(f, filename)
        else:
            with open(f, "rb") as source:
                units, summary = parse_daily_file(source, filename)
        unit_rows.extend(units)
        summary_rows.append(summary)

    unit_df = pd.DataFrame(unit_rows).dropna(subset=["ReportDate"])
    summary_df = pd.DataFrame(summary_rows).dropna(subset=["ReportDate"])
    unit_df = unit_df.sort_values(["ReportDate", "Unit"])
    summary_df = summary_df.sort_values("ReportDate")
    return unit_df, summary_df


# ----------------------------------------------------------------------
# Sidebar — file upload (stand-in for the Power BI folder connector)
# ----------------------------------------------------------------------
st.sidebar.header("Data source")
uploaded_files = st.sidebar.file_uploader(
    "Upload daily plant status reports (.xlsx)",
    type="xlsx",
    accept_multiple_files=True,
)

data_folder = Path(__file__).parent / "Data"
data_files = sorted(data_folder.glob("*.xlsx"))
source_files = data_files + list(uploaded_files or [])

# Keep one copy when an uploaded report already exists in the Data folder.
unique_sources = {}
for source in source_files:
    source_name = getattr(source, "name", Path(source).name)
    unique_sources[source_name.lower()] = source
source_files = list(unique_sources.values())

if not source_files:
    st.title("Plant Daily Operational Dashboard")
    st.info(
        "Add one or more `DAILY PLANT STATUS FOR <DD MONTH YYYY>.xlsx` files "
        "to the Data folder or upload them in the sidebar."
    )
    st.stop()

unit_df, summary_df = load_all(source_files)

if summary_df.empty:
    st.error(
        "No valid reports could be parsed. Check the filenames follow the "
        "`... FOR DD MONTH YYYY.xlsx` pattern."
    )
    st.stop()

# ----------------------------------------------------------------------
# Header row: title + cumulative date picker (defaults to latest day)
# ----------------------------------------------------------------------
all_dates = sorted(summary_df["ReportDate"].unique())
title_col, date_col = st.columns([4, 1], gap="small")
with title_col:
    st.markdown(
        "<div style='font-size:18px; font-weight:600; padding-top:10px;'>"
        "Plant Daily Operational Dashboard</div>",
        unsafe_allow_html=True,
    )
with date_col:
    selected_date = st.selectbox(
        "Cumulative data as of", all_dates, index=len(all_dates) - 1, format_func=lambda d: d.strftime("%A, %d %B %Y")
    )

day_summary = summary_df[summary_df["ReportDate"] == selected_date].iloc[0]
day_units = unit_df[unit_df["ReportDate"] <= selected_date].copy()
selected_day_units = day_units[day_units["ReportDate"] == selected_date].copy()

summary_to_date = summary_df[summary_df["ReportDate"] <= selected_date]
summary_totals = summary_to_date[
    ["EnergySentOut", "MaxLoadGenerated", "AvgGenerationDay", "AvgInflow"]
].sum()
unit_totals = day_units.groupby("Unit", as_index=False)[
    ["ExpectedGen", "ActualGen", "EnergyDeficit"]
].sum()
unit_totals = unit_totals.merge(
    selected_day_units[["Unit", "AvailCap"]], on="Unit", how="left"
)

st.write("")

# ----------------------------------------------------------------------
# KPI cards
# ----------------------------------------------------------------------
total_deficit = day_units["EnergyDeficit"].sum()
k1, k2, k3, k4, k5 = st.columns(5, gap="small")

def kpi(col, label, value, tone=""):
    col.markdown(
        f"""<div class="kpi-card {tone}">
                <div class="kpi-label {tone}">{label}</div>
                <div class="kpi-value {tone}">{value}</div>
            </div>""",
        unsafe_allow_html=True,
    )

kpi(k1, "Energy sent out (MWh)", f"{summary_totals['EnergySentOut']:,.1f}")
kpi(k2, "Max load (MW)", f"{summary_totals['MaxLoadGenerated']:,.2f}")
kpi(k3, "Avg generation (MW)", f"{summary_totals['AvgGenerationDay']:,.2f}")
kpi(k4, "Avg inflow (m³/s)", f"{summary_totals['AvgInflow']:,.0f}")
kpi(k5, "Total energy deficit (MWh)", f"{total_deficit:,.1f}", tone="bad")

st.write("")

# ----------------------------------------------------------------------
# Bar chart + unit status + gauge row
# ----------------------------------------------------------------------
chart_col, table_col, gauge_col = st.columns([2, 1.15, 1], gap="small")

with chart_col:
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Cumulative expected vs actual generation by unit</div>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_bar(x=unit_totals["Unit"], y=unit_totals["ExpectedGen"], name="Expected", marker_color=EXPECTED_COLOR)
    fig.add_bar(x=unit_totals["Unit"], y=unit_totals["ActualGen"], name="Actual", marker_color=ACTUAL_COLOR)
    fig.update_layout(
        barmode="group",
        template="plotly_white",
        height=280,
        margin=dict(l=10, r=10, t=10, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", y=-0.15),
    )
    fig.update_yaxes(gridcolor=BORDER)
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

with table_col:
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Cumulative unit status</div>', unsafe_allow_html=True)

    display_df = unit_totals[["Unit", "AvailCap", "ActualGen", "EnergyDeficit"]].copy()
    operating = display_df[display_df["AvailCap"] > 0]
    lo, hi = (operating["EnergyDeficit"].min(), operating["EnergyDeficit"].max()) if not operating.empty else (0, 1)

    def deficit_style(row):
        if row["AvailCap"] == 0:
            return ["color: #898781"] * len(row)
        pct = (row["EnergyDeficit"] - lo) / (hi - lo) if hi > lo else 0
        bg = AMBER_BG if pct < 0.6 else RED_BG
        tx = AMBER_TX if pct < 0.6 else RED_TX
        styles = [""] * len(row)
        styles[-1] = f"background-color: {bg}; color: {tx};"
        return styles

    styled = display_df.style.set_properties(
        **{"background-color": "#FFFFFF", "color": "#222222"}
    ).apply(deficit_style, axis=1).format(
        {"AvailCap": "{:,.0f}", "ActualGen": "{:,.1f}", "EnergyDeficit": "{:,.1f}"}
    )
    st.dataframe(styled, hide_index=True, use_container_width=True, height=220)
    st.markdown("</div>", unsafe_allow_html=True)

    offline_units = selected_day_units[selected_day_units["AvailCap"] == 0]
    if not offline_units.empty:
        for _, u in offline_units.iterrows():
            comment = (u["Comments"] or "")[:120]
            st.markdown(
                f"""<div class="alert-box">
                        <div class="alert-heading">Unit {u['Unit']} offline</div>
                        <div class="alert-body">{comment}...</div>
                    </div>""",
                unsafe_allow_html=True,
            )

with gauge_col:
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Capacity utilization</div>', unsafe_allow_html=True)
    util = day_summary["DeclaredCapUtilPct"]
    gauge_fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=util,
            number={"suffix": "%", "font": {"size": 30}},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": ACTUAL_COLOR},
                "bgcolor": BORDER,
                "threshold": {"line": {"color": "#791F1F", "width": 3}, "value": 90},
            },
        )
    )
    gauge_fig.update_layout(
        template="plotly_white", height=280, margin=dict(l=10, r=10, t=10, b=10)
    )
    st.plotly_chart(gauge_fig, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

st.write("")

# ----------------------------------------------------------------------
# Trend line (always shows ALL uploaded days, not filtered by the picker)
# ----------------------------------------------------------------------
st.markdown('<div class="section-box">', unsafe_allow_html=True)
st.markdown('<div class="section-title">Actual vs expected generation over time (MWh)</div>', unsafe_allow_html=True)
daily_totals = unit_df.groupby("ReportDate")[["ExpectedGen", "ActualGen"]].sum().reset_index()
trend_fig = go.Figure()
trend_fig.add_scatter(
    x=daily_totals["ReportDate"], y=daily_totals["ExpectedGen"],
    name="Expected", mode="lines+markers", line=dict(color=EXPECTED_COLOR, width=2),
)
trend_fig.add_scatter(
    x=daily_totals["ReportDate"], y=daily_totals["ActualGen"],
    name="Actual", mode="lines+markers", line=dict(color=ACTUAL_COLOR, width=2),
)
trend_fig.update_layout(
    template="plotly_white", height=220, margin=dict(l=10, r=10, t=10, b=10),
    plot_bgcolor="white", paper_bgcolor="white",
    legend=dict(orientation="h", y=-0.25),
)
trend_fig.update_yaxes(gridcolor=BORDER)
st.plotly_chart(trend_fig, use_container_width=True)
st.markdown("</div>", unsafe_allow_html=True)

