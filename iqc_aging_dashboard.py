# DASHBOARD VERSION: BIG-NUMBER-TABLES-V2
import os
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="IQC Aging Analysis Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ------------------------------------------------------------
# STYLE
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    .stApp {background-color: #F6F8FC;}
    [data-testid="stSidebar"] {background-color: #FFFFFF; border-right: 1px solid #E8EDF5;}
    .block-container {padding-top: 1.1rem; padding-bottom: 2.5rem; max-width: 1800px;}

    .hero {
        background: linear-gradient(90deg, #0A3A8D 0%, #065AC9 100%);
        border-radius: 16px;
        padding: 22px 28px;
        margin-bottom: 18px;
        color: white;
        box-shadow: 0 8px 24px rgba(14, 75, 153, 0.14);
    }
    .hero h1 {font-size: 30px; margin: 0; font-weight: 800; letter-spacing: .2px;}
    .hero p {margin: 4px 0 0 0; opacity: .86; font-size: 14px;}

    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E7ECF4;
        border-radius: 15px;
        padding: 18px 19px;
        min-height: 120px;
        box-shadow: 0 5px 18px rgba(31, 48, 78, 0.05);
    }
    .kpi-label {font-size: 12px; color: #667085; font-weight: 700; text-transform: uppercase;}
    .kpi-value {font-size: 27px; color: #172B4D; font-weight: 800; margin-top: 8px;}
    .kpi-note {font-size: 12px; color: #98A2B3; margin-top: 6px;}

    .section-title {
        font-size: 16px;
        font-weight: 800;
        color: #172B4D;
        margin: 10px 0 4px 0;
    }
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #E7ECF4;
        border-radius: 14px;
        padding: 14px 16px;
    }
    div[data-testid="stDataFrame"] {
        border: 1px solid #E7ECF4;
        border-radius: 12px;
        overflow: hidden;
    }

    .big-table-wrap {
        background: #FFFFFF;
        border: 1px solid #DDE4EE;
        border-radius: 14px;
        overflow-x: auto;
        box-shadow: 0 4px 14px rgba(31, 48, 78, 0.05);
        margin-top: 4px;
        margin-bottom: 14px;
    }
    .big-table {
        width: 100%;
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        background: #FFFFFF;
    }
    .big-table thead th {
        background: #F3F5FA;
        color: #566176;
        font-size: 18px;
        font-weight: 800;
        padding: 16px 14px;
        border-bottom: 1px solid #DDE4EE;
        border-right: 1px solid #E5EAF1;
        white-space: nowrap;
        text-align: left;
    }
    .big-table tbody td {
        color: #172B4D;
        font-size: 18px;
        font-weight: 650;
        padding: 16px 14px;
        border-bottom: 1px solid #E7ECF4;
        border-right: 1px solid #EEF1F5;
        white-space: nowrap;
    }
    .big-table tbody tr:nth-child(even) {background: #FBFCFE;}
    .big-table tbody tr:hover {background: #F2F6FF;}
    .big-table td.num {
        text-align: right;
        font-size: 28px !important;
        font-weight: 900 !important;
        color: #3B2BC5 !important;
        font-variant-numeric: tabular-nums;
    }
    .big-table td.rank {
        text-align: center;
        font-size: 28px !important;
        font-weight: 900 !important;
        color: #3B2BC5 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
REQUIRED_COLUMNS = [
    "Item",
    "Date Created",
    "Location",
    "Quantity Received",
    "Quantity Approved (Total)",
    "Quantity In Quarantine",
    "Quantity To Inspect",
    "Inspection Category",
]

NUMERIC_COLUMNS = [
    "Quantity Received",
    "Quantity Approved (Total)",
    "Quantity In Quarantine",
    "Quantity To Inspect",
    "SOH",
    "Aging day",
]

AGING_ORDER = ["0-7", "8-14", "15-30", "31-60", "61-90", "91-180", "181-365", ">365"]


def aging_bucket(days):
    if pd.isna(days):
        return "Unknown"
    days = int(days)
    if days <= 7:
        return "0-7"
    if days <= 14:
        return "8-14"
    if days <= 30:
        return "15-30"
    if days <= 60:
        return "31-60"
    if days <= 90:
        return "61-90"
    if days <= 180:
        return "91-180"
    if days <= 365:
        return "181-365"
    return ">365"


@st.cache_data(show_spinner=False)
def load_excel(source):
    df = pd.read_excel(source, sheet_name="Sheet2", engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    # Clean core fields
    df["Item"] = df["Item"].astype(str).str.strip()
    df["Location"] = df["Location"].fillna("Unknown").astype(str).str.strip()
    df["Inspection Category"] = df["Inspection Category"].fillna("Unknown").astype(str).str.strip()
    df["Date Created"] = pd.to_datetime(df["Date Created"], errors="coerce")

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Aging day in the workbook is a NOW()-Date formula. Recalculate in Python so
    # the dashboard is always current and does not depend on Excel cached formulas.
    today = pd.Timestamp.now().normalize()
    df["Aging day"] = (today - df["Date Created"].dt.normalize()).dt.days.clip(lower=0)
    df["Aging Bucket"] = df["Aging day"].apply(aging_bucket)
    df["Created Date"] = df["Date Created"].dt.date

    return df


def fmt_num(value):
    if pd.isna(value):
        return "-"
    return f"{value:,.0f}"


def render_big_table(df, numeric_cols=None, rank_col=None):
    """Render a custom HTML table so numeric font size is fully controllable."""
    import html as _html

    numeric_cols = set(numeric_cols or [])
    work = df.copy()

    parts = ['<div class="big-table-wrap"><table class="big-table"><thead><tr>']
    for col in work.columns:
        parts.append(f'<th>{_html.escape(str(col))}</th>')
    parts.append('</tr></thead><tbody>')

    for _, row in work.iterrows():
        parts.append('<tr>')
        for col in work.columns:
            value = row[col]
            if col in numeric_cols:
                display = fmt_num(value)
                parts.append(f'<td class="num">{_html.escape(display)}</td>')
            elif rank_col and col == rank_col:
                display = "-" if pd.isna(value) else str(value)
                parts.append(f'<td class="rank">{_html.escape(display)}</td>')
            else:
                display = "-" if pd.isna(value) else str(value)
                parts.append(f'<td class="text">{_html.escape(display)}</td>')
        parts.append('</tr>')

    parts.append('</tbody></table></div>')
    st.markdown(''.join(parts), unsafe_allow_html=True)


def kpi_card(label, value, note=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def chart_layout(fig, height=390):
    fig.update_layout(
        height=height,
        margin=dict(l=20, r=20, t=55, b=25),
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Arial", color="#344054"),
        title_font=dict(size=16, color="#172B4D"),
        legend_title_text="",
        hoverlabel=dict(bgcolor="white"),
    )
    fig.update_xaxes(showgrid=False, linecolor="#E7ECF4")
    fig.update_yaxes(gridcolor="#EEF2F7", zeroline=False)
    return fig


# ------------------------------------------------------------
# DATA INPUT
# ------------------------------------------------------------
default_file = "IQC Aging day(1).xlsx"
uploaded = st.sidebar.file_uploader("Upload IQC Aging Excel", type=["xlsx"])

try:
    if uploaded is not None:
        df = load_excel(uploaded)
        source_name = uploaded.name
    elif os.path.exists(default_file):
        df = load_excel(default_file)
        source_name = default_file
    else:
        st.info("Upload file 'IQC Aging day(1).xlsx' in the left sidebar to start.")
        st.stop()
except Exception as exc:
    st.error(f"Cannot read Excel file: {exc}")
    st.stop()


# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------
st.markdown(
    f"""
    <div class="hero">
        <h1>IQC AGING ANALYSIS DASHBOARD</h1>
        <p>Incoming Quality Control • Source: {source_name} • Refreshed: {datetime.now():%Y-%m-%d %H:%M}</p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# FILTERS
# ------------------------------------------------------------
st.sidebar.markdown("## 🔎 Filters")

valid_dates = df["Date Created"].dropna()
min_date = valid_dates.min().date()
max_date = valid_dates.max().date()

selected_dates = st.sidebar.date_input(
    "Date Created",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)

all_items = sorted(df["Item"].dropna().unique().tolist())
selected_items = st.sidebar.multiselect("Item", all_items, placeholder="All items")

all_locations = sorted(df["Location"].dropna().unique().tolist())
selected_locations = st.sidebar.multiselect("Location", all_locations, placeholder="All locations")

bucket_options = [b for b in AGING_ORDER if b in set(df["Aging Bucket"])]
selected_buckets = st.sidebar.multiselect(
    "Aging Day (Bucket)",
    bucket_options,
    default=bucket_options,
)

all_categories = sorted(df["Inspection Category"].dropna().unique().tolist())
selected_categories = st.sidebar.multiselect(
    "Inspection Category", all_categories, placeholder="All categories"
)

st.sidebar.markdown("---")
top_n = st.sidebar.slider("Top N Items", min_value=1, max_value=20, value=10, step=1)
rank_metric = st.sidebar.selectbox(
    "Rank Top Items By",
    ["Quantity To Inspect", "Quantity In Quarantine", "Quantity Received", "Aging day"],
)

# Apply filter
filtered = df.copy()
if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
    start_date, end_date = selected_dates
    mask = filtered["Date Created"].dt.date.between(start_date, end_date)
    filtered = filtered[mask]

if selected_items:
    filtered = filtered[filtered["Item"].isin(selected_items)]
if selected_locations:
    filtered = filtered[filtered["Location"].isin(selected_locations)]
if selected_buckets:
    filtered = filtered[filtered["Aging Bucket"].isin(selected_buckets)]
if selected_categories:
    filtered = filtered[filtered["Inspection Category"].isin(selected_categories)]

st.sidebar.caption(f"Rows after filters: {len(filtered):,} / {len(df):,}")

if filtered.empty:
    st.warning("No data matches the current filters.")
    st.stop()


# ------------------------------------------------------------
# KPI ROW
# ------------------------------------------------------------
cols = st.columns(5)
with cols[0]:
    kpi_card("Quantity Received", fmt_num(filtered["Quantity Received"].sum()), "Selected records")
with cols[1]:
    kpi_card("In Quarantine", fmt_num(filtered["Quantity In Quarantine"].sum()), "Current filter")
with cols[2]:
    kpi_card("To Inspect", fmt_num(filtered["Quantity To Inspect"].sum()), "Current filter")
with cols[3]:
    kpi_card("Oldest Aging Day", fmt_num(filtered["Aging day"].max()), "Days")
with cols[4]:
    kpi_card("Unique Items", fmt_num(filtered["Item"].nunique()), "Distinct item codes")

st.write("")


# ------------------------------------------------------------
# AGING + TREND
# ------------------------------------------------------------
left, right = st.columns([1, 1.35])

with left:
    aging = (
        filtered.groupby("Aging Bucket", observed=False)
        .agg(
            Quantity_To_Inspect=("Quantity To Inspect", "sum"),
            Rows=("Item", "size"),
        )
        .reset_index()
    )
    aging["Aging Bucket"] = pd.Categorical(aging["Aging Bucket"], AGING_ORDER, ordered=True)
    aging = aging.sort_values("Aging Bucket")

    fig = px.bar(
        aging,
        x="Aging Bucket",
        y="Quantity_To_Inspect",
        text_auto=",.0f",
        title="QUANTITY TO INSPECT BY AGING BUCKET",
        labels={"Quantity_To_Inspect": "Qty To Inspect", "Aging Bucket": "Aging Day"},
    )
    fig.update_traces(marker_color="#1769E0", textposition="outside")
    st.plotly_chart(chart_layout(fig), use_container_width=True)

with right:
    trend = (
        filtered.groupby("Created Date", as_index=False)
        .agg(
            Quantity_To_Inspect=("Quantity To Inspect", "sum"),
            Quantity_In_Quarantine=("Quantity In Quarantine", "sum"),
        )
        .sort_values("Created Date")
    )
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=trend["Created Date"], y=trend["Quantity_To_Inspect"],
            mode="lines+markers", name="To Inspect", line=dict(width=3, color="#1769E0")
        )
    )
    fig.add_trace(
        go.Scatter(
            x=trend["Created Date"], y=trend["Quantity_In_Quarantine"],
            mode="lines+markers", name="In Quarantine", line=dict(width=2, color="#F59E0B")
        )
    )
    fig.update_layout(title="IQC QUANTITY TREND BY CREATED DATE")
    st.plotly_chart(chart_layout(fig), use_container_width=True)


# ------------------------------------------------------------
# LOCATION SECTION
# ------------------------------------------------------------
st.markdown('<div class="section-title">LOCATION ANALYSIS</div>', unsafe_allow_html=True)
loc = (
    filtered.groupby("Location", as_index=False)
    .agg(
        Quantity_Received=("Quantity Received", "sum"),
        Quantity_In_Quarantine=("Quantity In Quarantine", "sum"),
        Quantity_To_Inspect=("Quantity To Inspect", "sum"),
        Unique_Items=("Item", "nunique"),
        Oldest_Aging_Day=("Aging day", "max"),
    )
    .sort_values("Quantity_To_Inspect", ascending=False)
)

c1, c2 = st.columns([1.15, 1])
with c1:
    fig = px.bar(
        loc,
        y="Location",
        x="Quantity_To_Inspect",
        orientation="h",
        text_auto=",.0f",
        title="QUANTITY TO INSPECT BY LOCATION",
        labels={"Quantity_To_Inspect": "Qty To Inspect"},
    )
    fig.update_traces(marker_color="#1769E0")
    fig.update_yaxes(categoryorder="total ascending")
    st.plotly_chart(chart_layout(fig, 430), use_container_width=True)

with c2:
    loc_display = loc.rename(columns={
        "Quantity_Received": "Qty Received",
        "Quantity_In_Quarantine": "In Quarantine",
        "Quantity_To_Inspect": "To Inspect",
        "Unique_Items": "Items",
        "Oldest_Aging_Day": "Oldest Aging",
    })
    render_big_table(
        loc_display,
        numeric_cols=["Qty Received", "In Quarantine", "To Inspect", "Items", "Oldest Aging"],
    )


# ------------------------------------------------------------
# TOP N ITEM SECTION
# ------------------------------------------------------------
st.markdown(f'<div class="section-title">TOP {top_n} ITEM ANALYSIS</div>', unsafe_allow_html=True)

item_summary = (
    filtered.groupby("Item", as_index=False)
    .agg(
        Quantity_Received=("Quantity Received", "sum"),
        Quantity_Approved=("Quantity Approved (Total)", "sum"),
        Quantity_In_Quarantine=("Quantity In Quarantine", "sum"),
        Quantity_To_Inspect=("Quantity To Inspect", "sum"),
        Oldest_Aging_Day=("Aging day", "max"),
        Locations=("Location", "nunique"),
        Receipts=("Item Receipt", "nunique"),
    )
)

metric_map = {
    "Quantity To Inspect": "Quantity_To_Inspect",
    "Quantity In Quarantine": "Quantity_In_Quarantine",
    "Quantity Received": "Quantity_Received",
    "Aging day": "Oldest_Aging_Day",
}
rank_col = metric_map[rank_metric]
top_items = item_summary.sort_values(rank_col, ascending=False).head(top_n).copy()
top_items["Rank"] = range(1, len(top_items) + 1)

fig = px.bar(
    top_items.sort_values(rank_col, ascending=True),
    x=rank_col,
    y="Item",
    orientation="h",
    text=rank_col,
    title=f"TOP {top_n} ITEMS BY {rank_metric.upper()}",
    labels={rank_col: rank_metric, "Item": "Item Code"},
    hover_data=["Quantity_Received", "Quantity_In_Quarantine", "Quantity_To_Inspect", "Oldest_Aging_Day"],
)
fig.update_traces(marker_color="#1769E0", texttemplate="%{text:,.0f}", textposition="outside")
st.plotly_chart(chart_layout(fig, max(440, 38 * top_n + 150)), use_container_width=True)

# Second Top N chart: aging risk
risk_top = item_summary.sort_values(["Oldest_Aging_Day", "Quantity_To_Inspect"], ascending=[False, False]).head(top_n)
fig = px.bar(
    risk_top.sort_values("Oldest_Aging_Day", ascending=True),
    x="Oldest_Aging_Day",
    y="Item",
    orientation="h",
    text="Oldest_Aging_Day",
    title=f"TOP {top_n} OLDEST ITEMS",
    labels={"Oldest_Aging_Day": "Oldest Aging Day", "Item": "Item Code"},
)
fig.update_traces(marker_color="#F59E0B", texttemplate="%{text:,.0f} days", textposition="outside")
st.plotly_chart(chart_layout(fig, max(440, 38 * top_n + 150)), use_container_width=True)

# Top N table
show = top_items[[
    "Rank", "Item", "Quantity_Received", "Quantity_Approved",
    "Quantity_In_Quarantine", "Quantity_To_Inspect",
    "Oldest_Aging_Day", "Locations", "Receipts"
]].rename(columns={
    "Quantity_Received": "Qty Received",
    "Quantity_Approved": "Qty Approved",
    "Quantity_In_Quarantine": "In Quarantine",
    "Quantity_To_Inspect": "To Inspect",
    "Oldest_Aging_Day": "Oldest Aging",
})

render_big_table(
    show,
    numeric_cols=[
        "Qty Received", "Qty Approved", "In Quarantine", "To Inspect",
        "Oldest Aging", "Locations", "Receipts"
    ],
    rank_col="Rank",
)


# ------------------------------------------------------------
# INSPECTION CATEGORY
# ------------------------------------------------------------
st.markdown('<div class="section-title">INSPECTION CATEGORY</div>', unsafe_allow_html=True)
cat = (
    filtered.groupby("Inspection Category", as_index=False)
    .agg(
        Quantity_To_Inspect=("Quantity To Inspect", "sum"),
        Quantity_In_Quarantine=("Quantity In Quarantine", "sum"),
        Rows=("Item", "size"),
    )
    .sort_values("Quantity_To_Inspect", ascending=False)
)
fig = px.bar(
    cat,
    x="Inspection Category",
    y=["Quantity_To_Inspect", "Quantity_In_Quarantine"],
    barmode="group",
    title="QUANTITY BY INSPECTION CATEGORY",
    labels={"value": "Quantity", "variable": "Metric"},
)
st.plotly_chart(chart_layout(fig, 420), use_container_width=True)


# ------------------------------------------------------------
# DETAIL TABLE + DOWNLOAD
# ------------------------------------------------------------
st.markdown('<div class="section-title">FILTERED DETAIL</div>', unsafe_allow_html=True)

detail_cols = [
    "Item", "Date Created", "Location", "Item Receipt", "Line",
    "Quantity Received", "Quantity Approved (Total)", "Quantity In Quarantine",
    "Quantity To Inspect", "Inspection Category", "Aging day", "Aging Bucket"
]
detail_cols = [c for c in detail_cols if c in filtered.columns]
detail = filtered[detail_cols].sort_values("Aging day", ascending=False)

st.dataframe(detail, use_container_width=True, hide_index=True, height=520)

csv = detail.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️ Export filtered data to CSV",
    data=csv,
    file_name=f"iqc_aging_filtered_{datetime.now():%Y%m%d_%H%M}.csv",
    mime="text/csv",
)

st.caption("Aging day is recalculated in Python from Date Created to today's date, so it stays current even if Excel formulas are not recalculated.")
