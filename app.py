"""
app.py  --  VECTRA | Strategic FP&A & PVM Intelligence
========================================================
Engineered & Architected by Dinu Denny Culas

Run from the folder that holds app.py, finance_engine.py and the two demo CSVs:

    streamlit run app.py

Streamlit re-runs this whole file from top to bottom on every click (like Excel recalculating
after you change an input cell), so the code reads as one flat script:

    1. Setup           page config, colours, styling, helper functions
    2. Welcome screen  onboarding cards: demo data or your own ERP export
    3. Sidebar         data source switch, downloads, filters, scenario controls
    4. Data            load, simulate, filter
    5. Dashboard       KPI row and four tabs
    6. Footer
"""

import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from finance_engine import (load_and_clean_opex, get_department_summary, calculate_pvm_analysis,
                            variance_flag, get_template_csv_bytes, get_template_excel_bytes)

# ======================================================================
# 1. SETUP
# ======================================================================
st.set_page_config(page_title="VECTRA | Strategic FP&A & PVM Intelligence", layout="wide")

# Green and red are reserved for Favorable (F) and Unfavorable (U), so they always mean the same thing.
GREEN = "#2E7D32"
RED = "#C62828"
GREY = "#90A4AE"
NAVY = "#1F3A5F"
DEPT_PALETTE = ["#1F77B4", "#607D8B", "#FF7F0E", "#9467BD", "#00897B", "#8D6E63", "#5C6BC0", "#AD1457"]

CREDIT_LINE = "Engineered & Architected by Dinu Denny Culas"
WATERMARK_TEXT = f"VECTRA | {CREDIT_LINE}"

VIEW_WELCOME = "Welcome Screen"
VIEW_DEMO = "Demo ERP Dataset"
VIEW_CUSTOM = "Custom ERP Upload"
VIEWS = [VIEW_WELCOME, VIEW_DEMO, VIEW_CUSTOM]

SOURCE_OPEX = "OPEX ledger row"
SOURCE_SALES = "Sales PVM ledger row"
SOURCE_TYPED = "Type 4 test numbers (PVM)"

DEMO_OPEX_PATH = Path(__file__).parent / "erp_opex_ledger.csv"
DEMO_SALES_PATH = Path(__file__).parent / "sales_pvm_ledger.csv"
DEMO_FILES_FOUND = DEMO_OPEX_PATH.exists() and DEMO_SALES_PATH.exists()

CSV_MIME = "text/csv"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

PAGE_CSS = """
<style>
    /* 1. Flagship Fonts: Outfit (Brand Display), Plus Jakarta Sans (UI), JetBrains Mono (Numbers) */
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@500;600&family=Outfit:wght@600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    /* 2. Global Canvas, Ambient Top Spotlight & Custom Minimal Scrollbar */
    html, body, [class*="css"], .stApp {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
        scroll-behavior: smooth;
        background: radial-gradient(ellipse 80% 50% at 50% -15%, rgba(10, 132, 255, 0.14), transparent 70%), #050507 !important;
    }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: #050507; }
    ::-webkit-scrollbar-thumb { background: #27272A; border-radius: 3px; }
    ::-webkit-scrollbar-thumb:hover { background: #0A84FF; }

    /* Hide Streamlit Default Top Chrome */
    header[data-testid="stHeader"] { background: transparent !important; }
    .stDeployButton { display: none !important; }

    /* 3. Smooth Cubic-Bezier Entrance Physics */
    @keyframes vectraFadeUp {
        0% { opacity: 0; transform: translateY(12px); }
        100% { opacity: 1; transform: translateY(0); }
    }

    /* 4. Signature VECTRA Hero Banner with Specular Top Rim */
    .vectra-banner {
        background: radial-gradient(120% 150% at 0% 0%, rgba(10, 132, 255, 0.25) 0%, rgba(18, 18, 23, 0.92) 55%, rgba(9, 9, 12, 0.98) 100%);
        backdrop-filter: blur(28px);
        -webkit-backdrop-filter: blur(28px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-top: 1px solid rgba(255, 255, 255, 0.24);
        border-radius: 22px;
        padding: 30px 36px;
        margin-bottom: 24px;
        box-shadow: 0 24px 60px rgba(0, 0, 0, 0.65), 0 0 40px rgba(10, 132, 255, 0.08);
        animation: vectraFadeUp 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }
    .vectra-banner h1 {
        font-family: 'Outfit', sans-serif !important;
        font-size: 2.65rem !important;
        font-weight: 800 !important;
        letter-spacing: 0.26em !important;
        text-transform: uppercase;
        background: linear-gradient(180deg, #FFFFFF 0%, #94A3B8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0 !important;
        padding: 0 !important;
    }
    .vectra-banner p {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        color: #94A3B8;
        font-size: 0.96rem;
        font-weight: 500;
        letter-spacing: 0.03em;
        margin: 8px 0 0 0;
    }

    /* 5. Section Titles & Custom Classes */
    h2, h3 {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: -0.02em !important;
        color: #F8FAFC !important;
    }
    .vectra-card-title {
        font-family: 'Outfit', sans-serif !important;
        font-size: 1.28rem;
        font-weight: 600;
        color: #F8FAFC;
        letter-spacing: -0.01em;
        margin-bottom: 6px;
    }
    .vectra-card-text {
        font-size: 0.92rem;
        color: #94A3B8;
        line-height: 1.55;
        margin-bottom: 14px;
    }
    .vectra-badge {
        background: linear-gradient(135deg, rgba(255, 255, 255, 0.07) 0%, rgba(255, 255, 255, 0.02) 100%);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-top: 1px solid rgba(255, 255, 255, 0.2);
        border-radius: 12px;
        padding: 12px 14px;
        font-size: 0.82rem;
        letter-spacing: 0.01em;
        color: #E2E8F0;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.25);
    }
    .vectra-footer {
        text-align: center;
        color: #64748B;
        font-size: 0.85rem;
        font-weight: 500;
        letter-spacing: 0.02em;
        padding: 20px 0;
    }

    /* 6. KPI Metric Cards — Specular Glass + Strict Containment + Tabular Numbers */
    div[data-testid="stMetric"] {
        background: linear-gradient(180deg, rgba(28, 28, 34, 0.78) 0%, rgba(14, 14, 18, 0.92) 100%) !important;
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.07) !important;
        border-top: 1px solid rgba(255, 255, 255, 0.18) !important;
        border-radius: 18px !important;
        padding: 18px 20px !important;
        box-shadow: 0 12px 32px rgba(0, 0, 0, 0.5) !important;
        overflow: hidden !important;
        box-sizing: border-box !important;
        animation: vectraFadeUp 0.55s cubic-bezier(0.16, 1, 0.3, 1) forwards;
        transition: all 0.28s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-4px);
        border-color: rgba(10, 132, 255, 0.45) !important;
        box-shadow: 0 20px 44px rgba(0, 0, 0, 0.65), 0 0 24px rgba(10, 132, 255, 0.15) !important;
    }
    div[data-testid="stMetricValue"],
    div[data-testid="stMetricValue"] > div {
        font-family: 'Outfit', sans-serif !important;
        font-variant-numeric: tabular-nums !important;
        font-size: clamp(1.12rem, 1.45vw, 1.68rem) !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
        line-height: 1.25 !important;
        color: #FFFFFF !important;
        white-space: normal !important;
        word-break: break-word !important;
        overflow-wrap: anywhere !important;
    }
    div[data-testid="stMetricLabel"] > div {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-size: 0.75rem !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.08em !important;
        color: #94A3B8 !important;
    }
    div[data-testid="stMetricDelta"],
    div[data-testid="stMetricDelta"] > div {
        font-family: 'JetBrains Mono', monospace !important;
        font-size: 0.74rem !important;
        font-weight: 500 !important;
        white-space: normal !important;
        line-height: 1.3 !important;
    }

    /* 7. Bento-Box Glass Containers — Clips Inner Square Chart Corners & Eliminates Scroll */
    div[data-testid="stPlotlyChart"] {
        background: linear-gradient(180deg, rgba(22, 22, 27, 0.75) 0%, rgba(12, 12, 16, 0.92) 100%) !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-top: 1px solid rgba(255, 255, 255, 0.18) !important;
        border-radius: 18px !important;
        padding: 14px 18px 18px 18px !important;
        box-shadow: 0 14px 36px rgba(0, 0, 0, 0.45) !important;
        overflow: hidden !important;
        transition: border-color 0.25s ease !important;
    }
    div[data-testid="stPlotlyChart"]:hover {
        border-color: rgba(10, 132, 255, 0.35) !important;
    }
    /* Forces Plotly's internal SVG canvas to be 100% transparent with rounded corners */
    div[data-testid="stPlotlyChart"] .js-plotly-plot,
    div[data-testid="stPlotlyChart"] .plot-container,
    div[data-testid="stPlotlyChart"] .svg-container,
    div[data-testid="stPlotlyChart"] svg.main-svg {
        background: transparent !important;
        border-radius: 14px !important;
        overflow: hidden !important;
    }

    /* Rounded Glass Frame for DataFrames */
    div[data-testid="stDataFrame"] {
        border: 1px solid rgba(255, 255, 255, 0.09) !important;
        border-radius: 14px !important;
        overflow: hidden !important;
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.4) !important;
    }

    /* 8. Floating Segmented Pill Tabs */
    div[data-baseweb="tab-list"] {
        background: rgba(18, 18, 22, 0.9) !important;
        backdrop-filter: blur(16px);
        padding: 6px !important;
        border-radius: 14px !important;
        gap: 6px !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        margin-bottom: 20px !important;
    }
    button[data-baseweb="tab"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        border-radius: 10px !important;
        padding: 9px 18px !important;
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        color: #94A3B8 !important;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    button[data-baseweb="tab"]:hover {
        color: #F8FAFC !important;
        background: rgba(255, 255, 255, 0.04) !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(180deg, rgba(10, 132, 255, 0.3) 0%, rgba(10, 132, 255, 0.14) 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(10, 132, 255, 0.55) !important;
        box-shadow: 0 4px 16px rgba(10, 132, 255, 0.25) !important;
    }
    div[data-baseweb="tab-highlight"] { display: none !important; }

    /* 9. Sidebar, Tactile Glass Buttons & Expanders */
    section[data-testid="stSidebar"] {
        background-color: #09090B !important;
        border-right: 1px solid rgba(255, 255, 255, 0.07) !important;
    }
    div.stDownloadButton > button, div.stButton > button {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-weight: 600 !important;
        border-radius: 12px !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        border-top: 1px solid rgba(255, 255, 255, 0.22) !important;
        background: linear-gradient(180deg, rgba(38, 38, 44, 0.85) 0%, rgba(22, 22, 26, 0.95) 100%) !important;
        padding: 0.55rem 1.1rem !important;
        transition: all 0.22s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }
    div.stDownloadButton > button:hover, div.stButton > button:hover {
        transform: translateY(-1px);
        border-color: #0A84FF !important;
        box-shadow: 0 6px 20px rgba(10, 132, 255, 0.28) !important;
    }
    details[data-testid="stExpander"] {
        background: rgba(18, 18, 22, 0.7) !important;
        border: 1px solid rgba(255, 255, 255, 0.08) !important;
        border-radius: 14px !important;
    }
</style>
"""
BANNER_HTML = ("<div class='vectra-banner'><h1>VECTRA</h1>"
               "<p>Strategic FP&amp;A &amp; PVM Intelligence</p></div>")
BADGE_HTML = "<div class='vectra-badge'>Engineered &amp; Architected by <b>Dinu Denny Culas</b></div>"
FOOTER_HTML = "<div class='vectra-footer'>Engineered &amp; Architected by <b>Dinu Denny Culas</b></div>"

st.markdown(PAGE_CSS, unsafe_allow_html=True)


# ----------------------------------------------------------------------
# Formatting helpers
# ----------------------------------------------------------------------
def fmt_inr(value):
    """Compact Indian style for KPI cards: ₹1.26 Cr, ₹4.50 L, ₹850."""
    sign = ""
    if value < 0:
        sign = "-"
    size = abs(value)
    if size >= 1e7:
        return f"{sign}₹{size / 1e7:,.2f} Cr"
    if size >= 1e5:
        return f"{sign}₹{size / 1e5:,.2f} L"
    return f"{sign}₹{size:,.0f}"


def fmt_inr_full(value, decimals=0):
    """Full amount with Indian digit grouping: ₹12,58,500."""
    sign = ""
    if value < 0:
        sign = "-"
    whole, _, fraction = f"{abs(value):.{decimals}f}".partition(".")
    if len(whole) > 3:
        head, last_three = whole[:-3], whole[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ",".join(groups + [last_three])
    if fraction:
        whole = f"{whole}.{fraction}"
    return f"{sign}₹{whole}"


def plus_sign(value):
    """'+' for positive numbers (negative numbers already carry their own '-')."""
    if value > 0:
        return "+"
    return ""


def fmt_signed_inr(value):
    """Full amount with an explicit + for positive numbers."""
    return plus_sign(value) + fmt_inr_full(value)


def fmt_pct(value, decimals=1):
    """Percentage with sign, or n/a when the budget was zero."""
    if pd.isna(value):
        return "n/a"
    return f"{value:+.{decimals}f}%"


def py_num(value):
    """A number as it is written in Python code, with brackets around negatives: 1_827_400."""
    text = f"{value:_.12g}"
    if value < 0:
        return f"({text})"
    return text


def safe_divide(numerator, denominator):
    """EXCEL EQUIVALENT: =IFERROR(numerator / denominator, 0)"""
    if denominator == 0:
        return 0.0
    return numerator / denominator


def delta_text(variance, flag):
    """Text such as '+₹1.26 Cr (U)'. Streamlit colours the arrow by the first character."""
    return f"{plus_sign(variance)}{fmt_inr(variance)} ({flag})"


def style_flag(flag):
    """Table cell style: F is green, U is red."""
    if flag == "F":
        return "background-color: #E8F5E9; color: #1B5E20; font-weight: 600"
    if flag == "U":
        return "background-color: #FDECEA; color: #B71C1C; font-weight: 600"
    return ""


def compare_numbers(name, hand_value, engine_value, is_percent=False, tolerance=0.01):
    """One row of the verifier table: hand calculation versus the engine's value."""
    both_undefined = pd.isna(hand_value) and pd.isna(engine_value)
    matches = both_undefined or abs(hand_value - engine_value) <= tolerance
    if is_percent:
        hand_text, engine_text = fmt_pct(hand_value, 2), fmt_pct(engine_value, 2)
    else:
        hand_text, engine_text = fmt_inr_full(hand_value, 2), fmt_inr_full(engine_value, 2)
    return {"Check": name, "Hand calculation": hand_text, "Engine value": engine_text,
            "Match": "✅" if matches else "❌"}


def compare_text(name, hand_text, engine_text):
    """Same as compare_numbers, for F / U / On Track flags."""
    return {"Check": name, "Hand calculation": hand_text, "Engine value": engine_text,
            "Match": "✅" if hand_text == engine_text else "❌"}


def finish_chart(fig, height=420):
    """Give every chart a clean look with proper breathing room for 2-line axis labels."""
    fig.update_layout(
        template="plotly_dark",
        height=height,
        margin=dict(t=45, b=78, l=60, r=28),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Plus Jakarta Sans, sans-serif", color="#E2E8F0", size=11),
    )
    fig.update_xaxes(
        automargin=True,
        tickangle=0,
        tickfont=dict(size=10.5, color="#CBD5E1"),
        gridcolor="rgba(255,255,255,0.04)",
    )
    fig.update_yaxes(
        automargin=True,
        tickfont=dict(size=10.5, color="#CBD5E1"),
        gridcolor="rgba(255,255,255,0.06)",
    )


# ----------------------------------------------------------------------
# Cached loaders (the cache key is the file content, so new files refresh automatically)
# ----------------------------------------------------------------------
@st.cache_data
def get_opex_data(file_bytes, file_name):
    source = io.BytesIO(file_bytes)
    source.name = file_name
    return load_and_clean_opex(source)


@st.cache_data
def get_pvm_data(file_bytes, file_name):
    source = io.BytesIO(file_bytes)
    source.name = file_name
    return calculate_pvm_analysis(source)


@st.cache_data
def build_template_files():
    return {"opex_csv": get_template_csv_bytes("opex"), "opex_xlsx": get_template_excel_bytes("opex"),
            "sales_csv": get_template_csv_bytes("sales"), "sales_xlsx": get_template_excel_bytes("sales")}


def show_demo_downloads(key_prefix):
    """Download buttons for the raw, uncleaned demo files (so messy ERP rows can be inspected in Excel)."""
    if not DEMO_FILES_FOUND:
        st.caption("Demo files not found. Run `python data_generator.py` in this folder first.")
        return
    for path in [DEMO_OPEX_PATH, DEMO_SALES_PATH]:
        st.download_button(f"Download {path.name}", data=path.read_bytes(), file_name=path.name,
                           mime=CSV_MIME, key=f"{key_prefix}_demo_{path.name}")


def show_template_downloads(key_prefix):
    """Download buttons for the blank ERP templates (exact column headers plus example rows)."""
    templates = build_template_files()
    downloads = [
        ("OPEX template (.csv)", templates["opex_csv"], "erp_opex_template.csv", CSV_MIME),
        ("OPEX template (.xlsx)", templates["opex_xlsx"], "erp_opex_template.xlsx", XLSX_MIME),
        ("Sales PVM template (.csv)", templates["sales_csv"], "sales_pvm_template.csv", CSV_MIME),
        ("Sales PVM template (.xlsx)", templates["sales_xlsx"], "sales_pvm_template.xlsx", XLSX_MIME),
    ]
    for label, data, file_name, mime in downloads:
        st.download_button(label, data=data, file_name=file_name, mime=mime,
                           key=f"{key_prefix}_template_{file_name}")


# ----------------------------------------------------------------------
# Code snippets shown in Tab 4 (they mirror finance_engine.py and this file)
# ----------------------------------------------------------------------
SNIPPET_CLEANING = '''# 1) Trim text                         (Excel: =TRIM(B2))
df["Department"] = df["Department"].str.strip()

# 2) "₹1,25,000" -> 125000.0           (Excel: Find & Replace, then =VALUE())
for col in ["Budget_INR", "Actual_INR"]:
    text = df[col].astype(str).str.replace("₹", "", regex=False).str.replace(",", "", regex=False)
    df[col] = pd.to_numeric(text, errors="coerce")

# 3) Impute blanks                     (Excel: =IF(ISBLANK(Actual), Budget, Actual))
df["Actual_INR"] = df["Actual_INR"].fillna(df["Budget_INR"])
df["Cost_Type"] = df["Cost_Type"].fillna("Variable")'''

SNIPPET_VARIANCE = '''# Variance                             (Excel: =Actual - Budget)
df["Variance_INR"] = df["Actual_INR"] - df["Budget_INR"]

# Variance %, divided by |Budget|. A zero budget becomes NaN instead of a divide-by-zero error.
df["Variance_Pct"] = df["Variance_INR"] / df["Budget_INR"].abs().replace(0, np.nan) * 100

# Expense flag                         (Excel: =IF(Variance>0,"U",IF(Variance<0,"F","On Track")))
df["Variance_Flag"] = np.select([df["Variance_INR"] > 0, df["Variance_INR"] < 0],
                                ["U", "F"], default="On Track")

# Department totals: variance % is calculated on the TOTALS, never as an average of row percentages
summary = df.groupby(["Department", "Quarter"], as_index=False)[["Budget_INR", "Actual_INR"]].sum()
summary["Variance_INR"] = summary["Actual_INR"] - summary["Budget_INR"]
summary["Variance_Pct"] = summary["Variance_INR"] / summary["Budget_INR"].abs().replace(0, np.nan) * 100'''

SNIPPET_PVM = '''# Revenue                              (Excel: =Units * Price)
df["Budget_Revenue"] = df["Budget_Units"] * df["Budget_Price_INR"]
df["Actual_Revenue"] = df["Actual_Units"] * df["Actual_Price_INR"]
df["Total_Sales_Variance"] = df["Actual_Revenue"] - df["Budget_Revenue"]

# PRICE variance: the price gap, measured on the units actually sold
df["Price_Variance"] = (df["Actual_Price_INR"] - df["Budget_Price_INR"]) * df["Actual_Units"]

# VOLUME variance: the unit gap, measured at the budgeted price (so price changes cannot leak in)
df["Volume_Variance"] = (df["Actual_Units"] - df["Budget_Units"]) * df["Budget_Price_INR"]

# Revenue flag: higher than budget is Favorable
df["Price_Flag"] = np.select([df["Price_Variance"] > 0, df["Price_Variance"] < 0],
                             ["F", "U"], default="On Track")'''

SNIPPET_FLEX = '''volume_rate = volume_shock / 100
price_rate = price_shock / 100
inflation_rate = opex_inflation / 100

# Flexible budget switch: do Variable costs move with volume?
if flex_variable_costs:
    volume_flex_factor = 1 + volume_rate
else:
    volume_flex_factor = 1.0

# Revenue                              (Excel: =Units*(1+Vol%) * Price*(1+Price%))
df_sales["Sim_Units"] = df_sales["Actual_Units"] * (1 + volume_rate)
df_sales["Sim_Price"] = df_sales["Actual_Price_INR"] * (1 + price_rate)
df_sales["Sim_Revenue"] = df_sales["Sim_Units"] * df_sales["Sim_Price"]

# OPEX                                 (Excel: =IF(Cost_Type="Variable", Actual*(1+Vol%)*(1+Infl%), Actual))
df_opex["Sim_Actual_INR"] = np.where(df_opex["Cost_Type"] == "Variable",
                                     df_opex["Actual_INR"] * volume_flex_factor * (1 + inflation_rate),
                                     df_opex["Actual_INR"])           # Fixed rows never change

# Profit bridge: static budget -> flexible budget -> rate variance (each step is its profit impact)
volume_effect = baseline_revenue * volume_rate
price_effect = baseline_revenue * (1 + volume_rate) * price_rate
rate_variance = -(variable_opex * volume_flex_factor * inflation_rate)
if flex_variable_costs:
    volume_cost_change = -(variable_opex * volume_rate)
else:
    volume_cost_change = 0.0'''

# ======================================================================
# 2. WELCOME SCREEN
# ======================================================================
# The sidebar radio below is bound to the key "data_source_view". It is created AFTER this
# section, so the welcome cards are allowed to change that value before the radio exists.
st.session_state.setdefault("data_source_view", VIEW_WELCOME)

st.markdown(BANNER_HTML, unsafe_allow_html=True)

if st.session_state["data_source_view"] == VIEW_WELCOME:
    st.header("Welcome to VECTRA")
    st.write("Turn raw ERP exports into audited variance insight. Choose how you would like to start.")

    demo_card, custom_card = st.columns(2)

    with demo_card:
        with st.container(border=True):
            st.markdown("<div class='vectra-card-title'>Explore Live Demo Dataset</div>"
                        "<div class='vectra-card-text'>A realistic 2026 ERP extract: 144 operating-expense rows "
                        "with deliberately messy data, plus quarterly sales for three product tiers.</div>",
                        unsafe_allow_html=True)
            if st.button("Launch VECTRA with Demo ERP Data", type="primary", key="launch_demo",
                         disabled=not DEMO_FILES_FOUND):
                st.session_state["data_source_view"] = VIEW_DEMO
                st.rerun()
            st.markdown("**Download Demo Datasets**")
            st.caption("The exact raw, uncleaned files, so you can inspect the messy ERP rows in Excel yourself.")
            show_demo_downloads("welcome")

    with custom_card:
        with st.container(border=True):
            st.markdown("<div class='vectra-card-title'>Analyze Your Own ERP Export (SAP / Oracle / NetSuite)</div>"
                        "<div class='vectra-card-text'>Upload an OPEX ledger and a Sales PVM ledger (.csv or .xlsx). "
                        "VECTRA checks the headers, repairs common ERP data issues and opens the full "
                        "workspace.</div>", unsafe_allow_html=True)
            opex_upload = st.file_uploader("OPEX Ledger (.csv or .xlsx)", type=["csv", "xlsx"],
                                           key="welcome_opex_upload")
            sales_upload = st.file_uploader("Sales PVM Ledger (.csv or .xlsx)", type=["csv", "xlsx"],
                                            key="welcome_sales_upload")

            if opex_upload is not None and sales_upload is not None:
                try:
                    get_opex_data(opex_upload.getvalue(), opex_upload.name)
                    get_pvm_data(sales_upload.getvalue(), sales_upload.name)
                except ValueError as error:
                    st.error(f"{error}\n\nDownload a blank ERP template below to see the exact column structure.")
                else:
                    st.session_state["custom_files"] = {
                        "opex_bytes": opex_upload.getvalue(), "opex_name": opex_upload.name,
                        "sales_bytes": sales_upload.getvalue(), "sales_name": sales_upload.name,
                    }
                    st.session_state["data_source_view"] = VIEW_CUSTOM
                    st.rerun()

            st.markdown("**Download Blank ERP Templates**")
            st.caption("The exact column structure, with example rows. Replace them with your numbers and upload above.")
            show_template_downloads("welcome")

    if "custom_files" in st.session_state:
        earlier = st.session_state["custom_files"]
        loaded_names = f"{earlier['opex_name']} and {earlier['sales_name']}"
        st.info(f"Your earlier upload ({loaded_names}) is still loaded. "
                f"Select **{VIEW_CUSTOM}** in the sidebar to return to it, or upload new files above.")

    st.markdown(FOOTER_HTML, unsafe_allow_html=True)

# ======================================================================
# 3. SIDEBAR  (the data source switch and both download sets are always available)
# ======================================================================
st.sidebar.markdown(BADGE_HTML, unsafe_allow_html=True)
st.sidebar.divider()
active_view = st.sidebar.radio("Data Source Mode", VIEWS, key="data_source_view")
with st.sidebar.expander("Download Demo Datasets"):
    show_demo_downloads("sidebar")
with st.sidebar.expander("Download Blank ERP Templates"):
    show_template_downloads("sidebar")

if active_view == VIEW_WELCOME:
    st.stop()

# ======================================================================
# 4. DATA
# ======================================================================
if active_view == VIEW_DEMO:
    if not DEMO_FILES_FOUND:
        st.error("Demo CSV files not found. Run `python data_generator.py` in this folder first.")
        st.stop()
    opex_bytes, opex_name = DEMO_OPEX_PATH.read_bytes(), DEMO_OPEX_PATH.name
    sales_bytes, sales_name = DEMO_SALES_PATH.read_bytes(), DEMO_SALES_PATH.name
    source_label = "Demo ERP dataset"
else:
    custom_files = st.session_state.get("custom_files")
    if custom_files is None:
        st.info("No custom files loaded yet. Choose **Welcome Screen** in the sidebar, then upload your "
                "OPEX and Sales PVM ledgers.")
        st.stop()
    opex_bytes, opex_name = custom_files["opex_bytes"], custom_files["opex_name"]
    sales_bytes, sales_name = custom_files["sales_bytes"], custom_files["sales_name"]
    source_label = f"Uploaded files ({opex_name}, {sales_name})"

try:
    df_opex = get_opex_data(opex_bytes, opex_name)
    df_sales = get_pvm_data(sales_bytes, sales_name)
except ValueError as error:
    st.error(f"{error}\n\nDownload a blank ERP template from the sidebar to see the exact column structure.")
    st.stop()

st.sidebar.caption(f"Active data: {len(df_opex):,} OPEX rows and {len(df_sales):,} sales rows loaded.")

all_departments = sorted(df_opex["Department"].unique())
dept_colors = {}
for position, dept in enumerate(all_departments):
    dept_colors[dept] = DEPT_PALETTE[position % len(DEPT_PALETTE)]

quarters_in_data = sorted(set(df_opex["Quarter"]) | set(df_sales["Quarter_Short"]))
quarter_options = ["All Quarters"] + quarters_in_data
data_period_text = f"{df_opex['Date'].min():%b %Y} to {df_opex['Date'].max():%b %Y}"

# ----------------------------------------------------------------------
# Sidebar: filters and the What-If scenario controls
# ----------------------------------------------------------------------
st.sidebar.divider()
st.sidebar.header("Filters")
selected_quarter = st.sidebar.selectbox("Quarter", quarter_options)
selected_departments = st.sidebar.multiselect("Department", all_departments, default=all_departments)
st.sidebar.caption("The Department filter applies to OPEX only (the sales ledger has no department field).")

st.sidebar.divider()
st.sidebar.subheader("What-If Scenario Simulator")
volume_shock = st.sidebar.slider("Volume Growth Shock (%)", min_value=-30, max_value=30, value=0)
price_shock = st.sidebar.slider("Price Adjustment (%)", min_value=-20, max_value=20, value=0)
opex_inflation = st.sidebar.slider("Variable OPEX Inflation (%)", min_value=-20, max_value=30, value=0)
flex_variable_costs = st.sidebar.checkbox("Flex Variable Costs with Volume", value=True)
st.sidebar.caption("Flexible Budget Modeling. Checked: Variable OPEX = Actual x (1 + volume) x (1 + inflation). "
                   "Unchecked: volume ignores costs (static budget). Fixed OPEX never changes.")

if len(selected_departments) == 0:
    st.warning("Select at least one department in the sidebar to see the dashboard.")
    st.stop()

# ----------------------------------------------------------------------
# Scenario columns, calculated on the full data and filtered afterwards
# ----------------------------------------------------------------------
volume_rate = volume_shock / 100
price_rate = price_shock / 100
inflation_rate = opex_inflation / 100

# Flexible budget switch: do Variable costs move with volume?
if flex_variable_costs:
    volume_flex_factor = 1 + volume_rate
else:
    volume_flex_factor = 1.0

# EXCEL EQUIVALENT: =Units*(1+Vol%) * Price*(1+Price%)
df_sales["Sim_Units"] = df_sales["Actual_Units"] * (1 + volume_rate)
df_sales["Sim_Price"] = df_sales["Actual_Price_INR"] * (1 + price_rate)
df_sales["Sim_Revenue"] = df_sales["Sim_Units"] * df_sales["Sim_Price"]

# EXCEL EQUIVALENT: =IF(Cost_Type="Variable", Actual*(1+Vol%)*(1+Infl%), Actual)
df_opex["Sim_Actual_INR"] = np.where(df_opex["Cost_Type"] == "Variable",
                                     df_opex["Actual_INR"] * volume_flex_factor * (1 + inflation_rate),
                                     df_opex["Actual_INR"])

# ----------------------------------------------------------------------
# Apply the filters (EXCEL EQUIVALENT: slicers on a Pivot Table)
# ----------------------------------------------------------------------
opex_view = df_opex[df_opex["Department"].isin(selected_departments)]
sales_view = df_sales
period_label = f"Full period ({data_period_text})"
quarter_dates = pd.Series(dtype="datetime64[ns]")

if selected_quarter != "All Quarters":
    opex_view = opex_view[opex_view["Quarter"] == selected_quarter]
    sales_view = df_sales[df_sales["Quarter_Short"] == selected_quarter]
    quarter_dates = df_opex.loc[df_opex["Quarter"] == selected_quarter, "Date"]
    period_label = selected_quarter
    if not quarter_dates.empty:
        period_label = f"{selected_quarter} ({quarter_dates.min():%b} to {quarter_dates.max():%b %Y})"

# ======================================================================
# 5. DASHBOARD HEADER AND KPI ROW
# ======================================================================
st.caption(f"Period: {period_label}  |  Data: {source_label}  |  Amounts in INR "
           "(₹ Cr = crore, ₹ L = lakh)  |  F = Favorable, U = Unfavorable")

if opex_view.empty or sales_view.empty:
    st.warning(f"No data found for {selected_quarter} in the active dataset "
               "(one of the two ledgers has no rows for it). Pick another quarter in the sidebar.")
    st.stop()

# EXCEL EQUIVALENT: =SUM() of the filtered Budget and Actual columns
revenue_budget = sales_view["Budget_Revenue"].sum()
revenue_actual = sales_view["Actual_Revenue"].sum()
revenue_variance = revenue_actual - revenue_budget

opex_budget = opex_view["Budget_INR"].sum()
opex_actual = opex_view["Actual_INR"].sum()
opex_variance = opex_actual - opex_budget

# EXCEL EQUIVALENT: =Revenue - OPEX   and   =Net / Revenue
net_actual = revenue_actual - opex_actual
net_budget = revenue_budget - opex_budget
net_variance = net_actual - net_budget
margin_actual = safe_divide(net_actual, revenue_actual) * 100
margin_budget = safe_divide(net_budget, revenue_budget) * 100

# Data audit: rows where the cleaning step wrote a note (Data_Fix_Note in finance_engine.py)
rows_fixed = int((opex_view["Data_Fix_Note"] != "").sum())
rows_in_view = len(opex_view)

kpi_revenue, kpi_opex, kpi_margin, kpi_audit = st.columns(4)

with kpi_revenue:
    # Revenue above budget is Favorable: green when up, red when down
    revenue_flag = variance_flag(revenue_variance, is_expense=False)
    st.metric("Total Actual Revenue", fmt_inr(revenue_actual),
              delta=delta_text(revenue_variance, revenue_flag) + " vs Budget", delta_color="normal")

with kpi_opex:
    # Expense above budget is Unfavorable: "inverse" turns an increase red and a decrease green
    opex_flag = variance_flag(opex_variance, is_expense=True)
    st.metric("Total Actual OPEX", fmt_inr(opex_actual),
              delta=delta_text(opex_variance, opex_flag) + " vs Budget", delta_color="inverse")

with kpi_margin:
    net_flag = variance_flag(net_variance, is_expense=False)
    st.metric("Net Operating Margin", fmt_inr(net_actual),
              delta=delta_text(net_variance, net_flag) + " vs Budget", delta_color="normal")
    st.caption(f"Margin: {margin_actual:.1f}% of revenue (Budget: {margin_budget:.1f}%)")

with kpi_audit:
    if rows_fixed == 0:
        audit_value = "All clean"
    else:
        audit_value = f"{rows_fixed} rows auto-cleaned"
    audit_share = safe_divide(rows_fixed, rows_in_view) * 100
    st.metric("Data Audit Status", audit_value,
              delta=f"{audit_share:.1f}% of {rows_in_view} ERP rows in view", delta_color="off",
              help="Rows repaired by load_and_clean_opex(): trimmed spaces, text amounts converted to numbers, "
                   "missing Actuals filled with Budget, missing Cost_Type set to Variable, blank names filled.")

if len(selected_departments) < len(all_departments):
    st.caption("Note: revenue is company-wide, so with a department subset selected the margin compares "
               "ALL revenue with the selected departments' OPEX only.")

tab_opex, tab_pvm, tab_whatif, tab_audit = st.tabs([
    "1. OPEX Variance & Run-Rate",
    "2. Price-Volume-Mix (PVM) Bridge",
    "3. What-If Scenario Simulator",
    "4. Calculation Engine & Formula Audit",
])

# ======================================================================
# TAB 1: OPEX VARIANCE & RUN-RATE
# ======================================================================
with tab_opex:
    left_col, right_col = st.columns(2)

    # ------------------------------------------------------------------
    # Left: Budget vs Actual by department. The Actual bar is red (U) or green (F).
    # ------------------------------------------------------------------
    with left_col:
        st.subheader("Budget vs Actual OPEX by Department")

        # EXCEL EQUIVALENT: Pivot Table, Rows = Department, Values = Sum of Budget and Actual
        dept_totals = opex_view.groupby("Department", as_index=False)[["Budget_INR", "Actual_INR"]].sum()
        dept_totals["Variance_INR"] = dept_totals["Actual_INR"] - dept_totals["Budget_INR"]
        budget_magnitude = dept_totals["Budget_INR"].abs().replace(0, np.nan)
        dept_totals["Variance_Pct"] = dept_totals["Variance_INR"] / budget_magnitude * 100
        dept_totals["Variance_Flag"] = variance_flag(dept_totals["Variance_INR"], is_expense=True)

        is_unfavorable = dept_totals["Variance_Flag"] == "U"
        is_favorable = dept_totals["Variance_Flag"] == "F"
        actual_colors = np.select([is_unfavorable, is_favorable], [RED, GREEN], default=GREY)
        bar_labels = dept_totals["Variance_Pct"].map(fmt_pct) + " (" + dept_totals["Variance_Flag"] + ")"
        axis_labels = dept_totals["Department"].str.wrap(12, break_long_words=False).str.replace(
            "\n", "<br>", regex=False)

        fig_dept = go.Figure()
        fig_dept.add_trace(go.Bar(x=axis_labels, y=dept_totals["Budget_INR"] / 1e5, name="Budget",
                                  marker_color=GREY,
                                  hovertemplate="%{x}<br>Budget: ₹%{y:,.1f} L<extra></extra>"))
        fig_dept.add_trace(go.Bar(x=axis_labels, y=dept_totals["Actual_INR"] / 1e5,
                                  name="Actual (green = F, red = U)", marker_color=actual_colors,
                                  text=bar_labels, textposition="outside",
                                  hovertemplate="%{x}<br>Actual: ₹%{y:,.1f} L<extra></extra>"))
        tallest_bar = max(dept_totals["Budget_INR"].max(), dept_totals["Actual_INR"].max()) / 1e5
        fig_dept.update_layout(barmode="group", yaxis_title="₹ Lakh", yaxis_range=[0, tallest_bar * 1.2],
                               legend=dict(orientation="h", y=1.02, yanchor="bottom", x=0))
        finish_chart(fig_dept)
        st.plotly_chart(fig_dept)
        st.caption(":green[■ Favorable (F): spent less than budget]  ·  "
                   ":red[■ Unfavorable (U): spent more than budget]")

    # ------------------------------------------------------------------
    # Right: monthly run-rate for every month in the data
    # ------------------------------------------------------------------
    with right_col:
        st.subheader("Monthly Run-Rate: Budget vs Actual")

        # EXCEL EQUIVALENT: Pivot Table, Rows = Date, Columns = Department, Values = Sum of Actual
        trend_source = df_opex[df_opex["Department"].isin(selected_departments)]
        monthly = trend_source.groupby(["Date", "Department"], as_index=False)[["Budget_INR", "Actual_INR"]].sum()

        fig_trend = go.Figure()
        for dept in selected_departments:
            dept_months = monthly[monthly["Department"] == dept]
            line_color = dept_colors[dept]
            # The dotted Budget line has no legend entry of its own: it shares a legend group
            # with the department's solid Actual line, so one click toggles both.
            fig_trend.add_trace(go.Scatter(x=dept_months["Date"], y=dept_months["Budget_INR"] / 1e5,
                                           mode="lines", name=f"{dept} (Budget)", legendgroup=dept,
                                           showlegend=False, line=dict(color=line_color, width=2, dash="dot"),
                                           hovertemplate="%{x|%b %Y}<br>Budget: ₹%{y:,.1f} L<extra></extra>"))
            fig_trend.add_trace(go.Scatter(x=dept_months["Date"], y=dept_months["Actual_INR"] / 1e5,
                                           mode="lines+markers", name=dept, legendgroup=dept,
                                           line=dict(color=line_color, width=3),
                                           hovertemplate="%{x|%b %Y}<br>Actual: ₹%{y:,.1f} L<extra></extra>"))

        if not quarter_dates.empty:
            shade_start = quarter_dates.min()
            shade_end = quarter_dates.max() + pd.DateOffset(months=1)
            fig_trend.add_vrect(x0=shade_start, x1=shade_end, fillcolor=GREY, opacity=0.2,
                                line_width=0, layer="below")

        fig_trend.update_layout(yaxis_title="₹ Lakh per month",
                                xaxis=dict(tickformat="%b<br>%Y", dtick="M1", tickfont=dict(size=10)),
                                legend=dict(orientation="h", y=1.02, yanchor="bottom", x=0))
        finish_chart(fig_trend)
        st.plotly_chart(fig_trend)
        st.caption("Dotted = Budget, solid = Actual. Every month is shown so spikes stay visible; "
                   "the selected quarter is shaded.")

    # ------------------------------------------------------------------
    # Bottom: department summary and the data-quality log
    # ------------------------------------------------------------------
    st.subheader("Cleaned Department Summary")
    dept_summary = get_department_summary(opex_view)
    summary_table = dept_summary.style.format({
        "Budget_INR": fmt_inr_full, "Actual_INR": fmt_inr_full,
        "Variance_INR": fmt_signed_inr, "Variance_Pct": fmt_pct,
    }).map(style_flag, subset=["Variance_Flag"])
    st.dataframe(summary_table, hide_index=True)

    with st.expander(f"ERP data quality fixes ({rows_fixed} rows auto-cleaned in this view)"):
        fixed_rows = opex_view[opex_view["Data_Fix_Note"] != ""].copy()
        if fixed_rows.empty:
            st.write("No data quality fixes were needed for the current selection.")
        else:
            st.caption("Cleaned values are shown. Your original file is never modified.")
            fixed_rows["Date"] = fixed_rows["Date"].dt.strftime("%b %Y")
            fix_columns = ["Date", "Department", "Account_Code", "Expense_Line_Item", "Cost_Type",
                           "Budget_INR", "Actual_INR", "Data_Fix_Note"]
            st.dataframe(fixed_rows[fix_columns].style.format({"Budget_INR": fmt_inr_full,
                                                               "Actual_INR": fmt_inr_full}),
                         hide_index=True)

# ======================================================================
# TAB 2: PRICE-VOLUME-MIX (PVM) BRIDGE
# ======================================================================
with tab_pvm:
    pvm_quarter_options = ["All Quarters"] + sorted(df_sales["Quarter_Short"].unique())
    default_index = 0
    if "Q3" in pvm_quarter_options:
        default_index = pvm_quarter_options.index("Q3")
    pvm_quarter = st.selectbox("Select Quarter", pvm_quarter_options, index=default_index, key="pvm_quarter")

    pvm_view = df_sales
    if pvm_quarter != "All Quarters":
        pvm_view = df_sales[df_sales["Quarter_Short"] == pvm_quarter]

    if pvm_view.empty:
        st.warning("No sales data for this quarter.")
    else:
        # EXCEL EQUIVALENT: Pivot Table, Rows = Product_Tier, Values = Sum of each variance column
        sum_columns = ["Budget_Units", "Actual_Units", "Budget_Revenue", "Actual_Revenue",
                       "Price_Variance", "Volume_Variance", "Total_Sales_Variance"]
        pvm_by_tier = pvm_view.groupby("Product_Tier", sort=False)[sum_columns].sum().reset_index()

        st.subheader(f"Price vs Volume vs Total Sales Variance by Product Tier ({pvm_quarter})")

        # One bar per (tier, variance type), tier by tier: stack() turns the three variance
        # columns into rows, keeping each tier's Price, Volume and Total together.
        variance_columns = ["Price_Variance", "Volume_Variance", "Total_Sales_Variance"]
        bars = pvm_by_tier.set_index("Product_Tier")[variance_columns].stack().reset_index()
        bars.columns = ["Product_Tier", "Variance_Type", "Variance_INR"]
        bars["Variance_Type"] = bars["Variance_Type"].map(
            {"Price_Variance": "Price", "Volume_Variance": "Volume", "Total_Sales_Variance": "Total"})
        bars["Variance_Flag"] = variance_flag(bars["Variance_INR"], is_expense=False)

        is_favorable = bars["Variance_Flag"] == "F"
        is_unfavorable = bars["Variance_Flag"] == "U"
        bar_colors = np.select([is_favorable, is_unfavorable], [GREEN, RED], default=GREY)
        bar_texts = bars["Variance_INR"].map(fmt_inr) + " (" + bars["Variance_Flag"] + ")"

        # Leave headroom above and below so the value labels never touch the axis
        lowest = min(0, bars["Variance_INR"].min() / 1e5)
        highest = max(0, bars["Variance_INR"].max() / 1e5)
        spread = max(highest - lowest, 1e-9)
        axis_bottom = 0
        if lowest < 0:
            axis_bottom = lowest - 0.18 * spread

        fig_pvm = go.Figure(go.Bar(x=[bars["Product_Tier"], bars["Variance_Type"]],
                                   y=bars["Variance_INR"] / 1e5, marker_color=bar_colors,
                                   text=bar_texts, textposition="outside", cliponaxis=False,
                                   hovertemplate="%{x}<br>₹%{y:,.1f} L<extra></extra>"))
        fig_pvm.update_layout(yaxis_title="₹ Lakh", showlegend=False,
                              yaxis_range=[axis_bottom, highest + 0.15 * spread])
        fig_pvm.add_hline(y=0, line_color="#546E7A", line_width=1)
        finish_chart(fig_pvm, height=460)
        st.plotly_chart(fig_pvm)
        st.caption(":green[■ Favorable (F): revenue above budget]  ·  :red[■ Unfavorable (U): revenue below budget]  ·  "
                   "Price = (Actual Price - Budget Price) x Actual Units  ·  "
                   "Volume = (Actual Units - Budget Units) x Budget Price")

        # --------------------------------------------------------------
        # Callout: how volume can mask a price problem (Enterprise License if present)
        # --------------------------------------------------------------
        if "Enterprise License" in pvm_by_tier["Product_Tier"].tolist():
            spotlight_tier = "Enterprise License"
        else:
            biggest_total = pvm_by_tier["Total_Sales_Variance"].abs().idxmax()
            spotlight_tier = pvm_by_tier.loc[biggest_total, "Product_Tier"]
        spotlight = pvm_by_tier[pvm_by_tier["Product_Tier"] == spotlight_tier].iloc[0]

        price_var = spotlight["Price_Variance"]
        volume_var = spotlight["Volume_Variance"]
        total_var = spotlight["Total_Sales_Variance"]
        total_flag = variance_flag(total_var, is_expense=False)

        # Unit-weighted average price.  EXCEL EQUIVALENT: =SUM(Revenue) / SUM(Units)
        actual_price = safe_divide(spotlight["Actual_Revenue"], spotlight["Actual_Units"])
        budget_price = safe_divide(spotlight["Budget_Revenue"], spotlight["Budget_Units"])
        price_change_pct = (safe_divide(actual_price, budget_price) - 1) * 100
        unit_change_pct = (safe_divide(spotlight["Actual_Units"], spotlight["Budget_Units"]) - 1) * 100

        if price_var < 0 and volume_var > 0:
            erosion_pct = safe_divide(abs(price_var), volume_var) * 100
            if total_var > 0:
                headline = "Volume is masking a pricing problem."
                verdict = (f"Total sales look healthy at **{fmt_inr(total_var)} (F)**, but only because "
                           "extra units are hiding the discounting.")
            else:
                headline = "Discounting has overwhelmed the volume gain."
                verdict = f"Total sales are **{fmt_inr(total_var)} (U)** despite the extra units."
            st.warning(
                f"**{spotlight_tier}, {pvm_quarter}: {headline}**\n\n"
                f"Units sold were **{unit_change_pct:+.1f}%** vs budget, a Volume variance of "
                f"**{fmt_inr(volume_var)} (F)**. But the realised price was **{price_change_pct:+.1f}%** vs budget, "
                f"a Price variance of **{fmt_inr(price_var)} (U)**. "
                f"Price erosion consumed **{erosion_pct:.0f}%** of the volume gain. {verdict}")
        elif price_var > 0 and volume_var < 0:
            st.warning(
                f"**{spotlight_tier}, {pvm_quarter}: a price premium is masking falling volume.**\n\n"
                f"Units sold were **{unit_change_pct:+.1f}%** vs budget (Volume variance "
                f"**{fmt_inr(volume_var)} (U)**), while the realised price was **{price_change_pct:+.1f}%** vs "
                f"budget (Price variance **{fmt_inr(price_var)} (F)**). "
                f"Total sales variance: **{fmt_inr(total_var)} ({total_flag})**.")
        else:
            price_flag_now = variance_flag(price_var, is_expense=False)
            volume_flag_now = variance_flag(volume_var, is_expense=False)
            st.info(
                f"**{spotlight_tier}, {pvm_quarter}:** Price variance is **{fmt_inr(price_var)} "
                f"({price_flag_now})** and Volume variance is **{fmt_inr(volume_var)} ({volume_flag_now})**, "
                f"for a total of **{fmt_inr(total_var)}**. No price-versus-volume masking in this period.")

        # Detail table
        pvm_table = pvm_by_tier[["Product_Tier", "Budget_Revenue", "Actual_Revenue", "Price_Variance",
                                 "Volume_Variance", "Total_Sales_Variance"]].copy()
        pvm_table["Price_Flag"] = variance_flag(pvm_table["Price_Variance"], is_expense=False)
        pvm_table["Volume_Flag"] = variance_flag(pvm_table["Volume_Variance"], is_expense=False)
        pvm_table["Total_Flag"] = variance_flag(pvm_table["Total_Sales_Variance"], is_expense=False)
        table_order = ["Product_Tier", "Budget_Revenue", "Actual_Revenue", "Price_Variance", "Price_Flag",
                       "Volume_Variance", "Volume_Flag", "Total_Sales_Variance", "Total_Flag"]
        pvm_styled = pvm_table[table_order].style.format({
            "Budget_Revenue": fmt_inr_full, "Actual_Revenue": fmt_inr_full, "Price_Variance": fmt_signed_inr,
            "Volume_Variance": fmt_signed_inr, "Total_Sales_Variance": fmt_signed_inr,
        }).map(style_flag, subset=["Price_Flag", "Volume_Flag", "Total_Flag"])
        st.dataframe(pvm_styled, hide_index=True)

# ======================================================================
# TAB 3: WHAT-IF SCENARIO SIMULATOR  (Flexible Budget Modeling)
# ======================================================================
with tab_whatif:
    st.subheader("What-If Scenario Simulator")
    st.caption(f"Scenario applies to: {period_label}  |  {len(selected_departments)} department(s)  |  "
               "Change the sliders in the sidebar and everything below recalculates.")
    if flex_variable_costs:
        st.caption("Mode: **Flexible Budget** (Variable OPEX flexes with volume, then takes inflation). "
                   "Fixed OPEX is unchanged.")
    else:
        st.caption("Mode: **Static Budget** (volume changes revenue only; costs ignore volume). "
                   "Tick the sidebar box to switch to a flexible budget.")

    # The Sim_ columns were calculated in section 4.  EXCEL EQUIVALENT: =SUM(Revenue) - SUM(OPEX)
    baseline_revenue = sales_view["Actual_Revenue"].sum()
    baseline_opex = opex_view["Actual_INR"].sum()
    baseline_profit = baseline_revenue - baseline_opex

    simulated_revenue = sales_view["Sim_Revenue"].sum()
    simulated_opex = opex_view["Sim_Actual_INR"].sum()
    simulated_profit = simulated_revenue - simulated_opex

    profit_change = simulated_profit - baseline_profit
    baseline_margin = safe_divide(baseline_profit, baseline_revenue) * 100
    simulated_margin = safe_divide(simulated_profit, simulated_revenue) * 100
    margin_change = simulated_margin - baseline_margin

    # EXCEL EQUIVALENT: =SUMIFS(Actual, Cost_Type, "Fixed") and =SUMIFS(Actual, Cost_Type, "Variable")
    is_fixed = opex_view["Cost_Type"] == "Fixed"
    is_variable = opex_view["Cost_Type"] == "Variable"
    fixed_opex = opex_view.loc[is_fixed, "Actual_INR"].sum()
    variable_opex = opex_view.loc[is_variable, "Actual_INR"].sum()

    profit_delta_color = "normal"
    if profit_change == 0:
        profit_delta_color = "off"          # grey when nothing has changed
    margin_delta_color = "normal"
    if abs(margin_change) < 0.05:
        margin_delta_color = "off"

    metric_baseline, metric_simulated, metric_margin = st.columns(3)
    with metric_baseline:
        st.metric("Baseline Net Profit", fmt_inr(baseline_profit))
        st.caption(f"Margin: {baseline_margin:.1f}%")
    with metric_simulated:
        profit_flag = variance_flag(profit_change, is_expense=False)
        st.metric("Simulated Net Profit", fmt_inr(simulated_profit),
                  delta=delta_text(profit_change, profit_flag) + " vs Baseline", delta_color=profit_delta_color)
    with metric_margin:
        st.metric("Simulated Margin", f"{simulated_margin:.1f}%",
                  delta=f"{margin_change:+.1f} pp vs Baseline", delta_color=margin_delta_color)

    st.caption(f"Cost behaviour in this view: Fixed OPEX {fmt_inr(fixed_opex)} (never changes)  |  "
               f"Variable OPEX {fmt_inr(variable_opex)} (moves with the sliders).")

    # ------------------------------------------------------------------
    # Profit bridge: static budget -> flexible budget -> rate variance.
    # Each step is shown as its IMPACT ON PROFIT (more revenue = +, more cost = -).
    # ------------------------------------------------------------------
    volume_effect = baseline_revenue * volume_rate
    price_effect = baseline_revenue * (1 + volume_rate) * price_rate       # price applies to the NEW volume
    rate_variance = -(variable_opex * volume_flex_factor * inflation_rate)  # input prices, on the NEW volume
    if flex_variable_costs:
        volume_cost_change = -(variable_opex * volume_rate)
    else:
        volume_cost_change = 0.0

    bridge_labels = ["Baseline<br>Net Profit", "Volume Shock<br>(Revenue)", "Price Adjustment<br>(Revenue)",
                     "Volume-Driven<br>Variable Cost Change", "Rate/Inflation<br>OPEX Variance",
                     "Simulated<br>Net Profit"]
    bridge_steps = [volume_effect, price_effect, volume_cost_change, rate_variance]
    bridge_values = [baseline_profit] + bridge_steps + [simulated_profit]

    bridge_texts = [fmt_inr(baseline_profit)]
    for step in bridge_steps:
        bridge_texts.append(plus_sign(step) + fmt_inr(step))
    bridge_texts.append(fmt_inr(simulated_profit))

    fig_bridge = go.Figure(go.Waterfall(
        x=bridge_labels,
        y=[value / 1e5 for value in bridge_values],
        measure=["absolute", "relative", "relative", "relative", "relative", "total"],
        text=bridge_texts, textposition="outside",
        increasing=dict(marker=dict(color=GREEN)),       # adds profit = Favorable
        decreasing=dict(marker=dict(color=RED)),         # reduces profit = Unfavorable
        totals=dict(marker=dict(color=NAVY)),
        connector=dict(line=dict(color=GREY)),
        hovertemplate="%{x}<br>₹%{y:,.1f} L<extra></extra>",
    ))
    fig_bridge.update_layout(yaxis_title="₹ Lakh", showlegend=False)
    finish_chart(fig_bridge, height=480)
    st.plotly_chart(fig_bridge)
    st.caption(":green[■ Adds to profit (F)]  ·  :red[■ Reduces profit (U)]  ·  :blue[■ Profit totals]  ·  "
               "Static to flexible budget = the volume-driven cost step; flexible budget to actual = the "
               "rate/inflation step.")

    # Reconciliation: the four steps must add up exactly to the simulated profit
    bridge_gap = abs(baseline_profit + sum(bridge_steps) - simulated_profit)
    if bridge_gap < 1:
        st.caption("Bridge check: the four steps reconcile exactly to Simulated Net Profit.")
    else:
        st.error(f"Bridge does not reconcile (gap of {fmt_inr(bridge_gap)}). Please review the formulas.")

# ======================================================================
# TAB 4: CALCULATION ENGINE & FORMULA AUDIT
# ======================================================================
with tab_audit:
    st.header("Under the Hood: Python & Enterprise Variance Accounting Engine")
    st.caption("Every number in this dashboard comes from the formulas below. Read the math, read the code, "
               "then use the verifier to prove it on any single row.")

    audit_cleaning, audit_variance, audit_pvm, audit_flex, audit_verify = st.tabs([
        "A. ERP Cleaning & Imputation",
        "B. OPEX Variance & F/U Rules",
        "C. Price-Volume-Mix (PVM) Decomposition",
        "D. Flexible Budget Modeling",
        "E. Interactive Verifier",
    ])

    # ------------------------------------------------------------------
    # A. Cleaning and imputation
    # ------------------------------------------------------------------
    with audit_cleaning:
        st.subheader("A. ERP Data Cleaning & Imputation")
        st.markdown("Raw ERP exports contain text amounts, stray spaces and blanks. The engine repairs them "
                    "**before** any total is calculated and writes a note on every repaired row.")
        st.latex(r"""
        \text{Actual}_{\text{clean}} =
        \begin{cases}
        \text{Actual}_{\text{reported}} & \text{if the cell has a value} \\
        \text{Budget} & \text{if the cell is blank (imputed)}
        \end{cases}
        """)
        st.code(SNIPPET_CLEANING, language="python")

        st.markdown("**How budget-imputation affects variance.** An imputed row has Actual = Budget, so its "
                    "variance is exactly zero. If the real spend was higher, the reported variance is understated:")
        st.latex(r"""
        \text{Variance}_{\text{imputed row}} = \text{Budget} - \text{Budget} = 0
        \quad\Rightarrow\quad
        \text{Reported variance} = \text{True variance} - \text{True variance}_{\text{of imputed rows}}
        """)

        was_imputed = df_opex["Data_Fix_Note"].str.contains("Missing Actual filled with Budget", regex=False)
        imputed_rows = int(was_imputed.sum())
        imputed_amount = df_opex.loc[was_imputed, "Actual_INR"].sum()        # equals their Budget
        reported_variance = df_opex["Variance_INR"].sum()

        if imputed_rows == 0:
            st.success("The active dataset has no imputed Actuals, so this assumption does not affect any variance.")
        else:
            reported_flag = variance_flag(reported_variance, is_expense=True)
            st.info(f"In the active dataset (all periods, all departments), **{imputed_rows} row(s)** worth "
                    f"**{fmt_inr_full(imputed_amount)}** were imputed at zero variance. "
                    f"Reported total OPEX variance: **{fmt_inr_full(reported_variance)} ({reported_flag})**.")
            sensitivity_pct = st.slider("Sensitivity: what if the imputed rows really ran this % above (+) "
                                        "or below (-) budget?", min_value=-20, max_value=40, value=10,
                                        key="impute_sensitivity")
            hidden_variance = imputed_amount * sensitivity_pct / 100
            corrected_variance = reported_variance + hidden_variance
            corrected_flag = variance_flag(corrected_variance, is_expense=True)
            sens_hidden, sens_reported, sens_corrected = st.columns(3)
            sens_hidden.metric("Variance hidden by imputation", fmt_inr_full(hidden_variance))
            sens_reported.metric("Reported variance", fmt_inr_full(reported_variance))
            sens_corrected.metric("Variance if rows ran as assumed", fmt_inr_full(corrected_variance),
                                  delta=f"{corrected_flag} (expense rule)", delta_color="off")
            st.caption("Hidden variance = imputed amount x sensitivity %.  "
                       "Corrected variance = reported variance + hidden variance.")

    # ------------------------------------------------------------------
    # B. Variance and F/U rules
    # ------------------------------------------------------------------
    with audit_variance:
        st.subheader("B. OPEX Variance & Favorable / Unfavorable Rules")
        st.latex(r"\text{Variance (INR)} = \text{Actual} - \text{Budget}")
        st.latex(r"\text{Variance}\% = \frac{\text{Actual} - \text{Budget}}{\left|\text{Budget}\right|} \times 100"
                 r"\qquad (\text{undefined, shown as n/a, when Budget} = 0)")
        st.latex(r"""
        \text{Flag}_{\text{expense}} =
        \begin{cases}
        \text{U (Unfavorable)} & \text{if Actual} > \text{Budget} \\
        \text{F (Favorable)} & \text{if Actual} < \text{Budget} \\
        \text{On Track} & \text{if Actual} = \text{Budget}
        \end{cases}
        """)
        st.markdown("Revenue and profit use the **opposite** rule: Actual above Budget is Favorable (F).")
        st.latex(r"""
        \text{Dept Variance}\% = \frac{\sum \text{Actual} - \sum \text{Budget}}{\left|\sum \text{Budget}\right|} \times 100
        \qquad (\text{totals first; never the average of row percentages})
        """)
        st.code(SNIPPET_VARIANCE, language="python")

    # ------------------------------------------------------------------
    # C. Price-Volume-Mix decomposition
    # ------------------------------------------------------------------
    with audit_pvm:
        st.subheader("C. Price-Volume-Mix (PVM) Decomposition")
        st.latex(r"\text{Price Variance} = (P_{\text{actual}} - P_{\text{budget}}) \times Q_{\text{actual}}")
        st.latex(r"\text{Volume Variance} = (Q_{\text{actual}} - Q_{\text{budget}}) \times P_{\text{budget}}")
        st.latex(r"""
        \text{Total Sales Variance} = P_A Q_A - P_B Q_B
        = \underbrace{(P_A - P_B)\,Q_A}_{\text{Price}} + \underbrace{(Q_A - Q_B)\,P_B}_{\text{Volume}}
        """)
        st.markdown("Price is measured on the units actually sold; Volume at the budgeted price. That is why the "
                    "two add up exactly to the total with nothing left over. (A separate **Mix** variance only "
                    "appears when several products are combined into one total.)")
        st.code(SNIPPET_PVM, language="python")

    # ------------------------------------------------------------------
    # D. Flexible budget modeling
    # ------------------------------------------------------------------
    with audit_flex:
        st.subheader("D. Flexible Budget Modeling")
        flex_status = "OFF"
        if flex_variable_costs:
            flex_status = "ON"
        st.markdown(f"Your current settings: volume **{volume_shock:+d}%**, price **{price_shock:+d}%**, "
                    f"variable OPEX inflation **{opex_inflation:+d}%**, flexible budget **{flex_status}**.")
        st.latex(r"\text{Revenue}_{sim} = \big[Q_A\,(1+v)\big] \times \big[P_A\,(1+p)\big]")
        st.latex(r"\text{Variable OPEX}_{sim} = \text{Variable OPEX}_{actual} \times (1+v)^{f} \times (1+i)"
                 r"\qquad f = 1 \text{ if Flex Variable Costs is ON, else } 0")
        st.latex(r"\text{Fixed OPEX}_{sim} = \text{Fixed OPEX}_{actual}")
        st.latex(r"\text{Profit}_{sim} = \text{Revenue}_{sim} - \text{Fixed OPEX}_{sim} - \text{Variable OPEX}_{sim}")
        st.markdown("**Profit bridge** (static budget, flexible budget, then rate variance):")
        st.latex(r"""
        \Delta\text{Profit} =
        \underbrace{R_0\,v}_{\text{volume (revenue)}}
        + \underbrace{R_0(1+v)\,p}_{\text{price (revenue)}}
        - \underbrace{V_0\,v\,f}_{\text{volume-driven cost}}
        - \underbrace{V_0(1+v)^{f}\,i}_{\text{rate / inflation}}
        """)
        st.markdown("Static budget = baseline cost at baseline volume. Flexible budget = the same cost behaviour at "
                    "the new volume (variable cost per unit stays constant, fixed cost stays flat). "
                    "**Volume-driven change** = flexible minus static. **Rate/inflation variance** = simulated "
                    "actual minus flexible.")
        st.code(SNIPPET_FLEX, language="python")

    # ------------------------------------------------------------------
    # E. Interactive single-row verifier
    # ------------------------------------------------------------------
    with audit_verify:
        st.subheader("E. Interactive Single-Row Verifier")
        st.caption("Pick a row (or type four numbers). The script below plugs the values into the exact formulas; "
                   "the table then compares each hand-calculated result with the value the engine produced.")

        verify_source = st.radio("What would you like to verify?", [SOURCE_OPEX, SOURCE_SALES, SOURCE_TYPED],
                                 horizontal=True, key="verify_source")

        script_lines = []        # Python formulas with the numbers plugged in
        check_rows = []          # one row per comparison

        # ---------------- OPEX ledger row ----------------
        if verify_source == SOURCE_OPEX:
            picker_rows = df_opex.head(2000)
            row_labels = []
            for number, row in enumerate(picker_rows.itertuples(), start=1):
                row_labels.append(f"#{number} | {row.Date:%b %Y} | {row.Department} | {row.Expense_Line_Item}")
            choice = st.selectbox("Pick an OPEX row", row_labels, key="verify_opex_row")
            engine_row = picker_rows.iloc[row_labels.index(choice)]

            budget = float(engine_row["Budget_INR"])
            actual = float(engine_row["Actual_INR"])
            cost_type = engine_row["Cost_Type"]
            show_columns = ["Date", "Department", "Expense_Line_Item", "Cost_Type", "Budget_INR", "Actual_INR",
                            "Data_Fix_Note"]
            st.dataframe(engine_row.to_frame().T[show_columns], hide_index=True)
            if "Missing Actual filled with Budget" in engine_row["Data_Fix_Note"]:
                st.warning("This row's Actual was blank in the ERP file and was imputed from Budget, "
                           "so its variance is zero by construction.")

            # Hand calculation (uses only the four input values, never the engine's output columns)
            hand_variance = actual - budget
            hand_pct = np.nan
            if budget != 0:
                hand_pct = hand_variance / abs(budget) * 100
            if actual > budget:
                hand_flag = "U"
            elif actual < budget:
                hand_flag = "F"
            else:
                hand_flag = "On Track"
            hand_sim = actual
            if cost_type == "Variable":
                hand_sim = actual * volume_flex_factor * (1 + inflation_rate)

            a, b = py_num(actual), py_num(budget)       # the numbers as they appear in Python code
            script_lines.append(f"# {choice}   (Cost_Type = {cost_type})")
            script_lines.append("# Step 1: variance")
            script_lines.append(f"variance_inr = {a} - {b}      # = {hand_variance:,.2f}")
            script_lines.append("# Step 2: variance %, divided by |Budget|")
            if budget != 0:
                script_lines.append(f"variance_pct = {py_num(hand_variance)} / abs({b}) * 100      # = {hand_pct:.2f}")
            else:
                script_lines.append("variance_pct = NaN      # Budget is zero, so a percentage is undefined")
            script_lines.append("# Step 3: expense flag (spending more than budget is Unfavorable)")
            script_lines.append(f"variance_flag = 'U' if {a} > {b} else ('F' if {a} < {b} else 'On Track')"
                                f"      # = '{hand_flag}'")
            script_lines.append("# Step 4: scenario simulator (uses the current sidebar settings)")
            if cost_type == "Variable":
                script_lines.append(f"sim_actual_inr = {a} * {volume_flex_factor:.4g} * (1 + {inflation_rate:.4g})"
                                    f"      # Variable: flexed with volume, then inflated = {hand_sim:,.2f}")
            else:
                script_lines.append(f"sim_actual_inr = {a}      # Fixed cost: never changes")

            check_rows.append(compare_numbers("Variance_INR", hand_variance, engine_row["Variance_INR"]))
            check_rows.append(compare_numbers("Variance_Pct (engine rounds to 2 dp)", hand_pct,
                                              engine_row["Variance_Pct"], is_percent=True, tolerance=0.006))
            check_rows.append(compare_numbers("Simulated Actual_INR", hand_sim, engine_row["Sim_Actual_INR"]))
            check_rows.append(compare_text("Variance_Flag", hand_flag, engine_row["Variance_Flag"]))

        # ---------------- Sales row or four typed numbers ----------------
        else:
            if verify_source == SOURCE_SALES:
                picker_rows = df_sales.head(2000)
                row_labels = []
                for number, row in enumerate(picker_rows.itertuples(), start=1):
                    row_labels.append(f"#{number} | {row.Quarter} | {row.Product_Tier}")
                choice = st.selectbox("Pick a Sales row", row_labels, key="verify_sales_row")
                engine_row = picker_rows.iloc[row_labels.index(choice)]
                budget_units = float(engine_row["Budget_Units"])
                actual_units = float(engine_row["Actual_Units"])
                budget_price = float(engine_row["Budget_Price_INR"])
                actual_price = float(engine_row["Actual_Price_INR"])
                show_columns = ["Quarter", "Product_Tier", "Budget_Units", "Actual_Units",
                                "Budget_Price_INR", "Actual_Price_INR"]
                st.dataframe(engine_row.to_frame().T[show_columns], hide_index=True)
                script_lines.append(f"# {choice}")
            else:
                st.caption("Defaults are the Enterprise License Q3 numbers. Change any of them. The engine side runs "
                           "the real calculate_pvm_analysis() function on a one-row file built from your numbers.")
                input_a, input_b, input_c, input_d = st.columns(4)
                budget_units = input_a.number_input("Budget Units", min_value=0.0, value=170.0, step=1.0, key="test_bu")
                actual_units = input_b.number_input("Actual Units", min_value=0.0, value=230.0, step=1.0, key="test_au")
                budget_price = input_c.number_input("Budget Price (INR)", min_value=0.0, value=150000.0,
                                                    step=1000.0, key="test_bp")
                actual_price = input_d.number_input("Actual Price (INR)", min_value=0.0, value=123000.0,
                                                    step=1000.0, key="test_ap")
                test_ledger = pd.DataFrame([{"Quarter": "Q1 2026", "Product_Tier": "Test Tier",
                                             "Budget_Units": budget_units, "Actual_Units": actual_units,
                                             "Budget_Price_INR": budget_price, "Actual_Price_INR": actual_price}])
                test_file = io.BytesIO(test_ledger.to_csv(index=False).encode("utf-8"))
                test_file.name = "verifier_input.csv"
                engine_row = calculate_pvm_analysis(test_file).iloc[0]
                script_lines.append("# Your four test numbers")

            # Hand calculation (uses only the four input values, never the engine's output columns)
            hand_budget_revenue = budget_units * budget_price
            hand_actual_revenue = actual_units * actual_price
            hand_total = hand_actual_revenue - hand_budget_revenue
            hand_price_var = (actual_price - budget_price) * actual_units
            hand_volume_var = (actual_units - budget_units) * budget_price
            if hand_price_var > 0:
                hand_price_flag = "F"
            elif hand_price_var < 0:
                hand_price_flag = "U"
            else:
                hand_price_flag = "On Track"

            bu, au = py_num(budget_units), py_num(actual_units)       # the numbers as they appear in Python code
            bp, ap = py_num(budget_price), py_num(actual_price)
            script_lines.append(f"budget_revenue = {bu} * {bp}      # = {hand_budget_revenue:,.2f}")
            script_lines.append(f"actual_revenue = {au} * {ap}      # = {hand_actual_revenue:,.2f}")
            script_lines.append(f"total_sales_variance = {py_num(hand_actual_revenue)} - {py_num(hand_budget_revenue)}"
                                f"      # = {hand_total:,.2f}")
            script_lines.append(f"price_variance = ({ap} - {bp}) * {au}      # = {hand_price_var:,.2f}")
            script_lines.append(f"volume_variance = ({au} - {bu}) * {bp}      # = {hand_volume_var:,.2f}")
            script_lines.append(f"proof = {py_num(hand_price_var)} + {py_num(hand_volume_var)}"
                                f"      # = {hand_price_var + hand_volume_var:,.2f} (must equal the total)")

            check_rows.append(compare_numbers("Budget_Revenue", hand_budget_revenue, engine_row["Budget_Revenue"]))
            check_rows.append(compare_numbers("Actual_Revenue", hand_actual_revenue, engine_row["Actual_Revenue"]))
            check_rows.append(compare_numbers("Total_Sales_Variance", hand_total, engine_row["Total_Sales_Variance"]))
            check_rows.append(compare_numbers("Price_Variance", hand_price_var, engine_row["Price_Variance"]))
            check_rows.append(compare_numbers("Volume_Variance", hand_volume_var, engine_row["Volume_Variance"]))
            check_rows.append(compare_numbers("Price + Volume = Total", hand_price_var + hand_volume_var,
                                              engine_row["Total_Sales_Variance"]))
            check_rows.append(compare_text("Price_Flag", hand_price_flag, engine_row["Price_Flag"]))

            if verify_source == SOURCE_SALES:
                hand_sim_units = actual_units * (1 + volume_rate)
                hand_sim_price = actual_price * (1 + price_rate)
                hand_sim_revenue = hand_sim_units * hand_sim_price
                script_lines.append("# Scenario simulator (uses the current sidebar settings)")
                script_lines.append(f"sim_units = {au} * (1 + {volume_rate:.4g})      # = {hand_sim_units:,.2f}")
                script_lines.append(f"sim_price = {ap} * (1 + {price_rate:.4g})      # = {hand_sim_price:,.2f}")
                script_lines.append(f"sim_revenue = {py_num(hand_sim_units)} * {py_num(hand_sim_price)}"
                                    f"      # = {hand_sim_revenue:,.2f}")
                check_rows.append(compare_numbers("Sim_Revenue", hand_sim_revenue, engine_row["Sim_Revenue"]))

        # ---------------- Show the substitution and the comparison ----------------
        st.markdown("**Step-by-step substitution (the exact Python formulas with the numbers plugged in):**")
        st.code("\n".join(script_lines), language="python")

        comparison = pd.DataFrame(check_rows)
        st.dataframe(comparison, hide_index=True)
        mismatches = int((comparison["Match"] == "❌").sum())
        if mismatches == 0:
            st.success(f"All {len(comparison)} checks match: the hand calculation and the engine agree.")
        else:
            st.error(f"{mismatches} check(s) do not match. Please review the formulas.")

# ======================================================================
# 6. FOOTER
# ======================================================================
st.divider()
if flex_variable_costs:
    simulator_mode = "Flexible budget (Variable OPEX flexes with volume)."
else:
    simulator_mode = "Static budget (volume shock changes revenue only)."
st.caption("Assumptions: missing ERP Actuals are filled with Budget; missing Cost_Type is treated as Variable. "
           f"Simulator mode: {simulator_mode}")
st.markdown(FOOTER_HTML, unsafe_allow_html=True)