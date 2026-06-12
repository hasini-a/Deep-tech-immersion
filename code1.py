"""
╔══════════════════════════════════════════════════════════╗
║        ExpiryGuard — Retail Expiry Risk Manager          ║
║        Pure Streamlit + Pandas. No other libs.           ║
╠══════════════════════════════════════════════════════════╣
║  Install : pip install streamlit pandas                  ║
║  Run     : streamlit run app.py                          ║
╚══════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
from datetime import date, timedelta
import os

# ─────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────
CSV_PATH           = "inventory.csv"
CRITICAL_DAYS      = 30
EXPIRING_SOON_DAYS = 60

CATEGORIES = [
    "Dairy","Beverages","Bakery","Dry Goods",
    "Canned","Snacks","Cereals","Deli","Frozen","Other",
]
STATUS_ORDER = ["Expired","Critical","Expiring Soon","Safe"]

# Palette tokens — used in both CSS and Python-rendered HTML
CLR = {
    "bg":           "#0b0f1a",
    "surface":      "#111827",
    "surface2":     "#1a2236",
    "border":       "#1f2d45",
    "accent":       "#f5a623",       # warm amber — the signature tone
    "safe":         "#10b981",
    "soon":         "#f59e0b",
    "critical":     "#ef4444",
    "expired":      "#be185d",
    "text":         "#e2e8f0",
    "muted":        "#64748b",
    "safe_bg":      "#052e16",
    "soon_bg":      "#451a03",
    "critical_bg":  "#450a0a",
    "expired_bg":   "#4a044e",
}

STATUS_CLR = {
    "Safe":          CLR["safe"],
    "Expiring Soon": CLR["soon"],
    "Critical":      CLR["critical"],
    "Expired":       CLR["expired"],
}
STATUS_BG = {
    "Safe":          CLR["safe_bg"],
    "Expiring Soon": CLR["soon_bg"],
    "Critical":      CLR["critical_bg"],
    "Expired":       CLR["expired_bg"],
}
STATUS_ICON = {
    "Safe":          "✦",
    "Expiring Soon": "◈",
    "Critical":      "◉",
    "Expired":       "✖",
}

# ─────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ExpiryGuard",
    page_icon="🛡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────
# DESIGN SYSTEM CSS
# ─────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

/* ── Reset & base ───────────────────────────────────── */
*, *::before, *::after {{ box-sizing: border-box; }}

html, body, [data-testid="stAppViewContainer"],
[data-testid="stAppViewContainer"] > .main {{
    background: {CLR["bg"]} !important;
    color: {CLR["text"]};
    font-family: 'Inter', sans-serif;
}}

/* hide default streamlit chrome */
#MainMenu, footer, header {{ visibility: hidden; }}
[data-testid="stDecoration"] {{ display: none; }}

/* ── Sidebar ────────────────────────────────────────── */
[data-testid="stSidebar"] {{
    background: {CLR["surface"]} !important;
    border-right: 1px solid {CLR["border"]};
}}
[data-testid="stSidebar"] * {{ color: {CLR["text"]} !important; }}
[data-testid="stSidebar"] .stRadio label {{
    padding: 9px 14px !important;
    border-radius: 10px !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    transition: background 0.2s, color 0.2s !important;
}}
[data-testid="stSidebar"] .stRadio label:hover {{
    background: {CLR["surface2"]} !important;
    color: {CLR["accent"]} !important;
}}
[data-testid="stSidebar"] [data-baseweb="radio"] input:checked + div {{
    background: {CLR["accent"]}22 !important;
}}

/* ── Logo strip in sidebar ───────────────────────────  */
.eg-logo {{
    font-family: 'Inter', sans-serif;
    font-size: 1.35rem;
    font-weight: 800;
    letter-spacing: -0.5px;
    color: {CLR["accent"]};
    display: flex; align-items: center; gap: 8px;
    padding: 4px 0 2px;
}}
.eg-logo span {{ color: {CLR["text"]}; font-weight: 300; }}
.eg-tagline {{
    font-size: 0.72rem; color: {CLR["muted"]};
    text-transform: uppercase; letter-spacing: 0.12em;
    margin-top: 2px;
}}

/* ── Page header band ───────────────────────────────── */
.page-header {{
    background: linear-gradient(135deg, {CLR["surface2"]} 0%, #0f1e35 60%, {CLR["bg"]} 100%);
    border: 1px solid {CLR["border"]};
    border-radius: 16px;
    padding: 28px 32px 22px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}}
.page-header::before {{
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, {CLR["accent"]}, #f97316, {CLR["critical"]}, {CLR["expired"]});
    border-radius: 16px 16px 0 0;
}}
.page-header h1 {{
    font-size: 1.9rem; font-weight: 800;
    color: {CLR["text"]} !important;
    margin: 0 0 4px; letter-spacing: -0.5px;
}}
.page-header p {{
    color: {CLR["muted"]}; font-size: 0.88rem; margin: 0;
}}

/* ── KPI cards ───────────────────────────────────────── */
.kpi-card {{
    background: {CLR["surface"]};
    border: 1px solid {CLR["border"]};
    border-radius: 14px;
    padding: 20px 22px 18px;
    position: relative;
    overflow: hidden;
    height: 110px;
    display: flex; flex-direction: column; justify-content: space-between;
}}
.kpi-card::after {{
    content: attr(data-icon);
    position: absolute; right: 16px; bottom: 12px;
    font-size: 2.4rem; opacity: 0.08;
    pointer-events: none;
}}
.kpi-label {{
    font-size: 0.72rem; font-weight: 600;
    text-transform: uppercase; letter-spacing: 0.1em;
    color: {CLR["muted"]};
}}
.kpi-value {{
    font-size: 2.2rem; font-weight: 800;
    letter-spacing: -1px; line-height: 1;
}}
.kpi-sub {{
    font-size: 0.72rem; color: {CLR["muted"]}; margin-top: 4px;
}}
.kpi-accent  {{ color: {CLR["accent"]};   border-top: 3px solid {CLR["accent"]};   }}
.kpi-safe    {{ color: {CLR["safe"]};     border-top: 3px solid {CLR["safe"]};     }}
.kpi-soon    {{ color: {CLR["soon"]};     border-top: 3px solid {CLR["soon"]};     }}
.kpi-critical{{ color: {CLR["critical"]}; border-top: 3px solid {CLR["critical"]}; }}
.kpi-expired {{ color: {CLR["expired"]};  border-top: 3px solid {CLR["expired"]};  }}

/* ── Section headings ───────────────────────────────── */
.section-title {{
    font-size: 1rem; font-weight: 700;
    color: {CLR["text"]};
    text-transform: uppercase; letter-spacing: 0.08em;
    display: flex; align-items: center; gap: 8px;
    margin-bottom: 14px;
}}
.section-title::after {{
    content: ''; flex: 1; height: 1px;
    background: linear-gradient(90deg, {CLR["border"]}, transparent);
}}

/* ── Glass panel ─────────────────────────────────────── */
.glass {{
    background: {CLR["surface"]};
    border: 1px solid {CLR["border"]};
    border-radius: 14px;
    padding: 22px 24px;
    margin-bottom: 16px;
}}

/* ── Horizontal bar (custom chart) ──────────────────── */
.hbar-wrap {{ margin-bottom: 10px; }}
.hbar-label {{
    font-size: 0.82rem; color: {CLR["text"]};
    display: flex; justify-content: space-between; margin-bottom: 4px;
}}
.hbar-track {{
    background: {CLR["surface2"]}; border-radius: 99px;
    height: 10px; overflow: hidden;
}}
.hbar-fill {{
    height: 100%; border-radius: 99px;
    transition: width 0.6s cubic-bezier(.4,0,.2,1);
}}

/* ── Alert strips ────────────────────────────────────── */
.alert-strip {{
    border-radius: 10px;
    padding: 12px 16px;
    margin: 6px 0;
    display: flex; flex-direction: column; gap: 3px;
    border: 1px solid;
}}
.alert-strip .as-title {{
    font-weight: 700; font-size: 0.92rem;
}}
.alert-strip .as-body {{
    font-size: 0.8rem; opacity: 0.85;
    font-family: 'JetBrains Mono', monospace;
}}
.as-expired  {{ background:{CLR["expired_bg"]}; border-color:{CLR["expired"]}44; color:{CLR["expired"]}; }}
.as-critical {{ background:{CLR["critical_bg"]}; border-color:{CLR["critical"]}44; color:{CLR["critical"]}; }}
.as-soon     {{ background:{CLR["soon_bg"]}; border-color:{CLR["soon"]}44; color:{CLR["soon"]}; }}
.as-ok       {{ background:{CLR["safe_bg"]}; border-color:{CLR["safe"]}44; color:{CLR["safe"]}; }}

/* ── Status pill ─────────────────────────────────────── */
.pill {{
    display: inline-block;
    padding: 2px 10px; border-radius: 99px;
    font-size: 0.75rem; font-weight: 700;
    letter-spacing: 0.05em;
    border: 1px solid;
}}
.pill-safe     {{ background:{CLR["safe_bg"]};     color:{CLR["safe"]};     border-color:{CLR["safe"]}55; }}
.pill-soon     {{ background:{CLR["soon_bg"]};     color:{CLR["soon"]};     border-color:{CLR["soon"]}55; }}
.pill-critical {{ background:{CLR["critical_bg"]}; color:{CLR["critical"]}; border-color:{CLR["critical"]}55; }}
.pill-expired  {{ background:{CLR["expired_bg"]};  color:{CLR["expired"]};  border-color:{CLR["expired"]}55; }}

/* ── Product card (manage page) ──────────────────────── */
.prod-card {{
    background: {CLR["surface"]};
    border: 1px solid {CLR["border"]};
    border-radius: 12px;
    padding: 14px 18px;
    margin: 6px 0;
    transition: border-color 0.2s;
}}
.prod-card:hover {{ border-color: {CLR["accent"]}66; }}
.prod-name  {{ font-weight: 700; font-size: 0.95rem; color: {CLR["text"]}; }}
.prod-meta  {{ font-size: 0.78rem; color: {CLR["muted"]}; margin-top: 3px;
               font-family: 'JetBrains Mono', monospace; }}

/* ── Streamlit widget overrides ──────────────────────── */
div[data-testid="metric-container"] {{
    background: {CLR["surface"]} !important;
    border: 1px solid {CLR["border"]} !important;
    border-radius: 12px !important;
    padding: 14px 18px !important;
}}
[data-testid="stTabs"] [role="tab"] {{
    font-weight: 600; font-size: 0.85rem;
    color: {CLR["muted"]} !important;
    border-bottom: 2px solid transparent !important;
    padding: 8px 16px !important;
}}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{
    color: {CLR["accent"]} !important;
    border-bottom-color: {CLR["accent"]} !important;
}}
.stButton > button {{
    background: linear-gradient(135deg, {CLR["accent"]}, #f97316) !important;
    color: #0b0f1a !important; font-weight: 700 !important;
    border: none !important; border-radius: 10px !important;
    padding: 8px 22px !important; letter-spacing: 0.03em;
    transition: opacity 0.2s !important;
}}
.stButton > button:hover {{ opacity: 0.85 !important; }}

.stDownloadButton > button {{
    background: {CLR["surface2"]} !important;
    color: {CLR["accent"]} !important; font-weight: 700 !important;
    border: 1px solid {CLR["accent"]}66 !important;
    border-radius: 10px !important; padding: 8px 22px !important;
}}
.stDownloadButton > button:hover {{
    background: {CLR["accent"]}22 !important;
}}

.stTextInput input, .stNumberInput input,
.stDateInput input, .stSelectbox > div > div {{
    background: {CLR["surface2"]} !important;
    color: {CLR["text"]} !important;
    border-color: {CLR["border"]} !important;
    border-radius: 9px !important;
}}
.stForm {{ background: {CLR["surface"]}; border-radius: 14px; padding: 4px; }}
[data-testid="stDataFrame"] {{ border-radius: 12px; overflow: hidden; }}

.stDivider {{ border-color: {CLR["border"]} !important; }}

/* ── Scrollbar ───────────────────────────────────────── */
::-webkit-scrollbar {{ width: 6px; height: 6px; }}
::-webkit-scrollbar-track {{ background: {CLR["bg"]}; }}
::-webkit-scrollbar-thumb {{ background: {CLR["border"]}; border-radius: 3px; }}
::-webkit-scrollbar-thumb:hover {{ background: {CLR["accent"]}66; }}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────
# HTML HELPERS
# ─────────────────────────────────────────────────────────

def page_header(title: str, subtitle: str):
    st.markdown(
        f'<div class="page-header"><h1>{title}</h1><p>{subtitle}</p></div>',
        unsafe_allow_html=True,
    )


def section(label: str):
    st.markdown(f'<div class="section-title">{label}</div>', unsafe_allow_html=True)


def kpi_card(label: str, value, sub: str, style: str, icon: str):
    st.markdown(
        f'<div class="kpi-card kpi-{style}" data-icon="{icon}">'
        f'  <div class="kpi-label">{label}</div>'
        f'  <div>'
        f'    <div class="kpi-value">{value}</div>'
        f'    <div class="kpi-sub">{sub}</div>'
        f'  </div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def hbar(label: str, value: float, max_val: float, color: str, suffix: str = ""):
    pct = min(100, round(value / max_val * 100)) if max_val > 0 else 0
    display = f"{value:,.0f}{suffix}"
    st.markdown(
        f'<div class="hbar-wrap">'
        f'  <div class="hbar-label"><span>{label}</span><span>{display}</span></div>'
        f'  <div class="hbar-track">'
        f'    <div class="hbar-fill" style="width:{pct}%;background:{color}"></div>'
        f'  </div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def status_pill(status: str) -> str:
    cls = {"Safe":"safe","Expiring Soon":"soon","Critical":"critical","Expired":"expired"}.get(status,"safe")
    return f'<span class="pill pill-{cls}">{STATUS_ICON[status]} {status}</span>'


def alert_strip(status: str, title: str, body: str):
    cls = {"Expired":"as-expired","Critical":"as-critical",
           "Expiring Soon":"as-soon","Safe":"as-ok"}.get(status,"as-ok")
    st.markdown(
        f'<div class="alert-strip {cls}">'
        f'<div class="as-title">{STATUS_ICON[status]} {title}</div>'
        f'<div class="as-body">{body}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def prod_card(row):
    clr = STATUS_CLR.get(row["Status"], CLR["text"])
    pill = status_pill(row["Status"])
    st.markdown(
        f'<div class="prod-card">'
        f'  <div class="prod-name">{row["Product Name"]} &nbsp;{pill}</div>'
        f'  <div class="prod-meta">'
        f'    Cat: {row["Category"]} &nbsp;·&nbsp; Batch: {row["Batch Number"]} &nbsp;·&nbsp;'
        f'    Qty: {row["Quantity"]} &nbsp;·&nbsp; Expires: {row["Expiry Date"]} &nbsp;·&nbsp;'
        f'    <span style="color:{clr};font-weight:700">{row["Days Remaining"]}d remaining</span>'
        f'  </div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────
# DATA LAYER
# ─────────────────────────────────────────────────────────

def generate_sample_data() -> pd.DataFrame:
    today = date.today()
    rows = [
        ("Whole Milk 1L",       "Dairy",     "B001", 50,  today-timedelta(10),  today+timedelta(5),   45.00),
        ("Cheddar Cheese 500g", "Dairy",     "B002", 30,  today-timedelta(20),  today+timedelta(22),  130.00),
        ("Orange Juice 2L",     "Beverages", "B003", 40,  today-timedelta(30),  today+timedelta(45),  95.00),
        ("White Bread",         "Bakery",    "B004", 25,  today-timedelta(3),   today+timedelta(3),   40.00),
        ("Pasta 500g",          "Dry Goods", "B005", 100, today-timedelta(60),  today+timedelta(180), 30.00),
        ("Tomato Sauce 400g",   "Canned",    "B006", 80,  today-timedelta(90),  today+timedelta(120), 35.00),
        ("Greek Yogurt 200g",   "Dairy",     "B007", 20,  today-timedelta(5),   today-timedelta(1),   28.00),
        ("Sparkling Water 1L",  "Beverages", "B008", 60,  today-timedelta(120), today+timedelta(90),  18.00),
        ("Granola Bars x6",     "Snacks",    "B009", 45,  today-timedelta(45),  today+timedelta(50),  75.00),
        ("Butter 250g",         "Dairy",     "B010", 18,  today-timedelta(15),  today+timedelta(18),  62.00),
        ("Corn Flakes 500g",    "Cereals",   "B011", 35,  today-timedelta(80),  today+timedelta(200), 80.00),
        ("Salted Crackers",     "Snacks",    "B012", 55,  today-timedelta(60),  today+timedelta(40),  42.00),
        ("Apple Juice 1L",      "Beverages", "B013", 28,  today-timedelta(40),  today-timedelta(5),   55.00),
        ("Sliced Ham 200g",     "Deli",      "B014", 12,  today-timedelta(8),   today+timedelta(10),  110.00),
        ("Cream Cheese 150g",   "Dairy",     "B015", 22,  today-timedelta(12),  today+timedelta(14),  85.00),
    ]
    df = pd.DataFrame(rows, columns=[
        "Product Name","Category","Batch Number","Quantity",
        "Manufacturing Date","Expiry Date","Cost Per Unit",
    ])
    df["Manufacturing Date"] = pd.to_datetime(df["Manufacturing Date"]).dt.date
    df["Expiry Date"]        = pd.to_datetime(df["Expiry Date"]).dt.date
    return df


def load_data() -> pd.DataFrame:
    if os.path.exists(CSV_PATH):
        df = pd.read_csv(CSV_PATH)
        df["Manufacturing Date"] = pd.to_datetime(df["Manufacturing Date"]).dt.date
        df["Expiry Date"]        = pd.to_datetime(df["Expiry Date"]).dt.date
        df["Cost Per Unit"]      = pd.to_numeric(df["Cost Per Unit"], errors="coerce").fillna(0)
    else:
        df = generate_sample_data()
        df.to_csv(CSV_PATH, index=False)
    return df


def save_data(df: pd.DataFrame):
    df.to_csv(CSV_PATH, index=False)


def classify_status(days: int) -> str:
    if days < 0:                   return "Expired"
    if days < CRITICAL_DAYS:       return "Critical"
    if days < EXPIRING_SOON_DAYS:  return "Expiring Soon"
    return "Safe"


def enrich(raw: pd.DataFrame) -> pd.DataFrame:
    today = date.today()
    d = raw.copy()
    d["Days Remaining"] = pd.to_datetime(d["Expiry Date"]).apply(lambda x: (x.date()-today).days)
    d["Status"]         = d["Days Remaining"].apply(classify_status)
    d["Risk Score"]     = d["Days Remaining"].apply(
        lambda x: 100 if x < 0 else max(0, round(100-(x/EXPIRING_SOON_DAYS)*100))
    )
    d["Est. Loss (₹)"]  = (d["Quantity"]*d["Cost Per Unit"]).round(2)
    return d


# ─────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────
if "data"     not in st.session_state: st.session_state.data     = load_data()
if "edit_idx" not in st.session_state: st.session_state.edit_idx = None

def gdf() -> pd.DataFrame:
    return enrich(st.session_state.data)


# ─────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        '<div class="eg-logo">🛡 ExpiryGuard<span></span></div>'
        '<div class="eg-tagline">Retail Risk Management</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    page = st.radio(
        "nav", label_visibility="collapsed",
        options=[
            "📊  Dashboard",
            "➕  Product Entry",
            "⚠️  Expiry Risk Analysis",
            "🔔  Alerts",
            "📈  Reports & Analytics",
            "🔍  Search & Filter",
        ],
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        f'<div style="font-size:0.78rem;color:{CLR["muted"]}">'
        f'📅 &nbsp;<b style="color:{CLR["text"]}">{date.today().strftime("%d %b %Y")}</b><br>'
        f'📦 &nbsp;<b style="color:{CLR["text"]}">{len(st.session_state.data)}</b> products'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⟳  Reload CSV", use_container_width=True):
        st.session_state.data = load_data()
        st.success("Data refreshed!")


# ══════════════════════════════════════════════════════════
# PAGE 1 ── DASHBOARD
# ══════════════════════════════════════════════════════════
if page == "📊  Dashboard":
    page_header(
        "Inventory Dashboard",
        f"Live expiry overview · {date.today().strftime('%A, %d %B %Y')}",
    )

    data = gdf()
    total    = len(data)
    safe     = int((data["Status"]=="Safe").sum())
    soon     = int((data["Status"]=="Expiring Soon").sum())
    critical = int((data["Status"]=="Critical").sum())
    expired  = int((data["Status"]=="Expired").sum())
    risk_val = float(data[data["Status"].isin(["Expired","Critical","Expiring Soon"])]["Est. Loss (₹)"].sum())

    # ── KPI row ──
    section("Key Metrics")
    k1,k2,k3,k4,k5 = st.columns(5)
    with k1: kpi_card("Total Products",   total,    "all SKUs",           "accent",   "📦")
    with k2: kpi_card("Safe",             safe,     "> 60 days",          "safe",     "✓")
    with k3: kpi_card("Expiring Soon",    soon,     "30 – 60 days",       "soon",     "⏱")
    with k4: kpi_card("Critical",         critical, "< 30 days",          "critical", "!")
    with k5: kpi_card("Expired",          expired,  "past date",          "expired",  "✖")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Risk loss callout ──
    loss_pct = round(risk_val / max(1, data["Est. Loss (₹)"].sum()) * 100)
    st.markdown(
        f'<div style="background:linear-gradient(135deg,{CLR["critical_bg"]},{CLR["expired_bg"]});'
        f'border:1px solid {CLR["critical"]}44;border-radius:12px;padding:16px 22px;'
        f'display:flex;align-items:center;justify-content:space-between;margin-bottom:24px">'
        f'  <div>'
        f'    <div style="font-size:0.72rem;color:{CLR["critical"]};font-weight:700;text-transform:uppercase;letter-spacing:.1em">Estimated Inventory at Risk</div>'
        f'    <div style="font-size:2rem;font-weight:800;color:{CLR["text"]}">₹{risk_val:,.2f}</div>'
        f'  </div>'
        f'  <div style="text-align:right">'
        f'    <div style="font-size:0.72rem;color:{CLR["muted"]}">of total value</div>'
        f'    <div style="font-size:1.6rem;font-weight:800;color:{CLR["critical"]}">{loss_pct}%</div>'
        f'  </div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    col_l, col_r = st.columns([1,1])

    # ── Status distribution (custom hbars) ──
    with col_l:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        section("Status Distribution")
        max_n = max(safe, soon, critical, expired, 1)
        hbar("✅ Safe",          safe,     max_n, CLR["safe"])
        hbar("◈ Expiring Soon",  soon,     max_n, CLR["soon"])
        hbar("◉ Critical",       critical, max_n, CLR["critical"])
        hbar("✖ Expired",        expired,  max_n, CLR["expired"])
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Category breakdown (hbars) ──
    with col_r:
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        section("Products by Category")
        cat_counts = (
            data.groupby("Category")["Product Name"]
            .count().sort_values(ascending=False)
        )
        max_c = cat_counts.max()
        colors_cycle = [CLR["accent"],"#f97316",CLR["safe"],CLR["soon"],CLR["critical"],
                        "#a78bfa","#38bdf8","#fb7185","#4ade80","#fbbf24"]
        for i,(cat,cnt) in enumerate(cat_counts.items()):
            hbar(cat, cnt, max_c, colors_cycle[i % len(colors_cycle)])
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Top 10 closest to expiry ──
    st.markdown("<br>", unsafe_allow_html=True)
    section("🔺 Top 10 Closest to Expiry")
    top10 = (
        data[data["Days Remaining"] >= 0]
        .sort_values("Days Remaining").head(10)
        [["Product Name","Category","Days Remaining","Quantity","Status","Risk Score"]]
        .reset_index(drop=True)
    )
    if top10.empty:
        st.info("No active products.")
    else:
        # Visual bar per row using hbars inline in the glass panel
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        max_days = max(top10["Days Remaining"].max(), 1)
        for _, row in top10.iterrows():
            c = STATUS_CLR.get(row["Status"], CLR["muted"])
            d = int(row["Days Remaining"])
            pill_html = status_pill(row["Status"])
            st.markdown(
                f'<div style="margin-bottom:12px">'
                f'  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">'
                f'    <span style="font-weight:600;font-size:0.88rem">{row["Product Name"]}</span>'
                f'    <span>{pill_html} &nbsp;<span style="font-size:0.78rem;color:{CLR["muted"]};font-family:JetBrains Mono,monospace">Qty {int(row["Quantity"])}</span></span>'
                f'  </div>'
                f'  <div class="hbar-track"><div class="hbar-fill" style="width:{round(d/max_days*100)}%;background:{c}"></div></div>'
                f'  <div style="font-size:0.72rem;color:{CLR["muted"]};margin-top:3px">{d} days remaining</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# PAGE 2 ── PRODUCT ENTRY
# ══════════════════════════════════════════════════════════
elif page == "➕  Product Entry":
    page_header("Product Entry", "Add new stock or manage existing inventory")

    tab_add, tab_manage = st.tabs(["  ✦  Add New Product  ","  ✦  Manage Existing  "])

    with tab_add:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.form("add_form", clear_on_submit=True):
            r1a,r1b = st.columns(2)
            name  = r1a.text_input("Product Name *", placeholder="e.g. Full Cream Milk 500ml")
            cat   = r1b.selectbox("Category *", CATEGORIES)

            r2a,r2b = st.columns(2)
            batch = r2a.text_input("Batch Number *", placeholder="e.g. B0042")
            qty   = r2b.number_input("Quantity *", min_value=1, step=1, value=10)

            r3a,r3b,r3c = st.columns(3)
            mfg = r3a.date_input("Manufacturing Date *", value=date.today()-timedelta(30))
            exp = r3b.date_input("Expiry Date *",        value=date.today()+timedelta(90))
            cost= r3c.number_input("Cost Per Unit (₹) *", min_value=0.0, step=5.0,
                                   value=50.0, format="%.2f")

            st.markdown("<br>", unsafe_allow_html=True)
            submitted = st.form_submit_button("＋  Add to Inventory", use_container_width=True)

        if submitted:
            errs = []
            if not name.strip():  errs.append("Product Name is required.")
            if not batch.strip(): errs.append("Batch Number is required.")
            if exp <= mfg:        errs.append("Expiry date must be after manufacturing date.")
            if errs:
                for e in errs: st.error(e)
            else:
                new = pd.DataFrame([{
                    "Product Name": name.strip(), "Category": cat,
                    "Batch Number": batch.strip(), "Quantity": int(qty),
                    "Manufacturing Date": mfg, "Expiry Date": exp,
                    "Cost Per Unit": float(cost),
                }])
                st.session_state.data = pd.concat([st.session_state.data, new], ignore_index=True)
                save_data(st.session_state.data)
                st.success(f"✅  '{name}' added successfully!")

    with tab_manage:
        raw = gdf()
        if raw.empty:
            st.info("No products yet.")
        else:
            srch = st.text_input("🔍  Filter by name", placeholder="Start typing…",
                                 label_visibility="collapsed")
            disp = raw.copy()
            if srch:
                disp = disp[disp["Product Name"].str.contains(srch, case=False, na=False)]

            for i, row in disp.iterrows():
                col_card, col_ed, col_del = st.columns([8, 1, 1])
                with col_card:
                    prod_card(row)
                with col_ed:
                    st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
                    if st.button("✏", key=f"ed_{i}"):
                        st.session_state.edit_idx = i
                with col_del:
                    st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
                    if st.button("✖", key=f"del_{i}"):
                        st.session_state.data = st.session_state.data.drop(index=i).reset_index(drop=True)
                        save_data(st.session_state.data)
                        st.success("Deleted.")
                        st.rerun()

            eidx = st.session_state.edit_idx
            if eidx is not None and eidx in st.session_state.data.index:
                r = st.session_state.data.loc[eidx]
                st.divider()
                st.markdown(
                    f'<div class="section-title">Editing: {r["Product Name"]}</div>',
                    unsafe_allow_html=True,
                )
                with st.form("edit_form"):
                    ea,eb = st.columns(2)
                    e_name = ea.text_input("Product Name", value=r["Product Name"])
                    ci = CATEGORIES.index(r["Category"]) if r["Category"] in CATEGORIES else 0
                    e_cat  = eb.selectbox("Category", CATEGORIES, index=ci)
                    ec,ed  = st.columns(2)
                    e_batch= ec.text_input("Batch Number", value=r["Batch Number"])
                    e_qty  = ed.number_input("Quantity", min_value=1, value=int(r["Quantity"]))
                    ee,ef,eg = st.columns(3)
                    e_mfg  = ee.date_input("Mfg Date",    value=r["Manufacturing Date"])
                    e_exp  = ef.date_input("Expiry Date", value=r["Expiry Date"])
                    e_cost = eg.number_input("Cost/Unit", min_value=0.0,
                                             value=float(r["Cost Per Unit"]), format="%.2f")
                    save_btn = st.form_submit_button("💾  Save Changes", use_container_width=True)

                if save_btn:
                    if e_exp <= e_mfg:
                        st.error("Expiry date must be after manufacturing date.")
                    else:
                        for k,v in zip(
                            ["Product Name","Category","Batch Number","Quantity",
                             "Manufacturing Date","Expiry Date","Cost Per Unit"],
                            [e_name, e_cat, e_batch, e_qty, e_mfg, e_exp, e_cost]
                        ):
                            st.session_state.data.at[eidx, k] = v
                        save_data(st.session_state.data)
                        st.session_state.edit_idx = None
                        st.success("Saved!")
                        st.rerun()


# ══════════════════════════════════════════════════════════
# PAGE 3 ── EXPIRY RISK ANALYSIS
# ══════════════════════════════════════════════════════════
elif page == "⚠️  Expiry Risk Analysis":
    page_header("Expiry Risk Analysis", "Automatic classification of all products by risk level")

    data = gdf()

    k1,k2,k3,k4 = st.columns(4)
    for col, status in zip([k1,k2,k3,k4], STATUS_ORDER):
        n = int((data["Status"]==status).sum())
        style = {"Expired":"expired","Critical":"critical","Expiring Soon":"soon","Safe":"safe"}[status]
        with col:
            kpi_card(status, n, f"{STATUS_ICON[status]} products", style, STATUS_ICON[status])

    st.markdown("<br>", unsafe_allow_html=True)

    # Days-remaining distribution as custom hbars (bucketed)
    section("Days-Remaining Distribution")
    st.markdown('<div class="glass">', unsafe_allow_html=True)
    buckets = [
        ("Expired  (< 0d)",    data[data["Days Remaining"] < 0],           CLR["expired"]),
        ("Critical (0–15d)",   data[data["Days Remaining"].between(0,14)],  CLR["critical"]),
        ("Critical (15–30d)",  data[data["Days Remaining"].between(15,29)], "#f97316"),
        ("Soon     (30–45d)",  data[data["Days Remaining"].between(30,44)], CLR["soon"]),
        ("Soon     (45–60d)",  data[data["Days Remaining"].between(45,59)], "#facc15"),
        ("Safe     (60–90d)",  data[data["Days Remaining"].between(60,89)], CLR["safe"]),
        ("Safe     (90–180d)", data[data["Days Remaining"].between(90,179)],"#6ee7b7"),
        ("Safe     (180d+)",   data[data["Days Remaining"] >= 180],         "#a7f3d0"),
    ]
    max_b = max((len(b[1]) for b in buckets), default=1)
    for label, sub, color in buckets:
        hbar(label, len(sub), max_b, color, " products")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    SHOW = ["Product Name","Category","Batch Number","Quantity",
            "Expiry Date","Days Remaining","Risk Score","Est. Loss (₹)"]

    for status in STATUS_ORDER:
        subset = data[data["Status"]==status].sort_values("Days Remaining").reset_index(drop=True)
        if subset.empty: continue
        c = STATUS_CLR[status]
        st.markdown(
            f'<div class="section-title" style="color:{c}">'
            f'{STATUS_ICON[status]} {status} &nbsp;<span style="font-size:0.78rem;'
            f'background:{STATUS_BG[status]};color:{c};padding:2px 10px;border-radius:99px">'
            f'{len(subset)} products</span></div>',
            unsafe_allow_html=True,
        )
        st.dataframe(
            subset[SHOW],
            use_container_width=True, hide_index=True,
            column_config={
                "Risk Score": st.column_config.ProgressColumn(
                    "Risk Score", min_value=0, max_value=100, format="%d"),
                "Days Remaining": st.column_config.NumberColumn("Days Left", format="%d d"),
                "Est. Loss (₹)":  st.column_config.NumberColumn(format="₹%.2f"),
            },
        )
        st.markdown("<br>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# PAGE 4 ── ALERTS
# ══════════════════════════════════════════════════════════
elif page == "🔔  Alerts":
    page_header("Alerts Centre", "Products requiring immediate or near-term action")

    data     = gdf()
    expired  = data[data["Status"]=="Expired"].sort_values("Days Remaining")
    critical = data[data["Status"]=="Critical"].sort_values("Days Remaining")
    soon     = data[data["Status"]=="Expiring Soon"].sort_values("Days Remaining")

    total_risk_val = float(data[data["Status"].isin(["Expired","Critical","Expiring Soon"])]["Est. Loss (₹)"].sum())

    # Big risk banner
    if len(expired)+len(critical) > 0:
        st.markdown(
            f'<div style="background:linear-gradient(135deg,{CLR["critical_bg"]},{CLR["expired_bg"]});'
            f'border:1px solid {CLR["critical"]}55;border-radius:12px;padding:18px 24px;margin-bottom:20px">'
            f'<div style="font-size:0.72rem;color:{CLR["critical"]};font-weight:700;text-transform:uppercase;letter-spacing:.1em">Action Required</div>'
            f'<div style="font-size:1.05rem;color:{CLR["text"]};margin-top:4px">'
            f'<b>{len(expired)}</b> expired &nbsp;·&nbsp; <b>{len(critical)}</b> critical &nbsp;·&nbsp; '
            f'Estimated loss: <b style="color:{CLR["critical"]}">₹{total_risk_val:,.2f}</b>'
            f'</div></div>',
            unsafe_allow_html=True,
        )

    # Expired
    section(f"✖ Expired  ({len(expired)})")
    if expired.empty:
        st.markdown('<div class="alert-strip as-ok"><div class="as-title">✦ No expired products</div></div>', unsafe_allow_html=True)
    else:
        for _, r in expired.iterrows():
            alert_strip("Expired", r["Product Name"],
                f"Batch {r['Batch Number']}  ·  {abs(r['Days Remaining'])}d ago ({r['Expiry Date']})  ·  "
                f"Qty: {r['Quantity']}  ·  Est. loss ₹{r['Est. Loss (₹)']:.2f}")
    st.markdown("<br>", unsafe_allow_html=True)

    # Critical
    section(f"◉ Critical  ({len(critical)})")
    if critical.empty:
        st.markdown('<div class="alert-strip as-ok"><div class="as-title">✦ No critical products</div></div>', unsafe_allow_html=True)
    else:
        for _, r in critical.iterrows():
            alert_strip("Critical", r["Product Name"],
                f"Batch {r['Batch Number']}  ·  expires in {r['Days Remaining']}d ({r['Expiry Date']})  ·  "
                f"Qty: {r['Quantity']}  ·  Risk {r['Risk Score']}/100")
    st.markdown("<br>", unsafe_allow_html=True)

    # Expiring soon
    section(f"◈ Expiring Soon  ({len(soon)})")
    if soon.empty:
        st.markdown('<div class="alert-strip as-ok"><div class="as-title">✦ Nothing in 30–60-day window</div></div>', unsafe_allow_html=True)
    else:
        for _, r in soon.iterrows():
            alert_strip("Expiring Soon", r["Product Name"],
                f"Batch {r['Batch Number']}  ·  expires in {r['Days Remaining']}d ({r['Expiry Date']})  ·  Qty: {r['Quantity']}")


# ══════════════════════════════════════════════════════════
# PAGE 5 ── REPORTS & ANALYTICS
# ══════════════════════════════════════════════════════════
elif page == "📈  Reports & Analytics":
    page_header("Reports & Analytics", "Inventory insights, category breakdown, financial risk")

    data = gdf()

    tab_cat, tab_status, tab_risk, tab_dl = st.tabs([
        "  📦 By Category  ","  📊 Status  ","  💸 Financial Risk  ","  📥 Download  "
    ])

    with tab_cat:
        st.markdown("<br>", unsafe_allow_html=True)
        cat_df = (
            data.groupby("Category")
            .agg(
                Products   =("Product Name","count"),
                Total_Qty  =("Quantity","sum"),
                At_Risk    =("Status", lambda x: x.isin(["Critical","Expiring Soon","Expired"]).sum()),
                Est_Loss   =("Est. Loss (₹)","sum"),
            ).reset_index().rename(columns={"Est_Loss":"Est. Loss (₹)"})
            .sort_values("Products", ascending=False)
        )
        st.dataframe(
            cat_df, use_container_width=True, hide_index=True,
            column_config={
                "Est. Loss (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                "At_Risk":       st.column_config.NumberColumn("At Risk"),
            },
        )
        st.markdown("<br>", unsafe_allow_html=True)
        section("Products per Category")
        st.markdown('<div class="glass">', unsafe_allow_html=True)
        colors_cycle = [CLR["accent"],"#f97316",CLR["safe"],CLR["soon"],CLR["critical"],
                        "#a78bfa","#38bdf8","#fb7185","#4ade80","#fbbf24"]
        max_p = cat_df["Products"].max()
        for i, row in cat_df.iterrows():
            hbar(row["Category"], row["Products"], max_p, colors_cycle[int(i) % len(colors_cycle)])
        st.markdown('</div>', unsafe_allow_html=True)

    with tab_status:
        st.markdown("<br>", unsafe_allow_html=True)
        status_df = (
            data.groupby("Status")
            .agg(Products=("Product Name","count"), Total_Qty=("Quantity","sum"))
            .reindex(STATUS_ORDER).fillna(0).astype(int).reset_index()
        )
        st.dataframe(status_df, use_container_width=True, hide_index=True)
        st.markdown("<br>", unsafe_allow_html=True)

        col_a, col_b = st.columns(2)
        with col_a:
            section("Product Count")
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            max_p = max(status_df["Products"].max(), 1)
            for _, row in status_df.iterrows():
                hbar(f"{STATUS_ICON[row['Status']]} {row['Status']}", row["Products"],
                     max_p, STATUS_CLR[row["Status"]])
            st.markdown('</div>', unsafe_allow_html=True)
        with col_b:
            section("Total Quantity")
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            max_q = max(status_df["Total_Qty"].max(), 1)
            for _, row in status_df.iterrows():
                hbar(f"{STATUS_ICON[row['Status']]} {row['Status']}", row["Total_Qty"],
                     max_q, STATUS_CLR[row["Status"]])
            st.markdown('</div>', unsafe_allow_html=True)

    with tab_risk:
        st.markdown("<br>", unsafe_allow_html=True)
        risk_data = data[data["Status"].isin(["Expired","Critical","Expiring Soon"])].copy()
        if risk_data.empty:
            st.success("🎉  No products at risk right now!")
        else:
            total_risk = float(risk_data["Est. Loss (₹)"].sum())
            rk1,rk2,rk3 = st.columns(3)
            with rk1: kpi_card("At-Risk Products", len(risk_data), "need attention", "critical","!")
            with rk2: kpi_card("At-Risk Units", int(risk_data["Quantity"].sum()), "total qty","soon","📦")
            with rk3: kpi_card("Total Est. Loss", f"₹{total_risk:,.0f}", "potential loss","expired","₹")

            st.markdown("<br>", unsafe_allow_html=True)
            section("Estimated Loss by Category")
            st.markdown('<div class="glass">', unsafe_allow_html=True)
            loss_cat = risk_data.groupby("Category")["Est. Loss (₹)"].sum().sort_values(ascending=False)
            max_l = loss_cat.max()
            for cat, val in loss_cat.items():
                hbar(cat, val, max_l, CLR["critical"], " ₹")
            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            section("At-Risk Product Details")
            st.dataframe(
                risk_data[["Product Name","Category","Status","Quantity",
                           "Cost Per Unit","Est. Loss (₹)","Days Remaining","Risk Score"]]
                .sort_values("Days Remaining").reset_index(drop=True),
                use_container_width=True, hide_index=True,
                column_config={
                    "Risk Score":     st.column_config.ProgressColumn("Risk Score", min_value=0, max_value=100, format="%d"),
                    "Est. Loss (₹)":  st.column_config.NumberColumn(format="₹%.2f"),
                    "Days Remaining": st.column_config.NumberColumn("Days Left", format="%d d"),
                },
            )

    with tab_dl:
        st.markdown("<br>", unsafe_allow_html=True)
        today_str = date.today().strftime("%Y%m%d")
        full_rpt = data[[
            "Product Name","Category","Batch Number","Quantity",
            "Manufacturing Date","Expiry Date","Cost Per Unit",
            "Days Remaining","Status","Risk Score","Est. Loss (₹)"
        ]].sort_values("Days Remaining").reset_index(drop=True)
        risk_rpt = full_rpt[full_rpt["Status"].isin(["Expired","Critical","Expiring Soon"])]

        st.markdown('<div class="glass">', unsafe_allow_html=True)
        section("Export Options")
        d1,d2 = st.columns(2)
        with d1:
            st.markdown(f'<div style="font-size:0.82rem;color:{CLR["muted"]};margin-bottom:8px">Full inventory — {len(full_rpt)} rows</div>', unsafe_allow_html=True)
            st.download_button("⬇  Full Inventory CSV",
                data=full_rpt.to_csv(index=False).encode(), file_name=f"inventory_{today_str}.csv",
                mime="text/csv", use_container_width=True)
        with d2:
            st.markdown(f'<div style="font-size:0.82rem;color:{CLR["muted"]};margin-bottom:8px">At-risk only — {len(risk_rpt)} rows</div>', unsafe_allow_html=True)
            st.download_button("⬇  At-Risk Products CSV",
                data=risk_rpt.to_csv(index=False).encode(), file_name=f"at_risk_{today_str}.csv",
                mime="text/csv", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# PAGE 6 ── SEARCH & FILTER
# ══════════════════════════════════════════════════════════
elif page == "🔍  Search & Filter":
    page_header("Search & Filter", "Find products by name, category, or risk status")

    data = gdf()

    st.markdown('<div class="glass">', unsafe_allow_html=True)
    f1,f2,f3 = st.columns([2,1.5,1.5])
    with f1:
        srch = st.text_input("🔍  Product name or batch number",
                             placeholder="e.g. Milk, B001…", label_visibility="collapsed")
    with f2:
        cats = ["All"] + sorted(data["Category"].dropna().unique().tolist())
        cat_sel = st.selectbox("Category", cats, label_visibility="collapsed")
    with f3:
        status_sel = st.selectbox("Status", ["All"]+STATUS_ORDER, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

    # Apply filters
    filt = data.copy()
    if srch:
        filt = filt[
            filt["Product Name"].str.contains(srch, case=False, na=False) |
            filt["Batch Number"].str.contains(srch, case=False, na=False)
        ]
    if cat_sel != "All":   filt = filt[filt["Category"]==cat_sel]
    if status_sel != "All": filt = filt[filt["Status"]==status_sel]
    filt = filt.sort_values("Days Remaining").reset_index(drop=True)

    st.markdown(
        f'<div style="font-size:0.82rem;color:{CLR["muted"]};margin:12px 0 8px">'
        f'<b style="color:{CLR["accent"]}">{len(filt)}</b> product(s) found</div>',
        unsafe_allow_html=True,
    )

    if filt.empty:
        st.info("No products match your filters.")
    else:
        SHOW = ["Product Name","Category","Batch Number","Quantity",
                "Expiry Date","Days Remaining","Status","Risk Score","Est. Loss (₹)"]

        def row_color(row):
            s = {
                "Safe":          "background-color:#052e16;color:#10b981",
                "Expiring Soon": "background-color:#451a03;color:#f59e0b",
                "Critical":      "background-color:#450a0a;color:#ef4444",
                "Expired":       "background-color:#4a044e;color:#be185d",
            }.get(row["Status"], "")
            return [s]*len(row)

        styled = filt[SHOW].style.apply(row_color, axis=1)
        st.dataframe(
            styled, use_container_width=True, hide_index=True,
            column_config={
                "Risk Score":     st.column_config.ProgressColumn("Risk Score", min_value=0, max_value=100, format="%d"),
                "Days Remaining": st.column_config.NumberColumn("Days Left", format="%d d"),
                "Est. Loss (₹)":  st.column_config.NumberColumn(format="₹%.2f"),
            },
        )
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button(
            "⬇  Download Filtered CSV",
            data=filt[SHOW].to_csv(index=False).encode(),
            file_name="filtered_inventory.csv", mime="text/csv",
        )