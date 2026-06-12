"""
Expiry Risk Management System for Small Retail Stores
======================================================
Run with:
    pip install streamlit pandas
    streamlit run app.py

Dependencies: streamlit, pandas — NOTHING ELSE.
"""

import streamlit as st
import pandas as pd
from datetime import date, timedelta
import os

# ──────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────
CSV_PATH          = "inventory.csv"
CRITICAL_DAYS     = 30
EXPIRING_SOON_DAYS= 60

CATEGORIES = [
    "Dairy", "Beverages", "Bakery", "Dry Goods",
    "Canned", "Snacks", "Cereals", "Deli", "Frozen", "Other"
]

STATUS_ORDER  = ["Expired", "Critical", "Expiring Soon", "Safe"]

# Emoji indicators used throughout (no external icon lib)
STATUS_EMOJI = {
    "Safe":          "✅",
    "Expiring Soon": "🟡",
    "Critical":      "🔴",
    "Expired":       "💀",
}

# ──────────────────────────────────────────────
# PAGE CONFIG  — must be very first Streamlit call
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="ExpiryGuard",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# MINIMAL CSS  — only layout polish, no charting
# ──────────────────────────────────────────────
st.markdown("""
<style>
/* Status alert strips */
.strip-expired  { border-left:4px solid #e05c8a; background:#2a0a1a; color:#f08fad;
                  padding:10px 14px; border-radius:0 8px 8px 0; margin:4px 0; }
.strip-critical { border-left:4px solid #ff6b47; background:#3d0f00; color:#ffb39a;
                  padding:10px 14px; border-radius:0 8px 8px 0; margin:4px 0; }
.strip-soon     { border-left:4px solid #f5a623; background:#3d2d00; color:#f5d08a;
                  padding:10px 14px; border-radius:0 8px 8px 0; margin:4px 0; }
.strip-info     { border-left:4px solid #3fa3f5; background:#0d2233; color:#7ec8f5;
                  padding:10px 14px; border-radius:0 8px 8px 0; margin:4px 0; }
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────
# DATA HELPERS
# ──────────────────────────────────────────────

def generate_sample_data() -> pd.DataFrame:
    """Bootstrap sample data so the app is never empty on first run."""
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
        "Manufacturing Date","Expiry Date","Cost Per Unit"
    ])
    df["Manufacturing Date"] = pd.to_datetime(df["Manufacturing Date"]).dt.date
    df["Expiry Date"]        = pd.to_datetime(df["Expiry Date"]).dt.date
    return df


def load_data() -> pd.DataFrame:
    """Load from CSV, or seed with sample data if none exists."""
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
    if days < 0:                    return "Expired"
    if days < CRITICAL_DAYS:        return "Critical"
    if days < EXPIRING_SOON_DAYS:   return "Expiring Soon"
    return "Safe"


def enrich(df: pd.DataFrame) -> pd.DataFrame:
    """Attach computed columns without touching the stored CSV."""
    today = date.today()
    df = df.copy()
    df["Days Remaining"] = pd.to_datetime(df["Expiry Date"]).apply(
        lambda d: (d.date() - today).days
    )
    df["Status"]         = df["Days Remaining"].apply(classify_status)
    # Risk score 0-100: higher = more urgent
    df["Risk Score"]     = df["Days Remaining"].apply(
        lambda d: 100 if d < 0 else max(0, round(100 - (d / EXPIRING_SOON_DAYS) * 100))
    )
    df["Est. Loss (₹)"]  = (df["Quantity"] * df["Cost Per Unit"]).round(2)
    return df


# ──────────────────────────────────────────────
# SESSION STATE
# ──────────────────────────────────────────────
if "data"       not in st.session_state:
    st.session_state.data       = load_data()
if "edit_idx"   not in st.session_state:
    st.session_state.edit_idx   = None


def df() -> pd.DataFrame:
    """Always return freshly enriched data."""
    return enrich(st.session_state.data)


# ──────────────────────────────────────────────
# SIDEBAR NAVIGATION
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🛡️ ExpiryGuard")
    st.caption("Retail Expiry Risk Management")
    st.divider()

    page = st.radio(
        "Go to",
        [
            "📊 Dashboard",
            "➕ Product Entry",
            "⚠️ Expiry Risk Analysis",
            "🔔 Alerts",
            "📈 Reports & Analytics",
            "🔍 Search & Filter",
        ],
        label_visibility="collapsed",
    )

    st.divider()
    st.caption(f"📅 Today: **{date.today().strftime('%d %b %Y')}**")
    st.caption(f"📦 Products: **{len(st.session_state.data)}**")

    if st.button("🔄 Reload from CSV"):
        st.session_state.data = load_data()
        st.success("Reloaded!")


# ══════════════════════════════════════════════
# PAGE 1 — DASHBOARD
# ══════════════════════════════════════════════
if page == "📊 Dashboard":
    st.title("📊 Inventory Dashboard")

    data = df()
    total    = len(data)
    safe     = int((data["Status"] == "Safe").sum())
    soon     = int((data["Status"] == "Expiring Soon").sum())
    critical = int((data["Status"] == "Critical").sum())
    expired  = int((data["Status"] == "Expired").sum())
    at_risk_loss = float(
        data[data["Status"].isin(["Expired","Critical","Expiring Soon"])]["Est. Loss (₹)"].sum()
    )

    # ── KPI row ──────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📦 Total Products",   total)
    c2.metric("✅ Safe",              safe)
    c3.metric("🟡 Expiring Soon",     soon)
    c4.metric("🔴 Critical",          critical)
    c5.metric("💀 Expired",           expired)

    st.divider()

    # ── Status distribution bar chart (native) ──
    left, right = st.columns(2)

    with left:
        st.subheader("Status Distribution")
        status_counts = (
            data["Status"]
            .value_counts()
            .reindex(STATUS_ORDER)
            .fillna(0)
            .rename("Products")
        )
        st.bar_chart(status_counts, color="#3fa3f5", height=260)

    # ── Category breakdown (native) ──
    with right:
        st.subheader("Products per Category")
        cat_counts = (
            data.groupby("Category")["Product Name"]
            .count()
            .sort_values(ascending=False)
            .rename("Count")
        )
        st.bar_chart(cat_counts, color="#f5a623", height=260)

    st.divider()

    # ── Top 10 closest to expiry ──
    st.subheader("🔺 Top 10 Closest to Expiry")
    top10 = (
        data[data["Days Remaining"] >= 0]
        .sort_values("Days Remaining")
        .head(10)[["Product Name","Category","Days Remaining","Quantity","Status","Risk Score"]]
        .reset_index(drop=True)
    )
    if top10.empty:
        st.info("No active products found.")
    else:
        # Use st.dataframe with column_config for a progress bar on risk score
        st.dataframe(
            top10,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Risk Score": st.column_config.ProgressColumn(
                    "Risk Score",
                    help="0 = no risk, 100 = expired",
                    min_value=0, max_value=100,
                    format="%d",
                ),
                "Days Remaining": st.column_config.NumberColumn(
                    "Days Left", format="%d days"
                ),
            },
        )

    st.divider()
    # Financial risk callout
    st.warning(
        f"💸 **Estimated value at risk** (Critical + Expiring Soon + Expired): "
        f"**₹{at_risk_loss:,.2f}**"
    )


# ══════════════════════════════════════════════
# PAGE 2 — PRODUCT ENTRY
# ══════════════════════════════════════════════
elif page == "➕ Product Entry":
    st.title("➕ Product Entry")

    tab_add, tab_manage = st.tabs(["Add New Product", "Manage Existing Products"])

    # ── ADD ────────────────────────────────────
    with tab_add:
        st.subheader("New Product Details")

        with st.form("add_form", clear_on_submit=True):
            r1a, r1b = st.columns(2)
            name    = r1a.text_input("Product Name *")
            cat     = r1b.selectbox("Category *", CATEGORIES)

            r2a, r2b = st.columns(2)
            batch   = r2a.text_input("Batch Number *")
            qty     = r2b.number_input("Quantity *", min_value=1, step=1, value=1)

            r3a, r3b, r3c = st.columns(3)
            mfg_dt  = r3a.date_input("Manufacturing Date *", value=date.today()-timedelta(30))
            exp_dt  = r3b.date_input("Expiry Date *",        value=date.today()+timedelta(90))
            cost    = r3c.number_input("Cost Per Unit (₹) *", min_value=0.0, step=1.0,
                                       value=10.0, format="%.2f")

            submitted = st.form_submit_button("💾 Add Product")

        if submitted:
            errors = []
            if not name.strip():    errors.append("Product Name is required.")
            if not batch.strip():   errors.append("Batch Number is required.")
            if exp_dt <= mfg_dt:    errors.append("Expiry date must be after manufacturing date.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                new_row = pd.DataFrame([{
                    "Product Name":       name.strip(),
                    "Category":           cat,
                    "Batch Number":       batch.strip(),
                    "Quantity":           int(qty),
                    "Manufacturing Date": mfg_dt,
                    "Expiry Date":        exp_dt,
                    "Cost Per Unit":      float(cost),
                }])
                st.session_state.data = pd.concat(
                    [st.session_state.data, new_row], ignore_index=True
                )
                save_data(st.session_state.data)
                st.success(f"✅ '{name}' added successfully!")

    # ── MANAGE ─────────────────────────────────
    with tab_manage:
        raw = df()
        if raw.empty:
            st.info("No products yet. Add one in the 'Add New Product' tab.")
        else:
            search = st.text_input("Filter by name", placeholder="Type to narrow list…")
            display = raw.copy()
            if search:
                display = display[
                    display["Product Name"].str.contains(search, case=False, na=False)
                ]

            for i, row in display.iterrows():
                em  = STATUS_EMOJI.get(row["Status"], "")
                col_info, col_ed, col_del = st.columns([7, 1, 1])

                with col_info:
                    st.markdown(
                        f"**{row['Product Name']}** &nbsp; {em} {row['Status']}  \n"
                        f"Cat: {row['Category']} | Batch: {row['Batch Number']} | "
                        f"Qty: {row['Quantity']} | Expires: {row['Expiry Date']} | "
                        f"Days left: {row['Days Remaining']}"
                    )

                with col_ed:
                    if st.button("✏️", key=f"ed_{i}"):
                        st.session_state.edit_idx = i

                with col_del:
                    if st.button("🗑️", key=f"del_{i}"):
                        st.session_state.data = (
                            st.session_state.data.drop(index=i).reset_index(drop=True)
                        )
                        save_data(st.session_state.data)
                        st.success("Deleted.")
                        st.rerun()

            # Inline edit form
            eidx = st.session_state.edit_idx
            if eidx is not None and eidx in st.session_state.data.index:
                r = st.session_state.data.loc[eidx]
                st.divider()
                st.subheader(f"Editing: {r['Product Name']}")

                with st.form("edit_form"):
                    ea, eb = st.columns(2)
                    e_name  = ea.text_input("Product Name", value=r["Product Name"])
                    cat_idx = CATEGORIES.index(r["Category"]) if r["Category"] in CATEGORIES else 0
                    e_cat   = eb.selectbox("Category", CATEGORIES, index=cat_idx)

                    ec, ed = st.columns(2)
                    e_batch = ec.text_input("Batch Number", value=r["Batch Number"])
                    e_qty   = ed.number_input("Quantity", min_value=1, value=int(r["Quantity"]))

                    ee, ef, eg = st.columns(3)
                    e_mfg   = ee.date_input("Mfg Date",    value=r["Manufacturing Date"])
                    e_exp   = ef.date_input("Expiry Date", value=r["Expiry Date"])
                    e_cost  = eg.number_input("Cost/Unit", min_value=0.0,
                                              value=float(r["Cost Per Unit"]), format="%.2f")

                    save_btn = st.form_submit_button("💾 Save Changes")

                if save_btn:
                    if e_exp <= e_mfg:
                        st.error("Expiry date must be after manufacturing date.")
                    else:
                        st.session_state.data.at[eidx, "Product Name"]       = e_name
                        st.session_state.data.at[eidx, "Category"]           = e_cat
                        st.session_state.data.at[eidx, "Batch Number"]       = e_batch
                        st.session_state.data.at[eidx, "Quantity"]           = e_qty
                        st.session_state.data.at[eidx, "Manufacturing Date"] = e_mfg
                        st.session_state.data.at[eidx, "Expiry Date"]        = e_exp
                        st.session_state.data.at[eidx, "Cost Per Unit"]      = e_cost
                        save_data(st.session_state.data)
                        st.session_state.edit_idx = None
                        st.success("Changes saved!")
                        st.rerun()


# ══════════════════════════════════════════════
# PAGE 3 — EXPIRY RISK ANALYSIS
# ══════════════════════════════════════════════
elif page == "⚠️ Expiry Risk Analysis":
    st.title("⚠️ Expiry Risk Analysis")

    data = df()

    # KPI strip
    c1, c2, c3, c4 = st.columns(4)
    for col, status in zip([c1,c2,c3,c4], STATUS_ORDER):
        n = int((data["Status"] == status).sum())
        col.metric(f"{STATUS_EMOJI[status]} {status}", n)

    st.divider()

    # Days-remaining distribution using Streamlit's native histogram via bar_chart
    st.subheader("Days-Remaining Distribution (bucketed)")
    active = data[data["Days Remaining"].between(-30, 180)].copy()
    if not active.empty:
        active["Bucket"] = pd.cut(
            active["Days Remaining"],
            bins=[-30, 0, 15, 30, 45, 60, 90, 120, 180],
            labels=["< 0","0-15","15-30","30-45","45-60","60-90","90-120","120-180"],
        )
        bucket_counts = (
            active.groupby("Bucket", observed=True)["Product Name"]
            .count()
            .rename("Products")
        )
        st.bar_chart(bucket_counts, color="#f5a623", height=220)
    st.caption("Threshold lines: 0 = Expired | 30 days = Critical | 60 days = Expiring Soon")

    st.divider()

    # Per-status colour-coded tables
    SHOW_COLS = [
        "Product Name","Category","Batch Number","Quantity",
        "Expiry Date","Days Remaining","Risk Score","Est. Loss (₹)"
    ]

    for status in STATUS_ORDER:
        subset = data[data["Status"] == status].sort_values("Days Remaining").reset_index(drop=True)
        if subset.empty:
            continue
        em = STATUS_EMOJI[status]
        st.subheader(f"{em} {status} — {len(subset)} product(s)")
        st.dataframe(
            subset[SHOW_COLS],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Risk Score": st.column_config.ProgressColumn(
                    "Risk Score", min_value=0, max_value=100, format="%d"
                ),
                "Days Remaining": st.column_config.NumberColumn("Days Left", format="%d d"),
                "Est. Loss (₹)":  st.column_config.NumberColumn("Est. Loss (₹)", format="₹%.2f"),
            },
        )


# ══════════════════════════════════════════════
# PAGE 4 — ALERTS
# ══════════════════════════════════════════════
elif page == "🔔 Alerts":
    st.title("🔔 Alerts Centre")

    data     = df()
    expired  = data[data["Status"] == "Expired"].sort_values("Days Remaining")
    critical = data[data["Status"] == "Critical"].sort_values("Days Remaining")
    soon     = data[data["Status"] == "Expiring Soon"].sort_values("Days Remaining")

    # ── Expired ───────────────────────────────
    st.subheader(f"💀 Expired  ({len(expired)})")
    if expired.empty:
        st.markdown(
            '<div class="strip-info">No expired products — great job!</div>',
            unsafe_allow_html=True,
        )
    else:
        for _, row in expired.iterrows():
            st.markdown(
                f'<div class="strip-expired">'
                f'<strong>{row["Product Name"]}</strong> (Batch {row["Batch Number"]}) — '
                f'expired <strong>{abs(row["Days Remaining"])} days ago</strong> on {row["Expiry Date"]} | '
                f'Qty: {row["Quantity"]} | Loss: ₹{row["Est. Loss (₹)"]:.2f}'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.divider()

    # ── Critical ──────────────────────────────
    st.subheader(f"🔴 Critical — expires within 30 days  ({len(critical)})")
    if critical.empty:
        st.markdown(
            '<div class="strip-info">No critical products right now.</div>',
            unsafe_allow_html=True,
        )
    else:
        for _, row in critical.iterrows():
            st.markdown(
                f'<div class="strip-critical">'
                f'<strong>{row["Product Name"]}</strong> — expires in '
                f'<strong>{row["Days Remaining"]} days</strong> ({row["Expiry Date"]}) | '
                f'Qty: {row["Quantity"]} | Risk: {row["Risk Score"]}/100'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.divider()

    # ── Expiring Soon ─────────────────────────
    st.subheader(f"🟡 Expiring Soon — 30–60 days  ({len(soon)})")
    if soon.empty:
        st.markdown(
            '<div class="strip-info">Nothing expiring in the 30–60-day window.</div>',
            unsafe_allow_html=True,
        )
    else:
        for _, row in soon.iterrows():
            st.markdown(
                f'<div class="strip-soon">'
                f'<strong>{row["Product Name"]}</strong> — expires in '
                f'<strong>{row["Days Remaining"]} days</strong> ({row["Expiry Date"]}) | '
                f'Qty: {row["Quantity"]}'
                f'</div>',
                unsafe_allow_html=True,
            )

    st.divider()

    # Total risk callout
    total_loss = float(
        data[data["Status"].isin(["Expired","Critical","Expiring Soon"])]["Est. Loss (₹)"].sum()
    )
    count_risk = len(expired) + len(critical) + len(soon)
    st.warning(
        f"📊 **{count_risk} products** at risk — estimated inventory loss: **₹{total_loss:,.2f}**"
    )


# ══════════════════════════════════════════════
# PAGE 5 — REPORTS & ANALYTICS
# ══════════════════════════════════════════════
elif page == "📈 Reports & Analytics":
    st.title("📈 Reports & Analytics")

    data = df()

    tab_cat, tab_status, tab_risk, tab_dl = st.tabs([
        "📦 By Category",
        "📊 Status Breakdown",
        "💸 Financial Risk",
        "📥 Download",
    ])

    # ── By Category ───────────────────────────
    with tab_cat:
        st.subheader("Inventory by Category")
        cat_df = (
            data.groupby("Category")
            .agg(
                Products   =("Product Name", "count"),
                Total_Qty  =("Quantity",      "sum"),
                At_Risk    =("Status",        lambda x: x.isin(["Critical","Expiring Soon","Expired"]).sum()),
                Est_Loss   =("Est. Loss (₹)", "sum"),
            )
            .reset_index()
            .rename(columns={"Est_Loss": "Est. Loss (₹)"})
            .sort_values("Products", ascending=False)
        )
        st.dataframe(
            cat_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Est. Loss (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                "At_Risk":       st.column_config.NumberColumn("At Risk"),
            },
        )
        st.bar_chart(
            cat_df.set_index("Category")[["Total_Qty","At_Risk"]],
            height=280,
        )

    # ── Status Breakdown ──────────────────────
    with tab_status:
        st.subheader("Products by Expiry Status")
        status_df = (
            data.groupby("Status")
            .agg(Products=("Product Name","count"), Total_Qty=("Quantity","sum"))
            .reindex(STATUS_ORDER).fillna(0).astype(int).reset_index()
        )
        st.dataframe(status_df, use_container_width=True, hide_index=True)

        st.subheader("Product Count per Status")
        st.bar_chart(
            status_df.set_index("Status")["Products"],
            color="#3fa3f5", height=240,
        )

        st.subheader("Total Quantity per Status")
        st.bar_chart(
            status_df.set_index("Status")["Total_Qty"],
            color="#f5a623", height=240,
        )

    # ── Financial Risk ────────────────────────
    with tab_risk:
        st.subheader("Quantity & Value at Risk")

        risk_data = data[data["Status"].isin(["Expired","Critical","Expiring Soon"])].copy()

        if risk_data.empty:
            st.success("No products at risk right now! 🎉")
        else:
            total_risk = float(risk_data["Est. Loss (₹)"].sum())
            r1, r2, r3 = st.columns(3)
            r1.metric("At-Risk Products",  len(risk_data))
            r2.metric("At-Risk Qty Units", int(risk_data["Quantity"].sum()))
            r3.metric("Total Est. Loss",   f"₹{total_risk:,.2f}")

            st.divider()

            # Loss per category bar chart
            st.subheader("Estimated Loss by Category")
            loss_cat = (
                risk_data.groupby("Category")["Est. Loss (₹)"]
                .sum()
                .sort_values(ascending=False)
            )
            st.bar_chart(loss_cat, color="#e05c8a", height=240)

            # Full at-risk table
            st.subheader("At-Risk Product Details")
            st.dataframe(
                risk_data[[
                    "Product Name","Category","Status","Quantity",
                    "Cost Per Unit","Est. Loss (₹)","Days Remaining","Risk Score"
                ]].sort_values("Days Remaining").reset_index(drop=True),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Risk Score":    st.column_config.ProgressColumn(
                                         "Risk Score", min_value=0, max_value=100, format="%d"),
                    "Est. Loss (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                    "Days Remaining":st.column_config.NumberColumn("Days Left", format="%d d"),
                },
            )

    # ── Download ──────────────────────────────
    with tab_dl:
        st.subheader("Export Reports as CSV")

        today_str = date.today().strftime("%Y%m%d")
        full_rpt  = data[[
            "Product Name","Category","Batch Number","Quantity",
            "Manufacturing Date","Expiry Date","Cost Per Unit",
            "Days Remaining","Status","Risk Score","Est. Loss (₹)"
        ]].sort_values("Days Remaining").reset_index(drop=True)

        risk_rpt  = full_rpt[full_rpt["Status"].isin(["Expired","Critical","Expiring Soon"])]

        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            st.download_button(
                "⬇️ Full Inventory CSV",
                data=full_rpt.to_csv(index=False).encode("utf-8"),
                file_name=f"inventory_{today_str}.csv",
                mime="text/csv",
                use_container_width=True,
            )
            st.caption(f"{len(full_rpt)} rows")

        with c_dl2:
            st.download_button(
                "⬇️ At-Risk Products CSV",
                data=risk_rpt.to_csv(index=False).encode("utf-8"),
                file_name=f"at_risk_{today_str}.csv",
                mime="text/csv",
                use_container_width=True,
            )
            st.caption(f"{len(risk_rpt)} rows")


# ══════════════════════════════════════════════
# PAGE 6 — SEARCH & FILTER
# ══════════════════════════════════════════════
elif page == "🔍 Search & Filter":
    st.title("🔍 Search & Filter")

    data = df()

    f1, f2, f3 = st.columns([2, 1.5, 1.5])
    with f1:
        search_term = st.text_input("Search name or batch", placeholder="e.g. Milk, B001…")
    with f2:
        cats = ["All"] + sorted(data["Category"].dropna().unique().tolist())
        cat_sel = st.selectbox("Category", cats)
    with f3:
        status_sel = st.selectbox("Status", ["All"] + STATUS_ORDER)

    # Apply filters
    filtered = data.copy()
    if search_term:
        filtered = filtered[
            filtered["Product Name"].str.contains(search_term, case=False, na=False) |
            filtered["Batch Number"].str.contains(search_term, case=False, na=False)
        ]
    if cat_sel != "All":
        filtered = filtered[filtered["Category"] == cat_sel]
    if status_sel != "All":
        filtered = filtered[filtered["Status"] == status_sel]

    filtered = filtered.sort_values("Days Remaining").reset_index(drop=True)

    st.divider()
    st.markdown(f"**{len(filtered)}** product(s) found")

    if filtered.empty:
        st.info("No products match your filters.")
    else:
        SHOW_COLS = [
            "Product Name","Category","Batch Number","Quantity",
            "Expiry Date","Days Remaining","Status","Risk Score","Est. Loss (₹)"
        ]

        # Row-level colour highlighting via pandas Styler
        def highlight_row(row):
            style = {
                "Safe":          "background-color:#0d3d2e; color:#3dd68c",
                "Expiring Soon": "background-color:#3d2d00; color:#f5a623",
                "Critical":      "background-color:#3d0f00; color:#ff6b47",
                "Expired":       "background-color:#2a0a1a; color:#e05c8a",
            }.get(row["Status"], "")
            return [style] * len(row)

        styled = filtered[SHOW_COLS].style.apply(highlight_row, axis=1)
        st.dataframe(
            styled,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Risk Score":     st.column_config.ProgressColumn(
                                      "Risk Score", min_value=0, max_value=100, format="%d"),
                "Days Remaining": st.column_config.NumberColumn("Days Left", format="%d d"),
                "Est. Loss (₹)":  st.column_config.NumberColumn(format="₹%.2f"),
            },
        )

        st.download_button(
            "⬇️ Download Filtered Results",
            data=filtered[SHOW_COLS].to_csv(index=False).encode("utf-8"),
            file_name="filtered_inventory.csv",
            mime="text/csv",
        )