"""
Management View · Spyne AR
Read-only executive scorecard pulling live from Google Sheets.
Run: streamlit run management_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
from io import BytesIO
from typing import Optional
import json

try:
    from streamlit_autorefresh import st_autorefresh
    _AUTOREFRESH_AVAILABLE = True
except ImportError:
    _AUTOREFRESH_AVAILABLE = False

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Management View · Spyne AR",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.block-container { padding-top: 1.2rem; }
div[data-testid="stMetric"] {
    background: var(--secondary-background-color);
    border: 1px solid rgba(128,128,128,0.18);
    border-radius: 10px;
    padding: 14px 18px;
}
div[data-testid="stMetric"] label {
    font-size: 12px !important; font-weight: 600 !important;
    text-transform: uppercase; letter-spacing: 0.05em; opacity: 0.65;
}
div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
    font-size: 26px !important; font-weight: 700 !important;
}
.rag-box {
    border-radius: 8px; padding: 12px 16px; text-align: center;
    font-weight: 700; font-size: 15px; margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
COLUMN_ALIASES = {
    "Entity Name":               ["entity name", "entity"],
    "EnterprisesID":             ["enterprisesid", "enterprise id", "enterprise_id", "eid"],
    "customer_name":             ["customer name", "customer_name", "customername", "client name", "client_name"],
    "Status":                    ["status", "invoice status"],
    "customer_status":           ["customer status", "customer_status", "client status"],
    "invoice_number":            ["invoice number", "invoice_number", "invoice no", "invoice no.", "inv number", "inv no"],
    "Product":                   ["product", "product name", "plan"],
    "date":                      ["date", "invoice date", "invoice_date"],
    "due_date":                  ["due date", "due_date", "payment due", "payment due date"],
    "email":                     ["email", "email address", "billing email"],
    "country":                   ["country", "region"],
    "created_by":                ["created by", "created_by"],
    "currency_code":             ["currency code", "currency_code", "currency"],
    "total":                     ["total", "invoice total", "gross amount"],
    "balance":                   ["balance", "remaining balance"],
    "Outstanding":               ["outstanding", "outstanding amount"],
    "last_payment_date":         ["last payment date", "last_payment_date", "last payment"],
    "Billing Terms":             ["billing terms", "billing_terms", "payment terms"],
    "Service Type":              ["service type", "service_type"],
    "Service_period_Start_date": ["service period start date", "service_period_start_date",
                                  "service start date", "service start", "start date"],
    "Service_period_End_date":   ["service period end date", "service_period_end_date",
                                  "service end date", "service end", "end date"],
    "Final USD":                 ["final usd", "final_usd", "amount usd", "usd amount", "outstanding usd"],
    "CSM":                       ["csm", "customer success manager", "account manager", "am"],
    "CSM Email":                 ["csm email", "csm_email", "csm email address"],
}

BUCKET_ORDER = ["0-15", "16-30", "31-45", "46-60", "61-90", "90+"]
RAG_COLORS   = {"Red": "#ef4444", "Amber": "#f59e0b", "Green": "#10b981"}
REFRESH_MS   = 5 * 60 * 1000  # 5 minutes


# ── Shared helpers ─────────────────────────────────────────────────────────────
def remap_columns(df: pd.DataFrame) -> pd.DataFrame:
    lookup = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            lookup[alias.lower().strip()] = canonical
    rename_map = {}
    for col in df.columns:
        key = col.lower().strip()
        if key in lookup and col != lookup[key]:
            rename_map[col] = lookup[key]
    return df.rename(columns=rename_map)


DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1pY_hPKVa8A-d6kbCnsuRdns4CiuRTh1QaIJRf5-ppOI/edit"

def _sheet_id_from_url(url: str) -> str:
    """Extract sheet ID from a Google Sheets URL or return as-is if already an ID."""
    import re
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    return m.group(1) if m else url.strip()

def load_from_gsheet_public(sheet_url_or_id: str, gid: str = "0") -> pd.DataFrame:
    """Load a publicly shared Google Sheet via CSV export — no credentials needed."""
    sheet_id = _sheet_id_from_url(sheet_url_or_id)
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    df = pd.read_csv(csv_url, header=0, dtype=str, keep_default_na=False)
    df = df.dropna(how="all")
    return df


def compute_aging_rag(df: pd.DataFrame) -> pd.DataFrame:
    """Add Aging (int), Bucket (str), RAG (str) columns to df. Returns copy."""
    df = df.copy()

    # Normalise columns
    df.columns = [str(c).strip().replace("\xa0", " ") for c in df.columns]
    df = remap_columns(df)

    # Date parsing
    date_cols = ["due_date", "date", "last_payment_date",
                 "Service_period_Start_date", "Service_period_End_date"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Numeric parsing — strip commas first (e.g. "3,000.00" → 3000.00)
    for col in ["Final USD", "total", "balance", "Outstanding"]:
        if col in df.columns:
            s = df[col].astype(str).str.replace(",", "", regex=False)
            df[col] = pd.to_numeric(s, errors="coerce").fillna(0)

    # Aging — when Service_period_Start_date is missing, fall back to invoice date
    today = pd.Timestamp.today().normalize()
    if "date" in df.columns:
        inv_aging = (today - df["date"]).dt.days.clip(lower=0)
        if "Service_period_Start_date" in df.columns:
            svc_aging = (today - df["Service_period_Start_date"]).dt.days.clip(lower=0)
            inv_after = df["date"] > df["Service_period_Start_date"]
            # When svc_start is NaT, fall back to invoice date aging
            effective_svc_aging = svc_aging.fillna(inv_aging)
            df["Aging"] = np.where(inv_after, inv_aging, effective_svc_aging)
        else:
            df["Aging"] = inv_aging
    else:
        df["Aging"] = 0

    df["Aging"] = pd.to_numeric(df["Aging"], errors="coerce").fillna(0).astype(int)

    # Bucket
    def aging_bucket(a):
        if a <= 15:  return "0-15"
        if a <= 30:  return "16-30"
        if a <= 45:  return "31-45"
        if a <= 60:  return "46-60"
        if a <= 90:  return "61-90"
        return "90+"

    df["Bucket"] = df["Aging"].apply(aging_bucket)

    # RAG (customer-level then broadcast)
    OVER_90 = {"90+"}
    OVER_30 = {"31-45", "46-60", "61-90", "90+"}

    cust_col = next(
        (c for c in df.columns if c.lower().replace(" ", "_") == "customer_name"), None
    )
    if cust_col is None:
        cust_col = next((c for c in df.columns if "customer" in c.lower()), None)

    if cust_col:
        if cust_col != "customer_name":
            df = df.rename(columns={cust_col: "customer_name"})
        customer_buckets = df.groupby("customer_name")["Bucket"].apply(set)

        def rag_for(bset):
            if bset & OVER_90: return "Red"
            if bset & OVER_30: return "Amber"
            return "Green"

        rag_map = customer_buckets.apply(rag_for)
        df["RAG"] = df["customer_name"].map(rag_map)
    else:
        df["RAG"] = "Green"

    return df


def fmt_usd(val):
    if abs(val) >= 1_000_000:
        return f"${val/1_000_000:.2f}M"
    if abs(val) >= 1_000:
        return f"${val/1_000:.1f}K"
    return f"${val:,.0f}"


def get_gcp_creds() -> Optional[dict]:
    """Return service account dict from st.secrets or credentials.json (optional — only needed for private sheets)."""
    try:
        return dict(st.secrets["gcp_service_account"])
    except (KeyError, FileNotFoundError):
        pass
    try:
        with open("credentials.json") as f:
            data = json.load(f)
        if "gcp_service_account" in data:
            return data["gcp_service_account"]
        if "type" in data and data.get("type") == "service_account":
            return data
    except Exception:
        pass
    return None


def get_sheet_url() -> Optional[str]:
    try:
        return st.secrets["gsheet_url"]
    except (KeyError, FileNotFoundError):
        return None


# ── Auto-refresh ───────────────────────────────────────────────────────────────
if _AUTOREFRESH_AVAILABLE:
    st_autorefresh(interval=REFRESH_MS, key="mgmt_refresh")

# ── Header ─────────────────────────────────────────────────────────────────────
hcol1, hcol2 = st.columns([8, 2])
with hcol1:
    st.title("📊 Management View · Spyne AR")
with hcol2:
    refresh_btn = st.button("🔄 Refresh now", use_container_width=True)

st.markdown(f"*Last loaded: {datetime.now().strftime('%d %b %Y %H:%M:%S')}*")
st.divider()

# ── Sidebar: sheet URL override ────────────────────────────────────────────────
with st.sidebar:
    st.header("Configuration")
    sheet_override = st.text_input(
        "Google Sheet URL or ID (optional)",
        placeholder="Paste URL or sheet ID to override secrets",
    )

# ── Resolve sheet URL & creds ──────────────────────────────────────────────────
# ── Sidebar: sheet URL + Excel fallback ───────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.subheader("Or upload Excel")
    uploaded_file = st.file_uploader("Upload AR Excel file", type=["xlsx", "xls"])

# Sheet URL: sidebar override → secrets → default hardcoded
sheet_url = sheet_override.strip() if sheet_override.strip() else (
    get_sheet_url() or DEFAULT_SHEET_URL
)

# ── Load data ──────────────────────────────────────────────────────────────────
@st.cache_data(ttl=290, show_spinner="Fetching Google Sheet…")
def _fetch_public(url):
    raw = load_from_gsheet_public(url)
    raw = remap_columns(raw)
    # Strip commas on balance before filtering — matches hosted app's load_data order
    if "balance" in raw.columns:
        raw["balance"] = pd.to_numeric(
            raw["balance"].astype(str).str.replace(",", "", regex=False), errors="coerce"
        ).fillna(0)
        raw = raw[raw["balance"] > 0].copy()
    return compute_aging_rag(raw)

@st.cache_data(show_spinner="Loading Excel…")
def _fetch_excel(file_bytes):
    raw = pd.read_excel(BytesIO(file_bytes))
    raw.columns = [str(c).strip().replace("\xa0", " ") for c in raw.columns]
    raw = remap_columns(raw)
    return compute_aging_rag(raw)

if refresh_btn:
    st.cache_data.clear()

if uploaded_file:
    try:
        df = _fetch_excel(uploaded_file.read())
    except Exception as e:
        st.error(f"Failed to load Excel: {e}")
        st.stop()
else:
    try:
        df = _fetch_public(sheet_url)
    except Exception as e:
        st.error(f"Failed to load Google Sheet: {e}")
        st.stop()

if df.empty:
    st.warning("The Google Sheet appears to be empty or has no readable data.")
    st.stop()

# ── Value columns — match hosted dashboard exactly ────────────────────────────
# Hosted app uses Final USD for all USD metrics; Outstanding (INR) for INR display
val_col    = "Final USD"   if "Final USD"   in df.columns else None
inr_col    = "Outstanding" if "Outstanding" in df.columns else None
billed_col = val_col

today = pd.Timestamp.today().normalize()

# ── KPI strip (5 cards) — identical to hosted dashboard ───────────────────────
_total_usd  = df[val_col].sum()                                          if val_col else 0
_at_risk    = df[df["RAG"] == "Red"][val_col].sum()                      if (val_col and "RAG" in df.columns) else 0
_customers  = df["customer_name"].nunique()                               if "customer_name" in df.columns else 0
_invoices   = len(df)
_avg_aging  = int(df["Aging"].mean())                                    if ("Aging" in df.columns and len(df)) else 0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Outstanding (USD)", f"${_total_usd:,.0f}")
k2.metric("At Risk (Red)",           f"${_at_risk:,.0f}")
k3.metric("Customers",               f"{_customers:,}")
k4.metric("Invoices",                f"{_invoices:,}")
k5.metric("Avg Aging (days)",        f"{_avg_aging}")

st.divider()

# ── Overview 4-metric row — matches hosted Overview tab ───────────────────────
_total_inr  = df[inr_col].sum() if inr_col else 0
_n_csms     = df["CSM"].nunique() if "CSM" in df.columns else 0

ov1, ov2, ov3, ov4 = st.columns(4)
ov1.metric("Outstanding (INR)", f"₹{_total_inr:,.0f}")
ov2.metric("Invoices",          f"{_invoices:,}")
ov3.metric("Customers",         f"{_customers:,}")
ov4.metric("CSMs",              f"{_n_csms:,}")

st.divider()

# ── RAG summary boxes ──────────────────────────────────────────────────────────
if "RAG" in df.columns:
    total_inv = len(df)
    rag_counts = df["RAG"].value_counts()
    rcol1, rcol2, rcol3 = st.columns(3)

    for col_widget, rag, color in [
        (rcol1, "Red",   "#fee2e2"),
        (rcol2, "Amber", "#fef3c7"),
        (rcol3, "Green", "#d1fae5"),
    ]:
        cnt = int(rag_counts.get(rag, 0))
        pct = cnt / total_inv * 100 if total_inv else 0
        col_widget.markdown(
            f'<div class="rag-box" style="background:{color}; color:#1f2937;">'
            f'🔴 {rag}<br><span style="font-size:22px">{cnt}</span>'
            f'<br><span style="font-size:12px;opacity:0.7">{pct:.1f}% of invoices</span></div>'
            .replace("🔴 Red", "🔴 Red")
            .replace("🔴 Amber", "🟡 Amber")
            .replace("🔴 Green", "🟢 Green"),
            unsafe_allow_html=True,
        )

    st.divider()

# ── Charts ─────────────────────────────────────────────────────────────────────
ch1, ch2 = st.columns(2)

with ch1:
    st.subheader("Outstanding by Aging Bucket")
    if val_col and "Bucket" in df.columns:
        bucket_df = (
            df.groupby("Bucket", sort=False)[val_col]
            .sum()
            .reindex(BUCKET_ORDER)
            .fillna(0)
            .reset_index()
        )
        bucket_df.columns = ["Bucket", "Outstanding"]
        fig1 = px.bar(
            bucket_df, x="Bucket", y="Outstanding",
            text_auto=",.0f",
            color="Bucket",
            color_discrete_sequence=px.colors.sequential.RdBu_r,
        )
        fig1.update_layout(showlegend=False, xaxis_title="", yaxis_title="USD",
                           margin=dict(t=20, b=0))
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.info("Bucket or value column not available.")

with ch2:
    st.subheader("Monthly Billed vs Outstanding (Last 6 Months)")
    if "date" in df.columns and val_col:
        df["_month"] = df["date"].dt.to_period("M")
        last_6 = sorted(df["_month"].dropna().unique())[-6:]
        mdf = df[df["_month"].isin(last_6)].copy()
        mdf["_month_str"] = mdf["_month"].dt.strftime("%b %Y")
        trend = (
            mdf.groupby("_month_str")
            .agg(Billed=(billed_col, "sum"), Outstanding=(val_col, "sum"))
            .reset_index()
            .rename(columns={"_month_str": "Month"})
        )
        trend["_sort"] = pd.to_datetime(trend["Month"], format="%b %Y")
        trend = trend.sort_values("_sort").drop(columns="_sort")
        trend_m = trend.melt(id_vars="Month", var_name="Metric", value_name="USD")
        fig2 = px.line(
            trend_m, x="Month", y="USD", color="Metric", markers=True,
            color_discrete_map={"Billed": "#3b82f6", "Outstanding": "#ef4444"},
        )
        fig2.update_layout(xaxis_title="", yaxis_title="USD",
                           legend_title="", margin=dict(t=20, b=0))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Date or total column not available for trend.")

st.divider()

# ── CSM performance table ──────────────────────────────────────────────────────
st.subheader("CSM Performance")
if "CSM" in df.columns and val_col:
    csm_grp = df.groupby("CSM")
    inv_count_col = "invoice_number" if "invoice_number" in df.columns else val_col
    csm_tbl = csm_grp.agg(
        Accounts=(inv_count_col, "count"),
        Outstanding=(val_col, "sum"),
    ).reset_index()
    if "RAG" in df.columns:
        red_per_csm = df[df["RAG"] == "Red"].groupby("CSM")[val_col].sum().rename("At Risk (Red)")
        csm_tbl = csm_tbl.merge(red_per_csm, on="CSM", how="left").fillna({"At Risk (Red)": 0})
    if "At Risk (Red)" in csm_tbl.columns:
        csm_tbl["At Risk (Red)"] = csm_tbl["At Risk (Red)"].apply(fmt_usd)
    csm_tbl["Outstanding"] = csm_tbl["Outstanding"].apply(fmt_usd)
    st.dataframe(csm_tbl, use_container_width=True, hide_index=True)
else:
    st.info("CSM column not found.")

st.divider()

# ── Top 5 at-risk ──────────────────────────────────────────────────────────────
st.subheader("Top 5 At-Risk Accounts")
if "RAG" in df.columns and val_col:
    at_risk = df[df["RAG"] == "Red"].copy() if "Red" in df["RAG"].values else df.copy()
    if at_risk.empty:
        at_risk = df.copy()
    at_risk = at_risk.sort_values(val_col, ascending=False).head(5)
    show_cols = [c for c in ["customer_name", val_col, "Aging", "CSM", "RAG"] if c in at_risk.columns]
    risk_display = at_risk[show_cols].copy()

    def rag_badge(r):
        colors = {"Red": "🔴", "Amber": "🟡", "Green": "🟢"}
        return f"{colors.get(r, '')} {r}"

    if "RAG" in risk_display.columns:
        risk_display["RAG"] = risk_display["RAG"].apply(rag_badge)
    if val_col in risk_display.columns:
        risk_display[val_col] = risk_display[val_col].apply(fmt_usd)
    st.dataframe(risk_display, use_container_width=True, hide_index=True)
else:
    st.info("RAG column not available.")

