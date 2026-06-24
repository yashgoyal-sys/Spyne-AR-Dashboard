"""
Collections Dashboard · Spyne AR
Google-Sheets-connected collections view with invoice-level detail.
Run: streamlit run collections_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
from io import BytesIO
from typing import Optional
import json

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Collections Dashboard · Spyne AR",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded",
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
.reason-panel {
    background: var(--secondary-background-color);
    border-left: 3px solid #3b82f6;
    border-radius: 4px;
    padding: 10px 14px;
    margin-top: 8px;
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

CURR_SYM = {"INR": "₹", "USD": "$", "EUR": "€", "GBP": "£",
             "AUD": "A$", "CAD": "C$", "SGD": "S$", "AED": "د.إ"}
BUCKET_ORDER = ["0-15", "16-30", "31-45", "46-60", "61-90", "90+"]
RAG_COLORS   = {"Red": "#ef4444", "Amber": "#f59e0b", "Green": "#10b981"}


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


def load_from_gsheet(sheet_url_or_id: str, creds_dict: dict) -> pd.DataFrame:
    """Fetch first worksheet from a Google Sheet and return as DataFrame."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        raise ImportError("Run: pip install gspread google-auth")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)

    if sheet_url_or_id.startswith("http"):
        sh = gc.open_by_url(sheet_url_or_id)
    else:
        sh = gc.open_by_key(sheet_url_or_id)

    ws = sh.get_worksheet(0)
    records = ws.get_all_records(expected_headers=[], numericise_ignore=["all"])
    if not records:
        return pd.DataFrame()
    return pd.DataFrame(records)


def compute_aging_rag(df: pd.DataFrame) -> pd.DataFrame:
    """Add Aging (int), Bucket (str), RAG (str) columns to df."""
    df = df.copy()
    df.columns = [str(c).strip().replace("\xa0", " ") for c in df.columns]
    df = remap_columns(df)

    date_cols = ["due_date", "date", "last_payment_date",
                 "Service_period_Start_date", "Service_period_End_date"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in ["Final USD", "total", "balance", "Outstanding"]:
        if col in df.columns:
            s = df[col].astype(str).str.replace(",", "", regex=False)
            df[col] = pd.to_numeric(s, errors="coerce").fillna(0)

    today = pd.Timestamp.today().normalize()
    if "date" in df.columns and "Service_period_Start_date" in df.columns:
        inv_after = df["date"] > df["Service_period_Start_date"]
        df["Aging"] = np.where(
            inv_after,
            (today - df["date"]).dt.days,
            (today - df["Service_period_Start_date"]).dt.days,
        ).clip(min=0)
    elif "date" in df.columns:
        df["Aging"] = (today - df["date"]).dt.days.clip(lower=0)
    else:
        df["Aging"] = 0

    df["Aging"] = pd.to_numeric(df["Aging"], errors="coerce").fillna(0).astype(int)

    def aging_bucket(a):
        if a <= 15:  return "0-15"
        if a <= 30:  return "16-30"
        if a <= 45:  return "31-45"
        if a <= 60:  return "46-60"
        if a <= 90:  return "61-90"
        return "90+"

    df["Bucket"] = df["Aging"].apply(aging_bucket)

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


def get_password() -> str:
    try:
        return st.secrets["dashboard_password"]
    except (KeyError, FileNotFoundError):
        return "spyne@2024"


# ── Session state init ─────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "invoice_reasons" not in st.session_state:
    st.session_state.invoice_reasons = {}  # {invoice_number: reason_text}
if "selected_invoice" not in st.session_state:
    st.session_state.selected_invoice = None


# ── Login page ─────────────────────────────────────────────────────────────────
def login_page():
    st.title("💳 Collections Dashboard · Spyne AR")
    st.markdown("---")
    lcol, _, rcol = st.columns([1, 2, 1])
    with lcol:
        st.subheader("Login")
        with st.form("login_form"):
            pwd = st.text_input("Password", type="password", placeholder="Enter password")
            submitted = st.form_submit_button("Login", use_container_width=True)
        if submitted:
            if pwd == get_password():
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password. Please try again.")


if not st.session_state.authenticated:
    login_page()
    st.stop()


# ── Main dashboard ─────────────────────────────────────────────────────────────
st.title("💳 Collections Dashboard · Spyne AR")

# ── Sidebar: data source + filters ────────────────────────────────────────────
with st.sidebar:
    st.header("Data Source")
    data_source = st.radio("Source", ["Google Sheets", "Excel Upload"], horizontal=False)
    sheet_override = ""
    uploaded_file = None

    if data_source == "Google Sheets":
        sheet_override = st.text_input(
            "Sheet URL or ID (optional)",
            placeholder="Override secrets value",
        )
    else:
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

    if st.button("🔄 Refresh data"):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.button("🚪 Logout", on_click=lambda: st.session_state.update({"authenticated": False}))


# ── Load data ──────────────────────────────────────────────────────────────────
@st.cache_data(ttl=290, show_spinner="Fetching Google Sheet…")
def _fetch_gsheet(url, creds_json_str):
    creds = json.loads(creds_json_str)
    raw = load_from_gsheet(url, creds)
    return compute_aging_rag(raw)


@st.cache_data(show_spinner="Reading Excel…")
def _fetch_excel(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes))
    return compute_aging_rag(df)


df = None
if data_source == "Google Sheets":
    sheet_url = sheet_override.strip() if sheet_override.strip() else get_sheet_url()
    creds_dict = get_gcp_creds()
    if not sheet_url or not creds_dict:
        st.info(
            "### Google Sheets not configured\n\n"
            "Set `gsheet_url` and `[gcp_service_account]` in `.streamlit/secrets.toml`, "
            "or paste a Sheet URL in the sidebar.\n\n"
            "Alternatively, switch to **Excel Upload** in the sidebar.",
            icon="ℹ️",
        )
        st.stop()
    try:
        df = _fetch_gsheet(sheet_url, json.dumps(creds_dict))
    except Exception as e:
        st.error(f"Failed to load Google Sheet: {e}")
        st.stop()
else:
    if uploaded_file is None:
        st.info("Upload an Excel file in the sidebar to get started.")
        st.stop()
    try:
        df = _fetch_excel(uploaded_file.read())
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

if df is None or df.empty:
    st.warning("No data loaded. Please check your data source.")
    st.stop()


# ── Sidebar filters ────────────────────────────────────────────────────────────
def col(name):
    return df[name] if name in df.columns else pd.Series(dtype="object")

with st.sidebar:
    st.header("Filters")

    csm_opts = sorted(col("CSM").dropna().unique().tolist()) if "CSM" in df.columns else []
    csm_sel = st.multiselect("CSM", csm_opts)

    country_opts = sorted(col("country").dropna().unique().tolist()) if "country" in df.columns else []
    country_sel = st.multiselect("Country", country_opts)

    bucket_sel = st.multiselect("Aging Bucket", BUCKET_ORDER)

    rag_sel = st.multiselect("RAG Status", ["Red", "Amber", "Green"])

    currency_opts = sorted(col("currency_code").dropna().unique().tolist()) if "currency_code" in df.columns else []
    currency_sel = st.selectbox("Currency", ["All"] + currency_opts)


# ── Apply filters ──────────────────────────────────────────────────────────────
fdf = df.copy()
if csm_sel     and "CSM"           in fdf.columns: fdf = fdf[fdf["CSM"].isin(csm_sel)]
if country_sel and "country"       in fdf.columns: fdf = fdf[fdf["country"].isin(country_sel)]
if bucket_sel  and "Bucket"        in fdf.columns: fdf = fdf[fdf["Bucket"].isin(bucket_sel)]
if rag_sel     and "RAG"           in fdf.columns: fdf = fdf[fdf["RAG"].isin(rag_sel)]
if currency_sel != "All" and "currency_code" in fdf.columns:
    fdf = fdf[fdf["currency_code"] == currency_sel]


# ── Value column — Outstanding (INR) ÷ 92 → USD ──────────────────────────────
INR_TO_USD = 92.0

if "Outstanding" in fdf.columns:
    fdf = fdf.copy()
    fdf["Outstanding_USD"] = fdf["Outstanding"] / INR_TO_USD
    val_col = "Outstanding_USD"
elif "Final USD" in fdf.columns:
    val_col = "Final USD"
elif "balance" in fdf.columns:
    val_col = "balance"
else:
    val_col = None

# ── KPI row ────────────────────────────────────────────────────────────────────
total_outstanding = fdf[val_col].sum() if val_col else 0
avg_aging         = int(fdf["Aging"].mean()) if "Aging" in fdf.columns and len(fdf) else 0
red_count         = int((fdf["RAG"] == "Red").sum()) if "RAG" in fdf.columns else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Invoices Shown",      f"{len(fdf):,}")
k2.metric("Total Outstanding",   fmt_usd(total_outstanding) if val_col else "—")
k3.metric("Avg Aging (days)",    f"{avg_aging}")
k4.metric("Red Accounts",        f"{red_count:,}")

st.divider()

# ── Outstanding by CSM bar chart ───────────────────────────────────────────────
if "CSM" in fdf.columns and val_col:
    st.subheader("Outstanding by CSM")
    csm_df = (
        fdf.groupby("CSM")[val_col].sum()
        .reset_index()
        .sort_values(val_col, ascending=True)
    )
    fig = px.bar(
        csm_df, x=val_col, y="CSM", orientation="h",
        text_auto=",.0f", color=val_col,
        color_continuous_scale="RdYlGn_r",
        labels={val_col: "Outstanding (USD)", "CSM": ""},
    )
    fig.update_layout(coloraxis_showscale=False, margin=dict(t=10, b=0))
    st.plotly_chart(fig, use_container_width=True)
    st.divider()

# ── Invoice-level table ────────────────────────────────────────────────────────
st.subheader("Invoice Detail")

TABLE_COLS = [c for c in [
    "invoice_number", "customer_name", "CSM", "currency_code",
    "balance", "Outstanding", "Final USD",
    "Aging", "Bucket", "RAG",
    "due_date", "Status", "country",
] if c in fdf.columns]

display_df = fdf[TABLE_COLS].copy()

# Format RAG as emoji badge for display
if "RAG" in display_df.columns:
    rag_emoji = {"Red": "🔴 Red", "Amber": "🟡 Amber", "Green": "🟢 Green"}
    display_df["RAG"] = display_df["RAG"].map(lambda r: rag_emoji.get(r, r))

# Build column_config
col_cfg = {}
if val_col in display_df.columns:
    col_cfg[val_col] = st.column_config.NumberColumn(val_col, format="$%,.0f")
if "balance" in display_df.columns:
    col_cfg["balance"] = st.column_config.NumberColumn("Balance (FC)", format="%,.2f")
if "Aging" in display_df.columns:
    col_cfg["Aging"] = st.column_config.NumberColumn("Aging (days)", format="%d")
if "due_date" in display_df.columns:
    col_cfg["due_date"] = st.column_config.DateColumn("Due Date", format="DD MMM YYYY")

st.dataframe(
    display_df.reset_index(drop=True),
    use_container_width=True,
    column_config=col_cfg,
    hide_index=True,
)

# ── Reason tracking panel ──────────────────────────────────────────────────────
st.divider()
st.subheader("Reason / Notes Tracker")

inv_col = "invoice_number"
if inv_col in fdf.columns:
    invoice_list = sorted(fdf[inv_col].dropna().astype(str).unique().tolist())
    sel_inv = st.selectbox(
        "Select invoice to add/view notes",
        ["— select —"] + invoice_list,
    )

    if sel_inv and sel_inv != "— select —":
        existing = st.session_state.invoice_reasons.get(sel_inv, "")
        with st.container():
            st.markdown(
                f'<div class="reason-panel">📋 Notes for <strong>{sel_inv}</strong></div>',
                unsafe_allow_html=True,
            )
            new_reason = st.text_area(
                "Reason / Notes", value=existing,
                placeholder="Enter follow-up notes, payment promises, escalation details…",
                key=f"reason_{sel_inv}",
                height=100,
            )
            if st.button("Save notes", key=f"save_{sel_inv}"):
                st.session_state.invoice_reasons[sel_inv] = new_reason
                st.success(f"Notes saved for {sel_inv}")

        # Show all saved notes
        if st.session_state.invoice_reasons:
            with st.expander("View all saved notes"):
                notes_df = pd.DataFrame(
                    [(k, v) for k, v in st.session_state.invoice_reasons.items() if v],
                    columns=["Invoice", "Notes"],
                )
                if not notes_df.empty:
                    st.dataframe(notes_df, use_container_width=True, hide_index=True)
                else:
                    st.info("No notes saved yet.")
else:
    st.info("invoice_number column not found — reason tracking unavailable.")

st.divider()

# ── Download filtered view ─────────────────────────────────────────────────────
st.subheader("Export")

def to_excel_bytes(data: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="AR Collections")
    return buf.getvalue()

export_df = fdf[TABLE_COLS].copy()
# Add saved notes to export
if "invoice_number" in export_df.columns and st.session_state.invoice_reasons:
    export_df["Notes"] = export_df["invoice_number"].astype(str).map(
        st.session_state.invoice_reasons
    ).fillna("")

st.download_button(
    label="⬇️ Download filtered view as Excel",
    data=to_excel_bytes(export_df),
    file_name=f"ar_collections_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=False,
)
