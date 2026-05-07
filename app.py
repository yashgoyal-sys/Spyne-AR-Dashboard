import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
from io import BytesIO
import os
import re
import requests
import smtplib
import ssl
import time
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    from streamlit_autorefresh import st_autorefresh as _st_autorefresh
    _AUTOREFRESH_AVAILABLE = True
except ImportError:
    _AUTOREFRESH_AVAILABLE = False

# ── Logo loader (base64 — works locally + on Streamlit Cloud) ─────────────────
def _load_logo_b64(filename="spyne_logo.png") -> str | None:
    path = os.path.join(os.path.dirname(__file__), filename)
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except Exception:
        return None

_LOGO_B64 = _load_logo_b64()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AR Collections Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    /* Metric cards — visible on dark background */
    div[data-testid="stMetric"] {
        background: #1e2530;
        border: 1px solid #2e3a4e;
        border-radius: 10px;
        padding: 14px 18px;
    }
    div[data-testid="stMetric"] label {
        color: #94a3b8 !important;
        font-size: 13px !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #f1f5f9 !important;
        font-size: 26px !important;
        font-weight: 700 !important;
    }

    /* Tab labels */
    div[data-testid="stTabs"] button {
        font-size: 14px;
        font-weight: 600;
        color: #94a3b8;
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #60a5fa !important;
        border-bottom: 2px solid #60a5fa !important;
    }

    /* General text contrast */
    .block-container { padding-top: 1.5rem; }

    /* RAG inline badges */
    .rag-red   { color: #f87171; font-weight: 700; }
    .rag-amber { color: #fbbf24; font-weight: 700; }
    .rag-green { color: #34d399; font-weight: 700; }

    /* Dividers */
    hr { border-color: #2e3a4e !important; }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #111827;
        border-right: 1px solid #1f2937;
    }
    section[data-testid="stSidebar"] .stMarkdown p {
        color: #94a3b8;
    }

    /* Buttons */
    div[data-testid="stButton"] button[kind="primary"] {
        background: #2563eb;
        color: #fff;
        border: none;
        border-radius: 6px;
        font-weight: 600;
    }
    div[data-testid="stButton"] button[kind="primary"]:hover {
        background: #1d4ed8;
    }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────────────────────
REASON_CATEGORIES = [
    "Dispute – Invoice Amount",
    "Dispute – Service Delivered",
    "Customer Unresponsive",
    "Payment in Progress",
    "Credit Note / Adjustment Pending",
    "Contract / PO Issue",
    "Escalated to Management",
    "Promised to Pay",
    "Write-off Candidate",
    "Legal / Collections",
    "Other",
]

RAG_COLORS   = {"Red": "#ef4444", "Amber": "#f59e0b", "Green": "#10b981"}
BUCKET_ORDER = ["0-15", "16-30", "31-45", "46-60", "61-90", "90+"]
DB_PATH      = os.path.join(os.path.dirname(__file__), "reasons.db")
CREDS_PATH   = os.path.join(os.path.dirname(__file__), "credentials.json")

# ── Role definitions ──────────────────────────────────────────────────────────
ROLES = ["admin", "executor", "viewer", "csm", "management"]

ROLE_LABELS = {
    "admin":      "🔑 Admin",
    "executor":   "⚡ Executor",
    "viewer":     "👁 Viewer",
    "csm":        "🤝 CSM",
    "management": "📊 Management",
}

ROLE_COLORS = {
    "admin":      "#7c3aed",
    "executor":   "#2563eb",
    "viewer":     "#0891b2",
    "csm":        "#d97706",
    "management": "#059669",
}

# ─────────────────────────────────────────────────────────────────────────────
# Permission matrix
# ─────────────────────────────────────────────────────────────────────────────
# view_overview     → can see the Overview tab
# view_reasons      → can see the Reasons & Actions tab
# invoice_drilldown → can see the Invoice Drilldown tab
# send_reminders    → can send email reminders
# edit_reasons      → can save/delete reasons
# zoho_pull         → can pull live data from Zoho Books
# refresh_data      → Refresh Data button is shown
# download          → can download Excel exports
# manage_users      → can open User Management panel
# csm_filter        → data is automatically filtered to the user's assigned CSM
# ─────────────────────────────────────────────────────────────────────────────
ROLE_PERMISSIONS = {
    "admin": {
        "view_overview":     True,
        "view_reasons":      True,
        "invoice_drilldown": True,
        "send_reminders":    True,
        "edit_reasons":      True,
        "zoho_pull":         True,
        "refresh_data":      True,
        "download":          True,
        "manage_users":      True,
        "csm_filter":        False,
    },
    "executor": {
        "view_overview":     True,
        "view_reasons":      True,
        "invoice_drilldown": True,
        "send_reminders":    True,
        "edit_reasons":      True,
        "zoho_pull":         True,
        "refresh_data":      True,
        "download":          True,
        "manage_users":      False,
        "csm_filter":        False,
    },
    "viewer": {
        # Viewer sees: CSM Summary · Customer Summary · Invoice Drilldown · Reasons & Actions
        # Viewer cannot: Overview · Send Reminders · Refresh · Zoho pull · edit
        "view_overview":     False,
        "view_reasons":      True,
        "invoice_drilldown": True,
        "send_reminders":    False,
        "edit_reasons":      False,
        "zoho_pull":         False,
        "refresh_data":      False,
        "download":          True,
        "manage_users":      False,
        "csm_filter":        False,
    },
    "csm": {
        # CSM: same visible tabs as Viewer, but data auto-filtered to their assigned CSM
        "view_overview":     False,
        "view_reasons":      True,
        "invoice_drilldown": True,
        "send_reminders":    False,
        "edit_reasons":      False,
        "zoho_pull":         False,
        "refresh_data":      False,
        "download":          True,
        "manage_users":      False,
        "csm_filter":        True,   # ← data will be sliced to their CSM name
    },
    "management": {
        # Management sees: Overview · CSM Summary · Customer Summary
        # Management cannot: Invoice Drilldown · Reasons & Actions · Send Reminders
        "view_overview":     True,
        "view_reasons":      False,
        "invoice_drilldown": False,
        "send_reminders":    False,
        "edit_reasons":      False,
        "zoho_pull":         False,
        "refresh_data":      True,
        "download":          True,
        "manage_users":      False,
        "csm_filter":        False,
    },
}

# Default role assignments (override via secrets.toml [roles] or credentials.json)
_DEFAULT_ROLES = {
    "admin":   "admin",
    "yash":    "admin",
    "sukriti": "executor",
    "finance": "executor",
    "vijay":   "executor",
}

# ── CSM assignment helpers ────────────────────────────────────────────────────
def _load_csm_assignments() -> dict:
    """Return {username: csm_display_name}.
    Layers: credentials.json → st.secrets [csm_assignments] → SQLite app_users.csm_name.
    Falls back to using the username itself as the CSM name.
    """
    assignments: dict = {}
    # credentials.json
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH, "r") as f:
                data = json.load(f)
            assignments.update({k.lower(): v for k, v in data.get("csm_assignments", {}).items()})
        except Exception:
            pass
    # st.secrets
    try:
        if "csm_assignments" in st.secrets:
            assignments.update({k.lower(): v for k, v in st.secrets["csm_assignments"].items()})
    except Exception:
        pass
    # SQLite (csm_name column — added during User Management)
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute(
                "SELECT username, csm_name FROM app_users WHERE csm_name IS NOT NULL AND csm_name != ''"
            ).fetchall()
        for uname, cname in rows:
            assignments[uname.lower()] = cname
    except Exception:
        pass
    return assignments


def _get_csm_name_for_user(username: str) -> str | None:
    """Return the CSM display name assigned to this username.
    Falls back to the username itself (title-cased) if no explicit assignment.
    """
    assignments = _load_csm_assignments()
    return assignments.get(username.lower(), username.title())

def _load_roles() -> dict:
    """Return {username: role}.
    Layers (lowest → highest priority): defaults → credentials.json → st.secrets → SQLite DB.
    """
    roles = _DEFAULT_ROLES.copy()
    # credentials.json
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH, "r") as f:
                data = json.load(f)
            roles.update({k.lower(): v.lower() for k, v in data.get("roles", {}).items()})
        except Exception:
            pass
    # st.secrets
    try:
        if "roles" in st.secrets:
            roles.update({k.lower(): v.lower() for k, v in st.secrets["roles"].items()})
    except Exception:
        pass
    # SQLite (highest priority — set by admin via UI)
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute("SELECT username, role FROM app_users").fetchall()
        for uname, role in rows:
            roles[uname.lower()] = role.lower()
    except Exception:
        pass
    return roles

def _get_role(username: str) -> str:
    """Return the role for a given username. Defaults to 'viewer'."""
    roles = _load_roles()
    return roles.get(username.lower(), "viewer")

def _can(permission: str) -> bool:
    """Check if the current user's role has a given permission."""
    role = st.session_state.get("_role", "viewer")
    return ROLE_PERMISSIONS.get(role, ROLE_PERMISSIONS["viewer"]).get(permission, False)

# ── Login helpers ─────────────────────────────────────────────────────────────
import json
import hashlib

def _hash_pw(password: str) -> str:
    return hashlib.sha256(password.strip().encode()).hexdigest()

def _load_users() -> dict:
    """Return {username: hashed_password}.
    Priority (highest → lowest):
      1. app_users SQLite table  (created via User Management panel)
      2. st.secrets [users]      (Streamlit Cloud)
      3. credentials.json users  (local dev)
      4. Hard-coded fallback
    All layers are merged so every source contributes; DB always wins on conflict.
    """
    users = {}
    # 3. credentials.json (local dev)
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH, "r") as f:
                data = json.load(f)
            for uname, pw in data.get("users", {}).items():
                users[uname.lower()] = pw if len(str(pw)) == 64 else _hash_pw(str(pw))
        except Exception:
            pass
    # 2. st.secrets (Streamlit Cloud) — merge, secrets win over credentials.json
    try:
        if "users" in st.secrets:
            for uname, pw in st.secrets["users"].items():
                users[uname.lower()] = pw if len(str(pw)) == 64 else _hash_pw(str(pw))
    except Exception:
        pass
    # Hard-coded fallback (only if nothing loaded yet)
    if not users:
        _default_pw = _hash_pw("spyne@2024")
        users = {
            "admin":   _default_pw,
            "sukriti": _default_pw,
            "yash":    _default_pw,
            "finance": _default_pw,
            "vijay":   _hash_pw("vijay@2026"),
        }
    # 1. SQLite app_users — always wins (applied last, overwrites anything above)
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute("SELECT username, password_hash FROM app_users").fetchall()
        for uname, pw_hash in rows:
            users[uname.lower()] = pw_hash
    except Exception:
        pass
    return users

def _login_page():
    """Render the login page. Returns True if already authenticated."""
    if st.session_state.get("_authenticated"):
        return True

    # Centre the form
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        st.markdown(
            "<h2 style='text-align:center;margin-bottom:4px;'>🔐 AR Collections Dashboard</h2>"
            "<p style='text-align:center;color:#6b7280;margin-bottom:24px;'>Spyne.ai · Finance Team</p>",
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            username = st.text_input("Username").strip().lower()
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Login", use_container_width=True, type="primary")

        if submitted:
            users = _load_users()
            if username and password and users.get(username) == _hash_pw(password):
                st.session_state["_authenticated"] = True
                st.session_state["_username"]      = username
                st.session_state["_role"]          = _get_role(username)
                st.rerun()
            else:
                st.error("❌ Invalid username or password.")
    return False

def load_credentials():
    """Read saved credentials and populate session_state defaults.
    Priority: st.secrets (cloud) → credentials.json (local) → blank.
    Called once per session (guarded by _creds_loaded flag)."""
    if st.session_state.get("_creds_loaded"):
        return

    creds = {"gmail": {}, "zoho": {}}

    # ── 1. Try st.secrets (Streamlit Cloud) ───────────────────────────────────
    try:
        if "gmail" in st.secrets:
            creds["gmail"] = dict(st.secrets["gmail"])
        if "zoho" in st.secrets:
            creds["zoho"] = dict(st.secrets["zoho"])
    except Exception:
        pass

    # ── 2. Overlay / fallback with credentials.json (local dev) ───────────────
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH, "r") as f:
                local = json.load(f)
            # local file only fills keys not already set by st.secrets
            for section in ("gmail", "zoho"):
                for key, val in local.get(section, {}).items():
                    creds[section].setdefault(key, val)
        except Exception:
            pass

    # ── 3. Push into session_state ────────────────────────────────────────────
    for section in ("gmail", "zoho"):
        for key, val in creds[section].items():
            if key not in st.session_state:
                st.session_state[key] = val

    st.session_state["_creds_loaded"] = True


def save_credentials():
    """Write current sidebar credentials to disk (preserves users/roles already in file)."""
    # Load existing to preserve users & roles
    existing = {}
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH, "r") as f:
                existing = json.load(f)
        except Exception:
            pass
    creds = {
        "gmail": {
            "smtp_user":   st.session_state.get("smtp_user",   ""),
            "smtp_pass":   st.session_state.get("smtp_pass",   ""),
            "smtp_sender": st.session_state.get("smtp_sender", "finance@spyne.ai"),
        },
        "zoho": {
            "zoho_dc":            st.session_state.get("zoho_dc",            "US (.com)"),
            "zoho_org_id_1":      st.session_state.get("zoho_org_id_1",      ""),
            "zoho_org_id_2":      st.session_state.get("zoho_org_id_2",      ""),
            "zoho_client_id":     st.session_state.get("zoho_client_id",     ""),
            "zoho_client_secret": st.session_state.get("zoho_client_secret", ""),
            "zoho_refresh_token": st.session_state.get("zoho_refresh_token", ""),
        },
        # Preserve users & roles sections unchanged
        "users": existing.get("users", {}),
        "roles": existing.get("roles", {}),
    }
    with open(CREDS_PATH, "w") as f:
        json.dump(creds, f, indent=2)

# ── Database ──────────────────────────────────────────────────────────────────
# ── Email helpers ─────────────────────────────────────────────────────────────
FINANCE_CC = "finance@spyne.ai"

# ── Currency symbol map (shared) ──────────────────────────────────────────────
CURR_SYM = {"INR":"₹","USD":"$","EUR":"€","GBP":"£","AUD":"A$","CAD":"C$",
            "NZD":"NZ$","SGD":"S$","HKD":"HK$","NOK":"kr ","SEK":"kr ","DKK":"kr "}

TEMPLATES = {
    "Final Reminder":        "final",
    "Urgent Reminder":       "urgent",
    "Friendly Reminder":     "friendly",
    "Subscription Invoice":  "subscription",
}

# ── Zoho Books data-center map ────────────────────────────────────────────────
# (auth_host, api_host)
ZOHO_DC_MAP = {
    "US (.com)":        ("accounts.zoho.com",     "www.zohoapis.com"),
    "India (.in)":      ("accounts.zoho.in",      "www.zohoapis.in"),
    "EU (.eu)":         ("accounts.zoho.eu",       "www.zohoapis.eu"),
    "Australia (.au)":  ("accounts.zoho.com.au",   "www.zohoapis.com.au"),
    "Japan (.jp)":      ("accounts.zoho.jp",       "www.zohoapis.jp"),
}

# ── Shared invoice table builder ───────────────────────────────────────────────
def _invoice_table_html(invoices_df: pd.DataFrame) -> str:

    def _clean(val, fallback="—"):
        """Return a display-safe string; replace blank/NaN/NaT with fallback."""
        s = str(val) if val is not None else ""
        s = s.strip()
        return fallback if s in ("", "nan", "NaT", "None", "NaN") else s

    rows = ""
    for i, (_, r) in enumerate(invoices_df.iterrows()):
        bg       = "#f8fafc" if i % 2 == 0 else "#ffffff"
        currency = _clean(r.get("currency_code", ""), "INR")   # default INR if blank
        amount   = r.get("total", r.get("Final USD", 0)) or 0
        balance  = r.get("balance", amount) or 0
        inv_date = _clean(str(r.get("date", ""))[:10])
        svc_s    = _clean(str(r.get("Service_period_Start_date", ""))[:10])
        svc_e    = _clean(str(r.get("Service_period_End_date",   ""))[:10])

        bal_style = "color:#000000;font-weight:700;"

        # Pay Now button — only rendered when a valid payment link exists
        pay_link = str(r.get("payment_link", "") or "").strip()
        if pay_link and pay_link.lower().startswith("http"):
            pay_cell = (
                f'<a href="{pay_link}" target="_blank" '
                f'style="display:inline-block;background:#2563eb;color:#ffffff;'
                f'padding:6px 16px;border-radius:5px;font-size:12px;font-weight:700;'
                f'text-decoration:none;letter-spacing:0.03em;">Pay Now →</a>'
            )
        else:
            pay_cell = '<span style="color:#000000;font-size:12px;">—</span>'

        rows += f"""
        <tr style="background:{bg};">
          <td style="padding:9px 12px;font-size:13px;color:#000000;border-bottom:1px solid #e2e8f0;">{_clean(r.get("invoice_number",""))}</td>
          <td style="padding:9px 12px;font-size:13px;color:#000000;border-bottom:1px solid #e2e8f0;">{inv_date}</td>
          <td style="padding:9px 12px;font-size:13px;color:#000000;border-bottom:1px solid #e2e8f0;">{currency}</td>
          <td style="padding:9px 12px;font-size:13px;color:#000000;border-bottom:1px solid #e2e8f0;font-weight:600;">{fmt_amount(amount, currency)}</td>
          <td style="padding:9px 12px;font-size:13px;border-bottom:1px solid #e2e8f0;{bal_style}">{fmt_amount(balance, currency)}</td>
          <td style="padding:9px 12px;font-size:13px;color:#000000;border-bottom:1px solid #e2e8f0;">{svc_s}</td>
          <td style="padding:9px 12px;font-size:13px;color:#000000;border-bottom:1px solid #e2e8f0;">{svc_e}</td>
          <td style="padding:9px 12px;border-bottom:1px solid #e2e8f0;text-align:center;">{pay_cell}</td>
        </tr>"""

    # Total row — group by currency; if currency_code missing/blank default to INR
    if "balance" in invoices_df.columns:
        tmp = invoices_df.copy()
        tmp["_curr"] = (tmp.get("currency_code") if "currency_code" in tmp.columns
                        else "INR")
        tmp["_curr"] = tmp["_curr"].astype(str).str.strip().replace(
            {"": "INR", "nan": "INR", "NaN": "INR", "None": "INR"}
        )
        by_curr = (tmp.groupby("_curr")["balance"].sum()
                   .reset_index()
                   .sort_values("balance", ascending=False))
        total_str = "  |  ".join(
            fmt_amount(row["balance"], row["_curr"])
            for _, row in by_curr.iterrows()
        ) or "—"
    else:
        total_usd = invoices_df["Final USD"].sum() if "Final USD" in invoices_df.columns else 0
        total_str = fmt_amount(total_usd, "USD")

    rows += f"""
        <tr style="background:#f1f5f9;">
          <td colspan="4" style="padding:10px 12px;font-size:13px;font-weight:700;color:#000000;">Total Outstanding</td>
          <td colspan="4" style="padding:10px 12px;font-size:13px;font-weight:700;color:#000000;">{total_str}</td>
        </tr>"""

    header = """
      <tr style="background:#1a1a2e;">
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Invoice No.</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Invoice Date</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Currency</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Invoice Amount</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Outstanding Balance</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Service Start</th>
        <th style="padding:10px 12px;text-align:left;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Service End</th>
        <th style="padding:10px 12px;text-align:center;font-size:12px;color:#ffffff;text-transform:uppercase;font-weight:600;">Payment</th>
      </tr>"""

    return f"""
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border:1px solid #e2e8f0;border-radius:6px;overflow:hidden;margin-bottom:24px;">
      {header}{rows}
    </table>"""


def _email_wrapper(header_color: str, header_title: str,
                   banner_color: str, banner_text: str,
                   customer: str, body_html: str,
                   custom_note: str, csm: str) -> str:
    note_block = (f'<p style="color:#000000;font-size:14px;background:#fffbeb;'
                  f'border-left:4px solid #f59e0b;padding:12px 16px;'
                  f'border-radius:4px;margin-bottom:20px;">{custom_note}</p>'
                  if custom_note.strip() else "")

    banner = (f'<tr><td style="background:{banner_color};padding:12px 32px;">'
              f'<span style="color:#fff;font-size:13px;font-weight:700;">{banner_text}</span>'
              f'</td></tr>' if banner_text else "")

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;background:#f4f4f4;margin:0;padding:0;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 0;">
  <tr><td align="center">
    <table width="700" cellpadding="0" cellspacing="0"
           style="background:#fff;border-radius:8px;overflow:hidden;
                  box-shadow:0 2px 8px rgba(0,0,0,0.08);">
      <tr><td style="background:{header_color};padding:22px 32px;">
        <h2 style="color:#fff;margin:0;font-size:20px;">{header_title}</h2>
        <p style="color:rgba(255,255,255,0.7);margin:4px 0 0;font-size:13px;">Spyne.ai – Finance Team</p>
      </td></tr>
      {banner}
      <tr><td style="padding:30px 32px;">
        <p style="color:#000000;font-size:15px;margin:0 0 18px;">
          Dear <strong>{customer}</strong>,
        </p>
        {body_html}
        {note_block}
        <p style="color:#000000;font-size:14px;margin:20px 0 4px;">Best regards,</p>
        <p style="color:#000000;font-size:14px;font-weight:700;margin:0;">Finance Team – Spyne.ai</p>
      </td></tr>
      <tr><td style="background:#f8fafc;padding:14px 32px;border-top:1px solid #e2e8f0;">
        <p style="color:#000000;font-size:11px;margin:0;text-align:center;">
          Automated reminder from Spyne.ai Finance · Please ignore if payment has been made.
        </p>
      </td></tr>
    </table>
  </td></tr>
</table>
</body></html>"""


# ── Master email builder ───────────────────────────────────────────────────────
def build_email(template_key: str, customer: str, invoices_df: pd.DataFrame,
                csm: str, custom_note: str) -> tuple[str, str]:
    """
    Returns (subject, html) for any of the 4 templates.
    invoices_df must have: invoice_number, date, currency_code, Final USD,
                           Service_period_Start_date, Service_period_End_date, Aging
    """
    n         = len(invoices_df)
    max_aging = int(invoices_df["Aging"].max()) if "Aging" in invoices_df.columns else 0
    inv_table = _invoice_table_html(invoices_df)

    # Build total string from outstanding balance (same as the table Total row)
    if "balance" in invoices_df.columns and "currency_code" in invoices_df.columns:
        _by_curr = (invoices_df.groupby("currency_code")["balance"]
                    .sum().reset_index()
                    .sort_values("balance", ascending=False))
        total_str = "  |  ".join(
            fmt_amount(r["balance"], r["currency_code"])
            for _, r in _by_curr.iterrows()
        )
        # scalar fallback for subject lines (use largest currency bucket)
        total = _by_curr["balance"].iloc[0] if len(_by_curr) else 0
        total_subject = fmt_amount(total, _by_curr["currency_code"].iloc[0]) if len(_by_curr) else "$0"
    else:
        total     = invoices_df["Final USD"].sum() if "Final USD" in invoices_df.columns else 0
        total_str = fmt_amount(total, "USD")
        total_subject = fmt_amount(total, "USD")

    if template_key == "final":
        subject = f"⚠️ Final Payment Reminder – {customer} | {total_subject} Outstanding"
        body = f"""
        <p style="color:#dc2626;font-size:15px;font-weight:700;
                  background:#fef2f2;border-left:4px solid #dc2626;
                  padding:14px 16px;border-radius:4px;margin-bottom:20px;">
          This is a <strong>Final Reminder</strong> for your outstanding dues.
          Failure to make the payment within <strong>7 working days</strong> of this email
          may result in <strong>disruption of your Spyne.ai services</strong>.
        </p>
        <p style="color:#000000;font-size:14px;margin:0 0 18px;">
          We urge you to treat this matter with the utmost priority.
          Please find the details of all outstanding invoices below:
        </p>
        {inv_table}
        <p style="color:#000000;font-size:14px;margin:0 0 18px;">
          If you have already initiated the payment, we request you to
          <strong>share the transaction / UTR details</strong> by replying to this email
          so we can update our records accordingly.
        </p>
        <p style="color:#000000;font-size:14px;margin:0 0 8px;">
          For any queries, please contact your Customer Success Manager
          <strong>{csm}</strong> immediately.
        </p>"""
        html = _email_wrapper("#991b1b","⚠️ Final Payment Reminder",
                               "#dc2626", f"{n} invoice(s) | {total_subject} outstanding",
                               customer, body, custom_note, csm)

    elif template_key == "urgent":
        subject = f"🔴 Urgent: Payment Required – {customer} | {total_subject} Outstanding"
        body = f"""
        <p style="color:#000000;font-size:15px;margin:0 0 16px;">
          We would like to draw your <strong>immediate attention</strong> to the following
          outstanding invoices that require immediate action.
        </p>
        <p style="color:#000000;font-size:14px;margin:0 0 18px;">
          The total outstanding amount of <strong>{total_str}</strong> across
          <strong>{n} invoice(s)</strong> is pending. We request your immediate action
          to avoid any impact on your account.
        </p>
        {inv_table}
        <p style="color:#000000;font-size:14px;margin:0 0 8px;">
          We request you to arrange the payment at the earliest and confirm
          by replying to this email. Please contact your CSM
          <strong>{csm}</strong> if you need any assistance.
        </p>"""
        html = _email_wrapper("#b45309","🔴 Urgent Payment Reminder",
                               "#d97706", f"{n} invoice(s) | {total_subject} outstanding",
                               customer, body, custom_note, csm)

    elif template_key == "friendly":
        subject = f"Friendly Reminder: Outstanding Invoices – {customer}"
        body = f"""
        <p style="color:#000000;font-size:15px;margin:0 0 16px;">
          Hope this email finds you well!
        </p>
        <p style="color:#000000;font-size:14px;margin:0 0 18px;">
          This is a friendly reminder that you have <strong>{n} invoice(s)</strong>
          that are currently outstanding, totalling <strong>{total_str}</strong>.
          We would appreciate it if you could arrange the payment at your earliest convenience.
        </p>
        {inv_table}
        <p style="color:#000000;font-size:14px;margin:0 0 8px;">
          If you have any questions or need clarification on any of these invoices,
          please feel free to reach out to your Customer Success Manager
          <strong>{csm}</strong> or simply reply to this email. We are happy to help!
        </p>
        <p style="color:#000000;font-size:14px;margin:12px 0 0;">
          Thank you for your continued partnership with Spyne.ai 🙏
        </p>"""
        html = _email_wrapper("#065f46","Friendly Payment Reminder",
                               "#10b981", f"{n} invoice(s) | {total_subject} outstanding",
                               customer, body, custom_note, csm)

    else:  # subscription
        inv_row = invoices_df.iloc[0]
        sym = CURR_SYM.get(str(inv_row.get("currency_code","")).upper(), "")
        inv_amt = inv_row.get("total", inv_row.get("Final USD", 0))
        subject = f"New Subscription Invoice – {customer} | {sym}{inv_amt:,.0f}"
        body = f"""
        <p style="color:#000000;font-size:15px;margin:0 0 16px;">
          We hope you are enjoying Spyne.ai!
        </p>
        <p style="color:#000000;font-size:14px;margin:0 0 18px;">
          A new <strong>subscription invoice</strong> has been generated for your account.
          Please find the details below and arrange payment as per your billing terms.
        </p>
        {inv_table}
        <p style="color:#000000;font-size:14px;margin:0 0 8px;">
          If you have any questions regarding this invoice, please contact your
          Customer Success Manager <strong>{csm}</strong> or reply to this email.
        </p>"""
        html = _email_wrapper("#1e40af","New Subscription Invoice",
                               "#2563eb", f"Invoice generated for {customer}",
                               customer, body, custom_note, csm)

    return subject, html


def send_reminder(smtp_cfg: dict, to: str, cc_list: list[str],
                  subject: str, html: str,
                  attachments=None) -> str:
    """
    Send one email. Returns 'sent' or raises Exception.

    attachments: list of (filename: str, pdf_bytes: bytes) tuples, or None.
    When attachments are present the email is sent as multipart/mixed so
    both the HTML body and the PDF files are included.
    """
    if attachments:
        # outer: mixed (body + attachments)
        msg = MIMEMultipart("mixed")
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(html, "html"))
        msg.attach(alt)
        for fname, fbytes in attachments:
            part = MIMEBase("application", "pdf")
            part.set_payload(fbytes)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", "attachment",
                            filename=fname)
            msg.attach(part)
    else:
        msg = MIMEMultipart("alternative")
        msg.attach(MIMEText(html, "html"))

    msg["Subject"] = subject
    msg["From"]    = smtp_cfg["sender"]
    msg["To"]      = to
    msg["Cc"]      = ", ".join(cc_list)

    recipients = [to] + cc_list
    ctx = ssl.create_default_context()

    last_err = None
    for attempt in range(1, 4):          # up to 3 attempts
        try:
            if smtp_cfg.get("use_tls", True):
                with smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"], timeout=60) as s:
                    s.ehlo()
                    s.starttls(context=ctx)
                    s.ehlo()
                    s.login(smtp_cfg["user"], smtp_cfg["password"])
                    s.sendmail(smtp_cfg["sender"], recipients, msg.as_string())
            else:
                with smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"],
                                      context=ctx, timeout=60) as s:
                    s.login(smtp_cfg["user"], smtp_cfg["password"])
                    s.sendmail(smtp_cfg["sender"], recipients, msg.as_string())
            return "sent"                # success — exit immediately
        except (smtplib.SMTPServerDisconnected,
                smtplib.SMTPConnectError,
                TimeoutError,
                OSError) as e:
            last_err = e
            if attempt < 3:
                import time; time.sleep(3 * attempt)   # back-off: 3s, 6s
            continue
        except Exception as e:
            raise e                      # non-retryable (auth, bad address, etc.)

    raise last_err                       # all retries exhausted


# build_customer_email_html removed — replaced by build_email()


# ── Zoho Books helpers ────────────────────────────────────────────────────────
@st.cache_data(ttl=3000, show_spinner=False)   # cache token ~50 min; Zoho tokens live 60 min
def get_zoho_token(client_id: str, client_secret: str,
                   refresh_token: str, dc: str) -> str:
    """Exchange a Zoho refresh token for a fresh access token."""
    auth_host, _ = ZOHO_DC_MAP.get(dc, ZOHO_DC_MAP["US (.com)"])
    resp = requests.post(
        f"https://{auth_host}/oauth/v2/token",
        data={
            "grant_type":    "refresh_token",
            "client_id":     client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
        },
        timeout=20,
    )
    resp.raise_for_status()
    data = resp.json()
    if "access_token" not in data:
        raise ValueError(f"Zoho token error: {data.get('error', data)}")
    return data["access_token"]


# All known Zoho Books field names for the invoice / payment URL.
# invoice_url is the standard field; others are tried as fallbacks.
_PAYMENT_LINK_FIELDS = [
    "invoice_url",          # standard Zoho Books field — tried first
    "invoiceurl",
    "payment_link",
    "zohosecurepay_link",
    "online_payment_link",
    "secure_payment_url",
    "invoice_payment_link",
    "paymentlink",
]

def _extract_payment_link(obj: dict) -> str:
    """Try every known field name; return the first non-empty URL found."""
    for field in _PAYMENT_LINK_FIELDS:
        val = str(obj.get(field, "") or "").strip()
        if val.startswith("http"):
            return val
    # Last-resort: scan all string values for a zoho secure-pay URL pattern
    for val in obj.values():
        s = str(val or "")
        if "zoho" in s.lower() and "secure" in s.lower() and s.startswith("http"):
            return s
    return ""


def _zoho_invoice_id_and_link(invoice_number: str, access_token: str,
                               org_ids: list, api_host: str):
    """
    Internal helper. Searches across org_ids for invoice_number.
    Returns (invoice_id, org_id_found, payment_link) or raises on error.
    Returns (None, None, None) only when invoice is not found in any org.
    """
    base    = f"https://{api_host}/books/v3"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}

    for org_id in org_ids:
        org_id = str(org_id).strip()
        if not org_id:
            continue

        # Step 1 — search by invoice number
        sr = requests.get(
            f"{base}/invoices",
            params={"invoice_number": invoice_number, "organization_id": org_id},
            headers=headers,
            timeout=25,
        )
        sr.raise_for_status()
        sr_data  = sr.json()
        inv_list = sr_data.get("invoices", [])
        if not inv_list:
            continue                        # not in this org — try next

        invoice_id = inv_list[0]["invoice_id"]

        # Check if payment_link is already in the list response (saves one call)
        link_from_list = _extract_payment_link(inv_list[0])
        if link_from_list:
            return invoice_id, org_id, link_from_list

        # Step 2 — fetch full invoice detail for payment_link
        dr = requests.get(
            f"{base}/invoices/{invoice_id}",
            params={"organization_id": org_id},
            headers=headers,
            timeout=25,
        )
        dr.raise_for_status()
        inv_detail   = dr.json().get("invoice", {})
        payment_link = _extract_payment_link(inv_detail)

        return invoice_id, org_id, payment_link

    return None, None, None


def fetch_zoho_invoice_pdf(invoice_number: str, access_token: str,
                            org_ids: list, dc: str):
    """
    Fetch the PDF for *invoice_number* from Zoho Books.
    Also captures the payment_link (SecurePay URL) from the invoice detail.
    Returns (pdf_bytes, invoice_id, org_id_used, payment_link) on success.
    Returns (None, None, None, None) when the invoice isn't found in any org.
    Raises on network/auth errors.
    """
    _, api_host = ZOHO_DC_MAP.get(dc, ZOHO_DC_MAP["US (.com)"])
    base    = f"https://{api_host}/books/v3"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}

    invoice_id, org_id, payment_link = _zoho_invoice_id_and_link(
        invoice_number, access_token, org_ids, api_host
    )
    if not invoice_id:
        return None, None, None, None       # not found in any org

    # Download PDF
    pdf_resp = requests.get(
        f"{base}/invoices/{invoice_id}",
        params={"organization_id": org_id, "accept": "pdf"},
        headers=headers,
        timeout=40,
    )
    pdf_resp.raise_for_status()
    return pdf_resp.content, invoice_id, org_id, payment_link


def fetch_zoho_payment_links(invoice_numbers: list, access_token: str,
                              org_ids: list, dc: str) -> tuple:
    """
    Batch-fetch Zoho SecurePay payment links for a list of invoice numbers.
    Lightweight — no PDF download, just JSON API calls.

    Returns (links_dict, errors_dict):
      links_dict  = {invoice_number: payment_link}   — found links
      errors_dict = {invoice_number: error_message}  — fetch failures
    """
    _, api_host = ZOHO_DC_MAP.get(dc, ZOHO_DC_MAP["US (.com)"])
    links  = {}
    errors = {}
    for inv_num in invoice_numbers:
        inv_num = str(inv_num).strip()
        if not inv_num or inv_num in links:
            continue
        try:
            _, _, payment_link = _zoho_invoice_id_and_link(
                inv_num, access_token, org_ids, api_host
            )
            if payment_link:
                links[inv_num] = payment_link
            else:
                errors[inv_num] = "Found in Zoho but no payment link field present"
        except Exception as e:
            errors[inv_num] = str(e)
    return links, errors


# ── Zoho Books · Full invoice + line-item pull ────────────────────────────────

def _zoho_list_invoices_page(access_token: str, org_id: str, api_host: str,
                              page: int = 1, per_page: int = 200,
                              status: str = "") -> dict:
    """Fetch one page of invoices from Zoho Books."""
    params = {"organization_id": org_id, "page": page, "per_page": per_page,
              "sort_column": "date", "sort_order": "D"}
    if status:
        params["status"] = status
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    resp = requests.get(f"https://{api_host}/books/v3/invoices",
                        params=params, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json()


def fetch_zoho_all_invoices(access_token: str, org_ids: list, dc: str,
                             statuses: list | None = None) -> pd.DataFrame:
    """
    Pull ALL invoices across all org_ids from Zoho Books.
    statuses: list of Zoho status strings, e.g. ['overdue','sent','draft']
              Pass None/[] to fetch all statuses.
    Returns a DataFrame with canonical column names ready for load_data().
    """
    _, api_host = ZOHO_DC_MAP.get(dc, ZOHO_DC_MAP["US (.com)"])
    all_rows = []

    fetch_statuses = statuses if statuses else [None]   # None = no status filter

    for org_id in org_ids:
        for status in fetch_statuses:
            page = 1
            while True:
                data = _zoho_list_invoices_page(
                    access_token, org_id, api_host, page=page, status=status or "")
                invoices = data.get("invoices", [])
                for inv in invoices:
                    # Pull custom fields into a flat dict
                    cf = {f.get("label","").lower().strip(): f.get("value","")
                          for f in inv.get("custom_fields", [])}
                    all_rows.append({
                        "invoice_number":         inv.get("invoice_number",""),
                        "customer_name":          inv.get("customer_name",""),
                        "email":                  inv.get("email",""),
                        "date":                   inv.get("date",""),
                        "due_date":               inv.get("due_date",""),
                        "Current Invoice Status": inv.get("status","").title(),
                        "currency_code":          inv.get("currency_code",""),
                        "total":                  float(inv.get("total") or 0),
                        "balance":                float(inv.get("balance") or 0),
                        "Final USD":              float(inv.get("balance") or 0),
                        "Outstanding":            float(inv.get("balance") or 0),
                        "CSM":                    inv.get("salesperson_name",""),
                        "CSM Email":              inv.get("salesperson_email",""),
                        "payment_link":           inv.get("payment_link","") or
                                                  inv.get("invoice_url",""),
                        "_zoho_invoice_id":       inv.get("invoice_id",""),
                        "_zoho_org_id":           org_id,
                        # Merge useful custom fields
                        **{k: cf.get(k,"") for k in ("product","service type","country","billing terms")},
                    })
                page_ctx = data.get("page_context", {})
                if not page_ctx.get("has_more_page", False):
                    break
                page += 1

    if not all_rows:
        return pd.DataFrame()

    df_z = pd.DataFrame(all_rows).drop_duplicates(subset=["invoice_number","_zoho_org_id"])
    # Parse dates
    for dcol in ("date", "due_date"):
        df_z[dcol] = pd.to_datetime(df_z[dcol], errors="coerce")
    return df_z


def fetch_zoho_lineitems_for_invoices(invoice_ids: list, access_token: str,
                                      org_id: str, dc: str,
                                      progress_cb=None) -> pd.DataFrame:
    """
    Fetch line items (product-level) for a list of Zoho invoice IDs.
    Returns a DataFrame: invoice_number, product, description, qty, rate, amount, tax.
    progress_cb(i, total) is called each iteration if provided.
    """
    _, api_host = ZOHO_DC_MAP.get(dc, ZOHO_DC_MAP["US (.com)"])
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    rows = []
    for i, (inv_id, inv_number) in enumerate(invoice_ids):
        if progress_cb:
            progress_cb(i, len(invoice_ids))
        try:
            resp = requests.get(
                f"https://{api_host}/books/v3/invoices/{inv_id}",
                params={"organization_id": org_id},
                headers=headers, timeout=20,
            )
            resp.raise_for_status()
            inv_detail = resp.json().get("invoice", {})
            for li in inv_detail.get("line_items", []):
                rows.append({
                    "invoice_number": inv_number,
                    "product":        li.get("name",""),
                    "description":    li.get("description",""),
                    "quantity":       float(li.get("quantity") or 0),
                    "unit":           li.get("unit",""),
                    "rate":           float(li.get("rate") or 0),
                    "amount":         float(li.get("item_total") or 0),
                    "tax_name":       li.get("tax_name",""),
                    "tax_pct":        float(li.get("tax_percentage") or 0),
                })
        except Exception:
            pass
    return pd.DataFrame(rows) if rows else pd.DataFrame()


def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS reasons (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                level           TEXT NOT NULL,
                identifier      TEXT NOT NULL,
                reason_category TEXT,
                reason_text     TEXT,
                action_owner    TEXT,
                next_action_date TEXT,
                updated_at      TEXT,
                UNIQUE(level, identifier)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sent_emails (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                invoice_no   TEXT,
                customer     TEXT,
                to_email     TEXT,
                cc_emails    TEXT,
                subject      TEXT,
                sent_at      TEXT,
                status       TEXT,
                error        TEXT,
                template     TEXT
            )
        """)
        # Migrate: add template column if it doesn't exist yet
        try:
            conn.execute("ALTER TABLE sent_emails ADD COLUMN template TEXT")
        except Exception:
            pass  # column already exists
        # ── User management table ─────────────────────────────────────────────
        conn.execute("""
            CREATE TABLE IF NOT EXISTS app_users (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                username      TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                role          TEXT NOT NULL DEFAULT 'viewer',
                csm_name      TEXT,
                created_by    TEXT,
                created_at    TEXT
            )
        """)
        # Migrate: add csm_name column if upgrading from older schema
        try:
            conn.execute("ALTER TABLE app_users ADD COLUMN csm_name TEXT")
        except Exception:
            pass


def log_email(invoice_nos, customer, to_email, cc_emails, subject, status,
              error="", template=""):
    """Log one row per invoice_no so reminder counts work correctly.
    invoice_nos can be a single string or a list of strings.
    """
    from datetime import timezone, timedelta
    _IST = timezone(timedelta(hours=5, minutes=30))
    sent_at = datetime.now(_IST).isoformat()

    if isinstance(invoice_nos, str):
        invoice_nos = [invoice_nos]
    with sqlite3.connect(DB_PATH) as conn:
        for inv_no in invoice_nos:
            conn.execute("""
                INSERT INTO sent_emails
                    (invoice_no, customer, to_email, cc_emails, subject, sent_at, status, error, template)
                VALUES (?,?,?,?,?,?,?,?,?)
            """, (str(inv_no).strip(), customer, to_email, cc_emails, subject,
                  sent_at, status, error, template))


def get_sent_log():
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql(
            "SELECT * FROM sent_emails ORDER BY sent_at DESC", conn
        )


def get_reminder_counts() -> dict:
    """Return {invoice_no: sent_count} for all successfully sent emails."""
    with sqlite3.connect(DB_PATH) as conn:
        df = pd.read_sql(
            "SELECT invoice_no, COUNT(*) as cnt FROM sent_emails "
            "WHERE status='sent' GROUP BY invoice_no",
            conn,
        )
    if df.empty:
        return {}
    return dict(zip(df["invoice_no"], df["cnt"]))


# ── User management DB helpers ────────────────────────────────────────────────
def _db_get_all_users() -> pd.DataFrame:
    """Return all users stored in app_users SQLite table."""
    with sqlite3.connect(DB_PATH) as conn:
        try:
            return pd.read_sql(
                "SELECT username, role, csm_name, created_by, created_at "
                "FROM app_users ORDER BY created_at",
                conn,
            )
        except Exception:
            return pd.DataFrame(columns=["username", "role", "csm_name", "created_by", "created_at"])


def _db_create_user(username: str, password: str, role: str, created_by: str,
                    csm_name: str = "") -> tuple[bool, str]:
    """Insert a new user. Returns (success, message)."""
    uname = username.strip().lower()
    if not uname:
        return False, "Username cannot be empty."
    if len(password) < 4:
        return False, "Password must be at least 4 characters."
    if role not in ROLES:
        return False, f"Invalid role: {role}"
    pw_hash = _hash_pw(password)
    try:
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute(
                "INSERT INTO app_users (username, password_hash, role, csm_name, created_by, created_at) "
                "VALUES (?,?,?,?,?,?)",
                (uname, pw_hash, role, csm_name.strip() or None,
                 created_by, datetime.now().isoformat()),
            )
        _sync_users_to_credentials()
        return True, f"User **{uname}** created successfully."
    except sqlite3.IntegrityError:
        return False, f"Username **{uname}** already exists."
    except Exception as e:
        return False, f"Error: {e}"


def _db_update_user(username: str, new_password: str | None, new_role: str | None,
                    new_csm_name: str | None = None) -> tuple[bool, str]:
    """Update password, role, and/or CSM name for an existing DB user."""
    uname = username.strip().lower()
    try:
        with sqlite3.connect(DB_PATH) as conn:
            if new_password:
                conn.execute("UPDATE app_users SET password_hash=? WHERE username=?",
                             (_hash_pw(new_password), uname))
            if new_role and new_role in ROLES:
                conn.execute("UPDATE app_users SET role=? WHERE username=?",
                             (new_role, uname))
            if new_csm_name is not None:
                conn.execute("UPDATE app_users SET csm_name=? WHERE username=?",
                             (new_csm_name.strip() or None, uname))
        _sync_users_to_credentials()
        return True, f"User **{uname}** updated."
    except Exception as e:
        return False, f"Error: {e}"


def _db_delete_user(username: str) -> tuple[bool, str]:
    """Delete a user from app_users table."""
    uname = username.strip().lower()
    try:
        with sqlite3.connect(DB_PATH) as conn:
            conn.execute("DELETE FROM app_users WHERE username=?", (uname,))
        _sync_users_to_credentials()
        return True, f"User **{uname}** deleted."
    except Exception as e:
        return False, f"Error: {e}"


def _sync_users_to_credentials():
    """Write all DB users back to credentials.json and secrets.toml (local files only)."""
    import re as _re
    # Gather DB data
    db_pw_map: dict[str, str] = {}
    db_roles:  dict[str, str] = {}
    db_csm:    dict[str, str] = {}
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute(
                "SELECT username, password_hash, role, csm_name FROM app_users"
            ).fetchall()
        for uname, pw_hash, role, csm_name in rows:
            db_pw_map[uname] = pw_hash
            db_roles[uname]  = role
            if csm_name:
                db_csm[uname] = csm_name
    except Exception:
        pass

    # ── credentials.json ──────────────────────────────────────────────────────
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH, "r") as f:
                creds = json.load(f)
        except Exception:
            creds = {}
        creds.setdefault("users", {}).update(db_pw_map)
        creds.setdefault("roles", {}).update(db_roles)
        creds.setdefault("csm_assignments", {}).update(db_csm)
        try:
            with open(CREDS_PATH, "w") as f:
                json.dump(creds, f, indent=2)
        except Exception:
            pass

    # ── .streamlit/secrets.toml ───────────────────────────────────────────────
    _secrets_path = os.path.join(os.path.dirname(__file__), ".streamlit", "secrets.toml")
    if not os.path.exists(_secrets_path):
        return
    try:
        with open(_secrets_path, "r") as f:
            toml_text = f.read()
    except Exception:
        toml_text = ""

    def _parse_toml_section(text, section):
        blk = _re.search(rf"\[{section}\](.*?)(?=\n\[|\Z)", text, _re.S)
        out = {}
        if blk:
            for m in _re.finditer(r'(\w[\w.]*)\s*=\s*"([^"]*)"', blk.group(1)):
                out[m.group(1).lower()] = m.group(2)
        return out

    def _build_toml_section(section_name: str, mapping: dict, comment: str = "") -> str:
        lines = []
        if comment:
            lines.append(comment)
        lines.append(f"[{section_name}]")
        for k, v in sorted(mapping.items()):
            lines.append(f'{k:<14}= "{v}"')
        return "\n".join(lines)

    all_users = _parse_toml_section(toml_text, "users")
    all_roles = _parse_toml_section(toml_text, "roles")
    all_csm   = _parse_toml_section(toml_text, "csm_assignments")
    all_users.update(db_pw_map)
    all_roles.update(db_roles)
    all_csm.update(db_csm)

    # Strip old managed sections and rebuild
    clean = toml_text
    for sec in ("users", "roles", "csm_assignments"):
        clean = _re.sub(rf"(?:# ──[^\n]*\n)?\[{sec}\].*?(?=\n\[|\Z)", "", clean, flags=_re.S)
    clean = clean.strip()

    blocks = [
        _build_toml_section("users",           all_users, "# ── App Users ────────────────────────────────────"),
        _build_toml_section("roles",           all_roles, "# ── User Roles ───────────────────────────────────"),
    ]
    if all_csm:
        blocks.append(_build_toml_section("csm_assignments", all_csm,
                                           "# ── CSM Assignments (username = CSM display name) ──"))
    try:
        with open(_secrets_path, "w") as f:
            f.write(clean + "\n\n" + "\n\n".join(blocks) + "\n")
    except Exception:
        pass


def _get_secrets_toml_snippet() -> str:
    """Generate a [users]+[roles]+[csm_assignments] TOML snippet for Streamlit Cloud."""
    all_users: dict[str, str] = {}
    all_roles: dict[str, str] = {}
    all_csm:   dict[str, str] = {}
    # From credentials.json
    if os.path.exists(CREDS_PATH):
        try:
            with open(CREDS_PATH, "r") as f:
                data = json.load(f)
            all_users.update(data.get("users", {}))
            all_roles.update(data.get("roles", {}))
            all_csm.update(data.get("csm_assignments", {}))
        except Exception:
            pass
    # Overlay DB
    try:
        with sqlite3.connect(DB_PATH) as conn:
            rows = conn.execute(
                "SELECT username, password_hash, role, csm_name FROM app_users"
            ).fetchall()
        for uname, pw_hash, role, csm_name in rows:
            all_users[uname] = pw_hash
            all_roles[uname] = role
            if csm_name:
                all_csm[uname] = csm_name
    except Exception:
        pass

    lines = ["[users]"]
    for u, pw in sorted(all_users.items()):
        lines.append(f'{u:<14}= "{pw}"')
    lines += ["", "[roles]"]
    for u, role in sorted(all_roles.items()):
        lines.append(f'{u:<14}= "{role}"')
    if all_csm:
        lines += ["", "# CSM display name → maps username to the CSM column in your sheet",
                  "[csm_assignments]"]
        for u, cname in sorted(all_csm.items()):
            lines.append(f'{u:<14}= "{cname}"')
    return "\n".join(lines)


def upsert_reason(level, identifier, category, text, owner, next_dt):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT INTO reasons
                (level, identifier, reason_category, reason_text, action_owner, next_action_date, updated_at)
            VALUES (?,?,?,?,?,?,?)
            ON CONFLICT(level, identifier) DO UPDATE SET
                reason_category  = excluded.reason_category,
                reason_text      = excluded.reason_text,
                action_owner     = excluded.action_owner,
                next_action_date = excluded.next_action_date,
                updated_at       = excluded.updated_at
        """, (level, identifier, category, text, owner, str(next_dt), datetime.now().isoformat()))


def delete_reason(level, identifier):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("DELETE FROM reasons WHERE level=? AND identifier=?", (level, identifier))


def get_reasons(level=None):
    with sqlite3.connect(DB_PATH) as conn:
        if level:
            return pd.read_sql("SELECT * FROM reasons WHERE level=? ORDER BY updated_at DESC", conn, params=(level,))
        return pd.read_sql("SELECT * FROM reasons ORDER BY updated_at DESC", conn)


# ── Data loading ──────────────────────────────────────────────────────────────
# Canonical column name → list of accepted aliases (all lowercase, stripped)
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
    "Customer CC Email":         ["customer cc email", "customer_cc_email", "cc email",
                                  "customer cc", "cc_email", "client cc email",
                                  "customer cc emails", "customer_cc_emails",
                                  "cc emails", "cc_emails"],
    "payment_link":              ["zohosecurepay", "zoho secure pay", "zoho securepay",
                                  "payment link", "payment_link", "pay link", "pay_link",
                                  "secure pay link", "secure payment link", "payment url"],
}

def remap_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename df columns to canonical names using case-insensitive alias matching."""
    lookup = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            lookup[alias] = canonical

    rename_map = {}
    for col in df.columns:
        key = col.lower().strip()
        if key in lookup and col != lookup[key]:
            rename_map[col] = lookup[key]

    return df.rename(columns=rename_map)


@st.cache_data(show_spinner="Loading data…")
def load_data(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes))

    # Normalise column names: strip whitespace, remove non-breaking spaces
    df.columns = [str(c).strip().replace("\xa0", " ") for c in df.columns]

    # De-duplicate column names (customer_status appears twice in spec)
    cols = []
    seen = {}
    for c in df.columns:
        if c in seen:
            seen[c] += 1
            cols.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            cols.append(c)
    df.columns = cols

    # Remap aliases → canonical names
    df = remap_columns(df)

    # Show actual column names in sidebar for debugging
    st.session_state["_raw_cols"] = list(df.columns)

    date_cols = ["due_date", "date", "last_payment_date",
                 "Service_period_Start_date", "Service_period_End_date"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    num_cols = ["Final USD", "total", "balance", "Outstanding"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── Aging: Today - Invoice Date if invoice date > service start, else Today - Service Start ──
    today = pd.Timestamp.today().normalize()
    if "date" in df.columns and "Service_period_Start_date" in df.columns:
        inv_after_start = df["date"] > df["Service_period_Start_date"]
        days_from_invoice = (today - df["date"]).dt.days
        days_from_start   = (today - df["Service_period_Start_date"]).dt.days
        df["Aging"] = np.where(inv_after_start, days_from_invoice, days_from_start).clip(min=0)
    elif "date" in df.columns:
        df["Aging"] = (today - df["date"]).dt.days.clip(lower=0)
    else:
        df["Aging"] = 0

    # ── Bucket ────────────────────────────────────────────────────────────────
    def aging_bucket(a):
        if a <= 15:  return "0-15"
        if a <= 30:  return "16-30"
        if a <= 45:  return "31-45"
        if a <= 60:  return "46-60"
        if a <= 90:  return "61-90"
        return "90+"

    df["Bucket"] = df["Aging"].apply(aging_bucket)

    # ── RAG: customer-level, then broadcast to invoice rows ───────────────────
    # Red   → customer has ANY invoice in 90+ bucket
    # Amber → customer has ANY invoice with aging > 30 (and not Red)
    # Green → all invoices ≤ 30 days
    OVER_90 = {"90+"}
    OVER_30 = {"31-45", "46-60", "61-90", "90+"}

    def customer_rag(buckets_set):
        if buckets_set & OVER_90: return "Red"
        if buckets_set & OVER_30: return "Amber"
        return "Green"

    # Auto-detect the customer name column (case-insensitive, strip)
    cust_col = next(
        (c for c in df.columns if c.lower().replace(" ", "_") == "customer_name"),
        None
    )
    if cust_col is None:
        # Fallback: pick first column whose name contains "customer"
        cust_col = next((c for c in df.columns if "customer" in c.lower()), None)

    if cust_col:
        if cust_col != "customer_name":
            df = df.rename(columns={cust_col: "customer_name"})
        customer_buckets = df.groupby("customer_name")["Bucket"].apply(set)
        rag_map = customer_buckets.apply(customer_rag)
        df["RAG"] = df["customer_name"].map(rag_map)
    else:
        df["RAG"] = "Green"

    return df


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_usd(val):
    if abs(val) >= 1_000_000:
        return f"${val/1_000_000:.2f}M"
    if abs(val) >= 1_000:
        return f"${val/1_000:.1f}K"
    return f"${val:,.0f}"

def fmt_inr(val):
    if abs(val) >= 1_000_0000:        # 1 Crore
        return f"₹{val/1_000_0000:.2f}Cr"
    if abs(val) >= 1_00_000:          # 1 Lakh
        return f"₹{val/1_00_000:.2f}L"
    if abs(val) >= 1_000:
        return f"₹{val/1_000:.1f}K"
    return f"₹{val:,.0f}"

def _indian_commas(n: float) -> str:
    """Format an integer using Indian comma placement: last 3 digits then groups of 2.
    e.g. 2154240 → '21,54,240'  |  21542400 → '2,15,42,400'
    """
    neg = n < 0
    s = str(int(abs(round(n))))
    if len(s) <= 3:
        result = s
    else:
        result = s[-3:]
        s = s[:-3]
        while len(s) > 2:
            result = s[-2:] + "," + result
            s = s[:-2]
        result = s + "," + result
    return ("-" if neg else "") + result

def fmt_amount(val: float, currency_code: str) -> str:
    """Format a monetary amount with currency symbol and locale-appropriate separators.
    INR  → Indian lakh/crore comma style  (₹21,54,240)
    Others → standard million comma style ($1,234,567)
    """
    sym = CURR_SYM.get(str(currency_code).upper(), "")
    if str(currency_code).upper() == "INR":
        return f"{sym}{_indian_commas(val)}"
    return f"{sym}{val:,.0f}"


def rag_badge(val):
    icons = {"Red": "🔴", "Amber": "🟡", "Green": "🟢"}
    return f"{icons.get(str(val), '⚪')} {val}"


def export_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def _fmt_fig(fig):
    """Apply rounded integer formatting to all hover labels and axis ticks."""
    fig.update_traces(
        hovertemplate=None,   # reset per-trace template first
    )
    # Apply a clean rounded hover across all traces
    for trace in fig.data:
        if hasattr(trace, "hovertemplate"):
            if trace.type == "pie":
                trace.hovertemplate = "<b>%{label}</b><br>%{value:,.0f}<br>%{percent}<extra></extra>"
            else:
                trace.hovertemplate = (
                    "<b>%{fullData.name}</b><br>"
                    "%{x}: <b>%{y:,.0f}</b><extra></extra>"
                    if trace.orientation != "h"
                    else "<b>%{fullData.name}</b><br>"
                         "%{y}: <b>%{x:,.0f}</b><extra></extra>"
                )
    fig.update_layout(
        yaxis=dict(tickformat=",.0f"),
        xaxis=dict(tickformat=",.0f"),
    )
    return fig


# ── Column-level filter widget (reusable) ────────────────────────────────────
def column_filters(df: pd.DataFrame, key_prefix: str = "cf") -> pd.DataFrame:
    """
    Renders one filter widget per column (text / multiselect / range) in a
    compact expander above the table and returns the filtered DataFrame.
    Columns with ≤30 unique values → multiselect
    Numeric columns              → min/max number inputs
    Other columns                → case-insensitive text search
    """
    with st.expander("🔍 Column Filters", expanded=False):
        filtered = df.copy()
        n_cols = len(df.columns)
        # layout: up to 4 widgets per row
        cols_per_row = 4
        col_chunks = [list(df.columns)[i:i+cols_per_row]
                      for i in range(0, n_cols, cols_per_row)]

        for chunk in col_chunks:
            row = st.columns(len(chunk))
            for widget_col, col_name in zip(row, chunk):
                series = df[col_name].dropna()
                unique_vals = series.unique()
                widget_key = f"{key_prefix}_{col_name}"

                with widget_col:
                    if pd.api.types.is_numeric_dtype(df[col_name]):
                        col_min = float(series.min()) if len(series) else 0.0
                        col_max = float(series.max()) if len(series) else 0.0
                        if col_min == col_max:
                            continue  # nothing to filter
                        lo = st.number_input(
                            f"{col_name} ≥", value=col_min,
                            min_value=col_min, max_value=col_max,
                            step=max((col_max - col_min) / 100, 0.01),
                            key=f"{widget_key}_lo", label_visibility="visible",
                        )
                        hi = st.number_input(
                            f"{col_name} ≤", value=col_max,
                            min_value=col_min, max_value=col_max,
                            step=max((col_max - col_min) / 100, 0.01),
                            key=f"{widget_key}_hi", label_visibility="visible",
                        )
                        if lo > col_min or hi < col_max:
                            filtered = filtered[
                                (filtered[col_name] >= lo) & (filtered[col_name] <= hi)
                            ]
                    elif len(unique_vals) <= 30:
                        choices = st.multiselect(
                            col_name,
                            options=sorted([str(v) for v in unique_vals]),
                            default=[],
                            key=widget_key,
                        )
                        if choices:
                            filtered = filtered[
                                filtered[col_name].astype(str).isin(choices)
                            ]
                    else:
                        text = st.text_input(
                            col_name, value="", key=widget_key,
                            placeholder="search…",
                        )
                        if text.strip():
                            filtered = filtered[
                                filtered[col_name].astype(str)
                                    .str.contains(text.strip(), case=False, na=False)
                            ]
        return filtered


# ── Reason form (reusable) ────────────────────────────────────────────────────
def reason_form(level: str, identifiers, label: str, df_ref: pd.DataFrame = None):
    existing_all = get_reasons(level)

    left, right = st.columns([2, 3])
    with left:
        selected = st.selectbox(f"Select {label}", sorted([str(x) for x in identifiers if pd.notna(x)]), key=f"sel_{level}")

    existing_row = (
        existing_all[existing_all["identifier"] == selected]
        if not existing_all.empty else pd.DataFrame()
    )
    existing = existing_row.iloc[0].to_dict() if not existing_row.empty else {}

    with right:
        if existing:
            st.success(f"✅ Reason on record — last updated {str(existing.get('updated_at',''))[:10]}")

    with st.form(f"form_{level}_{selected}", clear_on_submit=False):
        default_cat = existing.get("reason_category", REASON_CATEGORIES[0])
        cat_idx = REASON_CATEGORIES.index(default_cat) if default_cat in REASON_CATEGORIES else 0

        cat = st.selectbox("Reason Category", REASON_CATEGORIES, index=cat_idx)
        notes = st.text_area("Notes / Details", value=existing.get("reason_text", ""), height=100)

        c1, c2 = st.columns(2)
        with c1:
            owner = st.text_input("Action Owner", value=existing.get("action_owner", ""))
        with c2:
            raw_date = existing.get("next_action_date")
            try:
                default_date = date.fromisoformat(str(raw_date)[:10]) if raw_date else date.today()
            except Exception:
                default_date = date.today()
            next_dt = st.date_input("Next Action Date", value=default_date)

        _can_edit_reasons = _can("edit_reasons")
        col_save, col_del = st.columns([3, 1])
        with col_save:
            if st.form_submit_button("💾  Save", use_container_width=True, type="primary",
                                     disabled=not _can_edit_reasons):
                upsert_reason(level, selected, cat, notes, owner, next_dt)
                st.success("Saved!")
                st.rerun()
        with col_del:
            if existing and st.form_submit_button("🗑 Delete", use_container_width=True,
                                                  disabled=not _can_edit_reasons):
                delete_reason(level, selected)
                st.warning("Deleted.")
                st.rerun()
        if not _can_edit_reasons:
            st.caption("🔒 Your role is view-only. Contact an Admin to make changes.")

    st.divider()

    if not existing_all.empty:
        st.subheader(f"All saved {label}-level reasons")
        display = existing_all[["identifier", "reason_category", "reason_text",
                                 "action_owner", "next_action_date", "updated_at"]].copy()
        display.columns = [label, "Category", "Notes", "Owner", "Next Action", "Last Updated"]
        display["Last Updated"] = display["Last Updated"].str[:10]
        display["Next Action"] = display["Next Action"].str[:10]

        # ── For Customer level: join Invoice Value & Outstanding from live data ──
        if level == "customer" and df_ref is not None and not df_ref.empty:
            _cn = "customer_name"
            if _cn in df_ref.columns:

                def _cust_fc_string(cname, amount_col):
                    """Return currency-aware formatted string (e.g. ₹12,34,567 | $1,234)."""
                    rows = df_ref[df_ref[_cn] == cname]
                    if amount_col not in rows.columns:
                        return "—"
                    has_curr = "currency_code" in rows.columns
                    if has_curr:
                        tmp = rows.copy()
                        tmp["_curr"] = tmp["currency_code"].astype(str).str.strip().replace(
                            {"": "INR", "nan": "INR", "NaN": "INR", "None": "INR"})
                        parts = (tmp.groupby("_curr")[amount_col].sum()
                                    .reset_index()
                                    .sort_values(amount_col, ascending=False))
                        return "  |  ".join(
                            fmt_amount(r[amount_col], r["_curr"]) for _, r in parts.iterrows()
                        ) or "—"
                    return fmt_amount(rows[amount_col].sum(), "USD")

                # Invoice Value = sum of 'total' (gross invoice amount in native currency)
                _iv_col = "total"    if "total"    in df_ref.columns else \
                          "Final USD" if "Final USD" in df_ref.columns else None
                # Outstanding  = sum of 'balance' (remaining unpaid in native currency)
                _ob_col = "balance"  if "balance"  in df_ref.columns else \
                          "Final USD" if "Final USD" in df_ref.columns else None

                _cust_names = display[label].tolist()
                if _iv_col:
                    display["Invoice Value"] = display[label].apply(
                        lambda c: _cust_fc_string(c, _iv_col))
                if _ob_col:
                    display["Total Outstanding"] = display[label].apply(
                        lambda c: _cust_fc_string(c, _ob_col))

        st.dataframe(display, use_container_width=True, height=300)

        dl_bytes = export_excel(display)
        st.download_button(
            f"⬇ Download {label} Reasons",
            data=dl_bytes,
            file_name=f"{level}_reasons.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info(f"No {label}-level reasons saved yet.")


# ── Google Sheets helpers ─────────────────────────────────────────────────────
def parse_gsheet_url(url: str):
    """Return (spreadsheet_id, gid) from any Google Sheets URL, or raise ValueError."""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        raise ValueError("Could not find a spreadsheet ID in the URL.")
    sheet_id = match.group(1)
    gid_match = re.search(r"[#&?]gid=(\d+)", url)
    gid = gid_match.group(1) if gid_match else "0"
    return sheet_id, gid


@st.cache_data(ttl=120, show_spinner=False)
def fetch_gsheet(url: str) -> bytes:
    """Download a Google Sheet as xlsx bytes (sheet must be publicly shared). Cached 2 min."""
    sheet_id, gid = parse_gsheet_url(url)
    export_url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}"
        f"/export?format=xlsx&gid={gid}"
    )
    resp = requests.get(export_url, timeout=30)
    if resp.status_code == 401:
        raise PermissionError(
            "Sheet is private. Share it as 'Anyone with the link can view' and try again."
        )
    if resp.status_code != 200:
        raise ConnectionError(f"Google returned HTTP {resp.status_code}. Check the URL.")
    return resp.content


@st.cache_data(ttl=120, show_spinner=False)
def fetch_gsheet_private(url: str, creds_json: str) -> bytes:
    """Download a private Google Sheet using a service-account JSON string. Cached 2 min."""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        import json
    except ImportError:
        raise ImportError("Run:  pip install gspread google-auth")

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    info = json.loads(creds_json)
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc = gspread.authorize(creds)
    sheet_id, gid = parse_gsheet_url(url)
    sh = gc.open_by_key(sheet_id)
    worksheet = next((ws for ws in sh.worksheets() if str(ws.id) == gid), sh.sheet1)
    records = worksheet.get_all_records()
    buf = BytesIO()
    pd.DataFrame(records).to_excel(buf, index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────
init_db()

# ── Login gate ────────────────────────────────────────────────────────────────
if not _login_page():
    st.stop()

load_credentials()   # populate session_state from credentials.json (first run only)

# ── Logout in sidebar ─────────────────────────────────────────────────────────
with st.sidebar:
    _sb_role  = st.session_state.get("_role", "viewer")
    _sb_label = ROLE_LABELS.get(_sb_role, _sb_role.title())
    _sb_color = ROLE_COLORS.get(_sb_role, "#64748b")
    st.markdown(
        f"👤 **{st.session_state.get('_username', 'user').title()}**&nbsp;&nbsp;"
        f'<span style="background:{_sb_color}33;color:{_sb_color};'
        f'border:1px solid {_sb_color}66;border-radius:12px;'
        f'padding:1px 8px;font-size:11px;font-weight:700;">{_sb_label}</span>',
        unsafe_allow_html=True,
    )
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["_authenticated"] = False
        st.session_state["_username"]      = ""
        st.rerun()
    st.divider()

    # ── Admin: User Management panel ──────────────────────────────────────────
    if _can("manage_users"):
        with st.expander("👥 User Management", expanded=False):
            st.caption("Create, update, or delete users. Changes are saved immediately.")

            # ── Role reference card ──────────────────────────────────────────
            _role_ref_cols = st.columns(5)
            _role_desc = {
                "admin":      "Full access",
                "executor":   "Operational access",
                "viewer":     "Read-only (no Overview)",
                "csm":        "Filtered to own data",
                "management": "Summary views only",
            }
            for _rc, _rname in zip(_role_ref_cols, ROLES):
                _rc.markdown(
                    f"<div style='background:{ROLE_COLORS[_rname]}15;border:1px solid "
                    f"{ROLE_COLORS[_rname]}44;border-radius:8px;padding:6px 8px;text-align:center;'>"
                    f"<div style='font-size:11px;font-weight:700;color:{ROLE_COLORS[_rname]};'>"
                    f"{ROLE_LABELS[_rname]}</div>"
                    f"<div style='font-size:10px;color:#6b7280;margin-top:2px;'>{_role_desc[_rname]}</div>"
                    f"</div>", unsafe_allow_html=True)
            st.markdown("")

            # ── Create new user ──────────────────────────────────────────────
            st.markdown("#### ➕ Create User")
            with st.form("create_user_form", clear_on_submit=True):
                _nu_c1, _nu_c2 = st.columns(2)
                with _nu_c1:
                    _new_uname = st.text_input("Username", placeholder="e.g. ravi")
                    _new_pw    = st.text_input("Password", type="password",
                                               placeholder="min 4 chars")
                with _nu_c2:
                    _new_role  = st.selectbox("Role", ROLES,
                                              format_func=lambda r: ROLE_LABELS.get(r, r.title()),
                                              index=ROLES.index("viewer"))
                    _new_csm_name = st.text_input(
                        "CSM Name (only for CSM role)",
                        placeholder="Exact name as in sheet's CSM column",
                        help="Required when role = CSM. Must match the CSM column value exactly."
                    )
                _create_btn = st.form_submit_button("✅ Create User",
                                                    use_container_width=True, type="primary")
            if _create_btn:
                _ok, _msg = _db_create_user(
                    _new_uname, _new_pw, _new_role,
                    created_by=st.session_state.get("_username", "admin"),
                    csm_name=_new_csm_name,
                )
                if _ok:
                    st.success(_msg)
                    if _new_role == "csm" and not _new_csm_name.strip():
                        st.warning("⚠️ You created a CSM user without a CSM Name. "
                                   "Edit the user below to assign their CSM name, "
                                   "or data will fall back to their username.")
                else:
                    st.error(_msg)
                st.rerun()

            st.divider()

            # ── Current users table ──────────────────────────────────────────
            st.markdown("#### 👤 All Users")
            _all_db = _db_get_all_users()

            # Merge static + DB for display
            _static_roles = {}
            if os.path.exists(CREDS_PATH):
                try:
                    with open(CREDS_PATH, "r") as _f:
                        _static_roles = json.load(_f).get("roles", {})
                except Exception:
                    pass
            try:
                if "roles" in st.secrets:
                    _static_roles.update(dict(st.secrets["roles"]))
            except Exception:
                pass
            _db_unames = set(_all_db["username"].tolist()) if not _all_db.empty else set()

            def _role_tag(role: str) -> str:
                clr = ROLE_COLORS.get(role, "#64748b")
                lbl = ROLE_LABELS.get(role, role.title())
                return (f'<span style="background:{clr}22;color:{clr};border:1px solid {clr}55;'
                        f'border-radius:10px;padding:1px 8px;font-size:11px;font-weight:700;">{lbl}</span>')

            # Static config users (non-editable)
            _any_static = False
            for _u, _r in sorted({**_DEFAULT_ROLES, **_static_roles}.items()):
                if _u in _db_unames:
                    continue
                _any_static = True
                _uc1, _uc2, _uc3 = st.columns([3, 3, 2])
                with _uc1:
                    st.markdown(f"🔒 **{_u}**  \n<small style='color:#6b7280;'>static config</small>",
                                unsafe_allow_html=True)
                with _uc2:
                    st.markdown(_role_tag(_r), unsafe_allow_html=True)
                with _uc3:
                    st.caption("(config only)")

            # DB users (editable)
            if not _all_db.empty:
                for _, _row in _all_db.iterrows():
                    _uc1, _uc2, _uc3 = st.columns([3, 3, 2])
                    with _uc1:
                        _csm_hint = ""
                        if _row["role"] == "csm" and _row.get("csm_name"):
                            _csm_hint = f" · CSM: *{_row['csm_name']}*"
                        st.markdown(
                            f"✏️ **{_row['username']}**{_csm_hint}  \n"
                            f"<small style='color:#6b7280;'>by {_row.get('created_by','—')} "
                            f"· {str(_row.get('created_at',''))[:16]}</small>",
                            unsafe_allow_html=True)
                    with _uc2:
                        st.markdown(_role_tag(_row["role"]), unsafe_allow_html=True)
                    with _uc3:
                        if st.button("🗑 Delete", key=f"del_{_row['username']}",
                                     use_container_width=True):
                            _ok2, _msg2 = _db_delete_user(_row["username"])
                            st.toast(_msg2, icon="✅" if _ok2 else "❌")
                            st.rerun()
                    st.markdown("---")

            # ── Edit DB user ─────────────────────────────────────────────────
            if not _all_db.empty:
                st.markdown("#### ✏️ Edit DB User")
                with st.form("edit_user_form", clear_on_submit=True):
                    _edit_uname = st.selectbox("User to edit",
                                               options=sorted(_all_db["username"].tolist()))
                    _eu_c1, _eu_c2 = st.columns(2)
                    with _eu_c1:
                        _new_pw2      = st.text_input("New Password (leave blank to keep)",
                                                       type="password")
                        _new_csm_nm2  = st.text_input(
                            "CSM Name (leave blank to keep)",
                            placeholder="Exact name from CSM column in sheet",
                        )
                    with _eu_c2:
                        _new_role2 = st.selectbox(
                            "New Role", ROLES,
                            format_func=lambda r: ROLE_LABELS.get(r, r.title()),
                        )
                    _edit_btn = st.form_submit_button("💾 Save Changes", use_container_width=True)
                if _edit_btn:
                    _ok3, _msg3 = _db_update_user(
                        _edit_uname,
                        _new_pw2 or None,
                        _new_role2,
                        new_csm_name=_new_csm_nm2 if _new_csm_nm2.strip() else None,
                    )
                    if _ok3:
                        st.success(_msg3)
                    else:
                        st.error(_msg3)
                    st.rerun()

            st.divider()

            # ── Streamlit Cloud secrets snippet ──────────────────────────────
            st.markdown("#### ☁️ Streamlit Cloud Secrets")
            st.caption(
                "Paste this into **App Settings → Secrets** on Streamlit Cloud "
                "so users persist across redeploys."
            )
            _snippet = _get_secrets_toml_snippet()
            st.code(_snippet, language="toml")

    st.divider()

# ── Fixed Google Sheet ────────────────────────────────────────────────────────
_FIXED_SHEET_URL = "https://docs.google.com/spreadsheets/d/1pY_hPKVa8A-d6kbCnsuRdns4CiuRTh1QaIJRf5-ppOI/edit?gid=0#gid=0"

file_bytes = None

# ── Top banner: logo | title | refresh ────────────────────────────────────────
# Use st.columns for layout (avoids flex-clipping in st.markdown)
_uname_display  = st.session_state.get("_username", "user").title()
_role_now       = st.session_state.get("_role", "viewer")
_role_label     = ROLE_LABELS.get(_role_now, _role_now.title())
_role_color     = ROLE_COLORS.get(_role_now, "#64748b")

# Shared banner background injected via CSS on a known class
st.markdown("""
<style>
[data-testid="stHorizontalBlock"]:has(> div > [data-testid="stColumn"] > div > .banner-logo) {
    background: linear-gradient(135deg,#0f172a 0%,#0d2b52 55%,#1e3a8a 100%);
    border-radius: 14px;
    border: 1px solid rgba(96,165,250,0.15);
    box-shadow: 0 4px 24px rgba(0,0,0,0.35);
    padding: 4px 0;
    margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)

_col_logo, _col_title, _col_refresh = st.columns([2, 6, 2])

with _col_logo:
    # Logo source: local base64 PNG → Clearbit CDN → nothing (CSS class still sets bg)
    if _LOGO_B64:
        _logo_src = f"data:image/png;base64,{_LOGO_B64}"
    else:
        _logo_src = "https://logo.clearbit.com/spyne.ai"

    # White pill container so the full-colour logo is visible on dark AND light themes
    st.markdown(
        f'<div class="banner-logo" style="padding:16px 0 12px 20px;">'
        f'<div style="background:#ffffff;border-radius:10px;padding:8px 14px;'
        f'display:inline-block;box-shadow:0 2px 10px rgba(0,0,0,0.25);">'
        f'<img src="{_logo_src}" '
        f'style="height:36px;max-width:130px;object-fit:contain;display:block;" '
        f'onerror="this.parentElement.innerHTML=\'<span style=&quot;font-size:22px;'
        f'font-weight:900;color:#1e293b;letter-spacing:-1px;&quot;>spyne</span>\'" />'
        f'</div></div>',
        unsafe_allow_html=True,
    )

with _col_title:
    st.markdown(
        f'<div style="padding:14px 0 10px 8px;">'
        f'<div style="color:#f1f5f9;font-size:22px;font-weight:800;'
        f'letter-spacing:-0.5px;line-height:1.2;">AR Collections Dashboard</div>'
        f'<div style="color:#64748b;font-size:12px;margin-top:4px;">'
        f'Finance Team &nbsp;·&nbsp; Collections &nbsp;·&nbsp; Aging &nbsp;·&nbsp; Reminders'
        f'&nbsp;&nbsp;<span style="background:rgba(96,165,250,0.12);color:#93c5fd;'
        f'border:1px solid rgba(96,165,250,0.25);border-radius:20px;'
        f'padding:2px 10px;font-size:11px;font-weight:600;">👤 {_uname_display}</span>'
        f'&nbsp;<span style="background:{_role_color}22;color:{_role_color};'
        f'border:1px solid {_role_color}55;border-radius:20px;'
        f'padding:2px 10px;font-size:11px;font-weight:700;">{_role_label}</span>'
        f'</div></div>',
        unsafe_allow_html=True,
    )

with _col_refresh:
    st.markdown("<div style='padding-top:14px;'></div>", unsafe_allow_html=True)
    if _can("refresh_data"):
        if st.button("🔄 Refresh Data", use_container_width=True, type="primary"):
            fetch_gsheet.clear()
            st.session_state.pop("_gs_file_bytes", None)
            st.session_state["_gs_last_refresh"] = None
            st.rerun()
    else:
        st.markdown(
            "<div style='padding:8px 0;text-align:right;'>"
            "<span style='color:#6b7280;font-size:12px;'>🔒 Read-only access</span>"
            "</div>",
            unsafe_allow_html=True,
        )

# ── Data source selector ──────────────────────────────────────────────────────
from datetime import timezone, timedelta
_IST        = timezone(timedelta(hours=5, minutes=30))
_ist_time_str = ""

_src_tab_gs, _src_tab_zoho = st.tabs(["📄 Google Sheets", "🔗 Zoho Books (Live)"])

# ════════════════════════ SOURCE A · GOOGLE SHEETS ════════════════════════════
with _src_tab_gs:
    st.caption("Data is loaded from the fixed Google Sheet. Click **Refresh Data** (top-right) to re-fetch.")
    if "_gs_file_bytes" not in st.session_state or st.session_state["_gs_file_bytes"] is None:
        with st.spinner("Loading data from Google Sheets…"):
            try:
                file_bytes = fetch_gsheet(_FIXED_SHEET_URL)
                st.session_state["_gs_file_bytes"]   = file_bytes
                st.session_state["_gs_last_refresh"] = time.time()
                st.session_state["_active_source"]   = "gsheet"
            except PermissionError as e:
                st.error(str(e)); st.stop()
            except Exception as e:
                st.error(f"Error loading sheet: {e}"); st.stop()
    else:
        file_bytes = st.session_state["_gs_file_bytes"]
        st.session_state.setdefault("_active_source", "gsheet")

# ════════════════════════ SOURCE B · ZOHO BOOKS ═══════════════════════════════
with _src_tab_zoho:
    # Need credentials to be loaded first
    _zoho_ready_now = all([
        st.session_state.get("zoho_client_id","").strip(),
        st.session_state.get("zoho_client_secret","").strip(),
        st.session_state.get("zoho_refresh_token","").strip(),
        [o.strip() for o in [st.session_state.get("zoho_org_id_1",""),
                              st.session_state.get("zoho_org_id_2","")] if o.strip()],
    ])

    if not _zoho_ready_now:
        st.warning("⚙️ Zoho credentials not configured. Add them to `.streamlit/secrets.toml`.")
    else:
        _zb_c1, _zb_c2, _zb_c3 = st.columns([2, 2, 3])
        with _zb_c1:
            _zoho_statuses = st.multiselect(
                "Invoice Statuses to fetch",
                ["overdue", "sent", "draft", "paid", "partially_paid", "void"],
                default=["overdue", "sent"],
                key="zoho_pull_statuses",
            )
        with _zb_c2:
            _fetch_lineitems = st.checkbox(
                "📦 Include line items (products)",
                value=False, key="zoho_pull_lineitems",
                help="Fetches product/service details per invoice. Slower for large datasets."
            )
        with _zb_c3:
            st.markdown("<div style='padding-top:22px'></div>", unsafe_allow_html=True)
            _pull_zoho = st.button("🔗 Pull from Zoho Books", type="primary",
                                   use_container_width=True, key="pull_zoho_btn",
                                   disabled=not _can("zoho_pull"),
                                   help="🔒 Your role does not have permission to pull Zoho data." if not _can("zoho_pull") else None)

        if _pull_zoho:
            try:
                with st.spinner("🔑 Authenticating with Zoho Books…"):
                    _z_org_ids = [o.strip() for o in [
                        st.session_state.get("zoho_org_id_1",""),
                        st.session_state.get("zoho_org_id_2",""),
                    ] if o.strip()]
                    _z_dc = st.session_state.get("zoho_dc", "US (.com)")
                    _z_token = get_zoho_token(
                        st.session_state["zoho_client_id"],
                        st.session_state["zoho_client_secret"],
                        st.session_state["zoho_refresh_token"],
                        _z_dc,
                    )
                with st.spinner(f"📥 Fetching invoices ({', '.join(_zoho_statuses)}) from Zoho Books…"):
                    _z_df = fetch_zoho_all_invoices(_z_token, _z_org_ids, _z_dc,
                                                    statuses=_zoho_statuses)

                if _z_df.empty:
                    st.warning("No invoices returned from Zoho Books for the selected statuses.")
                else:
                    # Optional line items
                    if _fetch_lineitems:
                        _li_pairs = list(zip(
                            _z_df["_zoho_invoice_id"].tolist(),
                            _z_df["invoice_number"].tolist(),
                        ))
                        _prog = st.progress(0, text="Fetching line items…")
                        def _li_prog(i, total):
                            _prog.progress(int(i / max(total,1) * 100),
                                           text=f"Line items: {i}/{total}")
                        # Use first org_id that has invoices
                        _li_org = _z_org_ids[0]
                        _li_df = fetch_zoho_lineitems_for_invoices(
                            _li_pairs, _z_token, _li_org, _z_dc, _li_prog)
                        _prog.empty()
                        if not _li_df.empty:
                            st.session_state["_zoho_lineitems"] = _li_df
                            st.success(f"✅ {len(_li_df):,} line-item rows fetched")

                    st.session_state["_zoho_df"]           = _z_df
                    st.session_state["_zoho_last_refresh"] = time.time()
                    st.session_state["_active_source"]     = "zoho"
                    st.success(f"✅ {len(_z_df):,} invoices loaded from Zoho Books")
            except Exception as _ze:
                st.error(f"Zoho pull failed: {_ze}")

        # Show last-pull info & line-items preview
        if st.session_state.get("_active_source") == "zoho":
            _zts = st.session_state.get("_zoho_last_refresh")
            if _zts:
                _zt_str = datetime.fromtimestamp(_zts, tz=_IST).strftime('%d %b %Y, %I:%M:%S %p IST')
                st.caption(f"🕐 Zoho data as of **{_zt_str}**")

            if "_zoho_lineitems" in st.session_state and not st.session_state["_zoho_lineitems"].empty:
                with st.expander("📦 Line Items (Product-level)", expanded=False):
                    _li_show = st.session_state["_zoho_lineitems"]
                    st.dataframe(_li_show, use_container_width=True, height=300)
                    st.download_button(
                        "⬇ Download Line Items",
                        data=export_excel(_li_show),
                        file_name="zoho_line_items.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_lineitems",
                    )

# ── Resolve active DataFrame ──────────────────────────────────────────────────
if st.session_state.get("_active_source") == "zoho" and "_zoho_df" in st.session_state:
    _active_df_raw = st.session_state["_zoho_df"]
    _zts = st.session_state.get("_zoho_last_refresh")
    if _zts:
        _ist_time_str = datetime.fromtimestamp(_zts, tz=_IST).strftime('%d %b %Y, %I:%M:%S %p IST')
    df = _active_df_raw.copy()
    if "date" not in df.columns:
        df["date"] = pd.NaT
else:
    # Google Sheets path
    if "file_bytes" not in dir() or file_bytes is None:
        file_bytes = st.session_state.get("_gs_file_bytes")
    _last = st.session_state.get("_gs_last_refresh")
    if _last:
        _ist_time_str = datetime.fromtimestamp(_last, tz=_IST).strftime('%d %b %Y, %I:%M:%S %p IST')
    df = load_data(file_bytes)

# ── Exclude fully-paid / zero-balance invoices everywhere ─────────────────────
if "balance" in df.columns:
    df = df[df["balance"].fillna(0) > 0].copy()

# ── CSM role: restrict data to the user's assigned CSM ────────────────────────
if _can("csm_filter") and "CSM" in df.columns:
    _csm_user     = st.session_state.get("_username", "")
    _csm_assigned = _get_csm_name_for_user(_csm_user)
    # Case-insensitive match against the CSM column
    _csm_mask = df["CSM"].astype(str).str.strip().str.lower() == _csm_assigned.strip().lower()
    if _csm_mask.any():
        df = df[_csm_mask].copy()
    else:
        # No match — show warning but don't lock out entirely
        st.warning(
            f"⚠️ No data found for CSM **{_csm_assigned}** (your assigned CSM name). "
            "Ask an Admin to set your CSM assignment correctly in User Management."
        )

# Note: Invoice Status filtering is handled by the sidebar multiselect (default: overdue + sent)

# ─── Column presence helpers ──────────────────────────────────────────────────
def col(name):
    """Return df[name] if it exists, else an empty Series."""
    return df[name] if name in df.columns else pd.Series(dtype=str)

def fcol(name):
    """Return fdf[name] if it exists, else an empty Series."""
    return fdf[name] if name in fdf.columns else pd.Series(dtype=str)

# ─── Missing critical columns warning ────────────────────────────────────────
CRITICAL = ["customer_name", "invoice_number", "CSM", "Final USD"]
missing  = [c for c in CRITICAL if c not in df.columns]
if missing:
    st.warning(
        f"⚠️ Could not find these expected columns: **{', '.join(missing)}**\n\n"
        f"Columns detected: `{'` · `'.join(df.columns.tolist())}`\n\n"
        "Check the **Detected columns** expander in the sidebar. "
        "Rename your sheet headers to match or ask for help."
    )

# ─── Sidebar filters (global) ─────────────────────────────────────────────────
with st.sidebar:
    st.header("Global Filters")
    csm_options = sorted(col("CSM").dropna().unique())
    csm_sel = st.multiselect("CSM", csm_options)

    rag_options = sorted(col("RAG").dropna().unique())
    rag_sel = st.multiselect("RAG Status", rag_options)

    bucket_options = [b for b in BUCKET_ORDER if b in col("Bucket").values]
    bucket_sel = st.multiselect("Aging Bucket", bucket_options)

    country_options = sorted(col("country").dropna().unique())
    country_sel = st.multiselect("Country", country_options)

    product_options = sorted(col("Product").dropna().unique())
    product_sel = st.multiselect("Product", product_options)

    if "Current Invoice Status" in df.columns:
        _all_statuses   = sorted(df["Current Invoice Status"].astype(str).str.strip().dropna().unique())
        _default_statuses = [s for s in _all_statuses if s.lower() in ("overdue", "sent")]
        inv_status_sel  = st.multiselect(
            "Invoice Status",
            options=_all_statuses,
            default=_default_statuses,
            key="sidebar_inv_status",
            help="Default shows Overdue & Sent. Add other statuses to include them.",
        )
    else:
        inv_status_sel = []

    st.divider()
    cust_count = df["customer_name"].nunique() if "customer_name" in df.columns else "?"
    st.caption(f"Dataset: **{len(df):,}** invoices · **{cust_count}** customers")

    with st.expander("🔍 Detected columns"):
        raw = st.session_state.get("_raw_cols", list(df.columns))
        st.write(raw)


    # ── SMTP config — read silently from secrets / credentials.json ──────────
    SMTP_CFG = {
        "host":     "smtp.gmail.com",
        "port":     587,
        "user":     st.session_state.get("smtp_user",   ""),
        "password": st.session_state.get("smtp_pass",   ""),
        "sender":   st.session_state.get("smtp_sender", "finance@spyne.ai"),
        "use_tls":  True,
    }

    # ── Zoho config — read silently from secrets / credentials.json ──────────
    # Build org_ids list (filter empty)
    _org_ids = [o.strip() for o in [
        st.session_state.get("zoho_org_id_1", ""),
        st.session_state.get("zoho_org_id_2", ""),
    ] if o.strip()]

    ZOHO_CFG = {
        "org_ids":       _org_ids,
        "client_id":     st.session_state.get("zoho_client_id",     "").strip(),
        "client_secret": st.session_state.get("zoho_client_secret", "").strip(),
        "refresh_token": st.session_state.get("zoho_refresh_token", "").strip(),
        "dc":            st.session_state.get("zoho_dc", "US (.com)"),
    }
    ZOHO_READY = all([ZOHO_CFG["org_ids"], ZOHO_CFG["client_id"],
                      ZOHO_CFG["client_secret"], ZOHO_CFG["refresh_token"]])

# Apply global filters
fdf = df.copy()
if csm_sel:          fdf = fdf[fdf["CSM"].isin(csm_sel)]                             if "CSM"                    in fdf.columns else fdf
if rag_sel:          fdf = fdf[fdf["RAG"].isin(rag_sel)]                             if "RAG"                    in fdf.columns else fdf
if bucket_sel:       fdf = fdf[fdf["Bucket"].isin(bucket_sel)]                       if "Bucket"                 in fdf.columns else fdf
if country_sel:      fdf = fdf[fdf["country"].isin(country_sel)]                     if "country"                in fdf.columns else fdf
if product_sel:      fdf = fdf[fdf["Product"].isin(product_sel)]                     if "Product"                in fdf.columns else fdf
if inv_status_sel and "Current Invoice Status" in fdf.columns:
    fdf = fdf[fdf["Current Invoice Status"].isin(inv_status_sel)]

# ── KPI strip + last-refresh bar ─────────────────────────────────────────────
_ks_total_usd   = fdf["Final USD"].sum()   if "Final USD"      in fdf.columns else 0
_ks_customers   = fdf["customer_name"].nunique() if "customer_name" in fdf.columns else 0
_ks_invoices    = len(fdf)
_ks_overdue     = (fdf[fdf["RAG"] == "Red"]["Final USD"].sum()
                   if "RAG" in fdf.columns and "Final USD" in fdf.columns else 0)
_ks_avg_aging   = int(fdf["Aging"].mean()) if "Aging" in fdf.columns and len(fdf) else 0

# ── KPI cards — one st.markdown per column (avoids markdown parser issues) ────
_kpi_data = [
    ("💵", "Total Outstanding (USD)", f"${_ks_total_usd:,.0f}", "#60a5fa"),
    ("🔴", "At Risk (Red)",           f"${_ks_overdue:,.0f}",   "#f87171"),
    ("🏢", "Customers",               f"{_ks_customers:,}",     "#34d399"),
    ("🧾", "Invoices",                f"{_ks_invoices:,}",      "#a78bfa"),
    ("⏳", "Avg Aging (days)",        f"{_ks_avg_aging}",       "#fbbf24"),
]
_kpi_cols = st.columns(5)
for _col, (_icon, _label, _val, _accent) in zip(_kpi_cols, _kpi_data):
    _col.markdown(
        f"<div style='background:rgba(255,255,255,0.03);border:1px solid rgba(255,255,255,0.08);"
        f"border-radius:10px;padding:14px 16px;'>"
        f"<div style='color:{_accent};font-size:20px;margin-bottom:4px;'>{_icon}</div>"
        f"<div style='color:#f1f5f9;font-size:22px;font-weight:800;line-height:1.1;'>{_val}</div>"
        f"<div style='color:#64748b;font-size:11px;font-weight:600;text-transform:uppercase;"
        f"letter-spacing:0.06em;margin-top:4px;'>{_label}</div>"
        f"</div>",
        unsafe_allow_html=True,
    )

# ── Last-refresh note ─────────────────────────────────────────────────────────
if _ist_time_str:
    st.caption(f"🕐 Data as of **{_ist_time_str}** · Click **Refresh Data** to fetch latest from Google Sheets")

# ── TABS ──────────────────────────────────────────────────────────────────────
tab_overview, tab_csm, tab_customer, tab_invoices, tab_reasons, tab_email = st.tabs([
    "📈 Overview",
    "👤 CSM Summary",
    "🏢 Customer Summary",
    "🔍 Invoice Drilldown",
    "📝 Reasons & Actions",
    "📧 Send Reminders",
])

# ─────────────────────────── TAB 1 · OVERVIEW ────────────────────────────────
with tab_overview:
    if not _can("view_overview"):
        st.info("📈 The Overview dashboard is not available for your role. "
                "You have access to CSM Summary, Customer Summary, Invoice Drilldown, "
                "and Reasons & Actions.")
    else:
        total_inr   = fdf["Outstanding"].sum() if "Outstanding" in fdf.columns else 0
        n_invoices  = len(fdf)
        n_customers = fdf["customer_name"].nunique() if "customer_name" in fdf.columns else 0
        n_csms      = fdf["CSM"].nunique()           if "CSM"           in fdf.columns else 0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Outstanding (INR)", fmt_inr(total_inr))
        k2.metric("Invoices",          f"{n_invoices:,}")
        k3.metric("Customers",         f"{n_customers:,}")
        k4.metric("CSMs",              f"{n_csms:,}")

        st.divider()
        c1, c2 = st.columns(2)

        with c1:
            if "RAG" in fdf.columns:
                rag_data = fdf.groupby("RAG")["Final USD"].sum().reset_index()
                fig = px.pie(
                    rag_data, values="Final USD", names="RAG",
                    title="Outstanding by RAG Status",
                    color="RAG",
                    color_discrete_map=RAG_COLORS,
                    hole=0.4,
                )
                fig.update_traces(textinfo="percent+label")
                st.plotly_chart(_fmt_fig(fig), use_container_width=True)

        with c2:
            if "Bucket" in fdf.columns and "RAG" in fdf.columns:
                bucket_rag = (
                    fdf.groupby(["Bucket", "RAG"])["Final USD"]
                    .sum()
                    .reset_index()
                )
                # Enforce correct bucket order
                bucket_rag["Bucket"] = pd.Categorical(bucket_rag["Bucket"], categories=BUCKET_ORDER, ordered=True)
                bucket_rag = bucket_rag.sort_values("Bucket")
                fig = px.bar(
                    bucket_rag, x="Bucket", y="Final USD", color="RAG",
                    title="Outstanding by Aging Bucket & RAG",
                    color_discrete_map=RAG_COLORS,
                    category_orders={"Bucket": BUCKET_ORDER, "RAG": ["Green", "Amber", "Red"]},
                    text_auto=",.0f",
                )
                fig.update_layout(xaxis_title="", yaxis_title="USD", legend_title="RAG")
                st.plotly_chart(_fmt_fig(fig), use_container_width=True)

        st.divider()

        # ── Bucket × RAG summary table ────────────────────────────────────────────
        if "Bucket" in fdf.columns and "RAG" in fdf.columns:
            st.subheader("Aging Bucket × RAG Breakdown")
            pivot = (
                fdf.groupby(["Bucket", "RAG"])["Final USD"]
                .sum()
                .unstack(fill_value=0)
                .reindex(BUCKET_ORDER)
            )
            # Add totals
            for col in ["Red", "Amber", "Green"]:
                if col not in pivot.columns:
                    pivot[col] = 0
            pivot = pivot[["Green", "Amber", "Red"]]
            pivot["Total"] = pivot.sum(axis=1)
            pivot.loc["Grand Total"] = pivot.sum()

            fmt_map = {c: "${:,.0f}" for c in pivot.columns}
            def _col_color(col):
                colors = {"Green": "color:#10b981;font-weight:600",
                          "Amber": "color:#f59e0b;font-weight:600",
                          "Red":   "color:#ef4444;font-weight:600"}
                return [colors.get(col.name, "")] * len(col)

            styled = pivot.style.format(fmt_map).apply(_col_color, axis=0)
            st.dataframe(styled, use_container_width=True)

        c3, c4 = st.columns(2)

        with c3:
            if "country" in fdf.columns:
                country_data = (fdf.groupby("country")["Final USD"].sum()
                                 .reset_index()
                                 .sort_values("Final USD", ascending=False)
                                 .head(10))
                fig = px.bar(country_data, x="Final USD", y="country",
                             orientation="h", title="Top 10 Countries by Outstanding",
                             text_auto=",.0f")
                fig.update_layout(yaxis=dict(autorange="reversed"), yaxis_title="")
                st.plotly_chart(_fmt_fig(fig), use_container_width=True)

        with c4:
            if "Product" in fdf.columns:
                product_data = (fdf.groupby("Product")["Final USD"].sum()
                                  .reset_index()
                                  .sort_values("Final USD", ascending=False)
                                  .head(10))
                fig = px.bar(product_data, x="Final USD", y="Product",
                             orientation="h", title="Top 10 Products by Outstanding",
                             text_auto=",.0f")
                fig.update_layout(yaxis=dict(autorange="reversed"), yaxis_title="")
                st.plotly_chart(_fmt_fig(fig), use_container_width=True)

        # ── Currency-wise outstanding (FC) with RAG ───────────────────────────────
        if "currency_code" in fdf.columns and "balance" in fdf.columns and "RAG" in fdf.columns:
            st.divider()
            st.subheader("Currency-wise Outstanding (FC) by RAG")

            curr_rag = (
                fdf.groupby(["currency_code", "RAG"])["balance"]
                .sum()
                .reset_index()
                .sort_values("balance", ascending=False)
            )

            cc1, cc2 = st.columns(2)

            with cc1:
                fig = px.bar(
                    curr_rag, x="currency_code", y="balance", color="RAG",
                    title="FC Outstanding by Currency & RAG",
                    color_discrete_map=RAG_COLORS,
                    category_orders={"RAG": ["Green", "Amber", "Red"]},
                    text_auto=",.0f",
                )
                fig.update_layout(
                    xaxis_title="Currency", yaxis_title="FC Amount",
                    legend_title="RAG", xaxis_tickangle=-30,
                )
                st.plotly_chart(_fmt_fig(fig), use_container_width=True)

            with cc2:
                # Pivot: currency rows × RAG columns
                pivot_curr = (
                    curr_rag.pivot_table(index="currency_code", columns="RAG",
                                         values="balance", aggfunc="sum", fill_value=0)
                    .reset_index()
                )
                for r in ["Green", "Amber", "Red"]:
                    if r not in pivot_curr.columns:
                        pivot_curr[r] = 0
                pivot_curr = pivot_curr[["currency_code", "Green", "Amber", "Red"]]
                pivot_curr["Total (FC)"] = pivot_curr[["Green","Amber","Red"]].sum(axis=1)
                pivot_curr = pivot_curr.sort_values("Total (FC)", ascending=False)

                # Append totals row
                totals = {"currency_code": "Grand Total",
                          "Green": pivot_curr["Green"].sum(),
                          "Amber": pivot_curr["Amber"].sum(),
                          "Red":   pivot_curr["Red"].sum(),
                          "Total (FC)": pivot_curr["Total (FC)"].sum()}
                pivot_curr = pd.concat([pivot_curr, pd.DataFrame([totals])], ignore_index=True)

                CURRENCY_SYMBOLS = {
                    "INR": "₹", "USD": "$", "EUR": "€", "GBP": "£",
                    "AUD": "A$", "CAD": "C$", "NZD": "NZ$", "SGD": "S$",
                    "HKD": "HK$", "JPY": "¥", "CNY": "¥", "CHF": "CHF ",
                    "NOK": "kr ", "SEK": "kr ", "DKK": "kr ", "AED": "AED ",
                    "SAR": "﷼", "MYR": "RM ", "THB": "฿", "IDR": "Rp ",
                    "PHP": "₱", "KRW": "₩", "BRL": "R$", "MXN": "MX$",
                    "ZAR": "R ",
                }

                def fmt_currency_val(val, currency):
                    sym = CURRENCY_SYMBOLS.get(str(currency).upper(), "")
                    return f"{sym}{val:,.0f}"

                def _rag_style(col):
                    colors = {"Green": "color:#34d399;font-weight:600",
                              "Amber": "color:#fbbf24;font-weight:600",
                              "Red":   "color:#f87171;font-weight:600"}
                    return [colors.get(col.name, "")] * len(col)

                # Format each row with its own currency symbol
                display_curr = pivot_curr.rename(columns={"currency_code": "Currency"}).copy()
                for num_col in ["Green", "Amber", "Red", "Total (FC)"]:
                    display_curr[num_col] = display_curr.apply(
                        lambda r: fmt_currency_val(r[num_col],
                            r["Currency"] if r["Currency"] != "Grand Total" else ""),
                        axis=1,
                    )

                st.dataframe(
                    display_curr.style.apply(_rag_style, axis=0),
                    use_container_width=True,
                    hide_index=True,
                )

# ─────────────────────────── TAB 2 · CSM SUMMARY ─────────────────────────────
with tab_csm:
    # ── Pick value column ─────────────────────────────────────────────────────
    val_col = "Outstanding" if "Outstanding" in fdf.columns else "Final USD"
    val_fmt = "₹{:,.0f}"   if val_col == "Outstanding"      else "${:,.0f}"

    # ── Base aggregation using simple identifier names, then rename ───────────
    grp = fdf.groupby("CSM")
    csm_df = pd.DataFrame()
    csm_df["CSM"]              = grp[val_col].sum().index
    csm_df = csm_df.set_index("CSM")
    csm_df["Total Outstanding"] = grp[val_col].sum()
    csm_df["Avg Aging (days)"]  = grp["Aging"].mean()                    if "Aging"          in fdf.columns else 0
    csm_df["No. of Invoices"]   = grp["invoice_number"].count()          if "invoice_number" in fdf.columns else 0
    csm_df["No. of Customers"]  = grp["customer_name"].nunique()         if "customer_name"  in fdf.columns else 0
    csm_df = csm_df.reset_index()

    # ── RAG breakdown ─────────────────────────────────────────────────────────
    if "RAG" in fdf.columns:
        for rag in ["Red", "Amber", "Green"]:
            sub = fdf[fdf["RAG"] == rag].groupby("CSM")[val_col].sum().rename(f"{rag} Outstanding")
            csm_df = csm_df.merge(sub, on="CSM", how="left")
            csm_df[f"{rag} Outstanding"] = csm_df[f"{rag} Outstanding"].fillna(0)

    csm_df = csm_df.sort_values("Total Outstanding", ascending=False)

    # ── Column order ──────────────────────────────────────────────────────────
    rag_out_cols = [c for c in ["Red Outstanding", "Amber Outstanding", "Green Outstanding"]
                    if c in csm_df.columns]
    col_order = (["CSM", "Total Outstanding"]
                 + rag_out_cols
                 + ["Avg Aging (days)", "No. of Invoices", "No. of Customers"])
    csm_df = csm_df[[c for c in col_order if c in csm_df.columns]]

    # ── Search ────────────────────────────────────────────────────────────────
    search = st.text_input("🔍 Search CSM", placeholder="Type to filter…")
    show_df = csm_df[csm_df["CSM"].str.contains(search, case=False, na=False)] if search else csm_df

    # ── Format & style ────────────────────────────────────────────────────────
    fmt_map = {
        "Total Outstanding": val_fmt,
        "Avg Aging (days)":  "{:.0f}",
        "No. of Invoices":   "{:,.0f}",
        "No. of Customers":  "{:,.0f}",
    }
    for c in rag_out_cols:
        fmt_map[c] = val_fmt

    def _rag_col_style(col):
        colors = {
            "Red":   "color:#f87171;font-weight:700",
            "Amber": "color:#fbbf24;font-weight:700",
            "Green": "color:#34d399;font-weight:700",
        }
        for key, style in colors.items():
            if col.name.startswith(key):
                return [style] * len(col)
        return [""] * len(col)

    st.dataframe(
        show_df.style.format(fmt_map).apply(_rag_col_style, axis=0),
        use_container_width=True,
        height=min(80 + len(show_df) * 35, 520),
        hide_index=True,
    )

    st.download_button(
        "⬇ Download CSM Summary",
        data=export_excel(show_df),
        file_name="csm_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    # ── Bar chart – top CSMs ──────────────────────────────────────────────────
    _csm_max = max(len(csm_df), 1)
    top_n = st.slider("Show top N CSMs", 5, max(5, min(30, _csm_max)), min(15, max(5, _csm_max))) if _csm_max > 5 else _csm_max
    fig = px.bar(
        csm_df.head(top_n), x="CSM", y="Total Outstanding",
        title=f"Top {top_n} CSMs by Outstanding",
        color="Total Outstanding", color_continuous_scale="Reds",
        text_auto=",.0f",
    )
    fig.update_layout(xaxis_tickangle=-30, coloraxis_showscale=False)
    st.plotly_chart(_fmt_fig(fig), use_container_width=True)

    # ── Stacked RAG bar ───────────────────────────────────────────────────────
    if rag_out_cols:
        st.subheader("RAG Breakdown per CSM")
        melted = csm_df.head(top_n).melt(
            id_vars="CSM", value_vars=rag_out_cols, var_name="RAG", value_name="Amount"
        )
        melted["RAG"] = melted["RAG"].apply(lambda x: x.split()[0])  # strip suffix
        fig2 = px.bar(melted, x="CSM", y="Amount", color="RAG",
                      color_discrete_map=RAG_COLORS,
                      title="Outstanding by RAG per CSM", text_auto=",.0f",
                      labels={"Amount": "Outstanding (₹)" if val_col == "Outstanding" else "Outstanding (USD)"})
        fig2.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(_fmt_fig(fig2), use_container_width=True)

    st.divider()
    st.subheader("CSM Deep Dive")
    selected_csm = st.selectbox("Select a CSM to drill in", sorted(fdf["CSM"].dropna().unique()))
    csm_detail = fdf[fdf["CSM"] == selected_csm]

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Outstanding (INR)", fmt_inr(csm_detail["Outstanding"].sum()) if "Outstanding" in csm_detail.columns else "—")
    d2.metric("Outstanding (USD)", fmt_usd(csm_detail["Final USD"].sum()) if "Final USD" in csm_detail.columns else "—")
    d3.metric("Invoices", len(csm_detail))
    d4.metric("Customers", csm_detail["customer_name"].nunique())

    show_cols = [c for c in ["invoice_number","customer_name","Outstanding","Final USD","Aging","Bucket","RAG","due_date","Status","Product","country"] if c in csm_detail.columns]
    csm_detail_show = csm_detail[show_cols].copy()
    if "Final USD" in csm_detail_show.columns:
        csm_detail_show = csm_detail_show.sort_values("Final USD", ascending=False)
    csm_detail_show = column_filters(csm_detail_show, key_prefix="csm_dd")
    st.dataframe(csm_detail_show, use_container_width=True)

# ─────────────────────────── TAB 3 · CUSTOMER SUMMARY ───────────────────────
with tab_customer:
    val_col_c = "Outstanding" if "Outstanding" in fdf.columns else "Final USD"
    val_fmt_c = "₹{:,.0f}"   if val_col_c == "Outstanding"   else "${:,.0f}"
    val_lbl_c = "Outstanding (₹)" if val_col_c == "Outstanding" else "Outstanding (USD)"

    # ── Aggregation ───────────────────────────────────────────────────────────
    grp_c = fdf.groupby("customer_name")
    cust_df = pd.DataFrame()
    cust_df["Customer"]           = grp_c[val_col_c].sum().index
    cust_df = cust_df.set_index("Customer")
    cust_df["Total Outstanding"]  = grp_c[val_col_c].sum()
    cust_df["Avg Aging (days)"]   = grp_c["Aging"].mean()               if "Aging"          in fdf.columns else 0
    cust_df["No. of Invoices"]    = grp_c["invoice_number"].count()     if "invoice_number" in fdf.columns else 0
    # CSM — take the most frequent CSM per customer
    if "CSM" in fdf.columns:
        cust_df["CSM"] = fdf.groupby("customer_name")["CSM"].agg(
            lambda x: x.value_counts().index[0] if len(x) else ""
        )
    # Country
    if "country" in fdf.columns:
        cust_df["Country"] = fdf.groupby("customer_name")["country"].agg(
            lambda x: x.value_counts().index[0] if len(x) else ""
        )
    cust_df = cust_df.reset_index()

    # ── RAG breakdown ─────────────────────────────────────────────────────────
    if "RAG" in fdf.columns:
        for rag in ["Red", "Amber", "Green"]:
            sub = (fdf[fdf["RAG"] == rag]
                   .groupby("customer_name")[val_col_c].sum()
                   .rename(f"{rag} Outstanding"))
            cust_df = cust_df.merge(sub, left_on="Customer",
                                    right_on="customer_name", how="left")
            cust_df[f"{rag} Outstanding"] = cust_df[f"{rag} Outstanding"].fillna(0)

    cust_df = cust_df.sort_values("Total Outstanding", ascending=False)

    # ── Column order ──────────────────────────────────────────────────────────
    rag_out_cols_c = [c for c in ["Red Outstanding", "Amber Outstanding", "Green Outstanding"]
                      if c in cust_df.columns]
    col_order_c = (["Customer"]
                   + (["CSM"] if "CSM" in cust_df.columns else [])
                   + (["Country"] if "Country" in cust_df.columns else [])
                   + ["Total Outstanding"]
                   + rag_out_cols_c
                   + ["Avg Aging (days)", "No. of Invoices"])
    cust_df = cust_df[[c for c in col_order_c if c in cust_df.columns]]

    # ── Top KPIs ──────────────────────────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Customers",    f"{len(cust_df):,}")
    k2.metric("Total Outstanding",  fmt_inr(cust_df["Total Outstanding"].sum()) if val_col_c == "Outstanding"
                                    else fmt_usd(cust_df["Total Outstanding"].sum()))
    k3.metric("Avg Outstanding/Customer",
              fmt_inr(cust_df["Total Outstanding"].mean()) if val_col_c == "Outstanding"
              else fmt_usd(cust_df["Total Outstanding"].mean()))
    k4.metric("Avg Aging (days)",   f"{cust_df['Avg Aging (days)'].mean():.0f}" if "Avg Aging (days)" in cust_df.columns else "—")

    st.divider()

    # ── Filters row ───────────────────────────────────────────────────────────
    fc1, fc2, fc3 = st.columns([2, 2, 2])
    with fc1:
        csm_filter = st.multiselect("Filter by CSM", sorted(fdf["CSM"].dropna().unique()),
                                    key="cust_csm_filter") if "CSM" in fdf.columns else []
    with fc2:
        rag_filter = st.multiselect("Filter by RAG", ["Red", "Amber", "Green"],
                                    key="cust_rag_filter") if "RAG" in fdf.columns else []
    with fc3:
        search_c = st.text_input("🔍 Search Customer", placeholder="Type to filter…", key="cust_search")

    show_cust = cust_df.copy()
    if csm_filter:
        show_cust = show_cust[show_cust["CSM"].isin(csm_filter)]
    if rag_filter:
        # keep customers that have ANY invoice in the selected RAG buckets
        cust_in_rag = fdf[fdf["RAG"].isin(rag_filter)]["customer_name"].unique()
        show_cust = show_cust[show_cust["Customer"].isin(cust_in_rag)]
    if search_c:
        show_cust = show_cust[show_cust["Customer"].str.contains(search_c, case=False, na=False)]

    # ── Format & style ────────────────────────────────────────────────────────
    fmt_map_c = {
        "Total Outstanding": val_fmt_c,
        "Avg Aging (days)":  "{:.0f}",
        "No. of Invoices":   "{:,.0f}",
    }
    for c in rag_out_cols_c:
        fmt_map_c[c] = val_fmt_c

    def _rag_cust_style(col):
        colors = {
            "Red":   "color:#f87171;font-weight:700",
            "Amber": "color:#fbbf24;font-weight:700",
            "Green": "color:#34d399;font-weight:700",
        }
        for key, style in colors.items():
            if col.name.startswith(key):
                return [style] * len(col)
        return [""] * len(col)

    st.dataframe(
        show_cust.style.format(fmt_map_c).apply(_rag_cust_style, axis=0),
        use_container_width=True,
        height=min(80 + len(show_cust) * 35, 520),
        hide_index=True,
    )

    st.download_button(
        "⬇ Download Customer Summary",
        data=export_excel(show_cust),
        file_name="customer_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()

    # ── Bar chart – top customers ──────────────────────────────────────────────
    _cust_max = max(len(cust_df), 1)
    if _cust_max <= 5:
        top_n_c = _cust_max
    else:
        top_n_c = st.slider("Show top N customers", 5, min(30, _cust_max), min(15, _cust_max),
                            key="cust_top_n")
    fig_c = px.bar(
        cust_df.head(top_n_c), x="Customer", y="Total Outstanding",
        title=f"Top {top_n_c} Customers by Outstanding",
        color="Total Outstanding", color_continuous_scale="Reds",
        text_auto=",.0f",
    )
    fig_c.update_layout(xaxis_tickangle=-30, coloraxis_showscale=False)
    st.plotly_chart(_fmt_fig(fig_c), use_container_width=True)

    # ── Stacked RAG bar ───────────────────────────────────────────────────────
    if rag_out_cols_c:
        st.subheader("RAG Breakdown per Customer")
        melted_c = cust_df.head(top_n_c).melt(
            id_vars="Customer", value_vars=rag_out_cols_c, var_name="RAG", value_name="Amount"
        )
        melted_c["RAG"] = melted_c["RAG"].apply(lambda x: x.split()[0])
        fig_c2 = px.bar(
            melted_c, x="Customer", y="Amount", color="RAG",
            color_discrete_map=RAG_COLORS,
            title="Outstanding by RAG per Customer", text_auto=",.0f",
            labels={"Amount": val_lbl_c},
        )
        fig_c2.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(_fmt_fig(fig_c2), use_container_width=True)

    st.divider()

    # ── Customer Deep Dive ────────────────────────────────────────────────────
    st.subheader("Customer Deep Dive")
    selected_cust = st.selectbox(
        "Select a customer to drill in",
        sorted(fdf["customer_name"].dropna().unique()),
        key="cust_deep_select",
    )
    cust_detail = fdf[fdf["customer_name"] == selected_cust]

    cd1, cd2, cd3, cd4 = st.columns(4)
    cd1.metric("Outstanding (INR)", fmt_inr(cust_detail["Outstanding"].sum())
               if "Outstanding" in cust_detail.columns else "—")
    cd2.metric("Outstanding (USD)", fmt_usd(cust_detail["Final USD"].sum())
               if "Final USD" in cust_detail.columns else "—")
    cd3.metric("Invoices", len(cust_detail))
    cd4.metric("CSM", cust_detail["CSM"].mode()[0] if "CSM" in cust_detail.columns and len(cust_detail) else "—")

    dd_cols = [c for c in [
        "invoice_number", "CSM", "currency_code", "Outstanding", "Final USD",
        "Aging", "Bucket", "RAG", "due_date", "Status", "Product", "country",
    ] if c in cust_detail.columns]
    cust_detail_show = cust_detail[dd_cols].sort_values(
        "Outstanding" if "Outstanding" in dd_cols else dd_cols[0], ascending=False
    )
    cust_detail_show = column_filters(cust_detail_show, key_prefix="cust_dd")
    st.dataframe(cust_detail_show, use_container_width=True)


# ─────────────────────────── TAB 4 · INVOICE DRILLDOWN ───────────────────────
with tab_invoices:
    if not _can("invoice_drilldown"):
        st.info("📊 Invoice-level detail is not available for your role. "
                "Please contact an Admin if you need access.")
    else:
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            inv_csm = st.multiselect("CSM", sorted(fdf["CSM"].dropna().unique()), key="inv_csm")
        with f2:
            inv_rag = st.multiselect("RAG", sorted(fdf["RAG"].dropna().unique()) if "RAG" in fdf.columns else [], key="inv_rag")
        with f3:
            inv_bkt = st.multiselect("Bucket", [b for b in BUCKET_ORDER if b in fdf.get("Bucket", pd.Series()).values], key="inv_bkt")
        with f4:
            inv_status = st.multiselect("Status", sorted(fdf["Status"].dropna().unique()) if "Status" in fdf.columns else [], key="inv_status")

        filtered = fdf.copy()
        if inv_csm:    filtered = filtered[filtered["CSM"].isin(inv_csm)]
        if inv_rag:    filtered = filtered[filtered["RAG"].isin(inv_rag)]
        if inv_bkt:    filtered = filtered[filtered["Bucket"].isin(inv_bkt)]
        if inv_status: filtered = filtered[filtered["Status"].isin(inv_status)]

        if "balance" in filtered.columns and len(filtered):
            min_v, max_v = float(filtered["balance"].min()), float(filtered["balance"].max())
            if min_v < max_v:
                rng = st.slider("Filter by Balance Amount (FC)", min_v, max_v, (min_v, max_v), step=100.0)
                filtered = filtered[(filtered["balance"] >= rng[0]) & (filtered["balance"] <= rng[1])]

        total_inr_disp = fmt_inr(filtered["Outstanding"].sum()) if "Outstanding" in filtered.columns else "—"
        st.caption(f"Showing **{len(filtered):,}** invoices — Outstanding (INR): **{total_inr_disp}**")

        # ── Merge saved invoice-level reasons into the table ──────────────────────
        inv_reasons = get_reasons("invoice")
        base_cols = [c for c in [
            "invoice_number", "customer_name", "CSM",
            "currency_code", "balance",               # native currency amount
            "Aging", "Bucket", "RAG", "due_date", "Status",
            "Product", "country", "Billing Terms", "Service Type",
        ] if c in filtered.columns]

        display = filtered[base_cols].sort_values("balance", ascending=False).copy()
        display["invoice_number"] = display["invoice_number"].astype(str)

        # ── Reminder count per invoice ────────────────────────────────────────────
        _reminder_counts = get_reminder_counts()
        display["Reminders Sent"] = display["invoice_number"].map(_reminder_counts).fillna(0).astype(int)

        # Build a formatted "Amount" column: currency symbol + balance
        if "balance" in display.columns and "currency_code" in display.columns:
            display["Amount"] = display.apply(
                lambda r: f"{CURR_SYM.get(str(r['currency_code']).upper(), '')}{r['balance']:,.0f}",
                axis=1,
            )
            # keep balance as numeric for sort/slider but show Amount as the readable column
            display = display.drop(columns=["balance"])
            # reorder: put Amount right after CSM
            cols_order = ["invoice_number","customer_name","CSM","Amount","currency_code",
                          "Aging","Bucket","RAG","due_date","Status",
                          "Product","country","Billing Terms","Service Type"]
            display = display[[c for c in cols_order if c in display.columns]]

        # Attach existing reason columns
        if not inv_reasons.empty:
            inv_reasons = inv_reasons.rename(columns={
                "identifier":      "invoice_number",
                "reason_category": "Reason Category",
                "reason_text":     "Notes",
                "action_owner":    "Action Owner",
                "next_action_date":"Next Action Date",
            })[["invoice_number","Reason Category","Notes","Action Owner","Next Action Date"]]
            display = display.merge(inv_reasons, on="invoice_number", how="left")
        else:
            display["Reason Category"] = None
            display["Notes"]           = None
            display["Action Owner"]    = None
            display["Next Action Date"]= None

        display["Reason Category"]  = display["Reason Category"].astype(str).replace("nan","")
        display["Notes"]            = display["Notes"].astype(str).replace("nan","")
        display["Action Owner"]     = display["Action Owner"].astype(str).replace("nan","")
        display["Next Action Date"] = display["Next Action Date"].astype(str).replace("nan","")

        # ── Column filters ────────────────────────────────────────────────────────
        display = column_filters(display, key_prefix="inv_dd")

        # ── Editable table ────────────────────────────────────────────────────────
        st.info("Edit **Reason Category**, **Notes**, **Action Owner**, and **Next Action Date** inline. Click **Save Changes** when done.")

        edited = st.data_editor(
            display,
            use_container_width=True,
            height=480,
            disabled=[c for c in display.columns if c not in
                      ["Reason Category", "Notes", "Action Owner", "Next Action Date"]],
            column_config={
                "Amount":    st.column_config.TextColumn("Amount (FC)"),
                "Aging":     st.column_config.NumberColumn("Aging (days)", format="%d"),
                "due_date":  st.column_config.DateColumn("Due Date"),
                "Reason Category": st.column_config.SelectboxColumn(
                    "Reason Category",
                    options=[""] + REASON_CATEGORIES,
                    required=False,
                ),
                "Notes":           st.column_config.TextColumn("Notes", max_chars=500),
                "Action Owner":    st.column_config.TextColumn("Action Owner"),
                "Next Action Date":st.column_config.TextColumn("Next Action Date", help="YYYY-MM-DD"),
            },
            key="invoice_editor",
        )

        if st.button("💾 Save Changes", type="primary", key="save_inv"):
            saved, skipped = 0, 0
            for _, row in edited.iterrows():
                inv = str(row["invoice_number"])
                cat   = str(row.get("Reason Category", "") or "")
                notes = str(row.get("Notes", "") or "")
                owner = str(row.get("Action Owner", "") or "")
                ndate = str(row.get("Next Action Date", "") or "")
                if any([cat, notes, owner, ndate]):
                    try:
                        upsert_reason("invoice", inv, cat, notes, owner, ndate)
                        saved += 1
                    except Exception:
                        skipped += 1
            st.success(f"Saved {saved} invoice reason(s)." + (f" {skipped} skipped." if skipped else ""))
            st.rerun()

        st.download_button(
            "⬇ Download with Reasons",
            data=export_excel(edited),
            file_name="invoices_with_reasons.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ─────────────────────────── TAB 4 · REASONS & ACTIONS ───────────────────────
with tab_reasons:
    if not _can("view_reasons"):
        st.info("📝 Reasons & Actions is not available for your role. "
                "Please contact an Admin if you need access.")
    else:
        if not _can("edit_reasons"):
            st.info("👁 You have view-only access. Saving and deleting reasons requires Admin or Executor access.")
        rt1, rt2, rt3 = st.tabs(["🧾 Invoice Level", "🏢 Customer Level", "👤 CSM Level"])

        with rt1:
            st.markdown("Mark a reason / action for a specific **invoice**.")
            reason_form("invoice", df["invoice_number"].dropna().unique(), "Invoice")

        with rt2:
            st.markdown("Mark a reason / action for a specific **customer**.")
            reason_form("customer", df["customer_name"].dropna().unique(), "Customer", df_ref=fdf)

        with rt3:
            st.markdown("Mark a reason / action for a specific **CSM**.")
            reason_form("csm", df["CSM"].dropna().unique(), "CSM")

# ─────────────────────────── TAB 5 · SEND REMINDERS ─────────────────────────
with tab_email:
    if not _can("send_reminders"):
        st.info("📧 Sending reminders is not available for your role. "
                "Please contact an Admin or Executor if you need to send emails.")
    else:

        smtp_ready = all([SMTP_CFG.get("user"), SMTP_CFG.get("password")])
        if not smtp_ready:
            st.warning("⚙️ Gmail credentials not configured. Add `smtp_user` and `smtp_pass` to your Streamlit secrets.")

        # ── Column availability hint ───────────────────────────────────────────────
        _cc_detected = "Customer CC Email" in fdf.columns
        if _cc_detected:
            _cc_populated = fdf["Customer CC Email"].astype(str).str.strip().replace({"nan":"","None":""}).ne("").sum()
            st.success(f"✅ **Customer CC Email** column detected — {_cc_populated} invoice row(s) have a CC address.")
        else:
            # Show raw column names so user can spot the mismatch
            _raw_cols = [c for c in fdf.columns if "cc" in c.lower() or "customer" in c.lower()]
            _hint = f"Columns with 'customer'/'cc': `{', '.join(_raw_cols)}`" if _raw_cols else "No matching columns found."
            st.warning(f"⚠️ **Customer CC Email** column not found in sheet. {_hint}")

        # ── Template & recipient config (shared across both sub-tabs) ─────────────
        cfg1, cfg2 = st.columns([2, 3])
        with cfg1:
            selected_template = st.selectbox(
                "📋 Email Template",
                list(TEMPLATES.keys()),
                help="Choose the tone and content of your email",
            )
            template_key = TEMPLATES[selected_template]

        with cfg2:
            st.markdown("**Recipients**")
            rc1, rc2, rc3, rc4 = st.columns(4)
            send_to_customer = rc1.checkbox("✉️ Customer",         value=True, key="rc_customer")
            send_to_csm      = rc2.checkbox("👤 CSM",              value=True, key="rc_csm")
            send_to_finance  = rc3.checkbox("🏦 finance@spyne.ai", value=True, key="rc_finance")
            send_to_other    = rc4.checkbox("📨 Other",            value=False, key="rc_other")

            other_email = ""
            if send_to_other:
                other_email = st.text_input(
                    "Additional email address",
                    placeholder="e.g. manager@company.com",
                    key="rc_other_email",
                    label_visibility="collapsed",
                ).strip()
                if send_to_other and other_email and "@" not in other_email:
                    st.warning("⚠️ Enter a valid email address for the additional recipient.")

            if not any([send_to_customer, send_to_csm, send_to_finance,
                        (send_to_other and other_email and "@" in other_email)]):
                st.error("Select at least one recipient.")

        st.divider()
        et1, et2, et3 = st.tabs(["🏢 By Customer", "🧾 By Invoice", "📋 Sent Log"])

        def _build_cc(csm_email: str, customer_cc: str = "") -> list:
            """Build CC list based on recipient toggles + optional customer CC email(s)."""
            cc = []
            if send_to_csm and csm_email and "@" in csm_email:
                cc.append(csm_email)
            if send_to_finance:
                cc.append(FINANCE_CC)
            if send_to_other and other_email and "@" in other_email:
                cc.append(other_email)
            # Customer CC Email column — may contain multiple addresses separated by comma/semicolon
            for addr in re.split(r"[,;]+", str(customer_cc or "")):
                addr = addr.strip()
                if addr and "@" in addr and addr not in cc:
                    cc.append(addr)
            return cc

        # ═══════════════════════════════════════════════════════════════════════════
        # SUB-TAB A · CUSTOMER LEVEL (consolidated — one email per customer)
        # ═══════════════════════════════════════════════════════════════════════════
        with et1:
            st.markdown(f"Using template: **{selected_template}** · One consolidated email per customer with all outstanding invoices.")

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                c_csm = st.multiselect("CSM",    sorted(fdf["CSM"].dropna().unique())    if "CSM"    in fdf.columns else [], key="c_csm")
            with c2:
                c_rag = st.multiselect("RAG",    sorted(fdf["RAG"].dropna().unique())    if "RAG"    in fdf.columns else [], key="c_rag")
            with c3:
                c_bkt = st.multiselect("Bucket", [b for b in BUCKET_ORDER if b in fdf.get("Bucket", pd.Series()).values], key="c_bkt")
            with c4:
                _cis_opts = (sorted(fdf["Current Invoice Status"].astype(str).str.strip().dropna().unique())
                             if "Current Invoice Status" in fdf.columns else [])
                c_inv_status = st.multiselect("Invoice Status", _cis_opts, key="c_inv_status")

            cf = fdf.copy()
            if c_csm:       cf = cf[cf["CSM"].isin(c_csm)]
            if c_rag:       cf = cf[cf["RAG"].isin(c_rag)]
            if c_bkt:       cf = cf[cf["Bucket"].isin(c_bkt)]
            if c_inv_status and "Current Invoice Status" in cf.columns:
                cf = cf[cf["Current Invoice Status"].isin(c_inv_status)]
            if "email" in cf.columns:
                cf = cf[cf["email"].notna() & (cf["email"].str.strip() != "")]

            if "customer_name" in cf.columns:
                # Only include columns that actually exist in cf
                agg_dict = {}
                if "email"          in cf.columns: agg_dict["Email"]     = ("email",          "first")
                if "CSM"            in cf.columns: agg_dict["CSM"]       = ("CSM",            "first")
                if "RAG"            in cf.columns: agg_dict["RAG"]       = ("RAG",            "first")
                if "invoice_number" in cf.columns: agg_dict["Invoices"]  = ("invoice_number", "count")
                if "Aging"          in cf.columns: agg_dict["Max_Aging"] = ("Aging",          "max")
                if "CSM Email"      in cf.columns: agg_dict["CSM_Email"] = ("CSM Email",      "first")
                if "Customer CC Email" in cf.columns: agg_dict["Customer_CC"] = ("Customer CC Email", "first")

                cust_summary = cf.groupby("customer_name").agg(**agg_dict).reset_index()
                if "Max_Aging" in cust_summary.columns:
                    cust_summary = cust_summary.sort_values("Max_Aging", ascending=False)

                # Build per-customer currency-wise outstanding string
                # e.g. "₹4,838,819  |  $529,867"
                def fmt_cust_outstanding(cname):
                    rows = cf[cf["customer_name"] == cname]
                    if "balance" not in rows.columns or "currency_code" not in rows.columns:
                        return "—"
                    parts = (rows.groupby("currency_code")["balance"].sum()
                                 .reset_index()
                                 .sort_values("balance", ascending=False))
                    return "  |  ".join(
                        f"{CURR_SYM.get(str(r['currency_code']).upper(), '')}{r['balance']:,.0f}"
                        for _, r in parts.iterrows()
                    )

                cust_summary["Outstanding (FC)"] = cust_summary["customer_name"].apply(fmt_cust_outstanding)

                # Reorder columns
                col_order = ["customer_name","Email","CSM","RAG","Invoices",
                             "Outstanding (FC)","Max_Aging"]
                if "CSM_Email" in cust_summary.columns:
                    col_order.append("CSM_Email")
                if "Customer_CC" in cust_summary.columns:
                    col_order.append("Customer_CC")
                cust_summary = cust_summary[[c for c in col_order if c in cust_summary.columns]]

                # ── Upload a customer list to auto-select ─────────────────────────
                with st.expander("📂 Auto-select from uploaded list", expanded=False):
                    st.caption("Upload a CSV or Excel file with a column named **customer_name** "
                               "(or any column — first column is used as fallback).")
                    _upload_file = st.file_uploader(
                        "Upload customer list", type=["csv", "xlsx", "xls"],
                        key="cust_list_upload", label_visibility="collapsed"
                    )
                    _auto_selected_names: set = set()
                    if _upload_file is not None:
                        try:
                            if _upload_file.name.endswith(".csv"):
                                _ul_df = pd.read_csv(_upload_file)
                            else:
                                _ul_df = pd.read_excel(_upload_file)
                            # Find customer_name column (case-insensitive) or use first column
                            _cn_col = next(
                                (c for c in _ul_df.columns
                                 if c.strip().lower() in ("customer_name", "customer name",
                                                           "customername", "name")),
                                _ul_df.columns[0]
                            )
                            _auto_selected_names = set(
                                _ul_df[_cn_col].dropna().astype(str).str.strip().tolist()
                            )
                            st.success(f"✅ {len(_auto_selected_names)} customer name(s) loaded from **{_upload_file.name}** — "
                                       f"matching rows will be pre-selected below.")
                            # Show which names matched / didn't match
                            _all_cust = set(cust_summary["customer_name"].astype(str).tolist())
                            _matched   = _auto_selected_names & _all_cust
                            _unmatched = _auto_selected_names - _all_cust
                            st.caption(f"Matched: **{len(_matched)}**   |   Not found in current data: **{len(_unmatched)}**"
                                       + (f" — {sorted(_unmatched)}" if _unmatched else ""))
                        except Exception as _ul_err:
                            st.error(f"Could not read file: {_ul_err}")

                # Pre-tick Send? for uploaded names; default False for others
                cust_summary.insert(
                    0, "Send?",
                    cust_summary["customer_name"].astype(str).isin(_auto_selected_names)
                    if _auto_selected_names else False
                )

                edited_cust = st.data_editor(
                    cust_summary, use_container_width=True, height=360,
                    disabled=[c for c in cust_summary.columns if c != "Send?"],
                    column_config={
                        "Send?":            st.column_config.CheckboxColumn("Send?", default=False),
                        "Outstanding (FC)": st.column_config.TextColumn("Outstanding (FC)"),
                        "Customer_CC":      st.column_config.TextColumn("Customer CC Email"),
                        "Max_Aging":        st.column_config.NumberColumn("Max Aging (days)", format="%d"),
                        "Invoices":         st.column_config.NumberColumn("# Invoices", format="%d"),
                    },
                    key="cust_email_selector",
                )

                cust_to_send = edited_cust[edited_cust["Send?"] == True]
                st.caption(f"**{len(cust_to_send)}** customer(s) selected")

                st.divider()
                c_note = st.text_area("Additional note (optional)", height=70,
                                       placeholder="e.g. Please share UTR / transaction details upon payment.",
                                       key="c_note")

                if not cust_to_send.empty:
                    with st.expander("👁 Preview email for first selected customer"):
                        r0        = cust_to_send.iloc[0]
                        cname0    = r0["customer_name"]
                        _, phtml  = build_email(template_key, cname0,
                                                cf[cf["customer_name"]==cname0],
                                                r0.get("CSM",""), c_note)
                        st.components.v1.html(phtml, height=600, scrolling=True)

                st.divider()

                # ── Zoho toggles ───────────────────────────────────────────────────
                attach_pdfs_cust   = False
                fetch_plinks_cust  = False
                if ZOHO_READY:
                    zc1, zc2 = st.columns(2)
                    with zc1:
                        attach_pdfs_cust = st.checkbox(
                            "📎 Attach invoice PDFs",
                            value=False, key="attach_pdfs_cust",
                            help="Downloads each invoice PDF from Zoho Books and attaches it to the email.",
                        )
                    with zc2:
                        fetch_plinks_cust = st.checkbox(
                            "🔗 Add payment links (Pay Now button)",
                            value=True, key="fetch_plinks_cust",
                            help="Fetches the Zoho SecurePay link for each invoice and embeds a Pay Now button in the email.",
                        )
                else:
                    st.caption("📎 Zoho Books credentials not configured. Add Zoho keys to your Streamlit secrets to enable PDF attachments and payment links.")

                s1, s2 = st.columns([2, 3])
                with s1:
                    send_cust = st.button(
                        f"📤 Send to {len(cust_to_send)} Customer(s)",
                        type="primary",
                        disabled=(len(cust_to_send)==0 or not smtp_ready or
                                  not any([send_to_customer, send_to_csm, send_to_finance,
                                           (send_to_other and other_email and "@" in other_email)])),
                        use_container_width=True, key="send_cust_btn")
                with s2:
                    recip_summary = []
                    if send_to_customer: recip_summary.append("Customer")
                    if send_to_csm:      recip_summary.append("CSM")
                    if send_to_finance:  recip_summary.append("finance@spyne.ai")
                    if send_to_other and other_email and "@" in other_email:
                        recip_summary.append(other_email)
                    _has_cust_cc = "Customer CC Email" in cf.columns and cf["Customer CC Email"].astype(str).str.strip().replace({"nan":"","None":""}).ne("").any()
                    if _has_cust_cc:
                        recip_summary.append("Customer CC (from sheet)")
                    st.info("Sending to: " + " · ".join(recip_summary) if recip_summary else "No recipients selected")

                if send_cust and not cust_to_send.empty and smtp_ready:
                    # ── Get Zoho access token once if either Zoho feature is on ───
                    zoho_token = None
                    need_zoho  = (attach_pdfs_cust or fetch_plinks_cust) and ZOHO_READY
                    if need_zoho:
                        with st.spinner("🔑 Authenticating with Zoho Books…"):
                            try:
                                zoho_token = get_zoho_token(
                                    ZOHO_CFG["client_id"], ZOHO_CFG["client_secret"],
                                    ZOHO_CFG["refresh_token"], ZOHO_CFG["dc"],
                                )
                            except Exception as zt_err:
                                st.error(f"Zoho authentication failed: {zt_err}. Emails will be sent without PDFs/payment links.")

                    # ── Batch-fetch payment links only for SELECTED customers ──────
                    cf_send = cf.copy()
                    if zoho_token and fetch_plinks_cust and "invoice_number" in cf_send.columns:
                        selected_customers = cust_to_send["customer_name"].tolist()
                        all_inv_nos = (
                            cf_send[cf_send["customer_name"].isin(selected_customers)]
                            ["invoice_number"].dropna().astype(str).unique().tolist()
                        )
                        with st.spinner(f"🔗 Fetching payment links for {len(all_inv_nos)} invoice(s) across {len(selected_customers)} customer(s)…"):
                            try:
                                pay_links_map, pay_link_errs = fetch_zoho_payment_links(
                                    all_inv_nos, zoho_token,
                                    ZOHO_CFG["org_ids"], ZOHO_CFG["dc"],
                                )
                                cf_send["payment_link"] = cf_send["invoice_number"].astype(str).map(pay_links_map)
                                st.caption(f"🔗 Payment links found: {len(pay_links_map)} / {len(all_inv_nos)}")
                                if pay_link_errs:
                                    with st.expander(f"⚠️ {len(pay_link_errs)} invoice(s) had issues fetching payment link"):
                                        for inv, err in pay_link_errs.items():
                                            st.text(f"{inv}: {err}")
                            except Exception as pl_err:
                                st.warning(f"Payment links fetch failed: {pl_err}")

                    prog = st.progress(0, text="Sending…")
                    ok, fail, results = 0, 0, []
                    for idx, (_, crow) in enumerate(cust_to_send.iterrows()):
                        cname        = crow["customer_name"]
                        to_email     = str(crow.get("Email","")).strip() if send_to_customer else None
                        csm_email    = str(crow.get("CSM_Email", crow.get("CSM",""))).strip()
                        customer_cc  = str(crow.get("Customer_CC", "")).strip()
                        cc_list      = _build_cc(csm_email, customer_cc)
                        cust_invs = cf_send[cf_send["customer_name"]==cname]
                        subject, html = build_email(template_key, cname, cust_invs,
                                                    crow.get("CSM",""), c_note)

                        # ── Fetch PDFs for all invoices of this customer ───────────
                        attachments = []
                        pdf_notes   = []
                        if zoho_token and attach_pdfs_cust and "invoice_number" in cust_invs.columns:
                            for inv_no in cust_invs["invoice_number"].dropna().unique():
                                inv_no = str(inv_no).strip()
                                try:
                                    pdf_bytes, _, _org_used, _plink = fetch_zoho_invoice_pdf(
                                        inv_no, zoho_token,
                                        ZOHO_CFG["org_ids"], ZOHO_CFG["dc"],
                                    )
                                    if pdf_bytes:
                                        attachments.append((f"{inv_no}.pdf", pdf_bytes))
                                        pdf_notes.append(f"✅ {inv_no}")
                                    else:
                                        pdf_notes.append(f"⚠️ {inv_no} (not in Zoho)")
                                except Exception as pdf_err:
                                    pdf_notes.append(f"❌ {inv_no} ({pdf_err})")

                        # Determine actual TO
                        if to_email and "@" in to_email:
                            actual_to, actual_cc = to_email, cc_list
                        elif cc_list:
                            actual_to, actual_cc = cc_list[0], cc_list[1:]
                        else:
                            results.append({"Customer": cname, "Status": "⚠️ No recipients",
                                            "PDFs": ""})
                            continue
                        try:
                            send_reminder(SMTP_CFG, actual_to, actual_cc, subject, html,
                                          attachments=attachments or None)
                            inv_nos_sent = list(cust_invs["invoice_number"].dropna().astype(str).unique()) if "invoice_number" in cust_invs.columns else ["(consolidated)"]
                            log_email(inv_nos_sent, cname, actual_to,
                                      ", ".join(actual_cc), subject, "sent",
                                      template=selected_template)
                            ok += 1
                            status = f"✅ Sent" + (f" ({len(attachments)} PDF(s))" if attachments else "")
                            results.append({"Customer": cname, "To": actual_to,
                                            "CC": ", ".join(actual_cc),
                                            "Status": status,
                                            "PDFs": " | ".join(pdf_notes) if pdf_notes else "—"})
                        except Exception as e:
                            inv_nos_sent = list(cust_invs["invoice_number"].dropna().astype(str).unique()) if "invoice_number" in cust_invs.columns else ["(consolidated)"]
                            log_email(inv_nos_sent, cname, actual_to,
                                      ", ".join(actual_cc), subject, "failed", str(e),
                                      template=selected_template)
                            fail += 1
                            results.append({"Customer": cname, "To": actual_to,
                                            "Status": f"❌ {e}",
                                            "PDFs": " | ".join(pdf_notes) if pdf_notes else "—"})
                        prog.progress((idx+1)/len(cust_to_send), text=f"Sent {idx+1} of {len(cust_to_send)}…")
                    prog.empty()
                    if ok:   st.success(f"✅ {ok} reminder(s) sent.")
                    if fail: st.error(f"❌ {fail} failed.")
                    st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)
            else:
                st.info("customer_name column not detected.")

        # ═══════════════════════════════════════════════════════════════════════════
        # SUB-TAB B · INVOICE LEVEL (one email per invoice)
        # ═══════════════════════════════════════════════════════════════════════════
        with et2:
            st.markdown(f"Using template: **{selected_template}** · One email per invoice.")

            i1, i2, i3 = st.columns(3)
            with i1:
                e_csm = st.multiselect("CSM",    sorted(fdf["CSM"].dropna().unique())    if "CSM"    in fdf.columns else [], key="e_csm")
            with i2:
                e_rag = st.multiselect("RAG",    sorted(fdf["RAG"].dropna().unique())    if "RAG"    in fdf.columns else [], key="e_rag")
            with i3:
                e_bkt = st.multiselect("Bucket", [b for b in BUCKET_ORDER if b in fdf.get("Bucket", pd.Series()).values], key="e_bkt")

            ef = fdf.copy()
            if e_csm: ef = ef[ef["CSM"].isin(e_csm)]
            if e_rag: ef = ef[ef["RAG"].isin(e_rag)]
            if e_bkt: ef = ef[ef["Bucket"].isin(e_bkt)]
            if "email" in ef.columns:
                ef = ef[ef["email"].notna() & (ef["email"].str.strip() != "")]

            st.caption(f"{len(ef):,} invoices with valid email addresses")

            sel_cols = [c for c in ["invoice_number","customer_name","email","CSM",
                                     "CSM Email","Final USD","Aging","Bucket","RAG"] if c in ef.columns]
            sel_df = ef[sel_cols].copy()
            if "Aging" in sel_df.columns:
                sel_df = sel_df.sort_values("Aging", ascending=False)
            sel_df.insert(0, "Send?", False)

            edited_sel = st.data_editor(
                sel_df, use_container_width=True, height=340,
                disabled=[c for c in sel_df.columns if c != "Send?"],
                column_config={
                    "Send?":     st.column_config.CheckboxColumn("Send?", default=False),
                    "Final USD": st.column_config.NumberColumn("Final USD", format="$%.0f"),
                    "Aging":     st.column_config.NumberColumn("Aging (days)", format="%d"),
                },
                key="email_selector",
            )

            to_send = edited_sel[edited_sel["Send?"] == True]
            st.caption(f"**{len(to_send)}** invoice(s) selected")

            st.divider()
            custom_note = st.text_area("Additional note (optional)", height=70,
                                        placeholder="e.g. Please note our bank details have changed.",
                                        key="inv_note")

            if not to_send.empty:
                with st.expander("👁 Preview first email"):
                    r0       = to_send.iloc[0]
                    cname0   = str(r0.get("customer_name",""))
                    _, phtml = build_email(template_key, cname0,
                                           pd.DataFrame([r0.to_dict()]),
                                           str(r0.get("CSM","")), custom_note)
                    st.components.v1.html(phtml, height=560, scrolling=True)

            st.divider()

            # ── Zoho toggles ───────────────────────────────────────────────────────
            attach_pdfs_inv  = False
            fetch_plinks_inv = False
            if ZOHO_READY:
                zi1, zi2 = st.columns(2)
                with zi1:
                    attach_pdfs_inv = st.checkbox(
                        "📎 Attach invoice PDF",
                        value=False, key="attach_pdfs_inv",
                        help="Downloads the invoice PDF from Zoho Books and attaches it.",
                    )
                with zi2:
                    fetch_plinks_inv = st.checkbox(
                        "🔗 Add payment links (Pay Now button)",
                        value=True, key="fetch_plinks_inv",
                        help="Fetches the Zoho SecurePay link and embeds a Pay Now button in the email.",
                    )
            else:
                st.caption("📎 Zoho Books credentials not configured. Add Zoho keys to your Streamlit secrets to enable PDF attachments and payment links.")

            b1, b2 = st.columns([2, 3])
            with b1:
                _unique_custs = to_send["customer_name"].nunique() if "customer_name" in to_send.columns else len(to_send)
                send_clicked = st.button(
                    f"📤 Send {_unique_custs} Reminder(s)  ({len(to_send)} invoice(s))",
                    type="primary",
                    disabled=(len(to_send)==0 or not smtp_ready or
                              not any([send_to_customer, send_to_csm, send_to_finance,
                                       (send_to_other and other_email and "@" in other_email)])),
                    use_container_width=True, key="send_inv_btn")
            with b2:
                recip_inv = []
                if send_to_customer: recip_inv.append("Customer")
                if send_to_csm:      recip_inv.append("CSM")
                if send_to_finance:  recip_inv.append("finance@spyne.ai")
                if send_to_other and other_email and "@" in other_email:
                    recip_inv.append(other_email)
                _has_cust_cc_inv = "Customer CC Email" in ef.columns and ef["Customer CC Email"].astype(str).str.strip().replace({"nan":"","None":""}).ne("").any()
                if _has_cust_cc_inv:
                    recip_inv.append("Customer CC (from sheet)")
                st.info("Sending to: " + " · ".join(recip_inv) if recip_inv else "No recipients selected")

            if send_clicked and not to_send.empty and smtp_ready:
                # ── Get Zoho access token once if either Zoho feature is on ──────
                zoho_token_inv = None
                need_zoho_inv  = (attach_pdfs_inv or fetch_plinks_inv) and ZOHO_READY
                if need_zoho_inv:
                    with st.spinner("🔑 Authenticating with Zoho Books…"):
                        try:
                            zoho_token_inv = get_zoho_token(
                                ZOHO_CFG["client_id"], ZOHO_CFG["client_secret"],
                                ZOHO_CFG["refresh_token"], ZOHO_CFG["dc"],
                            )
                        except Exception as zt_err:
                            st.error(f"Zoho authentication failed: {zt_err}. Emails will be sent without PDFs/payment links.")

                # ── Batch-fetch ALL payment links upfront (one pass) ──────────────
                ef_send = ef.copy()
                if zoho_token_inv and fetch_plinks_inv and "invoice_number" in ef_send.columns:
                    sel_inv_nos = to_send["invoice_number"].dropna().astype(str).unique().tolist()
                    with st.spinner(f"🔗 Fetching payment links for {len(sel_inv_nos)} invoice(s)…"):
                        try:
                            pay_links_map_inv, pay_link_errs_inv = fetch_zoho_payment_links(
                                sel_inv_nos, zoho_token_inv,
                                ZOHO_CFG["org_ids"], ZOHO_CFG["dc"],
                            )
                            ef_send["payment_link"] = ef_send["invoice_number"].astype(str).map(pay_links_map_inv)
                            st.caption(f"🔗 Payment links found: {len(pay_links_map_inv)} / {len(sel_inv_nos)}")
                            if pay_link_errs_inv:
                                with st.expander(f"⚠️ {len(pay_link_errs_inv)} invoice(s) had issues fetching payment link"):
                                    for inv, err in pay_link_errs_inv.items():
                                        st.text(f"{inv}: {err}")
                        except Exception as pl_err:
                            st.warning(f"Payment links fetch failed: {pl_err}")

                # ── Group selected invoices by customer → 1 email per customer ──────
                _cust_groups = {}
                for _, row in to_send.iterrows():
                    cname = str(row.get("customer_name","")).strip() or "(unknown)"
                    _cust_groups.setdefault(cname, []).append(row)

                progress = st.progress(0, text="Sending…")
                ok2, fail2, results2 = 0, 0, []

                for idx, (customer, rows) in enumerate(_cust_groups.items()):
                    inv_nos = [str(r.get("invoice_number","")).strip() for r in rows]
                    first   = rows[0]   # use first row for contact details

                    # Build enriched invoice DataFrame for this customer
                    cust_inv_df = pd.concat(
                        [ef_send[ef_send["invoice_number"].astype(str) == n]
                         for n in inv_nos], ignore_index=True
                    )
                    if cust_inv_df.empty:
                        cust_inv_df = pd.DataFrame([r.to_dict() for r in rows])

                    to_email    = str(first.get("email","")).strip() if send_to_customer else None
                    csm_email   = str(first.get("CSM Email","")).strip()
                    customer_cc = str(first.get("Customer CC Email","")).strip()
                    cc_list     = _build_cc(csm_email, customer_cc)
                    subject, html = build_email(template_key, customer,
                                                cust_inv_df,
                                                str(first.get("CSM","")), custom_note)

                    # ── Fetch PDFs for all invoices of this customer ───────────────
                    attachments_inv = []
                    pdf_notes_inv   = []
                    if zoho_token_inv and attach_pdfs_inv:
                        for inv_no in inv_nos:
                            if not inv_no: continue
                            try:
                                pdf_bytes, _, _org_used, _plink = fetch_zoho_invoice_pdf(
                                    inv_no, zoho_token_inv,
                                    ZOHO_CFG["org_ids"], ZOHO_CFG["dc"],
                                )
                                if pdf_bytes:
                                    attachments_inv.append((f"{inv_no}.pdf", pdf_bytes))
                                    pdf_notes_inv.append(f"✅ {inv_no}")
                                else:
                                    pdf_notes_inv.append(f"⚠️ {inv_no} not found")
                            except Exception as pdf_err:
                                pdf_notes_inv.append(f"❌ {inv_no}: {pdf_err}")
                    pdf_note_str = " | ".join(pdf_notes_inv) if pdf_notes_inv else "—"

                    if to_email and "@" in to_email:
                        actual_to, actual_cc = to_email, cc_list
                    elif cc_list:
                        actual_to, actual_cc = cc_list[0], cc_list[1:]
                    else:
                        results2.append({"Customer": customer,
                                         "Invoices": ", ".join(inv_nos),
                                         "Status": "⚠️ No recipients", "PDFs": pdf_note_str})
                        continue

                    try:
                        send_reminder(SMTP_CFG, actual_to, actual_cc, subject, html,
                                      attachments=attachments_inv or None)
                        log_email(inv_nos, customer, actual_to, ", ".join(actual_cc),
                                  subject, "sent", template=selected_template)
                        ok2 += 1
                        status2 = "✅ Sent" + (f" ({len(attachments_inv)} PDF(s))" if attachments_inv else "")
                        results2.append({"Customer": customer,
                                         "Invoices": ", ".join(inv_nos),
                                         "To": actual_to, "Status": status2, "PDFs": pdf_note_str})
                    except Exception as e:
                        log_email(inv_nos, customer, actual_to, ", ".join(actual_cc),
                                  subject, "failed", str(e), template=selected_template)
                        fail2 += 1
                        results2.append({"Customer": customer,
                                         "Invoices": ", ".join(inv_nos),
                                         "To": actual_to, "Status": f"❌ {e}", "PDFs": pdf_note_str})

                    progress.progress((idx+1)/len(_cust_groups),
                                      text=f"Sent {idx+1} of {len(_cust_groups)} customer(s)…")
                progress.empty()
                if ok2:   st.success(f"✅ {ok2} email(s) sent to {ok2} customer(s).")
                if fail2: st.error(f"❌ {fail2} failed.")
                st.dataframe(pd.DataFrame(results2), use_container_width=True, hide_index=True)

        # ═══════════════════════════════════════════════════════════════════════════
        # SUB-TAB C · SENT LOG
        # ═══════════════════════════════════════════════════════════════════════════
        with et3:
            st.subheader("📋 Sent Email Log")
            log_df = get_sent_log()
            if not log_df.empty:
                # ── Format sent_at as readable IST string ──────────────────────────
                from datetime import timezone, timedelta
                _IST = timezone(timedelta(hours=5, minutes=30))

                def _fmt_ist(ts_str):
                    try:
                        dt = datetime.fromisoformat(str(ts_str).strip())
                        if dt.tzinfo is None:
                            dt = dt.replace(tzinfo=_IST)
                        else:
                            dt = dt.astimezone(_IST)
                        return dt.strftime("%d %b %Y, %I:%M %p IST")
                    except Exception:
                        return str(ts_str)[:19]

                log_df["sent_at_fmt"] = log_df["sent_at"].apply(_fmt_ist)

                # ── Build display dataframe ────────────────────────────────────────
                cols_available = [c for c in ["sent_at_fmt","invoice_no","customer","to_email",
                                               "cc_emails","subject","template","status","error"]
                                  if c in log_df.columns or c == "sent_at_fmt"]
                # template column may not exist in old rows — fill blank
                if "template" not in log_df.columns:
                    log_df["template"] = ""

                display_log = log_df[["sent_at_fmt","invoice_no","customer","to_email",
                                       "cc_emails","subject","template","status","error"]].copy()
                display_log.columns = ["Sent At (IST)","Invoice","Customer","To","CC",
                                        "Subject","Template","Status","Error"]

                # ── Summary KPIs ──────────────────────────────────────────────────
                total_sent = (display_log["Status"] == "sent").sum()
                total_failed = (display_log["Status"] == "failed").sum()
                unique_customers = display_log["Customer"].nunique()
                unique_invoices = display_log["Invoice"].nunique()

                lk1, lk2, lk3, lk4 = st.columns(4)
                lk1.metric("Total Reminders Sent", total_sent)
                lk2.metric("Failed", total_failed)
                lk3.metric("Unique Customers", unique_customers)
                lk4.metric("Unique Invoices", unique_invoices)

                st.divider()

                # ── Optional search filter ────────────────────────────────────────
                _log_search = st.text_input("🔍 Search log (customer, invoice, subject…)", key="log_search")
                if _log_search.strip():
                    _mask = display_log.apply(
                        lambda row: row.astype(str).str.contains(_log_search.strip(), case=False).any(),
                        axis=1
                    )
                    display_log = display_log[_mask]

                st.dataframe(display_log, use_container_width=True, height=450, hide_index=True)

                st.download_button(
                    "⬇ Download Log",
                    data=export_excel(display_log),
                    file_name="sent_email_log.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("No emails sent yet.")
