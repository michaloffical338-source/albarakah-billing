# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# + Admin + Booker separate login
# + Booker: Bill → admin Bills List (no download)
# + Booker: Load Form → admin Load Form list
# ============================================================

import os
import json
import hashlib
import pandas as pd
from datetime import datetime, date, timedelta
import streamlit as st
import streamlit.components.v1 as components
import xlsxwriter
from io import BytesIO

st.set_page_config(
    page_title="AL-BARAKAH ENTERPRISES",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# AUTH HELPERS
# ============================================================
USERS_FILE = "users.json"
BOOKERS_FILE = "bookers_registry.json"
PASSWORD_SALT = "albarakah_2026_secret_salt"

def hash_password(password):
    return hashlib.sha256((PASSWORD_SALT + password).encode()).hexdigest()

def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except Exception: pass
    return {}

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=4, ensure_ascii=False)

def load_bookers():
    if os.path.exists(BOOKERS_FILE):
        try:
            with open(BOOKERS_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except Exception: pass
    return {}

def save_bookers(bookers):
    with open(BOOKERS_FILE, "w", encoding="utf-8") as f:
        json.dump(bookers, f, indent=4, ensure_ascii=False)

def sanitize_username(u):
    return "".join(ch for ch in u if ch.isalnum() or ch in "_-.").lower()

def user_data_file(username):
    return f"billing_database_{username}.json"

def default_blank_db():
    return {
        "next_bill_no": 1, "bills": [], "bookers": [], "salesmen": [],
        "load_forms": [], "bookers_salaries": {}, "salesmen_salaries": {},
        "product_prices": {}, "petrol_expenses": [], "lunch_expenses": [],
        "discount_packages": default_discount_packages(),
        "custom_products": [], "dsr_forms": [], "credit_bills": [],
        "next_credit_id": 1,
        "wholesaler_rule_enabled": True,
        "wholesaler_discount_pct": 6.0,
    }

def default_discount_packages():
    return [
        {"id": 1, "name": "Package 1", "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0, "tier3_amount": 3100.0, "tier3_pct": 3.0, "active": True},
        {"id": 2, "name": "Package 2", "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0, "tier3_amount": 3100.0, "tier3_pct": 3.0, "active": True},
        {"id": 3, "name": "Package 3", "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0, "tier3_amount": 3100.0, "tier3_pct": 3.0, "active": True},
        {"id": 4, "name": "Package 4", "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0, "tier3_amount": 3100.0, "tier3_pct": 3.0, "active": True},
    ]

# ============================================================
# GLOBAL CSS
# ============================================================
st.markdown("""
<style>
    html, body, .stApp, .stApp p, .stApp span, .stApp div,
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
    .stApp label, .stApp li, .stApp a, [class*="css"] *,
    [data-testid="stMarkdownContainer"] *,
    [data-testid="stText"], [data-testid="stCaptionContainer"] *,
    [data-testid="stWidgetLabel"] *, [data-testid="stSelectbox"] *,
    [data-testid="stTextInput"] *, [data-testid="stNumberInput"] * { color: #000000; }

    .stDateInput, .stDateInput > div, .stDateInput > div > div,
    .stDateInput > div > div > input,
    [data-testid="stDateInput"], [data-testid="stDateInput"] > div,
    [data-testid="stDateInput"] > div > div, [data-testid="stDateInput"] input,
    [data-testid="stDateInput"] div[data-baseweb="input"],
    [data-testid="stDateInput"] div[data-baseweb="input"] > div,
    [data-testid="stDateInput"] div[data-baseweb="base-input"],
    [data-baseweb="datepicker"] input, div[data-baseweb="input"] > div,
    div[data-baseweb="base-input"], div[data-baseweb="base-input"] input {
        background-color: #ffffff !important; color: #000000 !important;
        -webkit-text-fill-color: #000000 !important; border-color: #90caf9 !important;
    }
    div[data-baseweb="calendar"], div[data-baseweb="calendar"] *,
    div[data-baseweb="datepicker"] *, div[data-baseweb="popover"] *,
    div[role="dialog"] *, div[role="dialog"] {
        background-color: #ffffff !important; color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }
    [data-testid="stDateInput"] svg, div[data-baseweb="datepicker"] svg { fill: #1976d2 !important; color: #1976d2 !important; }
    .stButton > button, .stButton > button p, .stButton > button span, .stButton > button div,
    .stDownloadButton > button, .stDownloadButton > button p,
    .stDownloadButton > button span, .stDownloadButton > button div { color: #ffffff !important; }

    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    [data-testid="stStatusWidget"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    header[data-testid="stHeader"] { background: transparent !important; box-shadow: none !important; }

    .stApp { background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%) !important; }
    .block-container {
        padding-top: 0.5rem !important; padding-left: 1.5rem !important;
        padding-right: 1.5rem !important; padding-bottom: 0.5rem !important;
        max-width: 100% !important;
    }
    .stApp .element-container { margin-bottom: 0.35rem !important; }
    .stApp [data-testid="stVerticalBlock"] > div { gap: 0.35rem !important; }
    .stApp label, .stApp [data-testid="stWidgetLabel"] label, .stApp [data-testid="stWidgetLabel"] p {
        font-size: 13px !important; font-weight: 600 !important; margin-bottom: 2px !important;
    }
    .stTextInput > div > div > input, .stNumberInput > div > div > input,
    .stDateInput > div > div > input, .stSelectbox > div > div > div,
    div[data-baseweb="input"] input, div[data-baseweb="select"] > div {
        padding: 4px 8px !important; min-height: 34px !important; font-size: 14px !important;
    }
    .stButton > button { padding: 4px 10px !important; min-height: 36px !important; font-size: 14px !important; }
    .stApp hr { margin: 6px 0 !important; }
    .stApp h3 { margin-top: 6px !important; margin-bottom: 4px !important; font-size: 18px !important; }
    .stTextInput > div > div > input, .stNumberInput > div > div > input,
    .stSelectbox > div > div > div, .stSelectbox > div > div,
    div[data-baseweb="select"] > div, div[data-baseweb="input"] input {
        background-color: #ffffff !important; color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        border: 2px solid #90caf9 !important; border-radius: 8px !important;
    }
    div[data-baseweb="select"] * { color: #000000 !important; }
    ul[role="listbox"] li, div[role="option"] { color: #000000 !important; background-color: #ffffff !important; }
    input[type="password"] { background-color: #ffffff !important; color: #000000 !important; -webkit-text-fill-color: #000000 !important; }

    .stButton > button {
        background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%) !important;
        border: none !important; font-weight: bold !important; border-radius: 8px !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #0288d1 0%, #0277bd 100%) !important;
        color: #ffffff !important; font-weight: bold !important; border-radius: 8px !important;
    }
    .stDataFrame, .stDataFrame * { color: #000000 !important; }
    .stDataFrame { background-color: #ffffff !important; border-radius: 10px !important; }
    label[data-baseweb="checkbox"] * { color: #000000 !important; }
    details summary, details summary * { color: #000000 !important; }
    .stAlert, .stAlert * { color: #000000 !important; }
    .stSuccess, .stSuccess * { color: #1b5e20 !important; }
    .stError, .stError * { color: #b71c1c !important; }
    .stWarning, .stWarning * { color: #e65100 !important; }
    .stInfo, .stInfo * { color: #0d47a1 !important; }
    [data-testid="stCaptionContainer"] * { color: #555555 !important; }

    .auto-dl-hidden div[data-testid="stDownloadButton"] {
        position: absolute !important; left: -9999px !important; top: -9999px !important;
        opacity: 0 !important; height: 0 !important;
    }
    .metric-card { background: #ffffff; border: 2px solid #90caf9; border-radius: 14px; padding: 22px; text-align: center; box-shadow: 0 3px 10px rgba(33,150,243,0.15); }
    .metric-card h3 { font-size: 13px !important; margin: 0 !important; font-weight: 700 !important; text-transform: uppercase; color: #0277bd !important; }
    .metric-card h1 { font-size: 30px !important; margin: 10px 0 0 0 !important; font-weight: 800 !important; color: #1976d2 !important; }
    .booker-row { background: #ffffff; border: 1px solid #90caf9; border-radius: 10px; padding: 12px 18px; margin-bottom: 8px; }
    .empty-box { background: #ffffff; border: 2px dashed #90caf9; border-radius: 12px; padding: 20px; text-align: center; font-size: 14px; color: #0277bd !important; }
    .hint-box { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 6px 10px; border-radius: 6px; font-size: 12px; margin-top: 2px; color: #1976d2 !important; }
    .summary-box { background: #ffffff; border: 2px solid #2196f3; border-radius: 12px; padding: 12px 18px; margin-bottom: 12px; color: #000000 !important; }
    .lf-simple-card {
        background: #ffffff; border-left: 6px solid #2196f3; border-radius: 12px;
        padding: 16px 22px; margin-bottom: 12px;
        box-shadow: 0 3px 10px rgba(33,150,243,0.15);
        display: flex; align-items: center; justify-content: space-between;
    }
    .lf-simple-card .lf-info { display: flex; flex-direction: column; gap: 4px; }
    .lf-simple-card .lf-line1 { font-size: 17px; font-weight: 700; color: #1976d2 !important; }
    .lf-simple-card .lf-line2 { font-size: 13px; color: #0277bd !important; }
    .lf-simple-card .lf-boxes {
        background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
        font-weight: 800; font-size: 20px; padding: 10px 18px; border-radius: 10px;
        text-align: center; min-width: 90px; color: #ffffff !important;
    }
    .lf-simple-card .lf-boxes small { display: block; font-size: 10px; font-weight: 500; opacity: 0.9; }
    .lf-simple-card.dsr { border-left-color: #e65100; }
    .lf-simple-card.dsr .lf-boxes { background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%); }
    .lf-simple-card.credit-pending { border-left-color: #e65100; }
    .lf-simple-card.credit-pending .lf-boxes { background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%); }
    .lf-simple-card.credit-paid { border-left-color: #2e7d32; background: #f1f8e9; }
    .lf-simple-card.credit-paid .lf-boxes { background: linear-gradient(135deg, #2e7d32 0%, #1b5e20 100%); }

    .sal-metric { display: inline-block; padding: 8px 14px; margin-right: 8px; margin-bottom: 6px; border-radius: 8px; font-size: 13px; font-weight: 600; }
    .sal-metric.base { background: #e3f2fd; color: #0d47a1 !important; }
    .sal-metric.adv { background: #fff3e0; color: #e65100 !important; }
    .sal-metric.short { background: #ffebee; color: #c62828 !important; }
    .sal-metric.remain { background: #e8f5e9; color: #1b5e20 !important; }
    .sal-metric.paid { background: #c8e6c9; color: #2e7d32 !important; }

    .status-pending { background: #fff3e0; color: #e65100 !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; }
    .status-paid { background: #c8e6c9; color: #1b5e20 !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; }
    .badge-pending { background: #ffe0b2; color: #e65100 !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; margin-left: 8px; }
    .badge-refreshed { background: #e1bee7; color: #6a1b9a !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; margin-left: 8px; }
    .badge-paid-credit { background: #c8e6c9; color: #1b5e20 !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; margin-left: 8px; }
    .badge-credit-tag { background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%); color: #ffffff !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; margin-left: 8px; }

    .auth-title { text-align: center; font-size: 36px; font-weight: 900; color: #1976d2; margin-bottom: 6px; margin-top: 20px; }
    .auth-subtitle { text-align: center; font-size: 14px; color: #0277bd; margin-bottom: 24px; font-weight: 600; }
    .auth-card { background: #ffffff; border: 3px solid #90caf9; border-radius: 16px; padding: 28px 32px; box-shadow: 0 8px 24px rgba(33,150,243,0.2); margin: 0 auto; }
    .stAlert { border-radius: 10px !important; }
    hr { border-color: #90caf9 !important; opacity: 0.6 !important; }

    .bk-product-card {
        background: #ffffff; border: 1.5px solid #90caf9; border-radius: 10px;
        padding: 8px 12px; margin-bottom: 6px;
        display: flex; align-items: center; justify-content: space-between;
        transition: all 0.15s ease;
    }
    .bk-product-card:hover { border-color: #2196f3; box-shadow: 0 2px 8px rgba(33,150,243,0.15); }
    .bk-product-card .bk-name { font-size: 13.5px; font-weight: 700; color: #1976d2; }
    .bk-product-card .bk-code { font-size: 11px; color: #666; }
    .bk-product-card .bk-tp {
        background: linear-gradient(135deg, #4caf50 0%, #2e7d32 100%);
        color: #fff !important; font-weight: 800; font-size: 13px;
        padding: 4px 12px; border-radius: 6px;
    }

    /* SIDEBAR (light green) */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f1f8e9 0%, #e8f5e9 50%, #dcedc8 100%) !important;
        border-right: 2px solid #a5d6a7 !important;
        box-shadow: 4px 0 20px rgba(76,175,80,0.12) !important;
    }
    section[data-testid="stSidebar"] > div:first-child { padding-top: 0 !important; }
    section[data-testid="stSidebar"] .block-container { padding: 0 !important; }
    section[data-testid="stSidebar"] ::-webkit-scrollbar { width: 6px; }
    section[data-testid="stSidebar"] ::-webkit-scrollbar-thumb { background: rgba(76,175,80,0.4); border-radius: 3px; }

    .sb-brand { text-align: center; padding: 22px 16px 18px; border-bottom: 1px solid rgba(76,175,80,0.20); position: relative; background: radial-gradient(circle at 50% 0%, rgba(76,175,80,0.15) 0%, transparent 70%); }
    .sb-brand-logo { display: inline-flex; align-items: center; justify-content: center; width: 56px; height: 56px; border-radius: 16px; background: linear-gradient(135deg, #66bb6a 0%, #2e7d32 100%); font-size: 28px; margin-bottom: 10px; box-shadow: 0 8px 20px rgba(76,175,80,0.35), inset 0 1px 0 rgba(255,255,255,0.35); }
    .sb-brand-name { font-size: 17px; font-weight: 800; color: #1b5e20 !important; letter-spacing: 2px; margin: 0; text-shadow: 0 1px 2px rgba(255,255,255,0.8); }
    .sb-brand-sub { font-size: 9px; color: #558b2f !important; letter-spacing: 4px; font-weight: 700; margin-top: 3px; }
    .sb-brand-dot { display: inline-block; width: 6px; height: 6px; background: #43a047; border-radius: 50%; margin-right: 4px; box-shadow: 0 0 8px #43a047; vertical-align: middle; }
    .sb-brand-status { font-size: 9px; color: #689f38 !important; letter-spacing: 1px; margin-top: 6px; font-weight: 700; }
    .sb-user { display: flex; align-items: center; gap: 12px; background: #ffffff; border: 1px solid rgba(76,175,80,0.25); border-radius: 12px; padding: 10px 12px; margin: 14px 14px 16px; position: relative; overflow: hidden; box-shadow: 0 2px 8px rgba(76,175,80,0.10); }
    .sb-user::before { content: ''; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: linear-gradient(180deg, #66bb6a 0%, #2e7d32 100%); }
    .sb-user-avatar { width: 40px; height: 40px; border-radius: 11px; background: linear-gradient(135deg, #66bb6a 0%, #2e7d32 100%); display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 16px; color: #ffffff !important; flex-shrink: 0; box-shadow: 0 4px 10px rgba(76,175,80,0.35), inset 0 1px 0 rgba(255,255,255,0.3); text-transform: uppercase; }
    .sb-user-info { display: flex; flex-direction: column; min-width: 0; }
    .sb-user-label { font-size: 9px; color: #689f38 !important; text-transform: uppercase; letter-spacing: 1.5px; font-weight: 800; margin-bottom: 2px; }
    .sb-user-name { font-size: 13px; font-weight: 800; color: #1b5e20 !important; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .sb-section-label { font-size: 9px; color: #558b2f !important; letter-spacing: 2.5px; font-weight: 800; padding: 0 20px 8px; text-transform: uppercase; }

    section[data-testid="stSidebar"] div[role="radiogroup"] { gap: 2px !important; padding: 0 10px 8px !important; display: flex !important; flex-direction: column !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label { background-color: #ffffff !important; border: 1px solid rgba(76,175,80,0.18) !important; border-radius: 10px !important; margin-bottom: 2px !important; padding: 9px 12px !important; transition: all 0.18s ease !important; cursor: pointer !important; display: flex !important; align-items: center !important; width: 100% !important; position: relative !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover { background-color: #f1f8e9 !important; border-color: #66bb6a !important; transform: translateX(2px); box-shadow: 0 2px 8px rgba(76,175,80,0.15); }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label p { color: #1b5e20 !important; font-size: 13.5px !important; font-weight: 600 !important; margin: 0 !important; letter-spacing: 0.2px !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover p { color: #1b5e20 !important; font-weight: 700 !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child { display: none !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label > div[data-testid="stMarkdownContainer"] { width: 100% !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) { background: linear-gradient(135deg, #66bb6a 0%, #2e7d32 100%) !important; border-color: #2e7d32 !important; box-shadow: 0 6px 16px rgba(76,175,80,0.4), inset 0 1px 0 rgba(255,255,255,0.25) !important; transform: translateX(0) !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p { color: #ffffff !important; font-weight: 700 !important; }
    section[data-testid="stSidebar"] hr { border-color: rgba(76,175,80,0.25) !important; margin: 10px 14px !important; opacity: 1 !important; }

    .sb-stats { padding: 4px 14px 12px; }
    .sb-stats-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 10px; }
    .sb-stat-card { background: #ffffff; border: 1px solid rgba(76,175,80,0.18); border-radius: 10px; padding: 10px 10px; text-align: center; transition: all 0.2s ease; box-shadow: 0 1px 4px rgba(76,175,80,0.08); }
    .sb-stat-card:hover { background: #f1f8e9; border-color: #66bb6a; transform: translateY(-1px); box-shadow: 0 4px 10px rgba(76,175,80,0.20); }
    .sb-stat-card .sb-stat-icon { font-size: 16px; margin-bottom: 4px; line-height: 1; }
    .sb-stat-card .sb-stat-value { font-size: 16px; font-weight: 800; color: #1b5e20 !important; line-height: 1.1; }
    .sb-stat-card .sb-stat-label { font-size: 8.5px; color: #558b2f !important; text-transform: uppercase; letter-spacing: 1px; font-weight: 800; margin-top: 3px; }
    .sb-stat-card.warn .sb-stat-value { color: #e65100 !important; }
    .sb-stat-card.good .sb-stat-value { color: #2e7d32 !important; }
    .sb-stat-card.blue .sb-stat-value { color: #1976d2 !important; }
    .sb-date { text-align: center; padding: 10px 14px 6px; font-size: 10px; color: #558b2f !important; letter-spacing: 1px; font-weight: 800; text-transform: uppercase; }

    section[data-testid="stSidebar"] .stButton > button {
        background: linear-gradient(135deg, #ffffff 0%, #f1f8e9 100%) !important;
        border: 1.5px solid rgba(239,68,68,0.35) !important;
        color: #c62828 !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        transition: all 0.2s ease !important;
        margin: 0 14px 16px !important;
        width: calc(100% - 28px) !important;
        box-shadow: 0 2px 6px rgba(239,68,68,0.10);
    }
    section[data-testid="stSidebar"] .stButton > button:hover {
        background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important;
        border-color: #ef4444 !important;
        color: #ffffff !important;
        box-shadow: 0 6px 14px rgba(239,68,68,0.4) !important;
    }
    section[data-testid="stSidebar"] .stButton > button p,
    section[data-testid="stSidebar"] .stButton > button span,
    section[data-testid="stSidebar"] .stButton > button div { color: inherit !important; }

    /* ============================================================
       PROFESSIONAL DASHBOARD
       ============================================================ */
    .dash-hero {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #3b82f6 100%);
        border-radius: 20px; padding: 26px 32px; margin-bottom: 22px;
        color: #fff; position: relative; overflow: hidden;
        box-shadow: 0 12px 32px rgba(37, 99, 235, 0.35);
    }
    .dash-hero::before { content: ''; position: absolute; top: -50%; right: -10%; width: 420px; height: 420px; background: radial-gradient(circle, rgba(255,255,255,0.18) 0%, transparent 70%); border-radius: 50%; }
    .dash-hero::after { content: ''; position: absolute; bottom: -80%; left: -10%; width: 320px; height: 320px; background: radial-gradient(circle, rgba(255,255,255,0.10) 0%, transparent 70%); border-radius: 50%; }
    .dash-hero-content { position: relative; z-index: 2; display: flex; justify-content: space-between; align-items: center; gap: 24px; flex-wrap: wrap; }
    .dash-hero-left { display: flex; flex-direction: column; gap: 6px; }
    .dash-hero-welcome { font-size: 11px; font-weight: 800; letter-spacing: 2.5px; text-transform: uppercase; color: rgba(255,255,255,0.75) !important; }
    .dash-hero-title { font-size: 28px; font-weight: 800; margin: 0; color: #fff !important; letter-spacing: -0.6px; }
    .dash-hero-sub { font-size: 13px; color: rgba(255,255,255,0.9) !important; font-weight: 500; }
    .dash-hero-right { display: flex; gap: 22px; align-items: center; }
    .dash-hero-stat { text-align: right; }
    .dash-hero-stat-label { font-size: 10px; text-transform: uppercase; letter-spacing: 1.6px; color: rgba(255,255,255,0.7) !important; font-weight: 700; }
    .dash-hero-stat-val { font-size: 24px; font-weight: 800; color: #fff !important; line-height: 1.1; }

    .dash-section-title { display: flex; align-items: center; gap: 10px; font-size: 15px; font-weight: 800; color: #1e3a8a !important; margin: 10px 0 14px 0; letter-spacing: 0.3px; }
    .dash-section-title::before { content: ''; display: inline-block; width: 4px; height: 18px; background: linear-gradient(180deg, #2563eb 0%, #3b82f6 100%); border-radius: 3px; }

    .kpi-card { background: #ffffff; border-radius: 16px; padding: 20px 22px; border: 1px solid rgba(37,99,235,0.12); box-shadow: 0 4px 16px rgba(37,99,235,0.08); display: flex; flex-direction: column; gap: 12px; transition: all 0.2s ease; position: relative; overflow: hidden; height: 100%; }
    .kpi-card:hover { transform: translateY(-3px); box-shadow: 0 12px 28px rgba(37,99,235,0.18); border-color: rgba(37,99,235,0.30); }
    .kpi-card::after { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: linear-gradient(90deg, #2563eb 0%, #3b82f6 100%); opacity: 0; transition: opacity 0.2s ease; }
    .kpi-card:hover::after { opacity: 1; }
    .kpi-top { display: flex; align-items: center; justify-content: space-between; }
    .kpi-icon { width: 44px; height: 44px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-size: 20px; background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); color: #1e40af !important; box-shadow: 0 3px 8px rgba(37,99,235,0.15); }
    .kpi-icon.green { background: linear-gradient(135deg, #d1fae5 0%, #a7f3d0 100%); color: #065f46 !important; }
    .kpi-icon.orange { background: linear-gradient(135deg, #fed7aa 0%, #fdba74 100%); color: #9a3412 !important; }
    .kpi-icon.purple { background: linear-gradient(135deg, #e9d5ff 0%, #d8b4fe 100%); color: #6b21a8 !important; }
    .kpi-icon.teal { background: linear-gradient(135deg, #ccfbf1 0%, #99f6e4 100%); color: #115e59 !important; }
    .kpi-icon.red { background: linear-gradient(135deg, #fecaca 0%, #fca5a5 100%); color: #991b1b !important; }
    .kpi-label { font-size: 10.5px; font-weight: 800; color: #64748b !important; text-transform: uppercase; letter-spacing: 1.2px; }
    .kpi-value { font-size: 32px; font-weight: 800; color: #0f172a !important; line-height: 1; letter-spacing: -1.2px; }
    .kpi-sub { font-size: 11.5px; color: #64748b !important; font-weight: 600; }
    .kpi-sub.blue { color: #2563eb !important; }
    .kpi-sub.green { color: #059669 !important; }
    .kpi-sub.orange { color: #ea580c !important; }
    .kpi-sub.red { color: #dc2626 !important; }
    .kpi-value-group { display: flex; align-items: baseline; gap: 6px; }
    .kpi-value-unit { font-size: 12px; color: #94a3b8 !important; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }

    .team-card { background: #ffffff; border-radius: 14px; padding: 14px 18px; border: 1px solid rgba(37,99,235,0.10); box-shadow: 0 2px 10px rgba(37,99,235,0.06); display: flex; align-items: center; gap: 14px; margin-bottom: 10px; transition: all 0.18s ease; }
    .team-card:hover { transform: translateX(3px); box-shadow: 0 8px 22px rgba(37,99,235,0.15); border-color: rgba(37,99,235,0.28); }
    .team-avatar { width: 46px; height: 46px; border-radius: 12px; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 17px; color: #fff !important; flex-shrink: 0; text-transform: uppercase; box-shadow: 0 4px 10px rgba(37,99,235,0.3); background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); }
    .team-avatar.booker { background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); }
    .team-avatar.salesman { background: linear-gradient(135deg, #8b5cf6 0%, #6d28d9 100%); box-shadow: 0 4px 10px rgba(139,92,246,0.3); }
    .team-info { flex: 1; min-width: 0; }
    .team-name { font-size: 14.5px; font-weight: 700; color: #0f172a !important; margin-bottom: 3px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .team-meta { font-size: 11.5px; color: #64748b !important; font-weight: 500; }
    .team-meta b { color: #2563eb !important; font-weight: 700; }
    .team-meta.warn b { color: #ea580c !important; }
    .team-meta.good b { color: #059669 !important; }
    .team-badge { background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); color: #1e40af !important; font-size: 10px; font-weight: 800; padding: 5px 10px; border-radius: 8px; letter-spacing: 0.8px; border: 1px solid rgba(37,99,235,0.15); flex-shrink: 0; }
    .team-badge.orange { background: linear-gradient(135deg, #fff7ed 0%, #ffedd5 100%); color: #9a3412 !important; border-color: rgba(234,88,12,0.2); }
    .empty-team { background: #ffffff; border: 2px dashed rgba(37,99,235,0.25); border-radius: 14px; padding: 28px; text-align: center; color: #64748b !important; font-size: 13px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

components.html("""
<script>
(function(){
    function killManageApp(){ try{ var doc=window.parent.document;
        ['[data-testid="manage-app-button"]','[data-testid="stAppDeployButton"]',
         '[data-testid="stCloudAppManageButton"]','.stAppDeployButton'].forEach(function(s){
            doc.querySelectorAll(s).forEach(function(el){ el.style.setProperty('display','none','important'); }); });
    }catch(e){} }
    function tick(){ killManageApp(); }
    setTimeout(tick,300); setTimeout(tick,1000); setInterval(tick,1500);
})();
</script>
""", height=0)

COMPANY_NAME = "AL-BARAKAH ENTERPRISES"

# ============================================================
# PRODUCT LIST
# ============================================================
PRODUCTS = sorted([
    {"code":"51","name":"BOOMZ LIQUID MANGO","price":135},
    {"code":"4","name":"BADAM DELIGHT CANDY BOX","price":135},
    {"code":"25","name":"BADAM CONE BOX","price":208},
    {"code":"42","name":"CHOCO BITE CRUSHED PEANUT","price":210},
    {"code":"36","name":"CHOKOZO CHOCOLATE","price":129},
    {"code":"35","name":"CHOKOZO STRAWBERRY","price":129},
    {"code":"38","name":"CHOCOFY CHOCOLATE","price":133},
    {"code":"37","name":"CHOCOFY STRAWBERRY","price":133},
    {"code":"29","name":"CHOCOLATE CONE","price":208},
    {"code":"11","name":"CHOCOLATE CONE WITH PEANUT CHUNKS","price":207},
    {"code":"45","name":"COCONUT WAALA","price":145},
    {"code":"7","name":"CRISPEE WAFER ORANGE","price":137},
    {"code":"6","name":"CRISPEE WAFER BANANA","price":137},
    {"code":"8","name":"CRISPEE WAFER STRAWBERRY","price":137},
    {"code":"17","name":"CUP CAKE CHOCOLATE","price":215},
    {"code":"41","name":"DONUT CAKE","price":227},
    {"code":"24","name":"FISHU BIG CHOCO STICK","price":180},
    {"code":"43","name":"HEART BROWMIES","price":224},
    {"code":"34","name":"JIM JAM","price":145},
    {"code":"5","name":"KHATU APPLE CANDY","price":369},
    {"code":"13","name":"KIDS JOY EGG CHOCOLATE WITH BISCUIT","price":446},
    {"code":"12","name":"KOKO MASTI CHOCOLATE TUBE","price":227},
    {"code":"16","name":"KOKO MASTI MILK CREAM TUBE","price":227},
    {"code":"14","name":"KOKO MASTI STRAWBERRY TUBE","price":227},
    {"code":"27","name":"KULFI PISTA MILKY LOLLIPOP BOX","price":137},
    {"code":"10","name":"LUSH STAR CHOC. HAZELNUT","price":282},
    {"code":"47","name":"MAGIC LOLLY POP","price":138},
    {"code":"44","name":"MAKHAN WAALA","price":144},
    {"code":"39","name":"MAKHAN BADAMI 10","price":218},
    {"code":"40","name":"MAKHAN BADAMI TOFFEE","price":140},
    {"code":"33","name":"MAX GUAVA 3-D JELLY","price":202},
    {"code":"19","name":"MAX STRAWBERRY 3-D JELLY","price":202},
    {"code":"18","name":"MELLOW JOY MANGO MARSHMALLOW","price":202},
    {"code":"26","name":"MICKEY POP FRUITY BOX","price":138},
    {"code":"9","name":"MINI CONE STRAWBERRY","price":245},
    {"code":"3","name":"MINT GUM CENTER FILLED COATED BUBBLE","price":203},
    {"code":"28","name":"NUT KHAT CHOCOLATE","price":135},
    {"code":"23","name":"O-MILK NATURAL OAT ENERGY","price":202},
    {"code":"1","name":"OKAY CHOCOLATE VANILLA LAYER CAKE","price":215},
    {"code":"2","name":"OKAY STRAWBERRY VANILLA LAYER CAKE","price":215},
    {"code":"46","name":"PANDA SPONGE CAKE","price":215},
    {"code":"49","name":"ROLLEX WAFER CHOCOLATE","price":224},
    {"code":"48","name":"ROLLEX WAFER STRAWBERRY","price":224},
    {"code":"30","name":"STRAWBERRY CONE","price":208},
    {"code":"15","name":"STRAWBERRY FLAVORED CONE BOX","price":208},
    {"code":"20","name":"SUPREME SOFT CAKE","price":215},
    {"code":"21","name":"SWISS ROLL CAKE STRAWBERRY","price":216},
    {"code":"22","name":"SWISS ROLL CAKE CHOCOLATE","price":216},
    {"code":"50","name":"YUMMY DONUT STRAWBERRY","price":224},
    {"code":"52","name":"GALE EGG BISCUIT","price":144},
    {"code":"53","name":"PEANUT PISTA","price":144},
    {"code":"54","name":"MOBILE CHOCO BEANS","price":187},
    {"code":"55","name":"KHOPRA TOFFEE","price":207},
    {"code":"56","name":"MILK DELIGHT CHOCO","price":224},
    {"code":"57","name":"DOUBLE DECKER CARS CHOCO","price":269},
    {"code":"58","name":"CONEX CONE ORANGE","price":246},
    {"code":"59","name":"CHOCO PILLOW","price":213},
    {"code":"60","name":"SWISS ROLL CAKE BANANA","price":216},
    {"code":"61","name":"GRAIN TREAT CEREAL","price":390},
    {"code":"62","name":"BUTTER COOKIES","price":144},
    {"code":"63","name":"BEN-10","price":380},
    {"code":"64","name":"SWEET HEART STRAW","price":224},
    {"code":"65","name":"MAX STRAWBERRY RS 2","price":140},
    {"code":"66","name":"OAT MILK","price":208},
    {"code":"67","name":"RINGO RS 10","price":144},
    {"code":"68","name":"RINGO RS 5","price":142},
    {"code":"69","name":"KIMS – DONUT CHOCO CHIP","price":224},
    {"code":"70","name":"KIMS – SUPERB M/P","price":148},
    {"code":"71","name":"KIMS – DANISH BUTTER COOKIES M/P","price":148},
    {"code":"72","name":"KIMS – ROYAL CHOCOLATE M/P","price":148},
    {"code":"73","name":"KIMS – GOLE EGG & MILK M/P","price":148},
    {"code":"74","name":"KIMS – MILKY WAY M/P","price":148},
    {"code":"76","name":"KIMS – CHOCODIP COOKIES M/P","price":148},
    {"code":"77","name":"KIMS – PEANUT PISTA M/P","price":148},
    {"code":"78","name":"KIMS – LEMON SANDWICH M/P","price":148},
    {"code":"102","name":"KIMS – COCONUT COOKIES M/P","price":148},
    {"code":"75","name":"KIMS – MILKY WAY T/P","price":137},
    {"code":"79","name":"KIMS – YUMTUM STRAWBERRY","price":141},
    {"code":"80","name":"KIMS – CHOCO BITE CONE","price":207},
    {"code":"81","name":"KIMS – TIK TOK ORANGE WAFER","price":135},
    {"code":"82","name":"KIMS – MAGIC STICKS WAFER ROLLIES","price":258},
    {"code":"83","name":"KIMS – FRUITO CONE JELLY JAR","price":269},
    {"code":"84","name":"KIMS – PANDA POP ORANGE & LIME","price":138},
    {"code":"85","name":"KIMS – PANDA POP STRAWBERRY & VANILLA","price":138},
    {"code":"86","name":"KIMS – ZOOLO SOUR CHEW BLUEBERRY","price":97},
    {"code":"87","name":"KIMS – CHOKOZO STRAWBERRY","price":129},
    {"code":"88","name":"KIMS – CHOKOZO CHOCOLATE CREAM","price":129},
    {"code":"89","name":"KIMS – CHOKOZO MILK MAZA","price":129},
    {"code":"90","name":"KIMS – SIP STRAWBERRY","price":328},
    {"code":"91","name":"KIMS – CHAMPION DELICIOUS MILK CHOCOLATE JAR","price":269},
    {"code":"92","name":"KIMS – CHOCO DELIGHT CREAMY CHOCOLATE","price":219},
    {"code":"93","name":"KIMS – NUT KHUT CHOCOLATE","price":135},
    {"code":"94","name":"KIMS – FIX CHOCOLATE CREAM","price":135},
    {"code":"95","name":"KIMS – SWISS ROLL STRAWBERRY & VANILLA","price":216},
    {"code":"96","name":"KIMS – SWISS ROLL CHOCOLATE","price":216},
    {"code":"97","name":"KIMS – SWISS ROLL BANANA","price":216},
    {"code":"98","name":"KIMS – KUP CAKE CHOCOLATE","price":213},
    {"code":"99","name":"KIMS – KUP CAKE STRAWBERRY","price":213},
    {"code":"100","name":"KIMS – DONUT STRAWBERRY","price":224},
    {"code":"101","name":"KIMS – BIG SIX BUBBLE GUM","price":406},
    {"code":"103","name":"KIMS – FOXI BUNTY SHEET","price":188},
    {"code":"104","name":"KIMS – SUPERB SOFT CAKE EGG & MILK","price":214},
    {"code":"105","name":"KIMS – SUPERB SOFT CAKE CHOCOLATE & MILK","price":214},
    {"code":"106","name":"KIMS – TWINKLE CAKE STRAWBERRY & VANILLA","price":138},
    {"code":"107","name":"KIMS – TWINKLE CAKE CHOCOLATE VANILLA","price":138},
    {"code":"108","name":"KIMS – SHAKE BANANA CAKE","price":135},
    {"code":"109","name":"KIMS – CHOCO DELIGHT CREAMY CHOCOLATE","price":219},
    {"code":"110","name":"KIMS – GUMMY GUAVA JELLY","price":202},
    {"code":"111","name":"KIMS – GUMMY STRAWBERRY JELLY","price":202},
    {"code":"112","name":"KIMS – AMROOD CANDY","price":139},
    {"code":"113","name":"KIMS – KHOPRA PLUS","price":139},
    {"code":"114","name":"KIMS – AAM MAZA CANDY","price":139},
    {"code":"115","name":"KIMS – FRUITO CANDY","price":139},
    {"code":"116","name":"KIMS – PAN MASALA CANDY BOX","price":139},
    {"code":"117","name":"KIMS – FANTO ORANGE CANDY BOX","price":139},
    {"code":"118","name":"KIMS – CHIPS FRIES (SPICY POTATO STICKS)","price":135}
], key=lambda x: x["name"])

# ============================================================
# AUTH SCREEN
# ============================================================
def render_auth_page():
    st.markdown(f"<div class='auth-title'>🧾 {COMPANY_NAME}</div>", unsafe_allow_html=True)
    st.markdown("<div class='auth-subtitle'>Billing Software 2026 — Login</div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<div class='auth-card'>", unsafe_allow_html=True)
        tab1, tab2, tab3 = st.tabs(["🔐 Admin Login", "👤 Booker Login", "📝 Signup"])

        with tab1:
            st.markdown("### 🔐 Admin Login")
            login_user = st.text_input("Username:", key="login_username", placeholder="admin username")
            login_pass = st.text_input("Password:", key="login_password", type="password", placeholder="password")
            if st.button("🔓 Admin Login", key="btn_login", use_container_width=True, type="primary"):
                users = load_users()
                uname = login_user.strip().lower()
                if uname == "" or login_pass == "": st.error("❌ Dono daalo")
                elif uname not in users: st.error("❌ Username exist nahi karta")
                elif users[uname].get("password_hash") != hash_password(login_pass): st.error("❌ Password galat")
                else:
                    st.session_state["logged_in_user"] = uname
                    st.session_state["display_name"] = users[uname].get("display_name", uname)
                    st.session_state["role"] = "admin"
                    st.session_state["page"] = "📊 Dashboard"
                    st.rerun()

        with tab2:
            st.markdown("### 👤 Booker Login")
            st.caption("Sirf wahi bookers jinka naam admin ne Bookers section mein add kiya hai")
            bk_user = st.text_input("Booker Name:", key="bk_login_user", placeholder="apka naam")
            bk_pass = st.text_input("Password:", key="bk_login_pass", type="password", placeholder="password")
            st.caption("💡 Default password: **1234** (Admin se change karwa lo)")
            if st.button("👤 Booker Login", key="btn_bk_login", use_container_width=True, type="primary"):
                registry = load_bookers()
                bk_uname = sanitize_username(bk_user.strip())
                if not bk_uname or not bk_pass: st.error("❌ Dono daalo")
                elif bk_uname not in registry: st.error("❌ Ye booker exist nahi karta. Admin se contact karo.")
                elif registry[bk_uname].get("password_hash") != hash_password(bk_pass): st.error("❌ Password galat")
                else:
                    admin_uname = registry[bk_uname].get("admin", "")
                    admin_db = load_database_for(admin_uname)
                    if registry[bk_uname].get("display_name") not in admin_db.get("bookers", []):
                        st.error("❌ Aapko admin ne bookers list se hata diya hai")
                    else:
                        st.session_state["logged_in_user"] = admin_uname
                        st.session_state["display_name"] = registry[bk_uname].get("display_name", bk_user)
                        st.session_state["role"] = "booker"
                        st.session_state["booker_name"] = registry[bk_uname].get("display_name", bk_user)
                        st.session_state["booker_admin"] = admin_uname
                        st.session_state["bk_page"] = "🧾 New Bill"
                        st.session_state["bk_bill_draft"] = []
                        st.session_state["bk_load_draft"] = []
                        st.session_state["bk_current_shop"] = ""
                        st.session_state["bk_bills_created"] = 0
                        st.rerun()

        with tab3:
            st.markdown("### 📝 Admin Signup")
            su_user = st.text_input("Naya Username:", key="su_username", placeholder="3-20 chars")
            su_pass = st.text_input("Password:", key="su_password", type="password")
            su_pass2 = st.text_input("Confirm:", key="su_password2", type="password")
            if st.button("✅ Signup Karo", key="btn_signup", use_container_width=True):
                users = load_users()
                uname_safe = sanitize_username(su_user.strip())
                if len(uname_safe) < 3: st.error("❌ Username chhota")
                elif len(su_pass) < 4: st.error("❌ Password chhota")
                elif su_pass != su_pass2: st.error("❌ Match nahi")
                elif uname_safe in users: st.error("❌ Exist karta hai")
                else:
                    users[uname_safe] = {"username": uname_safe, "display_name": su_user.strip(),
                        "password_hash": hash_password(su_pass),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                    save_users(users)
                    with open(user_data_file(uname_safe), "w", encoding="utf-8") as f:
                        json.dump(default_blank_db(), f, indent=4, ensure_ascii=False)
                    st.success(f"✅ Account ban gaya — username: **{uname_safe}**")
        st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# LOAD DB
# ============================================================
def load_database_for(username):
    if not username: return default_blank_db()
    fpath = user_data_file(username)
    if os.path.exists(fpath):
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k in ["bookers", "salesmen", "load_forms", "petrol_expenses",
                          "lunch_expenses", "custom_products", "dsr_forms", "credit_bills"]:
                    if k not in data: data[k] = []
                for k in ["bookers_salaries", "salesmen_salaries", "product_prices"]:
                    if k not in data: data[k] = {}
                if "discount_packages" not in data or not data["discount_packages"]:
                    data["discount_packages"] = default_discount_packages()
                else:
                    for p in data["discount_packages"]:
                        if "tier3_amount" not in p: p["tier3_amount"] = 0.0
                        if "tier3_pct" not in p: p["tier3_pct"] = 0.0
                if "bills" not in data: data["bills"] = []
                if "next_bill_no" not in data: data["next_bill_no"] = 1
                if "next_credit_id" not in data:
                    data["next_credit_id"] = max([c.get("id", 0) for c in data.get("credit_bills", [])] + [0]) + 1
                if "wholesaler_rule_enabled" not in data: data["wholesaler_rule_enabled"] = True
                if "wholesaler_discount_pct" not in data: data["wholesaler_discount_pct"] = 6.0
                return data
        except Exception: pass
    return default_blank_db()

# ============================================================
# CHECK LOGIN
# ============================================================
if "logged_in_user" not in st.session_state or not st.session_state["logged_in_user"]:
    render_auth_page()
    st.stop()

CURRENT_USER = st.session_state["logged_in_user"]
ROLE = st.session_state.get("role", "admin")

def save_database(db):
    try:
        with open(user_data_file(CURRENT_USER), "w", encoding="utf-8") as f:
            json.dump(db, f, indent=4, ensure_ascii=False)
    except Exception as e:
        st.warning(f"⚠️ Save: {e}")

def parse_date(dstr):
    try: return datetime.strptime(dstr, "%d-%m-%Y").date()
    except Exception: return None

if "database" not in st.session_state:
    st.session_state.database = load_database_for(CURRENT_USER)
if st.session_state.get("_db_user") != CURRENT_USER:
    st.session_state.database = load_database_for(CURRENT_USER)
    st.session_state["_db_user"] = CURRENT_USER

for _k, _v in [("_prev_prod", None), ("last_bill_no", None), ("download_file", None),
               ("_dl_counter", 0), ("page", "📊 Dashboard"), ("view_bill_key", None),
               ("view_lf_key", None), ("view_dsr_key", None), ("view_credit_key", None),
               ("bk_page", "🧾 New Bill"), ("bk_bill_draft", []), ("bk_load_draft", []),
               ("bk_current_shop", ""), ("bk_bills_created", 0),
               ("bk_reset_item_form", False), ("bk_reset_shop_form", False)]:
    if _k not in st.session_state: st.session_state[_k] = _v

db = st.session_state.database
for k in ["bookers", "salesmen", "load_forms", "petrol_expenses",
          "lunch_expenses", "custom_products", "dsr_forms", "credit_bills"]:
    if k not in db: db[k] = []
for k in ["bookers_salaries", "salesmen_salaries", "product_prices"]:
    if k not in db: db[k] = {}
if "discount_packages" not in db or not db["discount_packages"]:
    db["discount_packages"] = default_discount_packages()
for p in db["discount_packages"]:
    if "tier3_amount" not in p: p["tier3_amount"] = 0.0
    if "tier3_pct" not in p: p["tier3_pct"] = 0.0
if "next_credit_id" not in db:
    db["next_credit_id"] = max([c.get("id", 0) for c in db.get("credit_bills", [])] + [0]) + 1
if "wholesaler_rule_enabled" not in db: db["wholesaler_rule_enabled"] = True
if "wholesaler_discount_pct" not in db: db["wholesaler_discount_pct"] = 6.0

# ============================================================
# HELPERS
# ============================================================
def is_wholesaler(shop_name):
    if not shop_name: return False
    s = str(shop_name).lower()
    s_compact = "".join(ch for ch in s if ch.isalnum())
    for token in ["wholeseller", "wholesaler", "wholseller", "wholsaler",
                  "whole seller", "whole saler", "holeseller", "holesaler"]:
        if token.replace(" ", "") in s_compact: return True
    return False

def wholesaler_rule_active(): return bool(db.get("wholesaler_rule_enabled", True))
def get_wholesaler_pct():
    try: return float(db.get("wholesaler_discount_pct", 6.0))
    except Exception: return 6.0

def get_all_products():
    custom = db.get("custom_products", [])
    merged = list(PRODUCTS)
    seen_codes = {str(p["code"]) for p in merged}
    for cp in custom:
        code = str(cp.get("code", "")).strip()
        name = str(cp.get("name", "")).strip()
        if not code or not name or code in seen_codes: continue
        try: price = float(cp.get("price", 0))
        except Exception: price = 0.0
        merged.append({"code": code, "name": name, "price": price}); seen_codes.add(code)
    return sorted(merged, key=lambda x: x["name"])

def get_price(code, base_price):
    custom = db.get("product_prices", {})
    if str(code) in custom:
        try: return float(custom[str(code)])
        except Exception: return float(base_price)
    return float(base_price)

def get_all_active_packages(): return [p for p in db.get("discount_packages", []) if p.get("active")]

def get_package_discount_pct(bill_total):
    active_pkgs = get_all_active_packages()
    if not active_pkgs: return 0.0, None, None
    best_pct = 0.0; best_name = None; best_tier = None
    for pkg in active_pkgs:
        pkg_name = pkg.get("name", "Package")
        tiers = [
            (float(pkg.get("tier3_amount", 0) or 0), float(pkg.get("tier3_pct", 0) or 0), "Tier3"),
            (float(pkg.get("tier2_amount", 0) or 0), float(pkg.get("tier2_pct", 0) or 0), "Tier2"),
            (float(pkg.get("tier1_amount", 0) or 0), float(pkg.get("tier1_pct", 0) or 0), "Tier1"),
        ]
        tiers.sort(key=lambda x: x[0], reverse=True)
        for amt, pct, label in tiers:
            if amt > 0 and bill_total > amt and pct > best_pct:
                best_pct = pct; best_name = pkg_name
                best_tier = f"{label} (> {amt:,.0f})"; break
    if best_pct > 0: return best_pct, best_name, best_tier
    return 0.0, (active_pkgs[0].get("name") if active_pkgs else None), None

def get_effective_discount_pct(shop_name, bill_total):
    if wholesaler_rule_active() and is_wholesaler(shop_name):
        pct = get_wholesaler_pct()
        return pct, f"Wholesaler ({pct}%)", "Wholesaler Rule"
    return get_package_discount_pct(bill_total)

def get_dsr_credit_info(dsr):
    booker = dsr.get("booker", "").strip()
    credit_total = 0.0; credit_shops = {}; credit_details = []
    for sb in dsr.get("source_bills", []):
        shop = sb.get("shop", ""); bill_no = sb.get("bill_no", "")
        for cb in db.get("credit_bills", []):
            if (str(cb.get("booker", "")).strip() == booker
                and str(cb.get("bill_no", "")) == str(bill_no)
                and str(cb.get("shop", "")) == str(shop)):
                amt = float(cb.get("total_net", 0))
                credit_total += amt
                credit_shops[shop] = credit_shops.get(shop, 0) + amt
                credit_details.append({"shop": shop, "bill_no": bill_no, "amount": amt,
                    "status": cb.get("status", "pending"), "credit_id": cb.get("id")})
                break
    return credit_total, credit_shops, credit_details

def show_auto_download():
    if st.session_state.get("download_file"):
        fname, fdata = st.session_state["download_file"]
        st.session_state["download_file"] = None
        st.session_state["_dl_counter"] += 1
        st.markdown('<div class="auto-dl-hidden">', unsafe_allow_html=True)
        st.download_button(label=f"Download {fname}", data=fdata, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"auto_dl_{st.session_state['_dl_counter']}")
        st.markdown('</div>', unsafe_allow_html=True)
        components.html("""
        <script>
        (function(){ var tries=0; function attempt(){ tries++;
            var btns = window.parent.document.querySelectorAll('[data-testid="stDownloadButton"] button');
            if (btns.length > 0){ btns[btns.length - 1].click(); return; }
            if (tries < 25) setTimeout(attempt, 200); }
            setTimeout(attempt, 300); })();
        </script>""", height=0)

# ============================================================
# BOOKER UI
# ============================================================
def render_booker_sidebar():
    with st.sidebar:
        bk_name = st.session_state.get("booker_name", "Booker")
        avatar_letter = (bk_name[0] if bk_name else "B").upper()
        st.markdown(f"""
        <div class="sb-brand">
            <div class="sb-brand-logo">👤</div>
            <div class="sb-brand-name">BOOKER PANEL</div>
            <div class="sb-brand-sub">AL-BARAKAH</div>
            <div class="sb-brand-status"><span class="sb-brand-dot"></span>BOOKER LOGGED IN</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"""
        <div class="sb-user">
            <div class="sb-user-avatar">{avatar_letter}</div>
            <div class="sb-user-info">
                <div class="sb-user-label">Booker</div>
                <div class="sb-user-name">{bk_name}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown('<div class="sb-section-label">Menu</div>', unsafe_allow_html=True)

        bk_page = st.radio("BK MENU",
            ["🧾 New Bill", "🛒 All Products", "📦 Load Form Status"],
            key="bk_page_selector", label_visibility="collapsed")
        st.session_state["bk_page"] = bk_page

        st.markdown(f'<div class="sb-date">📅 {datetime.now().strftime("%A, %d %b %Y")}</div>', unsafe_allow_html=True)

        bill_count = st.session_state.get("bk_bills_created", 0)
        load_count = len(st.session_state.get("bk_load_draft", []))
        total_boxes = sum(int(it.get("Boxes", 0)) for it in st.session_state.get("bk_load_draft", []))
        st.markdown(f"""
        <div class="sb-stats">
            <div class="sb-section-label" style="padding: 0 6px 10px;">Session</div>
            <div class="sb-stats-grid">
                <div class="sb-stat-card blue"><div class="sb-stat-icon">🧾</div><div class="sb-stat-value">{bill_count}</div><div class="sb-stat-label">Bills</div></div>
                <div class="sb-stat-card good"><div class="sb-stat-icon">📦</div><div class="sb-stat-value">{total_boxes}</div><div class="sb-stat-label">Boxes</div></div>
                <div class="sb-stat-card"><div class="sb-stat-icon">🎯</div><div class="sb-stat-value">{load_count}</div><div class="sb-stat-label">Products</div></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("🚪  Logout", key="btn_bk_logout", use_container_width=True):
            for k in ["logged_in_user", "display_name", "role", "booker_name", "booker_admin",
                      "bk_bill_draft", "bk_load_draft", "bk_current_shop",
                      "bk_bill_items", "bk_page", "database", "_db_user",
                      "bk_reset_item_form", "bk_reset_shop_form"]:
                st.session_state[k] = None if k not in ["bk_bill_draft", "bk_load_draft", "bk_bill_items"] else []
            st.session_state["bk_bills_created"] = 0
            st.rerun()

def render_booker_new_bill():
    bk_name = st.session_state.get("booker_name", "")
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🧾 New Bill — <small style='color:#2e7d32;'>👤 {bk_name}</small></h1>", unsafe_allow_html=True)
    st.markdown("---")

    if st.session_state.get("bk_reset_item_form"):
        st.session_state["bk_product_sel"] = "-- Select --"
        st.session_state["bk_qty"] = 0
        st.session_state["bk_reset_item_form"] = False

    if wholesaler_rule_active():
        st.markdown(f"<div class='hint-box' style='background:#f3e5f5;border-left-color:#8e24aa;color:#6a1b9a !important;'>🏢 Wholesaler rule ON — 'whole seller' naam wali shop pe {get_wholesaler_pct()}% milega</div>", unsafe_allow_html=True)

    col_a, col_b = st.columns([3, 1])
    with col_a:
        shop = st.text_input("🏪 Shop Name:", value=st.session_state.get("bk_current_shop", ""),
                             key="bk_shop_input", placeholder="Pehle shop ka naam likho")
    with col_b:
        st.markdown("<br>", unsafe_allow_html=True)

    if shop:
        new_shop = shop.strip()
        if new_shop != st.session_state.get("bk_current_shop", ""):
            st.session_state["bk_current_shop"] = new_shop
            st.session_state["bk_bill_draft"] = []
            st.rerun()

    if not shop.strip():
        st.warning("⚠️ Pehle shop ka naam likho, phir item add karo")
        return

    if wholesaler_rule_active() and is_wholesaler(shop):
        st.markdown(f"<div class='hint-box' style='background:#f3e5f5;border-left-color:#8e24aa;color:#6a1b9a !important;font-weight:700;'>🏢 Wholesaler shop detected — {get_wholesaler_pct()}% discount lagega</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### ➕ Add Item")
    all_products = get_all_products()

    search_txt = st.text_input("🔍 Search:", key="bk_search", placeholder="Type name...")
    su = search_txt.strip().upper()
    filtered = [p for p in all_products if su in p["name"].upper() or su in str(p["code"]).upper()] if su else all_products
    filtered_labels = [f"{p['name']} (Code: {p['code']})" for p in filtered]

    c1, c2, c3 = st.columns([3, 1, 1])
    with c1:
        selected = st.selectbox("Product:", options=["-- Select --"] + filtered_labels, key="bk_product_sel")
    with c2:
        boxes = st.number_input("Boxes:", min_value=0, step=1, value=0, key="bk_qty")
    with c3:
        st.markdown("<br>", unsafe_allow_html=True)
        add_clicked = st.button("➕ Add Item", key="bk_add_item_btn", use_container_width=True, type="primary")

    selected_product = None
    if selected != "-- Select --":
        for p in all_products:
            if f"{p['name']} (Code: {p['code']})" == selected:
                selected_product = p; break
    if selected_product:
        tp = get_price(selected_product["code"], selected_product["price"])
        st.success(f"💰 {selected_product['name']} — TP: **Rs {tp:,.0f}** | Boxes: {boxes} → Total: **Rs {boxes * tp:,.0f}**")

    if add_clicked:
        if not selected_product:
            st.error("❌ Product select karo")
        elif boxes <= 0:
            st.error("❌ Boxes daalo")
        else:
            tp = get_price(selected_product["code"], selected_product["price"])
            st.session_state["bk_bill_draft"].append({
                "Code": selected_product["code"], "Product": selected_product["name"],
                "Boxes": int(boxes), "TP/Box": float(tp),
                "Gross": float(boxes * tp), "Discount %": 0.0, "Net": float(boxes * tp),
            })
            st.session_state["bk_reset_item_form"] = True
            st.rerun()

    st.markdown("---")
    draft = st.session_state.get("bk_bill_draft", [])
    if not draft:
        st.info("Koi item nahi. Upar se add karo.")
    else:
        st.markdown(f"### 📋 Current Bill — {shop}")
        st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"],
            "Boxes": it["Boxes"], "TP/Box": it["TP/Box"], "Total": it["Gross"]} for it in draft]),
            use_container_width=True, hide_index=True)

        c1, c2 = st.columns([3, 1])
        with c1:
            remove_idx = st.selectbox("Delete item:", options=["--"] + [f"{i+1}. {it['Product']} ({it['Boxes']} boxes)" for i, it in enumerate(draft)], key="bk_remove_sel")
        with c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑 Remove", key="bk_remove_btn", use_container_width=True):
                if remove_idx != "--":
                    idx = int(remove_idx.split(".")[0]) - 1
                    st.session_state["bk_bill_draft"].pop(idx)
                    st.rerun()

        gross_total = sum(float(it["Gross"]) for it in draft)
        st.markdown(f"<div class='summary-box'><b>Gross Total:</b> <span style='color:#1976d2;font-weight:800;font-size:18px;'>Rs {gross_total:,.0f}</span></div>", unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Done Bill (Save to Admin)", key="bk_done_bill", use_container_width=True, type="primary"):
                finalize_booker_bill(shop, draft)
        with c2:
            if st.button("🗑 Clear Draft", key="bk_clear_draft", use_container_width=True):
                st.session_state["bk_bill_draft"] = []
                st.rerun()

    st.markdown("---")
    load_draft = st.session_state.get("bk_load_draft", [])
    total_load_boxes = sum(int(it.get("Boxes", 0)) for it in load_draft)
    st.markdown(f"### 📦 Load Form — Session Total")
    if not load_draft:
        st.info("Abhi tak koi bill complete nahi hua. Pehle bill banao, load form automatically banega.")
    else:
        st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"]} for it in load_draft]),
                     use_container_width=True, hide_index=True)
        st.markdown(f"<div class='summary-box'><b>Total Products:</b> {len(load_draft)} | <b>Total Boxes:</b> <span style='color:#1976d2;font-weight:800;font-size:18px;'>{total_load_boxes}</span></div>", unsafe_allow_html=True)

        if st.button("📦 Done Load Form (Save to Admin)", key="bk_done_load", use_container_width=True, type="primary"):
            finalize_booker_load_form()

def finalize_booker_bill(shop, draft):
    """Save bill to admin DB. NO Excel download. Just save + reset draft."""
    global db
    bill_no = db["next_bill_no"]
    date_str = datetime.now().strftime("%d-%m-%Y")
    bk_name = st.session_state.get("booker_name", "")
    for it in draft:
        db["bills"].append({
            "Bill No": bill_no, "Date": date_str,
            "Shop": shop, "Order Booker": bk_name, "Salesman": "",
            "Delivery Man": "", "Code": it["Code"], "Product": it["Product"],
            "Boxes": it["Boxes"], "TP/Box": it["TP/Box"],
            "Discount %": it["Discount %"], "Gross": it["Gross"], "Net": it["Net"],
        })
    db["next_bill_no"] += 1
    save_database(db)

    # Add to load draft
    for it in draft:
        found = False
        for ld in st.session_state["bk_load_draft"]:
            if str(ld["Code"]) == str(it["Code"]):
                ld["Boxes"] += int(it["Boxes"]); found = True; break
        if not found:
            st.session_state["bk_load_draft"].append({
                "Code": it["Code"], "Product": it["Product"], "Boxes": int(it["Boxes"])})

    # NO EXCEL DOWNLOAD — just save and reset
    bill_total_net = sum(float(it["Net"]) for it in draft)
    pkg_pct, _, _ = get_effective_discount_pct(shop, bill_total_net)
    after_disc_total = bill_total_net - (bill_total_net * pkg_pct / 100)

    st.session_state["bk_bill_draft"] = []
    st.session_state["bk_bills_created"] = st.session_state.get("bk_bills_created", 0) + 1
    st.session_state["success_msg"] = f"✅ Bill #{bill_no} saved to Admin | {shop} | Net: Rs {after_disc_total:,.0f}"
    st.rerun()

def finalize_booker_load_form():
    """Save load form to admin DB."""
    global db
    load_draft = st.session_state.get("bk_load_draft", [])
    if not load_draft:
        st.session_state["error_msg"] = "❌ Load form empty"; return
    bk_name = st.session_state.get("booker_name", "")
    total_boxes = sum(int(it["Boxes"]) for it in load_draft)
    next_id = 1
    if db.get("load_forms"): next_id = max(lf.get("id", 0) for lf in db["load_forms"]) + 1
    lf_record = {
        "id": next_id, "date": datetime.now().strftime("%d-%m-%Y"),
        "time": datetime.now().strftime("%H:%M"),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "booker": bk_name, "items": load_draft,
        "total_boxes": total_boxes, "total_products": len(load_draft),
        "transferred_to_dsr": False, "refreshed": False,
        "source": "booker",
    }
    db["load_forms"].append(lf_record)
    save_database(db)
    st.session_state["bk_load_draft"] = []
    st.session_state["success_msg"] = f"✅ Load Form #{next_id} saved to Admin | {len(load_draft)} products, {total_boxes} boxes"
    st.rerun()

def render_booker_all_products():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🛒 All Products with TP</h1>", unsafe_allow_html=True)
    st.markdown("---")
    all_products = get_all_products()
    search = st.text_input("🔍 Search:", key="bk_prod_search", placeholder="Type product name or code...")
    su = search.strip().upper()
    shown = [p for p in all_products if su in p["name"].upper() or su in str(p["code"]).upper()] if su else all_products
    st.markdown(f"<div class='summary-box'><b>Total Products:</b> {len(all_products)} | <b>Showing:</b> {len(shown)}</div>", unsafe_allow_html=True)
    if not shown:
        st.info("Koi product nahi mila."); return
    for p in shown:
        tp = get_price(p["code"], p["price"])
        st.markdown(f"""
        <div class="bk-product-card">
            <div>
                <div class="bk-name">{p['name']}</div>
                <div class="bk-code">Code: {p['code']}</div>
            </div>
            <div class="bk-tp">Rs {tp:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)

def render_booker_load_status():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📦 Load Form Status</h1>", unsafe_allow_html=True)
    st.markdown("---")
    load_draft = st.session_state.get("bk_load_draft", [])
    if not load_draft:
        st.info("Abhi tak koi bill complete nahi hua. Pehle New Bill se bills banao.")
        return
    total_boxes = sum(int(it["Boxes"]) for it in load_draft)
    st.markdown(f"""<div class='summary-box'>
        <b>Products:</b> {len(load_draft)} &nbsp;|&nbsp;
        <b>Total Boxes:</b> <span style='color:#1976d2;font-size:18px;font-weight:800;'>{total_boxes}</span>
    </div>""", unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"]} for it in load_draft]),
                 use_container_width=True, hide_index=True)
    if st.button("📦 Done Load Form (Save to Admin)", key="bk_done_load_2", use_container_width=True, type="primary"):
        finalize_booker_load_form()

# ============================================================
# ROUTER
# ============================================================
if ROLE == "booker":
    render_booker_sidebar()
    bk_page = st.session_state.get("bk_page", "🧾 New Bill")
    if bk_page == "🧾 New Bill":
        render_booker_new_bill()
    elif bk_page == "🛒 All Products":
        render_booker_all_products()
    elif bk_page == "📦 Load Form Status":
        render_booker_load_status()
    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
    if st.session_state.get("error_msg"):
        st.error(st.session_state["error_msg"]); st.session_state["error_msg"] = None
    st.stop()

# ============================================================
# ADMIN SIDEBAR
# ============================================================
with st.sidebar:
    display_name = st.session_state.get("display_name", CURRENT_USER)
    avatar_letter = (display_name[0] if display_name else "A").upper()
    st.markdown(f"""
    <div class="sb-brand">
        <div class="sb-brand-logo">🧾</div>
        <div class="sb-brand-name">AL-BARAKAH</div>
        <div class="sb-brand-sub">ENTERPRISES</div>
        <div class="sb-brand-status"><span class="sb-brand-dot"></span>ADMIN</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown(f"""
    <div class="sb-user">
        <div class="sb-user-avatar">{avatar_letter}</div>
        <div class="sb-user-info">
            <div class="sb-user-label">Admin</div>
            <div class="sb-user-name">{display_name}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown('<div class="sb-section-label">Main Menu</div>', unsafe_allow_html=True)
    page = st.radio("MENU",
        ["📊 Dashboard", "🧾 Billing", "🛒 All Products", "🎁 Discount",
         "👤 Bookers", "💰 Bookers Salary", "🧑‍💼 Salesmen", "💰 Salesmen Salary",
         "💵 Daily Expense", "📋 Bills List", "💳 Credit Bills", "📦 Load Form",
         "📋 DSR", "🧮 Calculation"],
        key="page_selector", label_visibility="collapsed")
    st.session_state["page"] = page

    st.markdown(f'<div class="sb-date">📅 {datetime.now().strftime("%A, %d %b %Y")}</div>', unsafe_allow_html=True)

    active_pkgs = get_all_active_packages()
    all_prod_count = len(get_all_products())
    pending_credit = sum(1 for c in db.get("credit_bills", []) if c.get("status") == "pending")
    pending_amt = sum(float(c.get("total_net", 0)) for c in db.get("credit_bills", []) if c.get("status") == "pending")
    st.markdown(f"""
    <div class="sb-stats">
        <div class="sb-stats-grid">
            <div class="sb-stat-card blue"><div class="sb-stat-icon">📦</div><div class="sb-stat-value">{all_prod_count}</div><div class="sb-stat-label">Products</div></div>
            <div class="sb-stat-card"><div class="sb-stat-icon">🧾</div><div class="sb-stat-value">{len(db['bills'])}</div><div class="sb-stat-label">Bills</div></div>
            <div class="sb-stat-card"><div class="sb-stat-icon">👤</div><div class="sb-stat-value">{len(db.get('bookers', []))}</div><div class="sb-stat-label">Bookers</div></div>
            <div class="sb-stat-card warn"><div class="sb-stat-icon">💳</div><div class="sb-stat-value">{pending_credit}</div><div class="sb-stat-label">Cr. Pend</div></div>
        </div>
        <div style="text-align:center;font-size:10px;color:#558b2f;padding:4px 0 8px;font-weight:700;">
            Pending: <b style="color:#e65100;">Rs {pending_amt:,.0f}</b>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if st.button("🚪  Logout", key="btn_logout", use_container_width=True):
        for k in ["logged_in_user", "display_name", "role", "database", "_db_user"]:
            st.session_state[k] = None
        st.session_state["page"] = "📊 Dashboard"
        st.rerun()

# ============================================================
# ADMIN DASHBOARD (RESTORED - PROFESSIONAL)
# ============================================================
def render_admin_dashboard():
    display_name = st.session_state.get("display_name", CURRENT_USER)
    today_str = date.today().strftime("%A, %d %B %Y")
    bookers = db.get("bookers", []); salesmen = db.get("salesmen", [])
    total_products = len(get_all_products()); total_bills = len(db["bills"])
    total_load_forms = len(db.get("load_forms", [])); total_dsr = len(db.get("dsr_forms", []))
    pending_credits = [c for c in db.get("credit_bills", []) if c.get("status") == "pending"]
    paid_credits = [c for c in db.get("credit_bills", []) if c.get("status") == "paid"]
    pending_amt = sum(float(c.get("total_net", 0)) for c in pending_credits)
    paid_amt = sum(float(c.get("total_net", 0)) for c in paid_credits)

    # HERO BANNER
    st.markdown(f"""
    <div class="dash-hero">
        <div class="dash-hero-content">
            <div class="dash-hero-left">
                <div class="dash-hero-welcome">Welcome back</div>
                <h1 class="dash-hero-title">👋 {display_name}</h1>
                <div class="dash-hero-sub">🏢 {COMPANY_NAME} &nbsp;·&nbsp; 📅 {today_str}</div>
            </div>
            <div class="dash-hero-right">
                <div class="dash-hero-stat">
                    <div class="dash-hero-stat-label">Total Bills</div>
                    <div class="dash-hero-stat-val">{total_bills}</div>
                </div>
                <div class="dash-hero-stat">
                    <div class="dash-hero-stat-label">Products</div>
                    <div class="dash-hero-stat-val">{total_products}</div>
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # KPI 1 - Business Overview
    st.markdown('<div class="dash-section-title">Business Overview</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon">👤</div><div class="kpi-label">Bookers</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{len(bookers)}</div><div class="kpi-value-unit">members</div></div>
            <div class="kpi-sub blue">Order booking team</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon purple">🧑‍💼</div><div class="kpi-label">Salesmen</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{len(salesmen)}</div><div class="kpi-value-unit">members</div></div>
            <div class="kpi-sub blue">Field sales team</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon teal">📦</div><div class="kpi-label">Products</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{total_products}</div><div class="kpi-value-unit">items</div></div>
            <div class="kpi-sub blue">Active catalog</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon">🧾</div><div class="kpi-label">Total Bills</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{total_bills}</div><div class="kpi-value-unit">bills</div></div>
            <div class="kpi-sub blue">Generated so far</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # KPI 2 - Credit & Operations
    st.markdown('<div class="dash-section-title">Credit & Operations</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon orange">⏳</div><div class="kpi-label">Pending Credit</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{len(pending_credits)}</div><div class="kpi-value-unit">bills</div></div>
            <div class="kpi-sub orange">Rs {pending_amt:,.0f} receivable</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon green">✅</div><div class="kpi-label">Paid Credit</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{len(paid_credits)}</div><div class="kpi-value-unit">bills</div></div>
            <div class="kpi-sub green">Rs {paid_amt:,.0f} settled</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon">📦</div><div class="kpi-label">Load Forms</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{total_load_forms}</div><div class="kpi-value-unit">saved</div></div>
            <div class="kpi-sub blue">Stock dispatch forms</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="kpi-card">
            <div class="kpi-top"><div class="kpi-icon red">📋</div><div class="kpi-label">DSR Forms</div></div>
            <div class="kpi-value-group"><div class="kpi-value">{total_dsr}</div><div class="kpi-value-unit">active</div></div>
            <div class="kpi-sub blue">Daily sales reports</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    st.markdown("---")

    # TEAM SECTION
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="dash-section-title">👤 Order Booker Team</div>', unsafe_allow_html=True)
        if not bookers:
            st.markdown('<div class="empty-team">Abhi tak koi booker add nahi hua</div>', unsafe_allow_html=True)
        else:
            for b_name in bookers:
                sd = db.get("bookers_salaries", {}).get(b_name, {})
                base = sd.get("base_salary", 0); txns = sd.get("transactions", [])
                adv_p = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status", "pending") == "pending")
                short_p = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status", "pending") == "pending")
                remaining = base - adv_p - short_p
                if base == 0: sub_class = ""; sub_text = 'Salary not set'
                elif remaining > base * 0.7: sub_class = "good"; sub_text = f'Remaining: <b>Rs {remaining:,.0f}</b>'
                elif remaining > 0: sub_class = ""; sub_text = f'Remaining: <b>Rs {remaining:,.0f}</b>'
                else: sub_class = "warn"; sub_text = f'Over Paid: <b>Rs {abs(remaining):,.0f}</b>'
                al = (b_name[0] if b_name else "?").upper()
                st.markdown(f"""<div class="team-card">
                    <div class="team-avatar booker">{al}</div>
                    <div class="team-info"><div class="team-name">{b_name}</div><div class="team-meta {sub_class}">{sub_text}</div></div>
                    <div class="team-badge">BOOKER</div>
                </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="dash-section-title">🧑‍💼 Salesmen Team</div>', unsafe_allow_html=True)
        if not salesmen:
            st.markdown('<div class="empty-team">Abhi tak koi salesman add nahi hua</div>', unsafe_allow_html=True)
        else:
            for s_name in salesmen:
                sd = db.get("salesmen_salaries", {}).get(s_name, {})
                base = sd.get("base_salary", 0); txns = sd.get("transactions", [])
                adv_p = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status", "pending") == "pending")
                short_p = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status", "pending") == "pending")
                remaining = base - adv_p - short_p
                if base == 0: sub_class = ""; sub_text = 'Salary not set'
                elif remaining > base * 0.7: sub_class = "good"; sub_text = f'Remaining: <b>Rs {remaining:,.0f}</b>'
                elif remaining > 0: sub_class = ""; sub_text = f'Remaining: <b>Rs {remaining:,.0f}</b>'
                else: sub_class = "warn"; sub_text = f'Over Paid: <b>Rs {abs(remaining):,.0f}</b>'
                al = (s_name[0] if s_name else "?").upper()
                st.markdown(f"""<div class="team-card">
                    <div class="team-avatar salesman">{al}</div>
                    <div class="team-info"><div class="team-name">{s_name}</div><div class="team-meta {sub_class}">{sub_text}</div></div>
                    <div class="team-badge orange">SALESMAN</div>
                </div>""", unsafe_allow_html=True)

# ============================================================
# ADMIN DISCOUNT
# ============================================================
def render_admin_discount():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🎁 Discount Settings</h1>", unsafe_allow_html=True)
    st.markdown("---")
    wh_on = bool(db.get("wholesaler_rule_enabled", True))
    wh_pct = get_wholesaler_pct()
    st.markdown("### 🏢 Wholesaler Special Rule")
    wc1, wc2 = st.columns([1, 1])
    with wc1:
        new_on = st.checkbox("🏢 Rule Enable Karo", value=wh_on, key="wholesaler_toggle")
        if new_on != wh_on: db["wholesaler_rule_enabled"] = bool(new_on); save_database(db); st.rerun()
    with wc2:
        new_pct = st.number_input("Discount %:", value=float(wh_pct), min_value=0.0, max_value=100.0, step=0.5, key="wholesaler_pct_input")
        if abs(float(new_pct) - float(wh_pct)) > 0.001: db["wholesaler_discount_pct"] = float(new_pct); save_database(db); st.rerun()

    st.markdown("---")
    pkgs = db.get("discount_packages", [])
    for i, pkg in enumerate(pkgs):
        pid = pkg.get("id", i+1)
        with st.expander(f"🎁 {pkg.get('name', 'Package')} {'✅' if pkg.get('active') else '⭕'}"):
            c1, c2 = st.columns([1, 5])
            with c1: new_active = st.checkbox("Active", value=pkg.get("active", False), key=f"active_{pid}")
            with c2: new_name = st.text_input("Name:", value=pkg.get("name", ""), key=f"pkgname_{pid}")
            t1, t2, t3 = st.columns(3)
            with t1:
                ta1 = st.number_input("Min Rs (T1)", value=float(pkg.get("tier1_amount", 0)), key=f"t1a_{pid}")
                tp1 = st.number_input("% T1", value=float(pkg.get("tier1_pct", 0)), key=f"t1p_{pid}")
            with t2:
                ta2 = st.number_input("Min Rs (T2)", value=float(pkg.get("tier2_amount", 0)), key=f"t2a_{pid}")
                tp2 = st.number_input("% T2", value=float(pkg.get("tier2_pct", 0)), key=f"t2p_{pid}")
            with t3:
                ta3 = st.number_input("Min Rs (T3)", value=float(pkg.get("tier3_amount", 0)), key=f"t3a_{pid}")
                tp3 = st.number_input("% T3", value=float(pkg.get("tier3_pct", 0)), key=f"t3p_{pid}")
            if st.button("💾 Save", key=f"savepkg_{pid}", type="primary"):
                for pp in db["discount_packages"]:
                    if pp.get("id") == pid:
                        pp["name"] = new_name; pp["active"] = new_active
                        pp["tier1_amount"] = float(ta1); pp["tier1_pct"] = float(tp1)
                        pp["tier2_amount"] = float(ta2); pp["tier2_pct"] = float(tp2)
                        pp["tier3_amount"] = float(ta3); pp["tier3_pct"] = float(tp3)
                        break
                save_database(db); st.session_state["success_msg"] = "✅ Saved"; st.rerun()

# ============================================================
# ADMIN PRODUCTS
# ============================================================
def render_admin_products():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🛒 All Products</h1>", unsafe_allow_html=True)
    st.markdown("---")
    all_products = get_all_products()
    with st.expander("➕ Naya Product (Single)", expanded=False):
        c1, c2, c3, c4 = st.columns([1.5, 3.5, 1.5, 1])
        with c1: nc = st.text_input("Code:", key="new_prod_code")
        with c2: nn = st.text_input("Name:", key="new_prod_name")
        with c3: npr = st.number_input("Price:", min_value=0.0, step=1.0, key="new_prod_price", value=0.0)
        with c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕ Add", key="btn_add_new_product", type="primary"):
                cs = str(nc).strip(); ns = str(nn).strip()
                if not cs or not ns or npr <= 0: st.error("❌ Sab fields bharo")
                elif cs in {str(p["code"]) for p in all_products}: st.error("❌ Code exists")
                else:
                    db["custom_products"].append({"code": cs, "name": ns, "price": float(npr),
                        "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                    save_database(db); st.session_state["success_msg"] = f"✅ Added"; st.rerun()
    with st.expander("📤 Bulk Upload", expanded=False):
        up = st.file_uploader("Excel/CSV:", type=["xlsx","xls","csv"], key="bulk_upload_file")
        if up:
            try:
                df_u = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                df_u.columns = [str(c).strip().lower() for c in df_u.columns]
                if "code" in df_u.columns and "name" in df_u.columns and "price" in df_u.columns:
                    st.dataframe(df_u.head(10), use_container_width=True)
                    if st.button("📥 Import", key="btn_bulk_import", type="primary"):
                        existing = {str(p["code"]) for p in all_products}
                        added = 0
                        for _, r in df_u.iterrows():
                            try: c_s = str(r["code"]).strip(); n_s = str(r["name"]).strip(); p_f = float(r["price"])
                            except Exception: continue
                            if not c_s or not n_s or p_f <= 0 or c_s in existing: continue
                            db["custom_products"].append({"code": c_s, "name": n_s, "price": p_f,
                                "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                            existing.add(c_s); added += 1
                        save_database(db); st.session_state["success_msg"] = f"✅ {added} imported"; st.rerun()
            except Exception as e: st.error(f"❌ {e}")
    st.markdown(f"### 📋 Products ({len(all_products)})")
    search = st.text_input("🔍 Search:", key="prod_search")
    su = search.strip().upper()
    shown = [p for p in all_products if su in p["name"].upper() or su in str(p["code"])] if su else all_products
    edited_prices = db.get("product_prices", {})
    custom_codes = {str(cp.get("code", "")) for cp in db.get("custom_products", [])}
    for p in shown[:100]:
        code = str(p["code"]); base = float(p["price"]); cur = get_price(code, base)
        is_custom = code in custom_codes; is_edited = code in edited_prices
        c1, c2, c3, c4 = st.columns([1, 4, 2, 2])
        with c1: st.markdown(f"<div style='padding-top:8px;color:#1976d2;font-weight:700;'>{code}</div>", unsafe_allow_html=True)
        with c2:
            marks = (" 🆕" if is_custom else "") + (" ✏️" if is_edited else "")
            st.markdown(f"<div style='padding-top:6px;color:#1976d2;font-weight:600;'>{p['name']}{marks}</div>", unsafe_allow_html=True)
        with c3: np = st.number_input("Price", value=float(cur), min_value=0.0, step=1.0, key=f"price_{code}", label_visibility="collapsed")
        with c4:
            s1, s2 = st.columns(2)
            with s1:
                if st.button("💾", key=f"save_price_{code}"):
                    if float(np) == base: db.get("product_prices", {}).pop(code, None)
                    else: db.setdefault("product_prices", {})[code] = float(np)
                    save_database(db); st.session_state["success_msg"] = "✅ Saved"; st.rerun()
            with s2:
                if is_custom and st.button("🗑", key=f"del_custom_{code}"):
                    db["custom_products"] = [cp for cp in db["custom_products"] if str(cp.get("code", "")) != code]
                    save_database(db); st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()

# ============================================================
# ADMIN BOOKERS
# ============================================================
def render_admin_bookers():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>👤 Bookers</h1>", unsafe_allow_html=True)
    st.markdown("---")
    c1, c2 = st.columns([3, 1])
    with c1: nb = st.text_input("New Booker Name:", key="new_booker_name", label_visibility="collapsed")
    with c2: add = st.button("➕ Add Booker", key="btn_add_booker", type="primary", use_container_width=True)
    if add:
        name = nb.strip()
        if not name: st.error("❌ Name daalo")
        elif name in db.get("bookers", []): st.warning("⚠️ Exists")
        else:
            db.setdefault("bookers", []).append(name); db["bookers"] = sorted(db["bookers"]); save_database(db)
            registry = load_bookers()
            bk_uname = sanitize_username(name)
            if bk_uname not in registry:
                registry[bk_uname] = {
                    "display_name": name, "username": bk_uname, "admin": CURRENT_USER,
                    "password_hash": hash_password("1234"),
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                save_bookers(registry)
            st.session_state["success_msg"] = f"✅ '{name}' added | Default password: 1234"
            st.rerun()

    st.markdown("---")
    st.markdown(f"### 📋 Saved Bookers ({len(db.get('bookers', []))})")
    registry = load_bookers()
    for i, bk in enumerate(db.get("bookers", [])):
        c1, c2, c3 = st.columns([4, 1, 1])
        with c1:
            bk_uname = sanitize_username(bk)
            has_login = bk_uname in registry and registry[bk_uname].get("admin") == CURRENT_USER
            status = "🔐 Login Active" if has_login else "⚠️ No Login"
            st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#1976d2;'>👤 {bk}</b> <span style='color:#666;font-size:12px;margin-left:12px;'>{status}</span></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🔑 Set Password", key=f"setpw_{i}_{bk}", use_container_width=True):
                st.session_state["setpw_booker"] = bk
        with c3:
            if st.button("🗑 Delete", key=f"del_bk_{i}_{bk}", use_container_width=True):
                db["bookers"].remove(bk); save_database(db)
                bk_uname = sanitize_username(bk)
                if bk_uname in registry: del registry[bk_uname]; save_bookers(registry)
                st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()

    target = st.session_state.get("setpw_booker")
    if target:
        st.markdown("---")
        st.markdown(f"### 🔑 Set Password for: **{target}**")
        new_pw = st.text_input("Naya Password:", type="password", key="setpw_new")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save Password", key="setpw_save", type="primary", use_container_width=True):
                if len(new_pw) < 4: st.error("❌ Min 4 chars")
                else:
                    registry = load_bookers()
                    bk_uname = sanitize_username(target)
                    if bk_uname not in registry:
                        registry[bk_uname] = {"display_name": target, "username": bk_uname, "admin": CURRENT_USER,
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                    registry[bk_uname]["password_hash"] = hash_password(new_pw)
                    registry[bk_uname]["admin"] = CURRENT_USER
                    registry[bk_uname]["display_name"] = target
                    save_bookers(registry)
                    st.session_state["setpw_booker"] = None
                    st.session_state["success_msg"] = f"✅ Password set for {target}"
                    st.rerun()
        with c2:
            if st.button("❌ Cancel", key="setpw_cancel", use_container_width=True):
                st.session_state["setpw_booker"] = None; st.rerun()

# ============================================================
# ADMIN SALESMEN
# ============================================================
def render_admin_salesmen():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🧑‍💼 Salesmen</h1>", unsafe_allow_html=True)
    st.markdown("---")
    c1, c2 = st.columns([3, 1])
    with c1: ns = st.text_input("New Salesman:", key="new_salesman_name", label_visibility="collapsed")
    with c2: add = st.button("➕ Add", key="btn_add_salesman", type="primary", use_container_width=True)
    if add:
        name = ns.strip()
        if not name: st.error("❌ Name")
        elif name in db.get("salesmen", []): st.warning("⚠️ Exists")
        else:
            db.setdefault("salesmen", []).append(name); db["salesmen"] = sorted(db["salesmen"]); save_database(db)
            st.session_state["success_msg"] = f"✅ Added"; st.rerun()
    st.markdown("---")
    for i, sm in enumerate(db.get("salesmen", [])):
        c1, c2 = st.columns([5, 1])
        with c1: st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#1976d2;'>🧑‍💼 {sm}</b></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑", key=f"del_sm_{i}_{sm}"):
                db["salesmen"].remove(sm); save_database(db)
                st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()

# ============================================================
# ADMIN SALARIES
# ============================================================
def render_admin_salaries(role_type):
    if role_type == "bookers":
        title = "💰 Bookers Salary"; emoji = "👤"; names = db.get("bookers", []); sal_key = "bookers_salaries"
    else:
        title = "💰 Salesmen Salary"; emoji = "🧑‍💼"; names = db.get("salesmen", []); sal_key = "salesmen_salaries"
    st.markdown(f"<h1 style='color:#1976d2 !important;'>{title}</h1>", unsafe_allow_html=True)
    st.markdown("---")
    if sal_key not in db: db[sal_key] = {}
    if not names: st.info("Pehle add karo."); return
    for p_name in names:
        sd = db[sal_key].get(p_name, {"base_salary": 0, "transactions": []})
        base = sd.get("base_salary", 0); txns = sd.get("transactions", [])
        adv_p = sum(t["amount"] for t in txns if t.get("type")=="advanced" and t.get("status","pending")=="pending")
        short_p = sum(t["amount"] for t in txns if t.get("type")=="shortage" and t.get("status","pending")=="pending")
        remaining = base - adv_p - short_p
        with st.expander(f"💰 {emoji} {p_name} — Remaining: Rs {remaining:,.0f}"):
            c1, c2 = st.columns([3, 1])
            with c1: nb = st.number_input("Base Salary:", value=float(base), min_value=0.0, step=500.0, key=f"base_{role_type}_{p_name}")
            with c2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Save", key=f"savebase_{role_type}_{p_name}"):
                    db[sal_key].setdefault(p_name, {"base_salary": 0, "transactions": []})
                    db[sal_key][p_name]["base_salary"] = float(nb); save_database(db)
                    st.session_state["success_msg"] = "✅ Saved"; st.rerun()
            st.markdown(f"""<span class='sal-metric base'>Base: Rs {base:,.0f}</span>
                <span class='sal-metric adv'>Adv Pend: Rs {adv_p:,.0f}</span>
                <span class='sal-metric short'>Short Pend: Rs {short_p:,.0f}</span>
                <span class='sal-metric remain'>Remaining: Rs {remaining:,.0f}</span>""", unsafe_allow_html=True)
            st.markdown("**Add Transaction**")
            c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
            with c1: td = st.date_input("Date", value=date.today(), key=f"td_{role_type}_{p_name}")
            with c2: tt = st.selectbox("Type", ["Advanced", "Shortage"], key=f"tt_{role_type}_{p_name}")
            with c3: ta = st.number_input("Amount", min_value=0.0, step=100.0, key=f"ta_{role_type}_{p_name}")
            with c4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕", key=f"addtxn_{role_type}_{p_name}"):
                    if ta <= 0: st.session_state["error_msg"] = "❌ >0"
                    else:
                        db[sal_key].setdefault(p_name, {"base_salary": base, "transactions": []})
                        nid = max([t.get("id", 0) for t in txns] + [0]) + 1
                        db[sal_key][p_name]["transactions"].append({
                            "id": nid, "date": td.strftime("%d-%m-%Y"),
                            "time": datetime.now().strftime("%H:%M"),
                            "type": "advanced" if tt == "Advanced" else "shortage",
                            "amount": float(ta), "note": "", "status": "pending",
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                        save_database(db); st.session_state["success_msg"] = "✅"; st.rerun()
            for t in sorted(txns, key=lambda x: x.get("created_at", ""), reverse=True)[:10]:
                c1, c2, c3 = st.columns([4, 1, 1])
                with c1: st.markdown(f"📅 {t.get('date','')} · {t.get('type','')} · **Rs {t['amount']:,.0f}** · {t.get('status','').upper()}")
                with c2:
                    if t.get("status") == "pending" and st.button("✅ Paid", key=f"pd_{role_type}_{p_name}_{t['id']}"):
                        for tx in db[sal_key][p_name]["transactions"]:
                            if tx.get("id") == t["id"]: tx["status"] = "paid"; break
                        save_database(db); st.session_state["success_msg"] = "✅"; st.rerun()
                with c3:
                    if st.button("🗑", key=f"dt_{role_type}_{p_name}_{t['id']}"):
                        db[sal_key][p_name]["transactions"] = [x for x in db[sal_key][p_name]["transactions"] if x.get("id") != t["id"]]
                        save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()

# ============================================================
# ADMIN DAILY EXPENSE
# ============================================================
def render_admin_expense():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>💵 Daily Expense</h1>", unsafe_allow_html=True)
    st.markdown("---")
    if "petrol_expenses" not in db: db["petrol_expenses"] = []
    if "lunch_expenses" not in db: db["lunch_expenses"] = []
    total_petrol = sum(float(x.get("amount", 0)) for x in db["petrol_expenses"])
    total_lunch = sum(float(x.get("amount", 0)) for x in db["lunch_expenses"])
    st.markdown(f"<div class='summary-box'><b>Petrol:</b> Rs {total_petrol:,.0f} | <b>Lunch:</b> Rs {total_lunch:,.0f} | <b>Total:</b> Rs {total_petrol+total_lunch:,.0f}</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**⛽ Petrol**")
        c1a, c1b, c1c = st.columns([2, 2, 1])
        with c1a: ps = st.selectbox("Salesman:", ["-- Select --"] + db.get("salesmen", []), key="ps")
        with c1b: pa = st.number_input("Amount", min_value=0.0, step=50.0, key="pa")
        with c1c:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕", key="add_petrol"):
                if ps == "-- Select --" or pa <= 0: st.session_state["error_msg"] = "❌"
                else:
                    nid = max([x.get("id", 0) for x in db["petrol_expenses"]] + [0]) + 1
                    db["petrol_expenses"].append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "time": datetime.now().strftime("%H:%M"), "salesman": ps, "amount": float(pa)})
                    save_database(db); st.session_state["success_msg"] = "✅"; st.rerun()
    with c2:
        st.markdown("**🍽️ Lunch**")
        c2a, c2b = st.columns([3, 1])
        with c2a: la = st.number_input("Amount", min_value=0.0, step=50.0, key="la")
        with c2b:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕", key="add_lunch"):
                if la <= 0: st.session_state["error_msg"] = "❌"
                else:
                    nid = max([x.get("id", 0) for x in db["lunch_expenses"]] + [0]) + 1
                    db["lunch_expenses"].append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "time": datetime.now().strftime("%H:%M"), "amount": float(la)})
                    save_database(db); st.session_state["success_msg"] = "✅"; st.rerun()

# ============================================================
# ADMIN BILLING
# ============================================================
def render_admin_billing():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🧾 Billing</h1>", unsafe_allow_html=True)
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1: st.text_input("Bill No:", value=str(db["next_bill_no"]), disabled=True)
    with c2: st.text_input("Date:", value=datetime.now().strftime("%d-%m-%Y"), disabled=True)
    c1, c2, c3 = st.columns(3)
    with c1: shop = st.text_input("Shop:", key="shop_name")
    with c2:
        bks = db.get("bookers", [])
        if bks:
            sel = st.selectbox("Order Booker:", ["-- Select --"] + bks, key="order_booker_select")
            st.session_state["order_booker"] = "" if sel == "-- Select --" else sel
        else: st.text_input("Order Booker:", key="order_booker")
    with c3:
        sms = db.get("salesmen", [])
        if sms:
            sel = st.selectbox("Salesman:", ["-- Select --"] + sms, key="salesman_select")
            st.session_state["salesman"] = "" if sel == "-- Select --" else sel
        else: st.text_input("Salesman:", key="salesman")
    c1, c2 = st.columns([1, 3])
    with c1: s = st.text_input("🔍 Search:", key="search_text")
    with c2:
        all_p = get_all_products()
        su = s.strip().upper()
        f = [p["name"] for p in all_p if su in p["name"].upper()] if su else [p["name"] for p in all_p]
        if st.session_state.get("product_sel") and st.session_state["product_sel"] not in f:
            st.session_state["product_sel"] = ""
        ps = st.selectbox("Product:", [""] + f, key="product_sel")
    sel_p = next((p for p in get_all_products() if p["name"] == ps), None)
    tp_def = get_price(sel_p["code"], sel_p["price"]) if sel_p else 0.0
    if st.session_state["_prev_prod"] != ps:
        st.session_state["tp_box"] = float(tp_def); st.session_state["_prev_prod"] = ps
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: bx = st.number_input("Boxes:", min_value=0, step=1, key="boxes")
    with c2: tp = st.number_input("TP/Box:", min_value=0.0, step=1.0, key="tp_box")
    with c3: dc = st.number_input("Disc %:", min_value=0.0, step=0.5, key="discount")
    g = bx * tp; n = g - (g * dc / 100)
    with c4: st.text_input("Gross:", value=f"{g:.0f}", disabled=True)
    with c5: st.text_input("Net:", value=f"{n:.0f}", disabled=True)
    b1, b2 = st.columns(2)
    with b1: st.button("➕ Add Bill", key="btn_add", on_click=add_bill_callback, type="primary", use_container_width=True)
    with b2: st.button("🔄 Refresh", key="btn_refresh", on_click=refresh_callback, use_container_width=True)
    e1, e2, e3 = st.columns(3)
    with e1: st.button("📄 Export Bill", key="btn_export", on_click=export_bill_callback, use_container_width=True)
    with e2: st.button("📦 Export Load Form", key="btn_export_lf", on_click=export_load_form_from_billing_callback, use_container_width=True)
    with e3: st.button("🔄 Refresh Load Form", key="btn_load_refresh", on_click=refresh_load_form_callback, use_container_width=True)
    show_auto_download()

# ============================================================
# ADMIN BILLS LIST
# ============================================================
def render_admin_bills_list():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📋 Bills List</h1>", unsafe_allow_html=True)
    st.markdown("---")
    if not db["bills"]: st.info("Koi bill nahi."); return
    credit_map = {(cb.get("shop",""), cb.get("date",""), cb.get("booker",""), cb.get("bill_no","")): cb
                  for cb in db.get("credit_bills", [])}
    groups = {}
    for bi, b in enumerate(db["bills"]):
        key = (b.get("Shop",""), b.get("Date",""), b.get("Order Booker",""), b.get("Bill No",""))
        if key not in groups:
            groups[key] = {"shop": b.get("Shop",""), "date": b.get("Date",""), "booker": b.get("Order Booker",""),
                "salesman": b.get("Salesman",""), "bill_no": b.get("Bill No",""), "items": [], "idx": []}
        groups[key]["items"].append({"Code": b.get("Code"), "Product": b.get("Product"),
            "Boxes": b.get("Boxes"), "TP/Box": b.get("TP/Box"), "Discount %": b.get("Discount %"),
            "Gross": b.get("Gross"), "Net": b.get("Net")})
        groups[key]["idx"].append(bi)
    st.markdown(f"### 📋 Bills ({len(groups)})")
    for key, g in sorted(groups.items(), key=lambda kv: kv[0][1], reverse=True):
        shop = g["shop"] or "-"; dt = g["date"] or "-"; bk = g["booker"] or "-"; sm = g["salesman"] or "-"; bn = g["bill_no"]
        tb = sum(int(it.get("Boxes",0)) for it in g["items"])
        ck = (g["shop"] or "", g["date"] or "", g["booker"] or "", g["bill_no"] or "")
        ec = credit_map.get(ck)
        cbadge = ""
        if ec: cbadge = '<span class="badge-paid-credit">✅ Credit PAID</span>' if ec.get("status")=="paid" else '<span class="badge-credit-tag">💳 Credit PEND</span>'
        st.markdown(f"""<div class='lf-simple-card'>
            <div class='lf-info'><div class='lf-line1'>🏪 {shop} {cbadge}</div>
            <div class='lf-line2'>📅 {dt} · 👤 {bk} · 🧑‍💼 {sm} · 🧾 #{bn}</div></div>
            <div class='lf-boxes'>{tb}<small>BOXES</small></div>
        </div>""", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            if st.button(f"⬇️ Excel", key=f"dl_{key}_{dt}"):
                export_single_group_bill(shop, dt, bk, sm, g["items"], bn); st.rerun()
        with c2:
            if ec: st.button("💳 Already", key=f"cr_{key}_{dt}", disabled=True)
            else:
                if st.button("💳 Credit", key=f"cr_{key}_{dt}"):
                    move_group_to_credit(g); st.rerun()
        with c3:
            if st.button("🗑 Delete", key=f"del_{key}_{dt}"):
                idxs = set(g["idx"])
                db["bills"] = [b for i, b in enumerate(db["bills"]) if i not in idxs]
                save_database(db); st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()
        st.markdown("---")
    show_auto_download()

# ============================================================
# ADMIN CREDIT BILLS
# ============================================================
def render_admin_credit():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>💳 Credit Bills</h1>", unsafe_allow_html=True)
    st.markdown("---")
    cbs = db.get("credit_bills", [])
    if not cbs: st.info("Koi credit bill nahi."); return
    pend = [c for c in cbs if c.get("status") == "pending"]
    paid = [c for c in cbs if c.get("status") == "paid"]
    st.markdown(f"<div class='summary-box'><b>Pending:</b> {len(pend)} (Rs {sum(float(c.get('total_net',0)) for c in pend):,.0f}) | <b>Paid:</b> {len(paid)} (Rs {sum(float(c.get('total_net',0)) for c in paid):,.0f})</div>", unsafe_allow_html=True)
    for c in pend + paid:
        cid = c["id"]; is_paid = c.get("status") == "paid"
        card = "credit-paid" if is_paid else "credit-pending"
        badge = "✅ PAID" if is_paid else "⏳ PENDING"
        st.markdown(f"""<div class='lf-simple-card {card}'>
            <div class='lf-info'><div class='lf-line1'>🏪 {c.get('shop','')} · {badge}</div>
            <div class='lf-line2'>📅 {c.get('date','')} · 👤 {c.get('booker','')} · 💰 Rs {float(c.get('total_net',0)):,.0f}</div></div>
            <div class='lf-boxes'>{c.get('total_boxes',0)}<small>BOXES</small></div>
        </div>""", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if not is_paid:
                if st.button("✅ Mark Paid", key=f"pay_{cid}", type="primary"):
                    for cc in db["credit_bills"]:
                        if cc.get("id") == cid: cc["status"]="paid"; cc["paid_at"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S"); break
                    save_database(db); st.session_state["success_msg"] = "✅ Paid"; st.rerun()
            else:
                if st.button("↩️ Pending", key=f"unpay_{cid}"):
                    for cc in db["credit_bills"]:
                        if cc.get("id") == cid: cc["status"]="pending"; cc["paid_at"]=None; break
                    save_database(db); st.session_state["success_msg"] = "↩️"; st.rerun()
        with c2:
            if st.button("🗑 Delete", key=f"del_credit_{cid}"):
                db["credit_bills"] = [x for x in db["credit_bills"] if x.get("id") != cid]
                save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()

# ============================================================
# ADMIN LOAD FORM
# ============================================================
def render_admin_load_form():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📦 Load Forms</h1>", unsafe_allow_html=True)
    st.markdown("---")
    lfs = db.get("load_forms", [])
    if not lfs: st.info("Koi load form nahi."); return
    for i, lf in enumerate(sorted(lfs, key=lambda x: x.get("created_at",""), reverse=True)):
        lf_id = lf.get("id", i); bk = lf.get("booker", "?"); dt = lf.get("date", "")
        tb = lf.get("total_boxes", 0); items = lf.get("items", [])
        trf = lf.get("transferred_to_dsr", False); ref = lf.get("refreshed", False)
        src = lf.get("source", "admin")
        badge = f'✅ DSR #{lf.get("dsr_id")}' if trf else ('🔄 Refreshed' if ref else ('👤 Booker' if src == "booker" else '🟢 Active'))
        st.markdown(f"""<div class='lf-simple-card'>
            <div class='lf-info'><div class='lf-line1'>👤 {bk} <span class='badge-pending'>{badge}</span></div>
            <div class='lf-line2'>📅 {dt} · 📦 {len(items)} products</div></div>
            <div class='lf-boxes'>{tb}<small>BOXES</small></div>
        </div>""", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if not trf:
                if st.button("📤 Transfer to DSR", key=f"trf_{lf_id}"):
                    transfer_load_form_to_dsr(lf_id); st.rerun()
        with c2:
            if st.button("🗑 Delete", key=f"del_lf_{lf_id}"):
                db["load_forms"] = [x for x in db["load_forms"] if x.get("id") != lf_id]
                save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()
        st.markdown("---")

# ============================================================
# ADMIN DSR
# ============================================================
def render_admin_dsr():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📋 DSR</h1>", unsafe_allow_html=True)
    st.markdown("---")
    dsrs = db.get("dsr_forms", [])
    if not dsrs: st.info("Koi DSR nahi."); return
    for i, d in enumerate(sorted(dsrs, key=lambda x: x.get("created_at",""), reverse=True)):
        did = d.get("id", i); bk = d.get("booker", "?"); dt = d.get("date", "")
        tb = d.get("total_boxes", 0); tot = float(d.get("total_amount", 0))
        ct, _, _ = get_dsr_credit_info(d)
        fin = float(d.get("amount_to_collect", 0)) - ct
        st.markdown(f"""<div class='lf-simple-card dsr'>
            <div class='lf-info'><div class='lf-line1'>📋 DSR #{did} — {bk}</div>
            <div class='lf-line2'>📅 {dt} · 💰 Rs {tot:,.0f} · 💳 Rs {ct:,.0f} · <b style="color:#c62828;">FINAL: Rs {fin:,.0f}</b></div></div>
            <div class='lf-boxes'>{tb}<small>BOXES</small></div>
        </div>""", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📤 Excel", key=f"ex_dsr_{did}", type="primary"):
                export_dsr_excel(d); st.rerun()
        with c2:
            if st.button("🗑 Delete", key=f"del_dsr_{did}"):
                db["dsr_forms"] = [x for x in db["dsr_forms"] if x.get("id") != did]
                for x in db.get("load_forms", []):
                    if x.get("dsr_id") == did: x["transferred_to_dsr"] = False; x.pop("dsr_id", None)
                save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()
        st.markdown("---")

# ============================================================
# ADMIN CALCULATION
# ============================================================
def render_admin_calculation():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🧮 Cash Calculation</h1>", unsafe_allow_html=True)
    st.markdown("---")
    if "calc_reset_token" not in st.session_state: st.session_state["calc_reset_token"] = 0
    total = 0
    for section, denoms in [("💵 Notes", [5000, 1000, 500, 100, 50, 20, 10]), ("🪙 Coins", [5, 2, 1])]:
        st.markdown(f"### {section}")
        for d in denoms:
            c1, c2, c3 = st.columns([2, 2, 3])
            with c1: st.markdown(f"**₹{d}**")
            with c2: q = st.number_input(f"Qty", min_value=0, step=1, value=0, key=f"calc_{d}_{st.session_state['calc_reset_token']}", label_visibility="collapsed")
            with c3: amt = q * d; total += amt; st.markdown(f"**Rs {amt:,.0f}**")
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💰 Calculate", type="primary", use_container_width=True):
            st.session_state["calc_locked"] = total
    with c2:
        if st.button("🔄 Reset", use_container_width=True):
            st.session_state["calc_reset_token"] += 1; st.session_state["calc_locked"] = 0; st.rerun()
    st.markdown(f"<div style='background:linear-gradient(135deg,#1e3a8a,#2563eb);border-radius:16px;padding:24px;text-align:center;color:white;margin-top:16px;'><div style='font-size:12px;letter-spacing:2px;font-weight:700;'>LIVE TOTAL</div><div style='font-size:42px;font-weight:800;'>Rs {total:,.0f}</div></div>", unsafe_allow_html=True)
    locked = st.session_state.get("calc_locked", 0)
    if locked and locked > 0:
        st.markdown(f"<div style='background:linear-gradient(135deg,#065f46,#10b981);border-radius:16px;padding:24px;text-align:center;color:white;margin-top:10px;'><div style='font-size:12px;letter-spacing:2px;font-weight:700;'>🔒 LOCKED TOTAL</div><div style='font-size:42px;font-weight:800;'>Rs {locked:,.0f}</div></div>", unsafe_allow_html=True)

# ============================================================
# CALLBACKS
# ============================================================
def add_bill_callback():
    ps = st.session_state.get("product_sel", "")
    bx = st.session_state.get("boxes", 0)
    tp = st.session_state.get("tp_box", 0.0)
    dc = st.session_state.get("discount", 0.0)
    if not ps: st.session_state["error_msg"] = "❌ Select Product"; return
    if bx <= 0: st.session_state["error_msg"] = "❌ Enter Boxes"; return
    sel = next((p for p in get_all_products() if p["name"] == ps), None)
    if not sel: st.session_state["error_msg"] = "❌ Invalid"; return
    g = bx * tp; n = g - (g * dc / 100)
    db["bills"].append({"Bill No": db["next_bill_no"], "Date": datetime.now().strftime("%d-%m-%Y"),
        "Shop": st.session_state.get("shop_name","").strip(),
        "Order Booker": st.session_state.get("order_booker","").strip(),
        "Salesman": st.session_state.get("salesman","").strip(),
        "Delivery Man": "", "Code": sel["code"], "Product": sel["name"],
        "Boxes": bx, "TP/Box": tp, "Discount %": dc, "Gross": g, "Net": n})
    save_database(db)
    st.session_state["last_bill_no"] = db["next_bill_no"]
    st.session_state["success_msg"] = f"✅ Bill #{db['next_bill_no']}"
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None

def refresh_callback():
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None
    st.session_state["success_msg"] = "✅"

def _build_bill_excel(shop, bills, pkg_pct, pkg_name, tier_label):
    output = BytesIO(); wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Bill")
    ws.set_paper(9); ws.set_portrait(); ws.fit_to_pages(1, 1)
    for col, w in [("A",42.86),("B",12.71),("C",10.71),("D",10.71),("E",11.71),("F",11.14),("G",11.14),("H",12.14),("I",13.14),("J",13.14)]:
        ws.set_column(f"{col}:{col}", w)
    title = wb.add_format({"bold":True, "font_size":18, "align":"center", "border":2})
    hdr = wb.add_format({"bold":True, "font_size":11, "bg_color":"#BBDEFB", "align":"center", "border":2, "text_wrap":True})
    cl = wb.add_format({"font_size":12, "border":1, "align":"left"})
    cc = wb.add_format({"font_size":12, "border":1, "align":"center"})
    tt = wb.add_format({"bold":True, "font_size":12, "bg_color":"#FFF2CC", "align":"center", "border":2})
    dl = wb.add_format({"font_size":12, "border":1, "align":"center", "bg_color":"#E8F5E9", "bold":True})
    pi = wb.add_format({"font_size":10, "italic":True, "align":"left", "font_color":"#1b5e20"})
    ws.merge_range("A1:J1", COMPANY_NAME, title)
    ws.write("A3","Shop",hdr); ws.write("B3",shop,cc)
    ws.write("D3","Booker",hdr); ws.write("E3", bills[0].get("Order Booker","") if bills else "", cc)
    ws.write("G3","Bill No",hdr); ws.write("H3", bills[0].get("Bill No","") if bills else "", cc)
    ws.write("I3","Date",hdr); ws.write("J3", bills[0].get("Date","") if bills else "", cc)
    if pkg_pct > 0: ws.merge_range("A4:J4", f"🎁 Discount: {pkg_name} | {tier_label} | {pkg_pct}%", pi)
    else: ws.merge_range("A4:J4", "💡 No discount", pi)
    hr = 6
    for col, h in enumerate(["Product","Code","Boxes","TP/Box","Gross","Disc %","Net","Pkg Disc %","After Disc Net","Saved"]):
        ws.write(hr, col, h, hdr)
    row = hr + 1; gt = 0; tb = 0; nt = 0; ad = 0; sv = 0
    for b in bills:
        bn = float(b.get("Net",0)); bg = float(b.get("Gross",0)); bb = int(b.get("Boxes",0))
        an = bn - (bn * pkg_pct / 100); s = bn - an
        ws.write(row,0,b["Product"],cl); ws.write(row,1,b["Code"],cc); ws.write(row,2,bb,cc)
        ws.write(row,3,b.get("TP/Box",0),cc); ws.write(row,4,bg,cc); ws.write(row,5,b.get("Discount %",0),cc); ws.write(row,6,bn,cc)
        if pkg_pct > 0: ws.write(row,7,pkg_pct,dl); ws.write(row,8,an,dl); ws.write(row,9,s,dl)
        else: ws.write(row,7,0,cc); ws.write(row,8,bn,cc); ws.write(row,9,0,cc)
        gt += bg; tb += bb; nt += bn; ad += an; sv += s; row += 1
    ws.write(row,1,"TOTAL",tt); ws.write(row,2,tb,tt); ws.write(row,4,gt,tt); ws.write(row,6,nt,tt)
    ws.write(row,8,ad,tt); ws.write(row,9,sv,tt)
    row += 2
    ws.merge_range(row,0,row,6,"NET AMOUNT",hdr); ws.merge_range(row,7,row,9,f"Rs {ad:,.0f}",tt)
    wb.close(); output.seek(0)
    return output.getvalue()

def export_single_group_bill(shop, date_str, booker, salesman, items, bill_no):
    tot = sum(float(it.get("Net",0)) for it in items)
    pct, pn, tl = get_effective_discount_pct(shop, tot)
    fname = f"{shop}_{date_str.replace('-','')}.xlsx".replace("/","-").replace(" ","_").replace(":","")
    st.session_state["download_file"] = (fname, _build_bill_excel(shop, items, pct, pn, tl))
    st.session_state["success_msg"] = f"✅ Excel ready"

def export_bill_callback():
    if not db["bills"]: st.session_state["error_msg"] = "❌ No Bills"; return
    shop = st.session_state.get("shop_name", "").strip() or "Bill"
    for ch in ['\\','/',':','*','?','"','<','>','|']: shop = shop.replace(ch, "")
    sb = [b for b in db["bills"] if b["Shop"].strip() == shop]
    if not sb: st.session_state["error_msg"] = "❌ No bills"; return
    tot = sum(float(b.get("Net",0)) for b in sb)
    pct, pn, tl = get_effective_discount_pct(shop, tot)
    st.session_state["download_file"] = (f"{shop}.xlsx", _build_bill_excel(shop, sb, pct, pn, tl))
    db["next_bill_no"] += 1; save_database(db)
    st.session_state["success_msg"] = f"✅ Exported"

def export_load_form_from_billing_callback():
    bk = st.session_state.get("order_booker", "").strip()
    if not bk: st.session_state["error_msg"] = "❌ Select Booker"; return
    bbs = [b for b in db["bills"] if b["Order Booker"].strip() == bk]
    if not bbs: st.session_state["error_msg"] = "❌ No bills"; return
    sm = {}
    for b in bbs:
        c = b["Code"]
        if c not in sm: sm[c] = {"Code": c, "Product": b["Product"], "Boxes": 0}
        sm[c]["Boxes"] += b["Boxes"]
    items = list(sm.values()); tb = sum(it["Boxes"] for it in items)
    existing = None
    for lf in reversed(db.get("load_forms", [])):
        if (str(lf.get("booker","")).strip() == bk and not lf.get("transferred_to_dsr") and not lf.get("refreshed")):
            existing = lf; break
    if existing:
        existing["items"] = items; existing["total_boxes"] = tb; existing["total_products"] = len(items)
        existing["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        msg = f"🔄 Load Form #{existing['id']} updated"
    else:
        nid = max([lf.get("id",0) for lf in db.get("load_forms", [])] + [0]) + 1
        db.setdefault("load_forms", []).append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"),
            "time": datetime.now().strftime("%H:%M"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "booker": bk, "items": items, "total_boxes": tb, "total_products": len(items),
            "transferred_to_dsr": False, "refreshed": False})
        msg = f"✅ New Load Form #{nid}"
    save_database(db); st.session_state["success_msg"] = msg

def refresh_load_form_callback():
    bk = st.session_state.get("order_booker", "").strip()
    if not bk: st.session_state["error_msg"] = "❌ Enter Booker"; return
    db["bills"] = [b for b in db["bills"] if b["Order Booker"].strip() != bk]
    for lf in db.get("load_forms", []):
        if (str(lf.get("booker","")).strip() == bk and not lf.get("transferred_to_dsr") and not lf.get("refreshed")):
            lf["refreshed"] = True; lf["refreshed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_database(db)
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None
    st.session_state["success_msg"] = f"✅ Cleared"

def move_group_to_credit(group):
    shop = group.get("shop","") or "-"; dt = group.get("date","") or "-"
    bk = group.get("booker","") or "-"; sm = group.get("salesman","") or "-"; bn = group.get("bill_no","")
    items = group.get("items", [])
    for cb in db.get("credit_bills", []):
        if (cb.get("shop")==shop and cb.get("date")==dt and cb.get("booker")==bk and cb.get("bill_no")==bn):
            st.session_state["error_msg"] = "⚠️ Already in Credit"; return
    tb = sum(int(it.get("Boxes",0)) for it in items); tg = sum(float(it.get("Gross",0)) for it in items)
    tn = sum(float(it.get("Net",0)) for it in items)
    nid = db.get("next_credit_id", 1)
    db.setdefault("credit_bills", []).append({"id": nid, "shop": shop, "date": dt, "booker": bk, "salesman": sm,
        "bill_no": bn, "items": [dict(it) for it in items], "total_boxes": tb, "total_gross": tg, "total_net": tn,
        "status": "pending", "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "paid_at": None})
    db["next_credit_id"] = nid + 1; save_database(db)
    st.session_state["success_msg"] = f"💳 Credit #{nid}"

def transfer_load_form_to_dsr(lf_id):
    lf = next((x for x in db.get("load_forms", []) if x.get("id") == lf_id), None)
    if not lf or lf.get("transferred_to_dsr"): st.session_state["error_msg"] = "❌"; return
    bk = lf.get("booker",""); items = lf.get("items", []); allp = get_all_products()
    dsr_items = []; tb = 0; ta = 0.0
    for it in items:
        c = str(it.get("Code",""))
        pr = next((p for p in allp if str(p["code"]) == c), None)
        try: price = get_price(c, pr["price"]) if pr else 0.0
        except Exception: price = 0.0
        bx = int(it.get("Boxes",0)); lt = bx * float(price)
        dsr_items.append({"Code": c, "Product": it.get("Product",""), "Boxes": bx, "TP/Box": float(price),
            "Total": lt, "ReturnBoxes": 0, "ReturnAmount": 0.0})
        tb += bx; ta += lt
    bbs = [b for b in db.get("bills", []) if b.get("Order Booker","").strip() == bk]
    sgs = {}; src = []
    for b in bbs:
        s = b.get("Shop","-") or "-"; bn = b.get("Bill No","")
        net = float(b.get("Net",0)); gr = float(b.get("Gross",0))
        if s not in sgs: sgs[s] = {"shop": s, "gross": 0.0, "net": 0.0, "individual_discount": 0.0, "bill_nos": set()}
        sgs[s]["gross"] += gr; sgs[s]["net"] += net; sgs[s]["individual_discount"] += (gr - net)
        sgs[s]["bill_nos"].add(bn); src.append({"shop": s, "bill_no": bn, "net": net})
    sds = []; td = 0.0
    for s, g in sgs.items():
        pct, pn, _ = get_effective_discount_pct(s, g["net"])
        pd = g["net"] * pct / 100; std = g["individual_discount"] + pd
        sds.append({"shop": s, "gross": g["gross"], "net": g["net"], "individual_discount": g["individual_discount"],
            "package_pct": pct, "package_name": pn or "", "package_discount": pd,
            "total_discount": std, "bill_nos": sorted([str(x) for x in g["bill_nos"]])})
        td += std
    nid = max([x.get("id",0) for x in db.get("dsr_forms", [])] + [0]) + 1
    db.setdefault("dsr_forms", []).append({"id": nid, "source_load_form_id": lf_id, "booker": bk,
        "date": lf.get("date",""), "time": lf.get("time",""),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "items": dsr_items, "total_boxes": tb, "total_amount": ta,
        "total_return_boxes": 0, "total_return_amount": 0.0, "net_amount": ta,
        "shop_discounts": sds, "total_discount": td, "amount_to_collect": ta - td,
        "source_bills": src, "status": "open", "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    for x in db["load_forms"]:
        if x.get("id") == lf_id: x["transferred_to_dsr"] = True; x["dsr_id"] = nid; break
    save_database(db); st.session_state["success_msg"] = f"✅ DSR #{nid}"

def export_dsr_excel(dsr):
    ct, cs, _ = get_dsr_credit_info(dsr)
    fin = float(dsr.get("amount_to_collect", 0)) - ct
    output = BytesIO(); wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("DSR")
    title = wb.add_format({"bold":True, "font_size":16, "align":"center", "border":2, "bg_color":"#FFE0B2"})
    header = wb.add_format({"bold":True, "font_size":11, "bg_color":"#BBDEFB", "align":"center", "border":2})
    cl = wb.add_format({"font_size":11, "border":1, "align":"left"})
    cc = wb.add_format({"font_size":11, "border":1, "align":"center"})
    cn = wb.add_format({"font_size":11, "border":1, "align":"right", "num_format":"#,##0"})
    tot = wb.add_format({"bold":True, "font_size":11, "bg_color":"#FFF2CC", "align":"center", "border":2})
    hi = wb.add_format({"bold":True, "font_size":12, "bg_color":"#E8F5E9", "align":"center", "border":2})
    for col, w in [("A",10),("B",40),("C",10),("D",12),("E",14),("F",10),("G",14)]: ws.set_column(f"{col}:{col}", w)
    ws.merge_range("A1:G1", f"{COMPANY_NAME} — DSR #{dsr['id']}", title)
    ws.write("A3","Booker",header); ws.write("B3", dsr.get("booker",""), cl)
    ws.write("C3","Date",header); ws.write("D3", dsr.get("date",""), cc)
    ws.write("E3","Time",header); ws.write("F3", dsr.get("time",""), cc)
    ws.write("A5","Code",header); ws.write("B5","Product",header); ws.write("C5","Boxes",header)
    ws.write("D5","TP/Box",header); ws.write("E5","Total",header)
    row = 5
    for it in dsr.get("items", []):
        ws.write(row,0,str(it.get("Code","")),cc); ws.write(row,1,it.get("Product",""),cl)
        ws.write(row,2,int(it.get("Boxes",0)),cc); ws.write(row,3,float(it.get("TP/Box",0)),cn)
        ws.write(row,4,float(it.get("Total",0)),cn); row += 1
    ws.write(row,2,"TOTAL",tot); ws.write(row,4,float(dsr.get("total_amount",0)),tot)
    row += 2
    ws.merge_range(row,0,row,4,"Stock Value",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('total_amount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(−) Returns",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('total_return_amount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(=) Net Stock",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('net_amount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(−) Discount",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('total_discount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(=) Ye Lena Hai",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('amount_to_collect',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(−) Credit Bills",header); ws.merge_range(row,5,row,6,f"Rs {ct:,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"💰 FINAL YE LENA HAI",hi); ws.merge_range(row,5,row,6,f"Rs {fin:,.0f}",hi)
    wb.close(); output.seek(0)
    fname = f"DSR_{dsr['id']}_{dsr.get('booker','')}.xlsx".replace("/","-").replace(" ","_")
    st.session_state["download_file"] = (fname, output.getvalue())
    st.session_state["success_msg"] = f"✅ DSR #{dsr['id']} Excel ready"

# ============================================================
# RENDER ADMIN PAGE
# ============================================================
if st.session_state["page"] == "📊 Dashboard": render_admin_dashboard()
elif st.session_state["page"] == "🧾 Billing": render_admin_billing()
elif st.session_state["page"] == "🛒 All Products": render_admin_products()
elif st.session_state["page"] == "🎁 Discount": render_admin_discount()
elif st.session_state["page"] == "👤 Bookers": render_admin_bookers()
elif st.session_state["page"] == "💰 Bookers Salary": render_admin_salaries("bookers")
elif st.session_state["page"] == "🧑‍💼 Salesmen": render_admin_salesmen()
elif st.session_state["page"] == "💰 Salesmen Salary": render_admin_salaries("salesmen")
elif st.session_state["page"] == "💵 Daily Expense": render_admin_expense()
elif st.session_state["page"] == "📋 Bills List": render_admin_bills_list()
elif st.session_state["page"] == "💳 Credit Bills": render_admin_credit()
elif st.session_state["page"] == "📦 Load Form": render_admin_load_form()
elif st.session_state["page"] == "📋 DSR": render_admin_dsr()
elif st.session_state["page"] == "🧮 Calculation": render_admin_calculation()

if st.session_state.get("success_msg"):
    st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
if st.session_state.get("error_msg"):
    st.error(st.session_state["error_msg"]); st.session_state["error_msg"] = None
show_auto_download()
