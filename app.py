# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# Full code with Premium Login (Waves), Admin + Booker login
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

def get_booker_draft_file(booker_name):
    return f"booker_draft_{sanitize_username(booker_name)}.json"

def default_blank_db():
    return {
        "next_bill_no": 1, "bills": [], "bookers": [], "salesmen": [],
        "load_forms": [], "bookers_salaries": {}, "salesmen_salaries": {},
        "product_prices": {}, "petrol_expenses": [], "lunch_expenses": [],
        "misc_expenses": [], "booker_returns": [],
        "discount_packages": default_discount_packages(),
        "custom_products": [], "dsr_forms": [], "credit_bills": [],
        "deleted_products": [],
        "next_credit_id": 1,
        "wholesaler_packages": [
            {"id": "ws1", "name": "WS1", "discount_pct": 3.0},
            {"id": "ws2", "name": "WS2", "discount_pct": 5.0},
            {"id": "ws3", "name": "WS3", "discount_pct": 7.0},
        ],
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
# GLOBAL CSS (ANIMATED WAVES + GLASS LOGIN + OLD THEME)
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

    /* MAIN APP BACKGROUND */
    @keyframes gradientBG { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
    .stApp { background: linear-gradient(-45deg, #e3f2fd, #bbdefb, #e0f7fa, #e8eaf6) !important; background-size: 400% 400% !important; animation: gradientBG 15s ease infinite !important; overflow-x: hidden; }

    /* LOGIN SPECIFIC CSS */
    @keyframes waveMotion { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
    .login-bg {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: linear-gradient(270deg, #0f172a, #1e3a8a, #2563eb, #1e3a8a);
        background-size: 400% 400%;
        animation: waveMotion 15s ease infinite;
        z-index: -1;
    }

    /* ========================================================
       PROPER LOGIN CARD (middle column itself becomes the card)
       - ::before  = light golden glow coming from behind
       - ::after   = the card surface (cream / light golden)
       ======================================================== */
    @keyframes floatCard { 0% { transform: translateY(0px); } 50% { transform: translateY(-8px); } 100% { transform: translateY(0px); } }
    @keyframes goldPulse {
        0%   { opacity: 0.70; transform: scale(1); }
        50%  { opacity: 1;    transform: scale(1.07); }
        100% { opacity: 0.70; transform: scale(1); }
    }
    @keyframes cardShine {
        0%   { background-position: -300% 0; }
        100% { background-position: 300% 0; }
    }

    div[data-testid="stColumn"]:has(.login-card-marker),
    div[data-testid="column"]:has(.login-card-marker) {
        position: relative !important;
        padding: 38px 34px 26px 34px !important;
        margin-top: 4vh !important;
        background: transparent !important;
        overflow: visible !important;
        animation: floatCard 6s ease-in-out infinite;
        isolation: isolate;
    }
    /* soft golden light behind the card */
    div[data-testid="stColumn"]:has(.login-card-marker)::before,
    div[data-testid="column"]:has(.login-card-marker)::before {
        content: '';
        position: absolute;
        inset: -55px;
        border-radius: 60px;
        background:
            radial-gradient(circle at 50% 45%, rgba(255, 223, 100, 0.85) 0%, rgba(255, 200, 60, 0.55) 35%, rgba(255, 190, 40, 0.0) 70%);
        filter: blur(38px);
        z-index: -2;
        animation: goldPulse 4s ease-in-out infinite;
        pointer-events: none;
    }
    /* the actual card surface */
    div[data-testid="stColumn"]:has(.login-card-marker)::after,
    div[data-testid="column"]:has(.login-card-marker)::after {
        content: '';
        position: absolute;
        inset: 0;
        border-radius: 26px;
        background:
            linear-gradient(115deg, rgba(255,255,255,0) 35%, rgba(255,255,255,0.55) 50%, rgba(255,255,255,0) 65%),
            linear-gradient(160deg, #fffdf3 0%, #fff7d6 55%, #ffefb0 100%);
        background-size: 300% 100%, 100% 100%;
        animation: cardShine 7s linear infinite;
        border: 1.5px solid rgba(218, 165, 32, 0.55);
        box-shadow:
            0 25px 60px rgba(0, 0, 0, 0.35),
            0 0 0 5px rgba(255, 215, 0, 0.12),
            inset 0 1px 0 rgba(255, 255, 255, 0.9),
            inset 0 -12px 30px rgba(218, 165, 32, 0.12);
        z-index: -1;
        pointer-events: none;
    }

    .login-title {
        text-align: center;
        font-size: 36px;
        font-weight: 800;
        background: linear-gradient(135deg, #a8740a 0%, #daa520 50%, #f5c518 100%);
        -webkit-background-clip: text;
        background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: 2px;
        margin: 0 0 4px 0;
        line-height: 1.15;
    }
    .login-subtitle {
        text-align: center;
        font-size: 12px;
        color: #8b6508;
        letter-spacing: 4px;
        font-weight: 700;
        margin-bottom: 22px;
        text-transform: uppercase;
    }
    .login-divider {
        height: 2px; border-radius: 2px; margin: 0 auto 18px auto; width: 70px;
        background: linear-gradient(90deg, transparent, #daa520, transparent);
    }

    /* TABS INSIDE LOGIN CARD (GOLDEN THEME) */
    div[data-testid="stColumn"]:has(.login-card-marker) .stTabs [data-baseweb="tab-list"],
    div[data-testid="column"]:has(.login-card-marker) .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: rgba(255, 255, 255, 0.65);
        padding: 6px;
        border-radius: 12px;
        margin-bottom: 18px;
        border: 1px solid rgba(218, 165, 32, 0.3);
    }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTabs [data-baseweb="tab"],
    div[data-testid="column"]:has(.login-card-marker) .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 10px 15px;
        background-color: transparent;
        color: #8b6508 !important;
        font-weight: 700 !important;
        border: none;
        transition: all 0.3s ease;
    }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTabs [aria-selected="true"],
    div[data-testid="column"]:has(.login-card-marker) .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #b8860b !important;
        box-shadow: 0 2px 8px rgba(218, 165, 32, 0.25) !important;
    }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTabs [data-baseweb="tab-highlight"],
    div[data-testid="column"]:has(.login-card-marker) .stTabs [data-baseweb="tab-highlight"] {
        background-color: #daa520 !important;
    }

    /* INPUT FIELDS INSIDE LOGIN CARD (golden border) */
    div[data-testid="stColumn"]:has(.login-card-marker) .stTextInput > div > div > input,
    div[data-testid="column"]:has(.login-card-marker) .stTextInput > div > div > input,
    div[data-testid="stColumn"]:has(.login-card-marker) div[data-baseweb="input"] input,
    div[data-testid="column"]:has(.login-card-marker) div[data-baseweb="input"] input {
        background-color: rgba(255, 255, 255, 0.95) !important;
        border: 1.5px solid rgba(218, 165, 32, 0.55) !important;
        border-radius: 10px !important;
        color: #333 !important;
        transition: all 0.3s ease !important;
    }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTextInput > div > div > input:focus,
    div[data-testid="column"]:has(.login-card-marker) .stTextInput > div > div > input:focus {
        border-color: #daa520 !important;
        box-shadow: 0 0 12px rgba(218, 165, 32, 0.45) !important;
    }
    div[data-testid="stColumn"]:has(.login-card-marker) div[data-baseweb="input"],
    div[data-testid="column"]:has(.login-card-marker) div[data-baseweb="input"],
    div[data-testid="stColumn"]:has(.login-card-marker) div[data-baseweb="base-input"],
    div[data-testid="column"]:has(.login-card-marker) div[data-baseweb="base-input"] {
        background-color: transparent !important;
        border-color: transparent !important;
    }

    /* LOGIN BUTTON INSIDE CARD */
    div[data-testid="stColumn"]:has(.login-card-marker) .stButton > button,
    div[data-testid="column"]:has(.login-card-marker) .stButton > button {
        background: linear-gradient(135deg, #daa520 0%, #b8860b 100%) !important;
        border: none !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        border-radius: 10px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(218, 165, 32, 0.4) !important;
    }
    div[data-testid="stColumn"]:has(.login-card-marker) .stButton > button:hover,
    div[data-testid="column"]:has(.login-card-marker) .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 25px rgba(218, 165, 32, 0.6) !important;
        background: linear-gradient(135deg, #ffd700 0%, #daa520 100%) !important;
    }

    /* BUBBLES (For dashboard) */
    .bubbles { position: fixed; bottom: -100px; left: 0; width: 100%; height: 100%; z-index: 0; pointer-events: none; margin: 0; padding: 0; }
    .bubbles li { position: absolute; list-style: none; display: block; width: 40px; height: 40px; bottom: -160px; animation: square 25s infinite linear; border-radius: 50%; }
    .bubbles li:nth-child(1) { left: 10%; width: 80px; height: 80px; animation-delay: 0s; background: rgba(33,150,243,0.15); }
    .bubbles li:nth-child(2) { left: 20%; width: 30px; height: 30px; animation-delay: 2s; animation-duration: 12s; background: rgba(76,175,80,0.15); }
    .bubbles li:nth-child(3) { left: 35%; width: 100px; height: 100px; animation-delay: 4s; background: rgba(255,152,0,0.15); }
    .bubbles li:nth-child(4) { left: 50%; width: 60px; height: 60px; animation-delay: 0s; animation-duration: 18s; background: rgba(156,39,176,0.15); }
    .bubbles li:nth-child(5) { left: 65%; width: 40px; height: 40px; animation-delay: 1s; background: rgba(0,188,212,0.15); }
    .bubbles li:nth-child(6) { left: 80%; width: 120px; height: 120px; animation-delay: 3s; background: rgba(233,30,99,0.12); }
    .bubbles li:nth-child(7) { left: 90%; width: 50px; height: 50px; animation-delay: 5s; animation-duration: 14s; background: rgba(255,193,7,0.15); }
    .bubbles li:nth-child(8) { left: 25%; width: 90px; height: 90px; animation-delay: 7s; background: rgba(76,175,80,0.15); }
    .bubbles li:nth-child(9) { left: 75%; width: 70px; height: 70px; animation-delay: 2s; animation-duration: 22s; background: rgba(33,150,243,0.15); }
    @keyframes square { 0% { transform: translateY(0) rotate(0deg); opacity: 1; border-radius: 50%; } 100% { transform: translateY(-1000px) rotate(720deg); opacity: 0; border-radius: 50%; } }

    @keyframes heroShade { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
    .dash-hero { background: linear-gradient(-45deg, #1e3a8a, #2563eb, #3b82f6, #8b5cf6, #2563eb) !important; background-size: 300% 300% !important; animation: heroShade 8s ease infinite !important; border-radius: 20px; padding: 26px 32px; margin-bottom: 22px; color: #fff; position: relative; overflow: hidden; box-shadow: 0 12px 32px rgba(37, 99, 235, 0.35); z-index: 2; }

    .kpi-card { transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important; z-index: 5 !important; }
    .kpi-card:hover { transform: translateY(-8px) scale(1.02) !important; box-shadow: 0 15px 30px rgba(37,99,235,0.25) !important; }

    section[data-testid="stSidebar"] div[role="radiogroup"] > label { transition: all 0.25s ease !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover { transform: translateX(6px) !important; background-color: #e8f5e9 !important; }

    @keyframes fadeInPage { 0% { opacity: 0; transform: translateY(15px); } 100% { opacity: 1; transform: translateY(0); } }
    .block-container { animation: fadeInPage 0.6s ease-out !important; position: relative; z-index: 2; padding-top: 0.5rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; padding-bottom: 0.5rem !important; max-width: 100% !important; }

    html, body, .stApp, .stApp p, .stApp span, .stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp label, .stApp li, .stApp a, [class*="css"] *, [data-testid="stMarkdownContainer"] *, [data-testid="stText"], [data-testid="stCaptionContainer"] *, [data-testid="stWidgetLabel"] *, [data-testid="stSelectbox"] *, [data-testid="stTextInput"] *, [data-testid="stNumberInput"] * { color: #000000; }

    .stDateInput, .stDateInput > div, .stDateInput > div > div, .stDateInput > div > div > input, [data-testid="stDateInput"], [data-testid="stDateInput"] > div, [data-testid="stDateInput"] > div > div, [data-testid="stDateInput"] input, [data-testid="stDateInput"] div[data-baseweb="input"], [data-testid="stDateInput"] div[data-baseweb="input"] > div, [data-testid="stDateInput"] div[data-baseweb="base-input"], [data-baseweb="datepicker"] input, div[data-baseweb="input"] > div, div[data-baseweb="base-input"], div[data-baseweb="base-input"] input { background-color: #ffffff !important; color: #000000 !important; -webkit-text-fill-color: #000000 !important; border-color: #90caf9 !important; }
    div[data-baseweb="calendar"], div[data-baseweb="calendar"] *, div[data-baseweb="datepicker"] *, div[data-baseweb="popover"] *, div[role="dialog"] *, div[role="dialog"] { background-color: #ffffff !important; color: #000000 !important; -webkit-text-fill-color: #000000 !important; }
    [data-testid="stDateInput"] svg, div[data-baseweb="datepicker"] svg { fill: #1976d2 !important; color: #1976d2 !important; }

    .stButton > button, .stButton > button p, .stButton > button span, .stButton > button div, .stDownloadButton > button, .stDownloadButton > button p, .stDownloadButton > button span, .stDownloadButton > button div { color: #ffffff !important; }

    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    [data-testid="stStatusWidget"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    header[data-testid="stHeader"] { background: transparent !important; box-shadow: none !important; }

    .stApp .element-container { margin-bottom: 0.35rem !important; }
    .stApp [data-testid="stVerticalBlock"] > div { gap: 0.35rem !important; }
    .stApp label, .stApp [data-testid="stWidgetLabel"] label, .stApp [data-testid="stWidgetLabel"] p { font-size: 13px !important; font-weight: 600 !important; margin-bottom: 2px !important; }
    .stTextInput > div > div > input, .stNumberInput > div > div > input, .stDateInput > div > div > input, .stSelectbox > div > div > div, div[data-baseweb="input"] input, div[data-baseweb="select"] > div { padding: 4px 8px !important; min-height: 34px !important; font-size: 14px !important; }
    .stButton > button { padding: 4px 10px !important; min-height: 36px !important; font-size: 14px !important; }
    .stApp hr { margin: 6px 0 !important; }
    .stApp h3 { margin-top: 6px !important; margin-bottom: 4px !important; font-size: 18px !important; }

    .stTextInput > div > div > input, .stNumberInput > div > div > input, .stSelectbox > div > div > div, .stSelectbox > div > div, div[data-baseweb="select"] > div, div[data-baseweb="input"] input { background-color: #ffffff !important; color: #000000 !important; -webkit-text-fill-color: #000000 !important; border: 2px solid #90caf9 !important; border-radius: 8px !important; }
    div[data-baseweb="select"] * { color: #000000 !important; }
    ul[role="listbox"] li, div[role="option"] { color: #000000 !important; background-color: #ffffff !important; }
    input[type="password"] { background-color: #ffffff !important; color: #000000 !important; -webkit-text-fill-color: #000000 !important; }

    /* Re-apply golden input border inside the login card (must come AFTER the global input rule above) */
    div[data-testid="stColumn"]:has(.login-card-marker) .stTextInput > div > div > input,
    div[data-testid="column"]:has(.login-card-marker) .stTextInput > div > div > input,
    div[data-testid="stColumn"]:has(.login-card-marker) div[data-baseweb="input"] input,
    div[data-testid="column"]:has(.login-card-marker) div[data-baseweb="input"] input {
        border: 1.5px solid rgba(218, 165, 32, 0.6) !important;
        border-radius: 10px !important;
    }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTextInput > div > div > input:focus,
    div[data-testid="column"]:has(.login-card-marker) .stTextInput > div > div > input:focus,
    div[data-testid="stColumn"]:has(.login-card-marker) div[data-baseweb="input"] input:focus,
    div[data-testid="column"]:has(.login-card-marker) div[data-baseweb="input"] input:focus {
        border-color: #daa520 !important;
        box-shadow: 0 0 12px rgba(218, 165, 32, 0.45) !important;
    }
    div[data-testid="stColumn"]:has(.login-card-marker) div[data-baseweb="input"] > div,
    div[data-testid="column"]:has(.login-card-marker) div[data-baseweb="input"] > div,
    div[data-testid="stColumn"]:has(.login-card-marker) div[data-baseweb="base-input"],
    div[data-testid="column"]:has(.login-card-marker) div[data-baseweb="base-input"] {
        border-color: transparent !important;
        background-color: transparent !important;
    }

    .stDownloadButton > button { background: linear-gradient(135deg, #0288d1 0%, #0277bd 100%) !important; color: #ffffff !important; font-weight: bold !important; border-radius: 8px !important; }
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

    .auto-dl-hidden div[data-testid="stDownloadButton"] { position: absolute !important; left: -9999px !important; top: -9999px !important; opacity: 0 !important; height: 0 !important; }
    .metric-card { background: #ffffff; border: 2px solid #90caf9; border-radius: 14px; padding: 22px; text-align: center; box-shadow: 0 3px 10px rgba(33,150,243,0.15); }
    .metric-card h3 { font-size: 13px !important; margin: 0 !important; font-weight: 700 !important; text-transform: uppercase; color: #0277bd !important; }
    .metric-card h1 { font-size: 30px !important; margin: 10px 0 0 0 !important; font-weight: 800 !important; color: #1976d2 !important; }
    .booker-row { background: #ffffff; border: 1px solid #90caf9; border-radius: 10px; padding: 12px 18px; margin-bottom: 8px; }
    .empty-box { background: #ffffff; border: 2px dashed #90caf9; border-radius: 12px; padding: 20px; text-align: center; font-size: 14px; color: #0277bd !important; }
    .hint-box { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 6px 10px; border-radius: 6px; font-size: 12px; margin-top: 2px; color: #1976d2 !important; }
    .summary-box { background: #ffffff; border: 2px solid #2196f3; border-radius: 12px; padding: 12px 18px; margin-bottom: 12px; color: #000000 !important; }
    .lf-simple-card { background: #ffffff; border-left: 6px solid #2196f3; border-radius: 12px; padding: 16px 22px; margin-bottom: 12px; box-shadow: 0 3px 10px rgba(33,150,243,0.15); display: flex; align-items: center; justify-content: space-between; }
    .lf-simple-card .lf-info { display: flex; flex-direction: column; gap: 4px; }
    .lf-simple-card .lf-line1 { font-size: 17px; font-weight: 700; color: #1976d2 !important; }
    .lf-simple-card .lf-line2 { font-size: 13px; color: #0277bd !important; }
    .lf-simple-card .lf-boxes { background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%); font-weight: 800; font-size: 20px; padding: 10px 18px; border-radius: 10px; text-align: center; min-width: 90px; color: #ffffff !important; }
    .lf-simple-card .lf-boxes small { display: block; font-size: 10px; font-weight: 500; opacity: 0.9; }
    .lf-simple-card.dsr { border-left-color: #e65100; }
    .lf-simple-card.dsr .lf-boxes { background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%); }
    .lf-simple-card.credit-pending { border-left-color: #e65100; }
    .lf-simple-card.credit-pending .lf-boxes { background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%); }
    .lf-simple-card.credit-paid { border-left-color: #2e7d32; background: #f1f8e9; }
    .lf-simple-card.credit-paid .lf-boxes { background: linear-gradient(135deg, #2e7d32 0%, #1b5e20 100%); }
    .lf-simple-card.ret { border-left-color: #c62828; background: #ffebee; }
    .lf-simple-card.ret .lf-boxes { background: linear-gradient(135deg, #ef4444 0%, #c62828 100%); }

    .ret-item-row { background: #fff8e1; border: 1px solid #ffcc80; border-radius: 8px; padding: 6px 10px; margin-bottom: 5px; }
    .ret-item-row .ri-code { display: inline-block; background: #1976d2; color: #fff !important; font-weight: 800; font-size: 11px; padding: 2px 8px; border-radius: 5px; margin-right: 8px; }
    .ret-item-row .ri-name { font-weight: 600; color: #1976d2 !important; }
    .ret-item-row .ri-loaded { color: #2e7d32 !important; font-weight: 700; }

    .danger-box { background: #ffebee; border: 2px solid #c62828; border-radius: 12px; padding: 14px 18px; margin-bottom: 12px; }
    .danger-box b { color: #b71c1c !important; }

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
    .badge-return { background: linear-gradient(135deg, #ef4444 0%, #c62828 100%); color: #ffffff !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; margin-left: 8px; }
    .badge-selected { background: linear-gradient(135deg, #dc2626 0%, #991b1b 100%); color: #ffffff !important; padding: 3px 10px; border-radius: 6px; font-size: 11px; font-weight: 700; margin-left: 8px; }

    .bk-product-card { background: #ffffff; border: 1.5px solid #90caf9; border-radius: 10px; padding: 8px 12px; margin-bottom: 6px; display: flex; align-items: center; justify-content: space-between; transition: all 0.15s ease; }
    .bk-product-card:hover { border-color: #2196f3; box-shadow: 0 2px 8px rgba(33,150,243,0.15); transform: translateY(-2px); }
    .bk-product-card .bk-name { font-size: 13.5px; font-weight: 700; color: #1976d2; }
    .bk-product-card .bk-code { font-size: 11px; color: #666; }
    .bk-product-card .bk-tp { background: linear-gradient(135deg, #4caf50 0%, #2e7d32 100%); color: #fff !important; font-weight: 800; font-size: 13px; padding: 4px 12px; border-radius: 6px; }

    section[data-testid="stSidebar"] { background: linear-gradient(180deg, rgba(241,248,233,0.95) 0%, rgba(232,245,233,0.95) 50%, rgba(220,237,200,0.95) 100%) !important; border-right: 2px solid #a5d6a7 !important; box-shadow: 4px 0 20px rgba(76,175,80,0.12) !important; backdrop-filter: blur(10px) !important; z-index: 10 !important; }
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
    section[data-testid="stSidebar"] div[role="radiogroup"] > label { background-color: #ffffff !important; border: 1px solid rgba(76,175,80,0.18) !important; border-radius: 10px !important; margin-bottom: 2px !important; padding: 9px 12px !important; cursor: pointer !important; display: flex !important; align-items: center !important; width: 100% !important; position: relative !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover { border-color: #66bb6a !important; box-shadow: 0 2px 8px rgba(76,175,80,0.15); }
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

    section[data-testid="stSidebar"] .stButton > button { background: linear-gradient(135deg, #ffffff 0%, #f1f8e9 100%) !important; border: 1.5px solid rgba(239,68,68,0.35) !important; color: #c62828 !important; font-weight: 800 !important; border-radius: 10px !important; margin: 0 14px 16px !important; width: calc(100% - 28px) !important; box-shadow: 0 2px 6px rgba(239,68,68,0.10); }
    section[data-testid="stSidebar"] .stButton > button:hover { background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%) !important; border-color: #ef4444 !important; color: #ffffff !important; box-shadow: 0 6px 14px rgba(239,68,68,0.4) !important; transform: translateY(-2px) !important; }
    section[data-testid="stSidebar"] .stButton > button p, section[data-testid="stSidebar"] .stButton > button span, section[data-testid="stSidebar"] .stButton > button div { color: inherit !important; }

    .dash-hero::before { content: ''; position: absolute; top: -50%; right: -10%; width: 420px; height: 420px; background: radial-gradient(circle, rgba(255,255,255,0.18) 0%, transparent 70%); border-radius: 50%; }
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

    .kpi-card { background: #ffffff; border-radius: 16px; padding: 20px 22px; border: 1px solid rgba(37,99,235,0.12); box-shadow: 0 4px 16px rgba(37,99,235,0.08); display: flex; flex-direction: column; gap: 12px; position: relative; overflow: hidden; height: 100%; }
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

    .bb-summary { background: linear-gradient(135deg, #ffffff 0%, #f0f9ff 100%); border: 2px solid #2563eb; border-radius: 14px; padding: 16px 22px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 16px; }
    .bb-stat { text-align: center; }
    .bb-stat-label { font-size: 10px; color: #64748b !important; text-transform: uppercase; letter-spacing: 1.2px; font-weight: 800; }
    .bb-stat-val { font-size: 26px; font-weight: 800; color: #1976d2 !important; line-height: 1.1; }
    .bb-stat-val.green { color: #2e7d32 !important; }
    .bb-stat-val.red { color: #c62828 !important; }

    .item-line { display: flex; align-items: center; justify-content: space-between; background: #ffffff; border: 1px solid #b3e5fc; border-radius: 8px; padding: 8px 12px; margin-bottom: 6px; }
</style>

<ul class="bubbles">
    <li></li><li></li><li></li><li></li><li></li>
    <li></li><li></li><li></li><li></li>
</ul>
""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR TOGGLE JS
# ============================================================
components.html("""
<script>
(function(){
    function killManageApp() {
        try {
            var doc = window.parent.document;
            var sel = ['[data-testid="manage-app-button"]','[data-testid="stAppDeployButton"]','[data-testid="stCloudAppManageButton"]','.stAppDeployButton','iframe[title="streamlit_cloud_status"]','div[class*="manageApp"]','div[class*="ManageApp"]','button[class*="manageApp"]','button[class*="ManageApp"]'];
            sel.forEach(function(s){ doc.querySelectorAll(s).forEach(function(el){ el.style.setProperty('display','none','important'); el.style.setProperty('visibility','hidden','important'); el.style.setProperty('opacity','0','important'); }); });
            doc.querySelectorAll('button, a').forEach(function(el){ try { var t = (el.textContent || '').trim(); if (t === 'Manage app' || t === 'Manage App') { el.style.setProperty('display','none','important'); } } catch(e){} });
        } catch(e) {}
    }
    function attachToggle() {
        try {
            var doc = window.parent.document;
            if (!doc.querySelector('section[data-testid="stSidebar"]')) {
                var existing = doc.getElementById('custom-sidebar-toggle');
                if (existing) existing.parentNode.removeChild(existing); return;
            }
            var old = doc.getElementById('custom-sidebar-toggle');
            if (old) old.parentNode.removeChild(old);
            var btn = doc.createElement('button'); btn.id = 'custom-sidebar-toggle'; btn.title = 'Sidebar Open/Close'; btn.innerHTML = '\\u2630';
            var s = {'position':'fixed','top':'14px','left':'14px','z-index':'2147483647','background':'linear-gradient(135deg, #66bb6a 0%, #2e7d32 100%)','color':'#fff','border':'none','border-radius':'10px','padding':'8px 14px','font-size':'18px','font-weight':'bold','cursor':'pointer','box-shadow':'0 4px 14px rgba(76,175,80,0.5)','transition': 'all 0.3s ease'};
            for (var k in s) btn.style.setProperty(k, s[k], 'important');
            btn.onclick = function() {
                var targets = ['[data-testid="stSidebarCollapseButton"] button','[data-testid="stSidebarCollapsedControl"] button','[data-testid="collapsedControl"] button','[data-testid="stExpandSidebarButton"] button','button[kind="headerNoPadding"]'];
                for (var i = 0; i < targets.length; i++) { var el = doc.querySelector(targets[i]); if (el) { el.click(); return; } }
            };
            if (doc.body) doc.body.appendChild(btn);
        } catch(e) {}
    }
    function tick() { killManageApp(); attachToggle(); }
    setTimeout(tick, 300); setTimeout(tick, 1000); setTimeout(tick, 2000); setTimeout(tick, 3000); setInterval(tick, 1000);
})();
</script>
""", height=0)

COMPANY_NAME = "AL-BARAKAH ENTERPRISES"
PRODUCTS = []

# ============================================================
# AUTH SCREEN (PREMIUM ANIMATED WAVES + GOLDEN CARD)
# ============================================================
def render_auth_page():
    st.markdown("""
    <style>
    @keyframes waveMove {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    .login-bg {
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: linear-gradient(270deg, #0f172a, #1e3a8a, #2563eb, #3b82f6, #8b5cf6, #2563eb, #1e3a8a, #0f172a);
        background-size: 400% 400%;
        animation: waveMove 15s ease infinite;
        z-index: -1;
    }
    </style>
    <div class="login-bg"></div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        # The marker lets CSS turn this whole column into the login card
        st.markdown("""
        <div class="login-card-marker"></div>
        <div class="login-title">AL-BARAKAH</div>
        <div class="login-subtitle">ENTERPRISES 2026</div>
        <div class="login-divider"></div>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3 = st.tabs(["🔐 Admin", "👤 Booker", "📝 Signup"])

        with tab1:
            login_user = st.text_input("Username", key="login_username", placeholder="Enter admin username")
            login_pass = st.text_input("Password", key="login_password", type="password", placeholder="Enter password")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🔓 Login to Dashboard", key="btn_login", use_container_width=True, type="primary"):
                users = load_users()
                uname = login_user.strip().lower()
                if uname == "" or login_pass == "": st.error("❌ Both fields required")
                elif uname not in users: st.error("❌ Username not found")
                elif users[uname].get("password_hash") != hash_password(login_pass): st.error("❌ Incorrect password")
                else:
                    st.session_state["logged_in_user"] = uname
                    st.session_state["display_name"] = users[uname].get("display_name", uname)
                    st.session_state["role"] = "admin"
                    st.session_state["page"] = "📊 Dashboard"
                    st.rerun()

        with tab2:
            bk_user = st.text_input("Booker Name", key="bk_login_user", placeholder="Enter your name")
            bk_pass = st.text_input("Password", key="bk_login_pass", type="password", placeholder="Enter password")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("👤 Login as Booker", key="btn_bk_login", use_container_width=True, type="primary"):
                registry = load_bookers()
                bk_uname = sanitize_username(bk_user.strip())
                if not bk_uname or not bk_pass: st.error("❌ Both fields required")
                elif bk_uname not in registry: st.error("❌ Booker not found. Contact Admin.")
                elif registry[bk_uname].get("password_hash") != hash_password(bk_pass): st.error("❌ Incorrect password")
                else:
                    admin_uname = registry[bk_uname].get("admin", "")
                    admin_db = load_database_for(admin_uname)
                    if registry[bk_uname].get("display_name") not in admin_db.get("bookers", []):
                        st.error("❌ You have been removed from bookers list by Admin.")
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
                        draft_file = get_booker_draft_file(st.session_state["booker_name"])
                        if os.path.exists(draft_file):
                            try:
                                with open(draft_file, "r", encoding="utf-8") as f:
                                    draft_data = json.load(f)
                                    st.session_state["bk_current_shop"] = draft_data.get("shop", "")
                                    st.session_state["bk_bill_draft"] = draft_data.get("items", [])
                            except Exception: pass
                        st.rerun()

        with tab3:
            su_user = st.text_input("New Username", key="su_username", placeholder="3-20 characters")
            su_pass = st.text_input("Password", key="su_password", type="password")
            su_pass2 = st.text_input("Confirm Password", key="su_password2", type="password")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Create Account", key="btn_signup", use_container_width=True, type="primary"):
                users = load_users()
                uname_safe = sanitize_username(su_user.strip())
                if len(uname_safe) < 3: st.error("❌ Username too short")
                elif len(su_pass) < 4: st.error("❌ Password too short")
                elif su_pass != su_pass2: st.error("❌ Passwords do not match")
                elif uname_safe in users: st.error("❌ Username already exists")
                else:
                    users[uname_safe] = {"username": uname_safe, "display_name": su_user.strip(), "password_hash": hash_password(su_pass), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                    save_users(users)
                    with open(user_data_file(uname_safe), "w", encoding="utf-8") as f: json.dump(default_blank_db(), f, indent=4, ensure_ascii=False)
                    st.success(f"✅ Account created — Username: **{uname_safe}**")

# ============================================================
# LOAD DB & CHECK LOGIN
# ============================================================
def load_database_for(username):
    if not username: return default_blank_db()
    fpath = user_data_file(username)
    if os.path.exists(fpath):
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k in ["bookers", "salesmen", "load_forms", "petrol_expenses", "lunch_expenses", "misc_expenses", "booker_returns", "custom_products", "dsr_forms", "credit_bills", "deleted_products"]:
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
                if "next_credit_id" not in data: data["next_credit_id"] = max([c.get("id", 0) for c in data.get("credit_bills", [])] + [0]) + 1
                if "wholesaler_packages" not in data or not data["wholesaler_packages"]:
                    data["wholesaler_packages"] = [{"id": "ws1", "name": "WS1", "discount_pct": 3.0}, {"id": "ws2", "name": "WS2", "discount_pct": 5.0}, {"id": "ws3", "name": "WS3", "discount_pct": 7.0}]
                return data
        except Exception: pass
    return default_blank_db()

if "logged_in_user" not in st.session_state or not st.session_state["logged_in_user"]:
    render_auth_page()
    st.stop()

CURRENT_USER = st.session_state["logged_in_user"]
ROLE = st.session_state.get("role", "admin")

def save_database(db):
    try:
        with open(user_data_file(CURRENT_USER), "w", encoding="utf-8") as f: json.dump(db, f, indent=4, ensure_ascii=False)
    except Exception as e: st.warning(f"⚠️ Save Error: {e}")

def parse_date(dstr):
    try: return datetime.strptime(dstr, "%d-%m-%Y").date()
    except Exception: return None

if "database" not in st.session_state: st.session_state.database = load_database_for(CURRENT_USER)
if st.session_state.get("_db_user") != CURRENT_USER:
    st.session_state.database = load_database_for(CURRENT_USER)
    st.session_state["_db_user"] = CURRENT_USER

for _k, _v in [("_prev_prod", None), ("last_bill_no", None), ("download_file", None), ("_dl_counter", 0), ("page", "📊 Dashboard"), ("view_bill_key", None), ("view_lf_key", None), ("view_dsr_key", None), ("view_credit_key", None), ("bk_page", "🧾 New Bill"), ("bk_bill_draft", []), ("bk_load_draft", []), ("bk_current_shop", ""), ("bk_bills_created", 0), ("bk_reset_item_form", False), ("bk_reset_shop_form", False), ("prod_select_mode", False), ("prod_confirm_all", False)]:
    if _k not in st.session_state: st.session_state[_k] = _v

db = st.session_state.database
for k in ["bookers", "salesmen", "load_forms", "petrol_expenses", "lunch_expenses", "misc_expenses", "booker_returns", "custom_products", "dsr_forms", "credit_bills", "deleted_products"]:
    if k not in db: db[k] = []
for k in ["bookers_salaries", "salesmen_salaries", "product_prices"]:
    if k not in db: db[k] = {}
if "discount_packages" not in db or not db["discount_packages"]: db["discount_packages"] = default_discount_packages()
for p in db["discount_packages"]:
    if "tier3_amount" not in p: p["tier3_amount"] = 0.0
    if "tier3_pct" not in p: p["tier3_pct"] = 0.0
if "next_credit_id" not in db: db["next_credit_id"] = max([c.get("id", 0) for c in db.get("credit_bills", [])] + [0]) + 1
if "wholesaler_packages" not in db or not db["wholesaler_packages"]:
    db["wholesaler_packages"] = [{"id": "ws1", "name": "WS1", "discount_pct": 3.0}, {"id": "ws2", "name": "WS2", "discount_pct": 5.0}, {"id": "ws3", "name": "WS3", "discount_pct": 7.0}]

# ============================================================
# HELPERS
# ============================================================
def is_wholesaler(shop_name):
    if not shop_name: return False
    s = str(shop_name).lower().strip()
    for token in ["ws1", "ws2", "ws3"]:
        if s.endswith(token) or f" {token}" in s: return True
    s_compact = "".join(ch for ch in s if ch.isalnum())
    for token in ["wholeseller", "wholesaler", "wholseller", "wholsaler", "whole seller", "whole saler", "holeseller", "holesaler"]:
        if token.replace(" ", "") in s_compact: return True
    return False

def get_wholesaler_package_from_shop(shop_name):
    if not shop_name: return None
    s = str(shop_name).lower().strip()
    packages = db.get("wholesaler_packages", [])
    for token in ["ws1", "ws2", "ws3"]:
        if s.endswith(token) or f" {token}" in s:
            for p in packages:
                if p["id"].lower() == token: return p
    return None

def get_all_products():
    custom = db.get("custom_products", [])
    merged = list(PRODUCTS)
    seen_codes = {str(p["code"]) for p in merged}
    for cp in custom:
        code = str(cp.get("code", "")).strip()
        name = str(cp.get("name", "")).strip()
        if not code or not name: continue
        if code in seen_codes: continue
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
        tiers = [(float(pkg.get("tier3_amount", 0) or 0), float(pkg.get("tier3_pct", 0) or 0), "Tier3"), (float(pkg.get("tier2_amount", 0) or 0), float(pkg.get("tier2_pct", 0) or 0), "Tier2"), (float(pkg.get("tier1_amount", 0) or 0), float(pkg.get("tier1_pct", 0) or 0), "Tier1")]
        tiers.sort(key=lambda x: x[0], reverse=True)
        for amt, pct, label in tiers:
            if amt > 0 and bill_total > amt and pct > best_pct:
                best_pct = pct; best_name = pkg_name; best_tier = f"{label} (> {amt:,.0f})"; break
    if best_pct > 0: return best_pct, best_name, best_tier
    return 0.0, (active_pkgs[0].get("name") if active_pkgs else None), None

def get_effective_discount_pct(shop_name, bill_total):
    ws_pkg = get_wholesaler_package_from_shop(shop_name)
    if ws_pkg:
        pct = float(ws_pkg.get("discount_pct", 0))
        return pct, f"{ws_pkg['name']} ({pct}%)", "WS Package"
    if is_wholesaler(shop_name): return 0.0, "Wholesaler Keyword", "No Package Matched"
    return get_package_discount_pct(bill_total)

def get_dsr_credit_info(dsr):
    booker = dsr.get("booker", "").strip()
    dsr_date = dsr.get("date", "").strip()
    credit_total = 0.0; credit_shops = {}; credit_details = []
    for sb in dsr.get("source_bills", []):
        shop = sb.get("shop", ""); bill_no = sb.get("bill_no", "")
        for cb in db.get("credit_bills", []):
            if (str(cb.get("booker", "")).strip() == booker and 
                str(cb.get("bill_no", "")) == str(bill_no) and 
                str(cb.get("shop", "")) == str(shop) and
                str(cb.get("date", "")).strip() == dsr_date):
                amt = float(cb.get("total_net", 0)); credit_total += amt; credit_shops[shop] = credit_shops.get(shop, 0) + amt
                credit_details.append({"shop": shop, "bill_no": bill_no, "amount": amt, "status": cb.get("status", "pending"), "credit_id": cb.get("id")}); break
    return credit_total, credit_shops, credit_details

def get_lf_return_items(lf_id):
    result = {}
    for r in db.get("booker_returns", []):
        if r.get("load_form_id") != lf_id: continue
        items = r.get("items", [])
        for it in items: c = str(it.get("Code", "")); result[c] = result.get(c, 0) + int(it.get("ReturnBoxes", 0))
        if not items and r.get("boxes"): result["__misc__"] = result.get("__misc__", 0) + int(r.get("boxes", 0))
    return result

def get_booker_total_return_boxes(booker, in_range_func=None):
    total = 0
    for r in db.get("booker_returns", []):
        if str(r.get("booker", "")).strip() != booker: continue
        if in_range_func and not in_range_func(r.get("date", "")): continue
        items = r.get("items", [])
        if items: total += sum(int(it.get("ReturnBoxes", 0)) for it in items)
        else: total += int(r.get("boxes", 0))
    return total

def get_booker_returns_list(booker, in_range_func=None):
    out = []
    for r in db.get("booker_returns", []):
        if str(r.get("booker", "")).strip() != booker: continue
        if in_range_func and not in_range_func(r.get("date", "")): continue
        out.append(r)
    return out

def get_return_record_for_lf(lf_id):
    for r in db.get("booker_returns", []):
        if r.get("load_form_id") == lf_id: return r
    return None

def build_load_form_excel(lf):
    output = BytesIO(); wb = xlsxwriter.Workbook(output, {'in_memory': True}); ws = wb.add_worksheet("Load Form")
    title_fmt = wb.add_format({"bold": True, "font_size": 16, "align": "center", "border": 2, "bg_color": "#e2e8f0"})
    header_fmt = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f1f5f9", "align": "center", "border": 1})
    label_fmt = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f8fafc", "align": "left", "border": 1})
    val_fmt = wb.add_format({"font_size": 11, "align": "left", "border": 1})
    cl = wb.add_format({"font_size": 11, "border": 1, "align": "left"})
    cc = wb.add_format({"font_size": 11, "border": 1, "align": "center"})
    tot = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#fef3c7", "align": "center", "border": 1})
    net_fmt = wb.add_format({"bold": True, "font_size": 13, "bg_color": "#dcfce7", "align": "center", "border": 2})
    ws.merge_range("A1:C1", f"{COMPANY_NAME} — LOAD FORM", title_fmt)
    ws.write("A3", "Booker:", label_fmt); ws.write("B3", lf.get("booker", ""), val_fmt)
    ws.write("A4", "Date:", label_fmt); ws.write("B4", lf.get("date", ""), val_fmt); ws.write("C4", lf.get("time", ""), val_fmt)
    ws.write("A5", "Load Form #:", label_fmt); ws.write("B5", lf.get("id", ""), val_fmt)
    ws.write("A7", "Code", header_fmt); ws.write("B7", "Product", header_fmt); ws.write("C7", "Boxes", header_fmt)
    row = 7; total_boxes = 0
    for it in lf.get("items", []):
        bx = int(it.get("Boxes", 0)); ws.write(row, 0, str(it.get("Code", "")), cc); ws.write(row, 1, it.get("Product", ""), cl); ws.write(row, 2, bx, cc); total_boxes += bx; row += 1
    ws.write(row, 0, "TOTAL", tot); ws.write(row, 1, f"{len(lf.get('items', []))} Products", tot); ws.write(row, 2, total_boxes, tot); row += 2
    ws.merge_range(row, 0, row, 1, "TOTAL BOXES", header_fmt); ws.merge_range(row, 2, row, 2, total_boxes, net_fmt)
    code_width = max([len(str(it.get("Code", ""))) for it in lf.get("items", [])] + [4]) + 4
    prod_width = max([len(str(it.get("Product", ""))) for it in lf.get("items", [])] + [8]) + 6
    ws.set_column("A:A", max(14, code_width)); ws.set_column("B:B", max(30, prod_width)); ws.set_column("C:C", 12)
    ws.set_paper(9); ws.set_portrait(); ws.fit_to_pages(1, 1); wb.close(); output.seek(0); return output.getvalue()

def export_lf_excel_callback(lf_id):
    lf = next((x for x in db.get("load_forms", []) if x.get("id") == lf_id), None)
    if not lf: st.session_state["error_msg"] = "❌ Load Form not found"; return
    safe_booker = str(lf.get("booker", "booker")).replace(" ", "_").replace("/", "-")
    safe_date = str(lf.get("date", "")).replace("-", "")
    st.session_state["download_file"] = (f"LoadForm_{lf_id}_{safe_booker}_{safe_date}.xlsx", build_load_form_excel(lf))
    st.session_state["success_msg"] = f"✅ Load Form #{lf_id} Excel ready"

def show_auto_download():
    if st.session_state.get("download_file"):
        fname, fdata = st.session_state["download_file"]; st.session_state["download_file"] = None; st.session_state["_dl_counter"] += 1
        st.markdown('<div class="auto-dl-hidden">', unsafe_allow_html=True)
        st.download_button(label=f"Download {fname}", data=fdata, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"auto_dl_{st.session_state['_dl_counter']}")
        st.markdown('</div>', unsafe_allow_html=True)
        components.html("""<script>(function(){ var tries=0; function attempt(){ tries++; var btns = window.parent.document.querySelectorAll('[data-testid="stDownloadButton"] button'); if (btns.length > 0){ btns[btns.length - 1].click(); return; } if (tries < 25) setTimeout(attempt, 200); } setTimeout(attempt, 300); })();</script>""", height=0)

# ============================================================
# BOOKER UI
# ============================================================
def render_booker_sidebar():
    with st.sidebar:
        bk_name = st.session_state.get("booker_name", "Booker")
        avatar_letter = (bk_name[0] if bk_name else "B").upper()
        st.markdown(f"""<div class="sb-brand"><div class="sb-brand-logo">👤</div><div class="sb-brand-name">BOOKER PANEL</div><div class="sb-brand-sub">AL-BARAKAH</div></div>""", unsafe_allow_html=True)
        st.markdown(f"""<div class="sb-user"><div class="sb-user-avatar">{avatar_letter}</div><div class="sb-user-info"><div class="sb-user-label">Booker</div><div class="sb-user-name">{bk_name}</div></div></div>""", unsafe_allow_html=True)
        st.markdown('<div class="sb-section-label">Menu</div>', unsafe_allow_html=True)

        bk_page = st.radio("BK MENU", ["🧾 New Bill", "📋 My Bills History", "🛒 All Products", "📦 Load Form Status"], key="bk_page_selector", label_visibility="collapsed")
        st.session_state["bk_page"] = bk_page
        st.markdown(f'<div class="sb-date" style="text-align:center;padding:16px;font-size:11px;color:#94a3b8;font-weight:600;">📅 {datetime.now().strftime("%A, %d %b %Y")}</div>', unsafe_allow_html=True)

        bill_count = st.session_state.get("bk_bills_created", 0); load_count = len(st.session_state.get("bk_load_draft", [])); total_boxes = sum(int(it.get("Boxes", 0)) for it in st.session_state.get("bk_load_draft", []))
        st.markdown(f"""<div class="sb-stats"><div class="sb-section-label" style="padding: 0 4px 8px;">Session</div><div class="sb-stats-grid"><div class="sb-stat-card"><div class="sb-stat-icon">🧾</div><div class="sb-stat-value">{bill_count}</div><div class="sb-stat-label">Bills</div></div><div class="sb-stat-card"><div class="sb-stat-icon">📦</div><div class="sb-stat-value">{total_boxes}</div><div class="sb-stat-label">Boxes</div></div><div class="sb-stat-card"><div class="sb-stat-icon">🎯</div><div class="sb-stat-value">{load_count}</div><div class="sb-stat-label">Products</div></div></div></div>""", unsafe_allow_html=True)

        if st.button("🚪 Logout", key="btn_bk_logout", use_container_width=True):
            for k in ["logged_in_user", "display_name", "role", "booker_name", "booker_admin", "bk_bill_draft", "bk_load_draft", "bk_current_shop", "bk_bill_items", "bk_page", "database", "_db_user", "bk_reset_item_form", "bk_reset_shop_form"]:
                st.session_state[k] = None if k not in ["bk_bill_draft", "bk_load_draft", "bk_bill_items"] else []
            st.session_state["bk_bills_created"] = 0; st.rerun()

def save_booker_draft(shop, items):
    if not st.session_state.get("booker_name"): return
    draft_file = get_booker_draft_file(st.session_state["booker_name"])
    try:
        with open(draft_file, "w", encoding="utf-8") as f: json.dump({"shop": shop, "items": items}, f, indent=4, ensure_ascii=False)
    except Exception: pass

def clear_booker_draft():
    if not st.session_state.get("booker_name"): return
    draft_file = get_booker_draft_file(st.session_state["booker_name"])
    if os.path.exists(draft_file):
        try: os.remove(draft_file)
        except Exception: pass

def render_booker_new_bill():
    bk_name = st.session_state.get("booker_name", "")
    st.markdown(f"<h1 style='color:#0f172a;'>🧾 New Bill — <span style='color:#2563eb;'>{bk_name}</span></h1>", unsafe_allow_html=True)
    st.markdown("---")
    if st.session_state.get("bk_reset_item_form"):
        st.session_state["bk_product_sel"] = "-- Select --"; st.session_state["bk_qty"] = 0; st.session_state["bk_reset_item_form"] = False
    if is_wholesaler(st.session_state.get("bk_current_shop", "")): st.markdown(f"<div class='hint-box'>🏢 Wholesaler detected! WS1, WS2, WS3 discount applies.</div>", unsafe_allow_html=True)
    col_a, col_b = st.columns([3, 1])
    with col_a: shop = st.text_input("🏪 Shop Name", value=st.session_state.get("bk_current_shop", ""), key="bk_shop_input", placeholder="Enter shop name")
    with col_b: st.markdown("<br>", unsafe_allow_html=True)
    if shop:
        new_shop = shop.strip()
        if new_shop != st.session_state.get("bk_current_shop", ""):
            st.session_state["bk_current_shop"] = new_shop; st.session_state["bk_bill_draft"] = []; save_booker_draft(new_shop, st.session_state["bk_bill_draft"]); st.rerun()
    if not shop.strip(): st.warning("⚠️ Enter shop name first"); return
    if is_wholesaler(shop):
        ws_pkg = get_wholesaler_package_from_shop(shop)
        if ws_pkg: st.markdown(f"<div class='hint-box' style='background:#faf5ff;border-left-color:#9333ea;color:#6b21a8;'>🏢 Package: {ws_pkg['name']} — {ws_pkg['discount_pct']}% discount</div>", unsafe_allow_html=True)
    st.markdown("---"); st.markdown("### ➕ Add Item")
    all_products = get_all_products()
    search_txt = st.text_input("🔍 Search Product", key="bk_search", placeholder="Type name...")
    su = search_txt.strip().upper()
    filtered = [p for p in all_products if su in p["name"].upper() or su in str(p["code"]).upper()] if su else all_products
    filtered_labels = [f"{p['name']} (Code: {p['code']})" for p in filtered]
    c1, c2, c3 = st.columns([3, 1, 1])
    with c1: selected = st.selectbox("Product", options=["-- Select --"] + filtered_labels, key="bk_product_sel")
    with c2: boxes = st.number_input("Boxes", min_value=0, step=1, value=0, key="bk_qty")
    with c3: st.markdown("<br>", unsafe_allow_html=True); add_clicked = st.button("➕ Add", key="bk_add_item_btn", use_container_width=True, type="primary")
    selected_product = None
    if selected != "-- Select --":
        for p in all_products:
            if f"{p['name']} (Code: {p['code']})" == selected: selected_product = p; break
    if selected_product:
        tp = get_price(selected_product["code"], selected_product["price"]); st.success(f"💰 {selected_product['name']} — TP: **Rs {tp:,.0f}** | Boxes: {boxes} → Total: **Rs {boxes * tp:,.0f}**")
    if add_clicked:
        if not selected_product: st.error("❌ Select Product")
        elif boxes <= 0: st.error("❌ Enter Boxes")
        else:
            tp = get_price(selected_product["code"], selected_product["price"])
            st.session_state["bk_bill_draft"].append({"Code": selected_product["code"], "Product": selected_product["name"], "Boxes": int(boxes), "TP/Box": float(tp), "Gross": float(boxes * tp), "Discount %": 0.0, "Net": float(boxes * tp)})
            save_booker_draft(shop, st.session_state["bk_bill_draft"]); st.session_state["bk_reset_item_form"] = True; st.rerun()
    st.markdown("---")
    draft = st.session_state.get("bk_bill_draft", [])
    if not draft: st.info("No items added yet.")
    else:
        st.markdown(f"### 📋 Current Bill — {shop}")
        st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"], "TP/Box": it["TP/Box"], "Total": it["Gross"]} for it in draft]), use_container_width=True, hide_index=True)
        c1, c2 = st.columns([3, 1])
        with c1: remove_idx = st.selectbox("Delete item:", options=["--"] + [f"{i+1}. {it['Product']} ({it['Boxes']} boxes)" for i, it in enumerate(draft)], key="bk_remove_sel")
        with c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑 Remove", key="bk_remove_btn", use_container_width=True):
                if remove_idx != "--":
                    idx = int(remove_idx.split(".")[0]) - 1; st.session_state["bk_bill_draft"].pop(idx); save_booker_draft(shop, st.session_state["bk_bill_draft"]); st.rerun()
        gross_total = sum(float(it["Gross"]) for it in draft); st.markdown(f"<div class='summary-box'><b>Gross Total:</b> <span style='color:#2563eb;font-weight:700;font-size:20px;'>Rs {gross_total:,.0f}</span></div>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Done Bill (Save to Admin)", key="bk_done_bill", use_container_width=True, type="primary"): finalize_booker_bill(shop, draft)
        with c2:
            if st.button("🗑 Clear Draft", key="bk_clear_draft", use_container_width=True): st.session_state["bk_bill_draft"] = []; clear_booker_draft(); st.rerun()
    st.markdown("---"); load_draft = st.session_state.get("bk_load_draft", []); total_load_boxes = sum(int(it.get("Boxes", 0)) for it in load_draft)
    st.markdown(f"### 📦 Load Form — Session Total")
    if not load_draft: st.info("No bills completed yet. Create a bill first.")
    else:
        st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"]} for it in load_draft]), use_container_width=True, hide_index=True)
        st.markdown(f"<div class='summary-box'><b>Total Products:</b> {len(load_draft)} | <b>Total Boxes:</b> <span style='color:#2563eb;font-weight:700;font-size:20px;'>{total_load_boxes}</span></div>", unsafe_allow_html=True)
        if st.button("📦 Done Load Form (Save to Admin)", key="bk_done_load", use_container_width=True, type="primary"): finalize_booker_load_form()

def finalize_booker_bill(shop, draft):
    global db
    bill_no = db["next_bill_no"]; date_str = datetime.now().strftime("%d-%m-%Y"); bk_name = st.session_state.get("booker_name", "")
    for it in draft: db["bills"].append({"Bill No": bill_no, "Date": date_str, "Shop": shop, "Order Booker": bk_name, "Salesman": "", "Delivery Man": "", "Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"], "TP/Box": it["TP/Box"], "Discount %": it["Discount %"], "Gross": it["Gross"], "Net": it["Net"]})
    db["next_bill_no"] += 1; save_database(db)
    for it in draft:
        found = False
        for ld in st.session_state["bk_load_draft"]:
            if str(ld["Code"]) == str(it["Code"]): ld["Boxes"] += int(it["Boxes"]); found = True; break
        if not found: st.session_state["bk_load_draft"].append({"Code": it["Code"], "Product": it["Product"], "Boxes": int(it["Boxes"])})
    bill_total_net = sum(float(it["Net"]) for it in draft); pkg_pct, _, _ = get_effective_discount_pct(shop, bill_total_net); after_disc_total = bill_total_net - (bill_total_net * pkg_pct / 100)
    st.session_state["bk_bill_draft"] = []; clear_booker_draft(); st.session_state["bk_bills_created"] = st.session_state.get("bk_bills_created", 0) + 1
    st.session_state["success_msg"] = f"✅ Bill #{bill_no} saved | {shop} | Net: Rs {after_disc_total:,.0f}"; st.rerun()

def finalize_booker_load_form():
    global db
    load_draft = st.session_state.get("bk_load_draft", [])
    if not load_draft: st.session_state["error_msg"] = "❌ Load form empty"; return
    bk_name = st.session_state.get("booker_name", ""); today_str = datetime.now().strftime("%d-%m-%Y"); total_boxes = sum(int(it["Boxes"]) for it in load_draft)
    existing_lf = None
    for lf in db.get("load_forms", []):
        if lf.get("booker") == bk_name and lf.get("date") == today_str and not lf.get("transferred_to_dsr"): existing_lf = lf; break
    if existing_lf:
        existing_lf["items"] = load_draft; existing_lf["total_boxes"] = total_boxes; existing_lf["total_products"] = len(load_draft); existing_lf["time"] = datetime.now().strftime("%H:%M"); existing_lf["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        msg = f"🔄 Load Form #{existing_lf['id']} updated"
    else:
        next_id = max([lf.get("id", 0) for lf in db["load_forms"]] + [0]) + 1
        db["load_forms"].append({"id": next_id, "date": today_str, "time": datetime.now().strftime("%H:%M"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "booker": bk_name, "items": load_draft, "total_boxes": total_boxes, "total_products": len(load_draft), "transferred_to_dsr": False, "refreshed": False, "source": "booker"})
        msg = f"✅ Load Form #{next_id} saved"
    save_database(db); st.session_state["bk_load_draft"] = []; st.session_state["success_msg"] = msg; st.rerun()

def render_booker_all_products():
    st.markdown(f"<h1 style='color:#0f172a;'>🛒 All Products with TP</h1>", unsafe_allow_html=True); st.markdown("---")
    all_products = get_all_products(); search = st.text_input("🔍 Search", key="bk_prod_search", placeholder="Type name...")
    su = search.strip().upper(); shown = [p for p in all_products if su in p["name"].upper() or su in str(p["code"]).upper()] if su else all_products
    st.markdown(f"<div class='summary-box'><b>Total Products:</b> {len(all_products)} | <b>Showing:</b> {len(shown)}</div>", unsafe_allow_html=True)
    if not shown: st.info("No products found."); return
    for p in shown:
        tp = get_price(p["code"], p["price"]); st.markdown(f"""<div class="lf-simple-card"><div class="lf-info"><div class="lf-line1">{p['name']}</div><div class="lf-line2">Code: {p['code']}</div></div><div class="lf-boxes" style="background:#dcfce7;color:#16a34a;">Rs {tp:,.0f}</div></div>""", unsafe_allow_html=True)

def render_booker_load_status():
    st.markdown(f"<h1 style='color:#0f172a;'>📦 Load Form Status</h1>", unsafe_allow_html=True); st.markdown("---")
    load_draft = st.session_state.get("bk_load_draft", [])
    if not load_draft: st.info("No bills completed yet."); return
    total_boxes = sum(int(it["Boxes"]) for it in load_draft)
    st.markdown(f"<div class='summary-box'><b>Products:</b> {len(load_draft)} &nbsp;|&nbsp; <b>Total Boxes:</b> <span style='color:#2563eb;font-size:18px;font-weight:700;'>{total_boxes}</span></div>", unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"]} for it in load_draft]), use_container_width=True, hide_index=True)
    if st.button("📦 Done Load Form", key="bk_done_load_2", use_container_width=True, type="primary"): finalize_booker_load_form()

def render_booker_bills_history():
    bk_name = st.session_state.get("booker_name", "")
    st.markdown(f"<h1 style='color:#0f172a;'>📋 My Bills History</h1>", unsafe_allow_html=True); st.markdown("---")
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Filter:", ["📅 Aaj (Today)", "📆 Custom Range", "🗓️ Specific", "📋 All Bills"], key="bk_history_filter")
    with c2: from_date = st.date_input("From:", value=today - timedelta(days=7), key="bk_history_from")
    with c3: to_date = st.date_input("To:", value=today, key="bk_history_to")
    groups = {}
    for b in db["bills"]:
        if str(b.get("Order Booker", "")).strip() != bk_name: continue
        bdate = parse_date(b.get("Date", ""))
        if filter_mode == "📅 Aaj (Today)" and bdate != today: continue
        elif filter_mode == "🗓️ Specific" and bdate != from_date: continue
        elif filter_mode == "📆 Custom Range" and (bdate is None or not (from_date <= bdate <= to_date)): continue
        key = (b.get("Shop",""), b.get("Date",""), b.get("Bill No",""))
        if key not in groups: groups[key] = {"shop": b.get("Shop",""), "date": b.get("Date",""), "bill_no": b.get("Bill No",""), "items": [], "idx": []}
        groups[key]["items"].append(b); groups[key]["idx"].append(db["bills"].index(b))
    if not groups: st.info("No bills found in this filter."); return
    for key, g in sorted(groups.items(), key=lambda kv: kv[0][1], reverse=True):
        shop = g["shop"]; dt = g["date"]; bn = g["bill_no"]; tb = sum(int(it.get("Boxes",0)) for it in g["items"]); t_net = sum(float(it.get("Net",0)) for it in g["items"])
        st.markdown(f"""<div class='lf-simple-card'><div class='lf-info'><div class='lf-line1'>🏪 {shop}</div><div class='lf-line2'>📅 {dt} · 🧾 #{bn} · 💰 Rs {t_net:,.0f}</div></div><div class='lf-boxes'>{tb}<small>BOXES</small></div></div>""", unsafe_allow_html=True)
        with st.expander("View Items"):
            st.dataframe(pd.DataFrame([{"Code": i["Code"], "Product": i["Product"], "Boxes": i["Boxes"], "Net": i["Net"]} for i in g["items"]]), use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🗑 Delete Bill", key=f"bk_del_{key}_{dt}", type="primary", use_container_width=True):
                idxs = set(g["idx"]); db["bills"] = [b for i, b in enumerate(db["bills"]) if i not in idxs]; save_database(db); st.session_state["success_msg"] = f"🗑 Bill #{bn} deleted"; st.rerun()
        with c2:
            if st.button("📤 Resend to Admin", key=f"bk_resend_{key}_{dt}", use_container_width=True):
                save_database(db); st.session_state["success_msg"] = f"✅ Bill #{bn} resent"; st.rerun()
        st.markdown("---")

# ============================================================
# ROUTER
# ============================================================
if ROLE == "booker":
    render_booker_sidebar()
    bk_page = st.session_state.get("bk_page", "🧾 New Bill")
    if bk_page == "🧾 New Bill": render_booker_new_bill()
    elif bk_page == "📋 My Bills History": render_booker_bills_history()
    elif bk_page == "🛒 All Products": render_booker_all_products()
    elif bk_page == "📦 Load Form Status": render_booker_load_status()
    if st.session_state.get("success_msg"): st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
    if st.session_state.get("error_msg"): st.error(st.session_state["error_msg"]); st.session_state["error_msg"] = None
    st.stop()

# ============================================================
# ADMIN SIDEBAR
# ============================================================
with st.sidebar:
    display_name = st.session_state.get("display_name", CURRENT_USER); avatar_letter = (display_name[0] if display_name else "A").upper()
    st.markdown(f"""<div class="sb-brand"><div class="sb-brand-logo">🧾</div><div class="sb-brand-name">AL-BARAKAH</div><div class="sb-brand-sub">ENTERPRISES</div></div>""", unsafe_allow_html=True)
    st.markdown(f"""<div class="sb-user"><div class="sb-user-avatar">{avatar_letter}</div><div class="sb-user-info"><div class="sb-user-label">Admin</div><div class="sb-user-name">{display_name}</div></div></div>""", unsafe_allow_html=True)
    st.markdown('<div class="sb-section-label">Main Menu</div>', unsafe_allow_html=True)
    page = st.radio("MENU", ["📊 Dashboard", "🧾 Billing", "🛒 All Products", "🎁 Discount", "👤 Bookers", "📈 Booker Bills", "💰 Bookers Salary", "🧑‍💼 Salesmen", "💰 Salesmen Salary", "💵 Daily Expense", "📋 Bills List", "💳 Credit Bills", "📦 Load Form", "📋 DSR", "🧮 Calculation"], key="page_selector", label_visibility="collapsed")
    st.session_state["page"] = page
    st.markdown(f'<div class="sb-date" style="text-align:center;padding:16px;font-size:11px;color:#94a3b8;font-weight:600;">📅 {datetime.now().strftime("%A, %d %b %Y")}</div>', unsafe_allow_html=True)
    all_prod_count = len(get_all_products()); pending_credit = sum(1 for c in db.get("credit_bills", []) if c.get("status") == "pending"); pending_amt = sum(float(c.get("total_net", 0)) for c in db.get("credit_bills", []) if c.get("status") == "pending")
    st.markdown(f"""<div class="sb-stats"><div class="sb-stats-grid"><div class="sb-stat-card"><div class="sb-stat-icon">📦</div><div class="sb-stat-value">{all_prod_count}</div><div class="sb-stat-label">Products</div></div><div class="sb-stat-card"><div class="sb-stat-icon">🧾</div><div class="sb-stat-value">{len(db['bills'])}</div><div class="sb-stat-label">Bills</div></div><div class="sb-stat-card"><div class="sb-stat-icon">👤</div><div class="sb-stat-value">{len(db.get('bookers', []))}</div><div class="sb-stat-label">Bookers</div></div><div class="sb-stat-card"><div class="sb-stat-icon">💳</div><div class="sb-stat-value">{pending_credit}</div><div class="sb-stat-label">Cr. Pend</div></div></div></div>""", unsafe_allow_html=True)
    if st.button("🚪 Logout", key="btn_logout", use_container_width=True):
        for k in ["logged_in_user", "display_name", "role", "database", "_db_user"]: st.session_state[k] = None
        st.session_state["page"] = "📊 Dashboard"; st.rerun()

# ============================================================
# ADMIN PAGES
# ============================================================
def render_admin_dashboard():
    display_name = st.session_state.get("display_name", CURRENT_USER); today_str = date.today().strftime("%A, %d %B %Y")
    bookers = db.get("bookers", []); salesmen = db.get("salesmen", []); total_products = len(get_all_products()); total_bills = len(db["bills"])
    pending_credits = [c for c in db.get("credit_bills", []) if c.get("status") == "pending"]; paid_credits = [c for c in db.get("credit_bills", []) if c.get("status") == "paid"]
    pending_amt = sum(float(c.get("total_net", 0)) for c in pending_credits); paid_amt = sum(float(c.get("total_net", 0)) for c in paid_credits)
    st.markdown(f"""<div class="dash-hero"><div class="dash-hero-content"><div class="dash-hero-left"><div style="font-size:13px;font-weight:600;color:#94a3b8;margin-bottom:8px;">WELCOME BACK</div><h1 class="dash-hero-title">{display_name}</h1><div class="dash-hero-sub">🏢 {COMPANY_NAME} &nbsp;·&nbsp; 📅 {today_str}</div></div><div class="dash-hero-right" style="display:flex;gap:32px;"><div class="dash-hero-stat"><div class="dash-hero-stat-label">Total Bills</div><div class="dash-hero-stat-val">{total_bills}</div></div><div class="dash-hero-stat"><div class="dash-hero-stat-label">Products</div><div class="dash-hero-stat-val">{total_products}</div></div></div></div></div>""", unsafe_allow_html=True)
    st.markdown('<div class="dash-section-title" style="font-size:18px;font-weight:700;color:#0f172a;margin-bottom:16px;">Business Overview</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon">👤</div><div class="kpi-label">Bookers</div></div><div class="kpi-value-group"><div class="kpi-value">{len(bookers)}</div><div class="kpi-value-unit">members</div></div><div class="kpi-sub blue">Booking team</div></div>""", unsafe_allow_html=True)
    with c2: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon purple">🧑‍💼</div><div class="kpi-label">Salesmen</div></div><div class="kpi-value-group"><div class="kpi-value">{len(salesmen)}</div><div class="kpi-value-unit">members</div></div><div class="kpi-sub blue">Field team</div></div>""", unsafe_allow_html=True)
    with c3: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon teal">📦</div><div class="kpi-label">Products</div></div><div class="kpi-value-group"><div class="kpi-value">{total_products}</div><div class="kpi-value-unit">items</div></div><div class="kpi-sub blue">Active catalog</div></div>""", unsafe_allow_html=True)
    with c4: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon">🧾</div><div class="kpi-label">Total Bills</div></div><div class="kpi-value-group"><div class="kpi-value">{total_bills}</div><div class="kpi-value-unit">bills</div></div><div class="kpi-sub blue">Generated</div></div>""", unsafe_allow_html=True)
    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="dash-section-title" style="font-size:18px;font-weight:700;color:#0f172a;margin-bottom:16px;">Credit & Operations</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon orange">⏳</div><div class="kpi-label">Pending Credit</div></div><div class="kpi-value-group"><div class="kpi-value">{len(pending_credits)}</div><div class="kpi-value-unit">bills</div></div><div class="kpi-sub orange">Rs {pending_amt:,.0f} receivable</div></div>""", unsafe_allow_html=True)
    with c2: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon green">✅</div><div class="kpi-label">Paid Credit</div></div><div class="kpi-value-group"><div class="kpi-value">{len(paid_credits)}</div><div class="kpi-value-unit">bills</div></div><div class="kpi-sub green">Rs {paid_amt:,.0f} settled</div></div>""", unsafe_allow_html=True)
    with c3: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon">📦</div><div class="kpi-label">Load Forms</div></div><div class="kpi-value-group"><div class="kpi-value">{len(db.get('load_forms', []))}</div><div class="kpi-value-unit">saved</div></div><div class="kpi-sub blue">Stock dispatch</div></div>""", unsafe_allow_html=True)
    with c4: st.markdown(f"""<div class="kpi-card"><div class="kpi-top"><div class="kpi-icon red">📋</div><div class="kpi-label">DSR Forms</div></div><div class="kpi-value-group"><div class="kpi-value">{len(db.get('dsr_forms', []))}</div><div class="kpi-value-unit">active</div></div><div class="kpi-sub blue">Daily reports</div></div>""", unsafe_allow_html=True)
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True); st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div style="font-size:16px;font-weight:700;color:#0f172a;margin-bottom:16px;">👤 Order Booker Team</div>', unsafe_allow_html=True)
        if not bookers: st.info("No bookers added yet.")
        else:
            for b_name in bookers:
                sd = db.get("bookers_salaries", {}).get(b_name, {}); base = sd.get("base_salary", 0); txns = sd.get("transactions", [])
                adv_p = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status", "pending") == "pending"); short_p = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status", "pending") == "pending"); remaining = base - adv_p - short_p
                sub_text = 'Salary not set' if base == 0 else f'Remaining: <b>Rs {remaining:,.0f}</b>'
                al = (b_name[0] if b_name else "?").upper()
                st.markdown(f"""<div class="team-card"><div class="team-avatar booker">{al}</div><div class="team-info"><div class="team-name">{b_name}</div><div class="team-meta">{sub_text}</div></div><div class="team-badge">BOOKER</div></div>""", unsafe_allow_html=True)
    with c2:
        st.markdown('<div style="font-size:16px;font-weight:700;color:#0f172a;margin-bottom:16px;">🧑‍💼 Salesmen Team</div>', unsafe_allow_html=True)
        if not salesmen: st.info("No salesmen added yet.")
        else:
            for s_name in salesmen:
                sd = db.get("salesmen_salaries", {}).get(s_name, {}); base = sd.get("base_salary", 0); txns = sd.get("transactions", [])
                adv_p = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status", "pending") == "pending"); short_p = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status", "pending") == "pending"); remaining = base - adv_p - short_p
                sub_text = 'Salary not set' if base == 0 else f'Remaining: <b>Rs {remaining:,.0f}</b>'
                al = (s_name[0] if s_name else "?").upper()
                st.markdown(f"""<div class="team-card"><div class="team-avatar salesman">{al}</div><div class="team-info"><div class="team-name">{s_name}</div><div class="team-meta">{sub_text}</div></div><div class="team-badge" style="background:#faf5ff;color:#9333ea;">SALESMAN</div></div>""", unsafe_allow_html=True)

def render_admin_booker_bills():
    st.markdown(f"<h1 style='color:#0f172a;'>📈 Booker Bills & Returns</h1>", unsafe_allow_html=True); st.markdown("---")
    bookers = db.get("bookers", [])
    if not bookers: st.info("Add bookers first."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Filter Mode:", ["📅 Aaj (Today)", "📆 Custom Range", "🗓️ Specific", "📋 All"], key="bb_filter_mode")
    with c2: from_date = st.date_input("From:", value=today, key="bb_from_date")
    with c3: to_date = st.date_input("To:", value=today, key="bb_to_date")
    def in_range(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if filter_mode == "📅 Aaj (Today)": return d == today
        elif filter_mode == "🗓️ Specific": return d == from_date
        elif filter_mode == "📆 Custom Range": return from_date <= d <= to_date
        else: return True
    grand_sold = 0; grand_ret = 0; grand_net = 0
    for bk in bookers:
        bk_bills = [b for b in db.get("bills", []) if str(b.get("Order Booker", "")).strip() == bk and in_range(b.get("Date", ""))]
        bk_returns = get_booker_returns_list(bk, in_range); total_sold = sum(int(b.get("Boxes", 0)) for b in bk_bills); total_ret = get_booker_total_return_boxes(bk, in_range); net_boxes = total_sold - total_ret
        grand_sold += total_sold; grand_ret += total_ret; grand_net += net_boxes
        shop_groups = {}
        for b in bk_bills:
            shop = b.get("Shop", "-") or "-"
            if shop not in shop_groups: shop_groups[shop] = {"boxes": 0, "items": [], "bill_nos": set(), "dates": set()}
            shop_groups[shop]["boxes"] += int(b.get("Boxes", 0)); shop_groups[shop]["items"].append(b); shop_groups[shop]["bill_nos"].add(str(b.get("Bill No", ""))); shop_groups[shop]["dates"].add(b.get("Date", ""))
        expander_label = f"👤 {bk} — Sold: {total_sold} | Returns: {total_ret} | 🟢 Net: {net_boxes} boxes"
        with st.expander(expander_label, expanded=False):
            st.markdown(f"""<div class='bb-summary'><div class='bb-stat'><div class='bb-stat-label'>Total Sold</div><div class='bb-stat-val'>{total_sold}</div></div><div class='bb-stat'><div class='bb-stat-label'>Returns</div><div class='bb-stat-val red'>-{total_ret}</div></div><div class='bb-stat'><div class='bb-stat-label'>Net Boxes</div><div class='bb-stat-val green'>{net_boxes}</div></div><div class='bb-stat'><div class='bb-stat-label'>Shops</div><div class='bb-stat-val' style='font-size:20px;'>{len(shop_groups)}</div></div></div>""", unsafe_allow_html=True)
            st.markdown("**🏪 Shop-wise Bills**")
            if not shop_groups: st.info("No bills found.")
            else:
                for shop, info in sorted(shop_groups.items()):
                    st.markdown(f"""<div class='lf-simple-card'><div class='lf-info'><div class='lf-line1'>🏪 {shop}</div><div class='lf-line2'>🧾 {len(info['bill_nos'])} bill(s) · 📅 {', '.join(sorted(info['dates']))}</div></div><div class='lf-boxes'>{info['boxes']}<small>BOXES</small></div></div>""", unsafe_allow_html=True)
                    with st.expander(f"📋 {shop} — full items ({len(info['items'])})"):
                        st.dataframe(pd.DataFrame([{"Bill No": b.get("Bill No"), "Date": b.get("Date"), "Code": b.get("Code"), "Product": b.get("Product"), "Boxes": b.get("Boxes"), "TP/Box": b.get("TP/Box"), "Net": b.get("Net")} for b in info["items"]]), use_container_width=True, hide_index=True)
            st.markdown("---"); st.markdown("**↩️ Returns (item-wise)**")
            if bk_returns:
                for r in sorted(bk_returns, key=lambda x: x.get("created_at", ""), reverse=True):
                    lf_id = r.get("load_form_id", "-"); items = r.get("items", []); ret_total = sum(int(it.get("ReturnBoxes", 0)) for it in items) if items else int(r.get("boxes", 0))
                    st.markdown(f"""<div class='lf-simple-card ret'><div class='lf-info'><div class='lf-line1'>↩️ Return — {ret_total} boxes <span class='badge-return'>LF #{lf_id}</span></div><div class='lf-line2'>📅 {r.get('date','')} · {r.get('note','') or '—'}</div></div><div class='lf-boxes' style="background:#fee2e2;color:#dc2626;">{ret_total}<small>RET</small></div></div>""", unsafe_allow_html=True)
                    if items: st.dataframe(pd.DataFrame([{"Code": it.get("Code", ""), "Product": it.get("Product", ""), "Return Boxes": int(it.get("ReturnBoxes", 0))} for it in items if int(it.get("ReturnBoxes", 0)) > 0]), use_container_width=True, hide_index=True)
            st.markdown("---"); st.markdown("**🧾 All Bills**")
            if bk_bills: st.dataframe(pd.DataFrame([{"Bill No": b.get("Bill No"), "Date": b.get("Date"), "Shop": b.get("Shop"), "Code": b.get("Code"), "Product": b.get("Product"), "Boxes": b.get("Boxes"), "TP/Box": b.get("TP/Box"), "Gross": b.get("Gross"), "Net": b.get("Net")} for b in bk_bills]), use_container_width=True, hide_index=True)
    st.markdown("---")
    st.markdown(f"""<div class='bb-summary'><div class='bb-stat'><div class='bb-stat-label'>All Bookers Sold</div><div class='bb-stat-val'>{grand_sold}</div></div><div class='bb-stat'><div class='bb-stat-label'>Total Returns</div><div class='bb-stat-val red'>-{grand_ret}</div></div><div class='bb-stat'><div class='bb-stat-label'>NET</div><div class='bb-stat-val green'>{grand_net}</div></div></div>""", unsafe_allow_html=True)

def render_admin_discount():
    st.markdown(f"<h1 style='color:#0f172a;'>🎁 Discount Settings</h1>", unsafe_allow_html=True); st.markdown("---")
    st.markdown("### 🏢 Wholesaler Packages (WS1, WS2, WS3)")
    for pkg in db.get("wholesaler_packages", []):
        with st.expander(f"🏢 {pkg['name']} — Current: {pkg['discount_pct']}%"):
            c1, c2 = st.columns([3, 1])
            with c1: np = st.number_input(f"{pkg['name']} Discount %:", value=float(pkg['discount_pct']), min_value=0.0, max_value=100.0, step=0.5, key=f"ws_pct_{pkg['id']}")
            with c2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Save", key=f"ws_save_{pkg['id']}", type="primary"):
                    for pp in db["wholesaler_packages"]:
                        if pp["id"] == pkg["id"]: pp["discount_pct"] = float(np); break
                    save_database(db); st.session_state["success_msg"] = f"✅ {pkg['name']} saved"; st.rerun()
    st.markdown("---"); st.markdown("### 🎁 Regular Discount Packages")
    for i, pkg in enumerate(db.get("discount_packages", [])):
        pid = pkg.get("id", i+1)
        with st.expander(f"🎁 {pkg.get('name', 'Package')} {'✅' if pkg.get('active') else '⭕'}"):
            c1, c2 = st.columns([1, 5])
            with c1: new_active = st.checkbox("Active", value=pkg.get("active", False), key=f"active_{pid}")
            with c2: new_name = st.text_input("Name:", value=pkg.get("name", ""), key=f"pkgname_{pid}")
            t1, t2, t3 = st.columns(3)
            with t1: ta1 = st.number_input("Min Rs (T1)", value=float(pkg.get("tier1_amount", 0)), key=f"t1a_{pid}"); tp1 = st.number_input("% T1", value=float(pkg.get("tier1_pct", 0)), key=f"t1p_{pid}")
            with t2: ta2 = st.number_input("Min Rs (T2)", value=float(pkg.get("tier2_amount", 0)), key=f"t2a_{pid}"); tp2 = st.number_input("% T2", value=float(pkg.get("tier2_pct", 0)), key=f"t2p_{pid}")
            with t3: ta3 = st.number_input("Min Rs (T3)", value=float(pkg.get("tier3_amount", 0)), key=f"t3a_{pid}"); tp3 = st.number_input("% T3", value=float(pkg.get("tier3_pct", 0)), key=f"t3p_{pid}")
            if st.button("💾 Save", key=f"savepkg_{pid}", type="primary"):
                for pp in db["discount_packages"]:
                    if pp.get("id") == pid:
                        pp["name"] = new_name; pp["active"] = new_active; pp["tier1_amount"] = float(ta1); pp["tier1_pct"] = float(tp1); pp["tier2_amount"] = float(ta2); pp["tier2_pct"] = float(tp2); pp["tier3_amount"] = float(ta3); pp["tier3_pct"] = float(tp3); break
                save_database(db); st.session_state["success_msg"] = "✅ Saved"; st.rerun()

def render_admin_products():
    st.markdown(f"<h1 style='color:#0f172a;'>🛒 All Products</h1>", unsafe_allow_html=True); st.markdown("---")
    all_products = get_all_products()
    with st.expander("➕ Naya Product (Single)", expanded=False):
        c1, c2, c3, c4 = st.columns([1.5, 3.5, 1.5, 1])
        with c1: nc = st.text_input("Code", key="new_prod_code")
        with c2: nn = st.text_input("Name", key="new_prod_name")
        with c3: npr = st.number_input("Price", min_value=0.0, step=1.0, key="new_prod_price", value=0.0)
        with c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕ Add", key="btn_add_new_product", type="primary"):
                cs = str(nc).strip(); ns = str(nn).strip()
                if not cs or not ns or npr <= 0: st.error("❌ Fill all fields")
                elif cs in {str(p["code"]) for p in all_products}: st.error("❌ Code exists")
                else:
                    db["custom_products"].append({"code": cs, "name": ns, "price": float(npr), "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}); save_database(db); st.session_state["success_msg"] = f"✅ Added"; st.rerun()
    with st.expander("📤 Bulk Upload", expanded=False):
        up = st.file_uploader("Excel/CSV", type=["xlsx","xls","csv"], key="bulk_upload_file")
        if up:
            try:
                df_u = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up); df_u.columns = [str(c).strip().lower() for c in df_u.columns]
                if "code" in df_u.columns and "name" in df_u.columns and "price" in df_u.columns:
                    st.dataframe(df_u.head(10), use_container_width=True)
                    if st.button("📥 Import", key="btn_bulk_import", type="primary"):
                        existing = {str(p["code"]) for p in all_products}; added = 0; updated = 0
                        for _, r in df_u.iterrows():
                            try:
                                c_raw = r["code"]
                                if pd.isna(c_raw): continue
                                c_s = str(int(float(c_raw))) if isinstance(c_raw, (int, float)) or str(c_raw).replace('.','',1).isdigit() else str(c_raw).strip()
                                n_s = str(r["name"]).strip(); p_f = float(r["price"])
                            except Exception: continue
                            if not c_s or not n_s or p_f < 0: continue
                            if c_s in existing:
                                for cp in db.get("custom_products", []):
                                    if str(cp.get("code")) == c_s: cp["price"] = p_f; cp["name"] = n_s; updated += 1; break
                                else: db["custom_products"].append({"code": c_s, "name": n_s, "price": p_f, "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}); updated += 1
                            else: db["custom_products"].append({"code": c_s, "name": n_s, "price": p_f, "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}); existing.add(c_s); added += 1
                        save_database(db); st.session_state["success_msg"] = f"✅ {added} Added, {updated} Updated"; st.rerun()
                else: st.error("❌ File must have 'code', 'name', and 'price' columns.")
            except Exception as e: st.error(f"❌ Error reading file: {e}")
    st.markdown(f"""<div class="danger-box"><b>⚠️ DANGER ZONE — Product Delete Options</b><br><span style="color:#64748b;font-size:12px;">Total active: <b>{len(all_products)}</b></span></div>""", unsafe_allow_html=True)
    dz1, dz2 = st.columns([1, 1])
    with dz1: select_mode = st.toggle("☑️ Select to Delete Mode", value=st.session_state.get("prod_select_mode", False), key="prod_select_toggle"); st.session_state["prod_select_mode"] = select_mode
    with dz2:
        if st.button("🗑️ DELETE ALL PRODUCTS", type="primary", use_container_width=True): st.session_state["prod_confirm_all"] = True; st.rerun()
    if st.session_state.get("prod_confirm_all"):
        st.error("⚠️ **Are you sure? All products will be permanently deleted.**")
        cc1, cc2 = st.columns(2)
        with cc1:
            if st.button("✅ YES, DELETE ALL", type="primary", use_container_width=True): db["custom_products"] = []; db["product_prices"] = {}; save_database(db); st.session_state["prod_confirm_all"] = False; st.session_state["success_msg"] = f"🗑️ All products deleted"; st.rerun()
        with cc2:
            if st.button("❌ Cancel", use_container_width=True): st.session_state["prod_confirm_all"] = False; st.rerun()
    st.markdown("---"); st.markdown(f"### 📋 Products ({len(all_products)})")
    search = st.text_input("🔍 Search", key="prod_search"); su = search.strip().upper(); shown = [p for p in all_products if su in p["name"].upper() or su in str(p["code"])] if su else all_products
    edited_prices = db.get("product_prices", {}); custom_codes = {str(cp.get("code", "")) for cp in db.get("custom_products", [])}
    if not shown: st.info("No products found.")
    if "prod_selected_codes" not in st.session_state: st.session_state["prod_selected_codes"] = []
    if st.session_state.get("prod_select_mode", False):
        sa1, sa2, sa3 = st.columns(3)
        with sa1:
            if st.button("☑️ Select All Visible", use_container_width=True): st.session_state["prod_selected_codes"] = sorted({str(p["code"]) for p in shown}); st.rerun()
        with sa2:
            if st.button("⬜ Clear Selection", use_container_width=True): st.session_state["prod_selected_codes"] = []; st.rerun()
        with sa3:
            sel_count = len(st.session_state.get("prod_selected_codes", []))
            if st.button(f"🗑️ Delete Selected ({sel_count})", type="primary", use_container_width=True):
                if not sel_count: st.session_state["error_msg"] = "❌ Select products first"
                else:
                    sel_codes = list(st.session_state["prod_selected_codes"]); db["custom_products"] = [cp for cp in db.get("custom_products", []) if str(cp.get("code", "")) not in set(sel_codes)]
                    for c in sel_codes: db.get("product_prices", {}).pop(c, None)
                    save_database(db); st.session_state["prod_selected_codes"] = []; st.session_state["success_msg"] = f"🗑️ {sel_count} products deleted"; st.rerun()
    for p in shown[:200]:
        code = str(p["code"]); base = float(p["price"]); cur = get_price(code, base); is_custom = code in custom_codes; is_edited = code in edited_prices; selected_now = code in st.session_state.get("prod_selected_codes", [])
        if st.session_state.get("prod_select_mode", False):
            c0, c1, c2, c3, c4 = st.columns([0.5, 1, 4, 2, 2])
            with c0:
                chk = st.checkbox("", value=selected_now, key=f"chk_{code}", label_visibility="collapsed")
                if chk and not selected_now: st.session_state["prod_selected_codes"].append(code); st.rerun()
                elif (not chk) and selected_now: st.session_state["prod_selected_codes"].remove(code); st.rerun()
        else: c1, c2, c3, c4 = st.columns([1, 4, 2, 2])
        with c1: st.markdown(f"<div style='padding-top:8px;color:#2563eb;font-weight:600;'>{code}</div>", unsafe_allow_html=True)
        with c2: marks = (" 🆕" if is_custom else "") + (" ✏️" if is_edited else ""); st.markdown(f"<div style='padding-top:8px;color:#0f172a;font-weight:500;'>{p['name']}{marks}</div>", unsafe_allow_html=True)
        with c3: np = st.number_input("Price", value=float(cur), min_value=0.0, step=1.0, key=f"price_{code}", label_visibility="collapsed")
        with c4:
            s1, s2 = st.columns(2)
            with s1:
                if st.button("💾", key=f"save_price_{code}", use_container_width=True):
                    if float(np) == base: db.get("product_prices", {}).pop(code, None)
                    else: db.setdefault("product_prices", {})[code] = float(np)
                    save_database(db); st.session_state["success_msg"] = "✅ Saved"; st.rerun()
            with s2:
                if st.button("🗑", key=f"del_{code}", use_container_width=True):
                    db["custom_products"] = [cp for cp in db.get("custom_products", []) if str(cp.get("code", "")) != code]; db.get("product_prices", {}).pop(code, None)
                    save_database(db); st.session_state["success_msg"] = f"🗑 Deleted {code}"; st.rerun()

def render_admin_bookers():
    st.markdown(f"<h1 style='color:#0f172a;'>👤 Bookers</h1>", unsafe_allow_html=True); st.markdown("---")
    c1, c2 = st.columns([3, 1])
    with c1: nb = st.text_input("New Booker Name", key="new_booker_name", label_visibility="collapsed", placeholder="Enter booker name")
    with c2: add = st.button("➕ Add Booker", key="btn_add_booker", type="primary", use_container_width=True)
    if add:
        name = nb.strip()
        if not name: st.error("❌ Name required")
        elif name in db.get("bookers", []): st.warning("⚠️ Already exists")
        else:
            db.setdefault("bookers", []).append(name); db["bookers"] = sorted(db["bookers"]); save_database(db)
            registry = load_bookers(); bk_uname = sanitize_username(name)
            if bk_uname not in registry: registry[bk_uname] = {"display_name": name, "username": bk_uname, "admin": CURRENT_USER, "password_hash": hash_password("1234"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}; save_bookers(registry)
            st.session_state["success_msg"] = f"✅ '{name}' added | Default password: 1234"; st.rerun()
    st.markdown("---"); st.markdown(f"### 📋 Saved Bookers ({len(db.get('bookers', []))})")
    registry = load_bookers()
    for i, bk in enumerate(db.get("bookers", [])):
        c1, c2, c3 = st.columns([4, 1, 1])
        with c1:
            bk_uname = sanitize_username(bk); has_login = bk_uname in registry and registry[bk_uname].get("admin") == CURRENT_USER
            status = "🔐 Login Active" if has_login else "⚠️ No Login"
            st.markdown(f"<div class='booker-row'><b style='font-size:15px;color:#0f172a;'>{bk}</b> <span style='color:#64748b;font-size:12px;margin-left:12px;'>{status}</span></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🔑 Set Password", key=f"setpw_{i}_{bk}", use_container_width=True): st.session_state["setpw_booker"] = bk
        with c3:
            if st.button("🗑 Delete", key=f"del_bk_{i}_{bk}", use_container_width=True):
                db["bookers"].remove(bk); save_database(db); bk_uname = sanitize_username(bk)
                if bk_uname in registry: del registry[bk_uname]; save_bookers(registry)
                st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()
    target = st.session_state.get("setpw_booker")
    if target:
        st.markdown("---"); st.markdown(f"### 🔑 Set Password for: **{target}**")
        new_pw = st.text_input("New Password", type="password", key="setpw_new")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save Password", key="setpw_save", type="primary", use_container_width=True):
                if len(new_pw) < 4: st.error("❌ Min 4 characters")
                else:
                    registry = load_bookers(); bk_uname = sanitize_username(target)
                    if bk_uname not in registry: registry[bk_uname] = {"display_name": target, "username": bk_uname, "admin": CURRENT_USER, "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                    registry[bk_uname]["password_hash"] = hash_password(new_pw); registry[bk_uname]["admin"] = CURRENT_USER; registry[bk_uname]["display_name"] = target
                    save_bookers(registry); st.session_state["setpw_booker"] = None; st.session_state["success_msg"] = f"✅ Password set for {target}"; st.rerun()
        with c2:
            if st.button("❌ Cancel", key="setpw_cancel", use_container_width=True): st.session_state["setpw_booker"] = None; st.rerun()

def render_admin_salesmen():
    st.markdown(f"<h1 style='color:#0f172a;'>🧑‍💼 Salesmen</h1>", unsafe_allow_html=True); st.markdown("---")
    c1, c2 = st.columns([3, 1])
    with c1: ns = st.text_input("New Salesman", key="new_salesman_name", label_visibility="collapsed", placeholder="Enter salesman name")
    with c2: add = st.button("➕ Add Salesman", key="btn_add_salesman", type="primary", use_container_width=True)
    if add:
        name = ns.strip()
        if not name: st.error("❌ Name required")
        elif name in db.get("salesmen", []): st.warning("⚠️ Already exists")
        else: db.setdefault("salesmen", []).append(name); db["salesmen"] = sorted(db["salesmen"]); save_database(db); st.session_state["success_msg"] = f"✅ Added"; st.rerun()
    st.markdown("---")
    for i, sm in enumerate(db.get("salesmen", [])):
        c1, c2 = st.columns([5, 1])
        with c1: st.markdown(f"<div class='booker-row'><b style='font-size:15px;color:#0f172a;'>{sm}</b></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑", key=f"del_sm_{i}_{sm}"): db["salesmen"].remove(sm); save_database(db); st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()

def render_admin_salaries(role_type):
    if role_type == "bookers": title = "💰 Bookers Salary"; emoji = "👤"; names = db.get("bookers", []); sal_key = "bookers_salaries"
    else: title = "💰 Salesmen Salary"; emoji = "🧑‍💼"; names = db.get("salesmen", []); sal_key = "salesmen_salaries"
    st.markdown(f"<h1 style='color:#0f172a;'>{title}</h1>", unsafe_allow_html=True); st.markdown("---")
    if sal_key not in db: db[sal_key] = {}
    if not names: st.info("No members added yet."); return
    for p_name in names:
        sd = db[sal_key].get(p_name, {"base_salary": 0, "transactions": []}); base = sd.get("base_salary", 0); txns = sd.get("transactions", [])
        adv_p = sum(t["amount"] for t in txns if t.get("type")=="advanced" and t.get("status","pending")=="pending"); short_p = sum(t["amount"] for t in txns if t.get("type")=="shortage" and t.get("status","pending")=="pending"); remaining = base - adv_p - short_p
        with st.expander(f"💰 {emoji} {p_name} — Remaining: Rs {remaining:,.0f}"):
            c1, c2 = st.columns([3, 1])
            with c1: nb = st.number_input("Base Salary", value=float(base), min_value=0.0, step=500.0, key=f"base_{role_type}_{p_name}")
            with c2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Save", key=f"savebase_{role_type}_{p_name}"): db[sal_key].setdefault(p_name, {"base_salary": 0, "transactions": []}); db[sal_key][p_name]["base_salary"] = float(nb); save_database(db); st.session_state["success_msg"] = "✅ Saved"; st.rerun()
            st.markdown(f"""<span class='sal-metric base'>Base: Rs {base:,.0f}</span><span class='sal-metric adv'>Adv Pend: Rs {adv_p:,.0f}</span><span class='sal-metric short'>Short Pend: Rs {short_p:,.0f}</span><span class='sal-metric remain'>Remaining: Rs {remaining:,.0f}</span>""", unsafe_allow_html=True)
            st.markdown("**Add Transaction**")
            c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
            with c1: td = st.date_input("Date", value=date.today(), key=f"td_{role_type}_{p_name}")
            with c2: tt = st.selectbox("Type", ["Advanced", "Shortage"], key=f"tt_{role_type}_{p_name}")
            with c3: ta = st.number_input("Amount", min_value=0.0, step=100.0, key=f"ta_{role_type}_{p_name}")
            with c4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕", key=f"addtxn_{role_type}_{p_name}"):
                    if ta <= 0: st.session_state["error_msg"] = "❌ Amount must be >0"
                    else:
                        db[sal_key].setdefault(p_name, {"base_salary": base, "transactions": []}); nid = max([t.get("id", 0) for t in txns] + [0]) + 1
                        db[sal_key][p_name]["transactions"].append({"id": nid, "date": td.strftime("%d-%m-%Y"), "time": datetime.now().strftime("%H:%M"), "type": "advanced" if tt == "Advanced" else "shortage", "amount": float(ta), "note": "", "status": "pending", "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
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
                    if st.button("🗑", key=f"dt_{role_type}_{p_name}_{t['id']}"): db[sal_key][p_name]["transactions"] = [x for x in db[sal_key][p_name]["transactions"] if x.get("id") != t["id"]]; save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()

def render_admin_expense():
    st.markdown(f"<h1 style='color:#0f172a;'>💵 Daily Expense</h1>", unsafe_allow_html=True); st.markdown("---")
    if "petrol_expenses" not in db: db["petrol_expenses"] = []
    if "lunch_expenses" not in db: db["lunch_expenses"] = []
    if "misc_expenses" not in db: db["misc_expenses"] = []
    today = date.today(); c1, c2, c3 = st.columns([2, 2, 2])
    with c1: filter_mode = st.selectbox("Filter", ["📅 Aaj (Today)", "📆 This Month", "🗓️ Last 30 Days", "📋 All"], key="exp_filter_mode")
    with c2: from_date = st.date_input("From", value=today - timedelta(days=30), key="exp_from_date")
    with c3: to_date = st.date_input("To", value=today, key="exp_to_date")
    def date_match(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if filter_mode == "📅 Aaj (Today)": return d == today
        elif filter_mode == "📆 This Month": return d.year == today.year and d.month == today.month
        elif filter_mode == "🗓️ Last 30 Days": return today - timedelta(days=30) <= d <= today
        else: return True
    petrol_filtered = [x for x in db["petrol_expenses"] if date_match(x.get("date", ""))]; lunch_filtered = [x for x in db["lunch_expenses"] if date_match(x.get("date", ""))]; misc_filtered = [x for x in db["misc_expenses"] if date_match(x.get("date", ""))]
    total_petrol = sum(float(x.get("amount", 0)) for x in petrol_filtered); total_lunch = sum(float(x.get("amount", 0)) for x in lunch_filtered); total_misc = sum(float(x.get("amount", 0)) for x in misc_filtered)
    st.markdown(f"<div class='summary-box'><b>Petrol:</b> Rs {total_petrol:,.0f} | <b>Lunch:</b> Rs {total_lunch:,.0f} | <b>Misc:</b> Rs {total_misc:,.0f} | <b>GRAND TOTAL:</b> Rs {total_petrol+total_lunch+total_misc:,.0f}</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**⛽ Petrol**")
        c1a, c1b, c1c = st.columns([2, 2, 1])
        with c1a: ps = st.selectbox("Salesman", ["-- Select --"] + db.get("salesmen", []), key="ps")
        with c1b: pa = st.number_input("Amount", min_value=0.0, step=50.0, key="pa")
        with c1c:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕", key="add_petrol"):
                if ps == "-- Select --" or pa <= 0: st.session_state["error_msg"] = "❌ Invalid entry"
                else: nid = max([x.get("id", 0) for x in db["petrol_expenses"]] + [0]) + 1; db["petrol_expenses"].append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "time": datetime.now().strftime("%H:%M"), "salesman": ps, "amount": float(pa)}); save_database(db); st.session_state["success_msg"] = "✅"; st.rerun()
        if petrol_filtered:
            st.caption("Recent entries:")
            for x in sorted(petrol_filtered, key=lambda y: y.get("created_at",""), reverse=True)[:8]:
                cA, cB = st.columns([5, 1])
                with cA: st.markdown(f"📅 {x.get('date','')} · ⛽ **{x.get('salesman','')}** · Rs {float(x.get('amount',0)):,.0f}")
                with cB:
                    if st.button("🗑", key=f"del_pet_{x.get('id')}"): db["petrol_expenses"] = [y for y in db["petrol_expenses"] if y.get("id") != x.get("id")]; save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()
    with c2:
        st.markdown("**🍽️ Lunch**")
        c2a, c2b = st.columns([3, 1])
        with c2a: la = st.number_input("Amount", min_value=0.0, step=50.0, key="la")
        with c2b:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕", key="add_lunch"):
                if la <= 0: st.session_state["error_msg"] = "❌ Amount must be >0"
                else: nid = max([x.get("id", 0) for x in db["lunch_expenses"]] + [0]) + 1; db["lunch_expenses"].append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "time": datetime.now().strftime("%H:%M"), "amount": float(la)}); save_database(db); st.session_state["success_msg"] = "✅"; st.rerun()
        if lunch_filtered:
            st.caption("Recent entries:")
            for x in sorted(lunch_filtered, key=lambda y: y.get("created_at",""), reverse=True)[:8]:
                cA, cB = st.columns([5, 1])
                with cA: st.markdown(f"📅 {x.get('date','')} · 🍽️ Rs {float(x.get('amount',0)):,.0f}")
                with cB:
                    if st.button("🗑", key=f"del_lun_{x.get('id')}"): db["lunch_expenses"] = [y for y in db["lunch_expenses"] if y.get("id") != x.get("id")]; save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()
    st.markdown("---"); st.markdown("### 🧾 Miscellaneous Expense")
    mc1, mc2, mc3, mc4, mc5 = st.columns([2, 2, 1.5, 2, 1])
    with mc1: misc_desc = st.text_input("Item / Description", key="misc_desc", placeholder="pani / baraf / etc.")
    with mc2: misc_note = st.text_input("Note (optional)", key="misc_note", placeholder="extra detail")
    with mc3: misc_amt = st.number_input("Amount", min_value=0.0, step=50.0, key="misc_amt")
    with mc4: misc_date = st.date_input("Date", value=today, key="misc_date")
    with mc5:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕ Add", key="add_misc", type="primary", use_container_width=True):
            if not misc_desc.strip() or misc_amt <= 0: st.session_state["error_msg"] = "❌ Description and Amount required"
            else:
                nid = max([x.get("id", 0) for x in db.get("misc_expenses", [])] + [0]) + 1
                db.setdefault("misc_expenses", []).append({"id": nid, "date": misc_date.strftime("%d-%m-%Y"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "time": datetime.now().strftime("%H:%M"), "description": misc_desc.strip(), "note": misc_note.strip(), "amount": float(misc_amt)})
                save_database(db); st.session_state["success_msg"] = f"✅ Added Rs {misc_amt:,.0f}"; st.rerun()
    if misc_filtered:
        st.markdown("**📋 Misc Entries (filtered):**")
        for x in sorted(misc_filtered, key=lambda y: y.get("created_at",""), reverse=True):
            cA, cB = st.columns([6, 1])
            with cA: st.markdown(f"📅 {x.get('date','')} · 🧾 **{x.get('description','')}** · {x.get('note','') or '—'} · **Rs {float(x.get('amount',0)):,.0f}**")
            with cB:
                if st.button("🗑", key=f"del_misc_{x.get('id')}"): db["misc_expenses"] = [y for y in db["misc_expenses"] if y.get("id") != x.get("id")]; save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()

def render_admin_billing():
    st.markdown(f"<h1 style='color:#0f172a;'>🧾 Billing</h1>", unsafe_allow_html=True); st.markdown("---")
    c1, c2 = st.columns(2)
    with c1: st.text_input("Bill No", value=str(db["next_bill_no"]), disabled=True)
    with c2: st.text_input("Date", value=datetime.now().strftime("%d-%m-%Y"), disabled=True)
    c1, c2, c3 = st.columns(3)
    with c1: shop = st.text_input("Shop Name", key="shop_name", placeholder="Enter shop name")
    with c2:
        bks = db.get("bookers", [])
        if bks: sel = st.selectbox("Order Booker", ["-- Select --"] + bks, key="order_booker_select"); st.session_state["order_booker"] = "" if sel == "-- Select --" else sel
        else: st.text_input("Order Booker", key="order_booker")
    with c3:
        sms = db.get("salesmen", [])
        if sms: sel = st.selectbox("Salesman", ["-- Select --"] + sms, key="salesman_select"); st.session_state["salesman"] = "" if sel == "-- Select --" else sel
        else: st.text_input("Salesman", key="salesman")
    c1, c2 = st.columns([1, 3])
    with c1: s = st.text_input("🔍 Search", key="search_text", placeholder="Product name...")
    with c2:
        all_p = get_all_products(); su = s.strip().upper(); f = [p["name"] for p in all_p if su in p["name"].upper()] if su else [p["name"] for p in all_p]
        if st.session_state.get("product_sel") and st.session_state["product_sel"] not in f: st.session_state["product_sel"] = ""
        ps = st.selectbox("Product", [""] + f, key="product_sel")
    sel_p = next((p for p in get_all_products() if p["name"] == ps), None); tp_def = get_price(sel_p["code"], sel_p["price"]) if sel_p else 0.0
    if st.session_state["_prev_prod"] != ps: st.session_state["tp_box"] = float(tp_def); st.session_state["_prev_prod"] = ps
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: bx = st.number_input("Boxes", min_value=0, step=1, key="boxes")
    with c2: tp = st.number_input("TP/Box", min_value=0.0, step=1.0, key="tp_box")
    with c3: dc = st.number_input("Discount %", min_value=0.0, step=0.5, key="discount")
    g = bx * tp; n = g - (g * dc / 100)
    with c4: st.text_input("Gross", value=f"{g:.0f}", disabled=True)
    with c5: st.text_input("Net", value=f"{n:.0f}", disabled=True)
    b1, b2 = st.columns(2)
    with b1: st.button("➕ Add Bill", key="btn_add", on_click=add_bill_callback, type="primary", use_container_width=True)
    with b2: st.button("🔄 Refresh", key="btn_refresh", on_click=refresh_callback, use_container_width=True)
    e1, e2, e3 = st.columns(3)
    with e1: st.button("📄 Export Bill", key="btn_export", on_click=export_bill_callback, use_container_width=True)
    with e2: st.button("📦 Export Load Form", key="btn_export_lf", on_click=export_load_form_from_billing_callback, use_container_width=True)
    with e3: st.button("🔄 Refresh Load Form", key="btn_load_refresh", on_click=refresh_load_form_callback, use_container_width=True)
    show_auto_download()

def render_admin_bills_list():
    st.markdown(f"<h1 style='color:#0f172a;'>📋 Bills List</h1>", unsafe_allow_html=True); st.markdown("---")
    if not db["bills"]: st.info("No bills created yet."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Filter Mode", ["📅 Aaj (Today)", "📆 Custom Range", "🗓️ Specific", "📋 All Bills"], key="bills_filter_mode")
    with c2: from_date = st.date_input("From", value=today - timedelta(days=7), key="bills_from_date")
    with c3: to_date = st.date_input("To", value=today, key="bills_to_date")
    credit_map = {(cb.get("shop",""), cb.get("date",""), cb.get("booker",""), cb.get("bill_no","")): cb for cb in db.get("credit_bills", [])}
    groups = {}
    for bi, b in enumerate(db["bills"]):
        bdate = parse_date(b.get("Date", ""))
        if filter_mode == "📅 Aaj (Today)" and bdate != today: continue
        elif filter_mode == "🗓️ Specific" and bdate != from_date: continue
        elif filter_mode == "📆 Custom Range" and (bdate is None or not (from_date <= bdate <= to_date)): continue
        key = (b.get("Shop",""), b.get("Date",""), b.get("Order Booker",""), b.get("Bill No",""))
        if key not in groups: groups[key] = {"shop": b.get("Shop",""), "date": b.get("Date",""), "booker": b.get("Order Booker",""), "salesman": b.get("Salesman",""), "bill_no": b.get("Bill No",""), "items": [], "idx": []}
        groups[key]["items"].append({"Code": b.get("Code"), "Product": b.get("Product"), "Boxes": b.get("Boxes"), "TP/Box": b.get("TP/Box"), "Discount %": b.get("Discount %"), "Gross": b.get("Gross"), "Net": b.get("Net")}); groups[key]["idx"].append(bi)
    st.markdown(f"### 📋 Bills ({len(groups)})")
    for key, g in sorted(groups.items(), key=lambda kv: kv[0][1], reverse=True):
        shop = g["shop"] or "-"; dt = g["date"] or "-"; bk = g["booker"] or "-"; sm = g["salesman"] or "-"; bn = g["bill_no"]; tb = sum(int(it.get("Boxes",0)) for it in g["items"])
        ck = (g["shop"] or "", g["date"] or "", g["booker"] or "", g["bill_no"] or ""); ec = credit_map.get(ck); cbadge = ""
        if ec: cbadge = '<span class="badge-paid-credit">✅ Credit PAID</span>' if ec.get("status")=="paid" else '<span class="badge-credit-tag">💳 Credit PEND</span>'
        st.markdown(f"""<div class='lf-simple-card'><div class='lf-info'><div class='lf-line1'>🏪 {shop} {cbadge}</div><div class='lf-line2'>📅 {dt} · 👤 {bk} · 🧑‍💼 {sm} · 🧾 #{bn}</div></div><div class='lf-boxes'>{tb}<small>BOXES</small></div></div>""", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            if st.button(f"⬇️ Excel", key=f"dl_{key}_{dt}"): export_single_group_bill(shop, dt, bk, sm, g["items"], bn); st.rerun()
        with c2:
            if ec: st.button("💳 Already", key=f"cr_{key}_{dt}", disabled=True)
            else:
                if st.button("💳 Mark Credit", key=f"cr_{key}_{dt}"): move_group_to_credit(g); st.rerun()
        with c3:
            if st.button("🗑 Delete", key=f"del_{key}_{dt}"): idxs = set(g["idx"]); db["bills"] = [b for i, b in enumerate(db["bills"]) if i not in idxs]; save_database(db); st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()
        st.markdown("---")
    show_auto_download()

def render_admin_credit():
    st.markdown(f"<h1 style='color:#0f172a;'>💳 Credit Bills</h1>", unsafe_allow_html=True); st.markdown("---")
    cbs = db.get("credit_bills", [])
    if not cbs: st.info("No credit bills found."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Filter", ["📋 All Credit Bills", "📅 Aaj", "📆 Custom Range", "🗓️ Specific"], key="credit_filter_mode")
    with c2: from_date = st.date_input("From", value=today - timedelta(days=30), key="credit_from_date")
    with c3: to_date = st.date_input("To", value=today, key="credit_to_date")
    filtered_cbs = []
    for c in cbs:
        cdate = parse_date(c.get("date", ""))
        if filter_mode == "📅 Aaj" and cdate != today: continue
        elif filter_mode == "🗓️ Specific" and cdate != from_date: continue
        elif filter_mode == "📆 Custom Range" and (cdate is None or not (from_date <= cdate <= to_date)): continue
        filtered_cbs.append(c)
    pend = [c for c in filtered_cbs if c.get("status") == "pending"]; paid = [c for c in filtered_cbs if c.get("status") == "paid"]
    st.markdown(f"<div class='summary-box'><b>Pending:</b> {len(pend)} (Rs {sum(float(c.get('total_net',0)) for c in pend):,.0f}) | <b>Paid:</b> {len(paid)} (Rs {sum(float(c.get('total_net',0)) for c in paid):,.0f})</div>", unsafe_allow_html=True)
    for c in pend + paid:
        cid = c["id"]; is_paid = c.get("status") == "paid"; badge = "✅ PAID" if is_paid else "⏳ PENDING"
        st.markdown(f"""<div class='lf-simple-card'><div class='lf-info'><div class='lf-line1'>🏪 {c.get('shop','')} · {badge}</div><div class='lf-line2'>📅 {c.get('date','')} · 👤 {c.get('booker','')} · 💰 Rs {float(c.get('total_net',0)):,.0f}</div></div><div class='lf-boxes'>{c.get('total_boxes',0)}<small>BOXES</small></div></div>""", unsafe_allow_html=True)
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
            if st.button("🗑 Delete", key=f"del_credit_{cid}"): db["credit_bills"] = [x for x in db["credit_bills"] if x.get("id") != cid]; save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()

def render_admin_load_form():
    st.markdown(f"<h1 style='color:#0f172a;'>📦 Load Forms</h1>", unsafe_allow_html=True); st.caption("Booker-wise load forms · Salesman assign · Transfer to DSR"); st.markdown("---")
    lfs = db.get("load_forms", [])
    if not lfs: st.info("No load forms found."); return
    today = date.today(); allow_old = st.checkbox("☑️ Show old dates (Transfer allowed)", value=False, key="lf_allow_old_dates"); st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Filter Mode", ["📅 Aaj", "📆 Custom Range", "🗓️ Specific", "📋 All"], key="lf_filter_mode")
    with c2: from_date = st.date_input("From", value=today, key="lf_from_date")
    with c3: to_date = st.date_input("To", value=today, key="lf_to_date")
    def in_range(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if allow_old: return True
        if filter_mode == "📅 Aaj": return d == today
        elif filter_mode == "🗓️ Specific": return d == from_date
        elif filter_mode == "📆 Custom Range": return from_date <= d <= to_date
        else: return True
    bookers_in_lfs = sorted({lf.get("booker", "") for lf in lfs if lf.get("booker")})
    if not bookers_in_lfs: st.info("No load forms found."); return
    grand_lf = 0; grand_boxes = 0; grand_ret = 0; grand_net = 0
    for bk in bookers_in_lfs:
        bk_lfs = [lf for lf in lfs if lf.get("booker") == bk and in_range(lf.get("date", ""))]
        total_lf_boxes = sum(int(lf.get("total_boxes", 0)) for lf in bk_lfs); lf_ids = {lf.get("id") for lf in bk_lfs}; total_returns = 0
        for r in db.get("booker_returns", []):
            if r.get("load_form_id") not in lf_ids: continue
            items = r.get("items", [])
            if items: total_returns += sum(int(it.get("ReturnBoxes", 0)) for it in items)
            else: total_returns += int(r.get("boxes", 0))
        net_boxes = total_lf_boxes - total_returns; grand_lf += len(bk_lfs); grand_boxes += total_lf_boxes; grand_ret += total_returns; grand_net += net_boxes
        expander_label = f"👤 {bk} — Load Forms: {len(bk_lfs)} | Boxes: {total_lf_boxes} | Returns: {total_returns} | 🟢 Net: {net_boxes}"
        with st.expander(expander_label, expanded=False):
            st.markdown(f"""<div class='bb-summary'><div class='bb-stat'><div class='bb-stat-label'>Load Forms</div><div class='bb-stat-val'>{len(bk_lfs)}</div></div><div class='bb-stat'><div class='bb-stat-label'>Total Boxes</div><div class='bb-stat-val'>{total_lf_boxes}</div></div><div class='bb-stat'><div class='bb-stat-label'>Returns</div><div class='bb-stat-val red'>-{total_returns}</div></div><div class='bb-stat'><div class='bb-stat-label'>Net Boxes</div><div class='bb-stat-val green'>{net_boxes}</div></div></div>""", unsafe_allow_html=True)
            if not bk_lfs: st.info("No load forms in this filter.")
            for lf in sorted(bk_lfs, key=lambda x: x.get("created_at", ""), reverse=True):
                lf_id = lf.get("id"); dt = lf.get("date", ""); tb = lf.get("total_boxes", 0); items = lf.get("items", []); trf = lf.get("transferred_to_dsr", False); lf_ret_map = get_lf_return_items(lf_id); lf_ret_total = sum(lf_ret_map.values()); net_lf = tb - lf_ret_total; badge = f'✅ DSR #{lf.get("dsr_id")}' if trf else '🟢 Active'
                st.markdown(f"""<div class='lf-simple-card'><div class='lf-info'><div class='lf-line1'>📦 Load Form #{lf_id} <span class='badge-pending'>{badge}</span></div><div class='lf-line2'>📅 {dt} · 🕐 {lf.get('time','')} · 📦 {len(items)} products · ↩️ Returns: {lf_ret_total} · 🟢 Net: {net_lf}</div></div><div class='lf-boxes'>{tb}<small>BOXES</small></div></div>""", unsafe_allow_html=True)
                salesmen_list = db.get("salesmen", [])
                if salesmen_list:
                    current_salesman = lf.get("salesman", ""); default_idx = 0
                    if current_salesman in salesmen_list: default_idx = salesmen_list.index(current_salesman) + 1
                    sel_salesman = st.selectbox("👨‍💼 Assign Salesman", ["-- Select --"] + salesmen_list, index=default_idx, key=f"sm_select_{lf_id}")
                    if sel_salesman != "-- Select --" and sel_salesman != current_salesman: lf["salesman"] = sel_salesman; save_database(db); st.session_state["success_msg"] = f"✅ Salesman {sel_salesman} assigned"; st.rerun()
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    if st.button("↩️ Add Return", key=f"ret_open_{lf_id}", use_container_width=True): st.session_state[f"ret_show_{lf_id}"] = not st.session_state.get(f"ret_show_{lf_id}", False); st.rerun()
                with c2:
                    if not trf:
                        if st.button("📤 Transfer to DSR", key=f"trf_{lf_id}", type="primary", use_container_width=True): transfer_load_form_to_dsr(lf_id); st.rerun()
                    else: st.button(f"✅ DSR #{lf.get('dsr_id','')}", key=f"trfd_{lf_id}", disabled=True, use_container_width=True)
                with c3:
                    if st.button("📥 Download Excel", key=f"dl_lf_{lf_id}", use_container_width=True, type="primary"): export_lf_excel_callback(lf_id); st.rerun()
                with c4:
                    if st.button("🗑 Delete", key=f"del_lf_{lf_id}", use_container_width=True): db["load_forms"] = [x for x in db["load_forms"] if x.get("id") != lf_id]; db["booker_returns"] = [r for r in db.get("booker_returns", []) if r.get("load_form_id") != lf_id]; save_database(db); st.session_state["success_msg"] = "🗑 Deleted"; st.rerun()
                if st.session_state.get(f"ret_show_{lf_id}", False):
                    st.markdown(f"<div style='background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:16px;margin:8px 0;'>", unsafe_allow_html=True); st.markdown(f"#### ↩️ Return Items — 👤 {bk} (Load Form #{lf_id})")
                    existing_ret = get_return_record_for_lf(lf_id); existing_map = {}
                    if existing_ret:
                        for it in existing_ret.get("items", []): existing_map[str(it.get("Code", ""))] = int(it.get("ReturnBoxes", 0))
                    hc1, hc2, hc3, hc4 = st.columns([1, 4, 2, 2])
                    with hc1: st.markdown("**Code**")
                    with hc2: st.markdown("**Product**")
                    with hc3: st.markdown("**Loaded**")
                    with hc4: st.markdown("**Return**")
                    return_inputs = {}
                    for idx, it in enumerate(items):
                        code = str(it.get("Code", "")); prod = it.get("Product", ""); loaded = int(it.get("Boxes", 0))
                        c1, c2, c3, c4 = st.columns([1, 4, 2, 2])
                        with c1: st.markdown(f"<div style='padding-top:8px;font-weight:600;color:#2563eb;'>{code}</div>", unsafe_allow_html=True)
                        with c2: st.markdown(f"<div style='padding-top:8px;font-weight:500;'>{prod}</div>", unsafe_allow_html=True)
                        with c3: st.markdown(f"<div style='padding-top:8px;color:#16a34a;font-weight:600;'>📦 {loaded}</div>", unsafe_allow_html=True)
                        with c4: return_inputs[code] = st.number_input("ret", min_value=0, max_value=loaded, value=int(existing_map.get(code, 0)), step=1, key=f"ret_qty_{lf_id}_{code}_{idx}", label_visibility="collapsed")
                    rc1, rc2, rc3 = st.columns([2, 3, 1])
                    with rc1:
                        default_date = today
                        if existing_ret:
                            try: default_date = datetime.strptime(existing_ret.get("date", ""), "%d-%m-%Y").date()
                            except Exception: default_date = today
                        r_date = st.date_input("Return Date", value=default_date, key=f"rd_{lf_id}")
                    with rc2: r_note = st.text_input("Note (optional)", value=existing_ret.get("note", "") if existing_ret else "", key=f"rn_{lf_id}")
                    with rc3:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("💾 Save Return", key=f"radd_{lf_id}", type="primary", use_container_width=True):
                            new_items = []; total_new = 0
                            for it in items:
                                code = str(it.get("Code", "")); qty = int(return_inputs.get(code, 0))
                                if qty > 0: new_items.append({"Code": code, "Product": it.get("Product", ""), "ReturnBoxes": qty}); total_new += qty
                            if existing_ret:
                                existing_ret["items"] = new_items; existing_ret["date"] = r_date.strftime("%d-%m-%Y"); existing_ret["note"] = r_note.strip(); existing_ret["total_return_boxes"] = total_new; existing_ret["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); msg = f"✅ Return updated ({total_new} boxes)"
                            else:
                                nid = max([r.get("id", 0) for r in db.get("booker_returns", [])] + [0]) + 1
                                db.setdefault("booker_returns", []).append({"id": nid, "booker": bk, "load_form_id": lf_id, "date": r_date.strftime("%d-%m-%Y"), "items": new_items, "total_return_boxes": total_new, "note": r_note.strip(), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}); msg = f"✅ Return saved ({total_new} boxes)"
                            save_database(db); st.session_state["success_msg"] = msg; st.session_state[f"ret_show_{lf_id}"] = False; st.rerun()
                    if existing_ret:
                        if st.button("🗑 Delete this Return", key=f"del_ret_{lf_id}"): db["booker_returns"] = [x for x in db["booker_returns"] if x.get("id") != existing_ret.get("id")]; save_database(db); st.session_state["success_msg"] = "🗑 Return deleted"; st.session_state[f"ret_show_{lf_id}"] = False; st.rerun()
                    st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("---")
    st.markdown(f"""<div class='bb-summary'><div class='bb-stat'><div class='bb-stat-label'>Load Forms</div><div class='bb-stat-val'>{grand_lf}</div></div><div class='bb-stat'><div class='bb-stat-label'>Total Boxes</div><div class='bb-stat-val'>{grand_boxes}</div></div><div class='bb-stat'><div class='bb-stat-label'>Returns</div><div class='bb-stat-val red'>-{grand_ret}</div></div><div class='bb-stat'><div class='bb-stat-label'>NET (All)</div><div class='bb-stat-val green'>{grand_net}</div></div></div>""", unsafe_allow_html=True)

def render_admin_dsr():
    st.markdown(f"<h1 style='color:#0f172a;'>📋 DSR</h1>", unsafe_allow_html=True); st.markdown("---")
    dsrs = db.get("dsr_forms", [])
    if not dsrs: st.info("No DSR forms found."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Filter Mode", ["📅 Aaj", "📆 Custom Range", "🗓️ Specific", "📋 All"], key="dsr_filter_mode")
    with c2: from_date = st.date_input("From", value=today - timedelta(days=7), key="dsr_from_date")
    with c3: to_date = st.date_input("To", value=today, key="dsr_to_date")
    filtered_dsrs = []
    for d in dsrs:
        d_date = parse_date(d.get("date", ""))
        if d_date is None: continue
        if filter_mode == "📅 Aaj" and d_date != today: continue
        elif filter_mode == "🗓️ Specific" and d_date != from_date: continue
        elif filter_mode == "📆 Custom Range" and not (from_date <= d_date <= to_date): continue
        filtered_dsrs.append(d)
    if not filtered_dsrs: st.info("No DSR found in this filter."); return
    for i, d in enumerate(sorted(filtered_dsrs, key=lambda x: x.get("created_at",""), reverse=True)):
        did = d.get("id", i); bk = d.get("booker", "?"); dt = d.get("date", ""); sm = d.get("salesman", ""); tb = d.get("total_boxes", 0); tot = float(d.get("total_amount", 0)); rtb = int(d.get("total_return_boxes", 0)); ct, _, _ = get_dsr_credit_info(d); fin = float(d.get("amount_to_collect", 0)) - ct; sm_txt = f" · 👨‍💼 {sm}" if sm else ""
        st.markdown(f"""<div class='lf-simple-card dsr'><div class='lf-info'><div class='lf-line1'>📋 DSR #{did} — {bk}</div><div class='lf-line2'>📅 {dt}{sm_txt} · 💰 Rs {tot:,.0f} · ↩️ {rtb} ret · 💳 Rs {ct:,.0f} · <b style="color:#dc2626;">FINAL: Rs {fin:,.0f}</b></div></div><div class='lf-boxes' style="background:#ffedd5;color:#ea580c;">{tb}<small>BOXES</small></div></div>""", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📤 Excel", key=f"ex_dsr_{did}", type="primary"): export_dsr_excel(d); st.rerun()
        with c2:
            if st.button("🗑 Delete", key=f"del_dsr_{did}"):
                db["dsr_forms"] = [x for x in db["dsr_forms"] if x.get("id") != did]
                for x in db.get("load_forms", []):
                    if x.get("dsr_id") == did: x["transferred_to_dsr"] = False; x.pop("dsr_id", None)
                save_database(db); st.session_state["success_msg"] = "🗑"; st.rerun()
        st.markdown("---")

def render_admin_calculation():
    st.markdown(f"<h1 style='color:#0f172a;'>🧮 Cash Calculation</h1>", unsafe_allow_html=True); st.markdown("---")
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
        if st.button("💰 Calculate", type="primary", use_container_width=True): st.session_state["calc_locked"] = total
    with c2:
        if st.button("🔄 Reset", use_container_width=True): st.session_state["calc_reset_token"] += 1; st.session_state["calc_locked"] = 0; st.rerun()
    st.markdown(f"<div style='background:#0f172a;border-radius:16px;padding:24px;text-align:center;color:white;margin-top:16px;'><div style='font-size:12px;letter-spacing:2px;font-weight:700;color:#94a3b8;'>LIVE TOTAL</div><div style='font-size:42px;font-weight:800;'>Rs {total:,.0f}</div></div>", unsafe_allow_html=True)
    locked = st.session_state.get("calc_locked", 0)
    if locked and locked > 0: st.markdown(f"<div style='background:#16a34a;border-radius:16px;padding:24px;text-align:center;color:white;margin-top:10px;'><div style='font-size:12px;letter-spacing:2px;font-weight:700;color:#dcfce7;'>🔒 LOCKED TOTAL</div><div style='font-size:42px;font-weight:800;'>Rs {locked:,.0f}</div></div>", unsafe_allow_html=True)

# ============================================================
# CALLBACKS
# ============================================================
def add_bill_callback():
    ps = st.session_state.get("product_sel", ""); bx = st.session_state.get("boxes", 0); tp = st.session_state.get("tp_box", 0.0); dc = st.session_state.get("discount", 0.0)
    if not ps: st.session_state["error_msg"] = "❌ Select Product"; return
    if bx <= 0: st.session_state["error_msg"] = "❌ Enter Boxes"; return
    sel = next((p for p in get_all_products() if p["name"] == ps), None)
    if not sel: st.session_state["error_msg"] = "❌ Invalid Product"; return
    g = bx * tp; n = g - (g * dc / 100)
    db["bills"].append({"Bill No": db["next_bill_no"], "Date": datetime.now().strftime("%d-%m-%Y"), "Shop": st.session_state.get("shop_name","").strip(), "Order Booker": st.session_state.get("order_booker","").strip(), "Salesman": st.session_state.get("salesman","").strip(), "Delivery Man": "", "Code": sel["code"], "Product": sel["name"], "Boxes": bx, "TP/Box": tp, "Discount %": dc, "Gross": g, "Net": n})
    save_database(db); st.session_state["last_bill_no"] = db["next_bill_no"]; st.session_state["success_msg"] = f"✅ Bill #{db['next_bill_no']}"
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None

def refresh_callback():
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None; st.session_state["success_msg"] = "✅ Cleared"

def _build_bill_excel(shop, bills, pkg_pct, pkg_name, tier_label, booker="", date_str=""):
    if not booker:
        try: booker = bills[0].get("Order Booker", "") if bills else ""
        except Exception: booker = ""
    if not date_str:
        try: date_str = bills[0].get("Date", "") if bills else ""
        except Exception: date_str = ""
    output = BytesIO(); wb = xlsxwriter.Workbook(output, {'in_memory': True}); ws = wb.add_worksheet("Bill")
    headers_txt = ["PRODUCT", "CODE", "BOXES", "TP/BOX", "GROSS", "DISC %", "NET", "PKG DISC %", "AFTER DISC NET", "SAVED"]
    max_lens = [len(h) for h in headers_txt]
    for b in bills:
        row_data = [str(b.get("Product", "")).upper(), str(b.get("Code", "")).upper(), str(b.get("Boxes", 0)), f"{float(b.get('TP/Box', 0)):,.0f}", f"{float(b.get('Gross', 0)):,.0f}", f"{float(b.get('Discount %', 0))}", f"{float(b.get('Net', 0)):,.0f}", f"{float(pkg_pct)}", f"{float(b.get('Net', 0)) * (1 - pkg_pct/100):,.0f}", f"{float(b.get('Net', 0)) * (pkg_pct/100):,.0f}"]
        for i, val in enumerate(row_data):
            if i < len(max_lens): max_lens[i] = max(max_lens[i], len(val))
    total_len_check = ["TOTAL", "", "999999", "", "99999999", "", "99999999", "", "99999999", "99999999"]
    for i, val in enumerate(total_len_check):
        if i < len(max_lens): max_lens[i] = max(max_lens[i], len(val))
    for i, ml in enumerate(max_lens): width = max(10, ml + 4); width = max(28, ml + 4) if i == 0 else width; ws.set_column(i, i, width)
    title = wb.add_format({"bold": True, "font_size": 18, "align": "center", "border": 2, "bg_color": "#e2e8f0"})
    info_lbl = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f8fafc", "align": "right", "border": 1, "valign": "vcenter"})
    info_val = wb.add_format({"bold": True, "font_size": 11, "align": "left", "border": 1, "valign": "vcenter", "font_color": "#2563eb"})
    info_date = wb.add_format({"bold": True, "font_size": 11, "align": "left", "border": 1, "valign": "vcenter", "font_color": "#16a34a"})
    hdr = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f1f5f9", "align": "center", "border": 1, "text_wrap": True, "valign": "vcenter"})
    cl = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "left", "text_wrap": True, "valign": "vcenter"})
    cc = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "center", "valign": "vcenter"})
    cn = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "right", "valign": "vcenter", "num_format": "#,##0"})
    tt = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#fef3c7", "align": "center", "border": 2})
    ttn = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#fef3c7", "align": "right", "border": 2, "num_format": "#,##0"})
    dl = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "center", "bg_color": "#dcfce7", "valign": "vcenter"})
    pi = wb.add_format({"bold": True, "font_size": 10, "italic": True, "align": "left", "font_color": "#16a34a", "border": 1})
    net_lbl = wb.add_format({"bold": True, "font_size": 13, "bg_color": "#0f172a", "font_color": "#ffffff", "align": "center", "border": 2, "valign": "vcenter"})
    net_val = wb.add_format({"bold": True, "font_size": 16, "bg_color": "#dcfce7", "align": "center", "border": 2, "num_format": '"RS "#,##0', "valign": "vcenter"})
    ws.merge_range("A1:J1", COMPANY_NAME, title)
    ws.merge_range("A3:B3", "ORDER BOOKER:", info_lbl); ws.merge_range("C3:E3", f"{booker or '-'}".upper(), info_val); ws.merge_range("F3:G3", "DATE:", info_lbl); ws.merge_range("H3:J3", f"{date_str or '-'}".upper(), info_date)
    ws.merge_range("A4:B4", "SHOP:", info_lbl); ws.merge_range("C4:E4", (shop or "-").upper(), info_val); ws.merge_range("F4:G4", "BILL NO:", info_lbl); ws.merge_range("H4:J4", f"#{str(bills[0].get('Bill No', '')).upper() if bills else ''}", info_val)
    if pkg_pct > 0: ws.merge_range("A5:J5", f"DISCOUNT: {pkg_name} | {tier_label} | {pkg_pct}%".upper(), pi)
    else: ws.merge_range("A5:J5", "NO DISCOUNT APPLIED", pi)
    hr = 6
    for col, h in enumerate(headers_txt): ws.write(hr, col, h, hdr)
    row = hr + 1; gt = 0; tb = 0; nt = 0; ad = 0; sv = 0
    for b in bills:
        bn = float(b.get("Net", 0)); bg = float(b.get("Gross", 0)); bb = int(b.get("Boxes", 0)); an = bn - (bn * pkg_pct / 100); s = bn - an
        ws.write(row, 0, str(b["Product"]).upper(), cl); ws.write(row, 1, str(b["Code"]).upper(), cc); ws.write(row, 2, bb, cc); ws.write(row, 3, float(b.get("TP/Box", 0)), cn); ws.write(row, 4, bg, cn); ws.write(row, 5, float(b.get("Discount %", 0)), cc); ws.write(row, 6, bn, cn)
        if pkg_pct > 0: ws.write(row, 7, pkg_pct, dl); ws.write(row, 8, an, dl); ws.write(row, 9, s, dl)
        else: ws.write(row, 7, 0, cc); ws.write(row, 8, bn, cn); ws.write(row, 9, 0, cn)
        gt += bg; tb += bb; nt += bn; ad += an; sv += s; row += 1
    ws.merge_range(row, 0, row, 1, "TOTAL", tt); ws.write(row, 2, tb, tt); ws.write(row, 3, "", tt); ws.write(row, 4, gt, ttn); ws.write(row, 5, "", tt); ws.write(row, 6, nt, ttn); ws.write(row, 7, "", tt); ws.write(row, 8, ad, ttn); ws.write(row, 9, sv, ttn); row += 2
    ws.merge_range(row, 0, row, 6, "NET AMOUNT", net_lbl); ws.merge_range(row, 7, row, 9, ad, net_val); row += 2
    foot_lbl = wb.add_format({"bold": True, "font_size": 9, "italic": True, "align": "right", "font_color": "#64748b"}); foot_val = wb.add_format({"bold": True, "font_size": 9, "italic": True, "align": "left", "font_color": "#64748b"})
    ws.merge_range(row, 0, row, 1, "BOOKER:", foot_lbl); ws.merge_range(row, 2, row, 4, (booker or "-").upper(), foot_val); ws.merge_range(row, 5, row, 6, "DATE:", foot_lbl); ws.merge_range(row, 7, row, 9, (date_str or "-").upper(), foot_val)
    ws.set_paper(9); ws.set_portrait(); ws.fit_to_pages(1, 1); wb.close(); output.seek(0); return output.getvalue()

def export_single_group_bill(shop, date_str, booker, salesman, items, bill_no):
    tot = sum(float(it.get("Net", 0)) for it in items); pct, pn, tl = get_effective_discount_pct(shop, tot)
    safe_shop = shop.replace("/", "-").replace(" ", "_").replace(":", ""); safe_bk = (booker or "bill").replace(" ", "_").replace("/", "-"); safe_date = (date_str or "").replace("-", "")
    st.session_state["download_file"] = (f"{safe_shop}_{safe_bk}_{safe_date}.xlsx", _build_bill_excel(shop, items, pct, pn, tl, booker=booker, date_str=date_str))
    st.session_state["success_msg"] = f"✅ Excel ready — {booker} | {date_str}"

def export_bill_callback():
    if not db["bills"]: st.session_state["error_msg"] = "❌ No Bills"; return
    shop = st.session_state.get("shop_name", "").strip() or "Bill"
    for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']: shop = shop.replace(ch, "")
    sb = [b for b in db["bills"] if b["Shop"].strip() == shop]
    if not sb: st.session_state["error_msg"] = "❌ No bills for this shop"; return
    tot = sum(float(b.get("Net", 0)) for b in sb); pct, pn, tl = get_effective_discount_pct(shop, tot)
    booker_val = sb[0].get("Order Booker", "") if sb else ""; date_val = sb[0].get("Date", "") if sb else ""
    st.session_state["download_file"] = (f"{shop}_{(booker_val or 'bill').replace(' ', '_')}.xlsx", _build_bill_excel(shop, sb, pct, pn, tl, booker=booker_val, date_str=date_val))
    db["next_bill_no"] += 1; save_database(db); st.session_state["success_msg"] = f"✅ Exported — {booker_val}"

def export_load_form_from_billing_callback():
    bk = st.session_state.get("order_booker", "").strip()
    if not bk: st.session_state["error_msg"] = "❌ Select Booker"; return
    bbs = [b for b in db["bills"] if b["Order Booker"].strip() == bk]
    if not bbs: st.session_state["error_msg"] = "❌ No bills for this booker"; return
    sm = {}
    for b in bbs:
        c = b["Code"]
        if c not in sm: sm[c] = {"Code": c, "Product": b["Product"], "Boxes": 0}
        sm[c]["Boxes"] += b["Boxes"]
    items = list(sm.values()); tb = sum(it["Boxes"] for it in items)
    existing = None
    for lf in reversed(db.get("load_forms", [])):
        if (str(lf.get("booker","")).strip() == bk and not lf.get("transferred_to_dsr") and not lf.get("refreshed")): existing = lf; break
    if existing: existing["items"] = items; existing["total_boxes"] = tb; existing["total_products"] = len(items); existing["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); msg = f"🔄 Load Form #{existing['id']} updated"
    else:
        nid = max([lf.get("id",0) for lf in db.get("load_forms", [])] + [0]) + 1
        db.setdefault("load_forms", []).append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"), "time": datetime.now().strftime("%H:%M"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "booker": bk, "items": items, "total_boxes": tb, "total_products": len(items), "transferred_to_dsr": False, "refreshed": False}); msg = f"✅ New Load Form #{nid}"
    save_database(db); st.session_state["success_msg"] = msg

def refresh_load_form_callback():
    bk = st.session_state.get("order_booker", "").strip()
    if not bk: st.session_state["error_msg"] = "❌ Enter Booker"; return
    db["bills"] = [b for b in db["bills"] if b["Order Booker"].strip() != bk]
    for lf in db.get("load_forms", []):
        if (str(lf.get("booker","")).strip() == bk and not lf.get("transferred_to_dsr") and not lf.get("refreshed")): lf["refreshed"] = True; lf["refreshed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_database(db)
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None; st.session_state["success_msg"] = f"✅ Cleared"

def move_group_to_credit(group):
    shop = group.get("shop","") or "-"; dt = group.get("date","") or "-"; bk = group.get("booker","") or "-"; sm = group.get("salesman","") or "-"; bn = group.get("bill_no",""); items = group.get("items", [])
    for cb in db.get("credit_bills", []):
        if (cb.get("shop")==shop and cb.get("date")==dt and cb.get("booker")==bk and cb.get("bill_no")==bn): st.session_state["error_msg"] = "⚠️ Already in Credit"; return
    tb = sum(int(it.get("Boxes",0)) for it in items); tg = sum(float(it.get("Gross",0)) for it in items); tn = sum(float(it.get("Net",0)) for it in items)
    nid = db.get("next_credit_id", 1)
    db.setdefault("credit_bills", []).append({"id": nid, "shop": shop, "date": dt, "booker": bk, "salesman": sm, "bill_no": bn, "items": [dict(it) for it in items], "total_boxes": tb, "total_gross": tg, "total_net": tn, "status": "pending", "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "paid_at": None})
    db["next_credit_id"] = nid + 1; save_database(db); st.session_state["success_msg"] = f"💳 Credit #{nid}"

def transfer_load_form_to_dsr(lf_id):
    lf = next((x for x in db.get("load_forms", []) if x.get("id") == lf_id), None)
    if not lf or lf.get("transferred_to_dsr"): st.session_state["error_msg"] = "❌"; return
    bk = lf.get("booker",""); items = lf.get("items", []); allp = get_all_products(); sm = lf.get("salesman", ""); returns_map = get_lf_return_items(lf_id)
    dsr_items = []; tb = 0; ta = 0.0; ret_tb = 0; ret_ta = 0.0
    for it in items:
        c = str(it.get("Code","")); pr = next((p for p in allp if str(p["code"]) == c), None)
        try: price = get_price(c, pr["price"]) if pr else 0.0
        except Exception: price = 0.0
        bx = int(it.get("Boxes",0)); rbx = int(returns_map.get(c, 0)); rbx = bx if rbx > bx else rbx; lt = bx * float(price); rlt = rbx * float(price)
        dsr_items.append({"Code": c, "Product": it.get("Product",""), "Boxes": bx, "TP/Box": float(price), "Total": lt, "ReturnBoxes": rbx, "ReturnAmount": rlt}); tb += bx; ta += lt; ret_tb += rbx; ret_ta += rlt
    bbs = [b for b in db.get("bills", []) if b.get("Order Booker","").strip() == bk]; sgs = {}; src = []
    for b in bbs:
        s = b.get("Shop","-") or "-"; bn = b.get("Bill No",""); net = float(b.get("Net",0)); gr = float(b.get("Gross",0))
        if s not in sgs: sgs[s] = {"shop": s, "gross": 0.0, "net": 0.0, "individual_discount": 0.0, "bill_nos": set()}
        sgs[s]["gross"] += gr; sgs[s]["net"] += net; sgs[s]["individual_discount"] += (gr - net); sgs[s]["bill_nos"].add(bn); src.append({"shop": s, "bill_no": bn, "net": net})
    sds = []; td = 0.0
    for s, g in sgs.items():
        pct, pn, _ = get_effective_discount_pct(s, g["net"]); pd = g["net"] * pct / 100; std = g["individual_discount"] + pd
        sds.append({"shop": s, "gross": g["gross"], "net": g["net"], "individual_discount": g["individual_discount"], "package_pct": pct, "package_name": pn or "", "package_discount": pd, "total_discount": std, "bill_nos": sorted([str(x) for x in g["bill_nos"]])}); td += std
    net_amount = ta - ret_ta; nid = max([x.get("id",0) for x in db.get("dsr_forms", [])] + [0]) + 1
    db.setdefault("dsr_forms", []).append({"id": nid, "source_load_form_id": lf_id, "booker": bk, "salesman": sm, "date": lf.get("date",""), "time": lf.get("time",""), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "items": dsr_items, "total_boxes": tb, "total_amount": ta, "total_return_boxes": ret_tb, "total_return_amount": ret_ta, "net_amount": net_amount, "shop_discounts": sds, "total_discount": td, "amount_to_collect": net_amount - td, "source_bills": src, "status": "open", "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    for x in db["load_forms"]:
        if x.get("id") == lf_id: x["transferred_to_dsr"] = True; x["dsr_id"] = nid; break
    save_database(db); st.session_state["success_msg"] = f"✅ DSR #{nid} created (Returns: {ret_tb} boxes)"

def export_dsr_excel(dsr):
    ct, cs, _ = get_dsr_credit_info(dsr); fin = float(dsr.get("amount_to_collect", 0)) - ct
    output = BytesIO(); wb = xlsxwriter.Workbook(output, {'in_memory': True}); ws = wb.add_worksheet("DSR")
    title = wb.add_format({"bold":True, "font_size":16, "align":"center", "border":2, "bg_color":"#e2e8f0"}); header = wb.add_format({"bold":True, "font_size":11, "bg_color":"#f1f5f9", "align":"center", "border":1}); cl = wb.add_format({"font_size":11, "border":1, "align":"left"}); cc = wb.add_format({"font_size":11, "border":1, "align":"center"}); cn = wb.add_format({"font_size":11, "border":1, "align":"right", "num_format":"#,##0"}); tot = wb.add_format({"bold":True, "font_size":11, "bg_color":"#fef3c7", "align":"center", "border":1}); hi = wb.add_format({"bold":True, "font_size":12, "bg_color":"#dcfce7", "align":"center", "border":2})
    for col, w in [("A",10),("B",40),("C",10),("D",12),("E",14),("F",10),("G",14)]: ws.set_column(f"{col}:{col}", w)
    ws.merge_range("A1:G1", f"{COMPANY_NAME} — DSR #{dsr['id']}", title)
    ws.write("A3","Booker",header); ws.write("B3", dsr.get("booker",""), cl); ws.write("C3","Salesman",header); ws.write("D3", dsr.get("salesman","") or "-", cl); ws.write("E3","Date",header); ws.write("F3", dsr.get("date",""), cc); ws.write("G3","Time",header)
    row = 5; ws.write(row,0,"Code",header); ws.write(row,1,"Product",header); ws.write(row,2,"Boxes",header); ws.write(row,3,"TP/Box",header); ws.write(row,4,"Total",header); ws.write(row,5,"Ret Bx",header); ws.write(row,6,"Ret Amt",header); row += 1
    for it in dsr.get("items", []):
        ws.write(row,0,str(it.get("Code","")),cc); ws.write(row,1,it.get("Product",""),cl); ws.write(row,2,int(it.get("Boxes",0)),cc); ws.write(row,3,float(it.get("TP/Box",0)),cn); ws.write(row,4,float(it.get("Total",0)),cn); ws.write(row,5,int(it.get("ReturnBoxes",0)),cc); ws.write(row,6,float(it.get("ReturnAmount",0)),cn); row += 1
    ws.write(row,2,"TOTAL",tot); ws.write(row,4,float(dsr.get("total_amount",0)),tot); ws.write(row,5,int(dsr.get("total_return_boxes",0)),tot); ws.write(row,6,float(dsr.get("total_return_amount",0)),tot); row += 2
    ws.merge_range(row,0,row,4,"Stock Value",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('total_amount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(−) Returns",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('total_return_amount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(=) Net Stock",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('net_amount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(−) Discount",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('total_discount',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(=) Ye Lena Hai",header); ws.merge_range(row,5,row,6,f"Rs {float(dsr.get('amount_to_collect',0)):,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"(−) Credit Bills",header); ws.merge_range(row,5,row,6,f"Rs {ct:,.0f}",cn); row += 1
    ws.merge_range(row,0,row,4,"💰 FINAL YE LENA HAI",hi); ws.merge_range(row,5,row,6,f"Rs {fin:,.0f}",hi)
    wb.close(); output.seek(0); st.session_state["download_file"] = (f"DSR_{dsr['id']}_{dsr.get('booker','')}.xlsx".replace("/","-").replace(" ","_"), output.getvalue()); st.session_state["success_msg"] = f"✅ DSR #{dsr['id']} Excel ready"

# ============================================================
# RENDER ADMIN PAGE
# ============================================================
if st.session_state["page"] == "📊 Dashboard": render_admin_dashboard()
elif st.session_state["page"] == "🧾 Billing": render_admin_billing()
elif st.session_state["page"] == "🛒 All Products": render_admin_products()
elif st.session_state["page"] == "🎁 Discount": render_admin_discount()
elif st.session_state["page"] == "👤 Bookers": render_admin_bookers()
elif st.session_state["page"] == "📈 Booker Bills": render_admin_booker_bills()
elif st.session_state["page"] == "💰 Bookers Salary": render_admin_salaries("bookers")
elif st.session_state["page"] == "🧑‍💼 Salesmen": render_admin_salesmen()
elif st.session_state["page"] == "💰 Salesmen Salary": render_admin_salaries("salesmen")
elif st.session_state["page"] == "💵 Daily Expense": render_admin_expense()
elif st.session_state["page"] == "📋 Bills List": render_admin_bills_list()
elif st.session_state["page"] == "💳 Credit Bills": render_admin_credit()
elif st.session_state["page"] == "📦 Load Form": render_admin_load_form()
elif st.session_state["page"] == "📋 DSR": render_admin_dsr()
elif st.session_state["page"] == "🧮 Calculation": render_admin_calculation()

if st.session_state.get("success_msg"): st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
if st.session_state.get("error_msg"): st.error(st.session_state["error_msg"]); st.session_state["error_msg"] = None
show_auto_download()
