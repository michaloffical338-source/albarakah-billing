# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# + Lamp Login (Native Streamlit Form — No more stuck)
# ============================================================

import os
import json
import hashlib
import random
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
# QUERY PARAM HANDLING (Lamp on/off)
# ============================================================
if "light_on" not in st.session_state:
    st.session_state["light_on"] = False
if "auth_error" not in st.session_state:
    st.session_state["auth_error"] = ""

def _get_qp():
    try:
        return dict(st.query_params)
    except Exception:
        try:
            return st.experimental_get_query_params()
        except Exception:
            return {}

def _clear_qp():
    try:
        st.query_params.clear()
    except Exception:
        try:
            st.experimental_set_query_params()
        except Exception:
            pass

def _qp_val(qp, key):
    if key not in qp:
        return None
    v = qp[key]
    if isinstance(v, list):
        return v[0] if v else None
    return v

def _handle_query_params():
    qp = _get_qp()
    handled = False
    lamp_val = _qp_val(qp, "lamp")
    if lamp_val == "on":
        st.session_state["light_on"] = True
        handled = True
    elif lamp_val == "off":
        st.session_state["light_on"] = False
        handled = True
    if handled:
        _clear_qp()
        st.rerun()

_handle_query_params()

# ============================================================
# AUTH HELPERS
# ============================================================
USERS_FILE = "users.json"
PASSWORD_SALT = "albarakah_2026_secret_salt"

def hash_password(password):
    return hashlib.sha256((PASSWORD_SALT + password).encode()).hexdigest()

def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_users(users):
    with open(USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, indent=4, ensure_ascii=False)

def sanitize_username(u):
    s = "".join(ch for ch in u if ch.isalnum() or ch in "_-.")
    return s.lower()

def user_data_file(username):
    return f"billing_database_{username}.json"

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

def default_blank_db():
    return {
        "next_bill_no": 1, "bills": [], "bookers": [], "salesmen": [],
        "load_forms": [], "bookers_salaries": {}, "salesmen_salaries": {},
        "product_prices": {}, "petrol_expenses": [], "lunch_expenses": [],
        "discount_packages": default_discount_packages()
    }

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
    {"code":"90","name":"KIMS – SIR STRAWBERRY","price":328},
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

PRODUCT_NAMES = [p["name"] for p in PRODUCTS]
COMPANY_NAME = "AL-BARAKAH ENTERPRISES"

# ============================================================
# GLOBAL APP CSS (used when logged in)
# ============================================================
st.markdown("""
<style>
    html, body, .stApp, .stApp *, .stApp p, .stApp span, .stApp div,
    .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
    .stApp label, .stApp li, .stApp a, [class*="css"] *,
    [data-testid="stMarkdownContainer"] *,
    [data-testid="stText"], [data-testid="stCaptionContainer"] *,
    [data-testid="stWidgetLabel"] *, [data-testid="stSelectbox"] *,
    [data-testid="stTextInput"] *, [data-testid="stNumberInput"] * {
        color: #000000 !important;
    }

    .stDateInput, .stDateInput > div, .stDateInput > div > div,
    .stDateInput > div > div > input,
    [data-testid="stDateInput"], [data-testid="stDateInput"] > div,
    [data-testid="stDateInput"] > div > div,
    [data-testid="stDateInput"] input,
    [data-testid="stDateInput"] div[data-baseweb="input"],
    [data-testid="stDateInput"] div[data-baseweb="input"] > div,
    [data-testid="stDateInput"] div[data-baseweb="base-input"],
    [data-baseweb="datepicker"] input,
    div[data-baseweb="input"] > div,
    div[data-baseweb="base-input"],
    div[data-baseweb="base-input"] input {
        background-color: #ffffff !important;
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        border-color: #90caf9 !important;
    }
    div[data-baseweb="calendar"], div[data-baseweb="calendar"] *,
    div[data-baseweb="datepicker"] *, div[data-baseweb="popover"] *,
    div[role="dialog"] *, div[role="dialog"] {
        background-color: #ffffff !important;
        color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }
    [data-testid="stDateInput"] svg, div[data-baseweb="datepicker"] svg {
        fill: #1976d2 !important;
        color: #1976d2 !important;
    }

    .stButton > button, .stButton > button p, .stButton > button span, .stButton > button div,
    .stDownloadButton > button, .stDownloadButton > button p,
    .stDownloadButton > button span, .stDownloadButton > button div {
        color: #ffffff !important;
    }

    h1[style*="color:#1976d2"], h1[style*="color: #1976d2"] { color: #1976d2 !important; }
    .metric-card h3 { color: #0277bd !important; }
    .metric-card h1 { color: #1976d2 !important; }
    .person-card .name { color: #1976d2 !important; }
    .person-card .sub { color: #0277bd !important; }
    .person-card .badge { color: #ffffff !important; }
    .lf-simple-card .lf-line1 { color: #1976d2 !important; }
    .lf-simple-card .lf-line2 { color: #0277bd !important; }
    .lf-simple-card .lf-boxes { color: #ffffff !important; }
    .sal-metric { color: #000000 !important; }
    .sal-metric.base { color: #0d47a1 !important; }
    .sal-metric.adv { color: #e65100 !important; }
    .sal-metric.short { color: #c62828 !important; }
    .sal-metric.remain { color: #1b5e20 !important; }
    .sal-metric.paid { color: #2e7d32 !important; }
    .hint-box { color: #1976d2 !important; }
    .summary-box { color: #000000 !important; }
    .booker-row { color: #1976d2 !important; }

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
    .stApp [data-testid="stVerticalBlockBorderWrapper"] > div { gap: 0.35rem !important; }
    .stApp label, .stApp [data-testid="stWidgetLabel"] label, .stApp [data-testid="stWidgetLabel"] p {
        font-size: 13px !important; font-weight: 600 !important; margin-bottom: 2px !important;
    }
    .stTextInput > div > div > input, .stNumberInput > div > div > input,
    .stDateInput > div > div > input, .stSelectbox > div > div > div,
    div[data-baseweb="input"] input, div[data-baseweb="select"] > div {
        padding: 4px 8px !important; min-height: 34px !important; font-size: 14px !important;
    }
    .stNumberInput button { padding: 2px 4px !important; min-height: 34px !important; }
    .stButton > button { padding: 4px 10px !important; min-height: 36px !important; font-size: 14px !important; }
    .stApp hr { margin: 6px 0 !important; }
    .stApp p { margin-bottom: 3px !important; }
    .stApp h3 { margin-top: 6px !important; margin-bottom: 4px !important; font-size: 18px !important; }

    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #bbdefb 0%, #90caf9 100%) !important; }
    section[data-testid="stSidebar"] * { color: #000000 !important; }
    section[data-testid="stSidebar"] .stRadio label { font-size: 15px !important; font-weight: 600 !important; }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label {
        background-color: #ffffff !important; border: 1px solid #90caf9 !important;
        border-radius: 8px !important; margin-bottom: 6px !important; padding: 6px 10px !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
        background-color: #e3f2fd !important; border-color: #2196f3 !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] input[type="radio"] { accent-color: #2196f3 !important; }

    .stTextInput > div > div > input, .stNumberInput > div > div > input,
    .stSelectbox > div > div > div, .stSelectbox > div > div,
    div[data-baseweb="select"] > div, div[data-baseweb="input"] input {
        background-color: #ffffff !important; color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
        border: 2px solid #90caf9 !important; border-radius: 8px !important;
    }
    div[data-baseweb="select"] * { color: #000000 !important; }
    ul[role="listbox"] li, div[role="option"] { color: #000000 !important; background-color: #ffffff !important; }
    .stTextInput > div > div > input:focus, .stNumberInput > div > div > input:focus { border-color: #2196f3 !important; }
    .stTextInput input::placeholder, .stNumberInput input::placeholder {
        color: #888888 !important; -webkit-text-fill-color: #888888 !important; opacity: 1 !important;
    }
    input[type="password"] {
        background-color: #ffffff !important; color: #000000 !important;
        -webkit-text-fill-color: #000000 !important;
    }

    .stButton > button {
        background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%) !important;
        border: none !important; font-weight: bold !important; border-radius: 8px !important;
    }
    .stButton > button:hover { background: linear-gradient(135deg, #1976d2 0%, #0d47a1 100%) !important; }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #0288d1 0%, #0277bd 100%) !important;
        color: #ffffff !important; font-weight: bold !important; border-radius: 8px !important;
    }

    .stDataFrame, .stDataFrame * { color: #000000 !important; }
    .stDataFrame { background-color: #ffffff !important; border-radius: 10px !important; }
    label[data-baseweb="checkbox"] * { color: #000000 !important; }
    details summary, details summary *, .streamlit-expanderHeader, .streamlit-expanderHeader * { color: #000000 !important; }

    .stAlert, .stAlert * { color: #000000 !important; }
    .stSuccess, .stSuccess * { color: #1b5e20 !important; }
    .stError, .stError * { color: #b71c1c !important; }
    .stWarning, .stWarning * { color: #e65100 !important; }
    .stInfo, .stInfo * { color: #0d47a1 !important; }
    [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * { color: #555555 !important; }
    .empty-box { color: #0277bd !important; }
    .auto-dl-hidden div[data-testid="stDownloadButton"] {
        position: absolute !important; left: -9999px !important; top: -9999px !important;
        opacity: 0 !important; height: 0 !important;
    }

    .metric-card {
        background: #ffffff; border: 2px solid #90caf9; border-radius: 14px;
        padding: 22px; text-align: center; box-shadow: 0 3px 10px rgba(33,150,243,0.15);
    }
    .metric-card h3 { font-size: 13px !important; margin: 0 !important; font-weight: 700 !important; text-transform: uppercase; }
    .metric-card h1 { font-size: 34px !important; margin: 10px 0 0 0 !important; font-weight: 800 !important; }

    .booker-row { background: #ffffff; border: 1px solid #90caf9; border-radius: 10px; padding: 12px 18px; margin-bottom: 8px; }

    .person-card {
        background: #ffffff; border-left: 6px solid #2196f3; border-radius: 12px;
        padding: 14px 18px; margin-bottom: 10px;
        box-shadow: 0 3px 10px rgba(33,150,243,0.12);
        display: flex; align-items: center; justify-content: space-between;
    }
    .person-card .info { display: flex; flex-direction: column; gap: 3px; }
    .person-card .name { font-size: 16px; font-weight: 700; }
    .person-card .sub { font-size: 12px; }
    .person-card .badge {
        background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
        font-weight: 800; font-size: 14px; padding: 6px 12px; border-radius: 8px;
        min-width: 60px; text-align: center;
    }
    .sal-badge { background: linear-gradient(135deg, #0288d1 0%, #0277bd 100%); }

    .empty-box {
        background: #ffffff; border: 2px dashed #90caf9; border-radius: 12px;
        padding: 20px; text-align: center; font-size: 14px;
    }
    .hint-box {
        background: #e3f2fd; border-left: 4px solid #2196f3; padding: 6px 10px;
        border-radius: 6px; font-size: 12px; margin-top: 2px;
    }
    .summary-box {
        background: #ffffff; border: 2px solid #2196f3; border-radius: 12px;
        padding: 12px 18px; margin-bottom: 12px;
    }
    .lf-simple-card {
        background: #ffffff; border-left: 6px solid #2196f3; border-radius: 12px;
        padding: 16px 22px; margin-bottom: 12px;
        box-shadow: 0 3px 10px rgba(33,150,243,0.15);
        display: flex; align-items: center; justify-content: space-between;
    }
    .lf-simple-card .lf-info { display: flex; flex-direction: column; gap: 4px; }
    .lf-simple-card .lf-line1 { font-size: 17px; font-weight: 700; }
    .lf-simple-card .lf-line2 { font-size: 13px; }
    .lf-simple-card .lf-boxes {
        background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
        font-weight: 800; font-size: 20px; padding: 10px 18px; border-radius: 10px;
        text-align: center; min-width: 90px;
    }
    .lf-simple-card .lf-boxes small { display: block; font-size: 10px; font-weight: 500; opacity: 0.9; }

    .sal-metric {
        display: inline-block; padding: 8px 14px; margin-right: 8px; margin-bottom: 6px;
        border-radius: 8px; font-size: 13px; font-weight: 600;
    }
    .sal-metric.base { background: #e3f2fd; }
    .sal-metric.adv { background: #fff3e0; }
    .sal-metric.short { background: #ffebee; }
    .sal-metric.remain { background: #e8f5e9; }
    .sal-metric.paid { background: #c8e6c9; }

    .exp-row {
        background: #ffffff; border: 1px solid #90caf9; border-radius: 10px;
        padding: 8px 14px; margin-bottom: 6px;
        display: flex; align-items: center; justify-content: space-between;
    }
    .exp-row .exp-left { display: flex; flex-direction: column; gap: 2px; }
    .exp-row .exp-name { font-size: 14px; font-weight: 700; color: #1976d2; }
    .exp-row .exp-date { font-size: 11px; color: #0277bd; }
    .exp-row .exp-amt {
        background: linear-gradient(135deg, #2196f3 0%, #1976d2 100%);
        color: #ffffff !important; font-weight: 800; font-size: 15px;
        padding: 5px 12px; border-radius: 8px; min-width: 80px; text-align: center;
    }
    .exp-row.lunch .exp-amt { background: linear-gradient(135deg, #ff9800 0%, #f57c00 100%); }
    .exp-row.lunch .exp-name { color: #e65100; }

    .status-pending {
        background: #fff3e0; color: #e65100 !important; padding: 3px 10px;
        border-radius: 6px; font-size: 11px; font-weight: 700;
    }
    .status-paid {
        background: #c8e6c9; color: #1b5e20 !important; padding: 3px 10px;
        border-radius: 6px; font-size: 11px; font-weight: 700;
    }

    .disc-card {
        background: #ffffff; border: 2px solid #90caf9; border-radius: 12px;
        padding: 10px 14px; margin-bottom: 10px;
    }
    .disc-card.active {
        border: 3px solid #2e7d32;
        box-shadow: 0 4px 14px rgba(46,125,50,0.25);
        background: #f1f8e9;
    }
    .disc-title { font-size: 15px; font-weight: 800; color: #1976d2; margin-bottom: 4px; }
    .disc-title-active { color: #2e7d32 !important; }
    .disc-badge-active {
        background: #c8e6c9; color: #1b5e20 !important;
        font-size: 11px; font-weight: 700; padding: 3px 10px;
        border-radius: 6px; margin-left: 8px;
    }

    .full-bill-box {
        background: #ffffff; border: 3px solid #2196f3; border-radius: 14px;
        padding: 18px 22px; margin: 10px 0 18px 0;
        box-shadow: 0 6px 20px rgba(33,150,243,0.25);
    }
    .full-bill-title { font-size: 20px; font-weight: 800; color: #1976d2; margin-bottom: 8px; }
    .full-bill-meta { font-size: 13px; color: #0277bd; margin-bottom: 12px; }

    .stAlert { border-radius: 10px !important; }
    hr { border-color: #90caf9 !important; opacity: 0.6 !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# SIDEBAR TOGGLE
# ============================================================
def inject_sidebar_toggle():
    components.html("""
    <script>
    (function(){
        function killManageApp() {
            try {
                var doc = window.parent.document;
                var sel = ['[data-testid="manage-app-button"]','[data-testid="stAppDeployButton"]',
                    '[data-testid="stCloudAppManageButton"]','.stAppDeployButton',
                    'iframe[title="streamlit_cloud_status"]','div[class*="manageApp"]',
                    'div[class*="ManageApp"]','button[class*="manageApp"]','button[class*="ManageApp"]'];
                sel.forEach(function(s){
                    doc.querySelectorAll(s).forEach(function(el){
                        el.style.setProperty('display','none','important');
                        el.style.setProperty('visibility','hidden','important');
                        el.style.setProperty('opacity','0','important');
                    });
                });
                doc.querySelectorAll('button, a').forEach(function(el){
                    try {
                        var t = (el.textContent || '').trim();
                        if (t === 'Manage app' || t === 'Manage App') el.style.setProperty('display','none','important');
                    } catch(e){}
                });
            } catch(e) {}
        }
        function attachToggle() {
            try {
                var doc = window.parent.document;
                if (!doc.querySelector('section[data-testid="stSidebar"]')) return;
                var old = doc.getElementById('custom-sidebar-toggle');
                if (old) old.parentNode.removeChild(old);
                var btn = doc.createElement('button');
                btn.id = 'custom-sidebar-toggle';
                btn.title = 'Sidebar';
                btn.innerHTML = '\\u2630';
                var s = {
                    'position':'fixed','top':'14px','left':'14px','z-index':'2147483647',
                    'background':'linear-gradient(135deg, #2196f3 0%, #1976d2 100%)',
                    'color':'#fff','border':'none','border-radius':'10px','padding':'8px 14px',
                    'font-size':'20px','font-weight':'bold','cursor':'pointer',
                    'box-shadow':'0 3px 10px rgba(33,150,243,0.5)'
                };
                for (var k in s) btn.style.setProperty(k, s[k], 'important');
                btn.onclick = function() {
                    var targets = ['[data-testid="stSidebarCollapseButton"] button',
                        '[data-testid="stSidebarCollapsedControl"] button',
                        '[data-testid="collapsedControl"] button',
                        '[data-testid="stExpandSidebarButton"] button',
                        'button[kind="headerNoPadding"]'];
                    for (var i = 0; i < targets.length; i++) {
                        var el = doc.querySelector(targets[i]);
                        if (el) { el.click(); return; }
                    }
                };
                if (doc.body) doc.body.appendChild(btn);
            } catch(e) {}
        }
        function tick() { killManageApp(); attachToggle(); }
        setTimeout(tick, 300); setTimeout(tick, 1000); setTimeout(tick, 2000);
        setInterval(tick, 1500);
    })();
    </script>
    """, height=0)

# ============================================================
# AUTH PAGE
# ============================================================
if not st.session_state.get("logged_in_user"):
    random.seed(7)
    _dust = ""
    for _ in range(30):
        _l = random.randint(8, 92)
        _t = random.randint(0, 62)
        _dur = round(random.uniform(5, 12), 2)
        _dly = round(random.uniform(0, 5), 2)
        _sz = random.choice([1, 1, 2, 2])
        _dust += (
            f'<span style="left:{_l}%;top:{_t}%;width:{_sz}px;height:{_sz}px;'
            f'animation-duration:{_dur}s;animation-delay:{_dly}s;"></span>'
        )

    _light_on = st.session_state.get("light_on", False)
    _lit = "lit" if _light_on else ""
    _cord_href = "?lamp=off" if _light_on else "?lamp=on"
    _hint = "☀ LIGHT ON · Fill the form below" if _light_on else "▼ PULL THE CORD ▼"
    _form_top = "420px" if _light_on else "260px"

    st.markdown(f"""
    <style>
        html, body, .stApp, [data-testid="stAppViewContainer"] {{
            background: #0a0c10 !important;
        }}
        header[data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        #MainMenu, footer,
        section[data-testid="stSidebar"],
        .stAppDeployButton,
        [data-testid="manage-app-button"] {{
            display: none !important;
            visibility: hidden !important;
        }}

        .block-container {{
            padding-top: {_form_top} !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
            padding-bottom: 3rem !important;
            max-width: 440px !important;
            margin: 0 auto !important;
            position: relative;
            z-index: 100;
            transition: padding-top 0.9s cubic-bezier(0.2,0.9,0.3,1);
        }}

        .lamp-scene {{
            position: fixed;
            inset: 0;
            z-index: 0;
            pointer-events: none;
            overflow: hidden;
            background:
              radial-gradient(ellipse at 50% 5%, rgba(38,42,48,0.72) 0%, rgba(22,25,29,0.72) 24%, rgba(11,13,16,0.98) 65%),
              linear-gradient(180deg, #171a1f 0%, #0e1013 52%, #090b0e 100%);
            transition: background 1s ease;
        }}
        .lamp-scene.lit {{
            background:
              radial-gradient(ellipse at 50% 17%, rgba(94,65,30,0.55) 0%, rgba(43,31,19,0.48) 22%, rgba(18,16,14,0.92) 55%, #0a0c0f 100%),
              linear-gradient(180deg, #1b1d20 0%, #111214 55%, #0a0c0f 100%);
        }}

        .lamp-scene .ceiling {{
            position: absolute; top: 0; left: 50%;
            transform: translateX(-50%);
            width: 76px; height: 14px;
            background: linear-gradient(180deg, #4b5057 0%, #22262b 38%, #0b0d10 100%);
            border-radius: 0 0 9px 9px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.75), inset 0 1px 0 rgba(255,255,255,0.12);
        }}
        .lamp-scene .ceiling::before {{
            content: ''; position: absolute; top: 0; left: 50%;
            transform: translateX(-50%);
            width: 32px; height: 4px;
            background: #747980; border-radius: 3px;
        }}
        .lamp-scene .wire {{
            position: absolute; top: 12px; left: 50%;
            transform: translateX(-50%);
            width: 3px; height: 124px;
            background: linear-gradient(90deg, #1b1e22 0%, #9da2a8 50%, #25282c 100%);
            box-shadow: 0 0 5px rgba(0,0,0,0.85);
        }}
        .lamp-scene .lamp {{
            position: absolute; top: 128px; left: 50%;
            transform: translateX(-50%);
            width: 250px; height: 150px;
        }}
        .lamp-scene .lamp::before {{
            content: ''; position: absolute; top: -8px; left: 50%;
            transform: translateX(-50%);
            width: 54px; height: 13px; border-radius: 50%;
            background: linear-gradient(180deg, #9a7041, #3a2818 70%, #17100a);
            box-shadow: 0 3px 8px rgba(0,0,0,0.8);
        }}
        .lamp-scene .lamp-shade {{
            position: absolute; top: 0; left: 15px;
            width: 220px; height: 104px;
            border-radius: 112px 112px 22px 22px / 100px 100px 24px 24px;
            background:
              radial-gradient(ellipse at 50% 3%, #6d4b29 0%, #3d2a19 23%, #21160d 52%, #0d0905 88%),
              linear-gradient(180deg, #4b3420 0%, #120c07 100%);
            border: 1px solid rgba(170,120,65,0.24);
            box-shadow:
              inset 0 5px 14px rgba(196,145,80,0.22),
              inset 0 -20px 34px rgba(0,0,0,0.95),
              0 8px 28px rgba(0,0,0,0.78);
            transition: box-shadow 0.9s ease, filter 0.9s ease;
        }}
        .lamp-scene.lit .lamp-shade {{
            filter: brightness(1.13);
            box-shadow:
              inset 0 5px 14px rgba(255,210,135,0.34),
              inset 0 -20px 34px rgba(0,0,0,0.9),
              0 8px 28px rgba(0,0,0,0.72),
              0 0 72px 18px rgba(255,194,82,0.23);
        }}
        .lamp-scene .bulb {{
            position: absolute; top: 80px; left: 50%;
            transform: translateX(-50%);
            width: 50px; height: 45px;
            border-radius: 50% 50% 44% 44% / 56% 56% 44% 44%;
            background: radial-gradient(circle at 50% 35%, #383b3f 0%, #17191c 60%, #070809 100%);
            transition: all 0.7s cubic-bezier(0.2, 0.9, 0.3, 1);
            box-shadow: inset 0 -7px 13px rgba(0,0,0,0.9);
        }}
        .lamp-scene.lit .bulb {{
            background: radial-gradient(circle at 50% 34%, #ffffff 0%, #fff9df 22%, #ffd96a 53%, #ff9e1a 100%);
            box-shadow:
              0 0 24px 9px rgba(255,239,174,0.98),
              0 0 58px 24px rgba(255,202,84,0.72),
              0 0 120px 50px rgba(255,169,42,0.38);
        }}
        .lamp-scene .beam {{
            position: absolute;
            top: 210px; left: 50%;
            transform: translateX(-50%);
            width: min(980px, 90vw); height: 760px;
            background: radial-gradient(
              ellipse at 50% 0%,
              rgba(255,244,198,0.58) 0%, rgba(255,226,146,0.34) 18%,
              rgba(255,202,91,0.16) 42%, rgba(255,180,55,0.055) 66%,
              transparent 84%);
            opacity: 0;
            transition: opacity 1.05s ease;
            filter: blur(8px);
        }}
        .lamp-scene.lit .beam {{ opacity: 1; }}

        .lamp-scene .dust {{
            position: absolute;
            top: 235px; left: 50%;
            transform: translateX(-50%);
            width: 760px; height: 690px;
            opacity: 0;
            transition: opacity 1.2s ease 0.2s;
        }}
        .lamp-scene.lit .dust {{ opacity: 1; }}
        .lamp-scene .dust span {{
            position: absolute;
            width: 2px; height: 2px;
            background: radial-gradient(circle, #fff8e1 0%, #ffd86a 60%, transparent 100%);
            border-radius: 50%;
            box-shadow: 0 0 5px #ffd05a;
            animation: dust-drift linear infinite;
        }}
        @keyframes dust-drift {{
            0%   {{ transform: translateY(0) translateX(0); opacity: 0; }}
            20%  {{ opacity: 0.85; }}
            80%  {{ opacity: 0.85; }}
            100% {{ transform: translateY(220px) translateX(40px); opacity: 0; }}
        }}

        .lamp-scene .cord {{
            position: absolute;
            top: 126px; left: 50%;
            transform: translateX(57px);
            width: 26px; height: 250px;
            cursor: pointer;
            pointer-events: auto;
            text-decoration: none;
            transition: transform 0.25s ease;
        }}
        .lamp-scene .cord::before {{
            content: ''; position: absolute; top: 0; left: 12px;
            width: 2px; height: 222px;
            background: linear-gradient(180deg, #292d32 0%, #b7bdc3 38%, #60666d 66%, #171a1d 100%);
        }}
        .lamp-scene .cord::after {{
            content: ''; position: absolute; top: 0; left: 7px;
            width: 12px; height: 12px; border-radius: 50%;
            background: radial-gradient(circle at 35% 30%, #e8c98f 0%, #8c6132 45%, #2d1a0b 100%);
        }}
        .lamp-scene .cord .bead {{
            position: absolute;
            bottom: 0; left: 50%;
            transform: translateX(-50%);
            width: 22px; height: 28px;
            border-radius: 50% 50% 45% 45% / 60% 60% 40% 40%;
            background: radial-gradient(circle at 34% 28%, #f0ca82 0%, #a56c31 38%, #5a3516 70%, #241408 100%);
            border: 1px solid rgba(244,205,140,0.24);
            box-shadow:
              0 4px 10px rgba(0,0,0,0.88),
              0 0 16px 2px rgba(255,190,85,0.34),
              inset 0 -3px 5px rgba(0,0,0,0.5);
            transition: all 0.25s ease;
        }}
        .lamp-scene .cord:hover .bead {{
            transform: translateX(-50%) scale(1.16);
            box-shadow:
              0 4px 10px rgba(0,0,0,0.88),
              0 0 28px 7px rgba(255,202,100,0.9),
              inset 0 -3px 5px rgba(0,0,0,0.5);
        }}
        .lamp-scene .cord:active {{ transform: translateX(57px) scaleY(1.3); }}
        .lamp-scene .cord:hover {{ transform: translateX(57px) scaleY(1.05); }}

        .bottom-hint {{
            position: fixed;
            bottom: 38px; left: 50%;
            transform: translateX(-50%);
            color: #8a8f96;
            font-size: 11px;
            letter-spacing: 4px;
            text-transform: uppercase;
            font-weight: 500;
            animation: lamp-pulse 2.8s ease-in-out infinite;
            z-index: 50;
            text-align: center;
            white-space: nowrap;
            pointer-events: none;
        }}
        .bottom-hint.lit {{
            color: #a98a5e;
            animation: none;
        }}
        @keyframes lamp-pulse {{ 0%,100% {{ opacity: 0.42; }} 50% {{ opacity: 1; }} }}

        [data-testid="stForm"] {{
            background: linear-gradient(145deg, rgba(30,31,33,0.88), rgba(15,16,18,0.82)) !important;
            border: 1px solid rgba(255,203,113,0.28) !important;
            border-radius: 20px !important;
            padding: 26px 28px 22px !important;
            backdrop-filter: blur(18px) saturate(115%) !important;
            -webkit-backdrop-filter: blur(18px) saturate(115%) !important;
            box-shadow:
              0 20px 60px rgba(0,0,0,0.6),
              inset 0 1px 0 rgba(255,255,255,0.08),
              0 0 70px rgba(255,188,75,0.1) !important;
        }}

        [data-testid="stForm"] label,
        [data-testid="stForm"] [data-testid="stWidgetLabel"] p,
        [data-testid="stForm"] [data-testid="stWidgetLabel"] label {{
            color: #a98b62 !important;
            font-size: 10px !important;
            letter-spacing: 2px !important;
            text-transform: uppercase !important;
            font-weight: 750 !important;
        }}

        [data-testid="stForm"] input {{
            background: rgba(5,6,8,0.65) !important;
            color: #fff9e8 !important;
            -webkit-text-fill-color: #fff9e8 !important;
            border: 1px solid rgba(255,205,113,0.2) !important;
            border-radius: 11px !important;
            padding: 11px 14px !important;
            font-size: 14px !important;
            caret-color: #f4c96f !important;
        }}
        [data-testid="stForm"] input:focus {{
            border-color: rgba(255,205,113,0.6) !important;
            box-shadow: 0 0 0 3px rgba(255,200,100,0.1), 0 0 24px rgba(255,180,50,0.15) !important;
            outline: none !important;
        }}
        [data-testid="stForm"] input::placeholder {{
            color: #675a46 !important;
            -webkit-text-fill-color: #675a46 !important;
            font-style: italic !important;
        }}

        [data-testid="stForm"] button[kind="primary"],
        [data-testid="stForm"] button[kind="primaryFormSubmit"] {{
            background: linear-gradient(135deg, rgba(255,205,105,0.95), rgba(255,154,45,0.9)) !important;
            color: #100803 !important;
            font-weight: 850 !important;
            letter-spacing: 2px !important;
            text-transform: uppercase !important;
            border-radius: 11px !important;
            border: none !important;
            padding: 11px 20px !important;
            box-shadow: 0 4px 18px rgba(255,150,50,0.35), inset 0 1px 0 rgba(255,255,255,0.3) !important;
            transition: all 0.25s ease !important;
        }}
        [data-testid="stForm"] button[kind="primary"]:hover,
        [data-testid="stForm"] button[kind="primaryFormSubmit"]:hover {{
            background: linear-gradient(135deg, rgba(255,220,140,1), rgba(255,180,60,0.95)) !important;
            box-shadow: 0 6px 24px rgba(255,150,50,0.5) !important;
            transform: translateY(-1px);
        }}

        .stTabs [data-baseweb="tab-list"] {{
            background: rgba(0,0,0,0.45) !important;
            border-radius: 12px !important;
            padding: 4px !important;
            gap: 4px !important;
            border: 1px solid rgba(255,203,113,0.12) !important;
        }}
        .stTabs [data-baseweb="tab"] {{
            background: transparent !important;
            border-radius: 9px !important;
            padding: 9px 18px !important;
        }}
        .stTabs [data-baseweb="tab"] p {{
            color: #83745e !important;
            font-weight: 800 !important;
            font-size: 12px !important;
            letter-spacing: 1.5px !important;
        }}
        .stTabs [aria-selected="true"] {{
            background: linear-gradient(135deg, rgba(255,205,105,0.2), rgba(255,154,45,0.1)) !important;
            box-shadow: 0 2px 10px rgba(255,180,50,0.15), inset 0 0 0 1px rgba(255,210,120,0.25) !important;
        }}
        .stTabs [aria-selected="true"] p {{ color: #f5cc72 !important; }}
        .stTabs [data-baseweb="tab-highlight"],
        .stTabs [data-baseweb="tab-border"] {{ display: none !important; }}

        .auth-brand {{
            text-align: center;
            margin-bottom: 16px;
        }}
        .auth-brand h1 {{
            color: #f4c96f !important;
            font-size: 21px !important;
            font-weight: 850 !important;
            letter-spacing: 4px !important;
            text-shadow: 0 0 20px rgba(255,199,92,0.4) !important;
            margin: 0 0 3px 0 !important;
        }}
        .auth-brand p {{
            color: #a98a5e !important;
            font-size: 10px !important;
            letter-spacing: 5px !important;
            font-weight: 700 !important;
            margin: 0 !important;
        }}

        [data-testid="stAlert"] {{
            background: rgba(200,40,40,0.18) !important;
            border: 1px solid rgba(255,100,100,0.4) !important;
            border-radius: 11px !important;
        }}
        [data-testid="stAlert"] * {{ color: #ffb3b3 !important; }}

        @media (max-width: 600px) {{
            .lamp-scene .lamp {{
                top: 105px;
                transform: translateX(-50%) scale(0.82);
                transform-origin: top center;
            }}
            .lamp-scene .wire {{ height: 102px; }}
            .lamp-scene .cord {{
                top: 103px;
                transform: translateX(49px) scale(0.82);
                transform-origin: top left;
            }}
            .block-container {{ padding-top: 300px !important; }}
        }}
    </style>

    <div class="lamp-scene {_lit}">
        <div class="ceiling"></div>
        <div class="wire"></div>
        <div class="lamp">
            <div class="lamp-shade"></div>
            <div class="bulb"></div>
        </div>
        <div class="beam"></div>
        <div class="dust">{_dust}</div>
        <a href="{_cord_href}" class="cord" title="Pull cord">
            <div class="bead"></div>
        </a>
    </div>
    <div class="bottom-hint {_lit}">{_hint}</div>
    """, unsafe_allow_html=True)

    if _light_on:
        st.markdown("""
        <div class="auth-brand">
            <h1>AL-BARAKAH</h1>
            <p>ENTERPRISES</p>
        </div>
        """, unsafe_allow_html=True)

        tab1, tab2 = st.tabs(["🔐 LOGIN", "📝 SIGNUP"])

        with tab1:
            with st.form("login_form", clear_on_submit=False):
                login_user = st.text_input("USERNAME", key="login_u", placeholder="enter username")
                login_pass = st.text_input("PASSWORD", key="login_p", type="password", placeholder="enter password")
                sub1 = st.form_submit_button("🔓 LOGIN", use_container_width=True, type="primary")

                if sub1:
                    users = load_users()
                    u = sanitize_username(login_user.strip())
                    if u == "" or login_pass == "":
                        st.error("❌ Please fill all fields")
                    elif u not in users:
                        st.error("❌ Username not found. Please signup first.")
                    elif users[u].get("password_hash") != hash_password(login_pass):
                        st.error("❌ Incorrect password")
                    else:
                        st.session_state["logged_in_user"] = u
                        st.session_state["display_name"] = users[u].get("display_name", u)
                        st.session_state["page"] = "📊 Dashboard"
                        st.session_state["light_on"] = False
                        st.rerun()

        with tab2:
            with st.form("signup_form", clear_on_submit=False):
                su_user = st.text_input("NEW USERNAME", key="su_u", placeholder="min 3 letters/numbers")
                su_pass = st.text_input("NEW PASSWORD", key="su_p", type="password", placeholder="min 4 characters")
                su_pass2 = st.text_input("CONFIRM PASSWORD", key="su_p2", type="password", placeholder="repeat password")
                sub2 = st.form_submit_button("✅ CREATE ACCOUNT", use_container_width=True, type="primary")

                if sub2:
                    users = load_users()
                    u = sanitize_username(su_user.strip())
                    if u == "":
                        st.error("❌ Username required")
                    elif len(u) < 3:
                        st.error("❌ Username min 3 characters")
                    elif len(u) > 20:
                        st.error("❌ Username max 20 characters")
                    elif len(su_pass) < 4:
                        st.error("❌ Password min 4 characters")
                    elif su_pass != su_pass2:
                        st.error("❌ Passwords do not match")
                    elif u in users:
                        st.error(f"❌ Username '{u}' already taken")
                    else:
                        users[u] = {
                            "username": u,
                            "display_name": su_user.strip(),
                            "password_hash": hash_password(su_pass),
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        }
                        save_users(users)
                        with open(user_data_file(u), "w", encoding="utf-8") as f:
                            json.dump(default_blank_db(), f, indent=4, ensure_ascii=False)
                        st.session_state["logged_in_user"] = u
                        st.session_state["display_name"] = su_user.strip()
                        st.session_state["page"] = "📊 Dashboard"
                        st.session_state["light_on"] = False
                        st.rerun()

    st.stop()

# ============================================================
# LOGGED IN — LOAD USER DATA
# ============================================================
CURRENT_USER = st.session_state["logged_in_user"]

def load_database(username):
    fpath = user_data_file(username)
    if os.path.exists(fpath):
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k in ["bookers", "salesmen", "load_forms", "petrol_expenses", "lunch_expenses"]:
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
                return data
        except Exception:
            pass
    return default_blank_db()

def save_database(db):
    try:
        fpath = user_data_file(CURRENT_USER)
        with open(fpath, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=4, ensure_ascii=False)
    except Exception as e:
        st.warning(f"⚠️ Save warning: {e}")

def parse_date(dstr):
    try:
        return datetime.strptime(dstr, "%d-%m-%Y").date()
    except Exception:
        return None

if "database" not in st.session_state:
    st.session_state.database = load_database(CURRENT_USER)

if st.session_state.get("_db_user") != CURRENT_USER:
    st.session_state.database = load_database(CURRENT_USER)
    st.session_state["_db_user"] = CURRENT_USER

if "_prev_prod" not in st.session_state:
    st.session_state["_prev_prod"] = None
if "last_bill_no" not in st.session_state:
    st.session_state["last_bill_no"] = None
if "download_file" not in st.session_state:
    st.session_state["download_file"] = None
if "_dl_counter" not in st.session_state:
    st.session_state["_dl_counter"] = 0
if "page" not in st.session_state:
    st.session_state["page"] = "📊 Dashboard"
if "view_bill_key" not in st.session_state:
    st.session_state["view_bill_key"] = None
if "view_lf_key" not in st.session_state:
    st.session_state["view_lf_key"] = None

db = st.session_state.database
for k in ["bookers", "salesmen", "load_forms", "petrol_expenses", "lunch_expenses"]:
    if k not in db: db[k] = []
for k in ["bookers_salaries", "salesmen_salaries", "product_prices"]:
    if k not in db: db[k] = {}
if "discount_packages" not in db or not db["discount_packages"]:
    db["discount_packages"] = default_discount_packages()
for p in db["discount_packages"]:
    if "tier3_amount" not in p: p["tier3_amount"] = 0.0
    if "tier3_pct" not in p: p["tier3_pct"] = 0.0

inject_sidebar_toggle()

# ============================================================
# HELPERS
# ============================================================
def get_price(code, base_price):
    custom = db.get("product_prices", {})
    if str(code) in custom:
        try: return float(custom[str(code)])
        except Exception: return float(base_price)
    return float(base_price)

def get_all_active_packages():
    return [p for p in db.get("discount_packages", []) if p.get("active")]

def get_package_discount_pct(bill_total):
    active_pkgs = get_all_active_packages()
    if not active_pkgs:
        return 0.0, None, None
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
                best_tier = f"{label} (> {amt:,.0f})"
                break
    if best_pct > 0:
        return best_pct, best_name, best_tier
    return 0.0, (active_pkgs[0].get("name") if active_pkgs else None), None

def show_auto_download():
    if st.session_state.get("download_file"):
        fname, fdata = st.session_state["download_file"]
        st.session_state["download_file"] = None
        st.session_state["_dl_counter"] += 1
        st.markdown('<div class="auto-dl-hidden">', unsafe_allow_html=True)
        st.download_button(
            label=f"Download {fname}",
            data=fdata, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"auto_dl_{st.session_state['_dl_counter']}",
        )
        st.markdown('</div>', unsafe_allow_html=True)
        components.html("""
        <script>
        (function(){
            var tries = 0;
            function attempt(){
                tries++;
                var btns = window.parent.document.querySelectorAll('[data-testid="stDownloadButton"] button');
                if (btns.length > 0){ btns[btns.length - 1].click(); return; }
                if (tries < 25) setTimeout(attempt, 200);
            }
            setTimeout(attempt, 300);
        })();
        </script>
        """, height=0)

def export_single_group_bill(shop, date_str, booker, salesman, items, bill_no):
    bill_total_net = sum(float(it.get("Net", 0)) for it in items)
    pkg_pct, pkg_name, tier_label = get_package_discount_pct(bill_total_net)

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Bill")
    worksheet.set_paper(9); worksheet.set_portrait(); worksheet.fit_to_pages(1, 1)
    worksheet.set_column("A:A", 42.86); worksheet.set_column("B:B", 12.71)
    worksheet.set_column("C:C", 10.71); worksheet.set_column("D:D", 10.71)
    worksheet.set_column("E:E", 11.71); worksheet.set_column("F:F", 11.14)
    worksheet.set_column("G:G", 11.14); worksheet.set_column("H:H", 12.14)
    worksheet.set_column("I:I", 13.14); worksheet.set_column("J:J", 13.14)

    title = workbook.add_format({"bold":True, "font_size":18, "align":"center", "border":2})
    header = workbook.add_format({"bold":True, "font_size":11, "bg_color":"#BBDEFB", "align":"center", "border":2, "text_wrap": True})
    cell_left = workbook.add_format({"font_size":12, "border":1, "align":"left"})
    cell_center = workbook.add_format({"font_size":12, "border":1, "align":"center"})
    total = workbook.add_format({"bold":True, "font_size":12, "bg_color":"#FFF2CC", "align":"center", "border":2})
    disc_hl = workbook.add_format({"font_size":12, "border":1, "align":"center", "bg_color":"#E8F5E9", "bold": True})
    pkg_info = workbook.add_format({"font_size":10, "italic": True, "align":"left", "font_color":"#1b5e20"})

    worksheet.merge_range("A1:J1", COMPANY_NAME, title)
    worksheet.write("A3","Shop Name",header); worksheet.write("B3", shop, cell_center)
    worksheet.write("D3","Booker",header); worksheet.write("E3", booker, cell_center)
    worksheet.write("G3","Bill No",header); worksheet.write("H3", bill_no, cell_center)
    worksheet.write("I3","Date",header); worksheet.write("J3", date_str, cell_center)

    if pkg_pct > 0:
        worksheet.merge_range("A4:J4", f"🎁 Best Discount Applied: {pkg_name} | {tier_label} | {pkg_pct}% on each product", pkg_info)
    else:
        worksheet.merge_range("A4:J4", "💡 Koi discount apply nahi hua", pkg_info)

    start_row = 6
    headers = ["Product", "Code", "Boxes", "TP/Box", "Gross", "Disc %", "Net", "Pkg Disc %", "After Disc Net", "Saved"]
    for col, h in enumerate(headers):
        worksheet.write(start_row, col, h, header)
    row = start_row + 1

    gross_total = 0; total_boxes = 0; net_total = 0; after_disc_total = 0; saved_total = 0
    for it in items:
        b_net = float(it.get("Net", 0)); b_gross = float(it.get("Gross", 0))
        b_boxes = int(it.get("Boxes", 0))
        after_net = b_net - (b_net * pkg_pct / 100)
        saved = b_net - after_net

        worksheet.write(row, 0, it.get("Product",""), cell_left)
        worksheet.write(row, 1, it.get("Code",""), cell_center)
        worksheet.write(row, 2, b_boxes, cell_center)
        worksheet.write(row, 3, it.get("TP/Box", 0), cell_center)
        worksheet.write(row, 4, b_gross, cell_center)
        worksheet.write(row, 5, it.get("Discount %", 0), cell_center)
        worksheet.write(row, 6, b_net, cell_center)
        if pkg_pct > 0:
            worksheet.write(row, 7, pkg_pct, disc_hl)
            worksheet.write(row, 8, after_net, disc_hl)
            worksheet.write(row, 9, saved, disc_hl)
        else:
            worksheet.write(row, 7, 0, cell_center)
            worksheet.write(row, 8, b_net, cell_center)
            worksheet.write(row, 9, 0, cell_center)

        gross_total += b_gross; total_boxes += b_boxes
        net_total += b_net; after_disc_total += after_net; saved_total += saved
        row += 1

    worksheet.write(row, 1, "TOTAL", total)
    worksheet.write(row, 2, total_boxes, total)
    worksheet.write(row, 4, gross_total, total)
    worksheet.write(row, 6, net_total, total)
    worksheet.write(row, 7, "", total)
    worksheet.write(row, 8, after_disc_total, total)
    worksheet.write(row, 9, saved_total, total)

    row += 2
    worksheet.merge_range(row, 0, row, 6, "NET AMOUNT (After Package Discount)", header)
    worksheet.merge_range(row, 7, row, 9, f"Rs {after_disc_total:,.0f}", total)
    row += 1
    if pkg_pct > 0:
        worksheet.merge_range(row, 0, row, 6, "TOTAL SAVED BY DISCOUNT", header)
        worksheet.merge_range(row, 7, row, 9, f"Rs {saved_total:,.0f}", total)

    workbook.close(); output.seek(0)
    fname = f"{shop}_{date_str.replace('-','')}.xlsx".replace("/","-").replace(" ","_").replace(":","")
    st.session_state["download_file"] = (fname, output.getvalue())
    if pkg_pct > 0:
        st.session_state["success_msg"] = f"✅ Excel ready | {shop} | {pkg_pct}% discount applied"
    else:
        st.session_state["success_msg"] = f"✅ Excel ready | {shop}"

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    display_name = st.session_state.get("display_name", CURRENT_USER)
    st.markdown(f"""
    <div style='text-align:center; padding: 15px 0;'>
        <h2 style='color:#1976d2 !important; margin:0;'>🧾 AL-BARAKAH</h2>
        <p style='color:#0277bd !important; font-size:12px; margin:0; font-weight:600;'>ENTERPRISES</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown(f"""
    <div style='background:#ffffff;border:1px solid #90caf9;border-radius:8px;padding:8px 12px;margin-bottom:8px;text-align:center;'>
        <div style='font-size:11px;color:#0277bd;'>Logged in as</div>
        <div style='font-size:15px;font-weight:800;color:#1976d2;'>👤 {display_name}</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    page = st.radio(
        "MENU",
        [
            "📊 Dashboard", "🧾 Billing", "🛒 All Products",
            "🎁 Discount",
            "👤 Bookers", "💰 Bookers Salary",
            "🧑‍💼 Salesmen", "💰 Salesmen Salary",
            "💵 Daily Expense",
            "📋 Bills List", "📦 Load Form",
        ],
        key="page_selector",
        label_visibility="collapsed"
    )
    st.session_state["page"] = page

    st.markdown("---")
    active_pkgs = get_all_active_packages()
    st.markdown(f"""
    <div style='padding:6px 10px; color:#0277bd !important; font-size:12px;'>
        <p>📅 {datetime.now().strftime('%d-%m-%Y')}</p>
        <p>🎁 Active Packages: <b>{len(active_pkgs)}</b></p>
        <p>📦 Products: {len(PRODUCTS)}</p>
        <p>👤 Bookers: {len(db.get('bookers', []))}</p>
        <p>🧑‍💼 Salesmen: {len(db.get('salesmen', []))}</p>
        <p>🧾 Total Bills: {len(db['bills'])}</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    if st.button("🚪 Logout", key="btn_logout", use_container_width=True):
        st.session_state["logged_in_user"] = None
        st.session_state["display_name"] = None
        st.session_state["database"] = None
        st.session_state["_db_user"] = None
        st.session_state["light_on"] = False
        st.session_state["page"] = "📊 Dashboard"
        st.rerun()

# ============================================================
# PAGE: DASHBOARD
# ============================================================
def render_dashboard():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📊 Dashboard</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#0277bd;font-weight:500;'>Welcome to {COMPANY_NAME}</p>", unsafe_allow_html=True)
    st.markdown("---")

    bookers = db.get("bookers", [])
    salesmen = db.get("salesmen", [])

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"<div class='metric-card'><h3>TOTAL BOOKERS</h3><h1>{len(bookers)}</h1></div>", unsafe_allow_html=True)
    with c2:
        st.markdown(f"<div class='metric-card'><h3>TOTAL SALESMEN</h3><h1>{len(salesmen)}</h1></div>", unsafe_allow_html=True)
    with c3:
        st.markdown(f"<div class='metric-card'><h3>TOTAL PRODUCTS</h3><h1>{len(PRODUCTS)}</h1></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### 👤 Bookers")
        if not bookers:
            st.markdown("<div class='empty-box'>Abhi tak koi booker add nahi hua.</div>", unsafe_allow_html=True)
        else:
            for i, b_name in enumerate(bookers):
                sd = db.get("bookers_salaries", {}).get(b_name, {})
                base = sd.get("base_salary", 0)
                txns = sd.get("transactions", [])
                adv_p = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status", "pending") == "pending")
                short_p = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status", "pending") == "pending")
                remaining = base - adv_p - short_p
                st.markdown(f"""
                <div class='person-card'>
                    <div class='info'>
                        <div class='name'>👤 {b_name}</div>
                        <div class='sub'>Remaining Salary: Rs {remaining:,.0f}</div>
                    </div>
                    <div class='badge'>#{i+1}</div>
                </div>
                """, unsafe_allow_html=True)

    with c2:
        st.markdown("### 🧑‍💼 Salesmen")
        if not salesmen:
            st.markdown("<div class='empty-box'>Abhi tak koi salesman add nahi hua.</div>", unsafe_allow_html=True)
        else:
            for i, s_name in enumerate(salesmen):
                sd = db.get("salesmen_salaries", {}).get(s_name, {})
                base = sd.get("base_salary", 0)
                txns = sd.get("transactions", [])
                adv_p = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status", "pending") == "pending")
                short_p = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status", "pending") == "pending")
                remaining = base - adv_p - short_p
                st.markdown(f"""
                <div class='person-card'>
                    <div class='info'>
                        <div class='name'>🧑‍💼 {s_name}</div>
                        <div class='sub'>Remaining Salary: Rs {remaining:,.0f}</div>
                    </div>
                    <div class='badge sal-badge'>#{i+1}</div>
                </div>
                """, unsafe_allow_html=True)

# ============================================================
# PAGE: DISCOUNT
# ============================================================
def render_discount():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🎁 Discount Packages</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#0277bd;font-weight:500;'>Chaaron packages ek saath active ho sakte hain. Bill export pe sabse zyada % wala apply hoga.</p>", unsafe_allow_html=True)
    st.markdown("---")

    pkgs = db.get("discount_packages", [])
    if not pkgs:
        pkgs = default_discount_packages()
        db["discount_packages"] = pkgs
        save_database(db)

    for p in pkgs:
        if "tier3_amount" not in p: p["tier3_amount"] = 0.0
        if "tier3_pct" not in p: p["tier3_pct"] = 0.0

    active_names = [p.get("name", "Package") for p in pkgs if p.get("active")]
    active_str = ", ".join(active_names) if active_names else "Koi nahi"

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:15px;'>🎁 Active Packages:</b>
        <span style='color:#2e7d32;font-size:15px;font-weight:700;'>{active_str}</span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📝 Packages")

    for i, pkg in enumerate(pkgs):
        pid = pkg.get("id", i+1)
        is_active = pkg.get("active", False)
        card_class = "disc-card active" if is_active else "disc-card"
        badge = '<span class="disc-badge-active">ACTIVE</span>' if is_active else ""

        st.markdown(f"<div class='{card_class}'>", unsafe_allow_html=True)

        c1, c2 = st.columns([1, 5])
        with c1:
            new_active = st.checkbox("Active", value=is_active, key=f"active_{pid}")
            if new_active != is_active:
                for pp in db["discount_packages"]:
                    if pp.get("id") == pid:
                        pp["active"] = new_active
                        break
                save_database(db)
                st.rerun()
        with c2:
            new_name = st.text_input(
                f"Package {pid} Name",
                value=pkg.get("name", f"Package {pid}"),
                key=f"pkgname_{pid}",
                label_visibility="collapsed"
            )

        st.markdown(
            f"<div style='font-size:13px;color:#{'2e7d32' if is_active else '1976d2'};font-weight:800;margin:2px 0 4px 0;'>"
            f"🎁 {pkg.get('name','Package')} {badge}</div>",
            unsafe_allow_html=True
        )

        t1, t2, t3 = st.columns(3)
        with t1:
            st.markdown("<div style='font-size:11px;font-weight:700;color:#1976d2;'>Tier 1</div>", unsafe_allow_html=True)
            tc1, tc2 = st.columns(2)
            with tc1:
                ta1 = st.number_input("Min", value=float(pkg.get("tier1_amount", 0) or 0),
                                      min_value=0.0, step=50.0, key=f"t1a_{pid}", label_visibility="collapsed")
                st.caption("Min Rs")
            with tc2:
                tp1 = st.number_input("%", value=float(pkg.get("tier1_pct", 0) or 0),
                                      min_value=0.0, max_value=100.0, step=0.5, key=f"t1p_{pid}", label_visibility="collapsed")
                st.caption("Disc %")
        with t2:
            st.markdown("<div style='font-size:11px;font-weight:700;color:#1976d2;'>Tier 2</div>", unsafe_allow_html=True)
            tc1, tc2 = st.columns(2)
            with tc1:
                ta2 = st.number_input("Min", value=float(pkg.get("tier2_amount", 0) or 0),
                                      min_value=0.0, step=50.0, key=f"t2a_{pid}", label_visibility="collapsed")
                st.caption("Min Rs")
            with tc2:
                tp2 = st.number_input("%", value=float(pkg.get("tier2_pct", 0) or 0),
                                      min_value=0.0, max_value=100.0, step=0.5, key=f"t2p_{pid}", label_visibility="collapsed")
                st.caption("Disc %")
        with t3:
            st.markdown("<div style='font-size:11px;font-weight:700;color:#1976d2;'>Tier 3</div>", unsafe_allow_html=True)
            tc1, tc2 = st.columns(2)
            with tc1:
                ta3 = st.number_input("Min", value=float(pkg.get("tier3_amount", 0) or 0),
                                      min_value=0.0, step=50.0, key=f"t3a_{pid}", label_visibility="collapsed")
                st.caption("Min Rs")
            with tc2:
                tp3 = st.number_input("%", value=float(pkg.get("tier3_pct", 0) or 0),
                                      min_value=0.0, max_value=100.0, step=0.5, key=f"t3p_{pid}", label_visibility="collapsed")
                st.caption("Disc %")

        sc1, sc2 = st.columns([1, 5])
        with sc1:
            if st.button("💾 Save", key=f"savepkg_{pid}", use_container_width=True, type="primary"):
                for pp in db["discount_packages"]:
                    if pp.get("id") == pid:
                        pp["name"] = new_name.strip() or f"Package {pid}"
                        pp["tier1_amount"] = float(ta1); pp["tier1_pct"] = float(tp1)
                        pp["tier2_amount"] = float(ta2); pp["tier2_pct"] = float(tp2)
                        pp["tier3_amount"] = float(ta3); pp["tier3_pct"] = float(tp3)
                        break
                save_database(db)
                st.session_state["success_msg"] = "✅ Package saved"
                st.rerun()

        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("<hr style='margin:8px 0;'>", unsafe_allow_html=True)

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None

# ============================================================
# PAGE: ALL PRODUCTS
# ============================================================
def render_all_products():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🛒 All Products</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#0277bd;font-weight:500;'>Kisi bhi product ka price change karo</p>", unsafe_allow_html=True)
    st.markdown("---")

    c1, c2 = st.columns([3, 1])
    with c1:
        search = st.text_input("🔍 Search Product:", key="prod_search", placeholder="Type product name...")
    with c2:
        show_only_edited = st.checkbox("Sirf edited prices", key="prod_only_edited")

    search_upper = search.strip().upper()
    edited_prices = db.get("product_prices", {})

    shown = []
    for p in PRODUCTS:
        if search_upper and search_upper not in p["name"].upper(): continue
        if show_only_edited and str(p["code"]) not in edited_prices: continue
        shown.append(p)

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#0277bd;'>
            Total Products: <b>{len(PRODUCTS)}</b> &nbsp;|&nbsp;
            Showing: <b>{len(shown)}</b> &nbsp;|&nbsp;
            Edited Prices: <b>{len(edited_prices)}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    if edited_prices:
        c1, c2 = st.columns([1, 3])
        with c1:
            if st.button("🔄 Reset All Prices", key="reset_all_prices", use_container_width=True):
                db["product_prices"] = {}
                save_database(db)
                st.session_state["success_msg"] = "✅ Sab prices original pe reset ho gaye"
                st.rerun()

    st.markdown("---")
    if not shown:
        st.info("Is filter ke hisaab se koi product nahi mila.")
        return

    st.markdown(f"### 📋 Products ({len(shown)})")
    hc1, hc2, hc3, hc4 = st.columns([1, 4, 2, 2])
    with hc1: st.markdown("**Code**")
    with hc2: st.markdown("**Product**")
    with hc3: st.markdown("**Current Price**")
    with hc4: st.markdown("**Change / Reset**")
    st.markdown("<hr style='margin:6px 0;'>", unsafe_allow_html=True)

    for p in shown:
        code = str(p["code"])
        base = float(p["price"])
        current = get_price(code, base)
        is_edited = code in edited_prices

        c1, c2, c3, c4 = st.columns([1, 4, 2, 2])
        with c1:
            st.markdown(f"<div style='padding-top:8px;color:#1976d2;font-weight:700;'>{code}</div>", unsafe_allow_html=True)
        with c2:
            edited_mark = " ✏️" if is_edited else ""
            color = "#c62828" if is_edited else "#1976d2"
            st.markdown(f"<div style='padding-top:6px;color:{color};font-weight:600;font-size:14px;'>{p['name']}{edited_mark}</div>", unsafe_allow_html=True)
        with c3:
            new_price = st.number_input("Price", value=float(current), min_value=0.0, step=1.0,
                                        key=f"price_{code}", label_visibility="collapsed")
        with c4:
            sub1, sub2 = st.columns(2)
            with sub1:
                if st.button("💾 Save", key=f"save_price_{code}", use_container_width=True):
                    if float(new_price) == base:
                        db.get("product_prices", {}).pop(code, None)
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ {p['name']}: price original (Rs {base:,.0f})"
                    else:
                        db.setdefault("product_prices", {})[code] = float(new_price)
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ {p['name']}: naya price Rs {new_price:,.0f}"
                    st.rerun()
            with sub2:
                if is_edited:
                    if st.button("↩️ Reset", key=f"reset_price_{code}", use_container_width=True):
                        db.get("product_prices", {}).pop(code, None)
                        save_database(db)
                        st.session_state["success_msg"] = f"↩️ {p['name']}: original price (Rs {base:,.0f})"
                        st.rerun()

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"])
        st.session_state["success_msg"] = None

# ============================================================
# PAGE: BOOKERS
# ============================================================
def render_bookers():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>👤 Bookers</h1>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("### ➕ Add New Booker")
    c1, c2 = st.columns([3, 1])
    with c1:
        new_booker = st.text_input("Booker Name:", key="new_booker_name", placeholder="Enter booker name...", label_visibility="collapsed")
    with c2:
        add_clicked = st.button("➕ Add Booker", key="btn_add_booker", use_container_width=True, type="primary")

    if add_clicked:
        name = new_booker.strip() if new_booker else ""
        if name == "":
            st.error("❌ Please enter a booker name")
        elif name in db.get("bookers", []):
            st.warning(f"⚠️ '{name}' already exists")
        else:
            db.setdefault("bookers", []).append(name)
            db["bookers"] = sorted(db["bookers"])
            save_database(db)
            st.session_state["booker_msg"] = f"✅ '{name}' added"
            st.rerun()

    if st.session_state.get("booker_msg"):
        st.success(st.session_state["booker_msg"]); st.session_state["booker_msg"] = None

    st.markdown("---")
    st.markdown(f"### 📋 Saved Bookers ({len(db.get('bookers', []))})")

    bookers = db.get("bookers", [])
    if not bookers:
        st.info("Abhi tak koi booker add nahi hua.")
        return

    for i, b in enumerate(bookers):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#1976d2;'>👤 {b}</b></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑 Delete", key=f"del_booker_{i}_{b}", use_container_width=True):
                db["bookers"].remove(b)
                save_database(db)
                st.session_state["booker_msg"] = f"🗑 '{b}' deleted"
                st.rerun()

# ============================================================
# PAGE: SALESMEN
# ============================================================
def render_salesmen():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🧑‍💼 Salesmen</h1>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("### ➕ Add New Salesman")
    c1, c2 = st.columns([3, 1])
    with c1:
        new_s = st.text_input("Salesman Name:", key="new_salesman_name", placeholder="Enter salesman name...", label_visibility="collapsed")
    with c2:
        add_clicked = st.button("➕ Add Salesman", key="btn_add_salesman", use_container_width=True, type="primary")

    if add_clicked:
        name = new_s.strip() if new_s else ""
        if name == "":
            st.error("❌ Please enter a salesman name")
        elif name in db.get("salesmen", []):
            st.warning(f"⚠️ '{name}' already exists")
        else:
            db.setdefault("salesmen", []).append(name)
            db["salesmen"] = sorted(db["salesmen"])
            save_database(db)
            st.session_state["salesman_msg"] = f"✅ '{name}' added"
            st.rerun()

    if st.session_state.get("salesman_msg"):
        st.success(st.session_state["salesman_msg"]); st.session_state["salesman_msg"] = None

    st.markdown("---")
    st.markdown(f"### 📋 Saved Salesmen ({len(db.get('salesmen', []))})")

    salesmen = db.get("salesmen", [])
    if not salesmen:
        st.info("Abhi tak koi salesman add nahi hua.")
        return

    for i, s in enumerate(salesmen):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#1976d2;'>🧑‍💼 {s}</b></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑 Delete", key=f"del_salesman_{i}_{s}", use_container_width=True):
                db["salesmen"].remove(s)
                save_database(db)
                st.session_state["salesman_msg"] = f"🗑 '{s}' deleted"
                st.rerun()

# ============================================================
# PAGE: SALARY
# ============================================================
def render_salaries(role_type):
    if role_type == "bookers":
        title = "💰 Bookers Salary"; emoji = "👤"; names = db.get("bookers", [])
        sal_key = "bookers_salaries"; other_page = "👤 Bookers"
    else:
        title = "💰 Salesmen Salary"; emoji = "🧑‍💼"; names = db.get("salesmen", [])
        sal_key = "salesmen_salaries"; other_page = "🧑‍💼 Salesmen"

    st.markdown(f"<h1 style='color:#1976d2 !important;'>{title}</h1>", unsafe_allow_html=True)
    st.markdown("---")

    if sal_key not in db: db[sal_key] = {}
    if not names:
        st.info(f"❌ Abhi tak koi {role_type[:-1]} add nahi hua. Pehle **{other_page}** page pe add karo.")
        return

    total_base = 0; total_adv_pending = 0; total_adv_paid = 0
    total_short_pending = 0; total_short_paid = 0
    for n in names:
        sd = db[sal_key].get(n, {})
        base = sd.get("base_salary", 0)
        txns = sd.get("transactions", [])
        total_base += base
        for t in txns:
            is_pending = t.get("status", "pending") == "pending"
            if t.get("type") == "advanced":
                if is_pending: total_adv_pending += t["amount"]
                else: total_adv_paid += t["amount"]
            elif t.get("type") == "shortage":
                if is_pending: total_short_pending += t["amount"]
                else: total_short_paid += t["amount"]
    total_remaining = total_base - total_adv_pending - total_short_pending

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Overall Summary</b><br>
        <span style='color:#0277bd;'>
            Base: <b>Rs {total_base:,.0f}</b> | Adv Pend: <b>Rs {total_adv_pending:,.0f}</b> |
            Adv Paid: <b>Rs {total_adv_paid:,.0f}</b> | Short Pend: <b>Rs {total_short_pending:,.0f}</b> |
            Short Paid: <b>Rs {total_short_paid:,.0f}</b> |
            <b style='color:#1b5e20;'>Remaining: Rs {total_remaining:,.0f}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    for person_name in names:
        sd = db[sal_key].get(person_name, {"base_salary": 0, "transactions": []})
        base = sd.get("base_salary", 0)
        txns = sd.get("transactions", [])

        adv_pending = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status", "pending") == "pending")
        adv_paid = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status") == "paid")
        short_pending = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status", "pending") == "pending")
        short_paid = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status") == "paid")
        remaining = base - adv_pending - short_pending

        with st.expander(f"💰 {emoji} {person_name}  —  Remaining: Rs {remaining:,.0f}", expanded=False):
            c1, c2 = st.columns([3, 1])
            with c1:
                new_base = st.number_input("Base Salary (Rs):", value=float(base), min_value=0.0, step=500.0,
                                           key=f"base_{role_type}_{person_name}")
            with c2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Save Base", key=f"savebase_{role_type}_{person_name}", use_container_width=True):
                    db[sal_key].setdefault(person_name, {"base_salary": 0, "transactions": []})
                    db[sal_key][person_name]["base_salary"] = float(new_base)
                    save_database(db)
                    st.session_state["success_msg"] = f"✅ Base saved for {person_name}"
                    st.rerun()

            st.markdown(f"""
            <div style='margin-top:10px;'>
                <span class='sal-metric base'>Base: Rs {base:,.0f}</span>
                <span class='sal-metric adv'>Adv Pend: Rs {adv_pending:,.0f}</span>
                <span class='sal-metric paid'>Adv Paid: Rs {adv_paid:,.0f}</span>
                <span class='sal-metric short'>Short Pend: Rs {short_pending:,.0f}</span>
                <span class='sal-metric paid'>Short Paid: Rs {short_paid:,.0f}</span>
                <span class='sal-metric remain'>Remaining: Rs {remaining:,.0f}</span>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("**➕ Add Transaction**")

            c1, c2, c3, c4, c5 = st.columns([2, 2, 2, 3, 1])
            with c1: txn_date = st.date_input("Date", value=date.today(), key=f"txn_date_{role_type}_{person_name}")
            with c2: txn_type = st.selectbox("Type", ["Advanced", "Shortage"], key=f"txn_type_{role_type}_{person_name}")
            with c3: txn_amt = st.number_input("Amount", min_value=0.0, step=100.0, key=f"txn_amt_{role_type}_{person_name}")
            with c4: txn_note = st.text_input("Note", key=f"txn_note_{role_type}_{person_name}", placeholder="Reason...")
            with c5:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Add", key=f"addtxn_{role_type}_{person_name}", use_container_width=True):
                    if txn_amt <= 0:
                        st.session_state["error_msg"] = "❌ Amount > 0"
                    else:
                        db[sal_key].setdefault(person_name, {"base_salary": base, "transactions": []})
                        existing = db[sal_key][person_name].get("transactions", [])
                        next_id = max([t.get("id", 0) for t in existing] + [0]) + 1
                        db[sal_key][person_name]["transactions"].append({
                            "id": next_id, "date": txn_date.strftime("%d-%m-%Y"),
                            "time": datetime.now().strftime("%H:%M"),
                            "type": "advanced" if txn_type == "Advanced" else "shortage",
                            "amount": float(txn_amt), "note": txn_note.strip(),
                            "status": "pending",
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        })
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ {txn_type} Rs {txn_amt:,.0f} added"
                        st.rerun()

            if txns:
                st.markdown("**📋 History**")
                for idx, t in enumerate(sorted(txns, key=lambda x: x.get("created_at", ""), reverse=True)):
                    t_id = t.get("id")
                    is_pending = t.get("status", "pending") == "pending"
                    is_adv = t["type"] == "advanced"
                    if is_adv:
                        bt, bc, bg = "Advanced", "#e65100", "#fff3e0"
                    else:
                        bt, bc, bg = "Shortage", "#c62828", "#ffebee"
                    sb = '<span class="status-pending">⏳ PENDING</span>' if is_pending else '<span class="status-paid">✅ PAID</span>'
                    nt = f" — {t.get('note','')}" if t.get("note") else ""

                    c1, c2, c3 = st.columns([5, 1, 1])
                    with c1:
                        st.markdown(f"""
                        <div style='background:#ffffff;border:1px solid #e0e0e0;border-radius:8px;padding:8px 12px;margin-bottom:4px;'>
                            <span style='color:#0277bd;font-size:12px;'>📅 {t.get('date','')} · 🕐 {t.get('time','')}</span> &nbsp;
                            <span style='background:{bg};color:{bc};padding:2px 8px;border-radius:5px;font-size:11px;font-weight:700;'>{bt}</span> &nbsp;
                            {sb} &nbsp;
                            <b style='color:#1976d2;font-size:15px;'>Rs {t['amount']:,.0f}</b>
                            <span style='color:#666;font-size:12px;'>{nt}</span>
                        </div>
                        """, unsafe_allow_html=True)
                    with c2:
                        if is_pending:
                            if st.button("✅ Paid", key=f"paid_{role_type}_{person_name}_{t_id}_{idx}", use_container_width=True, type="primary"):
                                for tx in db[sal_key][person_name]["transactions"]:
                                    if tx.get("id") == t_id:
                                        tx["status"] = "paid"
                                        tx["paid_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        break
                                save_database(db)
                                st.session_state["success_msg"] = f"✅ Rs {t['amount']:,.0f} PAID"
                                st.rerun()
                        else:
                            st.markdown("<div style='padding-top:6px;color:#1b5e20;font-weight:700;font-size:13px;text-align:center;'>✅ PAID</div>", unsafe_allow_html=True)
                    with c3:
                        if st.button("🗑", key=f"deltxn_{role_type}_{person_name}_{t_id}_{idx}", use_container_width=True):
                            db[sal_key][person_name]["transactions"] = [x for x in db[sal_key][person_name]["transactions"] if x.get("id") != t_id]
                            save_database(db)
                            st.session_state["success_msg"] = "🗑 Deleted"
                            st.rerun()
            else:
                st.info("Koi transaction nahi.")

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None

# ============================================================
# PAGE: DAILY EXPENSE
# ============================================================
def render_daily_expense():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>💵 Daily Expense</h1>", unsafe_allow_html=True)
    st.markdown("---")

    if "petrol_expenses" not in db: db["petrol_expenses"] = []
    if "lunch_expenses" not in db: db["lunch_expenses"] = []

    petrol_list = db.get("petrol_expenses", [])
    lunch_list = db.get("lunch_expenses", [])

    st.markdown("### 🔎 Filter")
    today = date.today()
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        filter_mode = st.selectbox("Filter:", ["📅 Aaj", "📆 This Month", "🗓️ Last 30 Days", "📋 All"], key="exp_filter_mode")
    with c2: from_date = st.date_input("From:", value=today - timedelta(days=30), key="exp_from_date")
    with c3: to_date = st.date_input("To:", value=today, key="exp_to_date")

    def date_match(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if filter_mode == "📅 Aaj": return d == today
        elif filter_mode == "📆 This Month": return d.year == today.year and d.month == today.month
        elif filter_mode == "🗓️ Last 30 Days": return today - timedelta(days=30) <= d <= today
        return True

    pf = [x for x in petrol_list if date_match(x.get("date", ""))]
    lf = [x for x in lunch_list if date_match(x.get("date", ""))]
    tp = sum(float(x.get("amount", 0)) for x in pf)
    tl = sum(float(x.get("amount", 0)) for x in lf)

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#0277bd;'>
            ⛽ Petrol: <b>Rs {tp:,.0f}</b> &nbsp;|&nbsp;
            🍽️ Lunch: <b>Rs {tl:,.0f}</b> &nbsp;|&nbsp;
            <b style='color:#0d47a1;'>Grand Total: Rs {tp+tl:,.0f}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown(f"### ⛽ Petrol  <span style='font-size:14px;color:#0277bd;'>(Rs {tp:,.0f})</span>", unsafe_allow_html=True)
        ss = db.get("salesmen", [])
        if not ss:
            st.info("💡 Pehle Salesmen add karo.")
        else:
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1: p_s = st.selectbox("Salesman:", options=["-- Select --"] + ss, key="petrol_salesman", label_visibility="collapsed")
            with c2: p_a = st.number_input("Amount", min_value=0.0, step=50.0, key="petrol_amount", label_visibility="collapsed", placeholder="Amount")
            with c3:
                if st.button("➕", key="add_petrol", use_container_width=True, type="primary"):
                    if p_s == "-- Select --": st.session_state["error_msg"] = "❌ Select salesman"
                    elif p_a <= 0: st.session_state["error_msg"] = "❌ Enter amount"
                    else:
                        nid = max([x.get("id", 0) for x in db["petrol_expenses"]] + [0]) + 1
                        db["petrol_expenses"].append({
                            "id": nid, "date": datetime.now().strftime("%d-%m-%Y"),
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "time": datetime.now().strftime("%H:%M"),
                            "salesman": p_s, "amount": float(p_a),
                        })
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ Petrol Rs {p_a:,.0f} ({p_s})"
                        st.rerun()

            if pf:
                for idx, x in enumerate(sorted(pf, key=lambda x: x.get("created_at", ""), reverse=True)):
                    xid = x.get("id", idx)
                    cc1, cc2 = st.columns([5, 1])
                    with cc1:
                        st.markdown(f"""
                        <div class='exp-row'>
                            <div class='exp-left'>
                                <div class='exp-name'>🧑‍💼 {x.get('salesman','-')}</div>
                                <div class='exp-date'>📅 {x.get('date','')} · 🕐 {x.get('time','')}</div>
                            </div>
                            <div class='exp-amt'>Rs {float(x.get('amount',0)):,.0f}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with cc2:
                        if st.button("🗑", key=f"del_petrol_{xid}_{idx}", use_container_width=True):
                            db["petrol_expenses"] = [e for e in db["petrol_expenses"] if e.get("id") != xid]
                            save_database(db)
                            st.session_state["success_msg"] = "🗑 Deleted"
                            st.rerun()

    with col_right:
        st.markdown(f"### 🍽️ Lunch  <span style='font-size:14px;color:#e65100;'>(Rs {tl:,.0f})</span>", unsafe_allow_html=True)
        c1, c2 = st.columns([3, 1])
        with c1: l_a = st.number_input("Amount", min_value=0.0, step=50.0, key="lunch_amount", placeholder="Amount", label_visibility="collapsed")
        with c2:
            if st.button("➕", key="add_lunch", use_container_width=True, type="primary"):
                if l_a <= 0: st.session_state["error_msg"] = "❌ Enter amount"
                else:
                    nid = max([x.get("id", 0) for x in db["lunch_expenses"]] + [0]) + 1
                    db["lunch_expenses"].append({
                        "id": nid, "date": datetime.now().strftime("%d-%m-%Y"),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "time": datetime.now().strftime("%H:%M"),
                        "amount": float(l_a),
                    })
                    save_database(db)
                    st.session_state["success_msg"] = f"✅ Lunch Rs {l_a:,.0f}"
                    st.rerun()

        if lf:
            for idx, x in enumerate(sorted(lf, key=lambda x: x.get("created_at", ""), reverse=True)):
                xid = x.get("id", idx)
                cc1, cc2 = st.columns([5, 1])
                with cc1:
                    st.markdown(f"""
                    <div class='exp-row lunch'>
                        <div class='exp-left'>
                            <div class='exp-name'>🍽️ Lunch</div>
                            <div class='exp-date'>📅 {x.get('date','')} · 🕐 {x.get('time','')}</div>
                        </div>
                        <div class='exp-amt'>Rs {float(x.get('amount',0)):,.0f}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with cc2:
                    if st.button("🗑", key=f"del_lunch_{xid}_{idx}", use_container_width=True):
                        db["lunch_expenses"] = [e for e in db["lunch_expenses"] if e.get("id") != xid]
                        save_database(db)
                        st.session_state["success_msg"] = "🗑 Deleted"
                        st.rerun()

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
    if st.session_state.get("error_msg"):
        st.error(st.session_state["error_msg"]); st.session_state["error_msg"] = None

# ============================================================
# PAGE: BILLING
# ============================================================
def render_billing():
    st.markdown(f"<h2 style='color:#1976d2 !important;margin:0 0 6px 0;'>🧾 Billing</h2>", unsafe_allow_html=True)

    active_pkgs = get_all_active_packages()
    if active_pkgs:
        names = ", ".join([p.get("name", "Package") for p in active_pkgs])
        st.markdown(f"<div class='hint-box'>🎁 Active Packages ({len(active_pkgs)}): <b>{names}</b></div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='hint-box'>💡 Koi discount package active nahi.</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1: st.text_input("Bill No:", value=str(db["next_bill_no"]), disabled=True, key="dash_bill_no")
    with c2: st.text_input("Date:", value=datetime.now().strftime("%d-%m-%Y"), disabled=True, key="dash_bill_date")

    c1, c2, c3 = st.columns(3)
    with c1: shop_name = st.text_input("Shop:", key="shop_name", placeholder="Shop Name")
    with c2:
        sb = db.get("bookers", [])
        if sb:
            opts = ["-- Select --"] + sb
            sel = st.selectbox("Order Booker:", options=opts, key="order_booker_select")
            st.session_state["order_booker"] = "" if sel == "-- Select --" else sel
        else:
            st.text_input("Order Booker:", key="order_booker", placeholder="Booker")
    with c3:
        ss = db.get("salesmen", [])
        if ss:
            opts = ["-- Select --"] + ss
            sel = st.selectbox("Salesman:", options=opts, key="salesman_select")
            st.session_state["salesman"] = "" if sel == "-- Select --" else sel
        else:
            st.text_input("Salesman:", key="salesman", placeholder="Salesman")

    c1, c2 = st.columns([1, 3])
    with c1: search_text = st.text_input("🔍 Search:", key="search_text", placeholder="Type name...")
    with c2:
        su = search_text.strip().upper()
        fn = [p["name"] for p in PRODUCTS if su in p["name"].upper()] if su else PRODUCT_NAMES
        if st.session_state.get("product_sel") and st.session_state["product_sel"] not in fn:
            st.session_state["product_sel"] = ""
        product_sel = st.selectbox("Select Product:", options=[""] + fn, key="product_sel")

    sp = None
    if product_sel:
        for p in PRODUCTS:
            if p["name"] == product_sel: sp = p; break

    if sp:
        cp = get_price(sp["code"], sp["price"])
        bp = float(sp["price"])
        pn = f"  (edited — orig Rs {bp:,.0f})" if cp != bp else ""
        st.success(f"✅ {sp['name']} (Code: {sp['code']}) — Rs {cp:,.0f}{pn}")
        tpd = float(cp)
    else:
        st.error("No Product Selected"); tpd = 0.0

    if st.session_state["_prev_prod"] != product_sel:
        st.session_state["tp_box"] = tpd
        st.session_state["_prev_prod"] = product_sel

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: boxes = st.number_input("Boxes:", min_value=0, step=1, key="boxes")
    with c2: tp_box = st.number_input("TP/Box:", min_value=0.0, step=1.0, key="tp_box")
    with c3: discount = st.number_input("Disc %:", min_value=0.0, step=0.5, key="discount")
    gross = boxes * tp_box
    net = gross - (gross * discount / 100)
    with c4: st.text_input("Gross:", value=f"{gross:.0f}", disabled=True, key="gross_disp")
    with c5: st.text_input("Net:", value=f"{net:.0f}", disabled=True, key="net_disp")

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
    if st.session_state.get("error_msg"):
        st.error(st.session_state["error_msg"]); st.session_state["error_msg"] = None

    b1, b2 = st.columns(2)
    with b1: st.button("➕ Add Bill", key="btn_add", on_click=add_bill_callback, use_container_width=True, type="primary")
    with b2: st.button("🔄 Refresh", key="btn_refresh", on_click=refresh_callback, use_container_width=True)

    e1, e2, e3 = st.columns(3)
    with e1: st.button("📄 Export Bill", key="btn_export", on_click=export_bill_callback, use_container_width=True)
    with e2: st.button("📦 Export Load Form", key="btn_export_lf", on_click=export_load_form_from_billing_callback, use_container_width=True)
    with e3: st.button("🗑 Refresh Load Form", key="btn_load_refresh", on_click=refresh_load_form_callback, use_container_width=True)

    show_auto_download()

# ============================================================
# PAGE: BILLS LIST
# ============================================================
def render_bills_list():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📋 Bills List</h1>", unsafe_allow_html=True)
    st.markdown("---")

    if len(db["bills"]) == 0:
        st.info("❌ Koi bill nahi mila.")
        return

    st.markdown("### 🔎 Filter")
    today = date.today()
    c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Mode:", ["📅 Aaj", "📆 Custom Range", "🗓️ Specific Date", "📋 All"], key="bills_filter_mode")
    with c2: from_date = st.date_input("From:", value=today - timedelta(days=7), key="bills_from_date")
    with c3: to_date = st.date_input("To:", value=today, key="bills_to_date")

    c1, c2 = st.columns(2)
    with c1: search = st.text_input("🔍 Search:", key="bills_search")
    with c2:
        sl = sorted(set(b["Shop"] for b in db["bills"] if b["Shop"]))
        shop_filter = st.selectbox("Shop:", options=["All"] + sl, key="shop_filter")

    groups = {}
    for oi, b in enumerate(db["bills"]):
        bd = parse_date(b.get("Date", ""))
        if filter_mode == "📅 Aaj":
            if bd != today: continue
        elif filter_mode == "🗓️ Specific Date":
            if bd != from_date: continue
        elif filter_mode == "📆 Custom Range":
            if bd is None or not (from_date <= bd <= to_date): continue

        if search:
            s = search.upper()
            if not (s in str(b.get("Shop","")).upper() or s in str(b.get("Product","")).upper() or s in str(b.get("Order Booker","")).upper()):
                continue
        if shop_filter != "All" and b.get("Shop", "") != shop_filter: continue

        key = (b.get("Shop",""), b.get("Date",""), b.get("Order Booker",""))
        if key not in groups:
            groups[key] = {"shop": b.get("Shop",""), "date": b.get("Date",""),
                          "booker": b.get("Order Booker",""), "salesman": b.get("Salesman",""),
                          "items": [], "orig_indices": [], "bill_no": b.get("Bill No","")}
        groups[key]["items"].append({
            "Code": b.get("Code"), "Product": b.get("Product"),
            "Boxes": b.get("Boxes"), "TP/Box": b.get("TP/Box"),
            "Discount %": b.get("Discount %"), "Gross": b.get("Gross"), "Net": b.get("Net"),
        })
        groups[key]["orig_indices"].append(oi)

    if not groups:
        st.warning("❌ Koi bill nahi mila.")
        return

    all_items = []
    for g in groups.values(): all_items.extend(g["items"])
    tb = sum(int(r.get("Boxes",0)) for r in all_items)
    tg = sum(float(r.get("Gross",0)) for r in all_items)
    tn = sum(float(r.get("Net",0)) for r in all_items)
    us = len(set(g["shop"] for g in groups.values() if g["shop"]))

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#0277bd;'>
            Bills: <b>{len(groups)}</b> | Boxes: <b>{tb}</b> |
            Shops: <b>{us}</b> |
            Gross: <b>Rs {tg:,.0f}</b> | Net: <b>Rs {tn:,.0f}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    sk = sorted(groups.keys(), key=lambda k: (parse_date(k[1]) or date.min, k[0]), reverse=True)

    for idx, key in enumerate(sk):
        g = groups[key]
        shop = g["shop"] or "-"; date_str = g["date"] or "-"
        booker = g["booker"] or "-"; salesman = g["salesman"] or "-"
        bill_no = g["bill_no"]; items = g["items"]
        tb2 = sum(int(it.get("Boxes",0)) for it in items)
        tn2 = sum(float(it.get("Net",0)) for it in items)

        wkey = f"{shop}_{date_str}_{booker}_{bill_no}_{idx}".replace(" ","_").replace("/","_").replace(":","")
        iv = st.session_state.get("view_bill_key") == wkey

        c1, c2, c3, c4 = st.columns([4, 0.7, 1, 1])
        with c1:
            st.markdown(f"""
            <div class='lf-simple-card'>
                <div class='lf-info'>
                    <div class='lf-line1'>🏪 {shop}</div>
                    <div class='lf-line2'>📅 {date_str} &nbsp;·&nbsp; 👤 {booker} &nbsp;·&nbsp; 🧑‍💼 {salesman}</div>
                </div>
                <div class='lf-boxes'>{tb2}<small>BOXES</small></div>
            </div>
            """, unsafe_allow_html=True)
        with c2:
            if st.button("🔽" if iv else "👁️", key=f"eye_bill_{wkey}", use_container_width=True):
                st.session_state["view_bill_key"] = None if iv else wkey
                st.rerun()
        with c3:
            if st.button("⬇️ Excel", key=f"dl_bill_{wkey}", use_container_width=True, type="primary"):
                export_single_group_bill(shop, date_str, booker, salesman, items, bill_no)
                st.rerun()
        with c4:
            if st.button("🗑 Delete", key=f"del_bill_{wkey}", use_container_width=True):
                st.session_state["confirm_delete_group"] = g["orig_indices"]
                st.session_state["_confirm_group_label"] = f"{shop} | {date_str} | {booker}"

        if iv:
            st.markdown(f"""
            <div class='full-bill-box'>
                <div class='full-bill-title'>👁️ {shop} — Full Bill</div>
                <div class='full-bill-meta'>📅 {date_str} · 👤 {booker} · 🧑‍💼 {salesman} · 🧾 #{bill_no}</div>
            </div>
            """, unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(items), use_container_width=True, hide_index=True)
            pp, pn, _ = get_package_discount_pct(tn2)
            ad = tn2 - (tn2 * pp / 100)
            cc1, cc2, cc3 = st.columns(3)
            with cc1: st.markdown(f"<div class='metric-card'><h3>BOXES</h3><h1>{tb2}</h1></div>", unsafe_allow_html=True)
            with cc2: st.markdown(f"<div class='metric-card'><h3>NET</h3><h1>Rs {tn2:,.0f}</h1></div>", unsafe_allow_html=True)
            with cc3: st.markdown(f"<div class='metric-card'><h3>AFTER DISC{' ('+str(pp)+'%)' if pp>0 else ''}</h3><h1>Rs {ad:,.0f}</h1></div>", unsafe_allow_html=True)
            if pp > 0:
                st.markdown(f"<div class='hint-box'>🎁 {pn} | {pp}% | Saved: Rs {tn2-ad:,.0f}</div>", unsafe_allow_html=True)
            cl1, cl2 = st.columns([1, 4])
            with cl1:
                if st.button("❌ Close", key=f"close_{wkey}", use_container_width=True):
                    st.session_state["view_bill_key"] = None
                    st.rerun()
            with cl2:
                if st.button("⬇️ Excel", key=f"dl_from_view_{wkey}", use_container_width=True, type="primary"):
                    export_single_group_bill(shop, date_str, booker, salesman, items, bill_no)
                    st.rerun()
            st.markdown("---")

    if st.session_state.get("confirm_delete_group"):
        label = st.session_state.get("_confirm_group_label", "")
        st.warning(f"⚠️ Delete **{label}**?")
        cc1, cc2 = st.columns(2)
        with cc1:
            if st.button("✅ Yes Delete", key="cdy", use_container_width=True, type="primary"):
                idxs = set(st.session_state.get("confirm_delete_group", []))
                if idxs:
                    db["bills"] = [b for i, b in enumerate(db["bills"]) if i not in idxs]
                    save_database(db)
                    st.session_state["success_msg"] = "🗑 Deleted"
                st.session_state["confirm_delete_group"] = None
                st.session_state["_confirm_group_label"] = ""
                st.session_state["view_bill_key"] = None
                st.rerun()
        with cc2:
            if st.button("❌ Cancel", key="cdn", use_container_width=True):
                st.session_state["confirm_delete_group"] = None
                st.session_state["_confirm_group_label"] = ""
                st.rerun()

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
    show_auto_download()

# ============================================================
# PAGE: LOAD FORM
# ============================================================
def render_load_form():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📦 Load Forms</h1>", unsafe_allow_html=True)
    st.markdown("---")

    load_forms = db.get("load_forms", [])
    if not load_forms:
        st.info("❌ Koi load form nahi mila.")
        return

    st.markdown("### 🔎 Filter")
    today = date.today()
    c1, c2, c3 = st.columns(3)
    with c1: filter_mode = st.selectbox("Mode:", ["📅 Aaj", "📆 Custom Range", "🗓️ Specific Date", "📋 All"], key="lf_filter_mode")
    with c2: from_date = st.date_input("From:", value=today - timedelta(days=7), key="lf_from_date")
    with c3: to_date = st.date_input("To:", value=today, key="lf_to_date")

    bn = sorted(set(lf["booker"] for lf in load_forms if lf.get("booker")))
    bf = st.selectbox("Filter by Booker:", options=["All"] + bn, key="lf_booker_filter")

    fl = []
    for lf in load_forms:
        lfd = parse_date(lf.get("date", ""))
        if lfd is None: continue
        if filter_mode == "📅 Aaj":
            if lfd != today: continue
        elif filter_mode == "🗓️ Specific Date":
            if lfd != from_date: continue
        elif filter_mode == "📆 Custom Range":
            if not (from_date <= lfd <= to_date): continue
        if bf != "All" and lf.get("booker") != bf: continue
        fl.append(lf)

    if not fl:
        st.warning("❌ Koi load form nahi mila.")
        return

    tf = len(fl)
    tba = sum(lf.get("total_boxes", 0) for lf in fl)

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#0277bd;'>Load Forms: <b>{tf}</b> | Boxes: <b>{tba}</b></span>
    </div>
    """, unsafe_allow_html=True)

    if st.button("⬇️ Download All Filtered (Excel)", key="lf_dl_all", use_container_width=True):
        export_all_filtered_load_forms(fl); st.rerun()

    st.markdown("---")

    for idx, lf in enumerate(sorted(fl, key=lambda x: x.get("created_at", ""), reverse=True)):
        lfid = lf.get("id", idx)
        booker = lf.get("booker", "Unknown")
        date_str = lf.get("date", ""); time_str = lf.get("time", "")
        tb = lf.get("total_boxes", 0)
        items = lf.get("items", [])

        bi = []
        for it in items:
            code = it.get("Code", "")
            bp = 0.0
            for p in PRODUCTS:
                if str(p["code"]) == str(code): bp = get_price(p["code"], p["price"]); break
            bx = int(it.get("Boxes", 0))
            gr = bx * bp
            bi.append({"Code": code, "Product": it.get("Product",""), "Boxes": bx, "TP/Box": bp, "Discount %": 0, "Gross": gr, "Net": gr})

        wkey = f"lf_{lfid}_{idx}"
        iv = st.session_state.get("view_lf_key") == wkey

        c1, c2, c3, c4 = st.columns([4, 0.7, 1, 1])
        with c1:
            st.markdown(f"""
            <div class='lf-simple-card'>
                <div class='lf-info'>
                    <div class='lf-line1'>👤 {booker}</div>
                    <div class='lf-line2'>📅 {date_str} · 🕐 {time_str} · 📦 {len(items)} products</div>
                </div>
                <div class='lf-boxes'>{tb}<small>BOXES</small></div>
            </div>
            """, unsafe_allow_html=True)
        with c2:
            if st.button("🔽" if iv else "👁️", key=f"eye_lf_{wkey}", use_container_width=True):
                st.session_state["view_lf_key"] = None if iv else wkey
                st.rerun()
        with c3:
            if st.button("⬇️ Excel", key=f"dl_lf_{wkey}", use_container_width=True, type="primary"):
                export_single_group_bill(booker + " Load", date_str, booker, "", bi, lfid)
                st.rerun()
        with c4:
            if st.button("🗑 Delete", key=f"del_lf_{wkey}", use_container_width=True):
                db["load_forms"] = [x for x in db["load_forms"] if x.get("id") != lfid]
                save_database(db)
                st.session_state["success_msg"] = "🗑 Deleted"
                st.session_state["view_lf_key"] = None
                st.rerun()

        if iv:
            st.markdown(f"""
            <div class='full-bill-box'>
                <div class='full-bill-title'>👁️ Load Form — {booker}</div>
                <div class='full-bill-meta'>📅 {date_str} · 🕐 {time_str} · 📦 {len(items)} products · {tb} boxes</div>
            </div>
            """, unsafe_allow_html=True)
            if bi: st.dataframe(pd.DataFrame(bi), use_container_width=True, hide_index=True)
            tnet = sum(float(x.get("Net",0)) for x in bi)
            cc1, cc2, cc3 = st.columns(3)
            with cc1: st.markdown(f"<div class='metric-card'><h3>PRODUCTS</h3><h1>{len(items)}</h1></div>", unsafe_allow_html=True)
            with cc2: st.markdown(f"<div class='metric-card'><h3>BOXES</h3><h1>{tb}</h1></div>", unsafe_allow_html=True)
            with cc3: st.markdown(f"<div class='metric-card'><h3>NET</h3><h1>Rs {tnet:,.0f}</h1></div>", unsafe_allow_html=True)
            cl1, cl2 = st.columns([1, 4])
            with cl1:
                if st.button("❌ Close", key=f"close_lf_{wkey}", use_container_width=True):
                    st.session_state["view_lf_key"] = None
                    st.rerun()
            with cl2:
                if st.button("⬇️ Excel", key=f"dl_from_view_lf_{wkey}", use_container_width=True, type="primary"):
                    export_single_group_bill(booker + " Load", date_str, booker, "", bi, lfid)
                    st.rerun()
            st.markdown("---")

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None
    show_auto_download()

# ============================================================
# CALLBACKS
# ============================================================
def add_bill_callback():
    db = st.session_state.database
    ps = st.session_state.get("product_sel", "")
    bv = st.session_state.get("boxes", 0)
    tv = st.session_state.get("tp_box", 0.0)
    dv = st.session_state.get("discount", 0.0)

    if not ps: st.session_state["error_msg"] = "❌ Select Product"; return
    if bv <= 0: st.session_state["error_msg"] = "❌ Enter Boxes"; return

    sel = next((p for p in PRODUCTS if p["name"] == ps), None)
    if not sel: st.session_state["error_msg"] = "❌ Invalid Product"; return

    gv = bv * tv
    nv = gv - (gv * dv / 100)

    db["bills"].append({
        "Bill No": db["next_bill_no"], "Date": datetime.now().strftime("%d-%m-%Y"),
        "Shop": st.session_state.get("shop_name", "").strip(),
        "Order Booker": st.session_state.get("order_booker", "").strip(),
        "Salesman": st.session_state.get("salesman", "").strip(),
        "Delivery Man": "",
        "Code": sel["code"], "Product": sel["name"],
        "Boxes": bv, "TP/Box": tv, "Discount %": dv,
        "Gross": gv, "Net": nv
    })
    save_database(db)
    st.session_state["last_bill_no"] = db["next_bill_no"]
    st.session_state["success_msg"] = f"✅ Bill Added | #{db['next_bill_no']}"
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    st.session_state["_prev_prod"] = None
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0

def refresh_callback():
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    st.session_state["_prev_prod"] = None
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["success_msg"] = "✅ Ready"

def export_bill_callback():
    db = st.session_state.database
    if len(db["bills"]) == 0: st.session_state["error_msg"] = "❌ No Bills"; return
    shop = st.session_state.get("shop_name", "").strip() or "Bill"
    for ch in ['\\','/',':','*','?','"','<','>','|']: shop = shop.replace(ch, "")
    sb = [b for b in db["bills"] if b["Shop"].strip() == shop]
    if not sb: st.session_state["error_msg"] = "❌ No Bills for shop"; return
    items = [{"Code": b["Code"], "Product": b["Product"], "Boxes": b["Boxes"],
              "TP/Box": b["TP/Box"], "Discount %": b["Discount %"],
              "Gross": b["Gross"], "Net": b["Net"]} for b in sb]
    export_single_group_bill(shop, datetime.now().strftime("%d-%m-%Y"),
                             st.session_state.get("order_booker", ""),
                             st.session_state.get("salesman", ""), items,
                             st.session_state.get("last_bill_no") or (db["next_bill_no"] - 1))
    db["next_bill_no"] += 1; save_database(db)
    st.session_state["last_bill_no"] = None

def export_load_form_from_billing_callback():
    db = st.session_state.database
    booker = st.session_state.get("order_booker", "").strip()
    if not booker: st.session_state["error_msg"] = "❌ Select Booker"; return
    bb = [b for b in db["bills"] if b["Order Booker"].strip() == booker]
    if not bb: st.session_state["error_msg"] = f"❌ No bills for {booker}"; return
    summary = {}
    for b in bb:
        c = b["Code"]
        if c not in summary: summary[c] = {"Code": c, "Product": b["Product"], "Boxes": 0}
        summary[c]["Boxes"] += b["Boxes"]
    items = list(summary.values())
    tb = sum(it["Boxes"] for it in items)
    if "load_forms" not in db: db["load_forms"] = []
    nid = max([lf.get("id", 0) for lf in db["load_forms"]] + [0]) + 1
    db["load_forms"].append({
        "id": nid, "date": datetime.now().strftime("%d-%m-%Y"),
        "time": datetime.now().strftime("%H:%M"),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "booker": booker, "items": items,
        "total_boxes": tb, "total_products": len(items)
    })
    save_database(db)
    st.session_state["success_msg"] = f"✅ Load Form #{nid} saved"

def export_load_form_for_booker(booker, bb=None):
    db = st.session_state.database
    if bb is None: bb = [b for b in db["bills"] if b["Order Booker"].strip() == booker]
    if not bb: st.session_state["error_msg"] = "❌ No Bills"; return
    summary = {}
    for b in bb:
        c = b["Code"]
        if c not in summary: summary[c] = {"Product": b["Product"], "Boxes": 0}
        summary[c]["Boxes"] += b["Boxes"]
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Load Form")
    t = wb.add_format({"bold":True, "font_size":16, "align":"center", "border":2})
    h = wb.add_format({"bold":True, "font_size":12, "bg_color":"#BBDEFB", "align":"center", "border":2})
    cl = wb.add_format({"font_size":14, "border":1, "align":"left"})
    cc = wb.add_format({"font_size":14, "border":1, "align":"center"})
    tt = wb.add_format({"bold":True, "font_size":14, "bg_color":"#FFF2CC", "align":"center", "border":2})
    ws.set_column("A:A", 47.86); ws.set_column("B:B", 12.71)
    ws.merge_range("A1:B1", COMPANY_NAME, t)
    ws.write("A3","Order Booker",h); ws.write("B3",booker,cc)
    ws.write("A5","Product",h); ws.write("B5","Boxes",h)
    row = 5; tb = 0
    for item in summary.values():
        ws.write(row, 0, item["Product"], cl); ws.write(row, 1, item["Boxes"], cc)
        tb += item["Boxes"]; row += 1
    ws.write(row, 0, "TOTAL", tt); ws.write(row, 1, tb, tt)
    wb.close(); output.seek(0)
    st.session_state["download_file"] = (f"{booker}_Load_Form.xlsx", output.getvalue())

def export_all_filtered_load_forms(fl):
    if not fl: st.session_state["error_msg"] = "❌ No Load Forms"; return
    output = BytesIO(); wb = xlsxwriter.Workbook(output, {'in_memory': True})
    tf = wb.add_format({"bold":True, "font_size":14, "align":"center", "border":2, "bg_color":"#BBDEFB"})
    hf = wb.add_format({"bold":True, "bg_color":"#E3F2FD", "border":1, "align":"center"})
    cf = wb.add_format({"border":1}); cc = wb.add_format({"border":1, "align":"center"})
    tt = wb.add_format({"bold":True, "bg_color":"#FFF2CC", "border":1, "align":"center"})
    for lf in fl:
        booker = lf.get("booker", "Unknown"); ds = lf.get("date", ""); ts = lf.get("time", "")
        sn = f"{booker}_{ds}".replace("/","-")[:31] or f"LF_{lf.get('id')}"
        base = sn; c = 1; ex = wb.sheetnames
        while sn in ex: sn = f"{base[:28]}_{c}"; c += 1
        ws = wb.add_worksheet(sn)
        ws.set_column("A:A", 45); ws.set_column("B:B", 10); ws.set_column("C:C", 10)
        ws.merge_range("A1:C1", f"{COMPANY_NAME} - Load Form", tf)
        ws.write("A3", "Booker", hf); ws.write("B3", booker, cf)
        ws.write("A4", "Date", hf); ws.write("B4", f"{ds} {ts}", cf)
        ws.write("A6", "Product", hf); ws.write("B6", "Code", hf); ws.write("C6", "Boxes", hf)
        row = 6; tb = 0
        for it in lf.get("items", []):
            ws.write(row, 0, it["Product"], cf)
            ws.write(row, 1, it["Code"], cc)
            ws.write(row, 2, it["Boxes"], cc)
            tb += it["Boxes"]; row += 1
        ws.write(row, 0, "TOTAL", tt); ws.write(row, 1, "", tt); ws.write(row, 2, tb, tt)
    wb.close(); output.seek(0)
    st.session_state["download_file"] = (f"Load_Forms_{datetime.now().strftime('%d-%m-%Y_%H%M')}.xlsx", output.getvalue())

def refresh_load_form_callback():
    db = st.session_state.database
    booker = st.session_state.get("order_booker", "").strip()
    if not booker: st.session_state["error_msg"] = "❌ Enter Booker"; return
    db["bills"] = [b for b in db["bills"] if b["Order Booker"].strip() != booker]
    save_database(db)
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    st.session_state["_prev_prod"] = None
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["success_msg"] = f"✅ Cleared for {booker}"

# ============================================================
# RENDER SELECTED PAGE
# ============================================================
pg = st.session_state["page"]
if pg == "📊 Dashboard": render_dashboard()
elif pg == "🧾 Billing": render_billing()
elif pg == "🛒 All Products": render_all_products()
elif pg == "🎁 Discount": render_discount()
elif pg == "👤 Bookers": render_bookers()
elif pg == "💰 Bookers Salary": render_salaries("bookers")
elif pg == "🧑‍💼 Salesmen": render_salesmen()
elif pg == "💰 Salesmen Salary": render_salaries("salesmen")
elif pg == "💵 Daily Expense": render_daily_expense()
elif pg == "📋 Bills List": render_bills_list()
elif pg == "📦 Load Form": render_load_form()
