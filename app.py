# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# FINAL FIXED VERSION - Full Screen + Working Toggle + Clean Sidebar
# ============================================================

import os
import json
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
# CSS — FULL SCREEN + WORKING TOGGLE + CLEAN SIDEBAR + NO MANAGE APP
# ============================================================
st.markdown("""
<style>
    /* ==== Header transparent (VISIBLE - keeps sidebar toggle alive) ==== */
    header[data-testid="stHeader"] {
        background: transparent !important;
        box-shadow: none !important;
    }
    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    [data-testid="stStatusWidget"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }

    /* ==== Kill "Manage app" button completely ==== */
    [data-testid="manage-app-button"],
    [data-testid="stAppDeployButton"],
    [data-testid="stCloudAppManageButton"],
    .stAppDeployButton,
    iframe[title="streamlit_cloud_status"],
    div[class*="manageApp"],
    div[class*="ManageApp"],
    div[class*="manage-app"],
    button[class*="manageApp"],
    button[class*="ManageApp"],
    a[class*="manageApp"],
    a[class*="ManageApp"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        width: 0 !important;
        height: 0 !important;
        overflow: hidden !important;
    }

    /* Cover bottom-right corner as fallback */
    .stApp::after {
        content: "" !important;
        position: fixed !important;
        bottom: 0 !important;
        right: 0 !important;
        width: 280px !important;
        height: 70px !important;
        background: linear-gradient(135deg, #e0f7fa 0%, #e8f5e9 100%) !important;
        z-index: 2147483640 !important;
        pointer-events: none !important;
    }

    /* ==== Sidebar toggle (both collapsed & expanded) — always visible ==== */
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        z-index: 999999 !important;
    }
    [data-testid="stSidebarCollapsedControl"] button,
    [data-testid="collapsedControl"] button {
        background: linear-gradient(135deg, #4caf50 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 6px rgba(76,175,80,0.4) !important;
        padding: 4px 6px !important;
    }
    [data-testid="stSidebarCollapsedControl"] svg,
    [data-testid="collapsedControl"] svg {
        fill: #ffffff !important;
        color: #ffffff !important;
    }
    [data-testid="stSidebarCollapseButton"] button {
        background: linear-gradient(135deg, #4caf50 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
    }

    /* ==== Main background ==== */
    .stApp {
        background: linear-gradient(135deg, #e0f7fa 0%, #e8f5e9 100%) !important;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        padding-bottom: 1rem !important;
        max-width: 100% !important;
    }

    /* ==== Sidebar ==== */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #c8e6c9 0%, #b2ebf2 100%) !important;
    }

    /* ==== Sidebar radio options — BALANCED size ==== */
    section[data-testid="stSidebar"] .stRadio label p,
    section[data-testid="stSidebar"] .stRadio label span,
    section[data-testid="stSidebar"] .stRadio label {
        font-size: 15px !important;
        font-weight: 600 !important;
        color: #0d3b1e !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label {
        background-color: #ffffff !important;
        border: 1px solid #a5d6a7 !important;
        border-radius: 8px !important;
        margin-bottom: 6px !important;
        padding: 6px 10px !important;
        transition: all 0.15s ease !important;
        box-shadow: 0 1px 3px rgba(76,175,80,0.1) !important;
        cursor: pointer !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
        background-color: #e8f5e9 !important;
        border-color: #4caf50 !important;
        transform: translateX(3px) !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"],
    section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
        background: linear-gradient(135deg, #c8e6c9 0%, #b2ebf2 100%) !important;
        border-color: #4caf50 !important;
    }
    section[data-testid="stSidebar"] div[role="radiogroup"] input[type="radio"] {
        width: 16px !important;
        height: 16px !important;
    }

    /* ==== Inputs ==== */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div > div,
    .stDateInput > div > div > input {
        background-color: #ffffff !important;
        color: #1a1a1a !important;
        border: 2px solid #b2dfdb !important;
        border-radius: 8px !important;
    }
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus {
        border-color: #4caf50 !important;
    }

    /* ==== Buttons ==== */
    .stButton > button {
        background: linear-gradient(135deg, #4caf50 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 6px rgba(76,175,80,0.25) !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #388e3c 0%, #00897b 100%) !important;
    }
    .stButton > button p { color: #ffffff !important; }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #0288d1 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        font-weight: bold !important;
        border-radius: 8px !important;
    }
    .stDownloadButton > button p { color: #ffffff !important; }

    .auto-dl-hidden div[data-testid="stDownloadButton"] {
        position: absolute !important;
        left: -9999px !important;
        top: -9999px !important;
        opacity: 0 !important;
        height: 0 !important;
    }

    /* ==== Cards ==== */
    .metric-card {
        background: #ffffff;
        border: 2px solid #a5d6a7;
        border-radius: 14px;
        padding: 22px;
        text-align: center;
        box-shadow: 0 3px 10px rgba(76,175,80,0.15);
    }
    .metric-card h3 {
        color: #00796b !important;
        font-size: 13px !important;
        margin: 0 !important;
        font-weight: 700 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-card h1 {
        color: #2e7d32 !important;
        font-size: 34px !important;
        margin: 10px 0 0 0 !important;
        font-weight: 800 !important;
    }

    .booker-row {
        background: #ffffff;
        border: 1px solid #a5d6a7;
        border-radius: 10px;
        padding: 12px 18px;
        margin-bottom: 8px;
        box-shadow: 0 2px 6px rgba(76,175,80,0.1);
    }

    .hint-box {
        background: #e8f5e9;
        border-left: 4px solid #4caf50;
        padding: 8px 12px;
        border-radius: 6px;
        color: #2e7d32 !important;
        font-size: 13px;
        margin-top: 4px;
    }

    .summary-box {
        background: #ffffff;
        border: 2px solid #4caf50;
        border-radius: 12px;
        padding: 15px 20px;
        margin-bottom: 15px;
        box-shadow: 0 3px 10px rgba(76,175,80,0.15);
    }

    .lf-simple-card {
        background: #ffffff;
        border-left: 6px solid #4caf50;
        border-radius: 12px;
        padding: 16px 22px;
        margin-bottom: 12px;
        box-shadow: 0 3px 10px rgba(76,175,80,0.15);
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .lf-simple-card .lf-info { display: flex; flex-direction: column; gap: 4px; }
    .lf-simple-card .lf-line1 { font-size: 17px; font-weight: 700; color: #2e7d32; }
    .lf-simple-card .lf-line2 { font-size: 13px; color: #00695c; }
    .lf-simple-card .lf-boxes {
        background: linear-gradient(135deg, #4caf50 0%, #26a69a 100%);
        color: #ffffff;
        font-weight: 800;
        font-size: 20px;
        padding: 10px 18px;
        border-radius: 10px;
        text-align: center;
        min-width: 90px;
        box-shadow: 0 2px 6px rgba(76,175,80,0.35);
    }
    .lf-simple-card .lf-boxes small {
        display: block;
        font-size: 10px;
        font-weight: 500;
        opacity: 0.9;
    }

    .sal-metric {
        display: inline-block;
        padding: 8px 14px;
        margin-right: 8px;
        margin-bottom: 6px;
        border-radius: 8px;
        font-size: 13px;
        font-weight: 600;
    }
    .sal-metric.base { background: #e3f2fd; color: #0d47a1; }
    .sal-metric.adv { background: #fff3e0; color: #e65100; }
    .sal-metric.short { background: #ffebee; color: #c62828; }
    .sal-metric.remain { background: #e8f5e9; color: #1b5e20; }

    .stAlert { border-radius: 10px !important; }
    hr { border-color: #a5d6a7 !important; opacity: 0.6 !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# JS: Aggressive removal of "Manage app" + ensure sidebar toggle alive
# ============================================================
components.html("""
<script>
(function() {
    function nuke_manage_app() {
        try {
            var doc = window.parent.document;
            var sel = [
                '[data-testid="manage-app-button"]',
                '[data-testid="stAppDeployButton"]',
                '[data-testid="stCloudAppManageButton"]',
                '.stAppDeployButton',
                'iframe[title="streamlit_cloud_status"]',
                'div[class*="manageApp"]',
                'div[class*="ManageApp"]',
                'div[class*="manage-app"]',
                'button[class*="manageApp"]',
                'button[class*="ManageApp"]',
                'a[class*="manageApp"]',
                'a[class*="ManageApp"]'
            ];
            sel.forEach(function(s){
                doc.querySelectorAll(s).forEach(function(el){
                    el.style.setProperty('display','none','important');
                    el.style.setProperty('visibility','hidden','important');
                    el.style.setProperty('opacity','0','important');
                    el.style.setProperty('pointer-events','none','important');
                });
            });
            doc.querySelectorAll('button, a, div, span').forEach(function(el){
                try {
                    var t = (el.textContent || '').trim();
                    if (t === 'Manage app' || t === 'Manage App') {
                        var target = el.closest('button') || el.closest('a') || el;
                        target.style.setProperty('display','none','important');
                        target.style.setProperty('visibility','hidden','important');
                    }
                } catch(e){}
            });
        } catch(e) {}
    }

    function fix_sidebar_toggle() {
        try {
            var doc = window.parent.document;
            doc.querySelectorAll('[data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"]').forEach(function(el){
                el.style.setProperty('display', 'flex', 'important');
                el.style.setProperty('visibility', 'visible', 'important');
                el.style.setProperty('opacity', '1', 'important');
                el.style.setProperty('z-index', '999999', 'important');
            });
        } catch(e) {}
    }

    function tick() { nuke_manage_app(); fix_sidebar_toggle(); }

    setTimeout(tick, 200);
    setTimeout(tick, 600);
    setTimeout(tick, 1200);
    setTimeout(tick, 2500);
    setInterval(tick, 800);

    try {
        var obs = new MutationObserver(function(){ tick(); });
        obs.observe(window.parent.document.body, { childList: true, subtree: true });
    } catch(e) {}
})();
</script>
""", height=0)

COMPANY_NAME = "AL-BARAKAH ENTERPRISES"
DATA_FILE = "billing_database.json"

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

PRODUCT_NAMES = [p["name"] for p in PRODUCTS]

# ============================================================
# DATABASE
# ============================================================
def save_database(db):
    try:
        with open(DATA_FILE, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=4, ensure_ascii=False)
    except Exception as e:
        st.warning(f"⚠️ Save warning: {e}")

def load_database():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "bookers" not in data:
                    data["bookers"] = []
                if "salesmen" not in data:
                    data["salesmen"] = []
                if "load_forms" not in data:
                    data["load_forms"] = []
                if "bookers_salaries" not in data:
                    data["bookers_salaries"] = {}
                if "salesmen_salaries" not in data:
                    data["salesmen_salaries"] = {}
                return data
        except Exception:
            pass
    return {
        "next_bill_no": 1,
        "bills": [],
        "bookers": [],
        "salesmen": [],
        "load_forms": [],
        "bookers_salaries": {},
        "salesmen_salaries": {}
    }

def parse_date(dstr):
    try:
        return datetime.strptime(dstr, "%d-%m-%Y").date()
    except Exception:
        return None

if "database" not in st.session_state:
    st.session_state.database = load_database()
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

db = st.session_state.database
for k in ["bookers", "salesmen", "load_forms", "bookers_salaries", "salesmen_salaries"]:
    if k not in db:
        db[k] = [] if k in ["bookers", "salesmen", "load_forms"] else {}

# ============================================================
# AUTO-DOWNLOAD
# ============================================================
def show_auto_download():
    if st.session_state.get("download_file"):
        fname, fdata = st.session_state["download_file"]
        st.session_state["download_file"] = None
        st.session_state["_dl_counter"] += 1

        st.markdown('<div class="auto-dl-hidden">', unsafe_allow_html=True)
        st.download_button(
            label=f"Download {fname}",
            data=fdata,
            file_name=fname,
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

# ============================================================
# SIDEBAR NAVIGATION
# ============================================================
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 15px 0;'>
        <h2 style='color:#2e7d32 !important; margin:0;'>🧾 AL-BARAKAH</h2>
        <p style='color:#00695c !important; font-size:12px; margin:0; font-weight:600;'>ENTERPRISES</p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

    page = st.radio(
        "MENU",
        [
            "📊 Dashboard",
            "🧾 Billing",
            "👤 Bookers",
            "💰 Bookers Salary",
            "🧑‍💼 Salesmen",
            "💰 Salesmen Salary",
            "📋 Bills List",
            "📦 Load Form",
        ],
        key="page_selector",
        label_visibility="collapsed"
    )
    st.session_state["page"] = page

    st.markdown("---")
    st.markdown(f"""
    <div style='padding:10px; color:#00695c !important; font-size:12px;'>
        <p>📅 {datetime.now().strftime('%d-%m-%Y')}</p>
        <p>📦 Products: {len(PRODUCTS)}</p>
        <p>👤 Bookers: {len(db.get('bookers', []))}</p>
        <p>🧑‍💼 Salesmen: {len(db.get('salesmen', []))}</p>
        <p>🧾 Total Bills: {len(db['bills'])}</p>
        <p>📦 Saved Load Forms: {len(db.get('load_forms', []))}</p>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# PAGE: DASHBOARD
# ============================================================
def render_dashboard():
    st.markdown(f"<h1 style='color:#2e7d32 !important;'>📊 Dashboard</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#00695c;font-weight:500;'>Welcome to {COMPANY_NAME} — Overview & Statistics</p>", unsafe_allow_html=True)
    st.markdown("---")

    bills = db["bills"]
    total_bills = len(bills)
    total_boxes = sum(b["Boxes"] for b in bills) if bills else 0
    total_gross = sum(b["Gross"] for b in bills) if bills else 0
    total_net = sum(b["Net"] for b in bills) if bills else 0
    unique_shops = len(set(b["Shop"] for b in bills if b["Shop"])) if bills else 0
    unique_bookers = len(set(b["Order Booker"] for b in bills if b["Order Booker"])) if bills else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"<div class='metric-card'><h3>TOTAL BILLS</h3><h1>{total_bills}</h1></div>", unsafe_allow_html=True)
    with c2:
        st.markdown(f"<div class='metric-card'><h3>TOTAL BOXES</h3><h1>{total_boxes}</h1></div>", unsafe_allow_html=True)
    with c3:
        st.markdown(f"<div class='metric-card'><h3>UNIQUE SHOPS</h3><h1>{unique_shops}</h1></div>", unsafe_allow_html=True)
    with c4:
        st.markdown(f"<div class='metric-card'><h3>ORDER BOOKERS</h3><h1>{unique_bookers}</h1></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"<div class='metric-card'><h3>TOTAL GROSS AMOUNT</h3><h1>Rs {total_gross:,.0f}</h1></div>", unsafe_allow_html=True)
    with c2:
        st.markdown(f"<div class='metric-card'><h3>TOTAL NET AMOUNT</h3><h1>Rs {total_net:,.0f}</h1></div>", unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("### 🕐 Recent Bills (Last 5)")
    if bills:
        recent = bills[-5:][::-1]
        st.dataframe(pd.DataFrame(recent), use_container_width=True, hide_index=True)
    else:
        st.info("Abhi tak koi bill add nahi hua. 'Billing' page pe jao aur pehla bill banao.")

    st.markdown("<br>", unsafe_allow_html=True)
    if bills:
        st.markdown("### 🏆 Top Products (By Boxes)")
        prod_summary = {}
        for b in bills:
            name = b["Product"]
            if name not in prod_summary:
                prod_summary[name] = {"Product": name, "Boxes": 0, "Amount": 0}
            prod_summary[name]["Boxes"] += b["Boxes"]
            prod_summary[name]["Amount"] += b["Net"]
        top_products = sorted(prod_summary.values(), key=lambda x: x["Boxes"], reverse=True)[:5]
        st.dataframe(pd.DataFrame(top_products), use_container_width=True, hide_index=True)

# ============================================================
# PAGE: BOOKERS
# ============================================================
def render_bookers():
    st.markdown(f"<h1 style='color:#2e7d32 !important;'>👤 Bookers</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#00695c;font-weight:500;'>Order Bookers ko add aur manage karo</p>", unsafe_allow_html=True)
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
            st.warning(f"⚠️ '{name}' already exists in the list")
        else:
            db.setdefault("bookers", []).append(name)
            db["bookers"] = sorted(db["bookers"])
            save_database(db)
            st.session_state["booker_msg"] = f"✅ '{name}' added successfully"
            st.rerun()

    if st.session_state.get("booker_msg"):
        st.success(st.session_state["booker_msg"])
        st.session_state["booker_msg"] = None

    st.markdown("---")
    st.markdown(f"### 📋 Saved Bookers ({len(db.get('bookers', []))})")

    bookers = db.get("bookers", [])
    if not bookers:
        st.info("Abhi tak koi booker add nahi hua. Upar se add karo.")
        return

    for i, booker_name in enumerate(bookers):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#2e7d32;'>👤 {booker_name}</b></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑 Delete", key=f"del_booker_{i}_{booker_name}", use_container_width=True):
                db["bookers"].remove(booker_name)
                save_database(db)
                st.session_state["booker_msg"] = f"🗑 '{booker_name}' deleted"
                st.rerun()

# ============================================================
# PAGE: SALESMEN
# ============================================================
def render_salesmen():
    st.markdown(f"<h1 style='color:#2e7d32 !important;'>🧑‍💼 Salesmen</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#00695c;font-weight:500;'>Salesmen ko add aur manage karo</p>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("### ➕ Add New Salesman")
    c1, c2 = st.columns([3, 1])
    with c1:
        new_salesman = st.text_input("Salesman Name:", key="new_salesman_name", placeholder="Enter salesman name...", label_visibility="collapsed")
    with c2:
        add_clicked = st.button("➕ Add Salesman", key="btn_add_salesman", use_container_width=True, type="primary")

    if add_clicked:
        name = new_salesman.strip() if new_salesman else ""
        if name == "":
            st.error("❌ Please enter a salesman name")
        elif name in db.get("salesmen", []):
            st.warning(f"⚠️ '{name}' already exists in the list")
        else:
            db.setdefault("salesmen", []).append(name)
            db["salesmen"] = sorted(db["salesmen"])
            save_database(db)
            st.session_state["salesman_msg"] = f"✅ '{name}' added successfully"
            st.rerun()

    if st.session_state.get("salesman_msg"):
        st.success(st.session_state["salesman_msg"])
        st.session_state["salesman_msg"] = None

    st.markdown("---")
    st.markdown(f"### 📋 Saved Salesmen ({len(db.get('salesmen', []))})")

    salesmen = db.get("salesmen", [])
    if not salesmen:
        st.info("Abhi tak koi salesman add nahi hua. Upar se add karo.")
        return

    for i, salesman_name in enumerate(salesmen):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#2e7d32;'>🧑‍💼 {salesman_name}</b></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑 Delete", key=f"del_salesman_{i}_{salesman_name}", use_container_width=True):
                db["salesmen"].remove(salesman_name)
                save_database(db)
                st.session_state["salesman_msg"] = f"🗑 '{salesman_name}' deleted"
                st.rerun()

# ============================================================
# PAGE: SALARY
# ============================================================
def render_salaries(role_type):
    if role_type == "bookers":
        title = "💰 Bookers Salary"
        emoji = "👤"
        names = db.get("bookers", [])
        sal_key = "bookers_salaries"
        other_page = "👤 Bookers"
    else:
        title = "💰 Salesmen Salary"
        emoji = "🧑‍💼"
        names = db.get("salesmen", [])
        sal_key = "salesmen_salaries"
        other_page = "🧑‍💼 Salesmen"

    st.markdown(f"<h1 style='color:#2e7d32 !important;'>{title}</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#00695c;font-weight:500;'>Base Salary + Advanced + Shortage = Remaining</p>", unsafe_allow_html=True)
    st.markdown("---")

    if sal_key not in db:
        db[sal_key] = {}

    if not names:
        st.info(f"❌ Abhi tak koi {role_type[:-1]} add nahi hua. Pehle **{other_page}** page pe add karo.")
        return

    total_base = 0
    total_adv = 0
    total_short = 0
    for n in names:
        sd = db[sal_key].get(n, {})
        base = sd.get("base_salary", 0)
        txns = sd.get("transactions", [])
        total_base += base
        total_adv += sum(t["amount"] for t in txns if t.get("type") == "advanced")
        total_short += sum(t["amount"] for t in txns if t.get("type") == "shortage")
    total_remaining = total_base - total_adv - total_short

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#2e7d32;font-size:16px;'>📊 Overall Summary</b><br>
        <span style='color:#00695c;'>
            Total Base: <b>Rs {total_base:,.0f}</b> &nbsp;|&nbsp;
            Total Advanced: <b>Rs {total_adv:,.0f}</b> &nbsp;|&nbsp;
            Total Shortage: <b>Rs {total_short:,.0f}</b> &nbsp;|&nbsp;
            Total Remaining: <b>Rs {total_remaining:,.0f}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    for person_name in names:
        sd = db[sal_key].get(person_name, {"base_salary": 0, "transactions": []})
        base = sd.get("base_salary", 0)
        txns = sd.get("transactions", [])
        total_adv_p = sum(t["amount"] for t in txns if t.get("type") == "advanced")
        total_short_p = sum(t["amount"] for t in txns if t.get("type") == "shortage")
        remaining = base - total_adv_p - total_short_p

        with st.expander(f"💰 {emoji} {person_name}  —  Remaining: Rs {remaining:,.0f}", expanded=False):
            c1, c2 = st.columns([3, 1])
            with c1:
                new_base = st.number_input(
                    "Base Salary (Rs):",
                    value=float(base),
                    min_value=0.0,
                    step=500.0,
                    key=f"base_{role_type}_{person_name}"
                )
            with c2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 Save Base", key=f"savebase_{role_type}_{person_name}", use_container_width=True):
                    db[sal_key].setdefault(person_name, {"base_salary": 0, "transactions": []})
                    db[sal_key][person_name]["base_salary"] = float(new_base)
                    save_database(db)
                    st.session_state["success_msg"] = f"✅ Base salary saved for {person_name}"
                    st.rerun()

            st.markdown(f"""
            <div style='margin-top:10px;'>
                <span class='sal-metric base'>Base: Rs {base:,.0f}</span>
                <span class='sal-metric adv'>Advanced: Rs {total_adv_p:,.0f}</span>
                <span class='sal-metric short'>Shortage: Rs {total_short_p:,.0f}</span>
                <span class='sal-metric remain'>Remaining: Rs {remaining:,.0f}</span>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("**➕ Add Transaction**")

            c1, c2, c3, c4, c5 = st.columns([2, 2, 2, 3, 1])
            with c1:
                txn_date = st.date_input("Date", value=date.today(), key=f"txn_date_{role_type}_{person_name}")
            with c2:
                txn_type = st.selectbox("Type", ["Advanced", "Shortage"], key=f"txn_type_{role_type}_{person_name}")
            with c3:
                txn_amt = st.number_input("Amount", min_value=0.0, step=100.0, key=f"txn_amt_{role_type}_{person_name}")
            with c4:
                txn_note = st.text_input("Note (optional)", key=f"txn_note_{role_type}_{person_name}", placeholder="Reason...")
            with c5:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕ Add", key=f"addtxn_{role_type}_{person_name}", use_container_width=True):
                    if txn_amt <= 0:
                        st.session_state["error_msg"] = "❌ Amount must be greater than 0"
                    else:
                        db[sal_key].setdefault(person_name, {"base_salary": base, "transactions": []})
                        existing = db[sal_key][person_name].get("transactions", [])
                        next_id = max([t.get("id", 0) for t in existing] + [0]) + 1
                        new_txn = {
                            "id": next_id,
                            "date": txn_date.strftime("%d-%m-%Y"),
                            "time": datetime.now().strftime("%H:%M"),
                            "type": "advanced" if txn_type == "Advanced" else "shortage",
                            "amount": float(txn_amt),
                            "note": txn_note.strip(),
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        }
                        db[sal_key][person_name]["transactions"].append(new_txn)
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ {txn_type} Rs {txn_amt:,.0f} added for {person_name}"
                        st.rerun()

            if txns:
                st.markdown("**📋 History (date-wise)**")
                sorted_txns = sorted(txns, key=lambda x: x.get("created_at", ""), reverse=True)
                for idx, t in enumerate(sorted_txns):
                    t_id = t.get("id")
                    badge_color = "#e65100" if t["type"] == "advanced" else "#c62828"
                    badge_bg = "#fff3e0" if t["type"] == "advanced" else "#ffebee"
                    badge_text = "Advanced" if t["type"] == "advanced" else "Shortage"
                    note_txt = f" — {t.get('note','')}" if t.get("note") else ""

                    c1, c2 = st.columns([6, 1])
                    with c1:
                        st.markdown(f"""
                        <div style='background:#ffffff;border:1px solid #e0e0e0;border-radius:8px;padding:10px 14px;margin-bottom:6px;'>
                            <span style='color:#00695c;font-size:13px;'>📅 {t.get('date','')} · 🕐 {t.get('time','')}</span><br>
                            <span style='background:{badge_bg};color:{badge_color};padding:3px 10px;border-radius:6px;font-size:12px;font-weight:700;'>{badge_text}</span>
                            <b style='color:#2e7d32;font-size:15px;margin-left:8px;'>Rs {t['amount']:,.0f}</b>
                            <span style='color:#666;font-size:12px;'>{note_txt}</span>
                        </div>
                        """, unsafe_allow_html=True)
                    with c2:
                        if st.button("🗑", key=f"deltxn_{role_type}_{person_name}_{t_id}_{idx}", use_container_width=True):
                            db[sal_key][person_name]["transactions"] = [
                                x for x in db[sal_key][person_name]["transactions"] if x.get("id") != t_id
                            ]
                            save_database(db)
                            st.session_state["success_msg"] = f"🗑 Transaction deleted"
                            st.rerun()
            else:
                st.info("Koi transaction nahi. Upar se add karo.")

# ============================================================
# PAGE: BILLING
# ============================================================
def render_billing():
    st.markdown(f"<h1 style='color:#2e7d32 !important;'>🧾 Billing</h1>", unsafe_allow_html=True)
    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Bill No:", value=str(db["next_bill_no"]), disabled=True, key="dash_bill_no")
    with c2:
        st.text_input("Date:", value=datetime.now().strftime("%d-%m-%Y"), disabled=True, key="dash_bill_date")

    shop_name = st.text_input("Shop:", key="shop_name", placeholder="Enter Shop Name")

    saved_bookers = db.get("bookers", [])
    if saved_bookers:
        booker_options = ["-- Select Booker --"] + saved_bookers
        selected_bk = st.selectbox("Order Booker:", options=booker_options, key="order_booker_select")
        order_booker_value = "" if selected_bk == "-- Select Booker --" else selected_bk
        st.session_state["order_booker"] = order_booker_value
        st.markdown(f"<div class='hint-box'>💡 {len(saved_bookers)} bookers available — select karo aur aage barho</div>", unsafe_allow_html=True)
    else:
        order_booker = st.text_input("Order Booker:", key="order_booker", placeholder="Enter Order Booker")
        st.markdown("<div class='hint-box'>💡 Tip: 'Bookers' page pe jao aur bookers add karo — phir yahan dropdown milega</div>", unsafe_allow_html=True)

    saved_salesmen = db.get("salesmen", [])
    if saved_salesmen:
        salesman_options = ["-- Select Salesman --"] + saved_salesmen
        selected_sm = st.selectbox("Salesman:", options=salesman_options, key="salesman_select")
        salesman_value = "" if selected_sm == "-- Select Salesman --" else selected_sm
        st.session_state["salesman"] = salesman_value
        st.markdown(f"<div class='hint-box'>💡 {len(saved_salesmen)} salesmen available — select karo aur aage barho</div>", unsafe_allow_html=True)
    else:
        salesman = st.text_input("Salesman:", key="salesman", placeholder="Enter Salesman")
        st.markdown("<div class='hint-box'>💡 Tip: 'Salesmen' page pe jao aur salesmen add karo — phir yahan dropdown milega</div>", unsafe_allow_html=True)

    delivery_man = st.text_input("Delivery Man:", key="delivery_man", placeholder="Enter Delivery Man")

    st.markdown("---")
    st.markdown("**🔍 Search Product by Name:**")

    search_text = st.text_input("Search:", key="search_text", placeholder="Type product name...")
    search_upper = search_text.strip().upper()
    filtered_names = [p["name"] for p in PRODUCTS if search_upper in p["name"].upper()] if search_upper else PRODUCT_NAMES

    if st.session_state.get("product_sel") and st.session_state["product_sel"] not in filtered_names:
        st.session_state["product_sel"] = ""

    product_sel = st.selectbox("Select Product:", options=[""] + filtered_names, key="product_sel")

    selected_product = None
    if product_sel:
        for p in PRODUCTS:
            if p["name"] == product_sel:
                selected_product = p
                break

    if selected_product:
        st.success(f"✅ {selected_product['name']}  (Code: {selected_product['code']})")
        tp_default = float(selected_product["price"])
    else:
        st.error("No Product Selected")
        tp_default = 0.0

    if st.session_state["_prev_prod"] != product_sel:
        st.session_state["tp_box"] = tp_default
        st.session_state["_prev_prod"] = product_sel

    st.markdown("---")

    c1, c2, c3 = st.columns(3)
    with c1:
        boxes = st.number_input("Boxes:", min_value=0, step=1, key="boxes")
    with c2:
        tp_box = st.number_input("TP/Box:", min_value=0.0, step=1.0, key="tp_box")
    with c3:
        discount = st.number_input("Discount %:", min_value=0.0, step=0.5, key="discount")

    gross = boxes * tp_box
    net = gross - (gross * discount / 100)

    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Gross:", value=f"{gross:.2f}", disabled=True, key="gross_disp")
    with c2:
        st.text_input("Net:", value=f"{net:.2f}", disabled=True, key="net_disp")

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"])
        st.session_state["success_msg"] = None
    if st.session_state.get("error_msg"):
        st.error(st.session_state["error_msg"])
        st.session_state["error_msg"] = None

    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.button("➕ Add Bill", key="btn_add", on_click=add_bill_callback, use_container_width=True, type="primary")
    with c2:
        st.button("🔄 Refresh", key="btn_refresh", on_click=refresh_callback, use_container_width=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.button("📄 Export Bill", key="btn_export", on_click=export_bill_callback, use_container_width=True)
    with c2:
        st.button("📦 Export Load Form", key="btn_export_lf", on_click=export_load_form_from_billing_callback, use_container_width=True)
    with c3:
        st.button("🗑 Refresh Load Form", key="btn_load_refresh", on_click=refresh_load_form_callback, use_container_width=True)

    show_auto_download()

# ============================================================
# PAGE: BILLS LIST
# ============================================================
def render_bills_list():
    st.markdown(f"<h1 style='color:#2e7d32 !important;'>📋 Bills List</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#00695c;font-weight:500;'>Total {len(db['bills'])} bills in database</p>", unsafe_allow_html=True)
    st.markdown("---")

    if len(db["bills"]) == 0:
        st.info("❌ Koi bill nahi mila. Pehle 'Billing' page pe jaake bill banao.")
        return

    st.markdown("### 🔎 Filter Bills")
    today = date.today()

    c1, c2, c3 = st.columns(3)
    with c1:
        filter_mode = st.selectbox(
            "Filter Mode:",
            ["📅 Aaj Ki Bills (Today)", "📆 Custom Date Range", "🗓️ Specific Date", "📋 All Bills"],
            key="bills_filter_mode"
        )
    with c2:
        from_date = st.date_input("From Date:", value=today - timedelta(days=7), key="bills_from_date")
    with c3:
        to_date = st.date_input("To Date:", value=today, key="bills_to_date")

    c1, c2 = st.columns(2)
    with c1:
        search = st.text_input("🔍 Search (Shop / Product / Booker):", key="bills_search")
    with c2:
        shops_list = sorted(set(b["Shop"] for b in db["bills"] if b["Shop"]))
        shop_filter = st.selectbox("Filter by Shop:", options=["All"] + shops_list, key="shop_filter")

    df = pd.DataFrame(db["bills"])
    df["_parsed_date"] = df["Date"].apply(parse_date)
    filtered_df = df.copy()

    if filter_mode == "📅 Aaj Ki Bills (Today)":
        filtered_df = filtered_df[filtered_df["_parsed_date"] == today]
    elif filter_mode == "🗓️ Specific Date":
        filtered_df = filtered_df[filtered_df["_parsed_date"] == from_date]
    elif filter_mode == "📆 Custom Date Range":
        filtered_df = filtered_df[(filtered_df["_parsed_date"] >= from_date) & (filtered_df["_parsed_date"] <= to_date)]

    if search:
        s = search.upper()
        mask = (
            filtered_df["Shop"].astype(str).str.upper().str.contains(s, na=False) |
            filtered_df["Product"].astype(str).str.upper().str.contains(s, na=False) |
            filtered_df["Order Booker"].astype(str).str.upper().str.contains(s, na=False)
        )
        filtered_df = filtered_df[mask]

    if shop_filter != "All":
        filtered_df = filtered_df[filtered_df["Shop"] == shop_filter]

    filtered_df = filtered_df.drop(columns=["_parsed_date"])

    if len(filtered_df) > 0:
        total_boxes = int(filtered_df["Boxes"].sum())
        total_gross = float(filtered_df["Gross"].sum())
        total_net = float(filtered_df["Net"].sum())
        unique_shops = filtered_df["Shop"].nunique()

        st.markdown(f"""
        <div class='summary-box'>
            <b style='color:#2e7d32;font-size:16px;'>📊 Summary</b><br>
            <span style='color:#00695c;'>
                Bills: <b>{len(filtered_df)}</b> &nbsp;|&nbsp;
                Boxes: <b>{total_boxes}</b> &nbsp;|&nbsp;
                Shops: <b>{unique_shops}</b> &nbsp;|&nbsp;
                Gross: <b>Rs {total_gross:,.0f}</b> &nbsp;|&nbsp;
                Net: <b>Rs {total_net:,.0f}</b>
            </span>
        </div>
        """, unsafe_allow_html=True)

    if len(filtered_df) == 0:
        st.warning("❌ Is filter ke hisaab se koi bill nahi mila.")
    else:
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)
        st.caption(f"Showing {len(filtered_df)} of {len(df)} bills")

    st.markdown("---")
    if len(filtered_df) > 0:
        output = BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})
        ws = wb.add_worksheet("Bills")
        header_fmt = wb.add_format({"bold": True, "bg_color": "#D9EAD3", "border": 1, "align": "center"})
        cell_fmt = wb.add_format({"border": 1})

        for i, col in enumerate(filtered_df.columns):
            ws.write(0, i, col, header_fmt)
        for r, (_, row) in enumerate(filtered_df.iterrows(), start=1):
            for c, col in enumerate(filtered_df.columns):
                ws.write(r, c, row[col], cell_fmt)
        wb.close()
        output.seek(0)

        st.download_button(
            label=f"⬇️ Download Filtered Bills ({len(filtered_df)} rows)",
            data=output.getvalue(),
            file_name=f"bills_filtered_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_filtered_bills",
            use_container_width=True,
        )

# ============================================================
# PAGE: LOAD FORM
# ============================================================
def render_load_form():
    st.markdown(f"<h1 style='color:#2e7d32 !important;'>📦 Saved Load Forms</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#00695c;font-weight:500;'>Billing page se export kiye gaye load forms</p>", unsafe_allow_html=True)
    st.markdown("---")

    load_forms = db.get("load_forms", [])
    if not load_forms:
        st.info("❌ Abhi tak koi load form save nahi hua. Billing page pe '📦 Export Load Form' click karo.")
        return

    st.markdown("### 🔎 Filter Load Forms")
    today = date.today()

    c1, c2, c3 = st.columns(3)
    with c1:
        filter_mode = st.selectbox(
            "Filter Mode:",
            ["📅 Aaj Ke Load Forms (Today)", "📆 Custom Date Range", "🗓️ Specific Date", "📋 All Load Forms"],
            key="lf_filter_mode"
        )
    with c2:
        from_date = st.date_input("From Date:", value=today - timedelta(days=7), key="lf_from_date")
    with c3:
        to_date = st.date_input("To Date:", value=today, key="lf_to_date")

    all_booker_names = sorted(set(lf["booker"] for lf in load_forms if lf.get("booker")))
    booker_filter = st.selectbox("Filter by Booker:", options=["All"] + all_booker_names, key="lf_booker_filter")

    filtered_lfs = []
    for lf in load_forms:
        lf_date = parse_date(lf.get("date", ""))
        if lf_date is None:
            continue

        if filter_mode == "📅 Aaj Ke Load Forms (Today)":
            if lf_date != today:
                continue
        elif filter_mode == "🗓️ Specific Date":
            if lf_date != from_date:
                continue
        elif filter_mode == "📆 Custom Date Range":
            if not (from_date <= lf_date <= to_date):
                continue

        if booker_filter != "All" and lf.get("booker") != booker_filter:
            continue

        filtered_lfs.append(lf)

    if not filtered_lfs:
        st.warning("❌ Is filter ke hisaab se koi load form nahi mila.")
        return

    total_forms = len(filtered_lfs)
    total_boxes_all = sum(lf.get("total_boxes", 0) for lf in filtered_lfs)

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#2e7d32;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#00695c;'>
            Load Forms: <b>{total_forms}</b> &nbsp;|&nbsp;
            Total Boxes: <b>{total_boxes_all}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    if st.button("⬇️ Download All Filtered Load Forms (Excel)", key="lf_dl_all", use_container_width=True):
        export_all_filtered_load_forms(filtered_lfs)
        st.rerun()

    st.markdown("---")
    st.markdown(f"### 📋 Load Forms ({len(filtered_lfs)})")

    sorted_lfs = sorted(filtered_lfs, key=lambda x: x.get("created_at", ""), reverse=True)

    for idx, lf in enumerate(sorted_lfs):
        lf_id = lf.get("id", idx)
        booker = lf.get("booker", "Unknown")
        date_str = lf.get("date", "")
        time_str = lf.get("time", "")
        total_boxes = lf.get("total_boxes", 0)

        c1, c2 = st.columns([4, 1])
        with c1:
            st.markdown(f"""
            <div class='lf-simple-card'>
                <div class='lf-info'>
                    <div class='lf-line1'>👤 {booker}</div>
                    <div class='lf-line2'>📅 {date_str} &nbsp;·&nbsp; 🕐 {time_str}</div>
                </div>
                <div class='lf-boxes'>
                    {total_boxes}
                    <small>BOXES</small>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with c2:
            if st.button("⬇️ Download", key=f"dl_lf_{lf_id}_{idx}", use_container_width=True):
                booker_bills = [{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"]} for it in lf.get("items", [])]
                export_load_form_for_booker(booker, booker_bills)
                st.rerun()
            if st.button("🗑 Delete", key=f"del_lf_{lf_id}_{idx}", use_container_width=True):
                db["load_forms"] = [x for x in db["load_forms"] if x.get("id") != lf_id]
                save_database(db)
                st.session_state["success_msg"] = f"🗑 Load Form #{lf_id} deleted"
                st.rerun()

    show_auto_download()

# ============================================================
# CALLBACKS
# ============================================================
def add_bill_callback():
    db = st.session_state.database
    product_sel = st.session_state.get("product_sel", "")
    boxes_v = st.session_state.get("boxes", 0)
    tp_v = st.session_state.get("tp_box", 0.0)
    disc_v = st.session_state.get("discount", 0.0)

    if not product_sel:
        st.session_state["error_msg"] = "❌ Please Select Product from dropdown"
        return
    if boxes_v <= 0:
        st.session_state["error_msg"] = "❌ Enter Boxes"
        return

    sel = next((p for p in PRODUCTS if p["name"] == product_sel), None)
    if not sel:
        st.session_state["error_msg"] = "❌ Invalid Product"
        return

    gross_v = boxes_v * tp_v
    net_v = gross_v - (gross_v * disc_v / 100)

    bill = {
        "Bill No": db["next_bill_no"],
        "Date": datetime.now().strftime("%d-%m-%Y"),
        "Shop": st.session_state.get("shop_name", "").strip(),
        "Order Booker": st.session_state.get("order_booker", "").strip(),
        "Salesman": st.session_state.get("salesman", "").strip(),
        "Delivery Man": st.session_state.get("delivery_man", "").strip(),
        "Code": sel["code"],
        "Product": sel["name"],
        "Boxes": boxes_v,
        "TP/Box": tp_v,
        "Discount %": disc_v,
        "Gross": gross_v,
        "Net": net_v
    }

    db["bills"].append(bill)
    save_database(db)
    st.session_state["last_bill_no"] = db["next_bill_no"]
    st.session_state["success_msg"] = f"✅ Bill Added | Bill No: {db['next_bill_no']} | Total: {len(db['bills'])}"

    st.session_state["search_text"] = ""
    st.session_state["product_sel"] = ""
    st.session_state["_prev_prod"] = None
    st.session_state["boxes"] = 0
    st.session_state["tp_box"] = 0.0
    st.session_state["discount"] = 0.0


def refresh_callback():
    st.session_state["search_text"] = ""
    st.session_state["product_sel"] = ""
    st.session_state["_prev_prod"] = None
    st.session_state["boxes"] = 0
    st.session_state["tp_box"] = 0.0
    st.session_state["discount"] = 0.0
    st.session_state["success_msg"] = "✅ Ready For Next Product"


def export_bill_callback():
    db = st.session_state.database
    if len(db["bills"]) == 0:
        st.session_state["error_msg"] = "❌ No Bills Found"
        return

    shop = st.session_state.get("shop_name", "").strip() or "Bill"
    for ch in ['\\','/',':','*','?','"','<','>','|']:
        shop = shop.replace(ch, "")

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Bill")

    worksheet.set_paper(9)
    worksheet.set_portrait()
    worksheet.fit_to_pages(1, 1)
    worksheet.set_column("A:A", 47.86)
    worksheet.set_column("B:B", 23.71)
    worksheet.set_column("C:C", 12.71)
    worksheet.set_column("D:D", 12.71)
    worksheet.set_column("E:E", 12.71)
    worksheet.set_column("F:F", 12.14)
    worksheet.set_column("G:G", 13.14)
    worksheet.set_column("H:H", 14.14)

    title = workbook.add_format({"bold":True, "font_size":18, "align":"center", "border":2})
    header = workbook.add_format({"bold":True, "font_size":12, "bg_color":"#D9EAD3", "align":"center", "border":2})
    cell_left = workbook.add_format({"font_size":14, "border":1, "align":"left"})
    cell_center = workbook.add_format({"font_size":14, "border":1, "align":"center"})
    total = workbook.add_format({"bold":True, "font_size":14, "bg_color":"#FFF2CC", "align":"center", "border":2})

    worksheet.merge_range("A1:H1", COMPANY_NAME, title)
    worksheet.write("A3","Shop Name",header)
    worksheet.write("B3", st.session_state.get("shop_name", ""), cell_center)
    worksheet.write("D3","Booker",header)
    worksheet.write("E3", st.session_state.get("order_booker", ""), cell_center)
    worksheet.write("G3","Bill No",header)
    last_bill = st.session_state.get("last_bill_no") or (db["next_bill_no"] - 1)
    worksheet.write("H3", last_bill, cell_center)
    worksheet.write("G4","Date",header)
    worksheet.write("H4", datetime.now().strftime("%d-%m-%Y"), cell_center)

    row = 6
    headers = ["Product", "Code", "Boxes", "TP/Box", "Gross", "Discount %", "Net"]
    for col, value in enumerate(headers):
        worksheet.write(row, col, value, header)
    row += 1

    gross_total = 0
    total_boxes = 0
    shop_filter = st.session_state.get("shop_name", "").strip()

    for bill in db["bills"]:
        if bill["Shop"].strip() != shop_filter:
            continue
        worksheet.write(row, 0, bill["Product"], cell_left)
        worksheet.write(row, 1, bill["Code"], cell_center)
        worksheet.write(row, 2, bill["Boxes"], cell_center)
        worksheet.write(row, 3, bill["TP/Box"], cell_center)
        worksheet.write(row, 4, bill["Gross"], cell_center)
        worksheet.write(row, 5, bill["Discount %"], cell_center)
        excel_row = row + 1
        worksheet.write_formula(row, 6, f"=E{excel_row}-(E{excel_row}*F{excel_row}/100)", cell_center)
        gross_total += bill["Gross"]
        total_boxes += bill["Boxes"]
        row += 1

    worksheet.write(row, 1, "TOTAL", total)
    worksheet.write(row, 2, total_boxes, total)
    worksheet.write(row, 4, gross_total, total)
    worksheet.write_blank(row, 5, None, total)
    worksheet.write_formula(row, 6, f"=E{row+1}-(E{row+1}*F{row+1}/100)", total)

    workbook.close()
    output.seek(0)

    st.session_state["download_file"] = (f"{shop}.xlsx", output.getvalue())
    db["next_bill_no"] += 1
    save_database(db)
    st.session_state["last_bill_no"] = None
    st.session_state["success_msg"] = f"✅ Bill Exported | Next Bill No: {db['next_bill_no']}"


def export_load_form_from_billing_callback():
    db = st.session_state.database
    booker = st.session_state.get("order_booker", "").strip()

    if not booker:
        st.session_state["error_msg"] = "❌ Please select Order Booker first"
        return

    booker_bills = [b for b in db["bills"] if b["Order Booker"].strip() == booker]
    if not booker_bills:
        st.session_state["error_msg"] = f"❌ No bills found for booker: {booker}"
        return

    summary = {}
    for b in booker_bills:
        code = b["Code"]
        if code not in summary:
            summary[code] = {"Code": code, "Product": b["Product"], "Boxes": 0}
        summary[code]["Boxes"] += b["Boxes"]

    items = list(summary.values())
    total_boxes = sum(it["Boxes"] for it in items)

    if "load_forms" not in db:
        db["load_forms"] = []

    next_id = 1
    if db["load_forms"]:
        next_id = max(lf.get("id", 0) for lf in db["load_forms"]) + 1

    lf_record = {
        "id": next_id,
        "date": datetime.now().strftime("%d-%m-%Y"),
        "time": datetime.now().strftime("%H:%M"),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "booker": booker,
        "items": items,
        "total_boxes": total_boxes,
        "total_products": len(items)
    }
    db["load_forms"].append(lf_record)
    save_database(db)

    export_load_form_for_booker(booker, booker_bills)
    st.session_state["success_msg"] = f"✅ Load Form #{next_id} saved & exported | {len(items)} products, {total_boxes} boxes"


def export_load_form_for_booker(booker, booker_bills=None):
    db = st.session_state.database
    if booker_bills is None:
        booker_bills = [b for b in db["bills"] if b["Order Booker"].strip() == booker]
    if not booker_bills:
        st.session_state["error_msg"] = "❌ No Bills Found for this Booker"
        return

    summary = {}
    for bill in booker_bills:
        code = bill["Code"]
        if code not in summary:
            summary[code] = {"Product": bill["Product"], "Boxes": 0}
        summary[code]["Boxes"] += bill["Boxes"]

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Load Form")

    title = workbook.add_format({"bold":True, "font_size":16, "align":"center", "border":2})
    header = workbook.add_format({"bold":True, "font_size":12, "bg_color":"#D9EAD3", "align":"center", "border":2})
    cell_left = workbook.add_format({"font_size":14, "border":1, "align":"left"})
    cell_center = workbook.add_format({"font_size":14, "border":1, "align":"center"})
    total = workbook.add_format({"bold":True, "font_size":14, "bg_color":"#FFF2CC", "align":"center", "border":2})

    worksheet.set_column("A:A", 47.86)
    worksheet.set_column("B:B", 12.71)
    worksheet.merge_range("A1:B1", COMPANY_NAME, title)
    worksheet.write("A3","Order Booker",header)
    worksheet.write("B3",booker,cell_center)
    worksheet.write("A5","Product",header)
    worksheet.write("B5","Boxes",header)

    row = 5
    total_boxes = 0
    for item in summary.values():
        worksheet.write(row, 0, item["Product"], cell_left)
        worksheet.write(row, 1, item["Boxes"], cell_center)
        total_boxes += item["Boxes"]
        row += 1

    worksheet.write(row, 0, "TOTAL", total)
    worksheet.write(row, 1, total_boxes, total)

    workbook.close()
    output.seek(0)

    st.session_state["download_file"] = (f"{booker}_Load_Form.xlsx", output.getvalue())


def export_all_filtered_load_forms(filtered_lfs):
    if not filtered_lfs:
        st.session_state["error_msg"] = "❌ No Load Forms to export"
        return

    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})

    title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center", "border": 2, "bg_color": "#D9EAD3"})
    header_fmt = wb.add_format({"bold": True, "bg_color": "#E8F5E9", "border": 1, "align": "center"})
    cell_fmt = wb.add_format({"border": 1})
    cell_center = wb.add_format({"border": 1, "align": "center"})
    total_fmt = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "align": "center"})

    for lf in filtered_lfs:
        booker = lf.get("booker", "Unknown")
        date_str = lf.get("date", "")
        time_str = lf.get("time", "")
        sheet_name = f"{booker}_{date_str}".replace("/", "-")[:31] or f"LF_{lf.get('id')}"
        base_name = sheet_name
        counter = 1
        existing = wb.sheetnames
        while sheet_name in existing:
            sheet_name = f"{base_name[:28]}_{counter}"
            counter += 1

        ws = wb.add_worksheet(sheet_name)
        ws.set_column("A:A", 45)
        ws.set_column("B:B", 10)
        ws.set_column("C:C", 10)

        ws.merge_range("A1:C1", f"{COMPANY_NAME} - Load Form", title_fmt)
        ws.write("A3", "Booker", header_fmt)
        ws.write("B3", booker, cell_fmt)
        ws.write("A4", "Date", header_fmt)
        ws.write("B4", f"{date_str} {time_str}", cell_fmt)

        ws.write("A6", "Product", header_fmt)
        ws.write("B6", "Code", header_fmt)
        ws.write("C6", "Boxes", header_fmt)

        row = 6
        total_boxes = 0
        for it in lf.get("items", []):
            ws.write(row, 0, it["Product"], cell_fmt)
            ws.write(row, 1, it["Code"], cell_center)
            ws.write(row, 2, it["Boxes"], cell_center)
            total_boxes += it["Boxes"]
            row += 1

        ws.write(row, 0, "TOTAL", total_fmt)
        ws.write(row, 1, "", total_fmt)
        ws.write(row, 2, total_boxes, total_fmt)

    wb.close()
    output.seek(0)

    fname = f"Load_Forms_{datetime.now().strftime('%d-%m-%Y_%H%M')}.xlsx"
    st.session_state["download_file"] = (fname, output.getvalue())


def refresh_load_form_callback():
    db = st.session_state.database
    booker = st.session_state.get("order_booker", "").strip()
    if not booker:
        st.session_state["error_msg"] = "❌ Please Enter Order Booker"
        return
    db["bills"] = [b for b in db["bills"] if b["Order Booker"].strip() != booker]
    save_database(db)
    st.session_state["search_text"] = ""
    st.session_state["product_sel"] = ""
    st.session_state["_prev_prod"] = None
    st.session_state["boxes"] = 0
    st.session_state["tp_box"] = 0.0
    st.session_state["discount"] = 0.0
    st.session_state["success_msg"] = f"✅ Load Form Cleared | Booker: {booker}"

# ============================================================
# RENDER SELECTED PAGE
# ============================================================
if st.session_state["page"] == "📊 Dashboard":
    render_dashboard()
elif st.session_state["page"] == "🧾 Billing":
    render_billing()
elif st.session_state["page"] == "👤 Bookers":
    render_bookers()
elif st.session_state["page"] == "💰 Bookers Salary":
    render_salaries("bookers")
elif st.session_state["page"] == "🧑‍💼 Salesmen":
    render_salesmen()
elif st.session_state["page"] == "💰 Salesmen Salary":
    render_salaries("salesmen")
elif st.session_state["page"] == "📋 Bills List":
    render_bills_list()
elif st.session_state["page"] == "📦 Load Form":
    render_load_form()
