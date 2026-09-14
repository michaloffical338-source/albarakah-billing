# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# + Eye Button (Full Bill View) in Bills List & Load Form
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
# SIDEBAR TOGGLE + MANAGE APP KILLER
# ============================================================
components.html("""
<script>
(function(){
    function killManageApp() {
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
                'button[class*="manageApp"]',
                'button[class*="ManageApp"]'
            ];
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
                    if (t === 'Manage app' || t === 'Manage App') {
                        el.style.setProperty('display','none','important');
                    }
                } catch(e){}
            });
        } catch(e) {}
    }

    function attachToggle() {
        try {
            var doc = window.parent.document;
            var old = doc.getElementById('custom-sidebar-toggle');
            if (old) old.parentNode.removeChild(old);
            var btn = doc.createElement('button');
            btn.id = 'custom-sidebar-toggle';
            btn.title = 'Sidebar Open/Close';
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
                var targets = [
                    '[data-testid="stSidebarCollapseButton"] button',
                    '[data-testid="stSidebarCollapsedControl"] button',
                    '[data-testid="collapsedControl"] button',
                    '[data-testid="stExpandSidebarButton"] button',
                    'button[kind="headerNoPadding"]'
                ];
                for (var i = 0; i < targets.length; i++) {
                    var el = doc.querySelector(targets[i]);
                    if (el) { el.click(); return; }
                }
            };
            if (doc.body) doc.body.appendChild(btn);
        } catch(e) {}
    }

    function tick() { killManageApp(); attachToggle(); }
    setTimeout(tick, 300);
    setTimeout(tick, 1000);
    setTimeout(tick, 2000);
    setInterval(tick, 1500);
})();
</script>
""", height=0)

# ============================================================
# THEME CSS
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

    /* Full bill view box */
    .full-bill-box {
        background: #ffffff;
        border: 3px solid #2196f3;
        border-radius: 14px;
        padding: 18px 22px;
        margin: 10px 0 18px 0;
        box-shadow: 0 6px 20px rgba(33,150,243,0.25);
    }
    .full-bill-title {
        font-size: 20px; font-weight: 800; color: #1976d2; margin-bottom: 8px;
    }
    .full-bill-meta {
        font-size: 13px; color: #0277bd; margin-bottom: 12px;
    }

    .stAlert { border-radius: 10px !important; }
    hr { border-color: #90caf9 !important; opacity: 0.6 !important; }
</style>
""", unsafe_allow_html=True)

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

# ============================================================
# DEFAULT DISCOUNT PACKAGES
# ============================================================
def default_discount_packages():
    return [
        {"id": 1, "name": "Package 1",
         "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0,
         "tier3_amount": 3100.0, "tier3_pct": 3.0,
         "active": True},
        {"id": 2, "name": "Package 2",
         "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0,
         "tier3_amount": 3100.0, "tier3_pct": 3.0,
         "active": True},
        {"id": 3, "name": "Package 3",
         "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0,
         "tier3_amount": 3100.0, "tier3_pct": 3.0,
         "active": True},
        {"id": 4, "name": "Package 4",
         "tier1_amount": 1200.0, "tier1_pct": 1.0,
         "tier2_amount": 2100.0, "tier2_pct": 2.0,
         "tier3_amount": 3100.0, "tier3_pct": 3.0,
         "active": True},
    ]

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
                return data
        except Exception:
            pass
    return {
        "next_bill_no": 1, "bills": [], "bookers": [], "salesmen": [],
        "load_forms": [], "bookers_salaries": {}, "salesmen_salaries": {},
        "product_prices": {}, "petrol_expenses": [], "lunch_expenses": [],
        "discount_packages": default_discount_packages()
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

# ============================================================
# EXPORT SINGLE GROUP BILL
# ============================================================
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
        b_net = float(it.get("Net", 0))
        b_gross = float(it.get("Gross", 0))
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
    st.markdown("""
    <div style='text-align:center; padding: 15px 0;'>
        <h2 style='color:#1976d2 !important; margin:0;'>🧾 AL-BARAKAH</h2>
        <p style='color:#0277bd !important; font-size:12px; margin:0; font-weight:600;'>ENTERPRISES</p>
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
    <div style='padding:10px; color:#0277bd !important; font-size:12px;'>
        <p>📅 {datetime.now().strftime('%d-%m-%Y')}</p>
        <p>🎁 Active Packages: <b>{len(active_pkgs)}</b></p>
        <p>📦 Products: {len(PRODUCTS)}</p>
        <p>👤 Bookers: {len(db.get('bookers', []))}</p>
        <p>🧑‍💼 Salesmen: {len(db.get('salesmen', []))}</p>
        <p>🧾 Total Bills: {len(db['bills'])}</p>
    </div>
    """, unsafe_allow_html=True)

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
        <br><span style='font-size:12px;color:#0277bd;'>
            Har package apne aap check hoga — jo sabse zyada % de raha ho, wahi apply hoga.
        </span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📝 Packages")
    st.caption("👇 Har package ke saath **Active** checkbox hai — jitne chahiye utne active rakho.")

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
                st.session_state["success_msg"] = f"✅ Package saved"
                st.rerun()
        with sc2:
            st.markdown(
                f"<div style='padding-top:10px;font-size:11px;color:#0277bd;'>"
                f"&gt; Rs {ta1:,.0f} → {tp1}% &nbsp;|&nbsp; "
                f"&gt; Rs {ta2:,.0f} → {tp2}% &nbsp;|&nbsp; "
                f"&gt; Rs {ta3:,.0f} → {tp3}%</div>",
                unsafe_allow_html=True
            )

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
    st.markdown("<p style='color:#0277bd;font-weight:500;'>Order Bookers ko add aur manage karo</p>", unsafe_allow_html=True)
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
        st.success(st.session_state["booker_msg"]); st.session_state["booker_msg"] = None

    st.markdown("---")
    st.markdown(f"### 📋 Saved Bookers ({len(db.get('bookers', []))})")

    bookers = db.get("bookers", [])
    if not bookers:
        st.info("Abhi tak koi booker add nahi hua.")
        return

    for i, booker_name in enumerate(bookers):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#1976d2;'>👤 {booker_name}</b></div>", unsafe_allow_html=True)
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
    st.markdown(f"<h1 style='color:#1976d2 !important;'>🧑‍💼 Salesmen</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#0277bd;font-weight:500;'>Salesmen ko add aur manage karo</p>", unsafe_allow_html=True)
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
        st.success(st.session_state["salesman_msg"]); st.session_state["salesman_msg"] = None

    st.markdown("---")
    st.markdown(f"### 📋 Saved Salesmen ({len(db.get('salesmen', []))})")

    salesmen = db.get("salesmen", [])
    if not salesmen:
        st.info("Abhi tak koi salesman add nahi hua.")
        return

    for i, salesman_name in enumerate(salesmen):
        c1, c2 = st.columns([5, 1])
        with c1:
            st.markdown(f"<div class='booker-row'><b style='font-size:16px;color:#1976d2;'>🧑‍💼 {salesman_name}</b></div>", unsafe_allow_html=True)
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
        title = "💰 Bookers Salary"; emoji = "👤"; names = db.get("bookers", [])
        sal_key = "bookers_salaries"; other_page = "👤 Bookers"
    else:
        title = "💰 Salesmen Salary"; emoji = "🧑‍💼"; names = db.get("salesmen", [])
        sal_key = "salesmen_salaries"; other_page = "🧑‍💼 Salesmen"

    st.markdown(f"<h1 style='color:#1976d2 !important;'>{title}</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#0277bd;font-weight:500;'>Base - Pending Advanced - Pending Shortage = Remaining</p>", unsafe_allow_html=True)
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
            Base: <b>Rs {total_base:,.0f}</b> | Adv Pending: <b>Rs {total_adv_pending:,.0f}</b> |
            Adv Paid: <b>Rs {total_adv_paid:,.0f}</b> | Short Pending: <b>Rs {total_short_pending:,.0f}</b> |
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
                    st.session_state["success_msg"] = f"✅ Base salary saved for {person_name}"
                    st.rerun()

            st.markdown(f"""
            <div style='margin-top:10px;'>
                <span class='sal-metric base'>Base: Rs {base:,.0f}</span>
                <span class='sal-metric adv'>Adv Pending: Rs {adv_pending:,.0f}</span>
                <span class='sal-metric paid'>Adv Paid: Rs {adv_paid:,.0f}</span>
                <span class='sal-metric short'>Short Pending: Rs {short_pending:,.0f}</span>
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
            with c4: txn_note = st.text_input("Note (optional)", key=f"txn_note_{role_type}_{person_name}", placeholder="Reason...")
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
                            "id": next_id, "date": txn_date.strftime("%d-%m-%Y"),
                            "time": datetime.now().strftime("%H:%M"),
                            "type": "advanced" if txn_type == "Advanced" else "shortage",
                            "amount": float(txn_amt), "note": txn_note.strip(),
                            "status": "pending",
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        }
                        db[sal_key][person_name]["transactions"].append(new_txn)
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ {txn_type} Rs {txn_amt:,.0f} added (Pending)"
                        st.rerun()

            if txns:
                st.markdown("**📋 History (date-wise)**")
                sorted_txns = sorted(txns, key=lambda x: x.get("created_at", ""), reverse=True)
                for idx, t in enumerate(sorted_txns):
                    t_id = t.get("id")
                    status = t.get("status", "pending")
                    is_pending = status == "pending"
                    is_adv = t["type"] == "advanced"
                    if is_adv:
                        badge_text = "Advanced"; badge_color = "#e65100"; badge_bg = "#fff3e0"
                    else:
                        badge_text = "Shortage"; badge_color = "#c62828"; badge_bg = "#ffebee"
                    status_badge = f'<span class="status-pending">⏳ PENDING</span>' if is_pending else f'<span class="status-paid">✅ PAID</span>'
                    note_txt = f" — {t.get('note','')}" if t.get("note") else ""

                    c1, c2, c3 = st.columns([5, 1, 1])
                    with c1:
                        st.markdown(f"""
                        <div style='background:#ffffff;border:1px solid #e0e0e0;border-radius:8px;padding:8px 12px;margin-bottom:4px;'>
                            <span style='color:#0277bd;font-size:12px;'>📅 {t.get('date','')} · 🕐 {t.get('time','')}</span> &nbsp;
                            <span style='background:{badge_bg};color:{badge_color};padding:2px 8px;border-radius:5px;font-size:11px;font-weight:700;'>{badge_text}</span> &nbsp;
                            {status_badge} &nbsp;
                            <b style='color:#1976d2;font-size:15px;'>Rs {t['amount']:,.0f}</b>
                            <span style='color:#666;font-size:12px;'>{note_txt}</span>
                        </div>
                        """, unsafe_allow_html=True)
                    with c2:
                        if is_pending:
                            if st.button("✅ Paid", key=f"paid_{role_type}_{person_name}_{t_id}_{idx}",
                                         use_container_width=True, type="primary"):
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
                            db[sal_key][person_name]["transactions"] = [
                                x for x in db[sal_key][person_name]["transactions"] if x.get("id") != t_id
                            ]
                            save_database(db)
                            st.session_state["success_msg"] = "🗑 Transaction deleted"
                            st.rerun()
            else:
                st.info("Koi transaction nahi. Upar se add karo.")

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None

# ============================================================
# PAGE: DAILY EXPENSE
# ============================================================
def render_daily_expense():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>💵 Daily Expense</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#0277bd;font-weight:500;'>Petrol + Lunch — Date wise</p>", unsafe_allow_html=True)
    st.markdown("---")

    if "petrol_expenses" not in db: db["petrol_expenses"] = []
    if "lunch_expenses" not in db: db["lunch_expenses"] = []

    petrol_list = db.get("petrol_expenses", [])
    lunch_list = db.get("lunch_expenses", [])

    st.markdown("### 🔎 Filter")
    today = date.today()
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        filter_mode = st.selectbox("Filter:",
            ["📅 Aaj (Today)", "📆 This Month", "🗓️ Last 30 Days", "📋 All"],
            key="exp_filter_mode")
    with c2: from_date = st.date_input("From:", value=today - timedelta(days=30), key="exp_from_date")
    with c3: to_date = st.date_input("To:", value=today, key="exp_to_date")

    def date_match(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if filter_mode == "📅 Aaj (Today)": return d == today
        elif filter_mode == "📆 This Month": return d.year == today.year and d.month == today.month
        elif filter_mode == "🗓️ Last 30 Days": return today - timedelta(days=30) <= d <= today
        else: return True

    petrol_filtered = [x for x in petrol_list if date_match(x.get("date", ""))]
    lunch_filtered = [x for x in lunch_list if date_match(x.get("date", ""))]
    total_petrol = sum(float(x.get("amount", 0)) for x in petrol_filtered)
    total_lunch = sum(float(x.get("amount", 0)) for x in lunch_filtered)
    grand = total_petrol + total_lunch

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#0277bd;'>
            ⛽ Petrol: <b>Rs {total_petrol:,.0f}</b> &nbsp;|&nbsp;
            🍽️ Lunch: <b>Rs {total_lunch:,.0f}</b> &nbsp;|&nbsp;
            <b style='color:#0d47a1;'>Grand Total: Rs {grand:,.0f}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown(f"### ⛽ Petrol  <span style='font-size:14px;color:#0277bd;'>(Rs {total_petrol:,.0f})</span>", unsafe_allow_html=True)
        saved_salesmen = db.get("salesmen", [])
        if not saved_salesmen:
            st.info("💡 Pehle **Salesmen** page pe jao aur salesmen add karo.")
        else:
            st.markdown("**➕ Add Petrol Entry**")
            c1, c2, c3 = st.columns([2, 2, 1])
            with c1:
                p_salesman = st.selectbox("Salesman:", options=["-- Select --"] + saved_salesmen,
                                          key="petrol_salesman", label_visibility="collapsed")
            with c2:
                p_amount = st.number_input("Petrol Amount", min_value=0.0, step=50.0,
                                           key="petrol_amount", label_visibility="collapsed", placeholder="Amount")
            with c3:
                if st.button("➕", key="add_petrol", use_container_width=True, type="primary"):
                    if p_salesman == "-- Select --":
                        st.session_state["error_msg"] = "❌ Please select salesman"
                    elif p_amount <= 0:
                        st.session_state["error_msg"] = "❌ Enter petrol amount"
                    else:
                        next_id = max([x.get("id", 0) for x in db["petrol_expenses"]] + [0]) + 1
                        db["petrol_expenses"].append({
                            "id": next_id, "date": datetime.now().strftime("%d-%m-%Y"),
                            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "time": datetime.now().strftime("%H:%M"),
                            "salesman": p_salesman, "amount": float(p_amount),
                        })
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ Petrol Rs {p_amount:,.0f} ({p_salesman}) added"
                        st.rerun()

            st.markdown(f"**📋 Entries ({len(petrol_filtered)})**")
            if not petrol_filtered:
                st.info("Koi entry nahi.")
            else:
                sorted_p = sorted(petrol_filtered, key=lambda x: x.get("created_at", ""), reverse=True)
                for idx, x in enumerate(sorted_p):
                    x_id = x.get("id", idx)
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
                        if st.button("🗑", key=f"del_petrol_{x_id}_{idx}", use_container_width=True):
                            db["petrol_expenses"] = [e for e in db["petrol_expenses"] if e.get("id") != x_id]
                            save_database(db)
                            st.session_state["success_msg"] = "🗑 Petrol entry deleted"
                            st.rerun()

                st.markdown("**📊 Salesman-wise Total (filtered)**")
                per_sm = {}
                for x in petrol_filtered:
                    sm = x.get("salesman", "Unknown")
                    per_sm[sm] = per_sm.get(sm, 0) + float(x.get("amount", 0))
                for sm, amt in sorted(per_sm.items()):
                    st.markdown(f"<div style='padding:4px 8px;font-size:13px;color:#0277bd;'>🧑‍💼 <b>{sm}</b> — Rs {amt:,.0f}</div>", unsafe_allow_html=True)

    with col_right:
        st.markdown(f"### 🍽️ Lunch  <span style='font-size:14px;color:#e65100;'>(Rs {total_lunch:,.0f})</span>", unsafe_allow_html=True)
        st.markdown("**➕ Add Lunch Entry**")
        c1, c2 = st.columns([3, 1])
        with c1:
            l_amount = st.number_input("Lunch Amount", min_value=0.0, step=50.0,
                                       key="lunch_amount", placeholder="Amount", label_visibility="collapsed")
        with c2:
            if st.button("➕", key="add_lunch", use_container_width=True, type="primary"):
                if l_amount <= 0:
                    st.session_state["error_msg"] = "❌ Enter lunch amount"
                else:
                    next_id = max([x.get("id", 0) for x in db["lunch_expenses"]] + [0]) + 1
                    db["lunch_expenses"].append({
                        "id": next_id, "date": datetime.now().strftime("%d-%m-%Y"),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "time": datetime.now().strftime("%H:%M"),
                        "amount": float(l_amount),
                    })
                    save_database(db)
                    st.session_state["success_msg"] = f"✅ Lunch Rs {l_amount:,.0f} added"
                    st.rerun()

        st.markdown(f"**📋 Entries ({len(lunch_filtered)})**")
        if not lunch_filtered:
            st.info("Koi entry nahi.")
        else:
            sorted_l = sorted(lunch_filtered, key=lambda x: x.get("created_at", ""), reverse=True)
            for idx, x in enumerate(sorted_l):
                x_id = x.get("id", idx)
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
                    if st.button("🗑", key=f"del_lunch_{x_id}_{idx}", use_container_width=True):
                        db["lunch_expenses"] = [e for e in db["lunch_expenses"] if e.get("id") != x_id]
                        save_database(db)
                        st.session_state["success_msg"] = "🗑 Lunch entry deleted"
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
        st.markdown(f"<div class='hint-box'>🎁 Active Packages ({len(active_pkgs)}): <b>{names}</b> — Bill export pe best discount apply hoga</div>", unsafe_allow_html=True)
    else:
        st.markdown("<div class='hint-box'>💡 Koi discount package active nahi. '🎁 Discount' page se activate karo.</div>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1: st.text_input("Bill No:", value=str(db["next_bill_no"]), disabled=True, key="dash_bill_no")
    with c2: st.text_input("Date:", value=datetime.now().strftime("%d-%m-%Y"), disabled=True, key="dash_bill_date")

    c1, c2, c3 = st.columns(3)
    with c1:
        shop_name = st.text_input("Shop:", key="shop_name", placeholder="Shop Name")
    with c2:
        saved_bookers = db.get("bookers", [])
        if saved_bookers:
            booker_options = ["-- Select --"] + saved_bookers
            selected_bk = st.selectbox("Order Booker:", options=booker_options, key="order_booker_select")
            st.session_state["order_booker"] = "" if selected_bk == "-- Select --" else selected_bk
        else:
            st.text_input("Order Booker:", key="order_booker", placeholder="Booker")
    with c3:
        saved_salesmen = db.get("salesmen", [])
        if saved_salesmen:
            salesman_options = ["-- Select --"] + saved_salesmen
            selected_sm = st.selectbox("Salesman:", options=salesman_options, key="salesman_select")
            st.session_state["salesman"] = "" if selected_sm == "-- Select --" else selected_sm
        else:
            st.text_input("Salesman:", key="salesman", placeholder="Salesman")

    c1, c2 = st.columns([1, 3])
    with c1:
        search_text = st.text_input("🔍 Search:", key="search_text", placeholder="Type name...")
    with c2:
        search_upper = search_text.strip().upper()
        filtered_names = [p["name"] for p in PRODUCTS if search_upper in p["name"].upper()] if search_upper else PRODUCT_NAMES
        if st.session_state.get("product_sel") and st.session_state["product_sel"] not in filtered_names:
            st.session_state["product_sel"] = ""
        product_sel = st.selectbox("Select Product:", options=[""] + filtered_names, key="product_sel")

    selected_product = None
    if product_sel:
        for p in PRODUCTS:
            if p["name"] == product_sel:
                selected_product = p; break

    if selected_product:
        current_price = get_price(selected_product["code"], selected_product["price"])
        base_price = float(selected_product["price"])
        price_note = f"  (edited — original Rs {base_price:,.0f})" if current_price != base_price else ""
        st.success(f"✅ {selected_product['name']} (Code: {selected_product['code']}) — Rs {current_price:,.0f}{price_note}")
        tp_default = float(current_price)
    else:
        st.error("No Product Selected"); tp_default = 0.0

    if st.session_state["_prev_prod"] != product_sel:
        st.session_state["tp_box"] = tp_default
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
# PAGE: BILLS LIST (with Eye button)
# ============================================================
def render_bills_list():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📋 Bills List</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:#0277bd;font-weight:500;'>Saved bills — 👁️ Eye button se poora bill dekho</p>", unsafe_allow_html=True)
    st.markdown("---")

    if len(db["bills"]) == 0:
        st.info("❌ Koi bill nahi mila. Pehle 'Billing' page pe jaake bill banao.")
        return

    st.markdown("### 🔎 Filter Bills")
    today = date.today()
    c1, c2, c3 = st.columns(3)
    with c1:
        filter_mode = st.selectbox("Filter Mode:",
            ["📅 Aaj Ki Bills (Today)", "📆 Custom Date Range", "🗓️ Specific Date", "📋 All Bills"],
            key="bills_filter_mode")
    with c2: from_date = st.date_input("From Date:", value=today - timedelta(days=7), key="bills_from_date")
    with c3: to_date = st.date_input("To Date:", value=today, key="bills_to_date")

    c1, c2 = st.columns(2)
    with c1: search = st.text_input("🔍 Search (Shop / Product / Booker):", key="bills_search")
    with c2:
        shops_list = sorted(set(b["Shop"] for b in db["bills"] if b["Shop"]))
        shop_filter = st.selectbox("Filter by Shop:", options=["All"] + shops_list, key="shop_filter")

    groups = {}
    for orig_idx, b in enumerate(db["bills"]):
        bdate_str = b.get("Date", "")
        bdate = parse_date(bdate_str)

        if filter_mode == "📅 Aaj Ki Bills (Today)":
            if bdate != today: continue
        elif filter_mode == "🗓️ Specific Date":
            if bdate != from_date: continue
        elif filter_mode == "📆 Custom Date Range":
            if bdate is None or not (from_date <= bdate <= to_date): continue

        if search:
            s = search.upper()
            if not (s in str(b.get("Shop","")).upper() or
                    s in str(b.get("Product","")).upper() or
                    s in str(b.get("Order Booker","")).upper()):
                continue

        if shop_filter != "All" and b.get("Shop", "") != shop_filter:
            continue

        key = (b.get("Shop",""), b.get("Date",""), b.get("Order Booker",""))
        if key not in groups:
            groups[key] = {
                "shop": b.get("Shop",""),
                "date": b.get("Date",""),
                "booker": b.get("Order Booker",""),
                "salesman": b.get("Salesman",""),
                "items": [],
                "orig_indices": [],
                "bill_no": b.get("Bill No",""),
            }
        groups[key]["items"].append({
            "Code": b.get("Code"),
            "Product": b.get("Product"),
            "Boxes": b.get("Boxes"),
            "TP/Box": b.get("TP/Box"),
            "Discount %": b.get("Discount %"),
            "Gross": b.get("Gross"),
            "Net": b.get("Net"),
        })
        groups[key]["orig_indices"].append(orig_idx)

    if not groups:
        st.warning("❌ Is filter ke hisaab se koi bill nahi mila.")
        return

    all_items = []
    for g in groups.values():
        all_items.extend(g["items"])
    total_boxes = sum(int(r.get("Boxes",0)) for r in all_items)
    total_gross = sum(float(r.get("Gross",0)) for r in all_items)
    total_net = sum(float(r.get("Net",0)) for r in all_items)
    unique_shops = len(set(g["shop"] for g in groups.values() if g["shop"]))

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#0277bd;'>
            Bills: <b>{len(groups)}</b> | Boxes: <b>{total_boxes}</b> |
            Shops: <b>{unique_shops}</b> |
            Gross: <b>Rs {total_gross:,.0f}</b> | Net: <b>Rs {total_net:,.0f}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"### 📋 Bills ({len(groups)})")
    st.caption("👇 👁️ = Poora bill dekho | ⬇️ = Excel download | 🗑 = Delete")

    sorted_keys = sorted(groups.keys(), key=lambda k: (parse_date(k[1]) or date.min, k[0]), reverse=True)

    for idx, key in enumerate(sorted_keys):
        g = groups[key]
        shop = g["shop"] or "-"
        date_str = g["date"] or "-"
        booker = g["booker"] or "-"
        salesman = g["salesman"] or "-"
        bill_no = g["bill_no"]
        items = g["items"]
        total_b = sum(int(it.get("Boxes",0)) for it in items)
        total_n = sum(float(it.get("Net",0)) for it in items)

        wkey = f"{shop}_{date_str}_{booker}_{bill_no}_{idx}".replace(" ","_").replace("/","_").replace(":","")
        is_viewing = st.session_state.get("view_bill_key") == wkey

        c1, c2, c3, c4 = st.columns([4, 0.7, 1, 1])
        with c1:
            st.markdown(f"""
            <div class='lf-simple-card'>
                <div class='lf-info'>
                    <div class='lf-line1'>🏪 {shop}</div>
                    <div class='lf-line2'>📅 {date_str} &nbsp;·&nbsp; 👤 {booker} &nbsp;·&nbsp; 🧑‍💼 {salesman}</div>
                </div>
                <div class='lf-boxes'>{total_b}<small>BOXES</small></div>
            </div>
            """, unsafe_allow_html=True)
        with c2:
            eye_icon = "🔽" if is_viewing else "👁️"
            if st.button(eye_icon, key=f"eye_bill_{wkey}", use_container_width=True,
                         help="Poora bill dekho"):
                if is_viewing:
                    st.session_state["view_bill_key"] = None
                else:
                    st.session_state["view_bill_key"] = wkey
                st.rerun()
        with c3:
            if st.button("⬇️ Excel", key=f"dl_bill_{wkey}", use_container_width=True, type="primary"):
                export_single_group_bill(shop, date_str, booker, salesman, items, bill_no)
                st.rerun()
        with c4:
            if st.button("🗑 Delete", key=f"del_bill_{wkey}", use_container_width=True):
                st.session_state["confirm_delete_group"] = g["orig_indices"]
                st.session_state["_confirm_group_label"] = f"{shop} | {date_str} | {booker}"

        # ---- Full Bill View (Eye clicked) ----
        if is_viewing:
            st.markdown(f"""
            <div class='full-bill-box'>
                <div class='full-bill-title'>👁️ {shop} — Full Bill View</div>
                <div class='full-bill-meta'>
                    📅 {date_str} &nbsp;·&nbsp; 👤 Booker: <b>{booker}</b> &nbsp;·&nbsp; 🧑‍💼 Salesman: <b>{salesman}</b> &nbsp;·&nbsp; 🧾 Bill No: <b>{bill_no}</b>
                </div>
            </div>
            """, unsafe_allow_html=True)

            df_items = pd.DataFrame(items)
            st.dataframe(df_items, use_container_width=True, hide_index=True)

            # Big totals
            pkg_pct_preview, pkg_name_preview, _tier = get_package_discount_pct(total_n)
            after_disc = total_n - (total_n * pkg_pct_preview / 100)
            saved = total_n - after_disc

            cc1, cc2, cc3 = st.columns(3)
            with cc1:
                st.markdown(f"<div class='metric-card'><h3>TOTAL BOXES</h3><h1>{total_b}</h1></div>", unsafe_allow_html=True)
            with cc2:
                st.markdown(f"<div class='metric-card'><h3>NET TOTAL</h3><h1>Rs {total_n:,.0f}</h1></div>", unsafe_allow_html=True)
            with cc3:
                if pkg_pct_preview > 0:
                    st.markdown(f"<div class='metric-card'><h3>AFTER DISC ({pkg_pct_preview}%)</h3><h1>Rs {after_disc:,.0f}</h1></div>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<div class='metric-card'><h3>AFTER DISC</h3><h1>Rs {total_n:,.0f}</h1></div>", unsafe_allow_html=True)

            if pkg_pct_preview > 0:
                st.markdown(f"<div class='hint-box'>🎁 Package: <b>{pkg_name_preview}</b> | {pkg_pct_preview}% discount | Saved: <b>Rs {saved:,.0f}</b></div>", unsafe_allow_html=True)

            close_c1, close_c2 = st.columns([1, 4])
            with close_c1:
                if st.button("❌ Close View", key=f"close_view_bill_{wkey}", use_container_width=True):
                    st.session_state["view_bill_key"] = None
                    st.rerun()
            with close_c2:
                if st.button("⬇️ Download Excel", key=f"dl_from_view_{wkey}", use_container_width=True, type="primary"):
                    export_single_group_bill(shop, date_str, booker, salesman, items, bill_no)
                    st.rerun()

            st.markdown("---")

    if st.session_state.get("confirm_delete_group"):
        label = st.session_state.get("_confirm_group_label", "this bill group")
        st.warning(f"⚠️ Kya aap **{label}** ka bill delete karna chahte hain? Ye undo nahi hoga.")
        cc1, cc2 = st.columns(2)
        with cc1:
            if st.button("✅ Haan, Delete Kar Do", key="confirm_group_yes", use_container_width=True, type="primary"):
                idxs = set(st.session_state.get("confirm_delete_group", []))
                if idxs:
                    db["bills"] = [b for i, b in enumerate(db["bills"]) if i not in idxs]
                    save_database(db)
                    st.session_state["success_msg"] = f"🗑 {len(idxs)} bill line(s) deleted"
                st.session_state["confirm_delete_group"] = None
                st.session_state["_confirm_group_label"] = ""
                st.session_state["view_bill_key"] = None
                st.rerun()
        with cc2:
            if st.button("❌ Cancel", key="confirm_group_no", use_container_width=True):
                st.session_state["confirm_delete_group"] = None
                st.session_state["_confirm_group_label"] = ""
                st.rerun()

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None

    show_auto_download()

# ============================================================
# PAGE: LOAD FORM (with Eye button)
# ============================================================
def render_load_form():
    st.markdown(f"<h1 style='color:#1976d2 !important;'>📦 Load Forms</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#0277bd;font-weight:500;'>Saved load forms — 👁️ Eye button se poora load form dekho</p>", unsafe_allow_html=True)
    st.markdown("---")

    load_forms = db.get("load_forms", [])
    if not load_forms:
        st.info("❌ Abhi tak koi load form save nahi hua. Billing page pe '📦 Export Load Form' click karo.")
        return

    st.markdown("### 🔎 Filter Load Forms")
    today = date.today()
    c1, c2, c3 = st.columns(3)
    with c1:
        filter_mode = st.selectbox("Filter Mode:",
            ["📅 Aaj Ke Load Forms (Today)", "📆 Custom Date Range", "🗓️ Specific Date", "📋 All Load Forms"],
            key="lf_filter_mode")
    with c2: from_date = st.date_input("From Date:", value=today - timedelta(days=7), key="lf_from_date")
    with c3: to_date = st.date_input("To Date:", value=today, key="lf_to_date")

    all_booker_names = sorted(set(lf["booker"] for lf in load_forms if lf.get("booker")))
    booker_filter = st.selectbox("Filter by Booker:", options=["All"] + all_booker_names, key="lf_booker_filter")

    filtered_lfs = []
    for lf in load_forms:
        lf_date = parse_date(lf.get("date", ""))
        if lf_date is None: continue
        if filter_mode == "📅 Aaj Ke Load Forms (Today)":
            if lf_date != today: continue
        elif filter_mode == "🗓️ Specific Date":
            if lf_date != from_date: continue
        elif filter_mode == "📆 Custom Date Range":
            if not (from_date <= lf_date <= to_date): continue
        if booker_filter != "All" and lf.get("booker") != booker_filter: continue
        filtered_lfs.append(lf)

    if not filtered_lfs:
        st.warning("❌ Is filter ke hisaab se koi load form nahi mila.")
        return

    total_forms = len(filtered_lfs)
    total_boxes_all = sum(lf.get("total_boxes", 0) for lf in filtered_lfs)

    st.markdown(f"""
    <div class='summary-box'>
        <b style='color:#1976d2;font-size:16px;'>📊 Summary</b><br>
        <span style='color:#0277bd;'>
            Load Forms: <b>{total_forms}</b> | Total Boxes: <b>{total_boxes_all}</b>
        </span>
    </div>
    """, unsafe_allow_html=True)

    if st.button("⬇️ Download All Filtered Load Forms (Excel)", key="lf_dl_all", use_container_width=True):
        export_all_filtered_load_forms(filtered_lfs); st.rerun()

    st.markdown("---")
    st.markdown(f"### 📋 Load Forms ({len(filtered_lfs)})")

    sorted_lfs = sorted(filtered_lfs, key=lambda x: x.get("created_at", ""), reverse=True)
    for idx, lf in enumerate(sorted_lfs):
        lf_id = lf.get("id", idx)
        booker = lf.get("booker", "Unknown")
        date_str = lf.get("date", ""); time_str = lf.get("time", "")
        total_boxes = lf.get("total_boxes", 0)
        items = lf.get("items", [])

        bill_items = []
        for it in items:
            code = it.get("Code", "")
            base_price = 0.0
            for p in PRODUCTS:
                if str(p["code"]) == str(code):
                    base_price = get_price(p["code"], p["price"])
                    break
            boxes = int(it.get("Boxes", 0))
            gross = boxes * base_price
            bill_items.append({
                "Code": code, "Product": it.get("Product",""),
                "Boxes": boxes, "TP/Box": base_price,
                "Discount %": 0, "Gross": gross, "Net": gross,
            })

        wkey = f"lf_{lf_id}_{idx}"
        is_viewing = st.session_state.get("view_lf_key") == wkey

        c1, c2, c3, c4 = st.columns([4, 0.7, 1, 1])
        with c1:
            st.markdown(f"""
            <div class='lf-simple-card'>
                <div class='lf-info'>
                    <div class='lf-line1'>👤 {booker}</div>
                    <div class='lf-line2'>📅 {date_str} &nbsp;·&nbsp; 🕐 {time_str} &nbsp;·&nbsp; 📦 {len(items)} products</div>
                </div>
                <div class='lf-boxes'>{total_boxes}<small>BOXES</small></div>
            </div>
            """, unsafe_allow_html=True)
        with c2:
            eye_icon = "🔽" if is_viewing else "👁️"
            if st.button(eye_icon, key=f"eye_lf_{wkey}", use_container_width=True,
                         help="Poora load form dekho"):
                if is_viewing:
                    st.session_state["view_lf_key"] = None
                else:
                    st.session_state["view_lf_key"] = wkey
                st.rerun()
        with c3:
            if st.button("⬇️ Excel", key=f"dl_lf_{wkey}", use_container_width=True, type="primary"):
                export_single_group_bill(booker + " Load", date_str, booker, "", bill_items, lf_id)
                st.rerun()
        with c4:
            if st.button("🗑 Delete", key=f"del_lf_{wkey}", use_container_width=True):
                db["load_forms"] = [x for x in db["load_forms"] if x.get("id") != lf_id]
                save_database(db)
                st.session_state["success_msg"] = f"🗑 Load Form deleted"
                st.session_state["view_lf_key"] = None
                st.rerun()

        # ---- Full Load Form View (Eye clicked) ----
        if is_viewing:
            st.markdown(f"""
            <div class='full-bill-box'>
                <div class='full-bill-title'>👁️ Load Form — {booker}</div>
                <div class='full-bill-meta'>
                    📅 {date_str} &nbsp;·&nbsp; 🕐 {time_str} &nbsp;·&nbsp; 📦 {len(items)} products &nbsp;·&nbsp; Total: <b>{total_boxes} boxes</b>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if bill_items:
                df_items = pd.DataFrame([{
                    "Code": it["Code"], "Product": it["Product"],
                    "Boxes": it["Boxes"], "TP/Box": it["TP/Box"], "Net": it["Net"]
                } for it in bill_items])
                st.dataframe(df_items, use_container_width=True, hide_index=True)

            total_net = sum(float(it.get("Net", 0)) for it in bill_items)
            pkg_pct_preview, pkg_name_preview, _tier = get_package_discount_pct(total_net)
            after_disc = total_net - (total_net * pkg_pct_preview / 100)

            cc1, cc2, cc3 = st.columns(3)
            with cc1:
                st.markdown(f"<div class='metric-card'><h3>PRODUCTS</h3><h1>{len(items)}</h1></div>", unsafe_allow_html=True)
            with cc2:
                st.markdown(f"<div class='metric-card'><h3>TOTAL BOXES</h3><h1>{total_boxes}</h1></div>", unsafe_allow_html=True)
            with cc3:
                st.markdown(f"<div class='metric-card'><h3>NET TOTAL</h3><h1>Rs {total_net:,.0f}</h1></div>", unsafe_allow_html=True)

            close_c1, close_c2 = st.columns([1, 4])
            with close_c1:
                if st.button("❌ Close View", key=f"close_view_lf_{wkey}", use_container_width=True):
                    st.session_state["view_lf_key"] = None
                    st.rerun()
            with close_c2:
                if st.button("⬇️ Download Excel", key=f"dl_from_view_lf_{wkey}", use_container_width=True, type="primary"):
                    export_single_group_bill(booker + " Load", date_str, booker, "", bill_items, lf_id)
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
    product_sel = st.session_state.get("product_sel", "")
    boxes_v = st.session_state.get("boxes", 0)
    tp_v = st.session_state.get("tp_box", 0.0)
    disc_v = st.session_state.get("discount", 0.0)

    if not product_sel:
        st.session_state["error_msg"] = "❌ Please Select Product from dropdown"; return
    if boxes_v <= 0:
        st.session_state["error_msg"] = "❌ Enter Boxes"; return

    sel = next((p for p in PRODUCTS if p["name"] == product_sel), None)
    if not sel:
        st.session_state["error_msg"] = "❌ Invalid Product"; return

    gross_v = boxes_v * tp_v
    net_v = gross_v - (gross_v * disc_v / 100)

    bill = {
        "Bill No": db["next_bill_no"], "Date": datetime.now().strftime("%d-%m-%Y"),
        "Shop": st.session_state.get("shop_name", "").strip(),
        "Order Booker": st.session_state.get("order_booker", "").strip(),
        "Salesman": st.session_state.get("salesman", "").strip(),
        "Delivery Man": "",
        "Code": sel["code"], "Product": sel["name"],
        "Boxes": boxes_v, "TP/Box": tp_v, "Discount %": disc_v,
        "Gross": gross_v, "Net": net_v
    }

    db["bills"].append(bill); save_database(db)
    st.session_state["last_bill_no"] = db["next_bill_no"]
    st.session_state["success_msg"] = f"✅ Bill Added | Bill No: {db['next_bill_no']} | Total: {len(db['bills'])}"
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    st.session_state["_prev_prod"] = None
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0

def refresh_callback():
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    st.session_state["_prev_prod"] = None
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["success_msg"] = "✅ Ready For Next Product"

def export_bill_callback():
    db = st.session_state.database
    if len(db["bills"]) == 0:
        st.session_state["error_msg"] = "❌ No Bills Found"; return

    shop = st.session_state.get("shop_name", "").strip() or "Bill"
    for ch in ['\\','/',':','*','?','"','<','>','|']: shop = shop.replace(ch, "")

    shop_bills = [b for b in db["bills"] if b["Shop"].strip() == shop]
    if not shop_bills:
        st.session_state["error_msg"] = "❌ Is shop ke liye koi bill nahi mila"; return

    bill_total_net = sum(float(b.get("Net", 0)) for b in shop_bills)
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
    worksheet.write("D3","Booker",header); worksheet.write("E3", st.session_state.get("order_booker", ""), cell_center)
    worksheet.write("G3","Bill No",header)
    last_bill = st.session_state.get("last_bill_no") or (db["next_bill_no"] - 1)
    worksheet.write("H3", last_bill, cell_center)
    worksheet.write("I3","Date",header); worksheet.write("J3", datetime.now().strftime("%d-%m-%Y"), cell_center)

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
    for bill in shop_bills:
        b_net = float(bill.get("Net", 0))
        b_gross = float(bill.get("Gross", 0))
        b_boxes = int(bill.get("Boxes", 0))
        after_net = b_net - (b_net * pkg_pct / 100)
        saved = b_net - after_net

        worksheet.write(row, 0, bill["Product"], cell_left)
        worksheet.write(row, 1, bill["Code"], cell_center)
        worksheet.write(row, 2, b_boxes, cell_center)
        worksheet.write(row, 3, bill.get("TP/Box", 0), cell_center)
        worksheet.write(row, 4, b_gross, cell_center)
        worksheet.write(row, 5, bill.get("Discount %", 0), cell_center)
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
    st.session_state["download_file"] = (f"{shop}.xlsx", output.getvalue())
    db["next_bill_no"] += 1; save_database(db)
    st.session_state["last_bill_no"] = None
    if pkg_pct > 0:
        st.session_state["success_msg"] = f"✅ Bill Exported | {pkg_name} {tier_label} | {pkg_pct}% | Net: Rs {after_disc_total:,.0f} (Saved Rs {saved_total:,.0f})"
    else:
        st.session_state["success_msg"] = f"✅ Bill Exported | Next Bill No: {db['next_bill_no']}"

def export_load_form_from_billing_callback():
    db = st.session_state.database
    booker = st.session_state.get("order_booker", "").strip()
    if not booker:
        st.session_state["error_msg"] = "❌ Please select Order Booker first"; return
    booker_bills = [b for b in db["bills"] if b["Order Booker"].strip() == booker]
    if not booker_bills:
        st.session_state["error_msg"] = f"❌ No bills found for booker: {booker}"; return
    summary = {}
    for b in booker_bills:
        code = b["Code"]
        if code not in summary: summary[code] = {"Code": code, "Product": b["Product"], "Boxes": 0}
        summary[code]["Boxes"] += b["Boxes"]
    items = list(summary.values()); total_boxes = sum(it["Boxes"] for it in items)
    if "load_forms" not in db: db["load_forms"] = []
    next_id = 1
    if db["load_forms"]: next_id = max(lf.get("id", 0) for lf in db["load_forms"]) + 1
    lf_record = {
        "id": next_id, "date": datetime.now().strftime("%d-%m-%Y"),
        "time": datetime.now().strftime("%H:%M"),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "booker": booker, "items": items,
        "total_boxes": total_boxes, "total_products": len(items)
    }
    db["load_forms"].append(lf_record); save_database(db)
    export_load_form_for_booker(booker, booker_bills)
    st.session_state["success_msg"] = f"✅ Load Form #{next_id} saved & exported | {len(items)} products, {total_boxes} boxes"

def export_load_form_for_booker(booker, booker_bills=None):
    db = st.session_state.database
    if booker_bills is None:
        booker_bills = [b for b in db["bills"] if b["Order Booker"].strip() == booker]
    if not booker_bills:
        st.session_state["error_msg"] = "❌ No Bills Found for this Booker"; return
    summary = {}
    for bill in booker_bills:
        code = bill["Code"]
        if code not in summary: summary[code] = {"Product": bill["Product"], "Boxes": 0}
        summary[code]["Boxes"] += bill["Boxes"]
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Load Form")
    title = workbook.add_format({"bold":True, "font_size":16, "align":"center", "border":2})
    header = workbook.add_format({"bold":True, "font_size":12, "bg_color":"#BBDEFB", "align":"center", "border":2})
    cell_left = workbook.add_format({"font_size":14, "border":1, "align":"left"})
    cell_center = workbook.add_format({"font_size":14, "border":1, "align":"center"})
    total = workbook.add_format({"bold":True, "font_size":14, "bg_color":"#FFF2CC", "align":"center", "border":2})
    worksheet.set_column("A:A", 47.86); worksheet.set_column("B:B", 12.71)
    worksheet.merge_range("A1:B1", COMPANY_NAME, title)
    worksheet.write("A3","Order Booker",header); worksheet.write("B3",booker,cell_center)
    worksheet.write("A5","Product",header); worksheet.write("B5","Boxes",header)
    row = 5; total_boxes = 0
    for item in summary.values():
        worksheet.write(row, 0, item["Product"], cell_left)
        worksheet.write(row, 1, item["Boxes"], cell_center)
        total_boxes += item["Boxes"]; row += 1
    worksheet.write(row, 0, "TOTAL", total); worksheet.write(row, 1, total_boxes, total)
    workbook.close(); output.seek(0)
    st.session_state["download_file"] = (f"{booker}_Load_Form.xlsx", output.getvalue())

def export_all_filtered_load_forms(filtered_lfs):
    if not filtered_lfs:
        st.session_state["error_msg"] = "❌ No Load Forms to export"; return
    output = BytesIO(); wb = xlsxwriter.Workbook(output, {'in_memory': True})
    title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center", "border": 2, "bg_color": "#BBDEFB"})
    header_fmt = wb.add_format({"bold": True, "bg_color": "#E3F2FD", "border": 1, "align": "center"})
    cell_fmt = wb.add_format({"border": 1}); cell_center = wb.add_format({"border": 1, "align": "center"})
    total_fmt = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1, "align": "center"})
    for lf in filtered_lfs:
        booker = lf.get("booker", "Unknown"); date_str = lf.get("date", ""); time_str = lf.get("time", "")
        sheet_name = f"{booker}_{date_str}".replace("/", "-")[:31] or f"LF_{lf.get('id')}"
        base_name = sheet_name; counter = 1; existing = wb.sheetnames
        while sheet_name in existing:
            sheet_name = f"{base_name[:28]}_{counter}"; counter += 1
        ws = wb.add_worksheet(sheet_name)
        ws.set_column("A:A", 45); ws.set_column("B:B", 10); ws.set_column("C:C", 10)
        ws.merge_range("A1:C1", f"{COMPANY_NAME} - Load Form", title_fmt)
        ws.write("A3", "Booker", header_fmt); ws.write("B3", booker, cell_fmt)
        ws.write("A4", "Date", header_fmt); ws.write("B4", f"{date_str} {time_str}", cell_fmt)
        ws.write("A6", "Product", header_fmt); ws.write("B6", "Code", header_fmt); ws.write("C6", "Boxes", header_fmt)
        row = 6; total_boxes = 0
        for it in lf.get("items", []):
            ws.write(row, 0, it["Product"], cell_fmt)
            ws.write(row, 1, it["Code"], cell_center)
            ws.write(row, 2, it["Boxes"], cell_center)
            total_boxes += it["Boxes"]; row += 1
        ws.write(row, 0, "TOTAL", total_fmt); ws.write(row, 1, "", total_fmt); ws.write(row, 2, total_boxes, total_fmt)
    wb.close(); output.seek(0)
    fname = f"Load_Forms_{datetime.now().strftime('%d-%m-%Y_%H%M')}.xlsx"
    st.session_state["download_file"] = (fname, output.getvalue())

def refresh_load_form_callback():
    db = st.session_state.database
    booker = st.session_state.get("order_booker", "").strip()
    if not booker:
        st.session_state["error_msg"] = "❌ Please Enter Order Booker"; return
    db["bills"] = [b for b in db["bills"] if b["Order Booker"].strip() != booker]
    save_database(db)
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    st.session_state["_prev_prod"] = None
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["success_msg"] = f"✅ Load Form Cleared | Booker: {booker}"

# ============================================================
# RENDER SELECTED PAGE
# ============================================================
if st.session_state["page"] == "📊 Dashboard":
    render_dashboard()
elif st.session_state["page"] == "🧾 Billing":
    render_billing()
elif st.session_state["page"] == "🛒 All Products":
    render_all_products()
elif st.session_state["page"] == "🎁 Discount":
    render_discount()
elif st.session_state["page"] == "👤 Bookers":
    render_bookers()
elif st.session_state["page"] == "💰 Bookers Salary":
    render_salaries("bookers")
elif st.session_state["page"] == "🧑‍💼 Salesmen":
    render_salesmen()
elif st.session_state["page"] == "💰 Salesmen Salary":
    render_salaries("salesmen")
elif st.session_state["page"] == "💵 Daily Expense":
    render_daily_expense()
elif st.session_state["page"] == "📋 Bills List":
    render_bills_list()
elif st.session_state["page"] == "📦 Load Form":
    render_load_form()
