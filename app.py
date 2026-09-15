# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# Premium Lamp Login (Query-Param toggle) + Multi-User
# ============================================================

import os
import json
import random
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
# SESSION STATE INIT
# ============================================================
if "light_on" not in st.session_state:
    st.session_state["light_on"] = False
if "auth_error" not in st.session_state:
    st.session_state["auth_error"] = ""
if "auth_tab" not in st.session_state:
    st.session_state["auth_tab"] = "login"

# ============================================================
# HANDLE ?lamp=on/off QUERY PARAM (turn light ON/OFF)
# ============================================================
def handle_lamp_query():
    try:
        qp = st.query_params
    except Exception:
        try:
            qp = st.experimental_get_query_params()
        except Exception:
            qp = {}

    lamp_val = None
    if isinstance(qp, dict):
        v = qp.get("lamp")
        if isinstance(v, list):
            lamp_val = v[0] if v else None
        else:
            lamp_val = v

    if lamp_val:
        if lamp_val == "on":
            st.session_state["light_on"] = True
        elif lamp_val == "off":
            st.session_state["light_on"] = False
        # clear query params
        try:
            st.query_params.clear()
        except Exception:
            try:
                st.experimental_set_query_params()
            except Exception:
                pass
        st.rerun()

handle_lamp_query()

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
# LAMP SCENE (rendered when NOT logged in)
# ============================================================
def render_lamp_scene():
    light_on = st.session_state.get("light_on", False)
    lit = "lit" if light_on else ""
    lit_hint = "lit" if light_on else ""
    cord_href = "?lamp=off" if light_on else "?lamp=on"
    hint_text = "☀ LIGHT ON — Fill form and pull cord again" if light_on else "▼ PULL THE CORD ▼"

    # Generate dust particles (server side)
    dust_html = ""
    random.seed(7)
    for _ in range(28):
        left = random.randint(8, 92)
        top = random.randint(0, 60)
        dur = round(random.uniform(5.5, 11.5), 2)
        delay = round(random.uniform(0, 6), 2)
        size = random.choice([1, 1, 2, 2, 3])
        dust_html += (
            f'<div class="dust-particle" style="left:{left}%;top:{top}%;'
            f'width:{size}px;height:{size}px;animation-duration:{dur}s;'
            f'animation-delay:{delay}s;"></div>'
        )

    html = f"""
<style>
    /* Base dark-grey (not pure black) */
    html, body {{
        background: #0d0d14 !important;
    }}
    .stApp, [data-testid="stAppViewContainer"], .main {{
        background: transparent !important;
    }}
    .block-container {{
        padding-top: 0 !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
        padding-bottom: 0 !important;
        max-width: 100% !important;
    }}

    /* Hide Streamlit chrome during auth */
    header[data-testid="stHeader"],
    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    #MainMenu, footer,
    section[data-testid="stSidebar"],
    .stAppDeployButton,
    [data-testid="manage-app-button"] {{
        display: none !important;
    }}

    /* ============ LAMP SCENE ============ */
    .lamp-scene {{
        position: fixed;
        top: 0; left: 0;
        width: 100vw;
        height: 100vh;
        z-index: 0;
        pointer-events: none;
        overflow: hidden;
        background: radial-gradient(ellipse at 50% -10%, #1a1a24 0%, #0e0e16 45%, #06060a 100%);
        transition: background 1.4s ease;
    }}
    .lamp-scene.lit {{
        background: radial-gradient(ellipse at 50% 22%, #33241a 0%, #1a1108 42%, #0a0808 78%);
    }}

    /* Ceiling plate */
    .lamp-ceiling {{
        position: absolute;
        top: 0; left: 50%;
        transform: translateX(-50%);
        width: 90px; height: 14px;
        background: linear-gradient(180deg, #322a20 0%, #120e08 100%);
        border-radius: 0 0 10px 10px;
        box-shadow:
            0 3px 12px rgba(0,0,0,0.9),
            inset 0 1px 0 rgba(200,160,100,0.2);
    }}

    /* Chain (links) */
    .lamp-chain {{
        position: absolute;
        top: 12px; left: 50%;
        transform: translateX(-50%);
        width: 8px; height: 125px;
        background: repeating-linear-gradient(
            180deg,
            #2c261c 0px, #2c261c 2px,
            #14100a 2px, #14100a 5px,
            #4a3e2c 5px, #4a3e2c 6px,
            #2c261c 6px, #2c261c 9px,
            #14100a 9px, #14100a 12px
        );
        box-shadow: 0 0 8px rgba(0,0,0,0.85), 1px 0 2px rgba(0,0,0,0.6);
    }}

    /* Metallic dome shade */
    .lamp-shade {{
        position: absolute;
        top: 133px; left: 50%;
        transform: translateX(-50%);
        width: 250px; height: 120px;
        border-radius: 125px 125px 28px 28px / 120px 120px 30px 30px;
        background:
            radial-gradient(ellipse at 50% 0%, #5a4832 0%, #3a2a1a 28%, #1c1408 68%, #0a0806 100%),
            linear-gradient(180deg, #2a1e10 0%, #0a0604 100%);
        box-shadow:
            inset 0 -26px 40px rgba(0,0,0,0.96),
            inset 0 6px 16px rgba(180,130,70,0.35),
            inset 0 2px 0 rgba(230,190,130,0.25),
            0 10px 30px rgba(0,0,0,0.92);
        z-index: 3;
        transition: box-shadow 1s ease;
    }}
    .lamp-scene.lit .lamp-shade {{
        box-shadow:
            inset 0 -26px 40px rgba(0,0,0,0.96),
            inset 0 6px 16px rgba(220,170,100,0.55),
            inset 0 2px 0 rgba(255,230,170,0.45),
            0 10px 30px rgba(0,0,0,0.92),
            0 0 90px 22px rgba(255,190,80,0.18);
    }}
    .lamp-shade::before {{
        content: '';
        position: absolute;
        top: 8px; left: 50%;
        transform: translateX(-50%);
        width: 72%; height: 32px;
        border-radius: 50%;
        background: radial-gradient(ellipse, rgba(200,150,80,0.35) 0%, transparent 70%);
        pointer-events: none;
    }}
    .lamp-shade::after {{
        content: '';
        position: absolute;
        bottom: 0; left: 0;
        width: 100%; height: 22px;
        border-radius: 0 0 28px 28px / 0 0 30px 30px;
        background: linear-gradient(180deg, transparent 0%, #030202 100%);
    }}

    /* Bulb */
    .lamp-bulb {{
        position: absolute;
        top: 90px; left: 50%;
        transform: translateX(-50%);
        width: 58px; height: 52px;
        border-radius: 50% 50% 42% 42% / 55% 55% 45% 45%;
        background: radial-gradient(circle at 50% 35%, #2a2a2a 0%, #0a0a0a 75%);
        box-shadow: inset 0 -8px 15px rgba(0,0,0,0.9);
        z-index: 4;
        transition: all 1s cubic-bezier(0.2, 0.9, 0.3, 1);
    }}
    .lamp-scene.lit .lamp-bulb {{
        background: radial-gradient(circle at 50% 35%, #ffffff 0%, #fff8e1 22%, #ffd54f 55%, #ff9800 100%);
        box-shadow:
            0 0 30px 12px rgba(255, 235, 160, 0.95),
            0 0 75px 30px rgba(255, 200, 80, 0.6),
            0 0 160px 65px rgba(255, 170, 40, 0.32),
            inset 0 0 14px rgba(255, 255, 220, 0.9);
    }}

    /* Light beam */
    .lamp-beam {{
        position: absolute;
        top: 240px; left: 50%;
        transform: translateX(-50%);
        width: 1100px; height: 920px;
        background: radial-gradient(
            ellipse at 50% 0%,
            rgba(255, 235, 160, 0.48) 0%,
            rgba(255, 220, 130, 0.26) 20%,
            rgba(255, 200, 80, 0.11) 45%,
            rgba(255, 180, 60, 0.04) 68%,
            transparent 84%
        );
        opacity: 0;
        transition: opacity 1.4s ease;
        pointer-events: none;
        filter: blur(12px);
        z-index: 1;
    }}
    .lamp-scene.lit .lamp-beam {{ opacity: 1; }}

    /* Dust */
    .dust-container {{
        position: absolute;
        top: 250px; left: 50%;
        transform: translateX(-50%);
        width: 750px; height: 750px;
        opacity: 0;
        transition: opacity 1.8s ease 0.4s;
        pointer-events: none;
        z-index: 2;
    }}
    .lamp-scene.lit .dust-container {{ opacity: 1; }}
    .dust-particle {{
        position: absolute;
        background: radial-gradient(circle, #fff8e1 0%, #ffd54f 60%, transparent 100%);
        border-radius: 50%;
        box-shadow: 0 0 6px #ffd54f;
        animation: drift linear infinite;
    }}
    @keyframes drift {{
        0%   {{ transform: translate(0,0); opacity: 0; }}
        15%  {{ opacity: 1; }}
        85%  {{ opacity: 1; }}
        100% {{ transform: translate(45px, 250px); opacity: 0; }}
    }}

    /* Pull cord */
    .lamp-cord {{
        position: absolute;
        top: 145px; left: 50%;
        transform: translateX(58px);
        width: 30px;
        height: 250px;
        z-index: 30;
        pointer-events: auto;
        cursor: pointer;
        text-decoration: none;
        display: block;
        transform-origin: top center;
        transition: transform 0.35s cubic-bezier(0.34, 1.56, 0.64, 1);
    }}
    .lamp-cord:hover {{
        transform: translateX(58px) scaleY(1.06);
    }}
    .lamp-cord:active {{
        transform: translateX(58px) scaleY(1.45);
    }}
    .cord-line {{
        position: absolute;
        top: 0; left: 50%;
        transform: translateX(-50%);
        width: 3px; height: 205px;
        background: linear-gradient(180deg, #3a3a3a 0%, #d0d0d0 45%, #3a3a3a 100%);
        box-shadow: 0 0 4px rgba(0,0,0,0.85);
        border-radius: 2px;
    }}
    .cord-bead {{
        position: absolute;
        bottom: 12px; left: 50%;
        transform: translateX(-50%);
        width: 22px; height: 30px;
        border-radius: 50% 50% 45% 45% / 60% 60% 40% 40%;
        background:
            radial-gradient(circle at 35% 25%, #ffe0a0 0%, #c89050 30%, #7a5020 65%, #2a1a08 100%);
        box-shadow:
            0 5px 14px rgba(0,0,0,0.9),
            0 0 22px 4px rgba(255, 190, 90, 0.55),
            inset 0 -3px 6px rgba(0,0,0,0.55);
        transition: all 0.3s ease;
    }}
    .cord-bead::after {{
        content: '';
        position: absolute;
        top: 7px; left: 6px;
        width: 7px; height: 7px;
        border-radius: 50%;
        background: radial-gradient(circle, rgba(255,245,215,0.95) 0%, transparent 70%);
    }}
    .lamp-cord:hover .cord-bead {{
        box-shadow:
            0 5px 14px rgba(0,0,0,0.9),
            0 0 36px 10px rgba(255, 200, 100, 1),
            inset 0 -3px 6px rgba(0,0,0,0.55);
        transform: translateX(-50%) scale(1.18);
    }}

    /* Bottom hint */
    .bottom-hint {{
        position: fixed;
        bottom: 42px; left: 50%;
        transform: translateX(-50%);
        color: #6e5e42;
        font-size: 12px;
        letter-spacing: 5px;
        text-transform: uppercase;
        font-weight: 300;
        z-index: 100;
        pointer-events: none;
        animation: pulse 2.8s ease-in-out infinite;
        text-align: center;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    }}
    .bottom-hint.lit {{
        color: #b8945e;
        letter-spacing: 3px;
        font-size: 11px;
        animation: none;
    }}
    @keyframes pulse {{
        0%,100% {{ opacity: 0.4; }}
        50%     {{ opacity: 1; }}
    }}
</style>

<div class="lamp-scene {lit}">
    <div class="lamp-ceiling"></div>
    <div class="lamp-chain"></div>
    <div class="lamp-shade"></div>
    <div class="lamp-bulb"></div>
    <div class="lamp-beam"></div>
    <div class="dust-container">{dust_html}</div>
    <a href="{cord_href}" target="_top" class="lamp-cord" title="Pull me">
        <div class="cord-line"></div>
        <div class="cord-bead"></div>
    </a>
</div>
<div class="bottom-hint {lit_hint}">{hint_text}</div>
"""
    st.markdown(html, unsafe_allow_html=True)

# ============================================================
# LOGIN / SIGNUP FORM STYLING (dark gold, appears over lamp scene)
# ============================================================
AUTH_FORM_CSS = """
<style>
    /* Push Streamlit content so form sits below lamp beam */
    .block-container {
        padding-top: 320px !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        padding-bottom: 4rem !important;
        max-width: 100% !important;
        position: relative;
        z-index: 50;
    }

    /* Streamlit form styling */
    [data-testid="stForm"] {
        background: rgba(15, 12, 9, 0.72) !important;
        border: 1px solid rgba(255, 200, 100, 0.28) !important;
        border-radius: 18px !important;
        padding: 22px 24px !important;
        backdrop-filter: blur(14px) !important;
        -webkit-backdrop-filter: blur(14px) !important;
        box-shadow:
            0 14px 44px rgba(0,0,0,0.7),
            inset 0 1px 0 rgba(255, 220, 150, 0.15),
            0 0 90px rgba(255, 190, 80, 0.1) !important;
    }

    /* Labels */
    [data-testid="stForm"] label,
    [data-testid="stForm"] [data-testid="stWidgetLabel"] label,
    [data-testid="stForm"] [data-testid="stWidgetLabel"] p {
        color: #b8945e !important;
        font-size: 11px !important;
        letter-spacing: 2px !important;
        text-transform: uppercase !important;
        font-weight: 700 !important;
        margin-bottom: 4px !important;
    }

    /* Inputs */
    [data-testid="stForm"] input {
        background: rgba(0,0,0,0.55) !important;
        color: #fff8e1 !important;
        -webkit-text-fill-color: #fff8e1 !important;
        border: 1px solid rgba(255, 200, 100, 0.2) !important;
        border-radius: 10px !important;
        padding: 10px 14px !important;
        font-size: 14px !important;
        caret-color: #ffd54f !important;
    }
    [data-testid="stForm"] input:focus {
        border-color: rgba(255, 200, 100, 0.6) !important;
        box-shadow: 0 0 0 3px rgba(255, 200, 100, 0.12), 0 0 22px rgba(255, 180, 50, 0.18) !important;
        outline: none !important;
    }
    [data-testid="stForm"] input::placeholder {
        color: #6a5a3a !important;
        -webkit-text-fill-color: #6a5a3a !important;
        font-style: italic !important;
    }

    /* Primary submit button */
    [data-testid="stForm"] button[kind="primary"],
    [data-testid="stForm"] button[kind="primaryFormSubmit"] {
        background: linear-gradient(135deg, #ffb347 0%, #ff9800 100%) !important;
        color: #1a0e05 !important;
        font-weight: 800 !important;
        letter-spacing: 2px !important;
        text-transform: uppercase !important;
        border-radius: 10px !important;
        border: none !important;
        box-shadow: 0 4px 16px rgba(255,150,50,0.45), inset 0 1px 0 rgba(255,255,255,0.3) !important;
        padding: 10px 20px !important;
    }
    [data-testid="stForm"] button[kind="primary"]:hover,
    [data-testid="stForm"] button[kind="primaryFormSubmit"]:hover {
        background: linear-gradient(135deg, #ffc266 0%, #ffa726 100%) !important;
        box-shadow: 0 6px 22px rgba(255,150,50,0.65) !important;
    }

    /* Secondary buttons */
    [data-testid="stForm"] button[kind="secondary"] {
        background: transparent !important;
        color: #b8945e !important;
        border: 1px solid rgba(255, 200, 100, 0.3) !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        letter-spacing: 1px !important;
    }
    [data-testid="stForm"] button[kind="secondary"]:hover {
        background: rgba(255, 200, 100, 0.08) !important;
        border-color: rgba(255, 200, 100, 0.6) !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background: rgba(0,0,0,0.5) !important;
        border-radius: 12px !important;
        padding: 4px !important;
        gap: 4px !important;
        border: 1px solid rgba(255,200,100,0.12) !important;
    }
    .stTabs [data-baseweb="tab"] {
        background: transparent !important;
        color: #8a7a5a !important;
        font-weight: 700 !important;
        letter-spacing: 1.5px !important;
        border-radius: 9px !important;
        padding: 8px 16px !important;
    }
    .stTabs [data-baseweb="tab"] p {
        color: #8a7a5a !important;
        font-weight: 700 !important;
        font-size: 12px !important;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, rgba(255,200,80,0.2), rgba(255,150,40,0.1)) !important;
        box-shadow: 0 2px 10px rgba(255,180,50,0.25), inset 0 0 0 1px rgba(255,200,100,0.3) !important;
    }
    .stTabs [aria-selected="true"] p {
        color: #ffd54f !important;
    }
    .stTabs [data-baseweb="tab-highlight"] { display: none !important; }
    .stTabs [data-baseweb="tab-border"] { display: none !important; }

    /* Auth headings */
    .auth-title {
        text-align: center;
        color: #ffd54f;
        font-size: 22px;
        font-weight: 900;
        letter-spacing: 4px;
        margin-bottom: 2px;
        text-shadow: 0 0 22px rgba(255, 200, 80, 0.55);
    }
    .auth-sub {
        text-align: center;
        color: #a8845a;
        font-size: 10px;
        letter-spacing: 4px;
        font-weight: 600;
        margin-bottom: 16px;
    }
    .auth-msg {
        text-align: center;
        color: #ff9a5a;
        font-size: 12px;
        margin-bottom: 8px;
        font-weight: 600;
        letter-spacing: 0.5px;
    }

    /* Alerts inside form */
    [data-testid="stForm"] .stAlert,
    .stAlert { border-radius: 10px !important; }
</style>
"""

# ============================================================
# AUTH PAGE
# ============================================================
if not st.session_state.get("logged_in_user"):
    # Render lamp scene
    render_lamp_scene()

    if st.session_state.get("light_on"):
        st.markdown(AUTH_FORM_CSS, unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1, 1.25, 1])
        with col2:
            st.markdown(f"<div class='auth-title'>AL-BARAKAH</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='auth-sub'>ENTERPRISES</div>", unsafe_allow_html=True)

            if st.session_state.get("auth_error"):
                st.markdown(f"<div class='auth-msg'>⚠️ {st.session_state['auth_error']}</div>", unsafe_allow_html=True)
                st.session_state["auth_error"] = ""

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
                            st.session_state["auth_error"] = "Please fill all fields"
                            st.rerun()
                        elif u not in users:
                            st.session_state["auth_error"] = "Username not found. Please signup first."
                            st.rerun()
                        elif users[u].get("password_hash") != hash_password(login_pass):
                            st.session_state["auth_error"] = "Incorrect password"
                            st.rerun()
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
                            st.session_state["auth_error"] = "Username required"
                            st.rerun()
                        elif len(u) < 3:
                            st.session_state["auth_error"] = "Username min 3 characters"
                            st.rerun()
                        elif len(u) > 20:
                            st.session_state["auth_error"] = "Username max 20 characters"
                            st.rerun()
                        elif len(su_pass) < 4:
                            st.session_state["auth_error"] = "Password min 4 characters"
                            st.rerun()
                        elif su_pass != su_pass2:
                            st.session_state["auth_error"] = "Passwords do not match"
                            st.rerun()
                        elif u in users:
                            st.session_state["auth_error"] = f"Username '{u}' already taken"
                            st.rerun()
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
        with open(user_data_file(CURRENT_USER), "w", encoding="utf-8") as f:
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

for k, v in [("_prev_prod", None), ("last_bill_no", None), ("download_file", None),
             ("_dl_counter", 0), ("page", "📊 Dashboard"),
             ("view_bill_key", None), ("view_lf_key", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

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
# GLOBAL APP CSS (light blue theme, for logged-in app)
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

# Sidebar toggle
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
    st.session_state["success_msg"] = f"✅ Excel ready | {shop}" + (f" | {pkg_pct}% discount" if pkg_pct > 0 else "")

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
        ["📊 Dashboard", "🧾 Billing", "🛒 All Products", "🎁 Discount",
         "👤 Bookers", "💰 Bookers Salary",
         "🧑‍💼 Salesmen", "💰 Salesmen Salary",
         "💵 Daily Expense", "📋 Bills List", "📦 Load Form"],
        key="page_selector", label_visibility="collapsed"
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
            new_name = st.text_input(f"Package {pid} Name", value=pkg.get("name", f"Package {pid}"),
                                     key=f"pkgname_{pid}", label_visibility="collapsed")

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
                ta1 = st.number_input("Min", value=float(pkg.get("tier1_amount", 0) or 0), min_value=0.0, step=50.0, key=f"t1a_{pid}", label_visibility="collapsed")
                st.caption("Min Rs")
            with tc2:
                tp1 = st.number_input("%", value=float(pkg.get("tier1_pct", 0) or 0), min_value=0.0, max_value=100.0, step=0.5, key=f"t1p_{pid}", label_visibility="collapsed")
                st.caption("Disc %")
        with t2:
            st.markdown("<div style='font-size:11px;font-weight:700;color:#1976d2;'>Tier 2</div>", unsafe_allow_html=True)
            tc1, tc2 = st.columns(2)
            with tc1:
                ta2 = st.number_input("Min", value=float(pkg.get("tier2_amount", 0) or 0), min_value=0.0, step=50.0, key=f"t2a_{pid}", label_visibility="collapsed")
                st.caption("Min Rs")
            with tc2:
                tp2 = st.number_input("%", value=float(pkg.get("tier2_pct", 0) or 0), min_value=0.0, max_value=100.0, step=0.5, key=f"t2p_{pid}", label_visibility="collapsed")
                st.caption("Disc %")
        with t3:
            st.markdown("<div style='font-size:11px;font-weight:700;color:#1976d2;'>Tier 3</div>", unsafe_allow_html=True)
            tc1, tc2 = st.columns(2)
            with tc1:
                ta3 = st.number_input("Min", value=float(pkg.get("tier3_amount", 0) or 0), min_value=0.0, step=50.0, key=f"t3a_{pid}", label_visibility="collapsed")
                st.caption("Min Rs")
            with tc2:
                tp3 = st.number_input("%", value=float(pkg.get("tier3_pct", 0) or 0), min_value=0.0, max_value=100.0, step=0.5, key=f"t3p_{pid}", label_visibility="collapsed")
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

    shown = [p for p in PRODUCTS
             if (not search_upper or search_upper in p["name"].upper())
             and (not show_only_edited or str(p["code"]) in edited_prices)]

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
                        st.session_state["success_msg"] = f"✅ {p['name']}: price original"
                    else:
                        db.setdefault("product_prices", {})[code] = float(new_price)
                        save_database(db)
                        st.session_state["success_msg"] = f"✅ {p['name']}: Rs {new_price:,.0f}"
                    st.rerun()
            with sub2:
                if is_edited:
                    if st.button("↩️ Reset", key=f"reset_price_{code}", use_container_width=True):
                        db.get("product_prices", {}).pop(code, None)
                        save_database(db)
                        st.session_state["success_msg"] = f"↩️ {p['name']}: reset"
                        st.rerun()

    if st.session_state.get("success_msg"):
        st.success(st.session_state["success_msg"]); st.session_state["success_msg"] = None

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
                    st.session_state["success_msg"] = f"🗑 Deleted"
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
