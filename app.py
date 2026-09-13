# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# Dashboard + Sidebar + Light Blue/Green Theme + Fixed Collapse
# ============================================================

import os
import json
import pandas as pd
from datetime import datetime
import streamlit as st
import xlsxwriter
from io import BytesIO

st.set_page_config(
    page_title="AL-BARAKAH ENTERPRISES",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# LIGHT BLUE / GREEN THEME CSS
# ============================================================
st.markdown("""
<style>
    /* Header visible but transparent - KEEP the toggle button alive */
    header[data-testid="stHeader"] {
        background: transparent !important;
    }
    [data-testid="stToolbar"] { display: none !important; }
    [data-testid="stDecoration"] { display: none !important; }
    [data-testid="stStatusWidget"] { display: none !important; }
    [data-testid="manage-app-button"] { display: none !important; }
    [data-testid="stAppDeployButton"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    .stAppDeployButton { display: none !important; }

    /* Hide "Manage app" from Streamlit Cloud */
    iframe[title="streamlit_cloud_status"] {
        display: none !important;
    }
    div[class*="manageApp"] {
        display: none !important;
    }
    button[kind="header"] {
        display: none !important;
    }

    /* Cover "Manage app" corner overlay */
    .stApp::after {
        content: "";
        position: fixed;
        bottom: 0;
        right: 0;
        width: 240px;
        height: 60px;
        background: linear-gradient(135deg, #e0f7fa 0%, #e8f5e9 100%);
        z-index: 2147483646;
        pointer-events: none;
    }

    /* Force sidebar collapse/expand button to be visible */
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"] {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
        position: fixed !important;
        top: 10px !important;
        left: 10px !important;
        z-index: 999999 !important;
        background: linear-gradient(135deg, #4caf50 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
        box-shadow: 0 3px 10px rgba(76,175,80,0.5) !important;
        padding: 6px 12px !important;
        cursor: pointer !important;
    }
    [data-testid="stSidebarCollapsedControl"] button,
    [data-testid="collapsedControl"] button {
        background: transparent !important;
        color: #ffffff !important;
        border: none !important;
        font-size: 20px !important;
        padding: 4px 8px !important;
    }
    [data-testid="stSidebarCollapsedControl"]:hover,
    [data-testid="collapsedControl"]:hover {
        background: linear-gradient(135deg, #388e3c 0%, #00897b 100%) !important;
        box-shadow: 0 5px 14px rgba(76,175,80,0.7) !important;
        transform: scale(1.05) !important;
    }
    [data-testid="stSidebarCollapsedControl"] svg,
    [data-testid="collapsedControl"] svg {
        fill: #ffffff !important;
        color: #ffffff !important;
    }

    /* Full width layout */
    .block-container {
        padding-top: 1.5rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        padding-bottom: 1rem !important;
        max-width: 100% !important;
    }

    /* Main background */
    .stApp {
        background: linear-gradient(135deg, #e0f7fa 0%, #e8f5e9 100%) !important;
        color: #1a1a1a !important;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #c8e6c9 0%, #b2ebf2 100%) !important;
        border-right: 2px solid #a5d6a7 !important;
    }
    section[data-testid="stSidebar"] * {
        color: #1a1a1a !important;
    }
    [data-testid="stSidebarCollapseButton"] button {
        background: linear-gradient(135deg, #4caf50 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
    }
    [data-testid="stSidebarCollapseButton"] button:hover {
        background: linear-gradient(135deg, #388e3c 0%, #00897b 100%) !important;
    }

    /* Text colors */
    h1, h2, h3, h4, h5, h6, p, span, label, div {
        color: #1a1a1a !important;
    }

    /* Inputs */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div > div {
        background-color: #ffffff !important;
        color: #1a1a1a !important;
        border: 2px solid #b2dfdb !important;
        border-radius: 8px !important;
    }
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus {
        border-color: #4caf50 !important;
    }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #4caf50 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        box-shadow: 0 2px 6px rgba(76,175,80,0.25) !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #388e3c 0%, #00897b 100%) !important;
        box-shadow: 0 4px 12px rgba(76,175,80,0.45) !important;
    }
    .stButton > button p {
        color: #ffffff !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #0288d1 0%, #26a69a 100%) !important;
        color: #ffffff !important;
        border: none !important;
        font-weight: bold !important;
        border-radius: 8px !important;
    }
    .stDownloadButton > button p {
        color: #ffffff !important;
    }

    /* Radio nav */
    .stRadio > div { background-color: transparent !important; }
    .stRadio label {
        color: #1a1a1a !important;
        font-size: 15px !important;
        padding: 8px 10px !important;
        border-radius: 8px !important;
        cursor: pointer !important;
        font-weight: 500 !important;
    }
    div[role="radiogroup"] > label {
        background-color: #ffffff !important;
        border: 1px solid #b2dfdb !important;
        border-radius: 10px !important;
        margin-bottom: 8px !important;
        transition: all 0.2s ease !important;
    }
    div[role="radiogroup"] > label:hover {
        background-color: #e0f2f1 !important;
        border-color: #4caf50 !important;
        transform: translateX(3px) !important;
    }

    /* Metric cards */
    .metric-card {
        background: #ffffff;
        border: 2px solid #a5d6a7;
        border-radius: 14px;
        padding: 22px;
        text-align: center;
        box-shadow: 0 3px 10px rgba(76,175,80,0.15);
        transition: all 0.3s ease;
    }
    .metric-card:hover {
        box-shadow: 0 6px 18px rgba(76,175,80,0.3);
        transform: translateY(-3px);
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

    /* Dataframe */
    .stDataFrame {
        background-color: #ffffff !important;
        border-radius: 10px !important;
        border: 1px solid #b2dfdb !important;
    }

    .stAlert {
        border-radius: 10px !important;
    }

    hr {
        border-color: #a5d6a7 !important;
        opacity: 0.6 !important;
    }

    .streamlit-expanderHeader {
        background-color: #ffffff !important;
        border-radius: 8px !important;
        color: #1a1a1a !important;
    }
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
                return json.load(f)
        except Exception:
            pass
    return {"next_bill_no": 1, "bills": []}

if "database" not in st.session_state:
    st.session_state.database = load_database()
if "_prev_prod" not in st.session_state:
    st.session_state["_prev_prod"] = None
if "last_bill_no" not in st.session_state:
    st.session_state["last_bill_no"] = None
if "download_file" not in st.session_state:
    st.session_state["download_file"] = None
if "page" not in st.session_state:
    st.session_state["page"] = "📊 Dashboard"

db = st.session_state.database

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
        ["📊 Dashboard", "🧾 Billing", "📋 Bills List", "📦 Load Form"],
        key="page_selector",
        label_visibility="collapsed"
    )
    st.session_state["page"] = page

    st.markdown("---")
    st.markdown(f"""
    <div style='padding:10px; color:#00695c !important; font-size:12px;'>
        <p>📅 {datetime.now().strftime('%d-%m-%Y')}</p>
        <p>📦 Products: {len(PRODUCTS)}</p>
        <p>🧾 Total Bills: {len(db['bills'])}</p>
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
        st.markdown(f"""
        <div class='metric-card'>
            <h3>TOTAL BILLS</h3>
            <h1>{total_bills}</h1>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class='metric-card'>
            <h3>TOTAL BOXES</h3>
            <h1>{total_boxes}</h1>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class='metric-card'>
            <h3>UNIQUE SHOPS</h3>
            <h1>{unique_shops}</h1>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class='metric-card'>
            <h3>ORDER BOOKERS</h3>
            <h1>{unique_bookers}</h1>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
        <div class='metric-card'>
            <h3>TOTAL GROSS AMOUNT</h3>
            <h1>Rs {total_gross:,.0f}</h1>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class='metric-card'>
            <h3>TOTAL NET AMOUNT</h3>
            <h1>Rs {total_net:,.0f}</h1>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)

    st.markdown("### 🕐 Recent Bills (Last 5)")
    if bills:
        recent = bills[-5:][::-1]
        df = pd.DataFrame(recent)
        st.dataframe(df, use_container_width=True, hide_index=True)
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
        df_top = pd.DataFrame(top_products)
        st.dataframe(df_top, use_container_width=True, hide_index=True)

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
    order_booker = st.text_input("Order Booker:", key="order_booker", placeholder="Enter Order Booker")
    salesman = st.text_input("Salesman:", key="salesman", placeholder="Enter Salesman")
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
        st.button("📦 Export Load Form", key="btn_load", on_click=export_load_form_callback, use_container_width=True)
    with c3:
        st.button("🗑 Refresh Load Form", key="btn_load_refresh", on_click=refresh_load_form_callback, use_container_width=True)

    if st.session_state.get("download_file"):
        fname, fdata = st.session_state["download_file"]
        st.download_button(
            label=f"⬇️ Download {fname}",
            data=fdata,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_btn",
            use_container_width=True,
        )

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

    df = pd.DataFrame(db["bills"])

    c1, c2 = st.columns([2, 1])
    with c1:
        search = st.text_input("🔍 Search (Shop / Product / Booker):", key="bills_search")
    with c2:
        shop_filter = st.selectbox("Filter by Shop:", options=["All"] + sorted(set(b["Shop"] for b in db["bills"] if b["Shop"])), key="shop_filter")

    filtered_df = df.copy()
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

    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
    st.caption(f"Showing {len(filtered_df)} of {len(df)} bills")

# ============================================================
# PAGE: LOAD FORM
# ============================================================
def render_load_form():
    st.markdown(f"<h1 style='color:#2e7d32 !important;'>📦 Load Form</h1>", unsafe_allow_html=True)
    st.markdown("<p style='color:#00695c;font-weight:500;'>Order Booker ke hisaab se load form dekho</p>", unsafe_allow_html=True)
    st.markdown("---")

    bookers = sorted(set(b["Order Booker"] for b in db["bills"] if b["Order Booker"]))
    if not bookers:
        st.info("❌ Koi order booker nahi mila. Pehle 'Billing' page pe bill banao.")
        return

    selected_booker = st.selectbox("Select Order Booker:", options=bookers, key="lf_booker")

    booker_bills = [b for b in db["bills"] if b["Order Booker"].strip() == selected_booker]

    if not booker_bills:
        st.warning("Is booker ke liye koi bill nahi hai.")
        return

    summary = {}
    for b in booker_bills:
        code = b["Code"]
        if code not in summary:
            summary[code] = {"Code": code, "Product": b["Product"], "Boxes": 0}
        summary[code]["Boxes"] += b["Boxes"]

    df = pd.DataFrame(list(summary.values()))
    total_boxes = df["Boxes"].sum()

    st.markdown(f"**Order Booker:** {selected_booker}")
    st.markdown(f"**Total Products:** {len(summary)} | **Total Boxes:** {total_boxes}")
    st.markdown("---")
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    if st.button("📦 Export Load Form (Excel)", key="lf_export", use_container_width=True, type="primary"):
        export_load_form_for_booker(selected_booker)

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


def export_load_form_callback():
    booker = st.session_state.get("order_booker", "").strip()
    if not booker:
        st.session_state["error_msg"] = "❌ Please Enter Order Booker"
        return
    export_load_form_for_booker(booker)


def export_load_form_for_booker(booker):
    db = st.session_state.database
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
    st.session_state["success_msg"] = f"✅ Load Form Exported | Products: {len(summary)} | Boxes: {total_boxes}"


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
elif st.session_state["page"] == "📋 Bills List":
    render_bills_list()
elif st.session_state["page"] == "📦 Load Form":
    render_load_form()
