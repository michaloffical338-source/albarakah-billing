# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# (Google Sheets Cloud Database + Premium Animations)
# 100% Error Free - Final Working Code
# ============================================================

import os
import re
import json
import hashlib
import pandas as pd
from datetime import datetime, date, timedelta
import streamlit as st
import streamlit.components.v1 as components
import xlsxwriter
from io import BytesIO
from streamlit_gsheets import GSheetsConnection

st.set_page_config(
    page_title="AL-BARAKAH ENTERPRISES",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# GOOGLE SHEETS CONNECTION
# ============================================================
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception as e:
    st.error(f"Google Sheets connection failed. Check Secrets. Error: {e}")
    st.stop()

# ============================================================
# CONSTANTS & AUTH
# ============================================================
COMPANY_NAME = "AL-BARAKAH ENTERPRISES"
PRODUCTS = []
PASSWORD_SALT = st.secrets.get("PASSWORD_SALT", "albarakah_2026_secret_salt")
AUTH_SHEET = "Auth"
BOOKERS_SHEET = "Bookers"

def hash_password(password):
    return hashlib.sha256((PASSWORD_SALT + password).encode()).hexdigest()

def sanitize_username(u):
    return "".join(ch for ch in u if ch.isalnum() or ch in "_-.").lower()

def _ensure_sheet(sheet_name, columns):
    try:
        df = conn.read(worksheet=sheet_name, ttl=0)
        if df.empty or not set(columns).issubset(df.columns):
            df = pd.DataFrame(columns=columns)
        return df
    except Exception:
        try:
            df = pd.DataFrame(columns=columns)
            conn.create(worksheet=sheet_name, data=df)
            return df
        except Exception:
            return pd.DataFrame(columns=columns)

def load_users():
    cols = ["username", "display_name", "password_hash", "created_at"]
    df = _ensure_sheet(AUTH_SHEET, cols)
    users = {}
    if df.empty: return users
    for _, row in df.iterrows():
        u = str(row.get("username", "")).strip()
        if not u: continue
        users[u] = {
            "username": u, "display_name": str(row.get("display_name", u)),
            "password_hash": str(row.get("password_hash", "")),
            "created_at": str(row.get("created_at", "")),
        }
    return users

def save_users(users):
    rows = [{"username": uname, "display_name": info.get("display_name", uname), "password_hash": info.get("password_hash", ""), "created_at": info.get("created_at", "")} for uname, info in users.items()]
    df = pd.DataFrame(rows, columns=["username", "display_name", "password_hash", "created_at"])
    try: conn.update(worksheet=AUTH_SHEET, data=df)
    except Exception:
        try: conn.create(worksheet=AUTH_SHEET, data=df)
        except Exception as e: st.session_state["cloud_error"] = f"⚠️ Auth save failed: {e}"

def add_or_update_user(username, display_name, password_hash, created_at=None):
    users = load_users()
    users[username] = {"username": username, "display_name": display_name, "password_hash": password_hash, "created_at": created_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    save_users(users)

def load_bookers():
    cols = ["booker_username", "display_name", "admin", "password_hash", "created_at"]
    df = _ensure_sheet(BOOKERS_SHEET, cols)
    registry = {}
    if df.empty: return registry
    for _, row in df.iterrows():
        u = str(row.get("booker_username", "")).strip()
        if not u: continue
        registry[u] = {"username": u, "display_name": str(row.get("display_name", u)), "admin": str(row.get("admin", "")), "password_hash": str(row.get("password_hash", "")), "created_at": str(row.get("created_at", ""))}
    return registry

def save_bookers(registry):
    rows = [{"booker_username": uname, "display_name": info.get("display_name", uname), "admin": info.get("admin", ""), "password_hash": info.get("password_hash", ""), "created_at": info.get("created_at", "")} for uname, info in registry.items()]
    df = pd.DataFrame(rows, columns=["booker_username", "display_name", "admin", "password_hash", "created_at"])
    try: conn.update(worksheet=BOOKERS_SHEET, data=df)
    except Exception:
        try: conn.create(worksheet=BOOKERS_SHEET, data=df)
        except Exception as e: st.session_state["cloud_error"] = f"⚠️ Bookers save failed: {e}"

# ============================================================
# DEFAULTS
# ============================================================
def default_discount_packages():
    return [{"id": i, "name": f"Package {i}", "tier1_amount": 1200.0, "tier1_pct": 1.0, "tier2_amount": 2100.0, "tier2_pct": 2.0, "tier3_amount": 3100.0, "tier3_pct": 3.0, "active": True} for i in range(1, 5)]

def default_wholesaler_packages():
    return [{"id": "ws1", "name": "WS1", "discount_pct": 3.0}, {"id": "ws2", "name": "WS2", "discount_pct": 5.0}, {"id": "ws3", "name": "WS3", "discount_pct": 7.0}]

def default_blank_db():
    return {"next_bill_no": 1, "bills": [], "bookers": [], "salesmen": [], "load_forms": [], "bookers_salaries": {}, "salesmen_salaries": {}, "product_prices": {}, "petrol_expenses": [], "lunch_expenses": [], "misc_expenses": [], "booker_returns": [], "discount_packages": default_discount_packages(), "custom_products": [], "dsr_forms": [], "credit_bills": [], "deleted_products": [], "next_credit_id": 1, "wholesaler_packages": default_wholesaler_packages()}

# ============================================================
# CLOUD DB PER ADMIN
# ============================================================
def _worksheet_name_for_admin(username):
    return f"Data_{sanitize_username(username) or 'default'}"

def load_cloud_db(admin_username):
    sheet = _worksheet_name_for_admin(admin_username)
    try:
        df = conn.read(worksheet=sheet, ttl=0)
        if df.empty or "data" not in df.columns: return None
        data_str = df["data"].iloc[0]
        if not data_str or pd.isna(data_str): return None
        return json.loads(data_str)
    except Exception: return None

def save_cloud_db(admin_username, db_data):
    sheet = _worksheet_name_for_admin(admin_username)
    try:
        json_str = json.dumps(db_data, ensure_ascii=False)
        df = pd.DataFrame([{"data": json_str}])
        try: conn.update(worksheet=sheet, data=df)
        except Exception: conn.create(worksheet=sheet, data=df)
        return True
    except Exception as e:
        st.session_state["cloud_error"] = f"⚠️ Cloud Save Error: {e}"
        return False

def _draft_sheet(admin_username):
    return f"Drafts_{sanitize_username(admin_username) or 'default'}"

def load_all_drafts(admin_username):
    try:
        df = conn.read(worksheet=_draft_sheet(admin_username), ttl=0)
        if df.empty or "data" not in df.columns: return {}
        raw = df["data"].iloc[0]
        return json.loads(raw) if raw and not pd.isna(raw) else {}
    except Exception: return {}

def save_all_drafts(admin_username, drafts):
    try:
        df = pd.DataFrame([{"data": json.dumps(drafts, ensure_ascii=False)}])
        try: conn.update(worksheet=_draft_sheet(admin_username), data=df)
        except Exception: conn.create(worksheet=_draft_sheet(admin_username), data=df)
    except Exception: pass

def get_booker_draft(admin_username, booker_name):
    return load_all_drafts(admin_username).get(sanitize_username(booker_name), {})

def save_booker_draft(admin_username, booker_name, shop, items, salesman=""):
    drafts = load_all_drafts(admin_username)
    drafts[sanitize_username(booker_name)] = {"shop": shop, "items": items, "salesman": salesman, "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    save_all_drafts(admin_username, drafts)

def clear_booker_draft(admin_username, booker_name):
    drafts = load_all_drafts(admin_username)
    if sanitize_username(booker_name) in drafts:
        drafts.pop(sanitize_username(booker_name))
        save_all_drafts(admin_username, drafts)

# ============================================================
# DB SHAPE
# ============================================================
def _ensure_db_shape(db):
    if not db: db = default_blank_db()
    for k in ["bookers", "salesmen", "load_forms", "petrol_expenses", "lunch_expenses", "misc_expenses", "booker_returns", "custom_products", "dsr_forms", "credit_bills", "deleted_products", "bills"]:
        if k not in db or db[k] is None: db[k] = []
    for k in ["bookers_salaries", "salesmen_salaries", "product_prices"]:
        if k not in db or db[k] is None: db[k] = {}
    if not db.get("discount_packages"): db["discount_packages"] = default_discount_packages()
    else:
        for p in db["discount_packages"]:
            p.setdefault("tier3_amount", 0.0)
            p.setdefault("tier3_pct", 0.0)
    if not db.get("wholesaler_packages"): db["wholesaler_packages"] = default_wholesaler_packages()
    if "next_bill_no" not in db or not isinstance(db["next_bill_no"], int): db["next_bill_no"] = 1
    if "next_credit_id" not in db: db["next_credit_id"] = max([c.get("id", 0) for c in db.get("credit_bills", [])] + [0]) + 1
    for b in db.get("bookers", []): db["bookers_salaries"].setdefault(b, {"base_salary": 0, "transactions": []})
    for s in db.get("salesmen", []): db["salesmen_salaries"].setdefault(s, {"base_salary": 0, "transactions": []})
    return db

def load_database_for(admin_username):
    data = load_cloud_db(admin_username)
    if not data:
        data = default_blank_db()
        save_cloud_db(admin_username, data)
    return _ensure_db_shape(data)

def parse_date(dstr):
    try: return datetime.strptime(dstr, "%d-%m-%Y").date()
    except Exception: return None

# ============================================================
# GLOBAL CSS & ANIMATIONS
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
    @keyframes gradientBG { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
    .stApp { background: linear-gradient(-45deg, #e3f2fd, #bbdefb, #e0f7fa, #e8eaf6) !important; background-size: 400% 400% !important; animation: gradientBG 15s ease infinite !important; overflow-x: hidden; }
    @keyframes waveMotion { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
    .login-bg { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: linear-gradient(270deg, #0f172a, #1e3a8a, #2563eb, #1e3a8a); background-size: 400% 400%; animation: waveMotion 15s ease infinite; z-index: -1; }
    @keyframes floatCard { 0% { transform: translateY(0px); } 50% { transform: translateY(-8px); } 100% { transform: translateY(0px); } }
    @keyframes goldPulse { 0% { opacity: 0.70; transform: scale(1); } 50% { opacity: 1; transform: scale(1.07); } 100% { opacity: 0.70; transform: scale(1); } }
    @keyframes cardShine { 0% { background-position: -300% 0; } 100% { background-position: 300% 0; } }
    div[data-testid="stColumn"]:has(.login-card-marker), div[data-testid="column"]:has(.login-card-marker) { position: relative !important; padding: 38px 34px 26px 34px !important; margin-top: 4vh !important; background: transparent !important; overflow: visible !important; animation: floatCard 6s ease-in-out infinite; isolation: isolate; }
    div[data-testid="stColumn"]:has(.login-card-marker)::before, div[data-testid="column"]:has(.login-card-marker)::before { content: ''; position: absolute; inset: -55px; border-radius: 60px; background: radial-gradient(circle at 50% 45%, rgba(255, 223, 100, 0.85) 0%, rgba(255, 200, 60, 0.55) 35%, rgba(255, 190, 40, 0.0) 70%); filter: blur(38px); z-index: -2; animation: goldPulse 4s ease-in-out infinite; pointer-events: none; }
    div[data-testid="stColumn"]:has(.login-card-marker)::after, div[data-testid="column"]:has(.login-card-marker)::after { content: ''; position: absolute; inset: 0; border-radius: 26px; background: linear-gradient(115deg, rgba(255,255,255,0) 35%, rgba(255,255,255,0.55) 50%, rgba(255,255,255,0) 65%), linear-gradient(160deg, #fffdf3 0%, #fff7d6 55%, #ffefb0 100%); background-size: 300% 100%, 100% 100%; animation: cardShine 7s linear infinite; border: 1.5px solid rgba(218, 165, 32, 0.55); box-shadow: 0 25px 60px rgba(0, 0, 0, 0.35), 0 0 0 5px rgba(255, 215, 0, 0.12), inset 0 1px 0 rgba(255, 255, 255, 0.9), inset 0 -12px 30px rgba(218, 165, 32, 0.12); z-index: -1; pointer-events: none; }
    .login-title { text-align: center; font-size: 36px; font-weight: 800; background: linear-gradient(135deg, #a8740a 0%, #daa520 50%, #f5c518 100%); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: 2px; margin: 0 0 4px 0; line-height: 1.15; }
    .login-subtitle { text-align: center; font-size: 12px; color: #8b6508; letter-spacing: 4px; font-weight: 700; margin-bottom: 22px; text-transform: uppercase; }
    .login-divider { height: 2px; border-radius: 2px; margin: 0 auto 18px auto; width: 70px; background: linear-gradient(90deg, transparent, #daa520, transparent); }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTabs [data-baseweb="tab-list"] { gap: 8px; background-color: rgba(255, 255, 255, 0.65); padding: 6px; border-radius: 12px; margin-bottom: 18px; border: 1px solid rgba(218, 165, 32, 0.3); }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTabs [data-baseweb="tab"] { border-radius: 8px; padding: 10px 15px; color: #8b6508 !important; font-weight: 700 !important; border: none; }
    div[data-testid="stColumn"]:has(.login-card-marker) .stTabs [aria-selected="true"] { background-color: #ffffff !important; color: #b8860b !important; box-shadow: 0 2px 8px rgba(218, 165, 32, 0.25) !important; }
    .bubbles { position: fixed; bottom: -100px; left: 0; width: 100%; height: 100%; z-index: 0; pointer-events: none; margin: 0; padding: 0; }
    .bubbles li { position: absolute; list-style: none; display: block; width: 40px; height: 40px; bottom: -160px; animation: square 25s infinite linear; border-radius: 50%; }
    @keyframes square { 0% { transform: translateY(0) rotate(0deg); opacity: 1; border-radius: 50%; } 100% { transform: translateY(-1000px) rotate(720deg); opacity: 0; border-radius: 50%; } }
    .dash-hero { background: linear-gradient(-45deg, #1e3a8a, #2563eb, #3b82f6, #8b5cf6, #2563eb) !important; background-size: 300% 300% !important; animation: gradientBG 8s ease infinite !important; border-radius: 20px; padding: 26px 32px; margin-bottom: 22px; color: #fff; box-shadow: 0 12px 32px rgba(37, 99, 235, 0.35); z-index: 2; }
    .kpi-card { background: #ffffff; border-radius: 16px; padding: 20px 22px; border: 1px solid rgba(37,99,235,0.12); box-shadow: 0 4px 16px rgba(37,99,235,0.08); display: flex; flex-direction: column; gap: 12px; transition: all 0.3s; }
    .kpi-card:hover { transform: translateY(-8px) scale(1.02) !important; box-shadow: 0 15px 30px rgba(37,99,235,0.25) !important; }
    .block-container { animation: fadeInPage 0.6s ease-out !important; padding-top: 1rem !important; }
    @keyframes fadeInPage { 0% { opacity: 0; transform: translateY(15px); } 100% { opacity: 1; transform: translateY(0); } }
    html, body, .stApp, .stApp p, .stApp span, .stApp div, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6, .stApp label, .stApp li, .stApp a { color: #000000; }
    [data-testid="stToolbar"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }
</style>
<ul class="bubbles">
    <li style="left:10%; background:rgba(33,150,243,0.15);"></li>
    <li style="left:20%; width:30px; height:30px; background:rgba(76,175,80,0.15);"></li>
    <li style="left:35%; width:100px; height:100px; background:rgba(255,152,0,0.15);"></li>
    <li style="left:50%; width:60px; height:60px; background:rgba(156,39,176,0.15);"></li>
</ul>
""", unsafe_allow_html=True)

# ============================================================
# AUTH SCREEN
# ============================================================
def render_auth_page():
    st.markdown('<div class="login-bg"></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("""
        <div class="login-card-marker"></div>
        <div class="login-title">AL-BARAKAH</div>
        <div class="login-subtitle">ENTERPRISES 2026</div>
        <div class="login-divider"></div>
        """, unsafe_allow_html=True)
        
        tab1, tab2, tab3 = st.tabs(["🔐 Admin", "👤 Booker", "📝 Signup"])
        with tab1:
            login_user = st.text_input("Username", key="login_username")
            login_pass = st.text_input("Password", key="login_password", type="password")
            if st.button("🔓 Login", key="btn_login", use_container_width=True, type="primary"):
                users = load_users()
                uname = sanitize_username(login_user.strip())
                if not uname or not login_pass: st.error("❌ Both fields required")
                elif uname not in users: st.error("❌ Username not found")
                elif users[uname].get("password_hash") != hash_password(login_pass): st.error("❌ Incorrect password")
                else:
                    st.session_state["logged_in_user"] = uname
                    st.session_state["display_name"] = users[uname].get("display_name", uname)
                    st.session_state["role"] = "admin"
                    st.session_state["page"] = "📊 Dashboard"
                    st.session_state.pop("database", None)
                    st.rerun()

        with tab2:
            bk_user = st.text_input("Booker Name", key="bk_login_user")
            bk_pass = st.text_input("Password", key="bk_login_pass", type="password")
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
                        st.error("❌ You have been removed from bookers list.")
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
                        d = get_booker_draft(admin_uname, st.session_state["booker_name"])
                        if d:
                            st.session_state["bk_current_shop"] = d.get("shop", "")
                            st.session_state["bk_bill_draft"] = d.get("items", [])
                        st.session_state.pop("database", None)
                        st.rerun()

        with tab3:
            su_user = st.text_input("New Username", key="su_username")
            su_pass = st.text_input("Password", key="su_password", type="password")
            su_pass2 = st.text_input("Confirm Password", key="su_password2", type="password")
            if st.button("Create Account", key="btn_signup", use_container_width=True, type="primary"):
                users = load_users()
                uname_safe = sanitize_username(su_user.strip())
                if len(uname_safe) < 3: st.error("❌ Username too short")
                elif len(su_pass) < 4: st.error("❌ Password too short")
                elif su_pass != su_pass2: st.error("❌ Passwords do not match")
                elif uname_safe in users: st.error("❌ Username already exists")
                else:
                    add_or_update_user(username=uname_safe, display_name=su_user.strip(), password_hash=hash_password(su_pass))
                    save_cloud_db(uname_safe, default_blank_db())
                    st.success(f"✅ Account created — Username: **{uname_safe}**")

# ============================================================
# SESSION GUARD
# ============================================================
if not st.session_state.get("logged_in_user"):
    render_auth_page()
    st.stop()

CURRENT_USER = st.session_state["logged_in_user"]
ROLE = st.session_state.get("role", "admin")

st.session_state.database = load_database_for(CURRENT_USER)
db = st.session_state.database

for _k, _v in [
    ("_prev_prod", None), ("download_file", None), ("_dl_counter", 0),
    ("page", "📊 Dashboard"), ("bk_page", "🧾 New Bill"),
    ("bk_bill_draft", []), ("bk_load_draft", []),
    ("bk_current_shop", ""), ("bk_bills_created", 0),
    ("bk_reset_item_form", False),
    ("prod_select_mode", False), ("prod_confirm_all", False),
]:
    if _k not in st.session_state: st.session_state[_k] = _v

def save_database(db_data):
    ok = save_cloud_db(CURRENT_USER, db_data)
    if ok: st.session_state.database = db_data
    return ok

# ============================================================
# HELPERS
# ============================================================
def get_all_products():
    merged = list(PRODUCTS)
    seen = {str(p["code"]) for p in merged}
    for cp in db.get("custom_products", []):
        code = str(cp.get("code", "")).strip()
        name = str(cp.get("name", "")).strip()
        if not code or not name or code in seen: continue
        try: price = float(cp.get("price", 0))
        except Exception: price = 0.0
        merged.append({"code": code, "name": name, "price": price})
        seen.add(code)
    return sorted(merged, key=lambda x: x["name"])

def get_price(code, base_price):
    custom = db.get("product_prices", {})
    if str(code) in custom:
        try: return float(custom[str(code)])
        except Exception: return float(base_price)
    return float(base_price)

_WS_PKG_RE = re.compile(r"(?:^|[\s\-_])(ws1|ws2|ws3)(?:$|[\s\-_])", re.IGNORECASE)
_WS_WORD_RE = re.compile(r"\b(wholesaler|wholeseller|wholseller|wholsaler)\b", re.IGNORECASE)

def is_wholesaler(shop_name):
    if not shop_name: return False
    s = str(shop_name).strip()
    return bool(_WS_PKG_RE.search(s) or _WS_WORD_RE.search(s))

def get_wholesaler_package_from_shop(shop_name):
    if not shop_name: return None
    m = _WS_PKG_RE.search(str(shop_name).strip())
    if not m: return None
    token = m.group(1).lower()
    for p in db.get("wholesaler_packages", []):
        if p["id"].lower() == token: return p
    return None

def get_all_active_packages(): return [p for p in db.get("discount_packages", []) if p.get("active")]

def get_package_discount_pct(bill_total):
    active = get_all_active_packages()
    if not active: return 0.0, None, None
    best_pct = 0.0; best_name = None; best_tier = None
    for pkg in active:
        pkg_name = pkg.get("name", "Package")
        tiers = [
            (float(pkg.get("tier3_amount", 0) or 0), float(pkg.get("tier3_pct", 0) or 0), "Tier3"),
            (float(pkg.get("tier2_amount", 0) or 0), float(pkg.get("tier2_pct", 0) or 0), "Tier2"),
            (float(pkg.get("tier1_amount", 0) or 0), float(pkg.get("tier1_pct", 0) or 0), "Tier1"),
        ]
        tiers.sort(key=lambda x: x[0], reverse=True)
        for amt, pct, label in tiers:
            if amt > 0 and bill_total >= amt and pct > best_pct:
                best_pct = pct; best_name = pkg_name; best_tier = f"{label} (≥ {amt:,.0f})"
    if best_pct > 0: return best_pct, best_name, best_tier
    return 0.0, (active[0].get("name") if active else None), None

def get_effective_discount_pct(shop_name, bill_total):
    ws = get_wholesaler_package_from_shop(shop_name)
    if ws:
        pct = float(ws.get("discount_pct", 0))
        return pct, f"{ws['name']} ({pct}%)", "WS Package"
    if is_wholesaler(shop_name): return 0.0, "Wholesaler Keyword", "No Package Matched"
    return get_package_discount_pct(bill_total)

def get_lf_return_items(lf_id):
    result = {}
    for r in db.get("booker_returns", []):
        if r.get("load_form_id") != lf_id: continue
        for it in r.get("items", []):
            c = str(it.get("Code", ""))
            result[c] = result.get(c, 0) + int(it.get("ReturnBoxes", 0))
        if not r.get("items") and r.get("boxes"): result["__misc__"] = result.get("__misc__", 0) + int(r.get("boxes", 0))
    return result

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

def get_dsr_credit_info(dsr):
    booker = dsr.get("booker", "").strip()
    dsr_date = dsr.get("date", "").strip()
    credit_total = 0.0; credit_shops = {}; details = []
    for sb in dsr.get("source_bills", []):
        shop = sb.get("shop", ""); bill_no = sb.get("bill_no", "")
        for cb in db.get("credit_bills", []):
            if (str(cb.get("booker", "")).strip() == booker and str(cb.get("bill_no", "")) == str(bill_no) and str(cb.get("shop", "")) == str(shop) and str(cb.get("date", "")).strip() == dsr_date):
                amt = float(cb.get("total_net", 0)); credit_total += amt; credit_shops[shop] = credit_shops.get(shop, 0) + amt
                details.append({"shop": shop, "bill_no": bill_no, "amount": amt, "status": cb.get("status", "pending"), "credit_id": cb.get("id")})
                break
    return credit_total, credit_shops, details

def build_load_form_excel(lf):
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Load Form")
    tf = wb.add_format({"bold": True, "font_size": 16, "align": "center", "border": 2, "bg_color": "#e2e8f0"})
    hf = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f1f5f9", "align": "center", "border": 1})
    lf_fmt = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f8fafc", "border": 1})
    vf = wb.add_format({"font_size": 11, "border": 1})
    cl = wb.add_format({"font_size": 11, "border": 1, "align": "left"})
    cc = wb.add_format({"font_size": 11, "border": 1, "align": "center"})
    tt = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#fef3c7", "align": "center", "border": 1})
    nf = wb.add_format({"bold": True, "font_size": 13, "bg_color": "#dcfce7", "align": "center", "border": 2})
    ws.merge_range("A1:C1", f"{COMPANY_NAME} — LOAD FORM", tf)
    ws.write("A3", "Booker:", lf_fmt); ws.write("B3", lf.get("booker", ""), vf)
    ws.write("A4", "Date:", lf_fmt); ws.write("B4", lf.get("date", ""), vf); ws.write("C4", lf.get("time", ""), vf)
    ws.write("A5", "Load Form #:", lf_fmt); ws.write("B5", lf.get("id", ""), vf)
    sm = lf.get("salesman", "-") or "-"
    if sm == "-" and lf.get("items"): sm = lf["items"][0].get("Salesman", "-")
    ws.write("A6", "Salesman:", lf_fmt); ws.write("B6", sm, vf)
    ws.write("A8", "Code", hf); ws.write("B8", "Product", hf); ws.write("C8", "Boxes", hf)
    row = 8; total = 0
    for it in lf.get("items", []):
        bx = int(it.get("Boxes", 0))
        ws.write(row, 0, str(it.get("Code", "")), cc)
        ws.write(row, 1, it.get("Product", ""), cl)
        ws.write(row, 2, bx, cc)
        total += bx
        row += 1
    ws.write(row, 0, "TOTAL", tt); ws.write(row, 1, f"{len(lf.get('items', []))} Products", tt)
    ws.write(row, 2, total, tt)
    row += 2
    ws.merge_range(row, 0, row, 1, "TOTAL BOXES", hf); ws.merge_range(row, 2, row, 2, total, nf)
    ws.set_column("A:A", 14); ws.set_column("B:B", 32); ws.set_column("C:C", 12)
    ws.set_paper(9); ws.set_portrait(); ws.fit_to_pages(1, 1)
    wb.close(); output.seek(0); return output.getvalue()

def _build_bill_excel(shop, bills, pkg_pct, pkg_name, tier_label, booker="", date_str=""):
    booker = booker or (bills[0].get("Order Booker", "") if bills else "")
    date_str = date_str or (bills[0].get("Date", "") if bills else "")
    salesman = bills[0].get("Salesman", "") if bills else ""
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Bill")
    ws.set_column(0, 0, 28); ws.set_column(1, 9, 12)
    title = wb.add_format({"bold": True, "font_size": 18, "align": "center", "border": 2, "bg_color": "#e2e8f0"})
    il = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f8fafc", "align": "right", "border": 1})
    iv = wb.add_format({"bold": True, "font_size": 11, "align": "left", "border": 1, "font_color": "#2563eb"})
    idate = wb.add_format({"bold": True, "font_size": 11, "align": "left", "border": 1, "font_color": "#16a34a"})
    hf = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f1f5f9", "align": "center", "border": 1, "text_wrap": True})
    cl = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "left", "text_wrap": True})
    cc = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "center"})
    cn = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "right", "num_format": "#,##0"})
    tt = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#fef3c7", "align": "center", "border": 2})
    ttn = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#fef3c7", "align": "right", "border": 2, "num_format": "#,##0"})
    dl = wb.add_format({"bold": True, "font_size": 11, "border": 1, "align": "center", "bg_color": "#dcfce7"})
    pi = wb.add_format({"bold": True, "font_size": 10, "italic": True, "align": "left", "font_color": "#16a34a", "border": 1})
    nl = wb.add_format({"bold": True, "font_size": 13, "bg_color": "#0f172a", "font_color": "#ffffff", "align": "center", "border": 2})
    nv = wb.add_format({"bold": True, "font_size": 16, "bg_color": "#dcfce7", "align": "center", "border": 2, "num_format": '"RS "#,##0'})
    ws.merge_range("A1:J1", COMPANY_NAME, title)
    ws.merge_range("A3:B3", "ORDER BOOKER:", il); ws.merge_range("C3:E3", (booker or "-").upper(), iv)
    ws.merge_range("F3:G3", "DATE:", il); ws.merge_range("H3:J3", (date_str or "-").upper(), idate)
    ws.merge_range("A4:B4", "SHOP:", il); ws.merge_range("C4:E4", (shop or "-").upper(), iv)
    ws.merge_range("F4:G4", "SALESMAN:", il); ws.merge_range("H4:J4", (salesman or "-").upper(), iv)
    ws.merge_range("A5:B5", "BILL NO:", il); ws.merge_range("C5:E5", f"#{bills[0].get('Bill No', '') if bills else ''}", iv)
    if pkg_pct > 0: ws.merge_range("A6:J6", f"DISCOUNT: {pkg_name} | {tier_label} | {pkg_pct}%".upper(), pi)
    else: ws.merge_range("A6:J6", "NO DISCOUNT APPLIED", pi)
    headers = ["PRODUCT", "CODE", "BOXES", "TP/BOX", "GROSS", "DISC %", "NET", "PKG DISC %", "AFTER DISC NET", "SAVED"]
    for c, h in enumerate(headers): ws.write(7, c, h, hf)
    row = 8; gt = tb = nt = ad = sv = 0
    for b in bills:
        bn = float(b.get("Net", 0)); bg = float(b.get("Gross", 0)); bb = int(b.get("Boxes", 0))
        an = bn - (bn * pkg_pct / 100); s = bn - an
        ws.write(row, 0, str(b["Product"]).upper(), cl); ws.write(row, 1, str(b["Code"]).upper(), cc); ws.write(row, 2, bb, cc)
        ws.write(row, 3, float(b.get("TP/Box", 0)), cn); ws.write(row, 4, bg, cn); ws.write(row, 5, float(b.get("Discount %", 0)), cc); ws.write(row, 6, bn, cn)
        if pkg_pct > 0:
            ws.write(row, 7, pkg_pct, dl); ws.write(row, 8, an, dl); ws.write(row, 9, s, dl)
        else:
            ws.write(row, 7, 0, cc); ws.write(row, 8, bn, cn); ws.write(row, 9, 0, cn)
        gt += bg; tb += bb; nt += bn; ad += an; sv += s; row += 1
    ws.merge_range(row, 0, row, 1, "TOTAL", tt)
    ws.write(row, 2, tb, tt); ws.write(row, 3, "", tt); ws.write(row, 4, gt, ttn); ws.write(row, 5, "", tt); ws.write(row, 6, nt, ttn); ws.write(row, 7, "", tt); ws.write(row, 8, ad, ttn); ws.write(row, 9, sv, ttn)
    row += 2
    ws.merge_range(row, 0, row, 6, "NET AMOUNT", nl); ws.merge_range(row, 7, row, 9, ad, nv)
    ws.set_paper(9); ws.set_portrait(); ws.fit_to_pages(1, 1)
    wb.close(); output.seek(0); return output.getvalue()

def export_dsr_excel(dsr):
    ct, _, _ = get_dsr_credit_info(dsr)
    fin = float(dsr.get("amount_to_collect", 0)) - ct
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("DSR")
    title = wb.add_format({"bold": True, "font_size": 16, "align": "center", "border": 2, "bg_color": "#e2e8f0"})
    hf = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#f1f5f9", "align": "center", "border": 1})
    cl = wb.add_format({"font_size": 11, "border": 1, "align": "left"})
    cc = wb.add_format({"font_size": 11, "border": 1, "align": "center"})
    cn = wb.add_format({"font_size": 11, "border": 1, "align": "right", "num_format": "#,##0"})
    tot = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#fef3c7", "align": "center", "border": 1})
    hi = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#dcfce7", "align": "center", "border": 2})
    for col, w in [("A", 10), ("B", 40), ("C", 10), ("D", 12), ("E", 14), ("F", 10), ("G", 14)]: ws.set_column(f"{col}:{col}", w)
    ws.merge_range("A1:G1", f"{COMPANY_NAME} — DSR #{dsr['id']}", title)
    ws.write("A3", "Booker", hf); ws.write("B3", dsr.get("booker", ""), cl); ws.write("C3", "Salesman", hf); ws.write("D3", dsr.get("salesman", "") or "-", cl); ws.write("E3", "Date", hf); ws.write("F3", dsr.get("date", ""), cc)
    row = 5
    for c, h in enumerate(["Code", "Product", "Boxes", "TP/Box", "Total", "Ret Bx", "Ret Amt"]): ws.write(row, c, h, hf)
    row += 1
    for it in dsr.get("items", []):
        ws.write(row, 0, str(it.get("Code", "")), cc); ws.write(row, 1, it.get("Product", ""), cl); ws.write(row, 2, int(it.get("Boxes", 0)), cc); ws.write(row, 3, float(it.get("TP/Box", 0)), cn); ws.write(row, 4, float(it.get("Total", 0)), cn); ws.write(row, 5, int(it.get("ReturnBoxes", 0)), cc); ws.write(row, 6, float(it.get("ReturnAmount", 0)), cn); row += 1
    ws.write(row, 2, "TOTAL", tot); ws.write(row, 4, float(dsr.get("total_amount", 0)), tot); ws.write(row, 5, int(dsr.get("total_return_boxes", 0)), tot); ws.write(row, 6, float(dsr.get("total_return_amount", 0)), tot); row += 2
    for lbl, val in [("Stock Value", dsr.get("total_amount", 0)), ("(−) Returns", dsr.get("total_return_amount", 0)), ("(=) Net Stock", dsr.get("net_amount", 0)), ("(−) Discount", dsr.get("total_discount", 0)), ("(=) Ye Lena Hai", dsr.get("amount_to_collect", 0)), ("(−) Credit Bills", ct)]:
        ws.merge_range(row, 0, row, 4, lbl, hf); ws.merge_range(row, 5, row, 6, f"Rs {float(val):,.0f}", cn); row += 1
    ws.merge_range(row, 0, row, 4, "💰 FINAL YE LENA HAI", hi); ws.merge_range(row, 5, row, 6, f"Rs {fin:,.0f}", hi)
    wb.close(); output.seek(0); st.session_state["download_file"] = (f"DSR_{dsr['id']}_{dsr.get('booker', '')}.xlsx".replace("/", "-").replace(" ", "_"), output.getvalue()); st.session_state["success_msg"] = f"✅ DSR #{dsr['id']} ready"

def show_auto_download():
    if st.session_state.get("download_file"):
        fname, fdata = st.session_state["download_file"]
        st.session_state["download_file"] = None; st.session_state["_dl_counter"] += 1
        key = f"auto_dl_{st.session_state['_dl_counter']}"
        st.download_button(f"⬇️ Download {fname}", data=fdata, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=key)
        components.html(f"""<script>(function(){{ var t=0; function a(){{ t++; var b=window.parent.document.querySelectorAll('[data-testid="stDownloadButton"] button'); if(b.length>0){{ b[b.length-1].click(); return; }} if(t<25) setTimeout(a, 200); }} setTimeout(a, 400); }})();</script>""", height=0)

# ============================================================
# BOOKER UI
# ============================================================
def render_booker_sidebar():
    with st.sidebar:
        bk = st.session_state.get("booker_name", "Booker")
        st.markdown(f"""
        <div style="text-align:center;padding:20px;">
            <div style="font-size:24px;">👤</div>
            <div style="font-weight:800;color:#1b5e20;letter-spacing:2px;">BOOKER PANEL</div>
            <div style="font-size:10px;color:#558b2f;letter-spacing:4px;">AL-BARAKAH</div>
        </div>
        <div style="background:#fff;border-radius:12px;padding:10px 14px;margin:0 14px 16px;border-left:4px solid #66bb6a;">
            <div style="font-size:9px;color:#689f38;text-transform:uppercase;font-weight:800;">Booker</div>
            <div style="font-size:14px;font-weight:800;color:#1b5e20;">{bk}</div>
        </div>
        """, unsafe_allow_html=True)
        bk_page = st.radio("BK", ["🧾 New Bill", "📋 My Bills", "🛒 All Products", "📦 Load Form Status"], key="bk_page_selector", label_visibility="collapsed")
        st.session_state["bk_page"] = bk_page
        if st.button("🚪 Logout", key="btn_bk_logout", use_container_width=True):
            for k in ["logged_in_user", "display_name", "role", "booker_name", "booker_admin", "database", "bk_bill_draft", "bk_load_draft", "bk_current_shop", "bk_page", "bk_reset_item_form"]: st.session_state.pop(k, None)
            st.rerun()

def render_booker_new_bill():
    bk_name = st.session_state.get("booker_name", "")
    admin_uname = st.session_state.get("booker_admin", CURRENT_USER)
    st.markdown(f"<h1 style='color:#0f172a;'>🧾 New Bill — <span style='color:#2563eb;'>{bk_name}</span></h1>", unsafe_allow_html=True)
    st.markdown("---")
    if "bk_selected_salesman" not in st.session_state: st.session_state["bk_selected_salesman"] = "-- Select --"
    if st.session_state.get("bk_reset_item_form"):
        st.session_state["bk_product_sel"] = "-- Select --"; st.session_state["bk_qty"] = 0; st.session_state["bk_reset_item_form"] = False
    if is_wholesaler(st.session_state.get("bk_current_shop", "")): st.markdown("<div style='background:#e3f2fd;border-left:4px solid #2196f3;padding:8px;'>🏢 Wholesaler detected!</div>", unsafe_allow_html=True)
    
    ca, cb = st.columns(2)
    with ca: shop = st.text_input("🏪 Shop Name", value=st.session_state.get("bk_current_shop", ""), key="bk_shop_input")
    with cb:
        sms = db.get("salesmen", [])
        if sms:
            sel_sm = st.selectbox("🧑‍💼 Salesman", ["-- Select --"] + sms, key="bk_salesman_select")
            st.session_state["bk_selected_salesman"] = sel_sm
            current_sm = "" if sel_sm == "-- Select --" else sel_sm
        else:
            current_sm = st.text_input("🧑‍💼 Salesman", key="bk_salesman_manual")
            st.session_state["bk_selected_salesman"] = current_sm
    if shop:
        new_shop = shop.strip()
        if new_shop != st.session_state.get("bk_current_shop", ""):
            st.session_state["bk_current_shop"] = new_shop; st.session_state["bk_bill_draft"] = []; save_booker_draft(admin_uname, bk_name, new_shop, [], current_sm); st.rerun()
    if not shop.strip(): st.warning("⚠️ Enter shop name first"); return
    if is_wholesaler(shop):
        ws = get_wholesaler_package_from_shop(shop)
        if ws: st.markdown(f"<div style='background:#faf5ff;border-left:4px solid #9333ea;padding:8px;'>🏢 {ws['name']} — {ws['discount_pct']}%</div>", unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### ➕ Add Item")
    all_p = get_all_products()
    search = st.text_input("🔍 Search", key="bk_search")
    su = search.strip().upper()
    filtered = [p for p in all_p if su in p["name"].upper() or su in str(p["code"]).upper()] if su else all_p
    labels = [f"{p['name']} (Code: {p['code']})" for p in filtered]
    c1, c2, c3 = st.columns([3, 1, 1])
    with c1: selected = st.selectbox("Product", ["-- Select --"] + labels, key="bk_product_sel")
    with c2: boxes = st.number_input("Boxes", min_value=0, step=1, value=0, key="bk_qty")
    with c3: st.markdown("<br>", unsafe_allow_html=True); add_clicked = st.button("➕ Add", key="bk_add_item_btn", use_container_width=True, type="primary")
    sel_product = None
    if selected != "-- Select --":
        for p in all_p:
            if f"{p['name']} (Code: {p['code']})" == selected: sel_product = p; break
    if sel_product:
        tp = get_price(sel_product["code"], sel_product["price"]); st.success(f"💰 {sel_product['name']} — TP: Rs {tp:,.0f} | Total: Rs {boxes * tp:,.0f}")
    if add_clicked:
        if not sel_product: st.error("❌ Select Product")
        elif boxes <= 0: st.error("❌ Enter Boxes")
        else:
            tp = get_price(sel_product["code"], sel_product["price"])
            st.session_state["bk_bill_draft"].append({"Code": sel_product["code"], "Product": sel_product["name"], "Boxes": int(boxes), "TP/Box": float(tp), "Gross": float(boxes * tp), "Discount %": 0.0, "Net": float(boxes * tp), "Salesman": current_sm})
            save_booker_draft(admin_uname, bk_name, shop, st.session_state["bk_bill_draft"], current_sm); st.session_state["bk_reset_item_form"] = True; st.rerun()
    
    st.markdown("---")
    draft = st.session_state.get("bk_bill_draft", [])
    if not draft: st.info("No items added yet.")
    else:
        st.markdown(f"### 📋 Current Bill — {shop}")
        st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"], "TP/Box": it["TP/Box"], "Total": it["Gross"], "Salesman": it.get("Salesman", "")} for it in draft]), use_container_width=True, hide_index=True)
        c1, c2 = st.columns([3, 1])
        with c1: rem = st.selectbox("Delete:", ["--"] + [f"{i+1}. {it['Product']} ({it['Boxes']} boxes)" for i, it in enumerate(draft)], key="bk_remove_sel")
        with c2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑 Remove", key="bk_remove_btn", use_container_width=True):
                if rem != "--": st.session_state["bk_bill_draft"].pop(int(rem.split(".")[0]) - 1); save_booker_draft(admin_uname, bk_name, shop, st.session_state["bk_bill_draft"], current_sm); st.rerun()
        gt = sum(float(it["Gross"]) for it in draft)
        st.markdown(f"<div style='background:#fff;border:2px solid #2196f3;border-radius:12px;padding:12px;'><b>Gross:</b> <span style='color:#2563eb;font-weight:700;font-size:20px;'>Rs {gt:,.0f}</span></div>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Done Bill", key="bk_done_bill", use_container_width=True, type="primary"): finalize_booker_bill(shop, current_sm, draft, admin_uname, bk_name)
        with c2:
            if st.button("🗑 Clear Draft", key="bk_clear_draft", use_container_width=True): st.session_state["bk_bill_draft"] = []; clear_booker_draft(admin_uname, bk_name); st.rerun()
    
    st.markdown("---")
    ld = st.session_state.get("bk_load_draft", [])
    st.markdown("### 📦 Load Form — Session")
    if not ld: st.info("No bills completed yet.")
    else:
        tb = sum(int(it["Boxes"]) for it in ld)
        st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"], "Salesman": it.get("Salesman", "")} for it in ld]), use_container_width=True, hide_index=True)
        st.markdown(f"<div style='background:#fff;border:2px solid #2196f3;border-radius:12px;padding:12px;'><b>Boxes:</b> {tb}</div>", unsafe_allow_html=True)
        if st.button("📦 Done Load Form", key="bk_done_load", use_container_width=True, type="primary"): finalize_booker_load_form(admin_uname, bk_name)

def finalize_booker_bill(shop, salesman, draft, admin_uname, bk_name):
    global db
    bill_no = db["next_bill_no"]; date_str = datetime.now().strftime("%d-%m-%Y")
    for it in draft: db["bills"].append({"Bill No": bill_no, "Date": date_str, "Shop": shop, "Order Booker": bk_name, "Salesman": salesman, "Delivery Man": "", "Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"], "TP/Box": it["TP/Box"], "Discount %": it["Discount %"], "Gross": it["Gross"], "Net": it["Net"]})
    db["next_bill_no"] += 1
    if not save_database(db): st.session_state["error_msg"] = "❌ Cloud save failed"; return
    for it in draft:
        found = False
        for l in st.session_state["bk_load_draft"]:
            if str(l["Code"]) == str(it["Code"]): l["Boxes"] += int(it["Boxes"]); l["Salesman"] = salesman; found = True; break
        if not found: st.session_state["bk_load_draft"].append({"Code": it["Code"], "Product": it["Product"], "Boxes": int(it["Boxes"]), "Salesman": salesman})
    st.session_state["bk_bill_draft"] = []; clear_booker_draft(admin_uname, bk_name); st.session_state["bk_bills_created"] = st.session_state.get("bk_bills_created", 0) + 1; st.session_state["success_msg"] = f"✅ Bill #{bill_no} saved"; st.rerun()

def finalize_booker_load_form(admin_uname, bk_name):
    global db
    ld = st.session_state.get("bk_load_draft", [])
    if not ld: st.session_state["error_msg"] = "❌ Empty"; return
    today = datetime.now().strftime("%d-%m-%Y"); tb = sum(int(it["Boxes"]) for it in ld); sm = ld[0].get("Salesman", "Unknown"); existing = None
    for lf in db.get("load_forms", []):
        lf_sm = lf.get("salesman", "")
        if not lf_sm and lf.get("items"): lf_sm = lf["items"][0].get("Salesman", "Unknown")
        if (lf.get("booker") == bk_name and lf.get("date") == today and not lf.get("transferred_to_dsr") and lf_sm == sm): existing = lf; break
    if existing:
        existing["items"] = ld; existing["total_boxes"] = tb; existing["total_products"] = len(ld); existing["time"] = datetime.now().strftime("%H:%M"); existing["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); existing["salesman"] = sm; msg = f"🔄 LF #{existing['id']} updated"
    else:
        nid = max([lf.get("id", 0) for lf in db["load_forms"]] + [0]) + 1
        db["load_forms"].append({"id": nid, "date": today, "time": datetime.now().strftime("%H:%M"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "booker": bk_name, "salesman": sm, "items": ld, "total_boxes": tb, "total_products": len(ld), "transferred_to_dsr": False, "refreshed": False, "source": "booker"}); msg = f"✅ LF #{nid} saved"
    if not save_database(db): st.session_state["error_msg"] = "❌ Cloud save failed"; return
    st.session_state["bk_load_draft"] = []; st.session_state["success_msg"] = msg; st.rerun()

def render_booker_all_products():
    st.markdown("<h1 style='color:#0f172a;'>🛒 Products</h1>", unsafe_allow_html=True); st.markdown("---")
    all_p = get_all_products(); search = st.text_input("🔍 Search", key="bk_prod_search"); su = search.strip().upper()
    shown = [p for p in all_p if su in p["name"].upper() or su in str(p["code"]).upper()] if su else all_p
    st.markdown(f"<div style='background:#fff;border:2px solid #2196f3;border-radius:12px;padding:12px;'><b>Total:</b> {len(all_p)} | <b>Showing:</b> {len(shown)}</div>", unsafe_allow_html=True)
    if not shown: st.info("No products."); return
    for p in shown:
        tp = get_price(p["code"], p["price"])
        st.markdown(f"""<div style="background:#fff;border-left:6px solid #2196f3;border-radius:12px;padding:12px 18px;margin-bottom:8px;display:flex;justify-content:space-between;"><div><div style="font-weight:700;color:#1976d2;">{p['name']}</div><div style="font-size:12px;color:#666;">Code: {p['code']}</div></div><div style="background:#dcfce7;color:#16a34a;font-weight:800;padding:6px 14px;border-radius:8px;">Rs {tp:,.0f}</div></div>""", unsafe_allow_html=True)

def render_booker_load_status():
    admin_uname = st.session_state.get("booker_admin", CURRENT_USER); bk_name = st.session_state.get("booker_name", "")
    st.markdown("<h1 style='color:#0f172a;'>📦 Load Form Status</h1>", unsafe_allow_html=True); st.markdown("---")
    ld = st.session_state.get("bk_load_draft", [])
    if not ld: st.info("No bills yet."); return
    tb = sum(int(it["Boxes"]) for it in ld)
    st.dataframe(pd.DataFrame([{"Code": it["Code"], "Product": it["Product"], "Boxes": it["Boxes"], "Salesman": it.get("Salesman", "")} for it in ld]), use_container_width=True, hide_index=True)
    st.markdown(f"<div style='background:#fff;border:2px solid #2196f3;border-radius:12px;padding:12px;'><b>Boxes:</b> {tb}</div>", unsafe_allow_html=True)
    if st.button("📦 Done Load Form", key="bk_done_load_2", use_container_width=True, type="primary"): finalize_booker_load_form(admin_uname, bk_name)

def render_booker_bills_history():
    bk_name = st.session_state.get("booker_name", "")
    st.markdown("<h1 style='color:#0f172a;'>📋 My Bills</h1>", unsafe_allow_html=True); st.markdown("---")
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: fm = st.selectbox("Filter", ["📅 Aaj", "📆 Custom", "🗓️ Specific", "📋 All"], key="bk_history_filter")
    with c2: fd = st.date_input("From", value=today - timedelta(days=7), key="bk_history_from")
    with c3: td = st.date_input("To", value=today, key="bk_history_to")
    groups = {}
    for b in db["bills"]:
        if str(b.get("Order Booker", "")).strip() != bk_name: continue
        bd = parse_date(b.get("Date", ""))
        if fm == "📅 Aaj" and bd != today: continue
        elif fm == "🗓️ Specific" and bd != fd: continue
        elif fm == "📆 Custom" and (bd is None or not (fd <= bd <= td)): continue
        key = (b.get("Shop", ""), b.get("Date", ""), b.get("Bill No", ""))
        if key not in groups: groups[key] = {"shop": b.get("Shop", ""), "date": b.get("Date", ""), "bill_no": b.get("Bill No", ""), "items": [], "idx": []}
        groups[key]["items"].append(b); groups[key]["idx"].append(db["bills"].index(b))
    if not groups: st.info("No bills."); return
    for key, g in sorted(groups.items(), key=lambda kv: kv[0][1], reverse=True):
        tb = sum(int(it.get("Boxes", 0)) for it in g["items"]); tn = sum(float(it.get("Net", 0)) for it in g["items"]); sm = g["items"][0].get("Salesman", "-")
        st.markdown(f"""<div style="background:#fff;border-left:6px solid #2196f3;border-radius:12px;padding:12px 18px;margin-bottom:8px;display:flex;justify-content:space-between;"><div><div style="font-weight:700;color:#1976d2;">🏪 {g['shop']}</div><div style="font-size:12px;color:#666;">📅 {g['date']} · 🧑‍💼 {sm} · 🧾 #{g['bill_no']} · Rs {tn:,.0f}</div></div><div style="background:#e3f2fd;color:#1976d2;font-weight:800;padding:6px 14px;border-radius:8px;">{tb} BOXES</div></div>""", unsafe_allow_html=True)
        with st.expander("View Items"):
            st.dataframe(pd.DataFrame([{"Code": i["Code"], "Product": i["Product"], "Boxes": i["Boxes"], "Net": i["Net"], "Salesman": i.get("Salesman", "")} for i in g["items"]]), use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🗑 Delete", key=f"bk_del_{key}_{g['date']}", type="primary", use_container_width=True): idxs = set(g["idx"]); db["bills"] = [b for i, b in enumerate(db["bills"]) if i not in idxs]; save_database(db); st.session_state["success_msg"] = f"🗑 Deleted"; st.rerun()
        with c2:
            if st.button("📤 Resend", key=f"bk_resend_{key}_{g['date']}", use_container_width=True): save_database(db); st.session_state["success_msg"] = "✅ Resent"; st.rerun()
        st.markdown("---")

# ============================================================
# ROUTER — BOOKER
# ============================================================
if ROLE == "booker":
    render_booker_sidebar()
    p = st.session_state.get("bk_page", "🧾 New Bill")
    if p == "🧾 New Bill": render_booker_new_bill()
    elif p == "📋 My Bills": render_booker_bills_history()
    elif p == "🛒 All Products": render_booker_all_products()
    elif p == "📦 Load Form Status": render_booker_load_status()
    for k in ["success_msg", "error_msg", "cloud_error"]:
        if st.session_state.get(k): {"success_msg": st.success, "error_msg": st.error, "cloud_error": st.error}[k](st.session_state[k]); st.session_state[k] = None
    st.stop()

# ============================================================
# ADMIN SIDEBAR
# ============================================================
with st.sidebar:
    dn = st.session_state.get("display_name", CURRENT_USER)
    st.markdown(f"""<div style="text-align:center;padding:20px;"><div style="font-size:24px;">🧾</div><div style="font-weight:800;color:#1b5e20;letter-spacing:2px;">AL-BARAKAH</div><div style="font-size:10px;color:#558b2f;letter-spacing:4px;">ENTERPRISES</div></div><div style="background:#fff;border-radius:12px;padding:10px 14px;margin:0 14px 16px;border-left:4px solid #66bb6a;"><div style="font-size:9px;color:#689f38;text-transform:uppercase;font-weight:800;">Admin</div><div style="font-size:14px;font-weight:800;color:#1b5e20;">{dn}</div></div>""", unsafe_allow_html=True)
    page = st.radio("MENU", ["📊 Dashboard", "🧾 Billing", "🛒 All Products", "🎁 Discount", "👤 Bookers", "📈 Booker Bills", "💰 Bookers Salary", "🧑‍💼 Salesmen", "💰 Salesmen Salary", "💵 Daily Expense", "📋 Bills List", "💳 Credit Bills", "📦 Load Form", "📋 DSR", "🧮 Calculation"], key="page_selector", label_visibility="collapsed")
    st.session_state["page"] = page
    if st.button("🚪 Logout", key="btn_logout", use_container_width=True):
        for k in ["logged_in_user", "display_name", "role", "database", "page"]: st.session_state.pop(k, None)
        st.rerun()

# ============================================================
# ADMIN PAGES
# ============================================================
def render_admin_dashboard():
    dn = st.session_state.get("display_name", CURRENT_USER); today_str = date.today().strftime("%A, %d %B %Y")
    bk = db.get("bookers", []); sm = db.get("salesmen", []); tp = len(get_all_products()); tb = len(db["bills"])
    pc = [c for c in db.get("credit_bills", []) if c.get("status") == "pending"]; yc = [c for c in db.get("credit_bills", []) if c.get("status") == "paid"]
    pa = sum(float(c.get("total_net", 0)) for c in pc); ya = sum(float(c.get("total_net", 0)) for c in yc)
    st.markdown(f"""<div class="dash-hero"><div style="font-size:11px;font-weight:800;letter-spacing:2px;color:rgba(255,255,255,0.75);">WELCOME</div><h1 style="font-size:28px;margin:4px 0;color:#fff;">{dn}</h1><div style="font-size:13px;color:rgba(255,255,255,0.9);">🏢 {COMPANY_NAME} · 📅 {today_str}</div></div>""", unsafe_allow_html=True)
    st.markdown("### Business Overview")
    c1, c2, c3, c4 = st.columns(4)
    for col, (icon, lbl, val, sub) in zip([c1, c2, c3, c4], [("👤", "Bookers", len(bk), "Team"), ("🧑‍💼", "Salesmen", len(sm), "Field"), ("📦", "Products", tp, "Active"), ("🧾", "Bills", tb, "Total")]):
        with col:
            st.markdown(f"""<div class="kpi-card"><div style="font-size:24px;">{icon}</div><div style="font-size:11px;color:#64748b;font-weight:800;margin-top:8px;">{lbl.upper()}</div><div style="font-size:28px;font-weight:800;color:#0f172a;">{val}</div><div style="font-size:11px;color:#2563eb;">{sub}</div></div>""", unsafe_allow_html=True)
    st.markdown("### Credit & Ops")
    c1, c2, c3, c4 = st.columns(4)
    for col, (icon, lbl, val, sub) in zip([c1, c2, c3, c4], [("⏳", "Pending Credit", len(pc), f"Rs {pa:,.0f}"), ("✅", "Paid Credit", len(yc), f"Rs {ya:,.0f}"), ("📦", "Load Forms", len(db.get("load_forms", [])), "Saved"), ("📋", "DSR", len(db.get("dsr_forms", [])), "Active")]):
        with col:
            st.markdown(f"""<div class="kpi-card"><div style="font-size:24px;">{icon}</div><div style="font-size:11px;color:#64748b;font-weight:800;margin-top:8px;">{lbl.upper()}</div><div style="font-size:28px;font-weight:800;color:#0f172a;">{val}</div><div style="font-size:11px;color:#ea580c;">{sub}</div></div>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### 💾 Backup")
    bc1, bc2 = st.columns(2)
    with bc1:
        bdata = json.dumps(db, indent=4, ensure_ascii=False)
        st.download_button("📥 Download Backup", data=bdata, file_name=f"Backup_{CURRENT_USER}_{datetime.now().strftime('%Y%m%d')}.json", mime="application/json", use_container_width=True, type="primary")
    with bc2:
        up = st.file_uploader("📂 Restore", type=["json"], key="restore_uploader")
        if up is not None:
            try:
                nd = _ensure_db_shape(json.load(up))
                if save_database(nd): st.success("✅ Restored!")
            except Exception as e: st.error(f"❌ {e}")

def render_admin_booker_bills():
    st.markdown("<h1 style='color:#0f172a;'>📈 Booker Bills & Returns</h1>", unsafe_allow_html=True); st.markdown("---")
    bks = db.get("bookers", [])
    if not bks: st.info("Add bookers first."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: fm = st.selectbox("Filter", ["📅 Aaj", "📆 Custom", "🗓️ Specific", "📋 All"], key="bb_filter")
    with c2: fd = st.date_input("From", value=today, key="bb_from")
    with c3: td = st.date_input("To", value=today, key="bb_to")
    def in_range(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if fm == "📅 Aaj": return d == today
        elif fm == "🗓️ Specific": return d == fd
        elif fm == "📆 Custom": return fd <= d <= td
        return True
    bills_by_bk = {}
    for b in db.get("bills", []): bills_by_bk.setdefault(str(b.get("Order Booker", "")).strip(), []).append(b)
    returns_by_bk = {}
    for r in db.get("booker_returns", []): returns_by_bk.setdefault(str(r.get("booker", "")).strip(), []).append(r)
    gs = gr = gn = 0
    for bkn in bks:
        bb = [b for b in bills_by_bk.get(bkn, []) if in_range(b.get("Date", ""))]
        br = [r for r in returns_by_bk.get(bkn, []) if in_range(r.get("date", ""))]
        ts = sum(int(b.get("Boxes", 0)) for b in bb); tr = 0
        for r in br:
            items = r.get("items", []); tr += sum(int(it.get("ReturnBoxes", 0)) for it in items) if items else int(r.get("boxes", 0))
        nb = ts - tr; gs += ts; gr += tr; gn += nb
        with st.expander(f"👤 {bkn} — Sold: {ts} | Ret: {tr} | Net: {nb}"):
            if bb: st.dataframe(pd.DataFrame([{"Bill No": b.get("Bill No"), "Date": b.get("Date"), "Shop": b.get("Shop"), "Product": b.get("Product"), "Boxes": b.get("Boxes"), "Salesman": b.get("Salesman", ""), "Net": b.get("Net")} for b in bb]), use_container_width=True, hide_index=True)
            else: st.info("No bills.")
    st.markdown(f"<div style='background:#fff;border:2px solid #2563eb;border-radius:12px;padding:14px;display:flex;justify-content:space-around;text-align:center;'><div><div style='font-size:10px;color:#64748b;font-weight:800;'>SOLD</div><div style='font-size:24px;font-weight:800;color:#2563eb;'>{gs}</div></div><div><div style='font-size:10px;color:#64748b;font-weight:800;'>RETURNS</div><div style='font-size:24px;font-weight:800;color:#dc2626;'>-{gr}</div></div><div><div style='font-size:10px;color:#64748b;font-weight:800;'>NET</div><div style='font-size:24px;font-weight:800;color:#16a34a;'>{gn}</div></div></div>", unsafe_allow_html=True)

def render_admin_discount():
    st.markdown("<h1 style='color:#0f172a;'>🎁 Discount</h1>", unsafe_allow_html=True); st.markdown("---")
    st.markdown("### 🏢 Wholesaler Packages")
    for pkg in db.get("wholesaler_packages", []):
        with st.expander(f"🏢 {pkg['name']} — {pkg['discount_pct']}%"):
            c1, c2 = st.columns([3, 1])
            with c1: np = st.number_input(f"{pkg['name']} %", value=float(pkg['discount_pct']), min_value=0.0, max_value=100.0, step=0.5, key=f"ws_pct_{pkg['id']}")
            with c2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾", key=f"ws_save_{pkg['id']}", type="primary"):
                    for pp in db["wholesaler_packages"]:
                        if pp["id"] == pkg["id"]: pp["discount_pct"] = float(np); break
                    save_database(db); st.success("✅"); st.rerun()
    st.markdown("---")
    st.markdown("### 🎁 Regular Packages")
    for i, pkg in enumerate(db.get("discount_packages", [])):
        pid = pkg.get("id", i + 1)
        with st.expander(f"🎁 {pkg.get('name', 'Package')} {'✅' if pkg.get('active') else '⭕'}"):
            c1, c2 = st.columns([1, 5])
            with c1: na = st.checkbox("Active", value=pkg.get("active", False), key=f"active_{pid}")
            with c2: nn = st.text_input("Name", value=pkg.get("name", ""), key=f"pkgname_{pid}")
            t1, t2, t3 = st.columns(3)
            with t1: a1 = st.number_input("Min T1", value=float(pkg.get("tier1_amount", 0)), key=f"t1a_{pid}"); p1 = st.number_input("% T1", value=float(pkg.get("tier1_pct", 0)), key=f"t1p_{pid}")
            with t2: a2 = st.number_input("Min T2", value=float(pkg.get("tier2_amount", 0)), key=f"t2a_{pid}"); p2 = st.number_input("% T2", value=float(pkg.get("tier2_pct", 0)), key=f"t2p_{pid}")
            with t3: a3 = st.number_input("Min T3", value=float(pkg.get("tier3_amount", 0)), key=f"t3a_{pid}"); p3 = st.number_input("% T3", value=float(pkg.get("tier3_pct", 0)), key=f"t3p_{pid}")
            if st.button("💾 Save", key=f"savepkg_{pid}", type="primary"):
                for pp in db["discount_packages"]:
                    if pp.get("id") == pid: pp.update({"name": nn, "active": na, "tier1_amount": float(a1), "tier1_pct": float(p1), "tier2_amount": float(a2), "tier2_pct": float(p2), "tier3_amount": float(a3), "tier3_pct": float(p3)}); break
                save_database(db); st.success("✅"); st.rerun()

def render_admin_products():
    st.markdown("<h1 style='color:#0f172a;'>🛒 Products</h1>", unsafe_allow_html=True); st.markdown("---")
    all_p = get_all_products()
    with st.expander("➕ New Product", expanded=False):
        c1, c2, c3, c4 = st.columns([1.5, 3.5, 1.5, 1])
        with c1: nc = st.text_input("Code", key="np_code")
        with c2: nn = st.text_input("Name", key="np_name")
        with c3: npr = st.number_input("Price", min_value=0.0, step=1.0, key="np_price", value=0.0)
        with c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕", key="btn_add_prod", type="primary"):
                cs, ns = str(nc).strip(), str(nn).strip()
                if not cs or not ns or npr <= 0: st.error("❌ Fill all")
                elif cs in {str(p["code"]) for p in all_p}: st.error("❌ Code exists")
                else:
                    db["custom_products"].append({"code": cs, "name": ns, "price": float(npr), "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                    save_database(db); st.success("✅"); st.rerun()
    with st.expander("📤 Bulk Upload", expanded=False):
        up = st.file_uploader("Excel/CSV", type=["xlsx", "xls", "csv"], key="bulk_up")
        if up:
            try:
                dfu = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                dfu.columns = [str(c).strip().lower() for c in dfu.columns]
                if {"code", "name", "price"}.issubset(dfu.columns):
                    st.dataframe(dfu.head(10), use_container_width=True)
                    if st.button("📥 Import", key="btn_bulk_imp", type="primary"):
                        ex = {str(p["code"]) for p in all_p}; ad = upd = 0
                        for _, r in dfu.iterrows():
                            try:
                                cr = r["code"]
                                if pd.isna(cr): continue
                                cs = str(int(float(cr))) if str(cr).replace('.', '', 1).isdigit() else str(cr).strip()
                                ns = str(r["name"]).strip(); pf = float(r["price"])
                            except Exception: continue
                            if not cs or not ns or pf < 0: continue
                            if cs in ex:
                                for cp in db.get("custom_products", []):
                                    if str(cp.get("code")) == cs: cp["price"] = pf; cp["name"] = ns; upd += 1; break
                            else:
                                db["custom_products"].append({"code": cs, "name": ns, "price": pf, "added_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}); ex.add(cs); ad += 1
                        save_database(db); st.success(f"✅ {ad} added, {upd} updated"); st.rerun()
                else: st.error("❌ Need columns: code, name, price")
            except Exception as e: st.error(f"❌ {e}")
    st.markdown(f"<div style='background:#fef2f2;border:2px solid #dc2626;border-radius:12px;padding:14px;margin:12px 0;'><b style='color:#b91c1c;'>⚠️ DANGER ZONE</b><br><span style='font-size:12px;color:#64748b;'>Active: {len(all_p)}</span></div>", unsafe_allow_html=True)
    dz1, dz2 = st.columns(2)
    with dz1: sm_mode = st.toggle("☑️ Select Mode", value=st.session_state.get("prod_select_mode", False), key="prod_sel_tog"); st.session_state["prod_select_mode"] = sm_mode
    with dz2:
        if st.button("🗑 DELETE ALL", type="primary", use_container_width=True): st.session_state["prod_confirm_all"] = True; st.rerun()
    if st.session_state.get("prod_confirm_all"):
        st.error("⚠️ Sure?")
        cc1, cc2 = st.columns(2)
        with cc1:
            if st.button("✅ YES", type="primary", use_container_width=True): db["custom_products"] = []; db["product_prices"] = {}; save_database(db); st.session_state["prod_confirm_all"] = False; st.rerun()
        with cc2:
            if st.button("❌ Cancel", use_container_width=True): st.session_state["prod_confirm_all"] = False; st.rerun()
    st.markdown("---")
    search = st.text_input("🔍 Search", key="prod_search"); su = search.strip().upper()
    shown = [p for p in all_p if su in p["name"].upper() or su in str(p["code"])] if su else all_p
    ep = db.get("product_prices", {}); cust = {str(cp.get("code", "")) for cp in db.get("custom_products", [])}
    if "prod_selected_codes" not in st.session_state: st.session_state["prod_selected_codes"] = []
    if st.session_state.get("prod_select_mode", False):
        sa1, sa2, sa3 = st.columns(3)
        with sa1:
            if st.button("☑️ All Visible", use_container_width=True): st.session_state["prod_selected_codes"] = sorted({str(p["code"]) for p in shown}); st.rerun()
        with sa2:
            if st.button("⬜ Clear", use_container_width=True): st.session_state["prod_selected_codes"] = []; st.rerun()
        with sa3:
            sc = len(st.session_state.get("prod_selected_codes", []))
            if st.button(f"🗑 Delete ({sc})", type="primary", use_container_width=True):
                if sc:
                    sel = set(st.session_state["prod_selected_codes"]); db["custom_products"] = [cp for cp in db.get("custom_products", []) if str(cp.get("code", "")) not in sel]
                    for c in sel: db.get("product_prices", {}).pop(c, None)
                    save_database(db); st.session_state["prod_selected_codes"] = []; st.rerun()
    for p in shown[:200]:
        code = str(p["code"]); base = float(p["price"]); cur = get_price(code, base); ic = code in cust; ie = code in ep; sn = code in st.session_state.get("prod_selected_codes", [])
        if st.session_state.get("prod_select_mode", False):
            c0, c1, c2, c3, c4 = st.columns([0.5, 1, 4, 2, 2])
            with c0:
                chk = st.checkbox("", value=sn, key=f"chk_{code}", label_visibility="collapsed")
                if chk and not sn: st.session_state["prod_selected_codes"].append(code); st.rerun()
                elif not chk and sn: st.session_state["prod_selected_codes"].remove(code); st.rerun()
        else: c1, c2, c3, c4 = st.columns([1, 4, 2, 2])
        with c1: st.markdown(f"<div style='padding-top:8px;color:#2563eb;font-weight:600;'>{code}</div>", unsafe_allow_html=True)
        with c2: marks = (" 🆕" if ic else "") + (" ✏️" if ie else ""); st.markdown(f"<div style='padding-top:8px;color:#0f172a;'>{p['name']}{marks}</div>", unsafe_allow_html=True)
        with c3: np = st.number_input("Price", value=float(cur), min_value=0.0, step=1.0, key=f"price_{code}", label_visibility="collapsed")
        with c4:
            s1, s2 = st.columns(2)
            with s1:
                if st.button("💾", key=f"sv_{code}", use_container_width=True):
                    if float(np) == base: db.get("product_prices", {}).pop(code, None)
                    else: db.setdefault("product_prices", {})[code] = float(np)
                    save_database(db); st.success("✅"); st.rerun()
            with s2:
                if st.button("🗑", key=f"dl_{code}", use_container_width=True): db["custom_products"] = [cp for cp in db.get("custom_products", []) if str(cp.get("code", "")) != code]; db.get("product_prices", {}).pop(code, None); save_database(db); st.rerun()

def render_admin_bookers():
    st.markdown("<h1 style='color:#0f172a;'>👤 Bookers</h1>", unsafe_allow_html=True); st.markdown("---")
    c1, c2 = st.columns([3, 1])
    with c1: nb = st.text_input("New Booker", key="new_bk", label_visibility="collapsed")
    with c2:
        if st.button("➕ Add", key="btn_add_bk", type="primary", use_container_width=True):
            name = nb.strip()
            if not name: st.error("❌ Name")
            elif name in db.get("bookers", []): st.warning("⚠️ Exists")
            else:
                db.setdefault("bookers", []).append(name); db["bookers"] = sorted(db["bookers"]); save_database(db)
                reg = load_bookers(); bk_u = sanitize_username(name)
                if bk_u not in reg: reg[bk_u] = {"display_name": name, "username": bk_u, "admin": CURRENT_USER, "password_hash": hash_password("1234"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}; save_bookers(reg)
                st.success(f"✅ '{name}' added | Password: 1234"); st.rerun()
    st.markdown("---")
    reg = load_bookers()
    for i, bk in enumerate(db.get("bookers", [])):
        c1, c2, c3 = st.columns([4, 1, 1])
        with c1:
            u = sanitize_username(bk); hl = u in reg and reg[u].get("admin") == CURRENT_USER
            st.markdown(f"<div style='background:#fff;border:1px solid #90caf9;border-radius:10px;padding:12px 18px;'><b style='color:#0f172a;'>{bk}</b> <span style='color:#64748b;font-size:12px;margin-left:12px;'>{'🔐 Login Active' if hl else '⚠️ No Login'}</span></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🔑 PW", key=f"pw_{i}_{bk}", use_container_width=True): st.session_state["setpw_bk"] = bk
        with c3:
            if st.button("🗑", key=f"delbk_{i}_{bk}", use_container_width=True):
                db["bookers"].remove(bk)
                save_database(db)
                u = sanitize_username(bk)
                if u in reg: del reg[u]; save_bookers(reg)
                st.rerun()
                
    target = st.session_state.get("setpw_bk")
    if target:
        st.markdown("---"); st.markdown(f"### 🔑 Set Password: {target}")
        np = st.text_input("New Password", type="password", key="setpw_new")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("💾 Save", key="setpw_sv", type="primary", use_container_width=True):
                if len(np) < 4: st.error("❌ Min 4")
                else:
                    reg = load_bookers(); u = sanitize_username(target)
                    if u not in reg: reg[u] = {"display_name": target, "username": u, "admin": CURRENT_USER, "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                    reg[u]["password_hash"] = hash_password(np); reg[u]["admin"] = CURRENT_USER; reg[u]["display_name"] = target
                    save_bookers(reg); st.session_state["setpw_bk"] = None; st.success(f"✅ Done"); st.rerun()
        with c2:
            if st.button("❌ Cancel", key="setpw_cx", use_container_width=True): st.session_state["setpw_bk"] = None; st.rerun()

def render_admin_salesmen():
    st.markdown("<h1 style='color:#0f172a;'>🧑‍💼 Salesmen</h1>", unsafe_allow_html=True); st.markdown("---")
    c1, c2 = st.columns([3, 1])
    with c1: ns = st.text_input("New Salesman", key="new_sm", label_visibility="collapsed")
    with c2:
        if st.button("➕ Add", key="btn_add_sm", type="primary", use_container_width=True):
            name = ns.strip()
            if not name: st.error("❌ Name")
            elif name in db.get("salesmen", []): st.warning("⚠️ Exists")
            else: db.setdefault("salesmen", []).append(name); db["salesmen"] = sorted(db["salesmen"]); save_database(db); st.rerun()
    st.markdown("---")
    for i, s in enumerate(db.get("salesmen", [])):
        c1, c2 = st.columns([5, 1])
        with c1: st.markdown(f"<div style='background:#fff;border:1px solid #90caf9;border-radius:10px;padding:12px 18px;'><b>{s}</b></div>", unsafe_allow_html=True)
        with c2:
            if st.button("🗑", key=f"delsm_{i}_{s}"): db["salesmen"].remove(s); save_database(db); st.rerun()

def render_admin_salaries(role_type):
    title = "💰 Bookers Salary" if role_type == "bookers" else "💰 Salesmen Salary"
    emoji = "👤" if role_type == "bookers" else "🧑‍💼"
    names = db.get("bookers", []) if role_type == "bookers" else db.get("salesmen", [])
    sk = "bookers_salaries" if role_type == "bookers" else "salesmen_salaries"
    st.markdown(f"<h1 style='color:#0f172a;'>{title}</h1>", unsafe_allow_html=True); st.markdown("---")
    if not names: st.info("None added."); return
    db.setdefault(sk, {})
    for pn in names:
        sd = db[sk].get(pn, {"base_salary": 0, "transactions": []}); base = sd.get("base_salary", 0); txns = sd.get("transactions", [])
        ap = sum(t["amount"] for t in txns if t.get("type") == "advanced" and t.get("status") == "pending"); sp = sum(t["amount"] for t in txns if t.get("type") == "shortage" and t.get("status") == "pending"); rem = base - ap - sp
        with st.expander(f"💰 {emoji} {pn} — Remaining: Rs {rem:,.0f}"):
            c1, c2 = st.columns([3, 1])
            with c1: nb = st.number_input("Base", value=float(base), min_value=0.0, step=500.0, key=f"b_{role_type}_{pn}")
            with c2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾", key=f"sb_{role_type}_{pn}"): db[sk].setdefault(pn, {"base_salary": 0, "transactions": []}); db[sk][pn]["base_salary"] = float(nb); save_database(db); st.rerun()
            st.markdown(f"Base: Rs {base:,.0f} | Adv: Rs {ap:,.0f} | Short: Rs {sp:,.0f} | Rem: Rs {rem:,.0f}")
            c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
            with c1: td = st.date_input("Date", value=date.today(), key=f"td_{role_type}_{pn}")
            with c2: tt = st.selectbox("Type", ["Advanced", "Shortage"], key=f"tt_{role_type}_{pn}")
            with c3: ta = st.number_input("Amount", min_value=0.0, step=100.0, key=f"ta_{role_type}_{pn}")
            with c4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("➕", key=f"at_{role_type}_{pn}"):
                    if ta <= 0: st.error("❌")
                    else:
                        db[sk].setdefault(pn, {"base_salary": base, "transactions": []}); nid = max([t.get("id", 0) for t in txns] + [0]) + 1
                        db[sk][pn]["transactions"].append({"id": nid, "date": td.strftime("%d-%m-%Y"), "time": datetime.now().strftime("%H:%M"), "type": "advanced" if tt == "Advanced" else "shortage", "amount": float(ta), "note": "", "status": "pending", "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}); save_database(db); st.rerun()
            for t in sorted(txns, key=lambda x: x.get("created_at", ""), reverse=True)[:10]:
                c1, c2, c3 = st.columns([4, 1, 1])
                with c1: st.markdown(f"📅 {t.get('date', '')} · {t.get('type', '')} · Rs {t['amount']:,.0f} · {t.get('status', '').upper()}")
                with c2:
                    if t.get("status") == "pending" and st.button("✅", key=f"pd_{role_type}_{pn}_{t['id']}"):
                        for tx in db[sk][pn]["transactions"]:
                            if tx.get("id") == t["id"]: tx["status"] = "paid"; break
                        save_database(db); st.rerun()
                with c3:
                    if st.button("🗑", key=f"dt_{role_type}_{pn}_{t['id']}"): db[sk][pn]["transactions"] = [x for x in db[sk][pn]["transactions"] if x.get("id") != t["id"]]; save_database(db); st.rerun()

def render_admin_expense():
    st.markdown("<h1 style='color:#0f172a;'>💵 Daily Expense</h1>", unsafe_allow_html=True); st.markdown("---")
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: fm = st.selectbox("Filter", ["📅 Aaj", "📆 This Month", "🗓️ Last 30 Days", "📋 All"], key="exp_filt")
    with c2: fd = st.date_input("From", value=today - timedelta(days=30), key="exp_from")
    with c3: td = st.date_input("To", value=today, key="exp_to")
    def dm(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if fm == "📅 Aaj": return d == today
        elif fm == "📆 This Month": return d.year == today.year and d.month == today.month
        elif fm == "🗓️ Last 30 Days": return today - timedelta(days=30) <= d <= today
        return True
    pf = [x for x in db.get("petrol_expenses", []) if dm(x.get("date", ""))]; lf = [x for x in db.get("lunch_expenses", []) if dm(x.get("date", ""))]; mf = [x for x in db.get("misc_expenses", []) if dm(x.get("date", ""))]
    tp = sum(float(x.get("amount", 0)) for x in pf); tl = sum(float(x.get("amount", 0)) for x in lf); tm = sum(float(x.get("amount", 0)) for x in mf)
    st.markdown(f"<div style='background:#fff;border:2px solid #2196f3;border-radius:12px;padding:12px;'><b>Petrol:</b> Rs {tp:,.0f} | <b>Lunch:</b> Rs {tl:,.0f} | <b>Misc:</b> Rs {tm:,.0f} | <b>TOTAL:</b> Rs {tp+tl+tm:,.0f}</div>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**⛽ Petrol**")
        c1a, c1b, c1c = st.columns([2, 2, 1])
        with c1a: ps = st.selectbox("Salesman", ["-- Select --"] + db.get("salesmen", []), key="ps")
        with c1b: pa = st.number_input("Amount", min_value=0.0, step=50.0, key="pa")
        with c1c:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕", key="add_pt"):
                if ps == "-- Select --" or pa <= 0: st.error("❌")
                else: nid = max([x.get("id", 0) for x in db.get("petrol_expenses", [])] + [0]) + 1; db.setdefault("petrol_expenses", []).append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "time": datetime.now().strftime("%H:%M"), "salesman": ps, "amount": float(pa)}); save_database(db); st.rerun()
        for x in sorted(pf, key=lambda y: y.get("created_at", ""), reverse=True)[:8]:
            cA, cB = st.columns([5, 1])
            with cA: st.markdown(f"📅 {x.get('date', '')} · {x.get('salesman', '')} · Rs {float(x.get('amount', 0)):,.0f}")
            with cB:
                if st.button("🗑", key=f"dpt_{x.get('id')}"): db["petrol_expenses"] = [y for y in db["petrol_expenses"] if y.get("id") != x.get("id")]; save_database(db); st.rerun()
    with c2:
        st.markdown("**🍽️ Lunch**")
        c2a, c2b = st.columns([3, 1])
        with c2a: la = st.number_input("Amount", min_value=0.0, step=50.0, key="la")
        with c2b:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("➕", key="add_ln"):
                if la <= 0: st.error("❌")
                else: nid = max([x.get("id", 0) for x in db.get("lunch_expenses", [])] + [0]) + 1; db.setdefault("lunch_expenses", []).append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "time": datetime.now().strftime("%H:%M"), "amount": float(la)}); save_database(db); st.rerun()
        for x in sorted(lf, key=lambda y: y.get("created_at", ""), reverse=True)[:8]:
            cA, cB = st.columns([5, 1])
            with cA: st.markdown(f"📅 {x.get('date', '')} · Rs {float(x.get('amount', 0)):,.0f}")
            with cB:
                if st.button("🗑", key=f"dln_{x.get('id')}"): db["lunch_expenses"] = [y for y in db["lunch_expenses"] if y.get("id") != x.get("id")]; save_database(db); st.rerun()
    st.markdown("---")
    st.markdown("### 🧾 Misc")
    mc1, mc2, mc3, mc4, mc5 = st.columns([2, 2, 1.5, 2, 1])
    with mc1: md = st.text_input("Item", key="mi")
    with mc2: mn = st.text_input("Note", key="mn")
    with mc3: ma = st.number_input("Amt", min_value=0.0, step=50.0, key="ma")
    with mc4: mdt = st.date_input("Date", value=today, key="md")
    with mc5:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➕", key="add_mc", type="primary", use_container_width=True):
            if not md.strip() or ma <= 0: st.error("❌")
            else: nid = max([x.get("id", 0) for x in db.get("misc_expenses", [])] + [0]) + 1; db.setdefault("misc_expenses", []).append({"id": nid, "date": mdt.strftime("%d-%m-%Y"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "time": datetime.now().strftime("%H:%M"), "description": md.strip(), "note": mn.strip(), "amount": float(ma)}); save_database(db); st.rerun()
    for x in sorted(mf, key=lambda y: y.get("created_at", ""), reverse=True):
        cA, cB = st.columns([6, 1])
        with cA: st.markdown(f"📅 {x.get('date', '')} · {x.get('description', '')} · Rs {float(x.get('amount', 0)):,.0f}")
        with cB:
            if st.button("🗑", key=f"dmc_{x.get('id')}"): db["misc_expenses"] = [y for y in db["misc_expenses"] if y.get("id") != x.get("id")]; save_database(db); st.rerun()

def render_admin_billing():
    st.markdown("<h1 style='color:#0f172a;'>🧾 Billing</h1>", unsafe_allow_html=True); st.markdown("---")
    c1, c2 = st.columns(2)
    with c1: st.text_input("Bill No", value=str(db["next_bill_no"]), disabled=True)
    with c2: st.text_input("Date", value=datetime.now().strftime("%d-%m-%Y"), disabled=True)
    c1, c2, c3 = st.columns(3)
    with c1: shop = st.text_input("Shop", key="shop_name")
    with c2:
        bks = db.get("bookers", [])
        if bks: sel = st.selectbox("Booker", ["--"] + bks, key="ob_sel"); st.session_state["order_booker"] = "" if sel == "--" else sel
        else: st.text_input("Booker", key="order_booker")
    with c3:
        sms = db.get("salesmen", [])
        if sms: sel = st.selectbox("Salesman", ["--"] + sms, key="sm_sel"); st.session_state["salesman"] = "" if sel == "--" else sel
        else: st.text_input("Salesman", key="salesman")
    c1, c2 = st.columns([1, 3])
    with c1: s = st.text_input("🔍", key="search_text")
    with c2:
        allp = get_all_products(); su = s.strip().upper()
        fl = [p["name"] for p in allp if su in p["name"].upper()] if su else [p["name"] for p in allp]
        if st.session_state.get("product_sel") and st.session_state["product_sel"] not in fl: st.session_state["product_sel"] = ""
        ps = st.selectbox("Product", [""] + fl, key="product_sel")
    sp = next((p for p in get_all_products() if p["name"] == ps), None); tpd = get_price(sp["code"], sp["price"]) if sp else 0.0
    if st.session_state["_prev_prod"] != ps: st.session_state["tp_box"] = float(tpd); st.session_state["_prev_prod"] = ps
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: bx = st.number_input("Boxes", min_value=0, step=1, key="boxes")
    with c2: tp = st.number_input("TP", min_value=0.0, step=1.0, key="tp_box")
    with c3: dc = st.number_input("Disc%", min_value=0.0, step=0.5, key="discount")
    g = bx * tp; n = g - (g * dc / 100)
    with c4: st.text_input("Gross", value=f"{g:.0f}", disabled=True)
    with c5: st.text_input("Net", value=f"{n:.0f}", disabled=True)
    b1, b2 = st.columns(2)
    with b1: st.button("➕ Add Bill", key="btn_add", on_click=add_bill_callback, type="primary", use_container_width=True)
    with b2: st.button("🔄 Refresh", key="btn_refresh", on_click=refresh_callback, use_container_width=True)
    e1, e2, e3 = st.columns(3)
    with e1: st.button("📄 Export", key="btn_export", on_click=export_bill_callback, use_container_width=True)
    with e2: st.button("📦 Export LF", key="btn_exp_lf", on_click=export_load_form_from_billing_callback, use_container_width=True)
    with e3: st.button("🔄 Refresh LF", key="btn_ref_lf", on_click=refresh_load_form_callback, use_container_width=True)
    show_auto_download()

def render_admin_bills_list():
    st.markdown("<h1 style='color:#0f172a;'>📋 Bills List</h1>", unsafe_allow_html=True); st.markdown("---")
    if not db["bills"]: st.info("No bills."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: fm = st.selectbox("Filter", ["📅 Aaj", "📆 Custom", "🗓️ Specific", "📋 All"], key="bl_filt")
    with c2: fd = st.date_input("From", value=today - timedelta(days=7), key="bl_from")
    with c3: td = st.date_input("To", value=today, key="bl_to")
    cm = {(cb.get("shop", ""), cb.get("date", ""), cb.get("booker", ""), cb.get("bill_no", "")): cb for cb in db.get("credit_bills", [])}
    groups = {}
    for bi, b in enumerate(db["bills"]):
        bd = parse_date(b.get("Date", ""))
        if fm == "📅 Aaj" and bd != today: continue
        elif fm == "🗓️ Specific" and bd != fd: continue
        elif fm == "📆 Custom" and (bd is None or not (fd <= bd <= td)): continue
        key = (b.get("Shop", ""), b.get("Date", ""), b.get("Order Booker", ""), b.get("Bill No", ""))
        if key not in groups: groups[key] = {"shop": b.get("Shop", ""), "date": b.get("Date", ""), "booker": b.get("Order Booker", ""), "salesman": b.get("Salesman", ""), "bill_no": b.get("Bill No", ""), "items": [], "idx": []}
        groups[key]["items"].append({"Code": b.get("Code"), "Product": b.get("Product"), "Boxes": b.get("Boxes"), "TP/Box": b.get("TP/Box"), "Discount %": b.get("Discount %"), "Gross": b.get("Gross"), "Net": b.get("Net"), "Salesman": b.get("Salesman", "")}); groups[key]["idx"].append(bi)
    st.markdown(f"### Bills ({len(groups)})")
    for key, g in sorted(groups.items(), key=lambda kv: kv[0][1], reverse=True):
        tb = sum(int(it.get("Boxes", 0)) for it in g["items"]); ck = (g["shop"] or "", g["date"] or "", g["booker"] or "", g["bill_no"] or ""); ec = cm.get(ck); cbadge = ""
        if ec: cbadge = '<span style="background:#c8e6c9;color:#1b5e20;padding:3px 10px;border-radius:6px;font-size:11px;font-weight:700;margin-left:8px;">✅ PAID</span>' if ec.get("status") == "paid" else '<span style="background:#ff9800;color:#fff;padding:3px 10px;border-radius:6px;font-size:11px;font-weight:700;margin-left:8px;">💳 PENDING</span>'
        st.markdown(f"""<div style="background:#fff;border-left:6px solid #2196f3;border-radius:12px;padding:12px 18px;margin-bottom:8px;display:flex;justify-content:space-between;"><div><div style="font-weight:700;color:#1976d2;">🏪 {g['shop'] or '-'} {cbadge}</div><div style="font-size:12px;color:#666;">📅 {g['date'] or '-'} · 👤 {g['booker'] or '-'} · 🧑‍💼 {g['salesman'] or '-'} · 🧾 #{g['bill_no']}</div></div><div style="background:#e3f2fd;color:#1976d2;font-weight:800;padding:6px 14px;border-radius:8px;">{tb} BOXES</div></div>""", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("⬇️ Excel", key=f"dl_{key}_{g['date']}"):
                tot = sum(float(it.get("Net", 0)) for it in g["items"]); pct, pn, tl = get_effective_discount_pct(g["shop"], tot)
                ss = g["shop"].replace("/", "-").replace(" ", "_").replace(":", ""); sb = (g["booker"] or "bill").replace(" ", "_").replace("/", "-"); sd = (g["date"] or "").replace("-", "")
                st.session_state["download_file"] = (f"{ss}_{sb}_{sd}.xlsx", _build_bill_excel(g["shop"], g["items"], pct, pn, tl, booker=g["booker"], date_str=g["date"])); st.rerun()
        with c2:
            if ec: st.button("💳 Already", key=f"cr_{key}_{g['date']}", disabled=True)
            else:
                if st.button("💳 Mark Credit", key=f"cr_{key}_{g['date']}"): move_group_to_credit(g); st.rerun()
        with c3:
            if st.button("🗑 Delete", key=f"del_{key}_{g['date']}"): idxs = set(g["idx"]); db["bills"] = [b for i, b in enumerate(db["bills"]) if i not in idxs]; save_database(db); st.rerun()
        st.markdown("---")
    show_auto_download()

def render_admin_credit():
    st.markdown("<h1 style='color:#0f172a;'>💳 Credit Bills</h1>", unsafe_allow_html=True); st.markdown("---")
    cbs = db.get("credit_bills", [])
    if not cbs: st.info("None."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: fm = st.selectbox("Filter", ["📋 All", "📅 Aaj", "📆 Custom", "🗓️ Specific"], key="cfilt")
    with c2: fd = st.date_input("From", value=today - timedelta(days=30), key="cfrom")
    with c3: td = st.date_input("To", value=today, key="cto")
    fl = []
    for c in cbs:
        cd = parse_date(c.get("date", ""))
        if fm == "📅 Aaj" and cd != today: continue
        elif fm == "🗓️ Specific" and cd != fd: continue
        elif fm == "📆 Custom" and (cd is None or not (fd <= cd <= td)): continue
        fl.append(c)
    pend = [c for c in fl if c.get("status") == "pending"]; paid = [c for c in fl if c.get("status") == "paid"]
    st.markdown(f"<div style='background:#fff;border:2px solid #2196f3;border-radius:12px;padding:12px;'><b>Pending:</b> {len(pend)} (Rs {sum(float(c.get('total_net', 0)) for c in pend):,.0f}) | <b>Paid:</b> {len(paid)} (Rs {sum(float(c.get('total_net', 0)) for c in paid):,.0f})</div>", unsafe_allow_html=True)
    for c in pend + paid:
        cid = c["id"]; isp = c.get("status") == "paid"; badge = "✅ PAID" if isp else "⏳ PENDING"
        st.markdown(f"🏪 **{c.get('shop', '')}** · {badge} · 📅 {c.get('date', '')} · 👤 {c.get('booker', '')} · Rs {float(c.get('total_net', 0)):,.0f}")
        c1, c2 = st.columns(2)
        with c1:
            if not isp:
                if st.button("✅ Mark Paid", key=f"pay_{cid}", type="primary"):
                    for cc in db["credit_bills"]:
                        if cc.get("id") == cid: cc["status"] = "paid"; cc["paid_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); break
                    save_database(db); st.rerun()
            else:
                if st.button("↩️ Pending", key=f"unpay_{cid}"):
                    for cc in db["credit_bills"]:
                        if cc.get("id") == cid: cc["status"] = "pending"; cc["paid_at"] = None; break
                    save_database(db); st.rerun()
        with c2:
            if st.button("🗑", key=f"dcr_{cid}"): db["credit_bills"] = [x for x in db["credit_bills"] if x.get("id") != cid]; save_database(db); st.rerun()

def render_admin_load_form():
    st.markdown("<h1 style='color:#0f172a;'>📦 Load Forms</h1>", unsafe_allow_html=True); st.markdown("---")
    lfs = db.get("load_forms", [])
    if not lfs: st.info("No load forms."); return
    today = date.today(); ao = st.checkbox("☑️ Show old dates", value=False, key="lf_ao")
    c1, c2, c3 = st.columns(3)
    with c1: fm = st.selectbox("Filter", ["📅 Aaj", "📆 Custom", "🗓️ Specific", "📋 All"], key="lf_filt")
    with c2: fd = st.date_input("From", value=today, key="lf_from")
    with c3: td = st.date_input("To", value=today, key="lf_to")
    def ir(dstr):
        d = parse_date(dstr)
        if d is None: return False
        if ao: return True
        if fm == "📅 Aaj": return d == today
        elif fm == "🗓️ Specific": return d == fd
        elif fm == "📆 Custom": return fd <= d <= td
        return True
    bl = sorted({lf.get("booker", "") for lf in lfs if lf.get("booker")})
    for bkn in bl:
        blfs = [lf for lf in lfs if lf.get("booker") == bkn and ir(lf.get("date", ""))]
        tb = sum(int(lf.get("total_boxes", 0)) for lf in blfs); ids = {lf.get("id") for lf in blfs}; tr = 0
        for r in db.get("booker_returns", []):
            if r.get("load_form_id") not in ids: continue
            it = r.get("items", []); tr += sum(int(x.get("ReturnBoxes", 0)) for x in it) if it else int(r.get("boxes", 0))
        nb = tb - tr
        with st.expander(f"👤 {bkn} — LFs: {len(blfs)} | Boxes: {tb} | Ret: {tr} | Net: {nb}"):
            for lf in sorted(blfs, key=lambda x: x.get("created_at", ""), reverse=True):
                lfid = lf.get("id"); dt = lf.get("date", ""); ltb = lf.get("total_boxes", 0); items = lf.get("items", []); trf = lf.get("transferred_to_dsr", False)
                rmap = get_lf_return_items(lfid); rt = sum(rmap.values()); nl = ltb - rt; badge = f'✅ DSR #{lf.get("dsr_id")}' if trf else '🟢 Active'
                st.markdown(f"📦 **LF #{lfid}** — {badge} · 📅 {dt} · {len(items)} products · {ltb} boxes · Ret: {rt} · Net: {nl}")
                sms = db.get("salesmen", [])
                if sms:
                    cs = lf.get("salesman", ""); di = 0
                    if cs in sms: di = sms.index(cs) + 1
                    sels = st.selectbox("Salesman", ["--"] + sms, index=di, key=f"sm_sel_{lfid}")
                    if sels != "--" and sels != cs: lf["salesman"] = sels; save_database(db); st.rerun()
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    if st.button("↩️ Return", key=f"ret_{lfid}", use_container_width=True): st.session_state[f"rs_{lfid}"] = not st.session_state.get(f"rs_{lfid}", False); st.rerun()
                with c2:
                    if not trf:
                        if st.button("📤 DSR", key=f"trf_{lfid}", type="primary", use_container_width=True): transfer_load_form_to_dsr(lfid); st.rerun()
                    else: st.button(f"✅ DSR #{lf.get('dsr_id', '')}", key=f"trfd_{lfid}", disabled=True, use_container_width=True)
                with c3:
                    if st.button("📥 Excel", key=f"dlx_{lfid}", use_container_width=True):
                        lfo = next((x for x in db.get("load_forms", []) if x.get("id") == lfid), None)
                        if lfo: sb = str(lfo.get("booker", "bk")).replace(" ", "_").replace("/", "-"); sd = str(lfo.get("date", "")).replace("-", ""); st.session_state["download_file"] = (f"LF_{lfid}_{sb}_{sd}.xlsx", build_load_form_excel(lfo)); st.rerun()
                with c4:
                    if st.button("🗑", key=f"dlf_{lfid}", use_container_width=True): db["load_forms"] = [x for x in db["load_forms"] if x.get("id") != lfid]; db["booker_returns"] = [r for r in db.get("booker_returns", []) if r.get("load_form_id") != lfid]; save_database(db); st.rerun()
                if st.session_state.get(f"rs_{lfid}", False):
                    er = get_return_record_for_lf(lfid); em = {}
                    if er:
                        for it in er.get("items", []): em[str(it.get("Code", ""))] = int(it.get("ReturnBoxes", 0))
                    ri = {}
                    for idx, it in enumerate(items):
                        code = str(it.get("Code", "")); ld = int(it.get("Boxes", 0))
                        cc1, cc2, cc3, cc4 = st.columns([1, 4, 2, 2])
                        with cc1: st.markdown(f"**{code}**")
                        with cc2: st.markdown(it.get("Product", ""))
                        with cc3: st.markdown(f"📦 {ld}")
                        with cc4: ri[code] = st.number_input("ret", min_value=0, max_value=ld, value=int(em.get(code, 0)), step=1, key=f"rq_{lfid}_{code}_{idx}", label_visibility="collapsed")
                    rc1, rc2, rc3 = st.columns([2, 3, 1])
                    with rc1:
                        dd = today
                        if er:
                            try: dd = datetime.strptime(er.get("date", ""), "%d-%m-%Y").date()
                            except Exception: dd = today
                        rdt = st.date_input("Date", value=dd, key=f"rd_{lfid}")
                    with rc2: rnt = st.text_input("Note", value=er.get("note", "") if er else "", key=f"rn_{lfid}")
                    with rc3:
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.button("💾 Save", key=f"rsv_{lfid}", type="primary", use_container_width=True):
                            ni = []; tn = 0
                            for it in items:
                                c = str(it.get("Code", "")); q = int(ri.get(c, 0))
                                if q > 0: ni.append({"Code": c, "Product": it.get("Product", ""), "ReturnBoxes": q}); tn += q
                            if er:
                                er["items"] = ni; er["date"] = rdt.strftime("%d-%m-%Y"); er["note"] = rnt.strip(); er["total_return_boxes"] = tn; er["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            else:
                                nid = max([r.get("id", 0) for r in db.get("booker_returns", [])] + [0]) + 1
                                db.setdefault("booker_returns", []).append({"id": nid, "booker": bkn, "load_form_id": lfid, "date": rdt.strftime("%d-%m-%Y"), "items": ni, "total_return_boxes": tn, "note": rnt.strip(), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                            save_database(db); st.session_state[f"rs_{lfid}"] = False; st.rerun()
                    if er:
                        if st.button("🗑 Return", key=f"dret_{lfid}"): db["booker_returns"] = [x for x in db["booker_returns"] if x.get("id") != er.get("id")]; save_database(db); st.session_state[f"rs_{lfid}"] = False; st.rerun()
                st.markdown("---")

def render_admin_dsr():
    st.markdown("<h1 style='color:#0f172a;'>📋 DSR</h1>", unsafe_allow_html=True); st.markdown("---")
    ds = db.get("dsr_forms", [])
    if not ds: st.info("None."); return
    today = date.today(); c1, c2, c3 = st.columns(3)
    with c1: fm = st.selectbox("Filter", ["📅 Aaj", "📆 Custom", "🗓️ Specific", "📋 All"], key="dfilt")
    with c2: fd = st.date_input("From", value=today - timedelta(days=7), key="dfrom")
    with c3: td = st.date_input("To", value=today, key="dto")
    fl = []
    for d in ds:
        dd = parse_date(d.get("date", ""))
        if dd is None: continue
        if fm == "📅 Aaj" and dd != today: continue
        elif fm == "🗓️ Specific" and dd != fd: continue
        elif fm == "📆 Custom" and not (fd <= dd <= td): continue
        fl.append(d)
    for i, d in enumerate(sorted(fl, key=lambda x: x.get("created_at", ""), reverse=True)):
        did = d.get("id", i); ct, _, _ = get_dsr_credit_info(d); fin = float(d.get("amount_to_collect", 0)) - ct
        st.markdown(f"📋 **DSR #{did}** — {d.get('booker', '?')} · 📅 {d.get('date', '')} · Rs {float(d.get('total_amount', 0)):,.0f} · Final: **Rs {fin:,.0f}**")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📤 Excel", key=f"dx_{did}", type="primary"): export_dsr_excel(d); st.rerun()
        with c2:
            if st.button("🗑", key=f"dd_{did}"):
                db["dsr_forms"] = [x for x in db["dsr_forms"] if x.get("id") != did]
                for x in db.get("load_forms", []):
                    if x.get("dsr_id") == did: x["transferred_to_dsr"] = False; x.pop("dsr_id", None)
                save_database(db); st.rerun()
        st.markdown("---")

def render_admin_calculation():
    st.markdown("<h1 style='color:#0f172a;'>🧮 Cash Calculation</h1>", unsafe_allow_html=True); st.markdown("---")
    if "calc_reset_token" not in st.session_state: st.session_state["calc_reset_token"] = 0
    total = 0
    for section, denoms in [("💵 Notes", [5000, 1000, 500, 100, 50, 20, 10]), ("🪙 Coins", [5, 2, 1])]:
        st.markdown(f"### {section}")
        for d in denoms:
            c1, c2, c3 = st.columns([2, 2, 3])
            with c1: st.markdown(f"**₹{d}**")
            with c2: q = st.number_input("Q", min_value=0, step=1, value=0, key=f"calc_{d}_{st.session_state['calc_reset_token']}", label_visibility="collapsed")
            with c3: amt = q * d; total += amt; st.markdown(f"**Rs {amt:,.0f}**")
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💰 Calculate", type="primary", use_container_width=True): st.session_state["calc_locked"] = total
    with c2:
        if st.button("🔄 Reset", use_container_width=True): st.session_state["calc_reset_token"] += 1; st.session_state["calc_locked"] = 0; st.rerun()
    st.markdown(f"<div style='background:#0f172a;border-radius:16px;padding:24px;text-align:center;color:white;margin-top:16px;'><div style='font-size:12px;letter-spacing:2px;font-weight:700;color:#94a3b8;'>LIVE</div><div style='font-size:42px;font-weight:800;'>Rs {total:,.0f}</div></div>", unsafe_allow_html=True)
    lk = st.session_state.get("calc_locked", 0)
    if lk and lk > 0: st.markdown(f"<div style='background:#16a34a;border-radius:16px;padding:24px;text-align:center;color:white;margin-top:10px;'><div style='font-size:12px;letter-spacing:2px;font-weight:700;color:#dcfce7;'>🔒 LOCKED</div><div style='font-size:42px;font-weight:800;'>Rs {lk:,.0f}</div></div>", unsafe_allow_html=True)

# ============================================================
# CALLBACKS
# ============================================================
def add_bill_callback():
    ps = st.session_state.get("product_sel", ""); bx = st.session_state.get("boxes", 0); tp = st.session_state.get("tp_box", 0.0); dc = st.session_state.get("discount", 0.0)
    if not ps or bx <= 0: st.session_state["error_msg"] = "❌ Select product and boxes"; return
    sel = next((p for p in get_all_products() if p["name"] == ps), None)
    if not sel: st.session_state["error_msg"] = "❌ Invalid"; return
    g = bx * tp; n = g - (g * dc / 100)
    db["bills"].append({"Bill No": db["next_bill_no"], "Date": datetime.now().strftime("%d-%m-%Y"), "Shop": st.session_state.get("shop_name", "").strip(), "Order Booker": st.session_state.get("order_booker", "").strip(), "Salesman": st.session_state.get("salesman", "").strip(), "Delivery Man": "", "Code": sel["code"], "Product": sel["name"], "Boxes": bx, "TP/Box": tp, "Discount %": dc, "Gross": g, "Net": n})
    save_database(db); st.session_state["success_msg"] = f"✅ Bill #{db['next_bill_no']}"
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None

def refresh_callback():
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None; st.session_state["success_msg"] = "✅ Cleared"

def export_bill_callback():
    if not db["bills"]: st.session_state["error_msg"] = "❌ No bills"; return
    shop = st.session_state.get("shop_name", "").strip() or "Bill"
    for ch in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']: shop = shop.replace(ch, "")
    sb = [b for b in db["bills"] if b["Shop"].strip() == shop]
    if not sb: st.session_state["error_msg"] = "❌ No bills for shop"; return
    tot = sum(float(b.get("Net", 0)) for b in sb); pct, pn, tl = get_effective_discount_pct(shop, tot)
    bv = sb[0].get("Order Booker", "") if sb else ""; dv = sb[0].get("Date", "") if sb else ""
    st.session_state["download_file"] = (f"{shop}_{(bv or 'bill').replace(' ', '_')}.xlsx", _build_bill_excel(shop, sb, pct, pn, tl, booker=bv, date_str=dv)); st.session_state["success_msg"] = f"✅ Exported"

def export_load_form_from_billing_callback():
    bk = st.session_state.get("order_booker", "").strip()
    if not bk: st.session_state["error_msg"] = "❌ Select booker"; return
    bbs = [b for b in db["bills"] if b["Order Booker"].strip() == bk]
    if not bbs: st.session_state["error_msg"] = "❌ No bills"; return
    sm = {}
    for b in bbs:
        c = b["Code"]
        if c not in sm: sm[c] = {"Code": c, "Product": b["Product"], "Boxes": 0}
        sm[c]["Boxes"] += b["Boxes"]
    items = list(sm.values()); tb = sum(it["Boxes"] for it in items); existing = None
    for lf in reversed(db.get("load_forms", [])):
        if (str(lf.get("booker", "")).strip() == bk and not lf.get("transferred_to_dsr") and not lf.get("refreshed")): existing = lf; break
    if existing:
        existing["items"] = items; existing["total_boxes"] = tb; existing["total_products"] = len(items); existing["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S"); msg = f"🔄 LF #{existing['id']} updated"
    else:
        nid = max([lf.get("id", 0) for lf in db.get("load_forms", [])] + [0]) + 1
        db.setdefault("load_forms", []).append({"id": nid, "date": datetime.now().strftime("%d-%m-%Y"), "time": datetime.now().strftime("%H:%M"), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "booker": bk, "items": items, "total_boxes": tb, "total_products": len(items), "transferred_to_dsr": False, "refreshed": False}); msg = f"✅ LF #{nid}"
    save_database(db); st.session_state["success_msg"] = msg

def refresh_load_form_callback():
    bk = st.session_state.get("order_booker", "").strip()
    if not bk: st.session_state["error_msg"] = "❌"; return
    db["bills"] = [b for b in db["bills"] if b["Order Booker"].strip() != bk]
    for lf in db.get("load_forms", []):
        if (str(lf.get("booker", "")).strip() == bk and not lf.get("transferred_to_dsr") and not lf.get("refreshed")): lf["refreshed"] = True; lf["refreshed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_database(db)
    for k in ["search_text", "product_sel"]: st.session_state[k] = ""
    for k in ["boxes", "tp_box", "discount"]: st.session_state[k] = 0
    st.session_state["_prev_prod"] = None; st.session_state["success_msg"] = f"✅ Cleared"

def move_group_to_credit(group):
    shop = group.get("shop", "") or "-"; dt = group.get("date", "") or "-"; bk = group.get("booker", "") or "-"; sm = group.get("salesman", "") or "-"; bn = group.get("bill_no", ""); items = group.get("items", [])
    for cb in db.get("credit_bills", []):
        if (cb.get("shop") == shop and cb.get("date") == dt and cb.get("booker") == bk and cb.get("bill_no") == bn): st.session_state["error_msg"] = "⚠️ Already in credit"; return
    tb = sum(int(it.get("Boxes", 0)) for it in items); tg = sum(float(it.get("Gross", 0)) for it in items); tn = sum(float(it.get("Net", 0)) for it in items); nid = db.get("next_credit_id", 1)
    db.setdefault("credit_bills", []).append({"id": nid, "shop": shop, "date": dt, "booker": bk, "salesman": sm, "bill_no": bn, "items": [dict(it) for it in items], "total_boxes": tb, "total_gross": tg, "total_net": tn, "status": "pending", "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "paid_at": None})
    db["next_credit_id"] = nid + 1; save_database(db); st.session_state["success_msg"] = f"💳 Credit #{nid}"

def transfer_load_form_to_dsr(lf_id):
    lf = next((x for x in db.get("load_forms", []) if x.get("id") == lf_id), None)
    if not lf or lf.get("transferred_to_dsr"): st.session_state["error_msg"] = "❌ Invalid"; return
    bk = lf.get("booker", ""); items = lf.get("items", []); allp = get_all_products(); sm = lf.get("salesman", ""); rmap = get_lf_return_items(lf_id); di = []; tb = ta = rt = ra = 0
    for it in items:
        c = str(it.get("Code", "")); pr = next((p for p in allp if str(p["code"]) == c), None)
        try: price = get_price(c, pr["price"]) if pr else 0.0
        except Exception: price = 0.0
        bx = int(it.get("Boxes", 0)); rbx = min(int(rmap.get(c, 0)), bx); lt = bx * float(price); rlt = rbx * float(price)
        di.append({"Code": c, "Product": it.get("Product", ""), "Boxes": bx, "TP/Box": float(price), "Total": lt, "ReturnBoxes": rbx, "ReturnAmount": rlt})
        tb += bx; ta += lt; rt += rbx; ra += rlt
    bbs = [b for b in db.get("bills", []) if b.get("Order Booker", "").strip() == bk]; sg = {}; src = []
    for b in bbs:
        s = b.get("Shop", "-") or "-"; bn = b.get("Bill No", ""); net = float(b.get("Net", 0)); gr = float(b.get("Gross", 0))
        if s not in sg: sg[s] = {"shop": s, "gross": 0.0, "net": 0.0, "individual_discount": 0.0, "bill_nos": set()}
        sg[s]["gross"] += gr; sg[s]["net"] += net; sg[s]["individual_discount"] += (gr - net); sg[s]["bill_nos"].add(bn); src.append({"shop": s, "bill_no": bn, "net": net})
    sds = []; td = 0.0
    for s, g in sg.items():
        pct, pn, _ = get_effective_discount_pct(s, g["net"]); pd = g["net"] * pct / 100; std = g["individual_discount"] + pd
        sds.append({"shop": s, "gross": g["gross"], "net": g["net"], "individual_discount": g["individual_discount"], "package_pct": pct, "package_name": pn or "", "package_discount": pd, "total_discount": std, "bill_nos": sorted([str(x) for x in g["bill_nos"]])}); td += std
    na = ta - ra; nid = max([x.get("id", 0) for x in db.get("dsr_forms", [])] + [0]) + 1
    db.setdefault("dsr_forms", []).append({"id": nid, "source_load_form_id": lf_id, "booker": bk, "salesman": sm, "date": lf.get("date", ""), "time": lf.get("time", ""), "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "items": di, "total_boxes": tb, "total_amount": ta, "total_return_boxes": rt, "total_return_amount": ra, "net_amount": na, "shop_discounts": sds, "total_discount": td, "amount_to_collect": na - td, "source_bills": src, "status": "open", "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
    for x in db["load_forms"]:
        if x.get("id") == lf_id: x["transferred_to_dsr"] = True; x["dsr_id"] = nid; break
    save_database(db); st.session_state["success_msg"] = f"✅ DSR #{nid}"

# ============================================================
# ROUTER — ADMIN
# ============================================================
_page = st.session_state["page"]
if _page == "📊 Dashboard": render_admin_dashboard()
elif _page == "🧾 Billing": render_admin_billing()
elif _page == "🛒 All Products": render_admin_products()
elif _page == "🎁 Discount": render_admin_discount()
elif _page == "👤 Bookers": render_admin_bookers()
elif _page == "📈 Booker Bills": render_admin_booker_bills()
elif _page == "💰 Bookers Salary": render_admin_salaries("bookers")
elif _page == "🧑‍💼 Salesmen": render_admin_salesmen()
elif _page == "💰 Salesmen Salary": render_admin_salaries("salesmen")
elif _page == "💵 Daily Expense": render_admin_expense()
elif _page == "📋 Bills List": render_admin_bills_list()
elif _page == "💳 Credit Bills": render_admin_credit()
elif _page == "📦 Load Form": render_admin_load_form()
elif _page == "📋 DSR": render_admin_dsr()
elif _page == "🧮 Calculation": render_admin_calculation()

for k, fn in [("success_msg", st.success), ("error_msg", st.error), ("cloud_error", st.error)]:
    if st.session_state.get(k): fn(st.session_state[k]); st.session_state[k] = None
show_auto_download()
