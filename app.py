# ============================================================
# AL-BARAKAH ENTERPRISES - BILLING SOFTWARE 2026
# + Premium Animated Lamp Login + Multi-User System
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
# LAMP HTML — Premium Pendant Lamp + Login Form Inside Light
# ============================================================
def build_lamp_html(light_on):
    initial_class = "lit" if light_on else ""
    return """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  html, body {
    width: 100%;
    height: 100vh;
    background: #000;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    overflow: hidden;
    user-select: none;
  }
  .scene {
    position: relative;
    width: 100%;
    height: 100vh;
    background: radial-gradient(ellipse at 50% 0%, #0a0805 0%, #000 60%);
    overflow: hidden;
  }
  .scene.lit {
    background: radial-gradient(ellipse at 50% 15%, #1f1408 0%, #0a0604 40%, #000 75%);
    transition: background 1.2s ease;
  }

  /* ========= CEILING ========= */
  .ceiling {
    position: absolute;
    top: 0; left: 50%;
    transform: translateX(-50%);
    width: 56px; height: 10px;
    background: linear-gradient(180deg, #1a1a1a 0%, #050505 100%);
    border-radius: 0 0 6px 6px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.9);
    z-index: 5;
  }
  .ceiling::after {
    content: '';
    position: absolute;
    top: 0; left: 50%;
    transform: translateX(-50%);
    width: 22px; height: 3px;
    background: #2a2a2a;
    border-radius: 2px;
  }

  /* ========= WIRE ========= */
  .wire {
    position: absolute;
    top: 8px; left: 50%;
    transform: translateX(-50%);
    width: 3px; height: 130px;
    background: linear-gradient(90deg, #3a3a3a 0%, #6b6b6b 50%, #1a1a1a 100%);
    box-shadow: 0 0 4px rgba(0,0,0,0.8);
    z-index: 4;
  }

  /* ========= LAMP ASSEMBLY ========= */
  .lamp {
    position: absolute;
    top: 130px; left: 50%;
    transform: translateX(-50%);
    width: 220px; height: 130px;
    z-index: 6;
  }
  .lamp-shade {
    position: absolute;
    top: 0; left: 0;
    width: 220px; height: 100px;
    border-radius: 110px 110px 20px 20px / 100px 100px 22px 22px;
    background:
      radial-gradient(ellipse at 50% -10%, #4a3826 0%, #2b1d10 30%, #0f0a05 85%),
      linear-gradient(180deg, #2a1d10 0%, #0a0603 100%);
    box-shadow:
      inset 0 -18px 30px rgba(0,0,0,0.95),
      inset 0 4px 12px rgba(120, 80, 40, 0.25),
      0 6px 22px rgba(0,0,0,0.85);
    transition: box-shadow 0.9s ease;
    z-index: 3;
  }
  .lamp-shade::before {
    content: '';
    position: absolute;
    top: 6px; left: 50%;
    transform: translateX(-50%);
    width: 90%; height: 30px;
    border-radius: 50%;
    background: radial-gradient(ellipse, rgba(160,110,60,0.25) 0%, transparent 70%);
    pointer-events: none;
  }
  .lamp-shade::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0;
    width: 100%; height: 22px;
    border-radius: 0 0 22px 22px / 0 0 22px 22px;
    background: linear-gradient(180deg, transparent 0%, #050301 100%);
  }
  .scene.lit .lamp-shade {
    box-shadow:
      inset 0 -18px 30px rgba(0,0,0,0.95),
      inset 0 4px 12px rgba(160, 110, 60, 0.45),
      0 6px 22px rgba(0,0,0,0.85),
      0 0 60px 12px rgba(255, 200, 80, 0.18);
  }

  /* Bulb (visible under shade) */
  .bulb {
    position: absolute;
    top: 78px; left: 50%;
    transform: translateX(-50%);
    width: 46px; height: 42px;
    border-radius: 50% 50% 42% 42% / 55% 55% 45% 45%;
    background: radial-gradient(circle at 50% 35%, #2a2a2a 0%, #0a0a0a 75%);
    transition: all 0.9s cubic-bezier(0.2, 0.9, 0.3, 1);
    z-index: 5;
    box-shadow: inset 0 -6px 12px rgba(0,0,0,0.9);
  }
  .scene.lit .bulb {
    background: radial-gradient(circle at 50% 35%, #ffffff 0%, #fff8e1 25%, #ffd54f 55%, #ff9800 100%);
    box-shadow:
      0 0 22px 8px rgba(255, 235, 160, 0.95),
      0 0 55px 22px rgba(255, 200, 80, 0.6),
      0 0 110px 45px rgba(255, 170, 40, 0.3),
      inset 0 0 10px rgba(255, 255, 220, 0.9);
  }

  /* ========= LIGHT BEAM ========= */
  .beam {
    position: absolute;
    top: 220px; left: 50%;
    transform: translateX(-50%);
    width: 900px; height: 780px;
    background: radial-gradient(
      ellipse at 50% 0%,
      rgba(255, 235, 160, 0.55) 0%,
      rgba(255, 220, 130, 0.30) 18%,
      rgba(255, 200, 80, 0.14) 40%,
      rgba(255, 180, 60, 0.04) 65%,
      transparent 82%
    );
    opacity: 0;
    transition: opacity 1.2s ease;
    pointer-events: none;
    filter: blur(8px);
    z-index: 1;
  }
  .scene.lit .beam { opacity: 1; }

  /* Dust particles */
  .dust {
    position: absolute;
    top: 240px; left: 50%;
    transform: translateX(-50%);
    width: 700px; height: 700px;
    pointer-events: none;
    opacity: 0;
    transition: opacity 1.5s ease 0.3s;
    z-index: 2;
  }
  .scene.lit .dust { opacity: 1; }
  .dust span {
    position: absolute;
    width: 2px; height: 2px;
    background: radial-gradient(circle, #fff8e1 0%, #ffd54f 60%, transparent 100%);
    border-radius: 50%;
    box-shadow: 0 0 4px #ffd54f;
    animation: drift linear infinite;
  }
  @keyframes drift {
    0%   { transform: translateY(0) translateX(0); opacity: 0; }
    20%  { opacity: 0.9; }
    80%  { opacity: 0.9; }
    100% { transform: translateY(220px) translateX(40px); opacity: 0; }
  }

  /* ========= PULL CORD ========= */
  .cord {
    position: absolute;
    top: 130px; left: 50%;
    transform: translateX(48px);
    width: 2px;
    height: 220px;
    cursor: pointer;
    z-index: 10;
    transition: transform 0.4s ease;
    transform-origin: top center;
  }
  .cord::before {
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 100%; height: 100%;
    background: linear-gradient(180deg, #4a4a4a 0%, #b8b8b8 40%, #2a2a2a 100%);
    box-shadow: 1px 0 2px rgba(0,0,0,0.6);
  }
  .cord .bead {
    position: absolute;
    bottom: -14px; left: 50%;
    transform: translateX(-50%);
    width: 18px; height: 22px;
    border-radius: 50% 50% 45% 45% / 60% 60% 40% 40%;
    background:
      radial-gradient(circle at 35% 30%, #d4a35e 0%, #8a5a2a 40%, #3a2210 100%);
    box-shadow:
      0 3px 8px rgba(0,0,0,0.85),
      0 0 14px 2px rgba(255, 190, 90, 0.35),
      inset 0 -2px 4px rgba(0,0,0,0.5);
    transition: all 0.3s ease;
  }
  .cord .bead::after {
    content: '';
    position: absolute;
    top: 5px; left: 5px;
    width: 5px; height: 5px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(255,240,200,0.9) 0%, transparent 70%);
  }
  .cord:hover .bead {
    box-shadow:
      0 3px 8px rgba(0,0,0,0.85),
      0 0 24px 6px rgba(255, 200, 100, 0.9),
      inset 0 -2px 4px rgba(0,0,0,0.5);
    transform: translateX(-50%) scale(1.15);
  }
  .cord.pulled {
    animation: pullCord 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);
  }
  @keyframes pullCord {
    0%   { transform: translateX(48px) scaleY(1); }
    40%  { transform: translateX(48px) scaleY(1.55); }
    70%  { transform: translateX(48px) scaleY(0.92); }
    100% { transform: translateX(48px) scaleY(1); }
  }

  /* ========= LOGIN FORM (inside beam) ========= */
  /* Authentication is rendered by native Streamlit widgets. */
  .form-wrap { display:none !important; }
  .form-wrap {
    position: absolute;
    top: 340px; left: 50%;
    transform: translateX(-50%);
    width: 380px;
    opacity: 0;
    pointer-events: none;
    transition: opacity 0.9s ease 0.4s, transform 0.9s ease 0.4s;
    z-index: 8;
    transform: translateX(-50%) translateY(-15px);
  }
  .scene.lit .form-wrap {
    opacity: 1;
    pointer-events: auto;
    transform: translateX(-50%) translateY(0);
  }
  .glass {
    background: rgba(15, 12, 8, 0.65);
    border: 1px solid rgba(255, 200, 100, 0.28);
    border-radius: 18px;
    padding: 26px 28px 22px;
    backdrop-filter: blur(14px);
    -webkit-backdrop-filter: blur(14px);
    box-shadow:
      0 12px 40px rgba(0,0,0,0.6),
      inset 0 1px 0 rgba(255, 220, 150, 0.15),
      0 0 60px rgba(255, 190, 80, 0.08);
  }
  .brand {
    text-align: center;
    margin-bottom: 16px;
  }
  .brand h1 {
    color: #ffd54f;
    font-size: 20px;
    font-weight: 800;
    letter-spacing: 3px;
    text-shadow: 0 0 18px rgba(255, 200, 80, 0.5);
    margin-bottom: 2px;
  }
  .brand p {
    color: #a8845a;
    font-size: 10px;
    letter-spacing: 4px;
    font-weight: 600;
  }

  .tabs {
    display: flex;
    background: rgba(0,0,0,0.5);
    border-radius: 10px;
    padding: 3px;
    margin-bottom: 16px;
    border: 1px solid rgba(255, 200, 100, 0.12);
  }
  .tab {
    flex: 1;
    text-align: center;
    padding: 8px 0;
    color: #8a7a5a;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 1px;
    border-radius: 8px;
    cursor: pointer;
    transition: all 0.3s ease;
    border: none;
    background: transparent;
    font-family: inherit;
  }
  .tab.active {
    background: linear-gradient(135deg, rgba(255,200,80,0.18) 0%, rgba(255,150,40,0.10) 100%);
    color: #ffd54f;
    box-shadow: 0 2px 8px rgba(255,180,50,0.2), inset 0 0 0 1px rgba(255,200,100,0.25);
  }

  .f-field {
    margin-bottom: 12px;
    position: relative;
  }
  .f-field label {
    display: block;
    color: #a8845a;
    font-size: 10px;
    letter-spacing: 2px;
    font-weight: 700;
    margin-bottom: 5px;
    padding-left: 4px;
  }
  .f-field input {
    width: 100%;
    padding: 11px 14px;
    background: rgba(0,0,0,0.55);
    border: 1px solid rgba(255, 200, 100, 0.15);
    border-radius: 10px;
    color: #fff8e1;
    font-size: 14px;
    font-family: inherit;
    outline: none;
    transition: all 0.3s ease;
    caret-color: #ffd54f;
  }
  .f-field input::placeholder {
    color: #6a5a3a;
    font-style: italic;
  }
  .f-field input:focus {
    border-color: rgba(255, 200, 100, 0.55);
    background: rgba(0,0,0,0.7);
    box-shadow: 0 0 0 3px rgba(255, 200, 100, 0.1), 0 0 20px rgba(255, 180, 50, 0.15);
  }
  .f-field.hidden { display: none; }

  .hint-msg {
    text-align: center;
    color: #a8845a;
    font-size: 11px;
    margin-top: 14px;
    letter-spacing: 1px;
    line-height: 1.6;
  }
  .hint-msg b { color: #ffd54f; }

  /* Error state shake */
  .glass.shake {
    animation: shake 0.45s ease;
  }
  @keyframes shake {
    0%,100% { transform: translateX(0); }
    20%     { transform: translateX(-8px); }
    40%     { transform: translateX(8px); }
    60%     { transform: translateX(-6px); }
    80%     { transform: translateX(6px); }
  }

  /* ========= HINT AT BOTTOM ========= */
  .bottom-hint {
    position: absolute;
    bottom: 40px; left: 50%;
    transform: translateX(-50%);
    color: #666;
    font-size: 12px;
    letter-spacing: 5px;
    text-transform: uppercase;
    font-weight: 300;
    animation: pulse 2.8s ease-in-out infinite;
    transition: all 0.6s ease;
    z-index: 9;
    text-align: center;
    white-space: nowrap;
  }
  .scene.lit .bottom-hint {
    color: #a8845a;
    animation: none;
    letter-spacing: 3px;
    font-size: 11px;
  }
  @keyframes pulse {
    0%,100% { opacity: 0.35; }
    50%     { opacity: 0.95; }
  }
</style>
</head>
<body>
<div class="scene __INITIAL_CLASS__" id="scene">

  <!-- Ceiling mount -->
  <div class="ceiling"></div>

  <!-- Wire -->
  <div class="wire"></div>

  <!-- Lamp assembly -->
  <div class="lamp">
    <div class="lamp-shade"></div>
    <div class="bulb"></div>
  </div>

  <!-- Light beam -->
  <div class="beam"></div>

  <!-- Dust particles -->
  <div class="dust" id="dust"></div>

  <!-- Pull cord -->
  <div class="cord" id="cord"><div class="bead"></div></div>

  <!-- Login form -->
  <div class="form-wrap">
    <div class="glass" id="glass">
      <div class="brand">
        <h1>AL-BARAKAH</h1>
        <p>ENTERPRISES</p>
      </div>
      <div class="tabs">
        <button class="tab active" data-mode="login" type="button">LOGIN</button>
        <button class="tab" data-mode="signup" type="button">SIGNUP</button>
      </div>
      <div class="f-field">
        <label>USERNAME</label>
        <input type="text" id="fUser" placeholder="enter username" autocomplete="off" spellcheck="false">
      </div>
      <div class="f-field">
        <label>PASSWORD</label>
        <input type="password" id="fPass" placeholder="enter password" autocomplete="off">
      </div>
      <div class="f-field hidden" id="fPass2Wrap">
        <label>CONFIRM PASSWORD</label>
        <input type="password" id="fPass2" placeholder="repeat password" autocomplete="off">
      </div>
      <div class="hint-msg" id="hintMsg">
        Pull the cord <b>again</b> to sign in
      </div>
    </div>
  </div>

  <!-- Bottom hint -->
  <div class="bottom-hint" id="bottomHint">▼ PULL THE CORD ▼</div>

</div>

<script>
(function(){
  var scene = document.getElementById('scene');
  var cord = document.getElementById('cord');
  var glass = document.getElementById('glass');
  var bottomHint = document.getElementById('bottomHint');
  var hintMsg = document.getElementById('hintMsg');
  var fUser = document.getElementById('fUser');
  var fPass = document.getElementById('fPass');
  var fPass2 = document.getElementById('fPass2');
  var fPass2Wrap = document.getElementById('fPass2Wrap');
  var tabs = document.querySelectorAll('.tab');
  var mode = 'login';
  var pulled = false;

  /* ---- Dust particles ---- */
  var dust = document.getElementById('dust');
  for (var i = 0; i < 22; i++) {
    var s = document.createElement('span');
    s.style.left = (10 + Math.random() * 80) + '%';
    s.style.top = (Math.random() * 60) + '%';
    s.style.animationDuration = (5 + Math.random() * 6) + 's';
    s.style.animationDelay = (Math.random() * 5) + 's';
    var sz = 1 + Math.random() * 2;
    s.style.width = sz + 'px';
    s.style.height = sz + 'px';
    dust.appendChild(s);
  }

  /* ---- Tab switching ---- */
  tabs.forEach(function(t){
    t.addEventListener('click', function(){
      tabs.forEach(function(x){ x.classList.remove('active'); });
      t.classList.add('active');
      mode = t.getAttribute('data-mode');
      if (mode === 'signup') {
        fPass2Wrap.classList.remove('hidden');
        hintMsg.innerHTML = 'Pull the cord <b>again</b> to create account';
      } else {
        fPass2Wrap.classList.add('hidden');
        hintMsg.innerHTML = 'Pull the cord <b>again</b> to sign in';
      }
    });
  });

  /* ---- Native value setter for React inputs ---- */
  function setNativeValue(element, value) {
    var proto = Object.getPrototypeOf(element);
    var setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
    setter.call(element, value);
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
  }

  /* ---- Find streamlit hidden input by label marker ---- */
  function findStreamlitInput(marker) {
    try {
      var pd = window.parent.document;
      var labels = pd.querySelectorAll('label');
      for (var i = 0; i < labels.length; i++) {
        if (labels[i].textContent.indexOf(marker) !== -1) {
          var wrap = labels[i].closest('[data-testid="stTextInput"]');
          if (wrap) return wrap.querySelector('input');
        }
      }
    } catch(e) {}
    return null;
  }

  /* ---- Click streamlit hidden button by text ---- */
  function clickStreamlitButton(marker) {
    try {
      var pd = window.parent.document;
      var btns = pd.querySelectorAll('button');
      for (var i = 0; i < btns.length; i++) {
        if ((btns[i].textContent || '').indexOf(marker) !== -1) {
          btns[i].click();
          return true;
        }
      }
    } catch(e) {}
    return false;
  }

  /* ---- Cord click ---- */
  cord.addEventListener('click', function(){
    if (pulled) return;
    pulled = true;
    cord.classList.add('pulled');
    setTimeout(function(){ cord.classList.remove('pulled'); pulled = false; }, 650);

    var isLit = scene.classList.contains('lit');

    if (!isLit) {
      /* Turn ON: click the "TURN_ON" streamlit button */
      setTimeout(function(){
        clickStreamlitButton('__LAMP_TURN_ON__');
      }, 350);
    } else {
      /* Turn OFF: check form values first */
      var user = (fUser.value || '').trim();
      var pass = fPass.value || '';
      var pass2 = fPass2.value || '';

      if (user === '' || pass === '') {
        /* shake */
        glass.classList.add('shake');
        setTimeout(function(){ glass.classList.remove('shake'); }, 500);
        hintMsg.innerHTML = '<span style="color:#ff7b7b;">Please fill all fields</span>';
        setTimeout(function(){
          hintMsg.innerHTML = mode === 'signup'
            ? 'Pull the cord <b>again</b> to create account'
            : 'Pull the cord <b>again</b> to sign in';
        }, 2200);
        return;
      }

      if (mode === 'signup' && pass !== pass2) {
        glass.classList.add('shake');
        setTimeout(function(){ glass.classList.remove('shake'); }, 500);
        hintMsg.innerHTML = '<span style="color:#ff7b7b;">Passwords do not match</span>';
        setTimeout(function(){
          hintMsg.innerHTML = 'Pull the cord <b>again</b> to create account';
        }, 2200);
        return;
      }

      /* Set hidden streamlit inputs and click submit */
      var uIn = findStreamlitInput('__LU__');
      var pIn = findStreamlitInput('__LP__');
      var p2In = findStreamlitInput('__LP2__');
      var mIn = findStreamlitInput('__LM__');
      if (uIn) setNativeValue(uIn, user);
      if (pIn) setNativeValue(pIn, pass);
      if (p2In) setNativeValue(p2In, pass2);
      if (mIn) setNativeValue(mIn, mode);

      setTimeout(function(){
        clickStreamlitButton('__LAMP_SUBMIT__');
      }, 220);
    }
  });

  /* Auto-focus username when lit */
  if (scene.classList.contains('lit')) {
    setTimeout(function(){ try { fUser.focus(); } catch(e){} }, 800);
  }
})();
</script>
</body>
</html>
""".replace("__INITIAL_CLASS__", initial_class)

# ============================================================
# GLOBAL CSS (for the app itself, applies when logged in)
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
# SESSION STATE INIT
# ============================================================
if "light_on" not in st.session_state:
    st.session_state["light_on"] = False
if "auth_error" not in st.session_state:
    st.session_state["auth_error"] = ""

# ============================================================
# RELIABLE AUTH UI — NATIVE STREAMLIT CONTROLS
# ============================================================
if not st.session_state.get("logged_in_user"):
    st.markdown("""
    <style>
        section[data-testid="stSidebar"],
        header[data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        [data-testid="stStatusWidget"],
        #MainMenu, footer, .stAppDeployButton,
        [data-testid="manage-app-button"] { display:none !important; }

        html, body, .stApp {
            background:#11110f !important;
            overflow:hidden !important;
        }
        .block-container {
            padding:0 !important;
            max-width:100% !important;
            margin:0 !important;
        }
        iframe { border:none !important; pointer-events:none !important; position:relative !important; z-index:1 !important; }

        /* Premium authentication card behind the real Streamlit fields */
        .auth-card {
            pointer-events:none !important;
            position:fixed;
            top:350px;
            left:50%;
            transform:translateX(-50%);
            width:400px;
            min-height:245px;
            border:1px solid rgba(255,195,90,.30);
            border-radius:20px;
            background:linear-gradient(145deg,rgba(38,35,30,.96),rgba(20,19,17,.97));
            box-shadow:0 20px 70px rgba(0,0,0,.72),0 0 70px rgba(255,183,65,.13),inset 0 1px 0 rgba(255,225,165,.12);
            backdrop-filter:blur(18px);
            z-index:50;
        }
        .auth-title {
            position:fixed; top:368px; left:50%; transform:translateX(-50%);
            width:360px; text-align:center; z-index:70; pointer-events:none;
            color:#ffd36a; font-weight:800; letter-spacing:4px; font-size:20px;
            text-shadow:0 0 20px rgba(255,198,90,.38);
        }
        .auth-subtitle {
            position:fixed; top:397px; left:50%; transform:translateX(-50%);
            width:360px; text-align:center; z-index:70; pointer-events:none;
            color:#a88b61; font-size:9px; font-weight:700; letter-spacing:5px;
        }

        /* Real Streamlit widgets are positioned over the card. */
        div[data-testid="stRadio"] {
            position:fixed !important;
            top:424px !important;
            left:50% !important;
            transform:translateX(-50%) !important;
            width:340px !important;
            z-index:100 !important;
            background:rgba(7,7,6,.68) !important;
            border:1px solid rgba(255,196,88,.16) !important;
            border-radius:11px !important;
            padding:4px 10px !important;
        }
        div[data-testid="stRadio"] label,
        div[data-testid="stRadio"] p,
        div[data-testid="stRadio"] span { color:#d5b777 !important; }
        div[data-testid="stRadio"] label { font-size:12px !important; font-weight:700 !important; }

        div[data-testid="stTextInput"]:has(input[aria-label="USERNAME"]) {
            position:fixed !important; top:478px !important; left:50% !important;
            transform:translateX(-50%) !important; width:340px !important; z-index:100 !important;
        }
        div[data-testid="stTextInput"]:has(input[aria-label="PASSWORD"]) {
            position:fixed !important; top:550px !important; left:50% !important;
            transform:translateX(-50%) !important; width:340px !important; z-index:100 !important;
        }
        div[data-testid="stTextInput"]:has(input[aria-label="CONFIRM PASSWORD"]) {
            position:fixed !important; top:622px !important; left:50% !important;
            transform:translateX(-50%) !important; width:340px !important; z-index:100 !important;
        }
        div[data-testid="stTextInput"] label,
        div[data-testid="stTextInput"] label p { color:#a98c61 !important; font-size:9px !important; font-weight:800 !important; letter-spacing:2px !important; }
        div[data-testid="stTextInput"] input {
            background:rgba(7,7,7,.78) !important;
            color:#fff8e8 !important;
            -webkit-text-fill-color:#fff8e8 !important;
            border:1px solid rgba(255,198,90,.20) !important;
            border-radius:10px !important;
            height:43px !important;
        }
        div[data-testid="stTextInput"] input:focus {
            border-color:rgba(255,203,105,.65) !important;
            box-shadow:0 0 0 2px rgba(255,190,70,.10),0 0 20px rgba(255,185,60,.12) !important;
        }

        /* Reliable invisible Streamlit click target over the pull bead. */
        div[data-testid="stButton"] {
            position:fixed !important;
            top:350px !important;
            left:calc(50% + 5px) !important;
            width:95px !important;
            height:105px !important;
            z-index:2147483647 !important;
            margin:0 !important;
            padding:0 !important;
        }
        div[data-testid="stButton"] button {
            width:95px !important; height:105px !important;
            min-height:105px !important;
            opacity:0 !important; cursor:pointer !important;
            padding:0 !important; border:0 !important;
            background:transparent !important;
        }

        .auth-error {
            position:fixed; top:665px; left:50%; transform:translateX(-50%);
            width:380px; text-align:center; z-index:300;
            color:#ff9a8f; font-size:12px; font-weight:700;
            background:rgba(90,20,15,.78); border:1px solid rgba(255,100,80,.22);
            border-radius:9px; padding:8px 12px;
        }
        div[data-testid="stRadio"],
        div[data-testid="stTextInput"] {
            position:relative !important;
            z-index:1000 !important;
        }

        .auth-hint {
            position:fixed; top:700px; left:50%; transform:translateX(-50%);
            width:420px; text-align:center; z-index:60; pointer-events:none;
            color:#a88b61; font-size:10px; letter-spacing:2px;
        }

        @media(max-width:600px){
            .auth-card{width:calc(100vw - 32px);}
            div[data-testid="stRadio"],
            div[data-testid="stTextInput"]:has(input[aria-label="USERNAME"]),
            div[data-testid="stTextInput"]:has(input[aria-label="PASSWORD"]),
            div[data-testid="stTextInput"]:has(input[aria-label="CONFIRM PASSWORD"]){width:calc(100vw - 70px) !important;}
        }
    </style>
    """, unsafe_allow_html=True)

    # Visual lamp only. Authentication is handled by real Streamlit widgets,
    # so clicking the cord can never get disconnected from Python state.
    components.html(build_lamp_html(st.session_state["light_on"]), height=900, scrolling=False)

    # Decorative card/title appear only after the lamp is lit.
    if st.session_state["light_on"]:
        st.markdown('<div class="auth-card"></div><div class="auth-title">AL-BARAKAH</div><div class="auth-subtitle">ENTERPRISES</div>', unsafe_allow_html=True)

        mode = st.radio("MODE", ["LOGIN", "SIGNUP"], horizontal=True, key="auth_mode", label_visibility="collapsed")
        st.text_input("USERNAME", key="auth_username", placeholder="Enter username", label_visibility="visible")
        st.text_input("PASSWORD", key="auth_password", type="password", placeholder="Enter password", label_visibility="visible")
        if mode == "SIGNUP":
            st.text_input("CONFIRM PASSWORD", key="auth_password2", type="password", placeholder="Repeat password", label_visibility="visible")

        st.markdown('<div class="auth-hint">PULL THE CORD TO CONTINUE</div>', unsafe_allow_html=True)

    # A real Streamlit button is placed over the visible pull-cord bead.
    st.markdown('<div class="cord-hit-area">', unsafe_allow_html=True)
    pull_cord = st.button("PULL", key="pull_cord", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if pull_cord:
        if not st.session_state["light_on"]:
            st.session_state["light_on"] = True
            st.session_state["auth_error"] = ""
            st.rerun()
        else:
            users = load_users()
            uname_raw = st.session_state.get("auth_username", "").strip()
            uname = sanitize_username(uname_raw)
            pass_v = st.session_state.get("auth_password", "")
            pass2_v = st.session_state.get("auth_password2", "")

            if not uname or not pass_v:
                st.session_state["auth_error"] = "Please fill all required fields."
            elif mode == "LOGIN":
                if uname not in users:
                    st.session_state["auth_error"] = "Username not found. Please signup."
                elif users[uname].get("password_hash") != hash_password(pass_v):
                    st.session_state["auth_error"] = "Incorrect password."
                else:
                    st.session_state["logged_in_user"] = uname
                    st.session_state["display_name"] = users[uname].get("display_name", uname)
                    st.session_state["page"] = "📊 Dashboard"
                    st.session_state["light_on"] = False
                    st.rerun()
            else:
                if len(uname) < 3:
                    st.session_state["auth_error"] = "Username must be at least 3 characters."
                elif len(uname) > 20:
                    st.session_state["auth_error"] = "Username must be 20 characters or less."
                elif len(pass_v) < 4:
                    st.session_state["auth_error"] = "Password must be at least 4 characters."
                elif pass_v != pass2_v:
                    st.session_state["auth_error"] = "Passwords do not match."
                elif uname in users:
                    st.session_state["auth_error"] = f"Username '{uname}' already taken."
                else:
                    users[uname] = {
                        "username": uname,
                        "display_name": uname_raw,
                        "password_hash": hash_password(pass_v),
                        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    }
                    save_users(users)
                    with open(user_data_file(uname), "w", encoding="utf-8") as f:
                        json.dump(default_blank_db(), f, indent=4, ensure_ascii=False)
                    st.session_state["logged_in_user"] = uname
                    st.session_state["display_name"] = uname_raw
                    st.session_state["page"] = "📊 Dashboard"
                    st.session_state["light_on"] = False
                    st.rerun()

            if st.session_state.get("auth_error"):
                st.markdown(f'<div class="auth-error">⚠️ {st.session_state["auth_error"]}</div>', unsafe_allow_html=True)

    st.stop()

# ============================================================
# LOAD USER DATA (after login)
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
# PAGE: BILLS LIST
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
                "shop": b.get("Shop",""), "date": b.get("Date",""),
                "booker": b.get("Order Booker",""), "salesman": b.get("Salesman",""),
                "items": [], "orig_indices": [], "bill_no": b.get("Bill No",""),
            }
        groups[key]["items"].append({
            "Code": b.get("Code"), "Product": b.get("Product"),
            "Boxes": b.get("Boxes"), "TP/Box": b.get("TP/Box"),
            "Discount %": b.get("Discount %"), "Gross": b.get("Gross"), "Net": b.get("Net"),
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
        shop = g["shop"] or "-"; date_str = g["date"] or "-"
        booker = g["booker"] or "-"; salesman = g["salesman"] or "-"
        bill_no = g["bill_no"]; items = g["items"]
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
            if st.button(eye_icon, key=f"eye_bill_{wkey}", use_container_width=True, help="Poora bill dekho"):
                if is_viewing: st.session_state["view_bill_key"] = None
                else: st.session_state["view_bill_key"] = wkey
                st.rerun()
        with c3:
            if st.button("⬇️ Excel", key=f"dl_bill_{wkey}", use_container_width=True, type="primary"):
                export_single_group_bill(shop, date_str, booker, salesman, items, bill_no)
                st.rerun()
        with c4:
            if st.button("🗑 Delete", key=f"del_bill_{wkey}", use_container_width=True):
                st.session_state["confirm_delete_group"] = g["orig_indices"]
                st.session_state["_confirm_group_label"] = f"{shop} | {date_str} | {booker}"

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
# PAGE: LOAD FORM
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
            if st.button(eye_icon, key=f"eye_lf_{wkey}", use_container_width=True, help="Poora load form dekho"):
                if is_viewing: st.session_state["view_lf_key"] = None
                else: st.session_state["view_lf_key"] = wkey
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
        b_net = float(bill.get("Net", 0)); b_gross = float(bill.get("Gross", 0))
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
