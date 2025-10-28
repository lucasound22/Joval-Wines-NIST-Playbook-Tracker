import os
import io
import re
import json
import base64
import hashlib
import secrets
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

import streamlit as st
import mammoth
from bs4 import BeautifulSoup
import pandas as pd
from fpdf import FPDF

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

import logging

# === CONFIGURATION ===
ref_pattern = re.compile(r'^\d+(\.\d+)*\b')

logging.basicConfig(
    filename='audit.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

PLAYBOOKS_DIR = "playbooks"
USERS_FILE = "users.json"
Path(PLAYBOOKS_DIR).mkdir(exist_ok=True)
Path(USERS_FILE).touch(exist_ok=True)

# === DYNAMIC SIDEBAR STATE ===
if 'authenticated' in st.session_state and st.session_state.authenticated:
    sidebar_state = "expanded"
else:
    sidebar_state = "collapsed"

st.set_page_config(
    page_title="Joval Wines NIST Playbook Tracker",
    page_icon="wine",
    layout="wide",
    initial_sidebar_state=sidebar_state
)

# === HIDE STREAMLIT BRANDING ===
st.markdown("""
<style>
/* Remove footer, toolbar, share button, deploy button */
.stApp > footer {display: none !important;}
.stApp [data-testid="stToolbar"] {display: none !important;}
.stApp [data-testid="collapsedControl"] {display: none !important;}
.stDeployButton {display: none !important;}
.stApp [data-testid="stDecoration"] {display: none !important;}
</style>
""", unsafe_allow_html=True)

# === FULL CSS + JAVASCRIPT (ROBUST EXPAND & HIGHLIGHT) ===
st.markdown("""
<script src="https://cdn.tailwindcss.com"></script>
<script>
function scrollToSection(id) {
    const el = document.getElementById(id);
    if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        el.style.backgroundColor = '#fff3cd';
        setTimeout(() => { el.style.backgroundColor = ''; }, 2000);
    }
}
function highlightText(id, term) {
    const el = document.getElementById(id);
    if (!el || !term) return;
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null, false);
    const nodes = [];
    let node;
    while (node = walker.nextNode()) nodes.push(node);
    const regex = new RegExp('(' + term.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&') + ')', 'gi');
    nodes.forEach(n => {
        if (regex.test(n.nodeValue)) {
            const parts = n.nodeValue.split(regex);
            const fragment = document.createDocumentFragment();
            parts.forEach((part, i) => {
                if (i % 2 === 1) {
                    const span = document.createElement('span');
                    span.style.backgroundColor = '#fff3cd';
                    span.style.fontWeight = 'bold';
                    span.textContent = part;
                    fragment.appendChild(span);
                } else {
                    fragment.appendChild(document.createTextNode(part));
                }
            });
            n.parentNode.replaceChild(fragment, n);
        }
    });
}
function expandSection(id) {
    const target = document.getElementById(id);
    if (!target) return;
    let expander = target.closest('[data-testid="stExpander"]');
    if (!expander) {
        let p = target.parentElement;
        while (p && p.getAttribute('data-testid') !== 'stExpander') p = p.parentElement;
        expander = p;
    }
    if (expander) {
        const details = expander.querySelector('details');
        if (details && !details.open) {
            details.open = true;
        }
    }
}
</script>
<style>
:root{ --bg:#fff; --text:#000; --muted:#777; --joval-red:#800020; --section-bg: rgba(0,0,0,0.02); }
html, body, .stApp { background:var(--bg)!important; color:var(--text)!important; font-family: 'Helvetica', 'Arial', sans-serif !important; }
.sticky-header{position:sticky;top:0;z-index:9999;display:flex;align-items:center;justify-content:space-between;padding:10px 18px;background:linear-gradient(180deg, rgba(255,255,255,0.96), rgba(245,245,245,0.86));border-bottom:1px solid rgba(0,0,0,0.04); width:100%; box-sizing:border-box;}
.logo-left{flex-shrink:0; height:80px; margin-right:auto;}
.app-title{flex:1; text-align:center; font-family:'Helvetica', sans-serif; font-size:3rem; color:var(--text); font-weight:700; margin:0; text-shadow: 2px 2px 4px rgba(0,0,0,0.3);}
.nist-logo{flex-shrink:0; font-family:'Helvetica', sans-serif; font-size:2rem; color:var(--joval-red); letter-spacing:0.08rem; text-shadow:0 0 12px rgba(128,0,32,0.08); font-weight:700; margin-left:auto;}
.toc{ position:fixed; left:12px; top:84px; bottom:92px; width:260px; background:rgba(255,255,255,0.95); padding:10px; border-radius:8px; overflow:auto; border:1px solid rgba(0,0,0,0.03); z-index:900;}
.content-wrap{margin-left:284px; margin-right:24px; padding-top:0px; padding-bottom:100px; margin-top: 0px;}
.section-box{background:var(--section-bg); padding:12px; border-radius:8px; margin-bottom:12px; border:1px solid rgba(0,0,0,0.02);}
.scaled-img{max-width:90%; height:auto; border-radius:8px; box-shadow:0 6px 18px rgba(0,0,0,0.6); margin:12px 0; display:block;}
.playbook-table{border-collapse:collapse; width:100%; margin-top:8px; margin-bottom:8px;}
.playbook-table th, .playbook-table td{border:1px solid rgba(0,0,0,0.05); padding:8px; color:var(--text); text-align:left; vertical-align:top;}
.playbook-table th{background: rgba(0,0,0,0.02); color:#333; font-weight:700;}
.row-preview{white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:600px; display:inline-block;}
.stCheckbox>label, .stCheckbox>label>div { color: var(--muted) !important; }
.stTextInput > label { color: var(--text) !important; }
.stSelectbox > label { color: var(--text) !important; font-weight: bold; font-size: 1.1rem; }
.stSelectbox { margin-bottom: 0 !important; }
.bottom-toolbar{position:fixed; left:0; right:0; bottom:0; z-index:999; background:linear-gradient(180deg, rgba(255,255,255,0.95), rgba(245,245,245,0.9)); border-top:1px solid rgba(0,0,0,0.03); padding:10px 18px; display:flex; justify-content:space-between; align-items:center;}
.progress-wrap{width:360px; height:10px; background:rgba(0,0,0,0.03); border-radius:999px; overflow:hidden;}
.progress-fill{height:100%; width:0%; background:linear-gradient(90deg, var(--joval-red), #500010); box-shadow:0 0 12px rgba(128,0,32,0.2); transition: width 0.6s ease;}
.stButton>button, .stButton>button:hover, .stDownloadButton>button, .stDownloadButton>button:hover{ background:#000 !important; color:#fff !important; border:1px solid rgba(0,0,0,0.12); font-weight:700; }
.stButton button, .stDownloadButton button { background:#000 !important; color:#fff !important; border:1px solid rgba(0,0,0,0.12); font-weight:700; }
.stSidebar .stButton > button { color: #fff !important; background: var(--joval-red) !important; font-size: 1.1rem; font-weight: bold; }
.comments-title{color:var(--text); font-weight:700; margin-top:12px; margin-bottom:6px;}
a { color: var(--joval-red); cursor: pointer; }
.toc a { display:block; padding:6px 4px; color:var(--text); text-decoration:none; }
.toc a:hover { background: rgba(0,0,0,0.02); border-radius:4px; }
.refresh-btn { background:#000 !important; color:#fff !important; border:1px solid rgba(0,0,0,0.12); font-weight:700; padding: 0.5rem 1rem; border-radius: 0.25rem; cursor:pointer; margin-left:10px; }
.copyright { color: var(--text) !important; font-size: 1.1rem; margin-left: 20px; }
.sidebar-header { color: var(--text) !important; font-weight: bold; font-size: 1.2rem; }
.sidebar-subheader { color: var(--text) !important; font-weight: bold; font-size: 1.1rem; }
.sidebar-footer { color: var(--text) !important; font-size: 1.1rem; }
.section-title { font-size: 1.8rem !important; font-weight: bold !important; color: var(--text) !important; margin-top: 0 !important; margin-bottom: 4px !important; }
.instructional-text { 
  color: #d9534f !important; 
  font-size: 1.5rem !important; 
  font-weight: bold !important; 
  border: 2px solid #d9534f; 
  padding: 15px; 
  border-radius: 8px; 
  background: rgba(217,83,79,0.1); 
  margin: 20px 0 !important; 
  text-align: center; 
  line-height: 1.4 !important; 
}
.content-text { font-size: 1.1rem !important; }
[data-testid="stExpander"] > div:first-child {
  background:#f0f0f0 !important;
  color:#000 !important;
  padding: 12px 20px !important;
  border-radius: 8px !important;
  font-size: 1.3rem !important;
  font-weight: bold !important;
  border:1px solid rgba(0,0,0,0.12) !important;
  margin-bottom: 10px !important;
  cursor: pointer !important;
  transition: all 0.2s ease !important;
}
[data-testid="stExpander"] > div:first-child:hover {
  background:#e0e0f0 !important;
}
[data-testid="stExpander"] label { color: var(--text) !important; font-size: 1.5rem !important; font-weight: bold !important; }
[data-testid="stExpander"] [data-testid="stArrowToggle"] { color: var(--text) !important; }
.theme-label, .compliance-label { color: #000 !important; font-weight: bold; }
.stSidebar label { color: #000 !important; }
.light-theme .stSelectbox > label { color: #000 !important; }
.light-theme .stTextInput > label { color: #000 !important; }
.light-theme .stButton>button { background: #000 !important; color: #fff !important; }
.light-theme .stDownloadButton>button { background: #000 !important; color: #fff !important; }
.light-theme .sidebar-header, .light-theme .sidebar-subheader, .light-theme .sidebar-footer { color: #000 !important; }
.light-theme .comments-title { color: #000 !important; }
.light-theme .section-title { color: #000 !important; }
.light-theme .app-title { color: #000 !important; }
.light-theme .nist-logo { color: var(--joval-red); }
.light-theme [data-testid="stExpander"] > div:first-child { background: #f0f0f0 !important; color: #000 !important; border: 1px solid rgba(0,0,0,0.12) !important; }
.light-theme [data-testid="stExpander"] label { color: #333 !important; }
.light-theme .bottom-toolbar { background: linear-gradient(180deg, rgba(255,255,255,0.95), rgba(250,250,250,0.9)); border-top:1px solid rgba(0,0,0,0.03); }
.light-theme .progress-wrap { background: rgba(0,0,0,0.03); }
.light-theme .progress-fill { background: linear-gradient(90deg, var(--joval-red), #500010); box-shadow:0 0 12px rgba(128,0,32,0.2); }
.light-theme .bottom-toolbar div { color: #000 !important; }
.progress-wrap { position: relative; }
.progress-fill { position: absolute; top: 0; left: 0; }
.login-container { background: var(--bg); color: var(--text); padding: 2rem; text-align: center; max-width: 400px; margin: 0 auto; }
.login-title { color: var(--text) !important; font-size: 1.5rem; font-weight: bold; margin-bottom: 1rem; }
.login-subtitle { color: var(--joval-red) !important; font-size: 1.2rem; margin-bottom: 2rem; }
.stHelp { color: var(--text) !important; background: rgba(255,255,255,0.8) !important; border: 1px solid var(--joval-red) !important; }
@media (max-width: 768px) {
  .stApp { max-width: 100%; }
  .content-wrap { margin-left: 0 !important; margin-right: 0 !important; padding: 0 10px; }
  .toc { display: none !important; }
  .sticky-header { flex-direction: column; padding: 10px; }
  .app-title { font-size: 2rem; margin: 10px 0; }
  .section-title { font-size: 1.4rem !important; }
  .playbook-table { font-size: 0.8rem; }
  .stSelectbox, .stTextInput { width: 100%; }
  .bottom-toolbar { flex-direction: column; gap: 10px; padding: 10px; }
  .progress-wrap { width: 100%; }
  [data-testid="stExpander"] label { font-size: 1.2rem !important; }
  .instructional-text { font-size: 1.2rem !important; padding: 10px; }
}
.playbook-select-label { font-size: 2.5rem !important; font-weight: bold !important; color: var(--text) !important; }
.nist-incident-section { color: #d9534f !important; }
.security-icon { font-size: 1.2rem; opacity: 0.7; margin-right: 0.5rem; }
/* Search result styling */
.search-result-btn { background: #800020 !important; color: white !important; font-size: 0.9rem; padding: 6px 12px; border-radius: 6px; text-align: left; width: 100%; margin: 4px 0; }
.search-result-btn:hover { background: #a00030 !important; }
.search-result-snippet { font-size: 0.8rem; color: #555; margin-top: 2px; }
</style>
""", unsafe_allow_html=True)

# === USER MANAGEMENT (SECURE) ===
def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                content = f.read().strip()
                if content:
                    return {k.lower(): v for k, v in json.loads(content).items()}
        except (json.JSONDecodeError, ValueError):
            pass

    admin_email = "admin@joval.com"
    admin_hash = st.secrets.get("ADMIN_PASSWORD_HASH")
    if not admin_hash:
        st.error("ADMIN_PASSWORD_HASH not set in secrets.toml")
        st.stop()
    
    default_admin = {admin_email: {"role": "admin", "hash": admin_hash}}
    save_users(default_admin)
    return default_admin

def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump({k.lower(): v for k, v in users.items()}, f, indent=2)

def get_user_role(email):
    users = load_users()
    return users.get(email.lower(), {}).get("role", "user")

def create_user(email, role, password):
    users = load_users()
    email = email.lower()
    if email in users:
        return False, "User already exists."
    hash_pass = hashlib.sha256(password.encode()).hexdigest()
    users[email] = {"role": role, "hash": hash_pass}
    save_users(users)
    logging.info(f"User created: {email}, Role: {role}")
    return True, "User created successfully."

def reset_user_password(email, password):
    users = load_users()
    email = email.lower()
    if email not in users:
        return False, "User not found."
    hash_pass = hashlib.sha256(password.encode()).hexdigest()
    users[email]["hash"] = hash_pass
    save_users(users)
    logging.info(f"Password reset: {email}")
    return True, "Password reset successfully."

def delete_user(email):
    users = load_users()
    email = email.lower()
    if email in users:
        del users[email]
        save_users(users)
        logging.info(f"User deleted: {email}")
        return True, "User deleted successfully."
    return False, "User not found."

def authenticate():
    if 'login_attempts' not in st.session_state:
        st.session_state.login_attempts = 0
    if 'last_attempt' not in st.session_state:
        st.session_state.last_attempt = None
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'user' not in st.session_state:
        st.session_state.user = None

    if not st.session_state.authenticated:
        st.title("Joval Wines NIST Playbook Tracker")
        st.markdown("### Login Required")
        st.markdown(get_logo(), unsafe_allow_html=True)
        
        username = st.text_input("Username", key="username")
        password = st.text_input("Password", type="password", key="password")
        
        now = datetime.now()
        if st.button("Login"):
            if st.session_state.login_attempts >= 5:
                if st.session_state.last_attempt and (now - st.session_state.last_attempt).seconds < 300:
                    st.error("Too many failed attempts. Try again in 5 minutes.")
                    st.stop()
                else:
                    st.session_state.login_attempts = 0
            
            users = load_users()
            email = username if "@" in username else username + "@joval.com"
            email = email.lower()
            if email in users:
                if hashlib.sha256(password.encode()).hexdigest() == users[email]["hash"]:
                    st.session_state.authenticated = True
                    display_name = username.split("@")[0].title() if "@" in username else username.title()
                    st.session_state.user = {"email": email, "name": display_name, "role": users[email]["role"]}
                    st.session_state.login_attempts = 0
                    st.success("Logged in successfully!")
                    st.rerun()
                else:
                    st.session_state.login_attempts += 1
                    st.session_state.last_attempt = now
                    st.error("Invalid credentials.")
            else:
                st.session_state.login_attempts += 1
                st.session_state.last_attempt = now
                st.error("Invalid credentials.")
        st.stop()

    if st.sidebar.button("Logout"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    return st.session_state.user

# === ADMIN DASHBOARD ===
def admin_dashboard(user):
    if get_user_role(user["email"]) != "admin":
        st.error("Access denied. Admin only.")
        return

    st.title("Admin Dashboard")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Create User", "Reset Password", "List Users", "Delete User", "Upload Logo/Playbook"])

    with tab1:
        st.subheader("Create New User")
        email_input = st.text_input("User Email")
        email = email_input if "@" in email_input else email_input + "@joval.com"
        role = st.selectbox("Role", ["user", "admin"])
        generate_pass = st.checkbox("Generate Random Password", value=True)
        if generate_pass:
            password = secrets.token_urlsafe(16)
            st.markdown(f'<p style="font-size:2rem;">Generated Password: <strong>{password}</strong></p> <p>(Share securely; shown only once)</p>', unsafe_allow_html=True)
        else:
            password = st.text_input("Set Password", type="password")
        if st.button("Create User"):
            if email and password:
                success, msg = create_user(email, role, password)
                if success:
                    st.success(msg)
                else:
                    st.error(msg)
            else:
                st.error("Fill all fields.")

    with tab2:
        st.subheader("Reset User Password")
        email_input = st.text_input("User Email to Reset")
        email = email_input if "@" in email_input else email_input + "@joval.com"
        generate_pass = st.checkbox("Generate Random Password", value=True, key="reset_gen")
        if generate_pass:
            password = secrets.token_urlsafe(16)
            show_pass = f'<p style="font-size:2rem;">Generated Password: <strong>{password}</strong></p> <p>(Share securely; shown only once)</p>'
        else:
            password = st.text_input("Set New Password", type="password", key="reset_custom")
            show_pass = "Password set."
        if st.button("Reset Password"):
            if email and password:
                success, msg = reset_user_password(email, password)
                if success:
                    st.markdown(show_pass, unsafe_allow_html=True) if generate_pass else st.success(msg)
                else:
                    st.error(msg)
            else:
                st.error("Enter email and password.")

    with tab3:
        st.subheader("List Users")
        users = load_users()
        user_list = [{"Email": k, "Role": v["role"]} for k, v in users.items()]
        st.table(pd.DataFrame(user_list))

    with tab4:
        st.subheader("Delete User")
        email_input = st.text_input("User Email to Delete")
        email = email_input if "@" in email_input else email_input + "@joval.com"
        if st.button("Delete User"):
            if email:
                success, msg = delete_user(email)
                if success:
                    st.success(msg)
                else:
                    st.error(msg)
            else:
                st.error("Enter email.")

    with tab5:
        st.subheader("Upload Custom Logo")
        uploaded_logo = st.file_uploader("Upload Logo", type=["png", "jpg", "jpeg"])
        if uploaded_logo:
            st.session_state.logo_b64 = base64.b64encode(uploaded_logo.read()).decode()
            st.success("Logo uploaded!")
            st.rerun()
        st.subheader("Upload New Playbook")
        uploaded_playbook = st.file_uploader("Upload Word Doc", type=["docx"])
        if uploaded_playbook:
            file_path = os.path.join(PLAYBOOKS_DIR, uploaded_playbook.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_playbook.getbuffer())
            st.success(f"Playbook uploaded!")

    if st.button("Back to Main App"):
        st.session_state.admin_page = False
        st.rerun()

# === UTILITIES ===
def stable_key(playbook_name: str, title: str, level: int) -> str:
    base = f"{playbook_name}||{level}||{title}"
    return "sec_" + hashlib.md5(base.encode("utf-8")).hexdigest()

@st.cache_data
def progress_filepath(playbook_name: str) -> str:
    base = os.path.splitext(playbook_name)[0]
    return os.path.join(PLAYBOOKS_DIR, f"{base}_progress.json")

@st.cache_data
def load_progress(playbook_name: str):
    path = progress_filepath(playbook_name)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
                return data.get("completed", {}), data.get("comments", {})
        except Exception:
            return {}, {}
    return {}, {}

def save_progress(playbook_name: str, completed_map: dict, comments_map: dict) -> str:
    rec = {
        "playbook": playbook_name,
        "timestamp": datetime.now().isoformat(),
        "version": "1.0",
        "completed": completed_map,
        "comments": comments_map,
    }
    path = progress_filepath(playbook_name)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(rec, fh, indent=2)
    return path

def safe_image_display(src: str) -> bool:
    if not src:
        return False
    try:
        st.markdown(f"<img class='scaled-img' src='{src}'/>", unsafe_allow_html=True)
        return True
    except Exception:
        try:
            st.image(src)
            return True
        except Exception:
            return False

def clean_for_pdf(text: str) -> str:
    if not text:
        return ""
    replacements = {
        '\u2014': '-', '\u2013': '-', '\u2022': '*', '\u2026': '...',
        '\u201c': '"', '\u201d': '"', '\u2018': "'", '\u2019': "'",
        '\u00A0': ' ', '\u2191': ' (up) '
    }
    for unicode_char, replacement in replacements.items():
        text = text.replace(unicode_char, replacement)
    text = ''.join(char for char in text if ord(char) < 256 or char in replacements)
    return text

def calculate_badges(pct: int) -> List[str]:
    if pct >= 100:
        return ["Gold Star"]
    elif pct >= 80:
        return ["Silver Shield"]
    elif pct >= 50:
        return ["Bronze Medal"]
    elif pct >= 25:
        return ["Progress Starter"]
    elif pct > 0:
        return ["Just Started"]
    else:
        return ["Ready to Begin"]

def save_feedback(rating: int, comments: str):
    feedback_data = {"rating": rating, "comments": comments, "timestamp": datetime.now().isoformat()}
    with open("feedback.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(feedback_data) + "\n")

def show_feedback():
    with st.expander("Provide Feedback"):
        with st.form("feedback_form"):
            rating = st.slider("Rate this session (1-5)", 1, 5, 3)
            feedback_comments = st.text_area("Additional feedback")
            if st.form_submit_button("Submit"):
                save_feedback(rating, feedback_comments)
                st.success("Feedback submitted! Thank you.")

def get_logo():
    if "logo_b64" not in st.session_state:
        st.session_state.logo_b64 = None
    if st.session_state.logo_b64:
        return f'<img src="data:image/png;base64,{st.session_state.logo_b64}" class="logo-left" alt="Custom Logo" style="display: block; margin: 0 auto; max-width: 200px;" />'
    default_logo_path = "logo.png"
    if os.path.exists(default_logo_path):
        with open(default_logo_path, "rb") as f:
            logo_bytes = f.read()
            return f'<img src="data:image/png;base64,{base64.b64encode(logo_bytes).decode()}" class="logo-left" alt="Default Logo" style="display: block; margin: 0 auto; max-width: 200px;" />'
    return '<div class="logo-left"></div>'

def theme_selector():
    theme = st.sidebar.selectbox("Select Theme", ["Light", "Dark"], index=0, key="theme_selector")
    if theme == "Dark":
        st.markdown("""
        <style>
        :root { --bg:#000; --text:#fff; --muted:#aaa; --section-bg: rgba(255,255,255,0.02); }
        html, body, .stApp { background:var(--bg)!important; color:var(--text)!important; }
        .app-title { font-size: 3rem; text-shadow: 2px 2px 4px rgba(255,255,255,0.3); }
        .nist-logo { color: var(--joval-red) !important; }
        .dark-theme [data-testid="stExpander"] > div:first-child { background: #111 !important; color: #fff !important; border: 1px solid rgba(255,255,255,0.12) !important; }
        .dark-theme .bottom-toolbar { background: linear-gradient(180deg, rgba(0,0,0,0.95), rgba(10,10,10,0.9)); border-top:1px solid rgba(255,255,255,0.03); }
        .dark-theme .progress-wrap { background: rgba(255,255,255,0.03); }
        .dark-theme .progress-fill { background: linear-gradient(90deg, var(--joval-red), #500010); }
        </style>
        """, unsafe_allow_html=True)
    return theme

@st.cache_data(ttl=300)
def export_to_excel(completed_map: Dict, comments_map: Dict, selected_playbook: str, bulk_export: bool = False) -> bytes:
    if not OPENPYXL_AVAILABLE:
        return b""
    df_completed = pd.DataFrame(list(completed_map.items()), columns=["Task_Key", "Status"])
    df_comments = pd.DataFrame(list(comments_map.items()), columns=["Task_Key", "Comment"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_completed.to_excel(writer, sheet_name="Progress", index=False)
        df_comments.to_excel(writer, sheet_name="Comments", index=False)
        if bulk_export:
            for pb in playbooks:
                if pb != selected_playbook:
                    comp, _ = load_progress(pb)
                    df_pb = pd.DataFrame(list(comp.items()), columns=["Task_Key", "Status"])
                    sheet_name = re.sub(r'[^\w\-_]', '_', pb.replace('.docx', ''))[:31]
                    df_pb.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

@st.cache_data
def export_to_csv(completed_map: Dict, comments_map: Dict, selected_playbook: str) -> bytes:
    df = pd.DataFrame({
        "Task_Key": list(completed_map.keys()) + list(comments_map.keys()),
        "Status": [str(completed_map.get(k, '')) for k in completed_map.keys()] + [''] * len(comments_map),
        "Comment": [''] * len(completed_map) + [str(v) for v in comments_map.values()]
    })
    return df.to_csv(index=False).encode('utf-8')

def create_jira_ticket(summary: str, description: str):
    try:
        import requests
        jira_url = st.secrets["JIRA_URL"]
        email = st.secrets["JIRA_EMAIL"]
        token = st.secrets["JIRA_TOKEN"]
        project_key = st.secrets["JIRA_PROJECT_KEY"]
        auth = base64.b64encode(f"{email}:{token}".encode()).decode()
        headers = {"Authorization": f"Basic {auth}", "Content-Type": "application/json"}
        data = json.dumps({
            "fields": {
                "project": {"key": project_key},
                "summary": summary,
                "description": description,
                "issuetype": {"name": "Task"}
            }
        })
        response = requests.post(jira_url, headers=headers, data=data)
        if response.status_code == 201:
            ticket_key = response.json()["key"]
            st.success(f"Jira ticket created: {ticket_key}")
            logging.info(f"Jira ticket created: {ticket_key}")
        else:
            st.error(f"Failed to create Jira ticket: {response.text}")
    except Exception as e:
        st.error(f"Jira integration error: {e}")

# === CACHED PARSING ===
@st.cache_data(hash_funcs={Path: lambda p: str(p)})
def parse_playbook_cached(path: str) -> List[Dict[str, Any]]:
    with open(path, "rb") as fh:
        result = mammoth.convert_to_html(fh)
        html = result.value
    soup = BeautifulSoup(html, "html.parser")

    exclude_terms = ["table of contents", "document control", "document revision", "assumptions", "disclaimer"]
    def excluded(text: str) -> bool:
        if not text:
            return False
        tl = text.strip().lower()
        return any(ex in tl for ex in exclude_terms)

    sections = []
    stack = []

    for tag in soup.find_all(['h1','h2','h3','h4','p','table','img']):
        if tag.name.startswith('h') and tag.name[1:].isdigit():
            title = tag.get_text().strip()
            if excluded(title):
                continue
            level = int(tag.name[1])
            node = {"title": title, "level": level, "content": [], "subs": []}
            while stack and stack[-1]["level"] >= level:
                stack.pop()
            if stack:
                stack[-1]["subs"].append(node)
            else:
                sections.append(node)
            stack.append(node)
        elif tag.name == 'p':
            text = tag.get_text(separator="\n").strip()
            if text and stack:
                stack[-1]["content"].append({"type": "text", "value": text})
        elif tag.name == 'img':
            src = tag.get("src", "")
            if src and stack:
                stack[-1]["content"].append({"type": "image", "value": src})
        elif tag.name == 'table':
            rows = [[td.get_text(separator="\n").strip() for td in tr.find_all(["td","th"])] for tr in tag.find_all("tr")]
            if rows and stack:
                stack[-1]["content"].append({"type": "table", "value": rows})

    def reconstruct_tables_in_section(section):
        contents = section.get("content", [])
        i = 0
        new_contents = []
        header_keywords = ["reference", "step", "description", "ownership", "responsibility"]
        owner_keywords = ["incident response team", "irt", "ownership", "responsibility", "it team leadership", "risk management team", "grc"]
        while i < len(contents):
            item = contents[i]
            if item["type"] != "text":
                new_contents.append(item)
                i += 1
                continue
            txt = item["value"].strip()
            txt_lower = txt.lower()
            keyword_count = sum(1 for word in header_keywords if word in txt_lower)
            is_header_like = keyword_count >= 2
            if is_header_like or ref_pattern.match(txt):
                headers = ["Reference", "Step", "Description", "Ownership/Responsibility"]
                rows = []
                current_ref = current_step = ""
                current_desc_parts = []
                j = i if not is_header_like else i + 1
                while j < len(contents) and contents[j]["type"] == "text":
                    txt_j = contents[j]["value"].strip()
                    if ref_pattern.match(txt_j):
                        if current_ref:
                            desc = " ".join(current_desc_parts)
                            owner = current_desc_parts.pop() if current_desc_parts and any(p in current_desc_parts[-1].lower() for p in owner_keywords) else ""
                            rows.append([current_ref, current_step, desc, owner])
                            current_desc_parts = []
                        match_obj = ref_pattern.match(txt_j)
                        current_ref = match_obj.group(0)
                        current_step = txt_j[match_obj.end():].strip()
                    else:
                        current_desc_parts.append(txt_j)
                    j += 1
                if current_ref:
                    desc = " ".join(current_desc_parts)
                    owner = current_desc_parts.pop() if current_desc_parts and any(p in current_desc_parts[-1].lower() for p in owner_keywords) else ""
                    rows.append([current_ref, current_step, desc, owner])
                if rows:
                    new_contents.append({"type": "table", "value": [headers] + rows})
                i = j
            else:
                new_contents.append(item)
                i += 1
        section["content"] = new_contents

    def walk_and_reconstruct(nodes):
        for n in nodes:
            reconstruct_tables_in_section(n)
            if n.get("subs"):
                walk_and_reconstruct(n["subs"])

    walk_and_reconstruct(sections)

    def prune(node):
        kept_subs = [sub for sub in node.get("subs", []) if prune(sub)]
        node["subs"] = kept_subs
        return bool(node.get("content")) or bool(kept_subs)

    return [s for s in sections if prune(s)]

ACTION_HEADERS = {"reference", "ref", "step", "description", "ownership", "responsibility", "owner", "responsible"}

def is_action_table(rows: List[List[str]]) -> bool:
    if not rows:
        return False
    headers = [h.strip().lower() for h in rows[0]]
    hits = sum(1 for h in headers if any(k in h for k in ACTION_HEADERS))
    if hits >= 2 or (len(rows[0]) >= 4 and ref_pattern.match(rows[0][0].strip())):
        return True
    return False

def render_action_table(playbook_name: str, sec_key: str, rows: List[List[str]], completed_map: dict, comments_map: dict, autosave: bool, table_index: int = 0):
    default_headers = ["Reference", "Step", "Description", "Ownership/Responsibility"]
    headers = rows[0] if len(rows) > 0 and not ref_pattern.match(rows[0][0].strip() if rows[0] else "") else default_headers
    data_rows = rows[1:] if len(rows) > 1 else rows
    for row in data_rows:
        while len(row) < 4:
            row.append("")
    st.caption("Mark tasks complete and add notes.")
    cols = st.columns([1, 2, 4, 2, 1, 2])
    for i, h in enumerate(["Ref", "Step", "Desc", "Owner", "Done", "Comment"]):
        cols[i].write(h)
    changed = False
    table_key = f"{sec_key}::tbl::{table_index}"
    for ridx, row in enumerate(data_rows):
        row_key = f"{table_key}::row::{ridx}"
        comment_key = f"{row_key}::comment"
        ref = row[0]
        step = row[1]
        desc = " ".join(row[2:-1])
        owner = row[-1]
        prev_val = completed_map.get(row_key, False)
        prev_comment = comments_map.get(comment_key, "")
        cols = st.columns([1, 2, 4, 2, 1, 2])
        cols[0].write(ref)
        cols[1].write(step)
        cols[2].write(desc)
        cols[3].write(owner)
        new_val = cols[4].checkbox("", value=prev_val, key=f"cb_{row_key}")
        new_comment = cols[5].text_input("", value=prev_comment, key=f"ci_{comment_key}", label_visibility="collapsed")
        if new_val != prev_val:
            completed_map[row_key] = new_val
            changed = True
        if new_comment != prev_comment:
            comments_map[comment_key] = new_comment
            changed = True
    if autosave and changed:
        save_progress(playbook_name, completed_map, comments_map)

def render_generic_table(rows: List[List[str]]):
    if len(rows) > 1:
        df = pd.DataFrame(rows[1:], columns=rows[0])
    else:
        df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

def render_section_content(section: Dict[str, Any], playbook_name: str, completed_map: dict, comments_map: dict, autosave: bool, sec_key: str, is_sub: bool = False):
    table_idx = 0
    for item in section.get("content", []):
        t = item.get("type")
        if t == "text":
            text = item.get("value", "").replace("\n", "<br/>")
            st.markdown(f'<div class="content-text">{text}</div>', unsafe_allow_html=True)
        elif t == "image":
            safe_image_display(item.get("value", ""))
        elif t == "table":
            rows = item.get("value", [])
            if rows:
                if is_action_table(rows):
                    render_action_table(playbook_name, sec_key, rows, completed_map, comments_map, autosave, table_idx)
                    table_idx += 1
                else:
                    render_generic_table(rows)
    for sub in section.get("subs", []):
        sub_key = stable_key(playbook_name, sub["title"], sub["level"])
        st.markdown(f"<div id='{sub_key}' style='margin-top:12px;'><strong style='color:var(--text);'>{sub['title']}</strong></div>", unsafe_allow_html=True)
        render_section_content(sub, playbook_name, completed_map, comments_map, autosave, sub_key, True)
    if not is_sub:
        st.markdown("<div class='comments-title'>Comments / Notes</div>", unsafe_allow_html=True)
        prev_sec_comment = comments_map.get(sec_key, "")
        new_sec_comment = st.text_area("", value=prev_sec_comment, key=f"c::{sec_key}", height=120, label_visibility="collapsed")
        if new_sec_comment != prev_sec_comment:
            comments_map[sec_key] = new_sec_comment
            if autosave:
                save_progress(playbook_name, completed_map, comments_map)

def render_section(section: Dict[str, Any], playbook_name: str, completed_map: dict, comments_map: dict, autosave: bool):
    sec_key = stable_key(playbook_name, section["title"], section["level"])
    title_class = "nist-incident-section" if section["title"] == "NIST Incident Handling Categories" else ""
    st.markdown(f"<div class='section-title {title_class}' id='{sec_key}'>{section['title']}</div>", unsafe_allow_html=True)
    with st.expander("Expand section", expanded=False):
        render_section_content(section, playbook_name, completed_map, comments_map, autosave, sec_key)

# === PDF GENERATION – FIXED WITH VALIDATION ===
@st.cache_data(ttl=300)
def generate_pdf_bytes(sections: List[Dict[str, Any]], playbook_name: str) -> bytes:
    try:
        pdf = FPDF()
        pdf.set_margins(15, 10, 15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 10, txt=clean_for_pdf("Joval Wines - NIST Playbook Tracker"), ln=1, align='C')
        pdf.set_font("Arial", size=10)
        pdf.cell(0, 8, txt=clean_for_pdf(f"Playbook: {playbook_name}"), ln=1)
        pdf.ln(4)

        def add_section(pdf, section, indent=0):
            pdf.set_font("Arial", "B", 8)
            pdf.multi_cell(180, 6, clean_for_pdf("  " * indent + section["title"]), 0, 'L')
            pdf.ln(2)
            pdf.set_font("Arial", size=7)
            for it in section.get("content", []):
                if it["type"] == "text":
                    for line in it["value"].split("\n"):
                        pdf.multi_cell(180, 5, clean_for_pdf(line.strip()), 0, 'L')
                        pdf.ln(1)
                elif it["type"] == "table":
                    pdf.multi_cell(180, 5, clean_for_pdf("[Table]"), 0, 'L')
                    pdf.ln(1)
            for sub in section.get("subs", []):
                add_section(pdf, sub, indent + 2)
            pdf.ln(2)

        for s in sections:
            add_section(pdf, s)

        return pdf.output(dest='S')  # Returns bytes directly
    except Exception as e:
        st.error(f"PDF generation failed: {str(e)}")
        return b""

# === FULL-TEXT SEARCH ===
@st.cache_data(ttl=600)
def simple_search(query: str, playbooks_list: List[str], top_k: int = 12):
    results = []
    query_lower = query.lower()

    for pb in playbooks_list:
        path = os.path.join(PLAYBOOKS_DIR, pb)
        secs = parse_playbook_cached(path)

        for s in secs:
            parts = [s.get("title", "")]
            for c in s.get("content", []):
                if c.get("type") == "text":
                    parts.append(c.get("value", ""))
                elif c.get("type") == "table":
                    for row in c.get("value", []):
                        parts.append(" ".join(row))
            full_text = "\n".join(parts).lower()

            if query_lower in full_text:
                pos = full_text.find(query_lower)
                start = max(0, pos - 40)
                snippet = "..." + full_text[start:start + 120].replace("\n", " ") + "..."
                results.append({
                    "playbook": pb,
                    "title": s.get("title", ""),
                    "level": s.get("level", 2),
                    "anchor": stable_key(pb, s.get("title", ""), s.get("level", 2)),
                    "snippet": snippet
                })

    results.sort(key=lambda x: (len(x["title"]), x["snippet"].find(query_lower)))
    return results[:top_k]

# === MAIN APP ===
def main():
    user = authenticate()
    st.sidebar.info(f"Logged in as: {user['name']} ({user['email']}) - Role: {get_user_role(user['email'])}")

    if st.session_state.get('admin_page', False):
        admin_dashboard(user)
        return

    if get_user_role(user['email']) == "admin":
        if st.sidebar Sven("Admin Dashboard"):
            st.session_state.admin_page = True
            st.rerun()

    if 'gamify' not in st.session_state:
        st.session_state.gamify = False
    if 'gamify_count' not in st.session_state:
        st.session_state.gamify_count = 0

    theme_selector()

    logo_html = get_logo()

    global playbooks
    playbooks = sorted([f for f in os.listdir(PLAYBOOKS_DIR) if f.lower().endswith(".docx")])
    if not playbooks:
        st.error(f"No .docx files found in '{PLAYBOOKS_DIR}'.")
        return

    st.markdown(f"""
    <div class='sticky-header'>
        {logo_html}
        <div class='app-title'>Joval Wines NIST Playbook Tracker</div>
        <div class='nist-logo'>NIST</div>
    </div>
    """, unsafe_allow_html=True)

    # === SEARCH UI ===
    st.sidebar.markdown('<div class="sidebar-header">Search Playbooks</div>', unsafe_allow_html=True)
    query = st.sidebar.text_input("Search", key="search_query", placeholder="type keyword…")
    search_btn = st.sidebar.button("Search", key="search_btn")

    if search_btn and query.strip():
        results = simple_search(query.strip(), playbooks, 12)
        if results:
            st.sidebar.markdown("<h4 style='color:var(--joval-red);margin-top:12px;'>Results</h4>", unsafe_allow_html=True)
            for idx, r in enumerate(results):
                clean_name = r["playbook"].replace(".docx", "").split(" v")[0]
                anchor = r["anchor"]
                btn_key = f"nav_{idx}_{anchor}"

                if st.sidebar.button(
                    f"Go to {r['title']}",
                    key=btn_key,
                    help=r["snippet"],
                    use_container_width=True
                ):
                    st.session_state.select_playbook = r["playbook"]
                    st.session_state.pending_anchor = anchor
                    st.session_state.pending_highlight = query.strip()
                    st.rerun()
        else:
            st.sidebar.info("No matches found.")
    else:
        st.sidebar.markdown("<em style='color:#777;'>Enter a term to search across all playbooks.</em>", unsafe_allow_html=True)

    st.sidebar.markdown("---")
    autosave = st.sidebar.checkbox("Auto-save progress", value=True)
    bulk_export = st.sidebar.checkbox("Bulk export")
    st.sidebar.markdown("---")
    st.sidebar.markdown('<div class="sidebar-subheader">NIST Resources</div>', unsafe_allow_html=True)
    resources = {
        "Cybersecurity Framework": "https://www.nist.gov/cyberframework",
        "Incident Response (SP 800-61 Rev2)": "https://csrc.nist.gov/publications/detail/sp/800-61/rev-2/final",
        "Risk Management Framework": "https://csrc.nist.gov/projects/risk-management",
        "NICE Resources": "https://www.nist.gov/itl/applied-cybersecurity/nice/resources",
    }
    sel = st.sidebar.selectbox("Choose resource", ["(none)"] + list(resources.keys()))
    if sel != "(none)":
        st.sidebar.markdown(f"[Open → {sel}]({resources[sel]})", unsafe_allow_html=True)

    st.sidebar.markdown("---")
    st.sidebar.markdown('<div class="sidebar-footer">© Joval Wines</div>', unsafe_allow_html=True)
    st.sidebar.markdown('<div style="color: var(--text); font-weight: bold; font-size: 1.1rem;">Better Never Stops</div>', unsafe_allow_html=True)

    st.markdown('<div class="playbook-select-label">Select playbook</div>', unsafe_allow_html=True)
    selected_playbook = st.selectbox(
        "Select playbook",
        playbooks,
        index=playbooks.index(st.session_state.select_playbook) if "select_playbook" in st.session_state and st.session_state.select_playbook in playbooks else 0,
        key="select_playbook"
    )
    st.markdown('<div class="instructional-text">In the event of a cyber incident select the required playbook and complete each required step in the NIST "Incident Handling Categories" section</div>', unsafe_allow_html=True)

    # === LOAD PLAYBOOK ===
    parsed_key = f"parsed::{selected_playbook}"
    if parsed_key not in st.session_state:
        st.session_state[parsed_key] = parse_playbook_cached(os.path.join(PLAYBOOKS_DIR, selected_playbook))
    sections = st.session_state[parsed_key]

    if not sections:
        st.error("No playbook sections loaded—check playbooks folder.")
        st.stop()

    completed_map, comments_map = load_progress(selected_playbook)
    st.session_state[f"completed::{selected_playbook}"] = completed_map
    st.session_state[f"comments::{selected_playbook}"] = comments_map
    completed_map = st.session_state[f"completed::{selected_playbook}"]
    comments_map = st.session_state[f"comments::{selected_playbook}"]

    # === HANDLE PENDING NAVIGATION FROM SEARCH ===
    if "pending_anchor" in st.session_state:
        anchor_id = st.session_state.pending_anchor
        highlight_term = st.session_state.get("pending_highlight", "")
        
        js = f"""
        setTimeout(() => {{
            expandSection('{anchor_id}');
            scrollToSection('{anchor_id}');
            {f"highlightText('{anchor_id}', '{highlight_term}');" if highlight_term else ""}
        }}, 800);
        """
        st.markdown(f"<script>{js}</script>", unsafe_allow_html=True)
        
        # Clean up
        del st.session_state.pending_anchor
        if "pending_highlight" in st.session_state:
            del st.session_state.pending_highlight

    # === TOC ===
    toc_items = []
    def walk_toc(secs):
        for s in secs:
            anchor = stable_key(selected_playbook, s["title"], s["level"])
            toc_items.append({"title": s["title"], "anchor": anchor})
            if s.get("subs"):
                walk_toc(s["subs"])
    walk_toc(sections)

    toc_html = "<div class='toc'><h4>Table of Contents</h4>" + "".join(
        f"<a onclick=\"scrollToSection('{t['anchor']}')\" style='cursor:pointer;'>{t['title']}</a>"
        for t in toc_items
    ) + "</div>"
    st.markdown(toc_html, unsafe_allow_html=True)

    st.markdown('<div class="content-wrap">', unsafe_allow_html=True)
    for sec in sections:
        render_section(sec, selected_playbook, completed_map, comments_map, autosave)
    st.markdown('</div>', unsafe_allow_html=True)

    total_checks = len([k for k in completed_map if '::row::' in k])
    done_checks = sum(1 for v in completed_map.values() if v)
    pct = int((done_checks / max(total_checks, 1)) * 100)

    badges = calculate_badges(pct)
    col1, col2 = st.columns([3,1])
    with col1:
        st.info(f"Progress: {pct}% - {', '.join(badges)}")
    with col2:
        if st.button("Gamify!"):
            st.session_state.gamify = not st.session_state.gamify
            if st.session_state.gamify:
                st.session_state.gamify_count += 1
                if st.session_state.gamify_count % 2 == 1:
                    st.balloons()
                else:
                    st.snow()

    st.progress(pct / 100)
    if st.button("Refresh"):
        st.rerun()

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Save Progress"):
            path = save_progress(selected_playbook, completed_map, comments_map)
            st.success(f"Saved to {os.path.basename(path)}")
        if st.button("Create Jira Ticket"):
            summary = st.text_input("Ticket Summary", "Incident Response Progress")
            desc = f"Progress: {pct}% for {selected_playbook}"
            if st.button("Confirm Create"):
                create_jira_ticket(summary, desc)
    with c2:
        csv_data = export_to_csv(completed_map, comments_map, selected_playbook)
        st.download_button("Download CSV", csv_data, f"{os.path.splitext(selected_playbook)[0]}_progress.csv", "text/csv")
        if OPENPYXL_AVAILABLE:
            excel_data = export_to_excel(completed_map, comments_map, selected_playbook, bulk_export)
            st.download_button("Download Excel", excel_data, f"{os.path.splitext(selected_playbook)[0]}_progress.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        pdf_bytes = generate_pdf_bytes(sections, selected_playbook)
        if pdf_bytes and len(pdf_bytes) > 200:
            st.download_button(
                label="Export PDF",
                data=pdf_bytes,
                file_name=f"{os.path.splitext(selected_playbook)[0]}_export.pdf",
                mime="application/pdf"
            )
        else:
            st.warning("PDF export unavailable.")

    if autosave:
        save_progress(selected_playbook, completed_map, comments_map)

    show_feedback()

if __name__ == "__main__":
    main()
