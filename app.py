# app.py
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

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

import logging

# -------------------------------------------------
# CONFIG & LOGGING
# -------------------------------------------------
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

# -------------------------------------------------
# PAGE CONFIG & GLOBAL CSS (jovalwines.com.au style)
# -------------------------------------------------
st.set_page_config(
    page_title="Joval Wines NIST Playbook Tracker",
    page_icon="wine",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
/* ---------- Tailwind CDN (modern look) ---------- */
@import url('https://cdn.tailwindcss.com');

/* ---------- Core colours (jovalwines.com.au) ---------- */
:root{
    --bg:#ffffff;
    --text:#111111;
    --muted:#666666;
    --red:#800020;
    --card-bg:#fafafa;
}

/* ---------- Global ---------- */
html,body,.stApp{background:var(--bg)!important;color:var(--text)!important;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;}
.stApp > footer,.stApp [data-testid="stToolbar"],.stApp [data-testid="collapsedControl"],.stDeployButton{display:none!important;}

/* ---------- Header ---------- */
.sticky-header{
    position:sticky;top:0;z-index:9999;
    display:flex;align-items:center;justify-content:space-between;
    padding:1rem 2rem;background:#fff;
    border-bottom:1px solid #eee;box-shadow:0 2px 8px rgba(0,0,0,.05);
}
.logo-left{height:70px;}
.app-title{font-size:2.2rem;font-weight:700;color:var(--text);margin:0;}
.nist-logo{font-size:1.8rem;color:var(--red);font-weight:700;}

/* ---------- Sidebar ---------- */
.css-1d391kg{padding-top:1rem;}
.sidebar-header{font-weight:600;font-size:1.1rem;margin-bottom:.5rem;}
.sidebar-subheader{font-weight:600;margin-top:1rem;margin-bottom:.5rem;}

/* ---------- Main content ---------- */
.content-wrap{margin-left:280px;padding:2rem 2rem 6rem;}
.section-card{
    background:var(--card-bg);padding:1.5rem;border-radius:12px;
    margin-bottom:1.5rem;box-shadow:0 2px 6px rgba(0,0,0,.04);
    border:1px solid #eaeaea;
}
.section-title{font-size:1.7rem;font-weight:700;margin-bottom:.75rem;color:var(--text);}
.nist-incident-section{color:var(--red)!important;}

/* ---------- Buttons ---------- */
.stButton>button,.stDownloadButton>button{
    background:#000!important;color:#fff!important;
    border-radius:8px;padding:.5rem 1rem;font-weight:600;
}
.stButton>button:hover,.stDownloadButton>button:hover{opacity:.9;}

/* ---------- Progress bar ---------- */
.progress-wrap{height:12px;background:#e5e5e5;border-radius:999px;overflow:hidden;}
.progress-fill{height:100%;background:var(--red);transition:width .4s ease;}

/* ---------- Bottom toolbar ---------- */
.bottom-toolbar{
    position:fixed;bottom:0;left:0;right:0;z-index:999;
    background:#fff;border-top:1px solid #eee;
    padding:.75rem 2rem;display:flex;align-items:center;justify-content:space-between;
    box-shadow:0 -2px 8px rgba(0,0,0,.03);
}

/* ---------- Responsive ---------- */
@media (max-width:768px){
    .sticky-header{flex-direction:column;padding:1rem;}
    .app-title{font-size:1.8rem;}
    .content-wrap{margin-left:0;padding:1rem;}
}
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# USER MANAGEMENT (unchanged)
# -------------------------------------------------
def load_users(): ...
def save_users(users): ...
def get_user_role(email): ...
def create_user(email, role, password): ...
def reset_user_password(email, password): ...
def delete_user(email): ...
def authenticate(): ...

# -------------------------------------------------
# ADMIN DASHBOARD (unchanged)
# -------------------------------------------------
def admin_dashboard(user): ...

# -------------------------------------------------
# UTILITIES (unchanged except PDF generation removed)
# -------------------------------------------------
def stable_key(playbook_name: str, title: str, level: int) -> str: ...
@st.cache_data
def progress_filepath(playbook_name: str) -> str: ...
@st.cache_data
def load_progress(playbook_name: str): ...
def save_progress(playbook_name: str, completed_map: dict, comments_map: dict) -> str: ...
def safe_image_display(src: str) -> bool: ...
def calculate_badges(pct: int) -> List[str]: ...
def save_feedback(rating: int, comments: str): ...
def show_feedback(): ...
def get_logo(): ...
def theme_selector(): ...

# ---------- EXCEL / CSV (unchanged) ----------
@st.cache_data(ttl=300)
def export_to_excel(completed_map: Dict, comments_map: Dict, selected_playbook: str, bulk_export: bool = False) -> bytes: ...
@st.cache_data
def export_to_csv(completed_map: Dict, comments_map: Dict, selected_playbook: str) -> bytes: ...

# ---------- JIRA ----------
def create_jira_ticket(summary: str, description: str): ...

# -------------------------------------------------
# PLAYBOOK PARSING (unchanged)
# -------------------------------------------------
@st.cache_data(hash_funcs={Path: lambda p: str(p)})
def parse_playbook_cached(path: str) -> List[Dict[str, Any]]: ...

# -------------------------------------------------
# RENDERING HELPERS
# -------------------------------------------------
ACTION_HEADERS = {"reference","ref","step","description","ownership","responsibility","owner","responsible"}

def is_action_table(rows: List[List[str]]) -> bool: ...

def render_action_table(playbook_name, sec_key, rows, completed_map, comments_map, autosave, table_index=0): ...
def render_generic_table(rows: List[List[str]]): ...
def render_section_content(section, playbook_name, completed_map, comments_map, autosave, sec_key, is_sub=False): ...
def render_section(section, playbook_name, completed_map, comments_map, autosave): ...

# -------------------------------------------------
# MAIN APP
# -------------------------------------------------
def main():
    user = authenticate()
    st.sidebar.info(f"Logged in as: **{user['name']}** – *{get_user_role(user['email'])}*")

    # ----- ADMIN -----
    if get_user_role(user["email"]) == "admin":
        if st.sidebar.button("Admin Dashboard"):
            st.session_state.admin_page = True
            st.rerun()
    if st.session_state.get('admin_page', False):
        admin_dashboard(user)
        return

    # ----- GAMIFY -----
    if 'gamify' not in st.session_state: st.session_state.gamify = False
    if 'gamify_count' not in st.session_state: st.session_state.gamify_count = 0

    theme_selector()

    # ----- HEADER -----
    logo_html = get_logo()
    st.markdown(f"""
    <div class="sticky-header">
        {logo_html}
        <div class="app-title">Joval Wines NIST Playbook Tracker</div>
        <div class="nist-logo">NIST</div>
    </div>
    """, unsafe_allow_html=True)

    # ----- SIDEBAR CONTROLS (search removed) -----
    st.sidebar.markdown('<div class="sidebar-header">Controls</div>', unsafe_allow_html=True)
    autosave = st.sidebar.checkbox("Auto-save progress", value=True)
    bulk_export = st.sidebar.checkbox("Bulk export")
    st.sidebar.markdown("---")
    st.sidebar.markdown('<div class="sidebar-subheader">NIST Resources</div>', unsafe_allow_html=True)
    resources = {
        "Cybersecurity Framework":"https://www.nist.gov/cyberframework",
        "Incident Response (SP 800-61 Rev2)":"https://csrc.nist.gov/publications/detail/sp/800-61/rev-2/final",
        "Risk Management Framework":"https://csrc.nist.gov/projects/risk-management",
        "NICE Resources":"https://www.nist.gov/itl/applied-cybersecurity/nice/resources",
    }
    sel = st.sidebar.selectbox("Choose resource", ["(none)"] + list(resources.keys()))
    if sel != "(none)":
        st.sidebar.markdown(f"[Open → {sel}]({resources[sel]})", unsafe_allow_html=True)

    st.sidebar.markdown("---")
    st.sidebar.markdown('<div style="font-weight:600;">© Joval Wines</div>', unsafe_allow_html=True)
    st.sidebar.markdown('<div style="font-weight:600;">Better Never Stops</div>', unsafe_allow_html=True)

    # ----- PLAYBOOK SELECT -----
    playbooks = sorted([f for f in os.listdir(PLAYBOOKS_DIR) if f.lower().endswith(".docx")])
    if not playbooks:
        st.error(f"No .docx files found in '{PLAYBOOKS_DIR}'.")
        return

    st.markdown("<h2 style='margin-top:2rem;'>Select Playbook</h2>", unsafe_allow_html=True)
    selected_playbook = st.selectbox(
        "Playbook",
        options=[""] + playbooks,
        index=0,
        key="select_playbook"
    )

    if not selected_playbook:
        st.markdown("""
        <div style="text-align:center;margin:4rem 0;font-size:1.3rem;color:#800020;font-weight:600;">
            Please select a playbook from the dropdown above to begin.
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    # ----- INSTRUCTIONAL BANNER -----
    st.markdown("""
    <div style="background:#fff3cd;padding:1.2rem;border-radius:8px;border:2px solid #d9534f;
                text-align:center;font-size:1.4rem;font-weight:600;color:#d9534f;margin:1.5rem 0;">
        In the event of a cyber incident select the required playbook and complete each required step
        in the <strong>NIST "Incident Handling Categories"</strong> section.
    </div>
    """, unsafe_allow_html=True)

    # ----- LOAD PLAYBOOK -----
    parsed_key = f"parsed::{selected_playbook}"
    if parsed_key not in st.session_state:
        st.session_state[parsed_key] = parse_playbook_cached(os.path.join(PLAYBOOKS_DIR, selected_playbook))
    sections = st.session_state[parsed_key]

    completed_map, comments_map = load_progress(selected_playbook)
    st.session_state[f"completed::{selected_playbook}"] = completed_map
    st.session_state[f"comments::{selected_playbook}"] = comments_map
    completed_map = st.session_state[f"completed::{selected_playbook}"]
    comments_map = st.session_state[f"comments::{selected_playbook}"]

    # ----- TOC (fixed left side) -----
    toc_items = []
    def walk_toc(secs):
        for s in secs:
            anchor = stable_key(selected_playbook, s["title"], s["level"])
            toc_items.append({"title": s["title"], "anchor": anchor})
            if s.get("subs"): walk_toc(s["subs"])
    walk_toc(sections)

    toc_html = "<div style='position:fixed;left:1rem;top:110px;bottom:100px;width:250px;background:#fff;padding:1rem;border-radius:8px;overflow:auto;box-shadow:0 2px 6px rgba(0,0,0,.04);border:1px solid #eaeaea;'><h4 style='margin-top:0;'>Table of Contents</h4>" + "".join(
        f"<a href='#' onclick=\"document.getElementById('{a[\"anchor\"]}').scrollIntoView({{behavior:'smooth'}});return false;\" style='display:block;padding:4px 0;color:#111;text-decoration:none;'>{a['title']}</a>"
        for a in toc_items
    ) + "</div>"
    st.markdown(toc_html, unsafe_allow_html=True)

    # ----- CONTENT -----
    st.markdown('<div class="content-wrap">', unsafe_allow_html=True)
    for sec in sections:
        render_section(sec, selected_playbook, completed_map, comments_map, autosave)
    st.markdown('</div>', unsafe_allow_html=True)

    # ----- PROGRESS -----
    total_checks = len([k for k in completed_map if '::row::' in k])
    done_checks = sum(1 for v in completed_map.values() if v)
    pct = int((done_checks / max(total_checks, 1)) * 100)
    badges = calculate_badges(pct)

    col1, col2 = st.columns([3, 1])
    with col1:
        st.info(f"**Progress:** {pct}% – {', '.join(badges)}")
    with col2:
        if st.button("Gamify!"):
            st.session_state.gamify = not st.session_state.gamify
            if st.session_state.gamify:
                st.session_state.gamify_count += 1
                if st.session_state.gamify_count % 2 == 1:
                    st.balloons()
                else:
                    st.snow()

    st.markdown(f"<div class='progress-wrap'><div class='progress-fill' style='width:{pct}%'></div></div>", unsafe_allow_html=True)

    if st.button("Refresh"):
        st.rerun()

    # ----- ACTION BUTTONS (CSV / Excel / Jira) -----
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Save Progress"):
            path = save_progress(selected_playbook, completed_map, comments_map)
            st.success(f"Saved to `{os.path.basename(path)}`")
        if st.button("Create Jira Ticket"):
            summary = st.text_input("Ticket Summary", "Incident Response Progress")
            desc = f"Progress: {pct}% for {selected_playbook}"
            if st.button("Confirm"):
                create_jira_ticket(summary, desc)
    with c2:
        csv_data = export_to_csv(completed_map, comments_map, selected_playbook)
        st.download_button("Download CSV", csv_data,
                           f"{os.path.splitext(selected_playbook)[0]}_progress.csv",
                           "text/csv")
        if OPENPYXL_AVAILABLE:
            excel_data = export_to_excel(completed_map, comments_map, selected_playbook, bulk_export)
            st.download_button("Download Excel", excel_data,
                               f"{os.path.splitext(selected_playbook)[0]}_progress.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        pass   # PDF button removed

    # ----- AUTO-SAVE & FEEDBACK -----
    if autosave:
        save_progress(selected_playbook, completed_map, comments_map)

    show_feedback()

    # ----- BOTTOM TOOLBAR -----
    st.markdown(f"""
    <div class="bottom-toolbar">
        <div>© Joval Wines – Better Never Stops</div>
        <div>Progress: {pct}%</div>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
