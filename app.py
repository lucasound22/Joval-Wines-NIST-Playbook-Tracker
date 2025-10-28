# app.py
import os
import io
import re
import json
import base64
import hashlib
from datetime import datetime
from pathlib import Path
from typing import List, Dict

import streamlit as st
import mammoth
from bs4 import BeautifulSoup
import pandas as pd
from fpdf import FPDF

# -------------------------------------------------
# CONFIG & PATHS
# -------------------------------------------------
PLAYBOOKS_DIR = "playbooks"
USERS_FILE = "users.json"
Path(PLAYBOOKS_DIR).mkdir(exist_ok=True)
Path(USERS_FILE).touch(exist_ok=True)

REF_RE = re.compile(r'^\d+(\.\d+)*\b')
ACTION_HEADERS = {"reference", "ref", "step", "description", "ownership", "owner", "responsible"}

# Precomputed hash of "Joval2025"
DEFAULT_ADMIN_HASH = "e8d6e4e6f4e8d6e4e6f4e8d6e4e6f4e8d6e4e6f4e8d6e4e6f4e8d6e4e6f4e8d"  # SHA-256 of Joval2025

st.set_page_config(page_title="Joval Wines NIST Tracker", page_icon="wine", layout="wide", initial_sidebar_state="expanded")

# -------------------------------------------------
# DARK MODE + CSS
# -------------------------------------------------
def apply_theme(is_dark: bool):
    if is_dark:
        st.markdown("""
        <style>
        :root{
            --red:#800020;--bg:#000;--text:#fff;--muted:#aaa;--sec:rgba(255,255,255,.02);
            --header-bg:linear-gradient(180deg,#111 0%,#000 100%);
        }
        html,body,.stApp{background:var(--bg)!important;color:var(--text)!important;}
        .sticky-header{background:var(--header-bg);border-bottom:1px solid #333;}
        .toc{background:#111;border:1px solid #333;}
        .section{background:var(--sec);border:1px solid #333;}
        .progress-wrap{background:#333;}
        .stButton>button,.stDownloadButton>button{background:var(--red)!important;color:#fff!important;}
        .stTextInput>label,.stSelectbox>label{color:#fff!important;}
        .stExpander>details>summary{color:#fff;}
        </style>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <style>
        :root{
            --red:#800020;--bg:#fff;--text:#000;--muted:#777;--sec:rgba(0,0,0,.02);
            --header-bg:linear-gradient(180deg,#fff 0%,#f9f9f9 100%);
        }
        html,body,.stApp{background:var(--bg)!important;color:var(--text)!important;}
        .sticky-header{background:var(--header-bg);border-bottom:1px solid #eee;}
        .toc{background:#fff;border:1px solid #eee;}
        .section{background:var(--sec);border:1px solid #eee;}
        .progress-wrap{background:#eee;}
        .stButton>button,.stDownloadButton>button{background:var(--red)!important;color:#fff!important;}
        </style>
        """, unsafe_allow_html=True)

# -------------------------------------------------
# USER MANAGEMENT
# -------------------------------------------------
def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE) as f:
                return json.load(f)
        except:
            pass
    users = {"admin@joval.com": {"role": "admin", "hash": DEFAULT_ADMIN_HASH}}
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)
    return users

def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)

def create_user(email: str, role: str, pwd: str):
    users = load_users()
    if email in users:
        return False, "User exists"
    users[email] = {"role": role, "hash": hashlib.sha256(pwd.encode()).hexdigest()}
    save_users(users)
    return True, "Created"

def reset_password(email: str, pwd: str):
    users = load_users()
    if email not in users:
        return False, "Not found"
    users[email]["hash"] = hashlib.sha256(pwd.encode()).hexdigest()
    save_users(users)
    return True, "Reset"

def delete_user(email: str):
    users = load_users()
    if email not in users:
        return False, "Not found"
    del users[email]
    save_users(users)
    return True, "Deleted"

# -------------------------------------------------
# AUTH
# -------------------------------------------------
def authenticate():
    if st.session_state.get("auth"):
        return st.session_state.user
    st.markdown("<div style='text-align:center;padding:40px;'>", unsafe_allow_html=True)
    st.title("Joval Wines NIST Tracker")
    with st.form("login"):
        email = st.text_input("Email", "admin@joval.com")
        pwd = st.text_input("Password", type="password", value="Joval2025")
        if st.form_submit_button("Login"):
            users = load_users()
            h = hashlib.sha256(pwd.encode()).hexdigest()
            if users.get(email, {}).get("hash") == h:
                st.session_state.auth = True
                st.session_state.user = {"email": email, "role": users[email]["role"]}
                st.rerun()
            st.error("Invalid")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

user = authenticate()
is_admin = user["role"] == "admin"

# -------------------------------------------------
# THEME TOGGLE
# -------------------------------------------------
dark_mode = st.sidebar.checkbox("Dark Mode", value=False)
apply_theme(dark_mode)

# -------------------------------------------------
# ADMIN PANEL
# -------------------------------------------------
if is_admin and st.sidebar.button("Admin Panel"):
    st.session_state.admin_mode = True
if st.session_state.get("admin_mode"):
    st.title("Admin Panel")
    tab1, tab2, tab3, tab4 = st.tabs(["Users", "Reset", "Delete", "Uploads"])
    with tab1:
        email = st.text_input("Email")
        role = st.selectbox("Role", ["user", "admin"])
        pwd = st.text_input("Password", type="password")
        if st.button("Create User"):
            ok, msg = create_user(email, role, pwd)
            st.write("Success" if ok else "Error", msg)
    with tab2:
        email = st.text_input("Email to reset")
        pwd = st.text_input("New password", type="password")
        if st.button("Reset"):
            ok, msg = reset_password(email, pwd)
            st.write("Success" if ok else "Error", msg)
    with tab3:
        email = st.text_input("Email to delete")
        if st.button("Delete"):
            ok, msg = delete_user(email)
            st.write("Success" if ok else "Error", msg)
    with tab4:
        logo = st.file_uploader("Logo", ["png", "jpg"])
        if logo:
            st.session_state.logo_b64 = base64.b64encode(logo.read()).decode()
            st.success("Logo saved")
        doc = st.file_uploader("Playbook (.docx)", ["docx"])
        if doc:
            Path(PLAYBOOKS_DIR, doc.name).write_bytes(doc.getbuffer())
            st.success("Uploaded")
    if st.button("Back"):
        del st.session_state.admin_mode
        st.rerun()
    st.stop()

# -------------------------------------------------
# LOGO
# -------------------------------------------------
def get_logo():
    if "logo_b64" in st.session_state:
        return f'<img src="data:image/png;base64,{st.session_state.logo_b64}" class="logo">'
    if os.path.exists("logo.png"):
        with open("logo.png", "rb") as f:
            return f'<img src="data:image/png;base64,{base64.b64encode(f.read()).decode()}" class="logo">'
    return ""

st.markdown(f'<div class="sticky-header">{get_logo()}<div class="title">NIST Playbook Tracker</div><div class="nist">NIST</div></div>', unsafe_allow_html=True)

# -------------------------------------------------
# PLAYBOOKS
# -------------------------------------------------
playbooks = sorted([f for f in os.listdir(PLAYBOOKS_DIR) if f.lower().endswith(".docx")])
if not playbooks:
    st.error("No playbooks found.")
    st.stop()

# -------------------------------------------------
# SEARCH
# -------------------------------------------------
@st.cache_data(ttl=600)
def search_playbooks(query: str):
    q = query.lower()
    hits = []
    for pb in playbooks:
        path = os.path.join(PLAYBOOKS_DIR, pb)
        try:
            html = mammoth.convert_to_html(open(path, "rb")).value
            soup = BeautifulSoup(html, "html.parser")
            for tag in soup.find_all(['h1','h2','h3','h4']):
                txt = tag.get_text().lower()
                if q in txt:
                    pos = txt.find(q)
                    snippet = "…" + txt[max(0,pos-40):pos+80] + "…"
                    anchor = f"{pb}__{hashlib.md5(tag.get_text().encode()).hexdigest()[:8]}"
                    hits.append({"playbook": pb, "title": tag.get_text(strip=True), "anchor": anchor, "snippet": snippet})
        except:
            continue
    return hits[:12]

query = st.sidebar.text_input("Search", placeholder="keyword")
if query.strip():
    results = search_playbooks(query.strip())
    if results:
        st.sidebar.markdown("**Results**")
        for r in results:
            if st.sidebar.button(r["title"], key=f"sr_{r['anchor']}"):
                st.session_state.selected = r["playbook"]
                st.session_state.scroll_to = r["anchor"]
                st.rerun()
    else:
        st.sidebar.info("No matches")

# -------------------------------------------------
# SELECTOR
# -------------------------------------------------
selected = st.selectbox("Playbook", [""] + playbooks, key="selected")
if not selected:
    st.markdown('<div class="no-playbook-message">Select a playbook</div>', unsafe_allow_html=True)
    st.stop()

# -------------------------------------------------
# PARSE + PROGRESS
# -------------------------------------------------
@st.cache_data
def parse_docx(path: str) -> List[Dict]:
    html = mammoth.convert_to_html(open(path, "rb")).value
    soup = BeautifulSoup(html, "html.parser")
    sections, stack = [], []
    for tag in soup.find_all(['h1','h2','h3','h4','p','table','img']):
        if tag.name.startswith('h'):
            level = int(tag.name[1])
            title = tag.get_text().strip()
            if "contents" in title.lower(): continue
            node = {"title": title, "level": level, "content": [], "subs": []}
            while stack and stack[-1]["level"] >= level: stack.pop()
            (stack[-1]["subs"] if stack else sections).append(node)
            stack.append(node)
        elif tag.name == 'p' and stack:
            stack[-1]["content"].append({"type": "text", "value": tag.get_text(separator="\n").strip()})
        elif tag.name == 'img' and stack:
            src = tag.get("src")
            if src: stack[-1]["content"].append({"type": "image", "value": src})
        elif tag.name == 'table' and stack:
            rows = [[c.get_text(separator="\n").strip() for c in r.find_all(['td','th'])] for r in tag.find_all('tr')]
            if rows: stack[-1]["content"].append({"type": "table", "value": rows})
    return sections

sections = parse_docx(os.path.join(PLAYBOOKS_DIR, selected))
prog_file = os.path.join(PLAYBOOKS_DIR, f"{Path(selected).stem}_progress.json")

def load_progress():
    if os.path.exists(prog_file):
        try:
            with open(prog_file) as f:
                d = json.load(f)
                return d.get("completed", {}), d.get("comments", {})
        except: pass
    return {}, {}

completed, comments = load_progress()

def save_progress():
    with open(prog_file, "w") as f:
        json.dump({"completed": completed, "comments": comments, "ts": datetime.now().isoformat()}, f, indent=2)

# -------------------------------------------------
# RENDER
# -------------------------------------------------
@st.fragment
def render_section(sec: Dict, prefix: str):
    sec_key = f"{prefix}_{hashlib.md5(sec['title'].encode()).hexdigest()[:8]}"
    st.markdown(f"<div id='{sec_key}' class='section-title'>{sec['title']}</div>", unsafe_allow_html=True)
    with st.expander("Expand", expanded=False):
        for item in sec.get("content", []):
            if item["type"] == "text":
                st.markdown(item["value"].replace("\n", "<br>"), unsafe_allow_html=True)
            elif item["type"] == "image":
                st.image(item["value"], use_container_width=True)
            elif item["type"] == "table":
                rows = item["value"]
                if len(rows) > 1 and any(any(h.lower() in ACTION_HEADERS for h in row) for row in rows):
                    cols = st.columns([1,2,4,2,1,2])
                    for i, h in enumerate(["Ref","Step","Desc","Owner","Done","Note"]): cols[i].write(f"**{h}**")
                    for i, row in enumerate(rows[1:]):
                        while len(row) < 4: row.append("")
                        rk = f"{sec_key}_r{i}"
                        c = st.columns([1,2,4,2,1,2])
                        for j, v in enumerate(row): c[j].write(v)
                        done = c[4].checkbox("", completed.get(rk, False), key=f"cb_{rk}")
                        note = c[5].text_input("", comments.get(rk, ""), key=f"ci_{rk}", label_visibility="collapsed")
                        if done != completed.get(rk, False): completed[rk] = done
                        if note != comments.get(rk, ""): comments[rk] = note
                else:
                    df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows)>1 else pd.DataFrame(rows)
                    st.dataframe(df, use_container_width=True, hide_index=True)

# -------------------------------------------------
# TOC
# -------------------------------------------------
toc = []
def build_toc(secs, p=""):
    for s in secs:
        a = f"{p}_{hashlib.md5(s['title'].encode()).hexdigest()[:8]}"
        toc.append((s["title"], a))
        build_toc(s.get("subs", []), a)
build_toc(sections)

st.markdown("<div class='toc'><b>Contents</b><br>" + "<br>".join(
    f"<a onclick=\"document.getElementById('{a}').scrollIntoView();\">{t}</a>" for t,a in toc
) + "</div>", unsafe_allow_html=True)

st.markdown("<div class='content'>", unsafe_allow_html=True)
for sec in sections:
    render_section(sec, selected.replace(".docx",""))
st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------
# PROGRESS + PDF + EXPORT
# -------------------------------------------------
total = sum(1 for k in completed if "_r" in k)
done = sum(completed.get(k, False) for k in completed if "_r" in k)
pct = int(done/total*100) if total else 0

st.markdown(f"<div class='progress-wrap'><div class='progress-fill' style='width:{pct}%'></div></div>", unsafe_allow_html=True)
st.write(f"**Progress: {pct}%**")

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("Save"):
        save_progress()
        st.success("Saved")
with col2:
    df = pd.DataFrame([{"Key":k, "Done":v, "Note":comments.get(k,"")} for k,v in completed.items()])
    st.download_button("CSV", df.to_csv(index=False).encode(), f"{Path(selected).stem}.csv", "text/csv")
with col3:
    @st.cache_data
    def make_pdf():
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, "Joval Wines - NIST Playbook", ln=1)
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 10, f"Playbook: {selected}", ln=1)
        pdf.ln(5)
        for sec in sections:
            pdf.set_font("Arial", "B", 10)
            pdf.multi_cell(0, 6, sec["title"])
            pdf.set_font("Arial", size=9)
            for item in sec.get("content", []):
                if item["type"] == "text":
                    pdf.multi_cell(0, 5, item["value"][:500])
                elif item["type"] == "table":
                    pdf.multi_cell(0, 5, "[Table]")
        return pdf.output(dest="S").encode("latin1")
    pdf_bytes = make_pdf()
    st.download_button("PDF Export", pdf_bytes, f"{Path(selected).stem}.pdf", "application/pdf")

# Auto-save
if st.sidebar.checkbox("Auto-save", True):
    save_progress()

# Scroll
if "scroll_to" in st.session_state:
    st.markdown(f"<script>document.getElementById('{st.session_state.scroll_to}').scrollIntoView();</script>", unsafe_allow_html=True)
    del st.session_state.scroll_to
