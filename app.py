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

st.set_page_config(page_title="Joval Wines NIST Tracker", page_icon="wine", layout="wide", initial_sidebar_state="expanded")

# -------------------------------------------------
# MODERN UI (jovalwines.com.au style)
# -------------------------------------------------
st.markdown("""
<style>
    :root{
        --red:#800020;--bg:#fff;--text:#000;--muted:#777;--sec:rgba(0,0,0,.02);
        --header-bg:linear-gradient(180deg,#fff 0%,#f9f9f9 100%);
    }
    html,body,.stApp{background:var(--bg)!important;color:var(--text)!important;font-family:Helvetica,Arial,sans-serif;}
    .sticky-header{
        position:sticky;top:0;z-index:9999;display:flex;align-items:center;justify-content:space-between;
        padding:12px 20px;background:var(--header-bg);border-bottom:1px solid #eee;box-shadow:0 2px 8px rgba(0,0,0,.05);
    }
    .logo{max-height:70px;}
    .title{font-size:2.2rem;font-weight:700;color:var(--text);margin:0;}
    .nist{color:var(--red);font-size:1.8rem;font-weight:700;letter-spacing:1px;}
    .toc{
        position:fixed;left:12px;top:88px;bottom:92px;width:260px;background:#fff;padding:12px;
        border-radius:8px;overflow:auto;border:1px solid #eee;z-index:900;box-shadow:0 2px 10px rgba(0,0,0,.06);
    }
    .toc a{cursor:pointer;color:var(--text);display:block;padding:6px 4px;border-radius:4px;}
    .toc a:hover{background:rgba(128,0,32,.05);}
    .content{margin-left:284px;padding:0 24px 120px;}
    .section{background:var(--sec);padding:16px;border-radius:8px;margin-bottom:16px;border:1px solid #eee;}
    .section-title{font-size:1.6rem;font-weight:bold;color:var(--text);margin:0 0 8px 0;}
    .scaled-img{max-width:90%;border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,.15);margin:12px 0;}
    .progress-wrap{width:100%;height:10px;background:#eee;border-radius:999px;overflow:hidden;margin:12px 0;}
    .progress-fill{height:100%;background:linear-gradient(90deg,var(--red),#500010);transition:width .6s ease;}
    .stButton>button,.stDownloadButton>button{
        background:var(--red)!important;color:#fff!important;border:none!important;font-weight:700;padding:0.5rem 1rem;border-radius:6px;
    }
    .stSidebar .stButton>button{background:#000!important;color:#fff!important;}
    .search-result{background:var(--red);color:#fff;padding:8px 12px;border-radius:6px;margin:6px 0;cursor:pointer;font-size:0.95rem;}
    .search-result:hover{background:#a00030;}
    .no-playbook-message{text-align:center;font-size:1.4rem;color:var(--red);margin:80px 0;font-weight:bold;}
    .instructional-text{
        background:rgba(217,83,79,.1);border:2px solid #d9534f;padding:16px;border-radius:8px;
        text-align:center;font-weight:bold;color:#d9534f;font-size:1.3rem;margin:20px 0;
    }
    @media(max-width:768px){
        .toc{display:none;}.content{margin-left:0;padding:0 12px;}
        .sticky-header{flex-direction:column;padding:10px;text-align:center;}
        .title{font-size:1.8rem;}.nist{font-size:1.5rem;}
    }
</style>
<script src="https://cdn.tailwindcss.com"></script>
""", unsafe_allow_html=True)

# -------------------------------------------------
# USER AUTH
# -------------------------------------------------
def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE) as f:
                return json.load(f)
        except:
            pass
    admin_hash = st.secrets.get("ADMIN_PASSWORD_HASH")
    if not admin_hash:
        st.error("Set ADMIN_PASSWORD_HASH in secrets.toml")
        st.stop()
    users = {"admin@joval.com": {"role": "admin", "hash": admin_hash}}
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)
    return users

def authenticate():
    if st.session_state.get("auth"):
        return st.session_state.user
    st.markdown("<div class='login-container'>", unsafe_allow_html=True)
    st.title("Joval Wines NIST Tracker")
    with st.form("login_form"):
        email = st.text_input("Email", placeholder="admin@joval.com")
        pwd = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            users = load_users()
            h = hashlib.sha256(pwd.encode()).hexdigest()
            user_data = users.get(email)
            if user_data and user_data.get("hash") == h:
                st.session_state.auth = True
                st.session_state.user = {"email": email, "role": user_data["role"]}
                st.rerun()
            st.error("Invalid credentials")
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

user = authenticate()
st.sidebar.success(f"**{user['email']}** – {user['role'].title()}")

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
    st.error(f"No .docx files in `{PLAYBOOKS_DIR}` folder.")
    st.stop()

# -------------------------------------------------
# SEARCH
# -------------------------------------------------
@st.cache_data(ttl=600)
def search_playbooks(query: str, top_k: int = 12):
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
                    snippet = "…" + txt[max(0,pos-40):pos+80].replace("\n"," ") + "…"
                    anchor = f"{pb}__{hashlib.md5(tag.get_text().encode()).hexdigest()[:8]}"
                    hits.append({"playbook": pb, "title": tag.get_text(strip=True), "anchor": anchor, "snippet": snippet})
        except:
            continue
    hits.sort(key=lambda x: len(x["title"]))
    return hits[:top_k]

query = st.sidebar.text_input("Search Playbooks", placeholder="e.g. detection")
if query.strip():
    results = search_playbooks(query.strip())
    if results:
        st.sidebar.markdown("**Results**")
        for r in results:
            if st.sidebar.button(f"{r['title']}", key=f"res_{r['playbook']}_{r['anchor']}"):
                st.session_state.selected = r["playbook"]
                st.session_state.scroll_to = r["anchor"]
                st.rerun()
    else:
        st.sidebar.info("No matches")
else:
    st.sidebar.markdown("<em>Enter keyword to search all playbooks</em>", unsafe_allow_html=True)

# -------------------------------------------------
# PLAYBOOK SELECTOR
# -------------------------------------------------
selected = st.selectbox("Select Playbook", [""] + playbooks, key="selected")
if not selected:
    st.markdown('<div class="no-playbook-message">Please select a playbook to begin.</div>', unsafe_allow_html=True)
    st.stop()

st.markdown('<div class="instructional-text">Complete each step in the NIST "Incident Handling Categories" section</div>', unsafe_allow_html=True)

# -------------------------------------------------
# PARSE DOCX
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
            if not title or "table of contents" in title.lower():
                continue
            node = {"title": title, "level": level, "content": [], "subs": []}
            while stack and stack[-1]["level"] >= level:
                stack.pop()
            (stack[-1]["subs"] if stack else sections).append(node)
            stack.append(node)
        elif tag.name == 'p' and stack:
            stack[-1]["content"].append({"type": "text", "value": tag.get_text(separator="\n").strip()})
        elif tag.name == 'img' and stack:
            src = tag.get("src", "")
            if src:
                stack[-1]["content"].append({"type": "image", "value": src})
        elif tag.name == 'table' and stack:
            rows = [[c.get_text(separator="\n").strip() for c in r.find_all(['td','th'])] for r in tag.find_all('tr')]
            if rows:
                stack[-1]["content"].append({"type": "table", "value": rows})
    return sections

sections = parse_docx(os.path.join(PLAYBOOKS_DIR, selected))

# -------------------------------------------------
# PROGRESS
# -------------------------------------------------
prog_file = os.path.join(PLAYBOOKS_DIR, f"{Path(selected).stem}_progress.json")
def load_progress():
    if os.path.exists(prog_file):
        try:
            with open(prog_file) as f:
                data = json.load(f)
                return data.get("completed", {}), data.get("comments", {})
        except:
            pass
    return {}, {}
completed, comments = load_progress()

def save_progress():
    with open(prog_file, "w") as f:
        json.dump({"completed": completed, "comments": comments, "ts": datetime.now().isoformat()}, f, indent=2)

# -------------------------------------------------
# RENDER SECTION
# -------------------------------------------------
@st.fragment
def render_section(sec: Dict, key_prefix: str):
    sec_key = f"{key_prefix}_{hashlib.md5(sec['title'].encode()).hexdigest()[:8]}"
    st.markdown(f"<div id='{sec_key}' class='section-title'>{sec['title']}</div>", unsafe_allow_html=True)
    with st.expander("Expand section", expanded=False):
        for item in sec.get("content", []):
            if item["type"] == "text":
                st.markdown(item["value"].replace("\n", "<br>"), unsafe_allow_html=True)
            elif item["type"] == "image":
                try:
                    st.image(item["value"], use_container_width=True)
                except:
                    st.write("[Image not available]")
            elif item["type"] == "table":
                rows = item["value"]
                if len(rows) > 1 and any(any(h.lower() in ACTION_HEADERS for h in row) for row in rows):
                    # Action table
                    cols = st.columns([1, 2, 4, 2, 1, 2])
                    headers = ["Ref", "Step", "Desc", "Owner", "Done", "Note"]
                    for i, h in enumerate(headers): cols[i].write(f"**{h}**")
                    for r_idx, row in enumerate(rows[1:]):
                        while len(row) < 4: row.append("")
                        rk = f"{sec_key}_r{r_idx}"
                        c0,c1,c2,c3,c4,c5 = st.columns([1,2,4,2,1,2])
                        c0.write(row[0]); c1.write(row[1]); c2.write(row[2]); c3.write(row[3])
                        done = c4.checkbox("", completed.get(rk, False), key=f"cb_{rk}")
                        note = c5.text_input("", comments.get(rk, ""), key=f"ci_{rk}", label_visibility="collapsed")
                        if done != completed.get(rk, False): completed[rk] = done
                        if note != comments.get(rk, ""): comments[rk] = note
                else:
                    df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
                    st.dataframe(df, use_container_width=True, hide_index=True)

# -------------------------------------------------
# TOC + CONTENT
# -------------------------------------------------
toc = []
def build_toc(secs, prefix=""):
    for s in secs:
        anchor = f"{prefix}_{hashlib.md5(s['title'].encode()).hexdigest()[:8]}"
        toc.append((s["title"], anchor))
        build_toc(s.get("subs", []), anchor)
build_toc(sections)

toc_html = "<div class='toc'><b>Contents</b><br>" + "<br>".join(
    f"<a onclick=\"document.getElementById('{a}').scrollIntoView({{behavior:'smooth'}});\">{t}</a>"
    for t,a in toc) + "</div>"
st.markdown(toc_html, unsafe_allow_html=True)

st.markdown("<div class='content'>", unsafe_allow_html=True)
for sec in sections:
    render_section(sec, selected.replace(".docx", ""))
st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------
# PROGRESS + ACTIONS
# -------------------------------------------------
total = sum(1 for k in completed if k.endswith("_r"))
done = sum(completed.get(k, False) for k in completed if k.endswith("_r"))
pct = int(done / total * 100) if total else 0

st.markdown(f"<div class='progress-wrap'><div class='progress-fill' style='width:{pct}%'></div></div>", unsafe_allow_html=True)
st.write(f"**Progress: {pct}%**")

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("Save Progress"):
        save_progress()
        st.success("Saved")
with col2:
    df = pd.DataFrame([
        {"Key": k, "Done": completed.get(k, False), "Note": comments.get(k, "")}
        for k in completed
    ])
    csv = df.to_csv(index=False).encode()
    st.download_button("Download CSV", csv, f"{Path(selected).stem}_progress.csv", "text/csv")
with col3:
    if st.button("Create Jira Ticket"):
        with st.form("jira_form"):
            summary = st.text_input("Summary", f"NIST Progress - {selected}")
            desc = st.text_area("Description", f"Progress: {pct}%")
            if st.form_submit_button("Send"):
                st.success("Jira ticket would be created here.")

# Auto-save
if st.sidebar.checkbox("Auto-save progress", True):
    save_progress()

# Scroll to anchor from search
if "scroll_to" in st.session_state:
    st.markdown(f"<script>document.getElementById('{st.session_state.scroll_to}').scrollIntoView();</script>", unsafe_allow_html=True)
    del st.session_state.scroll_to
