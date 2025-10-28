# --------------------------------------------------------------
# app.py – Joval Wines NIST Playbook Tracker (FULL VERSION)
# --------------------------------------------------------------
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

try:
    from sklearn.feature_extraction.text import TfidfVectorizer
    SKLEARN_AVAILABLE = True
except Exception:
    SKLEARN_AVAILABLE = False

import logging

# ----------------------------------------------------------------------
# Configuration & Paths
# ----------------------------------------------------------------------
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

# ----------------------------------------------------------------------
# Streamlit Page Config
# ----------------------------------------------------------------------
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

# ----------------------------------------------------------------------
# Hide Streamlit UI clutter
# ----------------------------------------------------------------------
st.markdown("""
<style>
/* Footer, toolbar, share button, deploy button */
.stApp > footer {display:none !important;}
[data-testid="stToolbar"] {display:none !important;}
[data-testid="collapsedControl"] {display:none !important;}
.stDeployButton {display:none !important;}
[data-testid="stDecoration"] {display:none !important;}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# CSP + JavaScript for scroll / playbook switch / highlight
# ----------------------------------------------------------------------
st.markdown("""
<meta http-equiv="Content-Security-Policy"
      content="default-src 'self'; script-src 'self' 'unsafe-inline' https://cdn.tailwindcss.com; style-src 'self' 'unsafe-inline'; img-src 'self' data:;">
<script src="https://cdn.tailwindcss.com"></script>
<script>
function switchAndScroll(playbook, anchor, query='') {
    const url = new URL(window.location);
    url.searchParams.set('pb', playbook);
    url.searchParams.set('sec', anchor);
    if (query) url.searchParams.set('q', query);
    history.replaceState(null, '', url);

    // hidden button forces Streamlit rerun
    const btn = document.createElement('button');
    btn.id = 'hidden_rerun_btn';
    btn.style.display = 'none';
    document.body.appendChild(btn);
    btn.click();

    // store target for post-rerun
    localStorage.setItem('scroll_target', anchor);
    localStorage.setItem('highlight_query', query);
}
window.addEventListener('load', () => {
    const target = localStorage.getItem('scroll_target');
    const query   = localStorage.getItem('highlight_query');
    if (target) {
        setTimeout(() => {
            const el = document.getElementById(target);
            if (el) {
                el.scrollIntoView({behavior:'smooth', block:'start'});
                if (query) {
                    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
                    const nodes = []; let node;
                    while (node = walker.nextNode()) nodes.push(node);
                    const regex = new RegExp(query, 'gi');
                    nodes.forEach(n => {
                        if (regex.test(n.textContent)) {
                            const span = document.createElement('span');
                            span.style.backgroundColor = '#fff3cd';
                            span.innerHTML = n.textContent.replace(regex, '<mark>$&</mark>');
                            n.parentNode.replaceChild(span, n);
                        }
                    });
                }
                el.style.transition = 'background 0.5s';
                el.style.backgroundColor = '#fff3cd';
                setTimeout(() => el.style.backgroundColor = '', 3000);
            }
            localStorage.removeItem('scroll_target');
            localStorage.removeItem('highlight_query');
        }, 800);
    }
});
</script>
<style>
:root{--bg:#fff;--text:#000;--muted:#777;--joval-red:#800020;--section-bg:rgba(0,0,0,0.02);}
html,body,.stApp{background:var(--bg)!important;color:var(--text)!important;font-family:'Helvetica','Arial',sans-serif!important;}
.sticky-header{position:sticky;top:0;z-index:9999;display:flex;align-items:center;justify-content:space-between;padding:10px 18px;background:linear-gradient(180deg,rgba(255,255,255,0.96),rgba(245,245,245,0.86));border-bottom:1px solid rgba(0,0,0,0.04);width:100%;box-sizing:border-box;}
.logo-left{flex-shrink:0;height:80px;margin-right:auto;}
.app-title{flex:1;text-align:center;font-family:'Helvetica',sans-serif;font-size:3rem;color:var(--text);font-weight:700;margin:0;text-shadow:2px 2px 4px rgba(0,0,0,0.3);}
.nist-logo{flex-shrink:0;font-family:'Helvetica',sans-serif;font-size:2rem;color:var(--joval-red);letter-spacing:0.08rem;text-shadow:0 0 12px rgba(128,0,32,0.08);font-weight:700;margin-left:auto;}
.toc{position:fixed;left:12px;top:84px;bottom:92px;width:260px;background:rgba(255,255,255,0.95);padding:10px;border-radius:8px;overflow:auto;border:1px solid rgba(0,0,0,0.03);z-index:900;}
.content-wrap{margin-left:284px;margin-right:24px;padding-top:0;padding-bottom:100px;}
.section-box{background:var(--section-bg);padding:12px;border-radius:8px;margin-bottom:12px;border:1px solid rgba(0,0,0,0.02);}
.scaled-img{max-width:90%;height:auto;border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,0.6);margin:12px 0;display:block;}
.playbook-table{border-collapse:collapse;width:100%;margin-top:8px;margin-bottom:8px;}
.playbook-table th,.playbook-table td{border:1px solid rgba(0,0,0,0.05);padding:8px;color:var(--text);text-align:left;vertical-align:top;}
.playbook-table th{background:rgba(0,0,0,0.02);color:#333;font-weight:700;}
.stCheckbox>label,.stCheckbox>label>div{color:var(--muted)!important;}
.stTextInput>label,.stSelectbox>label{color:var(--text)!important;font-weight:bold;font-size:1.1rem;}
.stSelectbox{margin-bottom:0!important;}
.bottom-toolbar{position:fixed;left:0;right:0;bottom:0;z-index:999;background:linear-gradient(180deg,rgba(255,255,255,0.95),rgba(245,245,245,0.9));border-top:1px solid rgba(0,0,0,0.03);padding:10px 18px;display:flex;justify-content:space-between;align-items:center;}
.progress-wrap{width:360px;height:10px;background:rgba(0,0,0,0.03);border-radius:999px;overflow:hidden;}
.progress-fill{height:100%;width:0%;background:linear-gradient(90deg,var(--joval-red),#500010);box-shadow:0 0 12px rgba(128,0,32,0.2);transition:width .6s ease;}
.stButton>button,.stDownloadButton>button{background:#000!important;color:#fff!important;border:1px solid rgba(0,0,0,0.12);font-weight:700;}
.stSidebar .stButton>button{color:#fff!important;background:var(--joval-red)!important;font-size:1.1rem;font-weight:bold;}
.comments-title{color:var(--text);font-weight:700;margin-top:12px;margin-bottom:6px;}
a{color:var(--joval-red);cursor:pointer;}
.toc a{display:block;padding:6px 4px;color:var(--text);text-decoration:none;cursor:pointer;}
.toc a:hover{background:rgba(0,0,0,0.02);border-radius:4px;}
.search-result a{color:#800020;text-decoration:underline;cursor:pointer;}
.search-result a:hover{color:#a00030;}
.section-title{font-size:1.8rem!important;font-weight:bold!important;color:var(--text)!important;margin-top:0!important;margin-bottom:4px!important;}
.instructional-text{color:#d9534f!important;font-size:1.5rem!important;font-weight:bold!important;border:2px solid #d9534f;padding:15px;border-radius:8px;background:rgba(217,83,79,0.1);margin:20px 0!important;text-align:center;line-height:1.4!important;}
.content-text{font-size:1.1rem!important;}
[data-testid="stExpander"]>div:first-child{background:#f0f0f0!important;color:#000!important;padding:12px 20px!important;border-radius:8px!important;font-size:1.3rem!important;font-weight:bold!important;border:1px solid rgba(0,0,0,0.12)!important;margin-bottom:10px!important;cursor:pointer!important;transition:all .2s ease!important;}
[data-testid="stExpander"]>div:first-child:hover{background:#e0e0e0!important;}
.playbook-select-label{font-size:2.5rem!important;font-weight:bold!important;color:var(--text)!important;}
.nist-incident-section{color:#d9534f!important;}
@media(max-width:768px){
  .content-wrap{margin-left:0!important;margin-right:0!important;padding:0 10px;}
  .toc{display:none!important;}
  .sticky-header{flex-direction:column;padding:10px;}
  .app-title{font-size:2rem;margin:10px 0;}
  .section-title{font-size:1.4rem!important;}
  .bottom-toolbar{flex-direction:column;gap:10px;padding:10px;}
  .progress-wrap{width:100%;}
}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# USER MANAGEMENT (admin password from secrets only)
# ----------------------------------------------------------------------
def load_users() -> Dict[str, Dict]:
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                content = f.read().strip()
                if content:
                    return {k.lower(): v for k, v in json.loads(content).items()}
        except Exception:
            pass
    # default admin – password MUST be set in secrets.toml
    default_admin = {
        "admin@joval.com": {
            "role": "admin",
            "hash": hashlib.sha256(st.secrets["ADMIN_PASSWORD"].encode()).hexdigest()
        }
    }
    save_users(default_admin)
    return default_admin

def save_users(users: Dict):
    with open(USERS_FILE, "w") as f:
        json.dump({k.lower(): v for k, v in users.items()}, f, indent=2)

def get_user_role(email: str) -> str:
    return load_users().get(email.lower(), {}).get("role", "user")

def create_user(email: str, role: str, password: str):
    users = load_users()
    email = email.lower()
    if email in users:
        return False, "User already exists."
    users[email] = {"role": role, "hash": hashlib.sha256(password.encode()).hexdigest()}
    save_users(users)
    logging.info(f"User created: {email}, Role: {role}")
    return True, "User created."

def reset_user_password(email: str, password: str):
    users = load_users()
    email = email.lower()
    if email not in users:
        return False, "User not found."
    users[email]["hash"] = hashlib.sha256(password.encode()).hexdigest()
    save_users(users)
    logging.info(f"Password reset: {email}")
    return True, password

def delete_user(email: str):
    users = load_users()
    email = email.lower()
    if email in users:
        del users[email]
        save_users(users)
        logging.info(f"User deleted: {email}")
        return True, "User deleted."
    return False, "User not found."

# ----------------------------------------------------------------------
# AUTHENTICATION
# ----------------------------------------------------------------------
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
            email = username if "@" in username else f"{username}@joval.com"
            email = email.lower()
            if email in users and hashlib.sha256(password.encode()).hexdigest() == users[email]["hash"]:
                st.session_state.authenticated = True
                display = username.split("@")[0].title() if "@" in username else username.title()
                st.session_state.user = {"email": email, "name": display, "role": users[email]["role"]}
                st.session_state.login_attempts = 0
                st.success("Logged in!")
                st.rerun()
            else:
                st.session_state.login_attempts += 1
                st.session_state.last_attempt = now
                st.error("Invalid credentials.")
        st.stop()

    if st.sidebar.button("Logout"):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.session_state.authenticated = False
        st.rerun()

    return st.session_state.user

# ----------------------------------------------------------------------
# Helper utilities
# ----------------------------------------------------------------------
def stable_key(playbook_name: str, title: str, level: int) -> str:
    base = f"{playbook_name}||{level}||{title}"
    return "sec_" + hashlib.md5(base.encode()).hexdigest()

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
            pass
    return {}, {}

def save_progress(playbook_name: str, completed: dict, comments: dict) -> str:
    rec = {
        "playbook": playbook_name,
        "timestamp": datetime.now().isoformat(),
        "version": "1.0",
        "completed": completed,
        "comments": comments,
    }
    path = progress_filepath(playbook_name)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(rec, fh, indent=2)
    return path

def get_logo() -> str:
    if "logo_b64" not in st.session_state:
        st.session_state.logo_b64 = None
    if st.session_state.logo_b64:
        return f'<img src="data:image/png;base64,{st.session_state.logo_b64}" class="logo-left" alt="Logo" style="height:80px;"/>'
    default = "logo.png"
    if os.path.exists(default):
        with open(default, "rb") as f:
            return f'<img src="data:image/png;base64,{base64.b64encode(f.read()).decode()}" class="logo-left" alt="Logo" style="height:80px;"/>'
    return '<div class="logo-left"></div>'

# ----------------------------------------------------------------------
# Playbook parsing (cached per file)
# ----------------------------------------------------------------------
@st.cache_data(hash_funcs={Path: lambda p: str(p)})
def parse_playbook_cached(path: str) -> List[Dict[str, Any]]:
    with open(path, "rb") as fh:
        result = mammoth.convert_to_html(fh)
        html = result.value
    soup = BeautifulSoup(html, "html.parser")

    exclude = ["table of contents", "document control", "document revision", "assumptions", "disclaimer"]
    def excluded(txt: str) -> bool:
        return txt and any(e in txt.strip().lower() for e in exclude)

    sections = []
    stack = []
    for tag in soup.find_all(['h1','h2','h3','h4','p','table','img']):
        if tag.name.startswith('h') and tag.name[1:].isdigit():
            title = tag.get_text().strip()
            if excluded(title): continue
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
            txt = tag.get_text(separator="\n").strip()
            if txt and stack:
                stack[-1]["content"].append({"type": "text", "value": txt})
        elif tag.name == 'img':
            src = tag.get("src", "")
            if src and stack:
                stack[-1]["content"].append({"type": "image", "value": src})
        elif tag.name == 'table':
            rows = [[td.get_text(separator="\n").strip() for td in tr.find_all(["td","th"])] for tr in tag.find_all("tr")]
            if rows and stack:
                stack[-1]["content"].append({"type": "table", "value": rows})

    # ----- reconstruct action tables -----
    header_keywords = ["reference","step","description","ownership","responsibility"]
    def rebuild(section):
        i = 0
        new = []
        while i < len(section["content"]):
            it = section["content"][i]
            if it["type"] != "text":
                new.append(it); i += 1; continue
            txt = it["value"].lower()
            if sum(w in txt for w in header_keywords) >= 2 or ref_pattern.match(it["value"]):
                headers = ["Reference","Step","Description","Ownership/Responsibility"]
                rows = []; ref = step = ""; desc = []
                j = i + 1
                while j < len(section["content"]) and section["content"][j]["type"] == "text":
                    line = section["content"][j]["value"].strip()
                    if ref_pattern.match(line):
                        if ref:
                            rows.append([ref, step, " ".join(desc), ""])
                        m = ref_pattern.match(line)
                        ref, step = m.group(0), line[m.end():].strip()
                        desc = []
                    else:
                        desc.append(line)
                    j += 1
                if ref:
                    rows.append([ref, step, " ".join(desc), ""])
                if rows:
                    new.append({"type":"table","value":[headers]+rows})
                i = j
            else:
                new.append(it); i += 1
        section["content"] = new
    def walk(nodes):
        for n in nodes:
            rebuild(n)
            if n.get("subs"): walk(n["subs"])
    walk(sections)

    def prune(node):
        node["subs"] = [s for s in node.get("subs", []) if prune(s)]
        return bool(node.get("content")) or bool(node.get("subs"))
    return [s for s in sections if prune(s)]

# ----------------------------------------------------------------------
# Search across ALL playbooks (cached)
# ----------------------------------------------------------------------
@st.cache_data(ttl=600)
def run_search_assistant(query: str, playbooks_list: List[str], top_k: int = 7):
    corpus = []
    for pb in playbooks_list:
        secs = parse_playbook_cached(os.path.join(PLAYBOOKS_DIR, pb))
        for s in secs:
            parts = [s["title"]]
            for c in s.get("content", []):
                if c["type"] == "text":
                    parts.append(c["value"])
                elif c["type"] == "table":
                    for r in c["value"]:
                        parts.append(" ".join(r))
            corpus.append({
                "playbook": pb,
                "title": s["title"],
                "level": s["level"],
                "text": "\n".join(parts),
                "anchor": stable_key(pb, s["title"], s["level"])
            })
    if not corpus or not SKLEARN_AVAILABLE:
        return []
    vect = TfidfVectorizer(stop_words='english', max_features=20000)
    mat = vect.fit_transform([c["text"] for c in corpus])
    qv = vect.transform([query])
    sims = (mat @ qv.T).toarray().ravel()
    idxs = sims.argsort()[::-1][:top_k]
    return [corpus[i] for i in idxs if sims[i] > 0.05]

# ----------------------------------------------------------------------
# Rendering helpers (FIXED container usage)
# ----------------------------------------------------------------------
def render_action_table(pb: str, sec_key: str, rows: List[List[str]],
                       completed: dict, comments: dict, autosave: bool, tbl_idx: int):
    default_headers = ["Reference","Step","Description","Ownership/Responsibility"]
    headers = rows[0] if rows and not ref_pattern.match(rows[0][0] if rows[0] else "") else default_headers
    data = rows[1:] if len(rows) > 1 else rows
    for r in data:
        while len(r) < 4: r.append("")
    st.caption("Mark tasks complete and add notes.")
    cols = st.columns([1,2,4,2,1,2])
    for i, h in enumerate(["Ref","Step","Desc","Owner","Done","Comment"]):
        cols[i].write(h)
    changed = False
    tbl_key = f"{sec_key}::tbl::{tbl_idx}"
    for ridx, row in enumerate(data):
        row_key = f"{tbl_key}::row::{ridx}"
        c_key = f"{row_key}::c"
        prev_done = completed.get(row_key, False)
        prev_comm = comments.get(c_key, "")
        c = st.columns([1,2,4,2,1,2])
        for i in range(4): c[i].write(row[i])
        new_done = c[4].checkbox("", value=prev_done, key=f"cb_{row_key}")
        new_comm = c[5].text_input("", value=prev_comm, key=f"ci_{c_key}", label_visibility="collapsed")
        if new_done != prev_done:
            completed[row_key] = new_done
            changed = True
        if new_comm != prev_comm:
            comments[c_key] = new_comm
            changed = True
    if autosave and changed:
        save_progress(pb, completed, comments)

def render_section(sec: Dict, pb: str, completed: dict, comments: dict, autosave: bool):
    key = stable_key(pb, sec["title"], sec["level"])
    cls = "nist-incident-section" if sec["title"] == "NIST Incident Handling Categories" else ""
    st.markdown(f"<div class='section-title {cls}' id='{key}'>{sec['title']}</div>", unsafe_allow_html=True)

    # FIXED: use a container inside the expander so all widgets are legal
    with st.expander("Expand section", expanded=False):
        container = st.container()
        with container:
            tbl_idx = 0
            for it in sec.get("content", []):
                if it["type"] == "text":
                    st.markdown(f"<div class='content-text'>{it['value'].replace(chr(10),'<br/>')}</div>", unsafe_allow_html=True)
                elif it["type"] == "image":
                    try: st.image(it["value"])
                    except: pass
                elif it["type"] == "table":
                    rows = it["value"]
                    if rows and len(rows[0]) >= 4 and any("ref" in h.lower() for h in rows[0]):
                        render_action_table(pb, key, rows, completed, comments, autosave, tbl_idx)
                        tbl_idx += 1
                    else:
                        df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
                        st.dataframe(df, use_container_width=True, hide_index=True)
            # sub-sections (recursive)
            for sub in sec.get("subs", []):
                render_section(sub, pb, completed, comments, autosave)
            # section-level comment
            sec_comm_key = f"{key}::sec_comment"
            prev = comments.get(sec_comm_key, "")
            new = st.text_area("", value=prev, key=f"sec_c_{key}", height=100, label_visibility="collapsed")
            if new != prev:
                comments[sec_comm_key] = new
                if autosave: save_progress(pb, completed, comments)

# ----------------------------------------------------------------------
# Export helpers
# ----------------------------------------------------------------------
@st.cache_data
def export_csv(completed: dict, comments: dict, pb: str) -> bytes:
    df = pd.DataFrame({
        "Key": list(completed.keys()) + list(comments.keys()),
        "Status": [str(completed.get(k, '')) for k in completed.keys()] + [''] * len(comments),
        "Comment": [''] * len(completed) + [str(v) for v in comments.values()]
    })
    return df.to_csv(index=False).encode('utf-8')

@st.cache_data
def export_excel(completed: dict, comments: dict, pb: str) -> bytes:
    if not OPENPYXL_AVAILABLE:
        return b""
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        pd.DataFrame(list(completed.items()), columns=["Key","Done"]).to_excel(writer, sheet_name="Progress", index=False)
        pd.DataFrame(list(comments.items()), columns=["Key","Comment"]).to_excel(writer, sheet_name="Comments", index=False)
    return out.getvalue()

@st.cache_data(ttl=300)
def generate_pdf(sections: List[Dict], pb: str) -> bytes:
    try:
        pdf = FPDF()
        pdf.set_margins(15,10,15)
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(0,10,txt="Joval Wines - NIST Playbook Tracker", ln=1, align='C')
        pdf.set_font("Arial", size=10)
        pdf.cell(0,8,txt=f"Playbook: {pb}", ln=1)
        pdf.ln(4)
        def add(sec, indent=0):
            pdf.set_font("Arial","B",8)
            pdf.multi_cell(180,6,"  "*indent + sec["title"],0,'L')
            pdf.ln(2)
            pdf.set_font("Arial","",7)
            for it in sec.get("content",[]):
                if it["type"]=="text":
                    for line in it["value"].split("\n"):
                        pdf.multi_cell(180,5,line.strip(),0,'L')
                        pdf.ln(1)
                elif it["type"]=="table":
                    pdf.multi_cell(180,5,"[Table]",0,'L')
                    pdf.ln(1)
            for sub in sec.get("subs",[]):
                add(sub, indent+2)
            pdf.ln(2)
        for s in sections:
            add(s)
        return pdf.output(dest='S').encode('latin1')
    except Exception as e:
        st.error(f"PDF error: {e}")
        return b""

# ----------------------------------------------------------------------
# Admin Dashboard
# ----------------------------------------------------------------------
def admin_dashboard(user):
    if get_user_role(user["email"]) != "admin":
        st.error("Admin only.")
        return
    st.title("Admin Dashboard")
    t1,t2,t3,t4,t5 = st.tabs(["Create User","Reset Password","List Users","Delete User","Upload"])
    with t1:
        st.subheader("Create New User")
        email = st.text_input("Email")
        role = st.selectbox("Role",["user","admin"])
        gen = st.checkbox("Generate random password", value=True)
        if gen:
            pwd = secrets.token_urlsafe(16)
            st.markdown(f"**Generated:** `{pwd}` (share securely)")
        else:
            pwd = st.text_input("Password", type="password")
        if st.button("Create"):
            if email and pwd:
                ok, msg = create_user(email, role, pwd)
                st.write("Success" if ok else "Error", msg)
    with t2:
        st.subheader("Reset Password")
        email = st.text_input("User email")
        gen = st.checkbox("Generate random", value=True, key="rgen")
        if gen:
            pwd = secrets.token_urlsafe(16)
            st.markdown(f"**New:** `{pwd}`")
        else:
            pwd = st.text_input("New password", type="password")
        if st.button("Reset"):
            if email and pwd:
                ok, _ = reset_user_password(email, pwd)
                st.write("Success" if ok else "Error")
    with t3:
        st.subheader("All Users")
        st.table(pd.DataFrame([{"Email":k,"Role":v["role"]} for k,v in load_users().items()]))
    with t4:
        st.subheader("Delete User")
        email = st.text_input("Email to delete")
        if st.button("Delete"):
            ok, msg = delete_user(email)
            st.write("Success" if ok else "Error", msg)
    with t5:
        logo = st.file_uploader("Custom Logo", type=["png","jpg","jpeg"])
        if logo:
            st.session_state.logo_b64 = base64.b64encode(logo.read()).decode()
            st.success("Logo saved")
        pb_file = st.file_uploader("New Playbook (.docx)", type=["docx"])
        if pb_file:
            dest = os.path.join(PLAYBOOKS_DIR, pb_file.name)
            with open(dest, "wb") as f:
                f.write(pb_file.getbuffer())
            st.success("Playbook uploaded")
    if st.button("Back to App"):
        st.session_state.admin_page = False
        st.rerun()

# ----------------------------------------------------------------------
# MAIN APP
# ----------------------------------------------------------------------
def main():
    user = authenticate()
    st.sidebar.info(f"Logged in: **{user['name']}** – Role: {get_user_role(user['email'])}")

    if st.session_state.get('admin_page', False):
        admin_dashboard(user)
        return

    if get_user_role(user["email"]) == "admin":
        if st.sidebar.button("Admin Dashboard"):
            st.session_state.admin_page = True
            st.rerun()

    # ----- URL parameters -----
    params = st.query_params
    url_pb = params.get("pb", [None])[0]
    url_sec = params.get("sec", [None])[0]
    url_q = params.get("q", [None])[0]

    # ----- Playbooks list -----
    global playbooks
    playbooks = sorted([f for f in os.listdir(PLAYBOOKS_DIR) if f.lower().endswith(".docx")])
    if not playbooks:
        st.error("No playbooks found.")
        return

    # ----- Playbook selector (sync with URL) -----
    sel_idx = playbooks.index(url_pb) if url_pb in playbooks else 0
    selected = st.selectbox("Select playbook", playbooks, index=sel_idx, key="pb_select")
    if selected != url_pb:
        st.query_params["pb"] = selected

    # ----- Load selected playbook -----
    parsed_key = f"parsed::{selected}"
    if parsed_key not in st.session_state:
        st.session_state[parsed_key] = parse_playbook_cached(os.path.join(PLAYBOOKS_DIR, selected))
    sections = st.session_state[parsed_key]

    completed, comments = load_progress(selected)
    st.session_state[f"comp::{selected}"] = completed
    st.session_state[f"comm::{selected}"] = comments
    completed = st.session_state[f"comp::{selected}"]
    comments = st.session_state[f"comm::{selected}"]

    # ----- Header -----
    st.markdown(f"""
    <div class='sticky-header'>
        {get_logo()}
        <div class='app-title'>Joval Wines NIST Playbook Tracker</div>
        <div class='nist-logo'>NIST</div>
    </div>
    """, unsafe_allow_html=True)

    # ----- Sidebar search (all playbooks) -----
    st.sidebar.markdown("### Search All Playbooks")
    query = st.sidebar.text_input("Search", value=url_q or "", key="search_q")
    if st.sidebar.button("Search"):
        st.query_params["q"] = query

    if query:
        results = run_search_assistant(query, playbooks, 10)
        if results:
            st.sidebar.markdown("**Results**")
            for r in results:
                clean = r["playbook"].replace(".docx","").split(" v")[0]
                onclick = f"switchAndScroll('{r['playbook']}','{r['anchor']}','{query}')"
                st.sidebar.markdown(
                    f"<div class='search-result'>"
                    f"• **{clean}**<br>"
                    f"  <a onclick=\"{onclick}\">{r['title']}</a>"
                    f"</div>", unsafe_allow_html=True)
        else:
            st.sidebar.info("No matches.")

    # ----- Table of Contents -----
    toc_items = []
    def build_toc(secs):
        for s in secs:
            anchor = stable_key(selected, s["title"], s["level"])
            toc_items.append({"title": s["title"], "anchor": anchor})
            if s.get("subs"): build_toc(s["subs"])
    build_toc(sections)
    toc_html = "<div class='toc'><h4>Contents</h4>" + "".join(
        f"<a onclick=\"switchAndScroll('{selected}','{t['anchor']}')\">{t['title']}</a>"
        for t in toc_items) + "</div>"
    st.markdown(toc_html, unsafe_allow_html=True)

    # ----- Main content -----
    st.markdown('<div class="content-wrap">', unsafe_allow_html=True)
    st.markdown('<div class="instructional-text">In the event of a cyber incident select the required playbook and complete each required step in the NIST "Incident Handling Categories" section</div>', unsafe_allow_html=True)
    for sec in sections:
        render_section(sec, selected, completed, comments, autosave=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ----- Progress -----
    total = len([k for k in completed if "::row::" in k])
    done  = sum(completed.get(k, False) for k in completed if "::row::" in k)
    pct   = int(done / max(total, 1) * 100)
    st.progress(pct / 100)
    st.info(f"Progress: **{pct}%**")

    # ----- Export buttons -----
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Save Progress"):
            save_progress(selected, completed, comments)
            st.success("Saved")
        csv = export_csv(completed, comments, selected)
        st.download_button("CSV", csv, f"{os.path.splitext(selected)[0]}_progress.csv", "text/csv")
    with c2:
        if OPENPYXL_AVAILABLE:
            xls = export_excel(completed, comments, selected)
            st.download_button("Excel", xls, f"{os.path.splitext(selected)[0]}_progress.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        pdf = generate_pdf(sections, selected)
        if pdf:
            st.download_button("PDF", pdf, f"{os.path.splitext(selected)[0]}_export.pdf", "application/pdf")

    # ----- hidden rerun button for JS -----
    st.markdown("<button id='hidden_rerun_btn' style='display:none;'></button>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
