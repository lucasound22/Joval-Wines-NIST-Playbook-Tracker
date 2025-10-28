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

# === PAGE CONFIG & MODERN UI ===
st.set_page_config(
    page_title="Joval Wines NIST Playbook Tracker",
    page_icon="wine",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
/* Tailwind CDN */
@import url('https://cdn.tailwindcss.com');

/* Core Colors */
:root{
    --bg:#ffffff;
    --text:#111111;
    --muted:#666666;
    --red:#800020;
    --card-bg:#fafafa;
    --border:#eaeaea;
}

/* Global */
html,body,.stApp{background:var(--bg)!important;color:var(--text)!important;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;}
.stApp > footer,.stApp [data-testid="stToolbar"],.stApp [data-testid="collapsedControl"],.stDeployButton{display:none!important;}

/* Header - Fixed Layout */
.sticky-header{
    position:sticky;top:0;z-index:9999;
    display:flex;align-items:center;justify-content:space-between;
    padding:1.2rem 2rem;background:#fff;
    border-bottom:1px solid var(--border);box-shadow:0 2px 8px rgba(0,0,0,.05);
    min-height:100px;
}
.logo-left{height:140px;width:auto;}
.app-title{font-size:2.4rem;font-weight:700;color:var(--text);margin:0;text-align:center;flex:1;}
.nist-logo{font-size:2.2rem;color:var(--red);font-weight:700;height:140px;display:flex;align-items:center;}

/* Buttons - Compact & Clean */
.stButton>button,.stDownloadButton>button{
    background:#000!important;color:#fff!important;
    border-radius:8px;padding:0.6rem 1.2rem!important;font-weight:600;
    min-width:120px;text-align:center;
    box-shadow:0 2px 4px rgba(0,0,0,0.1);
}
.stButton>button:hover,.stDownloadButton>button:hover{opacity:.9;transform:translateY(-1px);}

/* Progress */
.progress-wrap{height:12px;background:#e5e5e5;border-radius:999px;overflow:hidden;margin:1rem 0;}
.progress-fill{height:100%;background:var(--red);transition:width .4s ease;}

/* Bottom Toolbar */
.bottom-toolbar{
    position:fixed;bottom:0;left:0;right:0;z-index:999;
    background:#fff;border-top:1px solid var(--border);
    padding:.75rem 2rem;display:flex;align-items:center;justify-content:space-between;
    box-shadow:0 -2px 8px rgba(0,0,0,.03);
}

/* Sidebar */
.css-1d391kg{padding-top:1rem;}
.sidebar-header{font-weight:600;font-size:1.1rem;margin-bottom:.5rem;}
.sidebar-subheader{font-weight:600;margin-top:1rem;margin-bottom:.5rem;}

/* Content */
.content-wrap{margin-left:280px;padding:2rem 2rem 6rem;}
.section-title{font-size:1.7rem;font-weight:700;margin-bottom:.75rem;color:var(--text);}
.nist-incident-section{color:var(--red)!important;}

/* Responsive */
@media (max-width:768px){
    .sticky-header{flex-direction:column;padding:1rem;min-height:auto;}
    .app-title{font-size:1.8rem;}
    .logo-left,.nist-logo{height:100px;}
    .content-wrap{margin-left:0;padding:1rem;}
}
</style>
""", unsafe_allow_html=True)

# === USER MANAGEMENT ===
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

def authenticate():
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
        
        if st.button("Login"):
            users = load_users()
            email = username if "@" in username else username + "@joval.com"
            email = email.lower()
            if email in users and hashlib.sha256(password.encode()).hexdigest() == users[email]["hash"]:
                st.session_state.authenticated = True
                display_name = username.split("@")[0].title() if "@" in username else username.title()
                st.session_state.user = {"email": email, "name": display_name, "role": users[email]["role"]}
                st.success("Logged in!")
                st.rerun()
            else:
                st.error("Invalid credentials.")
        st.stop()

    if st.sidebar.button("Logout"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

    return st.session_state.user

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

def reset_playbook_progress(playbook_name: str):
    """Fully reset progress for the playbook"""
    path = progress_filepath(playbook_name)
    if os.path.exists(path):
        os.remove(path)
    # Clear from session state
    for key in list(st.session_state.keys()):
        if key.startswith(f"completed::{playbook_name}") or key.startswith(f"comments::{playbook_name}"):
            del st.session_state[key]
    st.success(f"**{playbook_name}** has been reset.")
    st.rerun()

def get_logo():
    if "logo_b64" not in st.session_state:
        st.session_state.logo_b64 = None
    if st.session_state.logo_b64:
        return f'<img src="data:image/png;base64,{st.session_state.logo_b64}" class="logo-left" alt="Logo" />'
    default_logo_path = "logo.png"
    if os.path.exists(default_logo_path):
        with open(default_logo_path, "rb") as f:
            logo_bytes = f.read()
            return f'<img src="data:image/png;base64,{base64.b64encode(logo_bytes).decode()}" class="logo-left" alt="Logo" />'
    return '<div class="logo-left"></div>'

# === PLAYBOOK PARSING ===
@st.cache_data(hash_funcs={Path: lambda p: str(p)})
def parse_playbook_cached(path: str) -> List[Dict[str, Any]]:
    with open(path, "rb") as fh:
        result = mammoth.convert_to_html(fh)
        html = result.value
    soup = BeautifulSoup(html, "html.parser")

    exclude_terms = ["table of contents", "document control", "document revision", "assumptions", "disclaimer"]
    def excluded(text: str) -> bool:
        return text and any(ex in text.strip().lower() for ex in exclude_terms)

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

    # Reconstruct action tables
    def reconstruct_tables_in_section(section):
        contents = section.get("content", [])
        i = 0
        new_contents = []
        header_keywords = ["reference", "step", "description", "ownership", "responsibility"]
        while i < len(contents):
            item = contents[i]
            if item["type"] != "text":
                new_contents.append(item)
                i += 1
                continue
            txt = item["value"].strip()
            if ref_pattern.match(txt) or sum(1 for w in header_keywords if w in txt.lower()) >= 2:
                headers = ["Reference", "Step", "Description", "Ownership/Responsibility"]
                rows = []
                current_ref = current_step = ""
                current_desc_parts = []
                j = i if not ref_pattern.match(txt) else i
                while j < len(contents) and contents[j]["type"] == "text":
                    txt_j = contents[j]["value"].strip()
                    if ref_pattern.match(txt_j):
                        if current_ref:
                            desc = " ".join(current_desc_parts)
                            rows.append([current_ref, current_step, desc, ""])
                            current_desc_parts = []
                        match = ref_pattern.match(txt_j)
                        current_ref = match.group(0)
                        current_step = txt_j[match.end():].strip()
                    else:
                        current_desc_parts.append(txt_j)
                    j += 1
                if current_ref:
                    desc = " ".join(current_desc_parts)
                    rows.append([current_ref, current_step, desc, ""])
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

# === RENDERING (Fragmented for Speed) ===
ACTION_HEADERS = {"reference","ref","step","description","ownership","responsibility","owner","responsible"}

def is_action_table(rows: List[List[str]]) -> bool:
    if not rows: return False
    headers = [h.strip().lower() for h in rows[0]]
    return sum(1 for h in headers if any(k in h for k in ACTION_HEADERS)) >= 2

@st.fragment
def render_action_table(playbook_name, sec_key, rows, completed_map, comments_map, table_index=0):
    default_headers = ["Reference", "Step", "Description", "Ownership/Responsibility"]
    headers = rows[0] if len(rows) > 0 and not ref_pattern.match(rows[0][0].strip()) else default_headers
    data_rows = rows[1:] if len(rows) > 1 else rows
    for row in data_rows:
        while len(row) < 4: row.append("")
    
    cols = st.columns([1, 2, 4, 2, 1, 2])
    for i, h in enumerate(["Ref", "Step", "Desc", "Owner", "Done", "Note"]):
        cols[i].markdown(f"**{h}**")
    
    table_key = f"{sec_key}::tbl::{table_index}"
    for ridx, row in enumerate(data_rows):
        row_key = f"{table_key}::row::{ridx}"
        comment_key = f"{row_key}::comment"
        ref, step, desc, owner = row[0], row[1], " ".join(row[2:-1]), row[-1]
        
        cols = st.columns([1, 2, 4, 2, 1, 2])
        cols[0].write(ref)
        cols[1].write(step)
        cols[2].write(desc)
        cols[3].write(owner)
        
        prev_val = completed_map.get(row_key, False)
        prev_comment = comments_map.get(comment_key, "")
        
        new_val = cols[4].checkbox("", value=prev_val, key=f"cb_{row_key}")
        new_comment = cols[5].text_input("", value=prev_comment, key=f"ci_{comment_key}", label_visibility="collapsed")
        
        if new_val != prev_val:
            completed_map[row_key] = new_val
        if new_comment != prev_comment:
            comments_map[comment_key] = new_comment

def render_section_content(section, playbook_name, completed_map, comments_map, sec_key, is_sub=False):
    table_idx = 0
    for item in section.get("content", []):
        if item["type"] == "text":
            st.markdown(f'<div style="font-size:1.1rem;line-height:1.6;">{item["value"].replace("\n", "<br/>")}</div>', unsafe_allow_html=True)
        elif item["type"] == "image":
            st.image(item["value"], use_column_width=True)
        elif item["type"] == "table":
            rows = item["value"]
            if is_action_table(rows):
                render_action_table(playbook_name, sec_key, rows, completed_map, comments_map, table_idx)
                table_idx += 1
            else:
                df = pd.DataFrame(rows[1:], columns=rows[0]) if len(rows) > 1 else pd.DataFrame(rows)
                st.dataframe(df, use_container_width=True, hide_index=True)
    
    for sub in section.get("subs", []):
        sub_key = stable_key(playbook_name, sub["title"], sub["level"])
        st.markdown(f"<strong style='color:var(--text);'>{sub['title']}</strong>", unsafe_allow_html=True)
        render_section_content(sub, playbook_name, completed_map, comments_map, sub_key, True)
    
    if not is_sub:
        prev_sec_comment = comments_map.get(sec_key, "")
        new_sec_comment = st.text_area("", value=prev_sec_comment, key=f"c::{sec_key}", height=100, label_visibility="collapsed")
        if new_sec_comment != prev_sec_comment:
            comments_map[sec_key] = new_sec_comment

# === MAIN APP ===
def main():
    user = authenticate()
    st.sidebar.info(f"**{user['name']}** – *{get_user_role(user['email'])}*")

    if get_user_role(user["email"]) == "admin" and st.sidebar.button("Admin Dashboard"):
        st.session_state.admin_page = True
        st.rerun()
    if st.session_state.get('admin_page', False):
        # Admin dashboard code here (unchanged for brevity)
        st.stop()

    # Header
    logo_html = get_logo()
    st.markdown(f"""
    <div class="sticky-header">
        {logo_html}
        <div class="app-title">Joval Wines NIST Playbook Tracker</div>
        <div class="nist-logo">NIST</div>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    st.sidebar.checkbox("Auto-save progress", value=True, key="autosave")
    st.sidebar.checkbox("Bulk export", value=False, key="bulk_export")

    # Playbook Select
    playbooks = sorted([f for f in os.listdir(PLAYBOOKS_DIR) if f.lower().endswith(".docx")])
    if not playbooks:
        st.error("No playbooks found.")
        return

    selected_playbook = st.selectbox("Playbook", [""] + playbooks, key="select_playbook")
    if not selected_playbook:
        st.stop()

    # Instructional Banner
    st.markdown("""
    <div style="background:#fff3cd;padding:1.2rem;border-radius:8px;border:2px solid #d9534f;
                text-align:center;font-size:1.4rem;font-weight:600;color:#d9534f;margin:1.5rem 0;">
        In the event of a cyber incident select the required playbook and complete each required step
        in the <strong>NIST "Incident Handling Categories"</strong> section.
    </div>
    """, unsafe_allow_html=True)

    # Load & Parse
    parsed_key = f"parsed::{selected_playbook}"
    if parsed_key not in st.session_state:
        st.session_state[parsed_key] = parse_playbook_cached(os.path.join(PLAYBOOKS_DIR, selected_playbook))
    sections = st.session_state[parsed_key]

    completed_map, comments_map = load_progress(selected_playbook)
    st.session_state[f"completed::{selected_playbook}"] = completed_map
    st.session_state[f"comments::{selected_playbook}"] = comments_map

    # Render Content
    st.markdown('<div class="content-wrap">', unsafe_allow_html=True)
    for sec in sections:
        sec_key = stable_key(selected_playbook, sec["title"], sec["level"])
        title_class = "nist-incident-section" if sec["title"] == "NIST Incident Handling Categories" else ""
        st.markdown(f"<div class='section-title {title_class}' id='{sec_key}'>{sec['title']}</div>", unsafe_allow_html=True)
        with st.expander("Expand section", expanded=False):
            render_section_content(sec, selected_playbook, completed_map, comments_map, sec_key)
    st.markdown('</div>', unsafe_allow_html=True)

    # Progress
    total = len([k for k in completed_map if '::row::' in k])
    done = sum(1 for v in completed_map.values() if v)
    pct = int((done / max(total, 1)) * 100)
    st.info(f"**Progress:** {pct}%")

    st.markdown(f"<div class='progress-wrap'><div class='progress-fill' style='width:{pct}%'></div></div>", unsafe_allow_html=True)

    # Action Buttons
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        if st.button("Save Progress"):
            save_progress(selected_playbook, completed_map, comments_map)
            st.success("Saved!")
    with c2:
        if st.button("**Reset Playbook**"):
            reset_playbook_progress(selected_playbook)
    with c3:
        csv_data = pd.DataFrame({
            "Task": list(completed_map.keys()),
            "Status": [str(v) for v in completed_map.values()]
        }).to_csv(index=False).encode()
        st.download_button("Download CSV", csv_data, f"{selected_playbook}_progress.csv", "text/csv")

    # Auto-save
    if st.session_state.autosave:
        save_progress(selected_playbook, completed_map, comments_map)

    # Bottom
    st.markdown(f"""
    <div class="bottom-toolbar">
        <div>© Joval Wines – Better Never Stops</div>
        <div>Progress: {pct}%</div>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
