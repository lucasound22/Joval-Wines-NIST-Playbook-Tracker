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

# === PAGE CONFIG & REMOVE ALL STREAMLIT BRANDING ===
st.set_page_config(
    page_title="Joval Wines NIST Playbook Tracker",
    page_icon="wine",
    layout="wide",
    initial_sidebar_state="expanded"
)

# HIDE ALL STREAMLIT BRANDING
hide_streamlit_style = """
<style>
    #MainMenu, footer, header, .stDeployButton, [data-testid="stToolbar"], 
    [data-testid="stHeader"], [data-testid="stFooter"], [data-testid="stDecoration"],
    .css-1d391kg, .css-1v0mbdj, .css-1y0t5a4, .css-1v3fvcr, .css-1v0mbdj a, 
    .css-1v0mbdj button, .css-1v0mbdj img {display: none !important;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# === CUSTOM STYLES (Reskinned for Joval Risk Look) ===
st.markdown(f"""
<style>
/* Core Colors for Joval Risk Look (Black/Red/Gold) */
:root{{
    --bg:#000000; /* Black background */
    --text:#ffffff; /* White text */
    --muted:#aaaaaa; /* Light gray muted text */
    --red:#800020; /* Dark Maroon Red - Joval primary color */
    --gold:#FFD700; /* Gold/Yellow accent */
    --card-bg:#1a1a1a; /* Dark gray for cards/sections */
    --border:rgba(255,255,255,0.1); /* Subtle white border */
    --blue-shadow:#4169E1; /* Keeping the blue for NIST text shadow */
}}

/* Global Overrides for Dark Theme */
html,body,.stApp{{background:var(--bg)!important;color:var(--text)!important;font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;}}
.stApp > footer,.stApp [data-testid="stToolbar"],.stDeployButton{{display:none!important;}}
/* Streamlit Component Styling */
.stTextInput > div > div > input, .stTextArea > div > textarea, .stSelectbox > div > div, .stFileUploader > div:first-child {{
    background-color: var(--card-bg);
    border: 1px solid var(--border);
    color: var(--text);
}}
.st-bh, .st-ck, .st-cl, .st-cq, .st-cr, .st-cs, .st-cu, .st-cv, .st-cw {{
    color: var(--text) !important; /* Ensure labels/text inside components are white */
}}
.st-by {{ /* Sidebar background */
    background-color: var(--card-bg) !important;
}}

/* Header - Simplified & Centered for Reskin */
.sticky-header{{
    position:sticky;top:0;z-index:9999;
    display:flex;align-items:center;justify-content:center; /* Center alignment */
    padding:1.5rem 2rem;background:var(--card-bg); /* Dark card background for header */
    border-bottom:1px solid var(--border);
    box-shadow:0 4px 12px rgba(0,0,0,.6);
    min-height:100px; /* Reduced height */
}}
.logo-left{{height:160px;width:auto;position:absolute;left:2rem;}} /* Keep logo left if present */
.app-title{{font-size:2.8rem;font-weight:900;color:var(--text);margin:0;}}
.nist-text{{
    font-size:2.8rem;
    font-weight:900;
    color:var(--gold); /* Gold for NIST text */
    text-shadow: 1px 1px 2px var(--red), 0 0 4px rgba(128,0,32,0.3); /* Red shadow accent */
    letter-spacing:1px;
    margin-right:8px;
}}
.nist-text sup{{font-size:1.2rem;color:var(--muted);}}

/* Section Titles */
.section-title,
.stExpander > div > div > div > label > div > span,
.stExpander > div > div > div > label > div > div > span {{
    font-size:1.9rem !important;
    font-weight:700 !important;
    color:var(--gold) !important; /* Gold for section titles */
    margin-bottom:0.5rem !important;
}}
.nist-incident-section {{
    color:var(--red) !important; /* Red for the specific section title */
    font-size:1.9rem !important;
    font-weight:700 !important;
}}

/* Content */
.content-wrap{{margin-left:280px;padding:2rem 2rem 6rem;}}
.section-card{{
    background:var(--card-bg);padding:1.5rem;border-radius:12px;
    margin-bottom:1.5rem;box-shadow:0 2px 6px rgba(0,0,0,.2);
    border:1px solid var(--border);
}}

/* Buttons (Primary Black/Dark, Secondary Red/Gold) */
.stButton>button,.stDownloadButton>button{{
    background:var(--red)!important;color:#fff!important; /* Use Red as primary button color */
    border-radius:8px;padding:0.75rem 1.5rem!important;
    font-weight:600;font-size:1rem;
    width:100%!important;min-height:52px;
    text-align:center;margin:0.6rem 0;
    border:1px solid var(--red);
}}
.stButton>button:hover,.stDownloadButton>button:hover{{opacity:.85;background:#a00030!important;}}
button[kind="secondary"] {{ /* Secondary/sidebar buttons */
    padding: 0.4rem 0.8rem !important;
    font-size: 0.85rem !important;
    min-height: 36px !important;
    background: var(--card-bg) !important;
    color: var(--gold) !important;
    border: 1px solid var(--gold) !important;
}}
button[kind="secondary"]:hover {{
    background: rgba(255,215,0,0.1) !important;
}}


/* Progress */
.progress-wrap{{height:12px;background:var(--card-bg);border-radius:999px;overflow:hidden;margin:1rem 0;}}
.progress-fill{{height:100%;background:var(--gold);transition:width .4s ease;}} /* Gold progress bar */

/* Checkbox/Radio/etc. */
.stCheckbox > label > div:first-child > div, .stRadio > label > div:first-child > div {{
    border-color: var(--muted);
}}
.stCheckbox > label > div:first-child > div:has(input:checked) {{
    background-color: var(--red);
    border-color: var(--red);
}}
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
    return True, "Password reset successfully.", password

def delete_user(email):
    users = load_users()
    email = email.lower()
    if email in users:
        del users[email]
        save_users(users)
        logging.info(f"User deleted: {email}")
        return True, "User deleted successfully."
    return False, "User not found."

def update_user(old_email, new_email, new_role):
    users = load_users()
    old_email = old_email.lower()
    new_email = new_email.lower()
    if old_email not in users:
        return False, "User not found."
    if new_email != old_email and new_email in users:
        return False, "New email already exists."
    
    user_data = users.pop(old_email)
    user_data["role"] = new_role
    users[new_email] = user_data
    save_users(users)
    logging.info(f"User updated: {old_email} → {new_email}, Role: {new_role}")
    return True, "User updated successfully."

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
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Create User", "Reset Password", "List & Edit Users", "Delete User", "Upload Logo/Playbook"])

    users = load_users()
    user_emails = sorted(users.keys())

    with tab1:
        st.subheader("Create New User")
        email_input = st.text_input("User Email")
        email = email_input if "@" in email_input else email_input + "@joval.com"
        role = st.selectbox("Role", ["user", "admin"], key="create_role")
        generate_pass = st.checkbox("Generate Random Password", value=True)
        if generate_pass:
            password = secrets.token_urlsafe(16)
            st.code(password, language=None)
            st.info("Copy this password now — it will not be shown again.")
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
        if not user_emails:
            st.info("No users to reset.")
        else:
            selected_user = st.selectbox("Select User", user_emails, key="reset_select")
            generate_pass = st.checkbox("Generate Random Password", value=True, key="reset_gen2")
            if generate_pass:
                password = secrets.token_urlsafe(16)
                st.code(password, language=None)
            else:
                password = st.text_input("Set New Password", type="password", key="reset_custom2")
            if st.button("Reset Password"):
                if password:
                    success, msg, new_pass = reset_user_password(selected_user, password)
                    if success:
                        st.success(msg)
                        if generate_pass:
                            st.code(new_pass, language=None)
                            st.info("New password shown above — copy it now.")
                        else:
                            st.success("New password set successfully.")
                    else:
                        st.error(msg)
                else:
                    st.error("Enter a password.")

    with tab3:
        st.subheader("List & Edit Users")
        if not users:
            st.info("No users.")
        else:
            user_list = [{"Email": k, "Role": v["role"]} for k, v in users.items()]
            df = pd.DataFrame(user_list)
            st.table(df)

            st.markdown("---")
            st.markdown("### Edit User")
            edit_email = st.selectbox("Select User to Edit", user_emails, key="edit_select")
            current_role = users[edit_email]["role"]
            new_email_input = st.text_input("New Email (leave blank to keep)", value=edit_email, key="edit_email")
            new_role = st.selectbox("New Role", ["user", "admin"], index=0 if current_role == "user" else 1, key="edit_role")

            if st.button("Update User"):
                new_email = new_email_input if new_email_input else edit_email
                if new_email != edit_email or new_role != current_role:
                    success, msg = update_user(edit_email, new_email, new_role)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
                else:
                    st.info("No changes made.")

    with tab4:
        st.subheader("Delete User")
        if not user_emails:
            st.info("No users to delete.")
        else:
            delete_email = st.selectbox("Select User to Delete", user_emails, key="delete_select")
            if st.button("Delete User"):
                success, msg = delete_user(delete_email)
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)

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

@st.cache_data(show_spinner=False)
def load_progress(playbook_name: str):
    path = progress_filepath(playbook_name)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
                return (
                    data.get("completed", {}),
                    data.get("comments", {}),
                    data.get("expanders", {})
                )
        except Exception as e:
            st.warning(f"Failed to load progress: {e}")
            return {}, {}, {}
    return {}, {}, {}

def save_progress(playbook_name: str, completed_map: dict, comments_map: dict, expanders_map: dict):
    rec = {
        "playbook": playbook_name,
        "timestamp": datetime.now().isoformat(),
        "version": "1.0",
        "completed": completed_map,
        "comments": comments_map,
        "expanders": expanders_map
    }
    path = progress_filepath(playbook_name)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(rec, fh, indent=2)

def safe_image_display(src: str) -> bool:
    if not src:
        return False
    try:
        st.markdown(f"<img style='max-width:90%;height:auto;border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,0.6);margin:12px 0;display:block;' src='{src}'/>", unsafe_allow_html=True)
        return True
    except Exception:
        try:
            st.image(src)
            return True
        except Exception:
            return False

def calculate_badges(pct: int) -> List[str]:
    if pct >= 100: return ["Gold Star"]
    elif pct >= 80: return ["Silver Shield"]
    elif pct >= 50: return ["Bronze Medal"]
    elif pct >= 25: return ["Progress Starter"]
    elif pct > 0: return ["Just Started"]
    else: return ["Ready to Begin"]

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
        return f'<img src="data:image/png;base64,{st.session_state.logo_b64}" class="logo-left" alt="Custom Logo" />'
    default_logo_path = "logo.png"
    if os.path.exists(default_logo_path):
        with open(default_logo_path, "rb") as f:
            logo_bytes = f.read()
            return f'<img src="data:image/png;base64,{base64.b64encode(logo_bytes).decode()}" class="logo-left" alt="Default Logo" />'
    return '<div class="logo-left"></div>'

def theme_selector():
    # The custom CSS block handles the dark theme directly based on the root variables
    # This function is now redundant as the main CSS is a dark theme, 
    # but keeping it simple/non-functional to avoid breaking the logic flow.
    # The existing theme in the original code is complex and can be simplified by the new CSS.
    st.sidebar.selectbox("Select Theme", ["Dark"], index=0, key="theme_selector")
    pass

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
            playbooks = [f for f in os.listdir(PLAYBOOKS_DIR) if f.endswith(".docx")]
            for pb in playbooks:
                if pb != selected_playbook:
                    comp, _, _ = load_progress(pb)
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

# === PLAYBOOK PARSING ===
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

# === RENDERING ===
ACTION_HEADERS = {"reference","ref","step","description","ownership","responsibility","owner","responsible"}

def is_action_table(rows: List[List[str]]) -> bool:
    if not rows:
        return False
    headers = [h.strip().lower() for h in rows[0]]
    hits = sum(1 for h in headers if any(k in h for k in ACTION_HEADERS))
    return hits >= 2 or (len(rows[0]) >= 4 and ref_pattern.match(rows[0][0].strip()))

# Global counter for progress
task_counter = {"total": 0, "done": 0}

def render_action_table(playbook_name, sec_key, rows, completed_map, comments_map, autosave, table_index=0):
    global task_counter
    default_headers = ["Reference", "Step", "Description", "Ownership/Responsibility"]
    headers = rows[0] if len(rows) > 0 and not ref_pattern.match(rows[0][0].strip() if rows[0] else "") else default_headers
    data_rows = rows[1:] if len(rows) > 1 and not ref_pattern.match(rows[0][0].strip() if rows[0] else "") else rows
    for row in data_rows:
        while len(row) < 4:
            row.append("")

    task_counter["total"] += len(data_rows)

    st.caption("Mark tasks complete and add notes.")
    cols = st.columns([1, 2, 4, 2, 1, 2])
    for i, h in enumerate(["Ref", "Step", "Desc", "Owner", "Done", "Comment"]):
        cols[i].write(f"**{h}**") # Bolding headers for clarity

    changed = False
    table_key = f"{sec_key}::tbl::{table_index}"
    for ridx, row in enumerate(data_rows):
        row_key = f"{table_key}::row::{ridx}"
        comment_key = f"{row_key}::comment"
        ref = row[0]; step = row[1]; desc = row[2]; owner = row[3] # Taking first 4 elements as per default
        
        # Heuristic for combining description, which was common in the parsing
        if len(row) > 4:
            desc = " ".join(row[2:-1])
            owner = row[-1]
        
        prev_val = completed_map.get(row_key, False)
        prev_comment = comments_map.get(comment_key, "")

        cb_key = f"cb_{playbook_name}_{sec_key}_{table_index}_{ridx}"
        ci_key = f"ci_{playbook_name}_{sec_key}_{table_index}_{ridx}"

        cols = st.columns([1, 2, 4, 2, 1, 2])
        cols[0].write(ref); cols[1].write(step); cols[2].write(desc); cols[3].write(owner)
        
        # Checkbox logic
        new_val = cols[4].checkbox("", value=prev_val, key=cb_key, label_visibility="collapsed")
        
        # Text input logic
        new_comment = cols[5].text_input("", value=prev_comment, key=ci_key, label_visibility="collapsed")

        # Update logic
        if new_val != prev_val:
            completed_map[row_key] = new_val
            changed = True
            # Update the global counter (needed for the progress bar calculation)
            if new_val:
                task_counter["done"] += 1
            else:
                task_counter["done"] -= 1
                
        if new_comment != prev_comment:
            comments_map[comment_key] = new_comment
            changed = True

    if autosave and changed:
        # Pass the current expander states from the session to save_progress
        save_progress(playbook_name, completed_map, comments_map, st.session_state.get("expanders", {}))


def render_generic_table(rows: List[List[str]]):
    if len(rows) > 1:
        # Convert all content to string for DataFrame consistency
        data = [[str(cell) for cell in row] for row in rows[1:]]
        headers = [str(h) for h in rows[0]]
        df = pd.DataFrame(data, columns=headers)
    else:
        df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

def render_section_content(section, playbook_name, completed_map, comments_map, autosave, sec_key, is_sub=False):
    table_idx = 0
    for item in section.get("content", []):
        t = item.get("type")
        if t == "text":
            text = item.get("value", "").replace("\n", "<br/>")
            st.markdown(f'<div style="font-size:1.1rem;line-height:1.6;">{text}</div>', unsafe_allow_html=True)
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
        st.markdown("<div style='font-weight:700;margin-top:12px;margin-bottom:6px;'>Comments / Notes</div>", unsafe_allow_html=True)
        prev_sec_comment = comments_map.get(sec_key, "")
        sec_comment_key = f"sec_cmt_{playbook_name}_{sec_key}"
        new_sec_comment = st.text_area("", value=prev_sec_comment, key=sec_comment_key, height=120, label_visibility="collapsed")
        if new_sec_comment != prev_sec_comment:
            comments_map[sec_key] = new_sec_comment
            if autosave:
                save_progress(playbook_name, completed_map, comments_map, st.session_state.get("expanders", {}))

def get_expander_state_key(playbook_name: str, sec_key: str) -> str:
    return f"exp_{playbook_name}_{sec_key}"

def load_expander_states(playbook_name: str, sections: List[Dict]) -> Dict[str, bool]:
    _, _, saved_states = load_progress(playbook_name)
    states = {}
    for sec in sections:
        key = stable_key(playbook_name, sec["title"], sec["level"])
        state_key = get_expander_state_key(playbook_name, key)
        states[key] = saved_states.get(state_key, False)
        for sub in sec.get("subs", []):
            sub_key = stable_key(playbook_name, sub["title"], sub["level"])
            sub_state_key = get_expander_state_key(playbook_name, sub_key)
            states[sub_key] = saved_states.get(sub_state_key, False)
    return states

def save_expander_state(playbook_name: str, sec_key: str, state: bool):
    # Ensure that all current progress is loaded before saving a single expander state
    completed, comments, expanders = load_progress(playbook_name)
    expanders[get_expander_state_key(playbook_name, sec_key)] = state
    save_progress(playbook_name, completed, comments, expanders)

def render_section(section, playbook_name, completed_map, comments_map, autosave, expander_states):
    sec_key = stable_key(playbook_name, section["title"], section["level"])
    title_class = "nist-incident-section" if section["title"] == "NIST Incident Handling Categories" else "section-title"
    
    # Use a hidden markdown for the section title so the expander label is cleaner
    st.markdown(f"<div id='{sec_key}' class='{title_class}'>{section['title']}</div>", unsafe_allow_html=True)
    
    # Store expander state in session_state, initialized from saved state
    state_key = get_expander_state_key(playbook_name, sec_key)
    if state_key not in st.session_state:
        st.session_state[state_key] = expander_states.get(sec_key, False)

    # Use a key to link the expander state to the session state variable
    expander_label = section["title"].split(' - ')[-1].split(':')[-1].strip() # Cleaner label
    
    with st.expander(expander_label, expanded=st.session_state[state_key], key=f"expander_{sec_key}"):
        
        # Streamlit's expander widget automatically manages its state in st.session_state.
        # We need to manually check if the expander state has changed from the SAVED state 
        # and trigger a save only if it's different.
        current_state_from_widget = st.session_state[f"expander_{sec_key}"]
        
        # Check if the current state in the session state needs to be updated and saved
        if current_state_from_widget != st.session_state[state_key]:
            st.session_state[state_key] = current_state_from_widget
            if autosave:
                # Save the new state
                save_expander_state(playbook_name, sec_key, current_state_from_widget)
                expander_states[sec_key] = current_state_from_widget
        
        render_section_content(section, playbook_name, completed_map, comments_map, autosave, sec_key)

# === MAIN APP ===
def main():
    user = authenticate()
    st.sidebar.info(f"Logged in as: **{user['name']}** – *{get_user_role(user['email'])}*")

    if get_user_role(user["email"]) == "admin":
        if "admin_page" not in st.session_state:
            st.session_state.admin_page = False
        if st.sidebar.button("Admin Dashboard"):
            st.session_state.admin_page = True
            st.rerun()
        if st.session_state.admin_page:
            admin_dashboard(user)
            return

    # Header Bar
    st.markdown(f"""
    <div class="sticky-header">
        {get_logo()}
        <h1 class="app-title">
            <span class="nist-text">Joval NIST</span> Playbook Tracker
        </h1>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    playbooks = sorted([f for f in os.listdir(PLAYBOOKS_DIR) if f.endswith(".docx")])
    if not playbooks:
        st.error("No playbooks found in the 'playbooks' directory.")
        return

    # Use a cleaner display name for the selector
    playbook_options = {p: p.replace('.docx', '').replace('_', ' ').title() for p in playbooks}
    
    st.sidebar.markdown("## 📖 Select Playbook")
    selected_playbook_display = st.sidebar.selectbox(
        "Playbook", 
        list(playbook_options.values()), 
        key="playbook_select",
        label_visibility="collapsed"
    )
    
    # Map back to the filename
    selected_playbook = next(key for key, value in playbook_options.items() if value == selected_playbook_display)

    # Theme selector (non-functional/simple due to custom CSS handling dark mode)
    theme_selector()

    st.sidebar.markdown("---")
    st.sidebar.markdown("## 📊 Progress & Tools")
    autosave = st.sidebar.checkbox("Autosave Progress", value=True, key="autosave_cb")
    
    # Reset global counter
    task_counter["total"] = 0
    task_counter["done"] = 0

    # Load data
    sections = parse_playbook_cached(os.path.join(PLAYBOOKS_DIR, selected_playbook))
    completed_map, comments_map, _ = load_progress(selected_playbook)
    expander_states = load_expander_states(selected_playbook, sections)

    # Pre-calculate initial done count from loaded data (necessary for first run)
    initial_done_count = sum(1 for v in completed_map.values() if v is True)
    task_counter["done"] = initial_done_count
    
    # Calculate Total Tasks by traversing the structure *before* rendering
    def count_actions(nodes):
        count = 0
        for node in nodes:
            for item in node.get("content", []):
                if item.get("type") == "table" and is_action_table(item.get("value", [])):
                    rows = item.get("value", [])
                    data_rows = rows[1:] if len(rows) > 1 and not ref_pattern.match(rows[0][0].strip() if rows[0] else "") else rows
                    count += len(data_rows)
            count += count_actions(node.get("subs", []))
        return count

    # Only set total count if it's not already updated during the loop (which it shouldn't be at this point)
    if task_counter["total"] == 0: 
        total_tasks = count_actions(sections)
        task_counter["total"] = total_tasks
    else:
        # If it was updated during an un-cached run (e.g. from render_action_table), keep that value.
        total_tasks = task_counter["total"]

    # Post-rendering updates (actual done count might change in render_action_table)
    done_tasks = task_counter["done"]
    pct = (done_tasks / total_tasks * 100) if total_tasks > 0 else 0
    pct = round(pct)

    st.sidebar.markdown(f"**Completion:** {done_tasks} / {total_tasks} Actions")
    st.sidebar.markdown(f"""
    <div class="progress-wrap">
        <div class="progress-fill" style="width:{pct}%;"></div>
    </div>
    <div style='text-align:center;font-weight:700;'>{pct}% Complete</div>
    """, unsafe_allow_html=True)
    
    # Badges
    badges = calculate_badges(pct)
    st.sidebar.markdown("---")
    st.sidebar.markdown("## 🏆 Achievements")
    st.sidebar.markdown(f"**Current Status:** *{', '.join(badges)}*")

    # Export Tools
    st.sidebar.markdown("---")
    st.sidebar.markdown("## 📤 Export Data")
    if OPENPYXL_AVAILABLE:
        excel_data = export_to_excel(completed_map, comments_map, selected_playbook)
        st.sidebar.download_button(
            label="Download Progress (Excel)",
            data=excel_data,
            file_name=f"{os.path.splitext(selected_playbook)[0]}_progress.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="excel_download"
        )
        # Optional bulk export
        # st.sidebar.download_button(...)
    
    csv_data = export_to_csv(completed_map, comments_map, selected_playbook)
    st.sidebar.download_button(
        label="Download Progress (CSV)",
        data=csv_data,
        file_name=f"{os.path.splitext(selected_playbook)[0]}_progress.csv",
        mime="text/csv",
        key="csv_download"
    )

    # Main Content Area
    st.markdown("<div class='content-wrap'>", unsafe_allow_html=True)

    # Table of Contents
    st.markdown("## 🧭 Table of Contents")
    toc_search = st.text_input("Filter Sections", key="toc_filter", placeholder="Search for a section or control...")
    st.markdown('<div class="section-card" style="max-height: 400px; overflow-y: auto;">', unsafe_allow_html=True)
    
    def render_toc_link(sec, indent_level=0):
        sec_key = stable_key(selected_playbook, sec["title"], sec["level"])
        title_text = sec["title"]
        if toc_search.lower() in title_text.lower():
            st.markdown(
                f"<a href='#{sec_key}' class='toc-item' style='margin-left:{indent_level*15}px;font-weight:{600 - indent_level*100};'>{title_text}</a>",
                unsafe_allow_html=True
            )
        for sub in sec.get("subs", []):
            render_toc_link(sub, indent_level + 1)

    for section in sections:
        render_toc_link(section)

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("## 📜 Playbook Content")

    # Render sections
    for section in sections:
        st.markdown("<div class='section-card'>", unsafe_allow_html=True)
        render_section(section, selected_playbook, completed_map, comments_map, autosave, expander_states)
        st.markdown("</div>", unsafe_allow_html=True)

    show_feedback()

    # Final save of any latent changes
    if not autosave:
        if st.button("Manual Save Progress"):
            save_progress(selected_playbook, completed_map, comments_map, st.session_state.get("expanders", {}))
            st.success("Progress manually saved!")

    st.markdown("</div>", unsafe_allow_html=True) # close content-wrap

if __name__ == '__main__':
    main()
