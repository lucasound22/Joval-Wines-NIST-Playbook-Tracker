import os
import io
import re
import json
import base64
import hashlib
import secrets
import time
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any

import streamlit as st
import mammoth
from bs4 import BeautifulSoup
import pandas as pd

try:
    # Attempt to import openpyxl for Excel export
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

import logging

# === CONFIGURATION & NIST/SECURITY SETUP ===
# Regex to identify control/step references (e.g., 3.1.2, 5.2)
ref_pattern = re.compile(r'^\d+(\.\d+)*\b') 

# Define logging for security-relevant actions (AU-2)
LOG_FILE = 'security_audit.log'
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logging.getLogger().setLevel(logging.INFO) # Ensure logger is active

PLAYBOOKS_DIR = "playbooks"
USERS_FILE = "users.json"
Path(PLAYBOOKS_DIR).mkdir(exist_ok=True)
Path(USERS_FILE).touch(exist_ok=True)

# NIST Compliance Note (SC-28): In a production environment, file-based storage 
# for user credentials and progress is highly discouraged. A persistent, 
# encrypted database (like Firestore or an encrypted SQL database) should be used 
# to meet data integrity and persistence requirements.

# === PAGE CONFIG & REMOVE ALL STREAMLIT BRANDING ===
st.set_page_config(
    page_title="Joval Wines NIST Playbook Tracker",
    page_icon="🍷",
    layout="wide",
    initial_sidebar_state="expanded"
)

# HIDE ALL STREAMLIT BRANDING (NIST UI/UX control)
hide_streamlit_style = """
<style>
    #MainMenu, footer, header, .stDeployButton, [data-testid="stToolbar"], 
    [data-testid="stHeader"], [data-testid="stFooter"], [data-testid="stDecoration"],
    .css-1d391kg, .css-1v0mbdj, .css-1y0t5a4, .css-1v3fvcr, .css-1v0mbdj a, 
    .css-1v0mbdj button, .css-1v0mbdj img {display: none !important;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# === CUSTOM STYLES (Applying Joval Wines Red/Gold/Blue Skin) ===
st.markdown(f"""
<style>
/* Tailwind CDN */
@import url('https://cdn.tailwindcss.com');
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@100..900&display=swap');

/* Core Colors - Wine-themed */
:root{{
    --bg:#ffffff;
    --text:#111111;
    --muted:#666666;
    --red:#800020; /* Deep Burgundy Red */
    --gold:#FFD700; /* Gold Accent */
    --blue-shadow:#4169E1; /* Royal Blue for highlight */
    --card-bg:#fdfdfd;
    --border:#eaeaea;
    --progress-bg: #E0E0E0;
}}

/* Dark Theme Overrides */
.dark-mode {{
    --bg:#0e1117;
    --text:#e3e3e3;
    --muted:#aaaaaa;
    --card-bg:rgba(255,255,255,0.05);
    --border:rgba(255,255,255,0.15);
    --progress-bg: rgba(255,255,255,0.1);
}}

/* Global */
html,body,.stApp{{
    background:var(--bg)!important;
    color:var(--text)!important;
    font-family:'Inter', sans-serif;
    transition: background-color 0.3s, color 0.3s;
}}
.stApp > footer,.stApp [data-testid="stToolbar"],.stDeployButton{{display:none!important;}}

/* Header */
.sticky-header{{
    position:sticky;top:0;z-index:9999;
    display:flex;align-items:center;justify-content:space-between;
    padding:1.2rem 2rem;background:var(--bg);
    border-bottom:1px solid var(--border);box-shadow:0 2px 8px rgba(0,0,0,.05);
    min-height:120px;
    transition: background-color 0.3s, border-color 0.3s, box-shadow 0.3s;
}}
.logo-left{{height:160px;width:auto;margin-right:20px;}}
.app-title{{font-size:2.4rem;font-weight:700;color:var(--text);margin:0;text-align:center;flex:1;}}
.nist-text{{
    font-size:2.8rem;
    font-weight:900;
    color:var(--red);
    text-shadow: 1px 1px 2px var(--blue-shadow), 0 0 4px rgba(65,105,225,0.3);
    letter-spacing:1px;
    margin-right:8px;
}}
.dark-mode .nist-text {{
    color: var(--gold);
    text-shadow: 1px 1px 2px var(--blue-shadow), 0 0 4px rgba(65,105,225,0.5);
}}
.nist-text sup{{font-size:1.2rem;color:var(--muted);}}

/* Sidebar/TOC */
.stSidebar > div:first-child {{
    background-color: var(--card-bg) !important;
    border-right: 1px solid var(--border);
    padding-top: 2rem;
}}

/* Section Titles */
.section-title,
.stExpander > div > div > div > label > div > span,
.stExpander > div > div > div > label > div > div > span {{
    font-size:1.9rem !important;
    font-weight:700 !important;
    color:var(--text) !important;
    margin-bottom:0.5rem !important;
}}
.nist-incident-section {{
    color:var(--red) !important;
    font-size:1.9rem !important;
    font-weight:700 !important;
}}

/* TOC Search */
.toc-search input {{
    width: 100%;
    padding: 0.5rem;
    border: 1px solid var(--border);
    border-radius: 6px;
    font-size: 0.9rem;
    margin-bottom: 0.5rem;
    background: var(--bg);
    color: var(--text);
}}
.toc-item {{display:block;padding:4px 0;color:var(--text);text-decoration:none;font-size: 0.95rem;}}
.toc-item:hover {{color:var(--red);font-weight:600;}}
.toc-item.active {{color:var(--red);font-weight:700;}}

/* Content */
.content-wrap{{padding:2rem 2rem 6rem;}}
.section-card{{
    background:var(--card-bg);padding:1.5rem;border-radius:12px;
    margin-bottom:1.5rem;box-shadow:0 2px 6px rgba(0,0,0,.04);
    border:1px solid var(--border);
    transition: background-color 0.3s, border-color 0.3s, box-shadow 0.3s;
}}

/* Buttons */
.stButton>button,.stDownloadButton>button{{
    background:var(--red)!important;color:#fff!important;
    border-radius:8px;padding:0.75rem 1.5rem!important;
    font-weight:600;font-size:1rem;
    width:100%!important;min-height:52px;
    text-align:center;margin:0.6rem 0;
    transition: background-color 0.3s, opacity 0.3s;
}}
.stButton>button:hover,.stDownloadButton>button:hover{{background: #A00030!important;}}
.stDownloadButton>button{{background: #000!important;}}
.stDownloadButton>button:hover{{background: #333!important;}}

/* Progress Bar */
.progress-wrap{{height:12px;background:var(--progress-bg);border-radius:999px;overflow:hidden;margin:1rem 0;}}
.progress-fill{{height:100%;background:var(--red);transition:width .4s ease;}}

/* Action Table Styling */
.action-table-header {{
    font-weight: 700;
    padding: 8px 0;
    border-bottom: 2px solid var(--border);
    display: flex;
    margin-bottom: 10px;
}}
.action-table-row {{
    padding: 8px 0;
    border-bottom: 1px dashed rgba(128,0,32,0.1);
    display: flex;
    align-items: flex-start;
}}
.dark-mode .action-table-row {{
    border-bottom: 1px dashed rgba(255,255,255,0.1);
}}
.action-table-row:last-child {{border-bottom: none;}}

.action-ref {{ width: 8%; font-weight: 600; }}
.action-step {{ width: 15%; font-style: italic; }}
.action-desc {{ width: 35%; line-height: 1.4; }}
.action-owner {{ width: 15%; font-size: 0.9rem; color: var(--muted); }}
.action-done {{ width: 5%; }}
.action-comment {{ width: 22%; }}

</style>
""", unsafe_allow_html=True)

# === USER MANAGEMENT ===

def load_users():
    """Loads users from the JSON file, ensuring the file exists and handles decoding errors."""
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                content = f.read().strip()
                if content:
                    return {k.lower(): v for k, v in json.loads(content).items()}
        except (json.JSONDecodeError, ValueError):
            st.warning("User file corrupted. Resetting default admin.")
            pass

    # NIST Security Control (AC-7, AC-3): Ensure a secure default admin exists.
    admin_email = "admin@joval.com"
    # Use st.secrets for secure credential management
    admin_hash = st.secrets.get("ADMIN_PASSWORD_HASH") 
    if not admin_hash:
        # st.error("ADMIN_PASSWORD_HASH not set in secrets.toml. App cannot run securely.")
        # Fallback for local testing if secrets not configured
        admin_hash = hashlib.sha256("defaultpass123".encode()).hexdigest()
        # st.stop()
    
    default_admin = {admin_email: {"role": "admin", "hash": admin_hash}}
    save_users(default_admin)
    return default_admin

def save_users(users):
    """Saves the user dictionary back to the JSON file."""
    with open(USERS_FILE, "w") as f:
        # Ensure only lowercased emails are keys
        json.dump({k.lower(): v for k, v in users.items()}, f, indent=2)

def get_user_role(email):
    """Retrieves the role for a given email."""
    users = load_users()
    return users.get(email.lower(), {}).get("role", "user")

def validate_email(email):
    """Basic email format validation."""
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

def validate_password(password):
    """Basic password strength check (NIST IA-5)."""
    if len(password) < 12:
        return False
    if not any(c.isupper() for c in password):
        return False
    if not any(c.islower() for c in password):
        return False
    if not any(c.isdigit() for c in password):
        return False
    # Simplified check for special character for demo purposes
    if not re.search(r'[!@#$%^&*(),.?":{}|<>]', password):
        return False
    return True

def create_user(email, role, password):
    """Creates a new user with password hashing."""
    email = email.lower()
    if not validate_email(email):
        return False, "Invalid email format."
    if not validate_password(password):
        return False, "Password must be at least 12 characters and include uppercase, lowercase, number, and a symbol."
        
    users = load_users()
    if email in users:
        return False, "User already exists."
    
    # NIST IA-7: Password hashing for storage
    hash_pass = hashlib.sha256(password.encode()).hexdigest() # Placeholder for bcrypt/scrypt
    users[email] = {"role": role, "hash": hash_pass}
    save_users(users)
    logging.info(f"User created (AC-2, AC-7): {email}, Role: {role}")
    return True, "User created successfully."

def reset_user_password(email, password):
    """Resets a user's password."""
    email = email.lower()
    if not validate_email(email):
        return False, "Invalid email format."
    if not validate_password(password):
        return False, "New password failed complexity requirements."
        
    users = load_users()
    if email not in users:
        return False, "User not found."
    
    hash_pass = hashlib.sha256(password.encode()).hexdigest()
    users[email]["hash"] = hash_pass
    save_users(users)
    logging.info(f"Password reset (AC-2, IA-5): {email}")
    return True, "Password reset successfully.", password

def delete_user(email):
    """Deletes a user."""
    users = load_users()
    email = email.lower()
    if email in users:
        del users[email]
        save_users(users)
        logging.info(f"User deleted (AC-2): {email}")
        return True, "User deleted successfully."
    return False, "User not found."

def update_user(old_email, new_email, new_role):
    """Updates a user's email or role."""
    users = load_users()
    old_email = old_email.lower()
    new_email = new_email.lower()
    
    if old_email not in users:
        return False, "User not found."
    if new_email != old_email and new_email in users:
        return False, "New email already exists."
    if not validate_email(new_email):
        return False, "Invalid new email format."
    
    user_data = users.pop(old_email)
    user_data["role"] = new_role
    users[new_email] = user_data
    save_users(users)
    logging.info(f"User updated (AC-2): {old_email} -> {new_email}, Role: {new_role}")
    return True, "User updated successfully."

def authenticate():
    """Handles user login with basic brute-force protection (AC-7)."""
    if 'login_attempts' not in st.session_state:
        st.session_state.login_attempts = 0
    if 'last_attempt' not in st.session_state:
        st.session_state.last_attempt = None
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'user' not in st.session_state:
        st.session_state.user = None

    if not st.session_state.authenticated:
        # Custom Login Layout
        login_container = st.container()
        login_container.markdown(
            '<div class="flex flex-col items-center justify-center p-8 bg-white dark-mode:bg-[#1f2937] rounded-xl shadow-2xl mt-12 w-full max-w-lg mx-auto border border-gray-200 dark-mode:border-gray-700">', 
            unsafe_allow_html=True
        )
        login_container.markdown(get_logo(), unsafe_allow_html=True)
        login_container.markdown('<h1 class="app-title text-center !text-4xl !mt-4">NIST Playbook Tracker</h1>', unsafe_allow_html=True)
        
        with login_container.form("login_form"):
            st.markdown("### Login Required")
            
            username = st.text_input("Username (Email or prefix)", key="username_input")
            password = st.text_input("Password", type="password", key="password_input")
            
            login_button = st.form_submit_button("Log In")

        login_container.markdown('</div>', unsafe_allow_html=True)

        if login_button:
            now = datetime.now()
            
            # AC-7: Rate-limiting check
            if st.session_state.login_attempts >= 5:
                # 5-minute lockout (300 seconds)
                if st.session_state.last_attempt and (now - st.session_state.last_attempt).seconds < 300:
                    st.error("Too many failed attempts. Try again in 5 minutes.")
                    time.sleep(1) # Simple delay to mitigate rapid script attempts
                    st.rerun()

            # Attempt to normalize email
            email = username if "@" in username else username + "@joval.com"
            email = email.lower()
            
            users = load_users()
            is_valid = False
            
            if email in users:
                # AC-7: Use constant-time comparison if possible, but sha256 is fine for a demo
                if hashlib.sha256(password.encode()).hexdigest() == users[email]["hash"]:
                    is_valid = True
            
            if is_valid:
                st.session_state.authenticated = True
                display_name = username.split("@")[0].title() if "@" in username else username.title()
                st.session_state.user = {"email": email, "name": display_name, "role": users[email]["role"]}
                st.session_state.login_attempts = 0
                logging.info(f"Successful login (AC-7): {email}")
                st.success("Logged in successfully!")
                st.rerun()
            else:
                st.session_state.login_attempts += 1
                st.session_state.last_attempt = now
                logging.warning(f"Failed login attempt (AC-7): {email}. Attempts: {st.session_state.login_attempts}")
                st.error("Invalid credentials.")
                time.sleep(1) # Simple delay to mitigate rapid script attempts
                st.rerun() # Re-run to show error message and update attempt count

        st.stop() # Stop the rest of the app execution until authenticated

    # Logout button is always visible in the sidebar after auth
    if st.sidebar.button("Logout", key="logout_btn"):
        logging.info(f"User logged out: {st.session_state.user['email']}")
        for key in list(st.session_state.keys()):
            if key not in ['theme_selector']: # Keep theme state
                del st.session_state[key]
        st.rerun()

    return st.session_state.user

# === ADMIN DASHBOARD ===
def admin_dashboard(user):
    """Admin controls for user management and content upload (AC-3, AC-6)."""
    if get_user_role(user["email"]) != "admin":
        st.error("Access denied. Admin only.")
        return

    st.title("Admin Dashboard")
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Create User", "Reset Password", "List & Edit Users", "Delete User", "Upload Content"])

    users = load_users()
    # Filter out the currently logged-in admin from being deleted/edited by themselves
    user_emails = sorted([e for e in users.keys() if e != user["email"]])
    all_user_emails = sorted(users.keys())

    with tab1:
        st.subheader("Create New User (AC-2)")
        email_input = st.text_input("User Email", key="create_email_input")
        email = email_input if "@" in email_input else email_input + "@joval.com"
        role = st.selectbox("Role (AC-3)", ["user", "admin"], key="create_role")
        generate_pass = st.checkbox("Generate Random Password (IA-5)", value=True)
        password = ""
        if generate_pass:
            # IA-5(11): Randomly generated initial passwords
            password = secrets.token_urlsafe(16)
            st.code(password, language=None)
            st.warning("Copy this password now — it will not be shown again.")
        else:
            password = st.text_input("Set Password (Min 12 chars, complex)", type="password", key="create_pass")
            
        if st.button("Create User"):
            if email and password:
                success, msg = create_user(email, role, password)
                if success:
                    st.success(msg)
                    # Clear inputs after success
                    st.session_state.create_email_input = ""
                    if not generate_pass: st.session_state.create_pass = ""
                    # Trigger rerun to update user list in other tabs
                    st.rerun()
                else:
                    st.error(msg)
            else:
                st.error("Fill all fields.")

    with tab2:
        st.subheader("Reset User Password (AC-2, IA-5)")
        if not all_user_emails:
            st.info("No users to reset.")
        else:
            selected_user = st.selectbox("Select User", all_user_emails, key="reset_select")
            generate_pass = st.checkbox("Generate Random Password (IA-5)", value=True, key="reset_gen2")
            password = ""
            if generate_pass:
                password = secrets.token_urlsafe(16)
                st.code(password, language=None)
                st.warning("New password shown above — copy it now.")
            else:
                password = st.text_input("Set New Password (Min 12 chars, complex)", type="password", key="reset_custom2")
            if st.button("Reset Password"):
                if password:
                    success, msg, new_pass = reset_user_password(selected_user, password)
                    if success:
                        st.success(msg)
                    else:
                        st.error(msg)
                else:
                    st.error("Enter a password.")

    with tab3:
        st.subheader("List & Edit Users (AC-3)")
        if not users:
            st.info("No users.")
        else:
            user_list = [{"Email": k, "Role": v["role"]} for k, v in users.items()]
            df = pd.DataFrame(user_list)
            # Simple list table
            st.markdown(df.to_html(index=False, classes='table-auto w-full'), unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("### Edit User")
            
            # Selectable list for editing (include current user, but prevent role change if only one admin)
            edit_email = st.selectbox("Select User to Edit", all_user_emails, key="edit_select")
            if edit_email:
                current_role = users[edit_email]["role"]
                # Prevent editing own email or the last admin's role
                is_last_admin = (current_role == "admin" and sum(1 for v in users.values() if v["role"] == "admin") == 1)

                new_email_input = st.text_input("New Email", value=edit_email, key="edit_email_input")
                new_role = st.selectbox("New Role", ["user", "admin"], 
                                        index=0 if current_role == "user" else 1, 
                                        key="edit_role")
                
                if is_last_admin and edit_email == user["email"]:
                    st.warning("You are the only admin. Cannot demote your role.")
                    
                if st.button("Update User"):
                    new_email = new_email_input if new_email_input else edit_email
                    if new_email != edit_email or new_role != current_role:
                        if is_last_admin and new_role == "user":
                            st.error("Cannot demote the last remaining admin.")
                        else:
                            success, msg = update_user(edit_email, new_email, new_role)
                            if success:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
                    else:
                        st.info("No changes made.")

    with tab4:
        st.subheader("Delete User (AC-2)")
        if not user_emails:
            st.info("No non-admin users to delete, or you are the only one listed.")
        else:
            delete_email = st.selectbox("Select User to Delete", user_emails, key="delete_select")
            if st.button("Delete User", type="primary"):
                if delete_email == user["email"]:
                    st.error("You cannot delete your own active account.")
                else:
                    success, msg = delete_user(delete_email)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

    with tab5:
        st.subheader("Upload Custom Logo")
        uploaded_logo = st.file_uploader("Upload Logo (.png, .jpg)", type=["png", "jpg", "jpeg"])
        if uploaded_logo:
            # Store logo in session state for display
            st.session_state.logo_b64 = base64.b64encode(uploaded_logo.read()).decode()
            st.success("Logo uploaded!")
            st.rerun()
            
        st.subheader("Upload New Playbook (CM-4, CM-5)")
        uploaded_playbook = st.file_uploader("Upload Word Doc (.docx)", type=["docx"])
        if uploaded_playbook:
            file_path = Path(PLAYBOOKS_DIR) / uploaded_playbook.name
            with open(file_path, "wb") as f:
                f.write(uploaded_playbook.getbuffer())
            st.success(f"Playbook '{uploaded_playbook.name}' uploaded! Please refresh the page.")
            
    if st.button("Back to Main App", key="back_to_main_btn"):
        st.session_state.admin_page = False
        st.rerun()

# === UTILITIES ===
def stable_key(playbook_name: str, title: str, level: int) -> str:
    """Generates a stable, unique key for a section based on its content (CM-4)."""
    base = f"{playbook_name}||{level}||{title}"
    return "sec_" + hashlib.md5(base.encode("utf-8")).hexdigest()

def progress_filepath(playbook_name: str) -> str:
    """Determines the file path for saving progress JSON."""
    base = os.path.splitext(playbook_name)[0]
    return os.path.join(PLAYBOOKS_DIR, f"{base}_progress.json")

# Use a shorter TTL for development/testing but keep the cache for performance
@st.cache_data(show_spinner=False, ttl=60) 
def load_progress(playbook_name: str):
    """Loads completion progress, comments, and expander state."""
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
    """Saves completion progress, comments, and expander state (SI-12, SC-28)."""
    rec = {
        "playbook": playbook_name,
        "timestamp": datetime.now().isoformat(),
        "user_email": st.session_state.user["email"], # Log user making changes
        "version": "1.0",
        "completed": completed_map,
        "comments": comments_map,
        "expanders": expanders_map
    }
    path = progress_filepath(playbook_name)
    try:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(rec, fh, indent=2)
    except Exception as e:
        st.error(f"Error saving progress: {e}")

def safe_image_display(src: str) -> bool:
    """Displays base64 images from mammoth output safely."""
    if not src or not src.startswith("data:"):
        # Fallback for external URLs or broken links
        st.caption("Image placeholder (data not embedded or external link used)")
        return False
    try:
        # Improved styling for images
        st.markdown(f"<img style='max-width:90%;height:auto;border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,0.1);margin:12px 0;display:block;' src='{src}'/>", unsafe_allow_html=True)
        return True
    except Exception:
        return False

def calculate_badges(pct: int) -> List[str]:
    """Awards badges based on completion percentage."""
    if pct >= 100: return ["🏆 Gold Star: 100% Complete"]
    elif pct >= 90: return ["🥇 Platinum Shield"]
    elif pct >= 75: return ["🥈 Gold Tier"]
    elif pct >= 50: return ["🥉 Silver Tier"]
    elif pct >= 25: return ["Bronze Starter"]
    elif pct > 0: return ["Just Started"]
    else: return ["Ready to Begin"]

def save_feedback(rating: int, comments: str, user_email: str):
    """Saves user feedback to a JSONL file."""
    feedback_data = {
        "user": user_email,
        "rating": rating, 
        "comments": comments, 
        "timestamp": datetime.now().isoformat()
    }
    with open("feedback.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(feedback_data) + "\n")

def show_feedback(user_email: str):
    """Provides a form for users to submit feedback."""
    with st.expander("Provide Feedback (Help us improve)", expanded=False):
        with st.form("feedback_form"):
            rating = st.slider("Rate this tool (1=Poor, 5=Excellent)", 1, 5, 4)
            feedback_comments = st.text_area("Additional feedback or suggestions")
            if st.form_submit_button("Submit Feedback"):
                save_feedback(rating, feedback_comments, user_email)
                st.success("Feedback submitted! Thank you.")

def get_logo():
    """Retrieves the custom or default logo for the header."""
    if "logo_b64" not in st.session_state:
        st.session_state.logo_b64 = None
    if st.session_state.logo_b64:
        return f'<img src="data:image/png;base64,{st.session_state.logo_b64}" class="logo-left" alt="Custom Logo" />'
    
    # Placeholder/Default Logo Logic - Assumes a local file 'logo.png' might exist
    default_logo_path = "logo.png"
    if os.path.exists(default_logo_path):
        try:
            with open(default_logo_path, "rb") as f:
                logo_bytes = f.read()
                return f'<img src="data:image/png;base64,{base64.b64encode(logo_bytes).decode()}" class="logo-left" alt="Default Logo" />'
        except Exception:
            pass # Fall through to default placeholder
            
    # Fallback placeholder if no image file is found
    return '<div class="logo-left"></div>'

def theme_selector():
    """Handles the light/dark mode switching."""
    theme = st.sidebar.selectbox("Select Theme", ["Light", "Dark"], 
                                 index=st.session_state.get("theme_index", 0), 
                                 key="theme_selector")
    st.session_state.theme_index = ["Light", "Dark"].index(theme)

    if theme == "Dark":
        st.markdown("""
        <script>document.body.classList.add('dark-mode');</script>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <script>document.body.classList.remove('dark-mode');</script>
        """, unsafe_allow_html=True)
    return theme

@st.cache_data(ttl=300)
def export_to_excel(completed_map: Dict, comments_map: Dict, selected_playbook: str, all_playbooks: List[str], bulk_export: bool = False) -> bytes:
    """Exports data to a single Excel file (SI-12, CM-4)."""
    if not OPENPYXL_AVAILABLE:
        st.error("The openpyxl library is required for Excel export.")
        return b""
    
    output = io.BytesIO()
    
    # Prepare data for the current playbook
    progress_list = []
    for key, status in completed_map.items():
        # Task keys are long, try to infer the section/row ID
        section_key = key.split("::tbl::")[0]
        comment = comments_map.get(f"{section_key}::{key.split('::tbl::')[1]}::row::{key.split('::row::')[1]}::comment")
        progress_list.append({
            "Section_Key": section_key,
            "Task_Key": key,
            "Status": "Complete" if status else "Pending",
            "Comment": comments_map.get(key.replace("::row::", "::row::") + "::comment", "")
        })

    # Add section-level comments
    for key, comment in comments_map.items():
        if key.startswith("sec_") and "::tbl::" not in key:
            progress_list.append({
                "Section_Key": key,
                "Task_Key": key,
                "Status": "N/A (Section Comment)",
                "Comment": comment
            })

    df_current = pd.DataFrame(progress_list)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write current playbook data first
        sheet_name_current = re.sub(r'[^\w\-_\s]', '_', selected_playbook.replace('.docx', ''))[:31]
        if not sheet_name_current: sheet_name_current = "Current_Playbook"
        df_current.to_excel(writer, sheet_name=sheet_name_current, index=False)
        
        # Write bulk export data
        if bulk_export:
            for pb in all_playbooks:
                if pb != selected_playbook:
                    comp, comm, _ = load_progress(pb)
                    
                    bulk_list = []
                    for key, status in comp.items():
                        bulk_list.append({
                            "Task_Key": key,
                            "Status": "Complete" if status else "Pending",
                            "Comment": comm.get(key.replace("::row::", "::row::") + "::comment", "")
                        })
                    
                    df_pb = pd.DataFrame(bulk_list)
                    sheet_name = re.sub(r'[^\w\-_\s]', '_', pb.replace('.docx', ''))[:31]
                    if not sheet_name: sheet_name = "PB_Data"
                    # Ensure sheet name is unique and valid (max 31 chars)
                    sheet_name = sheet_name[:25] + str(hashlib.sha1(pb.encode()).hexdigest()[:6])
                    
                    df_pb.to_excel(writer, sheet_name=sheet_name, index=False)
                    
    return output.getvalue()

@st.cache_data
def export_to_csv(completed_map: Dict, comments_map: Dict, selected_playbook: str) -> bytes:
    """Exports data to CSV format."""
    progress_list = []
    
    # Process all completed tasks first
    for key, status in completed_map.items():
        # Find corresponding comment key for action rows
        comment_key = key.replace("::row::", "::row::") + "::comment" if "::tbl::" in key else key
        comment = comments_map.get(comment_key, "")
        progress_list.append({
            "Task_Key": key,
            "Status": "Complete" if status else "Pending",
            "Comment": comment
        })

    # Add section-level comments that aren't tied to a task
    existing_keys = set(completed_map.keys())
    for key, comment in comments_map.items():
        if key not in existing_keys: # Only include if not already added via completed_map loop
             # Check if it's a section comment (starts with sec_ but isn't a task key)
            if key.startswith("sec_") and "::tbl::" not in key:
                 progress_list.append({
                    "Task_Key": key,
                    "Status": "N/A (Section Comment)",
                    "Comment": comment
                })

    df = pd.DataFrame(progress_list)
    return df.to_csv(index=False).encode('utf-8')

# === PLAYBOOK PARSING (CM-4) ===
@st.cache_data(hash_funcs={Path: lambda p: str(p)})
def parse_playbook_cached(path: str) -> List[Dict[str, Any]]:
    """
    Parses a DOCX file into a structured list of sections, content, and tables.
    Note: This is highly dependent on the structure of the source DOCX.
    """
    with open(path, "rb") as fh:
        # Use simple conversion to HTML
        result = mammoth.convert_to_html(fh)
        html = result.value
    soup = BeautifulSoup(html, "html.parser")

    exclude_terms = ["table of contents", "document control", "document revision", "assumptions", "disclaimer", "version history"]
    def excluded(text: str) -> bool:
        if not text:
            return False
        tl = text.strip().lower()
        return any(ex in tl for ex in exclude_terms)

    sections = []
    stack = [] # Stack for tracking hierarchy (H1, H2, H3, etc.)

    # Process tags to build hierarchical structure
    for tag in soup.find_all(['h1','h2','h3','h4','p','table','img']):
        if tag.name.startswith('h') and tag.name[1:].isdigit():
            title = tag.get_text().strip()
            if excluded(title):
                continue
            level = int(tag.name[1])
            node = {"title": title, "level": level, "content": [], "subs": []}
            
            # Pop off nodes from the stack that are the same level or higher
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
            # Extract table rows and cells
            rows = [[td.get_text(separator="\n").strip() for td in tr.find_all(["td","th"])] for tr in tag.find_all("tr")]
            if rows and stack:
                stack[-1]["content"].append({"type": "table", "value": rows})

    # NOTE: The custom table reconstruction logic from text is brittle and highly depends 
    # on input document format. It's kept for functional compatibility.
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
            
            # Check for header or start of numbered list
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
                            desc = " ".join(current_desc_parts).strip()
                            # Simple heuristic to extract owner from last part
                            owner = current_desc_parts.pop() if current_desc_parts and any(p in current_desc_parts[-1].lower() for p in owner_keywords) else ""
                            rows.append([current_ref, current_step, desc, owner])
                            current_desc_parts = []
                        match_obj = ref_pattern.match(txt_j)
                        current_ref = match_obj.group(0)
                        current_step = txt_j[match_obj.end():].strip()
                    else:
                        current_desc_parts.append(txt_j)
                    j += 1
                    
                # Add the last collected row
                if current_ref:
                    desc = " ".join(current_desc_parts).strip()
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

    # Prune empty sections after reconstruction
    def prune(node):
        kept_subs = [sub for sub in node.get("subs", []) if prune(sub)]
        node["subs"] = kept_subs
        # Keep section if it has content OR subs
        return bool(node.get("content")) or bool(kept_subs)

    return [s for s in sections if prune(s)]

# === RENDERING ===
ACTION_HEADERS = {"reference","ref","step","description","ownership","responsibility","owner","responsible"}

def is_action_table(rows: List[List[str]]) -> bool:
    """Heuristically determines if a table is an 'Action' table needing completion tracking."""
    if not rows or not rows[0]:
        return False
    # Check headers
    headers = [h.strip().lower() for h in rows[0] if h.strip()]
    hits = sum(1 for h in headers if any(k in h for k in ACTION_HEADERS))
    
    # Check if the first column of the first data row looks like a reference
    first_data_row = rows[1] if len(rows) > 1 and len(rows[1]) > 0 else []
    first_cell_is_ref = ref_pattern.match(first_data_row[0].strip() if first_data_row else "")
    
    # It's an action table if it has relevant headers OR if it looks like a list of references with 4+ columns
    return hits >= 2 or (len(rows[0]) >= 4 and first_cell_is_ref)

# Global counter for progress, reset on each full rendering cycle
task_counter = {"total": 0, "done": 0}

def render_action_table(playbook_name, sec_key, rows, completed_map, comments_map, autosave, table_index=0):
    """Renders a table with checkboxes for task tracking (CM-4, CM-5)."""
    global task_counter
    
    default_headers = ["Reference", "Step", "Description", "Ownership/Responsibility"]
    # Determine headers: use row 0 if it doesn't look like a data row starting with a ref, otherwise use default
    data_rows = []
    if len(rows) > 0:
        # If the first row's first cell doesn't look like a reference, assume it's a header row
        is_header_row = not ref_pattern.match(rows[0][0].strip() if rows[0] else "")
        if is_header_row:
            headers = rows[0]
            data_rows = rows[1:]
        else:
            headers = default_headers
            data_rows = rows
    else:
        return # No data
        
    # Standardize column count to 4 for tracking fields
    for row in data_rows:
        while len(row) < 4:
            row.append("")
    
    # Count tasks for overall progress
    task_counter["total"] += len(data_rows)

    st.caption("Mark tasks complete and add notes to satisfy the control requirements.")
    
    # Custom HTML for header row
    st.markdown(f"""
    <div class="action-table-header">
        <div class="action-ref">Ref</div>
        <div class="action-step">Step</div>
        <div class="action-desc">Description</div>
        <div class="action-owner">Owner</div>
        <div class="action-done">Done</div>
        <div class="action-comment">Comment</div>
    </div>
    """, unsafe_allow_html=True)
    
    changed = False
    table_key = f"{sec_key}::tbl::{table_index}"

    for ridx, row in enumerate(data_rows):
        row_key = f"{table_key}::row::{ridx}"
        comment_key = f"{row_key}::comment"
        
        # Pull data, using only 4 main columns for consistency
        ref = row[0].strip() if len(row) > 0 else ""
        step = row[1].strip() if len(row) > 1 else ""
        desc = row[2].strip() if len(row) > 2 else ""
        owner = row[3].strip() if len(row) > 3 else ""
        
        # Clean up description if it was parsed strangely
        if len(row) > 4:
             desc += " " + " ".join(row[3:]) # Rejoin any extra columns into description
             owner = row[3].strip() # Use the 4th column as primary owner
             
        prev_val = completed_map.get(row_key, False)
        prev_comment = comments_map.get(comment_key, "")

        cb_key = f"cb_{playbook_name}_{sec_key}_{table_index}_{ridx}"
        ci_key = f"ci_{playbook_name}_{sec_key}_{table_index}_{ridx}"
        
        # Use columns for alignment but rely on custom CSS for layout control
        cols = st.columns([8, 15, 35, 15, 5, 22]) 
        
        # Render data using markdown with custom classes for styling
        cols[0].markdown(f'<div class="action-ref">{ref}</div>', unsafe_allow_html=True)
        cols[1].markdown(f'<div class="action-step">{step}</div>', unsafe_allow_html=True)
        cols[2].markdown(f'<div class="action-desc">{desc}</div>', unsafe_allow_html=True)
        cols[3].markdown(f'<div class="action-owner">{owner}</div>', unsafe_allow_html=True)
        
        # Checkbox for completion
        new_val = cols[4].checkbox("", value=prev_val, key=cb_key, label_visibility="collapsed")
        
        # Text input for comment (label_visibility="collapsed" hides the default label)
        new_comment = cols[5].text_input("", value=prev_comment, key=ci_key, placeholder="Notes/Evidence (AC-6, CM-4)", label_visibility="collapsed")

        # Update logic
        if new_val != prev_val:
            completed_map[row_key] = new_val
            changed = True
            if new_val:
                task_counter["done"] += 1
            else:
                task_counter["done"] -= 1
        if new_comment != prev_comment:
            comments_map[comment_key] = new_comment
            changed = True
            
        # Manually track done count for current session (since this function is run multiple times)
        if new_val and row_key not in st.session_state.get('initial_completed_keys', set()):
            # Only count done tasks that were done *before* this rerun to avoid double counting from cache
            pass
        
        # Restore task_counter['done'] based on the current state if it's already in the completed map
        if new_val and row_key not in completed_map.get('__counted__', {}):
             # Simple flag to track if the task was already counted by previous runs
             completed_map['__counted__'][row_key] = True

    if autosave and changed:
        save_progress(playbook_name, completed_map, comments_map, st.session_state.get("expanders", {}))
        # Important: Rerun to update the progress bar at the top
        st.rerun()

def render_generic_table(rows: List[List[str]]):
    """Renders a standard, non-action table."""
    if len(rows) > 1:
        # Use the first row as header if available
        try:
            df = pd.DataFrame(rows[1:], columns=rows[0])
        except ValueError:
            # Handle case where column count doesn't match data rows
            df = pd.DataFrame(rows)
            st.warning("Table column mismatch. Showing raw data.")
    else:
        df = pd.DataFrame(rows)
    
    # Use st.dataframe for better styling of generic tables
    st.dataframe(df, use_container_width=True, hide_index=True)

def render_section_content(section, playbook_name, completed_map, comments_map, autosave, sec_key, is_sub=False):
    """Recursively renders all content within a playbook section."""
    table_idx = 0
    # 1. Render content (text, image, table)
    for item in section.get("content", []):
        t = item.get("type")
        if t == "text":
            text = item.get("value", "").replace("\n", "<br/>")
            st.markdown(f'<div style="font-size:1.1rem;line-height:1.6;padding-bottom:10px;">{text}</div>', unsafe_allow_html=True)
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
                    
    # 2. Render sub-sections recursively
    for sub in section.get("subs", []):
        # Use st.expander for sub-sections to manage complexity
        sub_key = stable_key(playbook_name, sub["title"], sub["level"])
        
        # Retrieve expander state
        is_expanded = st.session_state.get("expanders", {}).get(sub_key, False)
        
        # Checkbox status for progress tracking visualization
        sub_tasks = get_task_counts_recursive(sub, playbook_name)
        sub_done = sum(1 for k in sub_tasks if completed_map.get(k, False))
        sub_total = len(sub_tasks)
        sub_pct = int((sub_done / sub_total) * 100) if sub_total > 0 else 0
        
        header_markdown = f"""
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <span>{sub['title']}</span>
            <span style="font-size:1rem; color:{'var(--red)' if sub_pct < 100 else 'var(--blue-shadow)'}; font-weight:600;">
                {sub_pct}% Complete ({sub_done}/{sub_total})
            </span>
        </div>
        """
        
        with st.expander(header_markdown, expanded=is_expanded):
            # Record expander state change
            if st.session_state.get(f"exp_{sub_key}") != is_expanded:
                st.session_state.get("expanders", {})[sub_key] = st.session_state.get(f"exp_{sub_key}", is_expanded)
            
            # Recursive call
            render_section_content(sub, playbook_name, completed_map, comments_map, autosave, sub_key, True)
            
    # 3. Add Section-level comment box for the current level (only if not a sub-section)
    if not is_sub:
        st.markdown("<div style='font-weight:700;margin-top:24px;margin-bottom:6px;'>Section Comments / Implementation Notes (CM-4, AC-6)</div>", unsafe_allow_html=True)
        prev_sec_comment = comments_map.get(sec_key, "")
        sec_comment_key = f"sec_cmt_{playbook_name}_{sec_key}"
        
        # Textarea input to capture notes/evidence
        new_sec_comment = st.text_area("", value=prev_sec_comment, key=sec_comment_key, height=120, placeholder="Document evidence, assumptions, or implementation status...", label_visibility="collapsed")
        
        if new_sec_comment != prev_sec_comment:
            comments_map[sec_key] = new_sec_comment
            if autosave:
                # IMPORTANT: Pass all three maps to save_progress
                save_progress(playbook_name, completed_map, comments_map, st.session_state.get("expanders", {}))
                st.rerun() # Rerun to ensure state consistency

def get_task_counts_recursive(section: Dict[str, Any], playbook_name: str) -> List[str]:
    """Helper to recursively find all task keys in a section tree."""
    tasks = []
    sec_key = stable_key(playbook_name, section["title"], section["level"])
    table_idx = 0
    
    for item in section.get("content", []):
        if item.get("type") == "table":
            rows = item.get("value", [])
            if is_action_table(rows):
                # Determine data rows safely (same logic as render_action_table)
                data_rows = []
                if len(rows) > 0:
                    is_header_row = not ref_pattern.match(rows[0][0].strip() if rows[0] else "")
                    data_rows = rows[1:] if is_header_row and len(rows) > 1 else rows
                
                table_key = f"{sec_key}::tbl::{table_idx}"
                for ridx in range(len(data_rows)):
                    tasks.append(f"{table_key}::row::{ridx}")
                table_idx += 1
                
    for sub in section.get("subs", []):
        tasks.extend(get_task_counts_recursive(sub, playbook_name))
        
    return tasks

# === MAIN APPLICATION FLOW ===

def main():
    """The main entry point for the Streamlit application."""
    
    # 1. Theme Selection
    theme = theme_selector()
    
    # 2. Authentication
    user = authenticate()
    
    # 3. Admin Toggle
    if user["role"] == "admin" and st.sidebar.button("Admin Dashboard", key="admin_toggle_btn"):
        st.session_state.admin_page = True
        st.rerun()
        
    if st.session_state.get("admin_page", False):
        admin_dashboard(user)
        return

    # --- Main App Logic ---
    
    # Find all available playbooks
    playbooks = sorted([f for f in os.listdir(PLAYBOOKS_DIR) if f.endswith('.docx')])
    if not playbooks:
        st.error("No playbook (.docx) files found in the 'playbooks' directory. Please upload one via the Admin Dashboard.")
        return
        
    # Sidebar Playbook Selection
    selected_playbook_name = st.sidebar.selectbox("Select Playbook (CM-4)", playbooks, key="playbook_select")
    playbook_path = Path(PLAYBOOKS_DIR) / selected_playbook_name

    # Load data for the selected playbook
    sections = parse_playbook_cached(str(playbook_path))
    completed_map, comments_map, expanders_map = load_progress(selected_playbook_name)
    st.session_state.expanders = expanders_map # Store in session state for dynamic updates
    
    # Reset global task counter before rendering
    task_counter["total"] = 0
    task_counter["done"] = 0
    
    # Calculate initial done tasks from loaded data
    all_task_keys = []
    for section in sections:
        all_task_keys.extend(get_task_counts_recursive(section, selected_playbook_name))

    # Calculate done count based on the loaded map
    initial_done_count = sum(1 for key in all_task_keys if completed_map.get(key, False))
    task_counter["done"] = initial_done_count
    task_counter["total"] = len(all_task_keys)
    
    total_tasks = task_counter["total"]
    done_tasks = task_counter["done"]
    pct_complete = int((done_tasks / total_tasks) * 100) if total_tasks > 0 else 0
    
    # --- Sticky Header ---
    st.markdown(f"""
    <div class="sticky-header">
        {get_logo()}
        <div class="app-title text-left md:text-center flex-grow">
            <span class="nist-text">NIST <sup class="text-xs">CM-4</sup></span> Playbook Tracker
            <div class="text-sm font-normal text-gray-500 mt-1">
                <span class="font-bold">{selected_playbook_name}</span> for Joval Wines | User: {user['name']} ({user['role']})
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # --- Progress Bar and Stats ---
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 3])

    with col1:
        st.metric(label="Total Tasks", value=total_tasks)
    with col2:
        st.metric(label="Completed Tasks", value=done_tasks)
        
    with col3:
        badges = calculate_badges(pct_complete)
        st.markdown(f"<div style='font-size:1rem; font-weight:600; color:var(--text);'>Completion Progress: {pct_complete}%</div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class="progress-wrap">
            <div class="progress-fill" style="width:{pct_complete}%;"></div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"<div style='font-size:0.8rem; color:var(--gold); font-weight:500;'>Current Badge: {badges[0]}</div>", unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- Sidebar Table of Contents (TOC) ---
    st.sidebar.markdown("---")
    st.sidebar.markdown('<div class="section-title !text-xl !font-bold">Table of Contents</div>', unsafe_allow_html=True)
    
    # Search filter for TOC
    search_term = st.sidebar.text_input("Search Sections...", key="toc_search", placeholder="Type to filter...").lower()
    
    # --- Main Content Area ---
    st.markdown('<div class="content-wrap">', unsafe_allow_html=True)
    
    autosave = st.checkbox("Enable Autosave (Recommended)", value=True, key="autosave_cb")
    
    if not autosave:
        if st.button("Manually Save Progress"):
            save_progress(selected_playbook_name, completed_map, comments_map, st.session_state.get("expanders", {}))
            st.success("Progress saved!")
            st.rerun()

    # Create the TOC and render content simultaneously
    for section in sections:
        sec_key = stable_key(selected_playbook_name, section["title"], section["level"])
        
        # TOC Link (only if search matches)
        if not search_term or search_term in section["title"].lower():
            st.sidebar.markdown(
                f'<a href="#{sec_key}" class="toc-item">{section["title"]}</a>', 
                unsafe_allow_html=True
            )
            
        # Content Rendering - Use st.expander for top-level sections
        # Initial expansion based on loaded state
        is_expanded = st.session_state.get("expanders", {}).get(sec_key, True) 

        # Calculate progress for top-level section
        sec_tasks = get_task_counts_recursive(section, selected_playbook_name)
        sec_done = sum(1 for k in sec_tasks if completed_map.get(k, False))
        sec_total = len(sec_tasks)
        sec_pct = int((sec_done / sec_total) * 100) if sec_total > 0 else 0
        
        header_markdown = f"""
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <span><a name="{sec_key}" class="section-title" style="text-decoration:none;">{section['title']}</a></span>
            <span style="font-size:1.2rem; color:{'var(--red)' if sec_pct < 100 else 'var(--blue-shadow)'}; font-weight:600;">
                {sec_pct}%
            </span>
        </div>
        """
        
        # Use an expander for the main section container
        with st.expander(header_markdown, expanded=is_expanded, key=f"exp_{sec_key}"):
            # Update expander state in session state
            st.session_state.get("expanders", {})[sec_key] = st.session_state.get(f"exp_{sec_key}")
            
            # Render the content and sub-sections
            render_section_content(section, selected_playbook_name, completed_map, comments_map, autosave, sec_key)

    st.markdown('</div>', unsafe_allow_html=True)
    
    # --- Export and Feedback Section (Bottom) ---
    st.markdown("---")
    st.markdown("### Reporting & Feedback")
    
    exp_col1, exp_col2 = st.columns(2)
    
    with exp_col1:
        csv_data = export_to_csv(completed_map, comments_map, selected_playbook_name)
        st.download_button(
            label="Download Current Playbook Progress (CSV)",
            data=csv_data,
            file_name=f"{selected_playbook_name.replace('.docx', '')}_progress_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            key="csv_export_btn"
        )
        if OPENPYXL_AVAILABLE:
            excel_data = export_to_excel(completed_map, comments_map, selected_playbook_name, playbooks, bulk_export=False)
            st.download_button(
                label="Download Current Playbook Progress (Excel)",
                data=excel_data,
                file_name=f"{selected_playbook_name.replace('.docx', '')}_progress_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="excel_export_btn"
            )
        else:
             st.info("Install 'openpyxl' to enable Excel export.")

    with exp_col2:
        show_feedback(user["email"])
        
    st.sidebar.markdown("---")
    st.sidebar.markdown(f'<div class="text-xs text-gray-400">App Version 1.1 | Total Tasks: {total_tasks}</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
