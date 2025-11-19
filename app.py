import streamlit as st
import pandas as pd
import sqlite3
import hashlib
import os
import smtplib
import time
import random
import urllib.request
from datetime import datetime
from io import BytesIO

# ReportLab Imports for PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

# Plotly for visualizations
import plotly.graph_objects as go
import plotly.express as px

# Email imports
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# 1. CONFIGURATION & NIST MAPPING
# ==========================================
DB_FILE = "joval_portal.db"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "joval.risk.portal@gmail.com"
SENDER_PASSWORD = "your_app_password_here" 

# --- NIST Cybersecurity Framework (CSF) Functions ---
NIST_FUNCTIONS = {
    "IDENTIFY": "Develop an organizational understanding to manage cybersecurity risk to systems, assets, data, and capabilities.",
    "PROTECT": "Develop and implement safeguards to ensure delivery of critical infrastructure services.",
    "DETECT": "Develop and implement activities to identify the occurrence of a cybersecurity event.",
    "RESPOND": "Develop and implement activities to take action regarding a detected cybersecurity incident.",
    "RECOVER": "Develop and implement activities to maintain plans for resilience and to restore any capabilities or services impaired due to a cybersecurity incident."
}

# ==========================================
# 2. DATABASE ENGINE (Stable Direct Connection)
# ==========================================
def get_connection():
    # Streamlit runs in a multi-threaded environment, but sqlite3 defaults to 
    # check_same_thread=True, which causes issues. Setting it to False allows threads 
    # to use the same connection, which is okay for a simple single-file app.
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def run_query(query, params=None, is_write=False):
    conn = get_connection()
    try:
        c = conn.cursor()
        if is_write:
            c.execute(query, params or ())
            conn.commit()
            return c.rowcount
        else:
            # Use pandas to read dataframes easily
            return pd.read_sql(query, conn, params=params)
    except Exception as e:
        # In a real app, this should be logged to an external system
        print(f"Database Error: {e}")
        return 0 if is_write else pd.DataFrame()
    finally:
        conn.close()

def init_db():
    """Initializes the database schema and seeds initial data."""
    conn = get_connection()
    c = conn.cursor()
    
    # Core Tables
    c.execute("""CREATE TABLE IF NOT EXISTS companies (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE)""")
    c.execute("""CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, email TEXT UNIQUE, password TEXT, role TEXT, company_id INTEGER)""")
    
    # Risks Table with Enterprise Columns (remediation, jira, NIST categories)
    c.execute("""CREATE TABLE IF NOT EXISTS risks (
        id INTEGER PRIMARY KEY AUTOINCREMENT, company_id INTEGER, title TEXT, description TEXT, 
        category TEXT, likelihood TEXT, impact TEXT, status TEXT, submitted_by TEXT, submitted_date TEXT, 
        risk_score INTEGER, approver_email TEXT, approver_notes TEXT, approved_by TEXT, approved_date TEXT, 
        workflow_step TEXT, remediation_plan TEXT, remediation_owner TEXT, remediation_date TEXT, jira_ticket_id TEXT
    )""")
    
    # Evidence Vault Table
    c.execute("""CREATE TABLE IF NOT EXISTS evidence (id INTEGER PRIMARY KEY AUTOINCREMENT, risk_id INTEGER, company_id INTEGER, file_name TEXT, upload_date TEXT, uploaded_by TEXT, file_data BLOB)""")
    
    # Vendor Tables with Scoring
    c.execute("""CREATE TABLE IF NOT EXISTS vendors (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, contact_email TEXT, risk_level TEXT, last_assessment_date TEXT, company_id INTEGER, vendor_score INTEGER)""")
    c.execute("""CREATE TABLE IF NOT EXISTS vendor_questionnaire (id INTEGER PRIMARY KEY AUTOINCREMENT, vendor_id INTEGER, question_id INTEGER, answer TEXT, answered_date TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS vendor_questions (id INTEGER PRIMARY KEY AUTOINCREMENT, question TEXT, weight INTEGER, company_id INTEGER)""")
    
    # Audit Trail
    c.execute("""CREATE TABLE IF NOT EXISTS audit_trail (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, user_email TEXT, action TEXT, details TEXT)""")

    # Safe Column Migration (Ensure all necessary columns exist)
    # This prevents errors when running the script after the initial creation
    cols = [("risks", "remediation_plan", "TEXT"), ("risks", "jira_ticket_id", "TEXT"), 
            ("vendors", "vendor_score", "INTEGER"), ("risks", "remediation_owner", "TEXT"), 
            ("risks", "remediation_date", "TEXT"), ("vendor_questions", "weight", "INTEGER")]
    for t, c_name, d_type in cols:
        try: c.execute(f"ALTER TABLE {t} ADD COLUMN {c_name} {d_type}")
        except: pass

    # Seed Data (Admin/Approvers/Users)
    c.execute("SELECT count(*) FROM companies")
    if c.fetchone()[0] == 0:
        companies = ["Joval Wines", "Joval Family Wines", "BNV", "BAM"]
        for comp in companies: c.execute("INSERT OR IGNORE INTO companies (name) VALUES (?)", (comp,))
        
        c.execute("SELECT id, name FROM companies")
        rows = c.fetchall()
        # Default password 'Joval2025'
        hashed = hashlib.sha256("Joval2025".encode()).hexdigest() 
        
        for cid, cname in rows:
            # Create a dedicated Approver for each company
            c.execute("INSERT OR IGNORE INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", 
                      (f"approver_{cname.lower().replace(' ', '')}", f"approver@{cname.lower().replace(' ', '')}.com.au", hashed, "Approver", cid))
            # Create a standard User for Joval Wines
            if cname == "Joval Wines":
                c.execute("INSERT OR IGNORE INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", 
                          ("user_joval", "user@jovalwines.com.au", hashed, "User", cid))
            # Create Global Admin
            if cname == "Joval Wines":
                c.execute("INSERT OR IGNORE INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", 
                          ("admin", "admin@jovalwines.com.au", hashed, "Admin", cid))
    
    # Seed Questions (NIST aligned - relevant to vendor risk)
    c.execute("SELECT count(*) FROM vendor_questions")
    if c.fetchone()[0] == 0:
        qs = [("Does the vendor provide visibility into their security controls? (IDENTIFY)", 10), 
              ("Is data processed by the vendor encrypted at rest and in transit? (PROTECT)", 15), 
              ("Does the vendor have a SOC/monitoring function? (DETECT)", 10), 
              ("Does the vendor have a documented Incident Response Plan? (RESPOND)", 10)]
        for q, w in qs: c.execute("INSERT INTO vendor_questions (question, weight, company_id) VALUES (?, ?, ?)", (q, w, 1))

    conn.commit()
    conn.close()

# Run initialization (will only create tables and seed data if DB file doesn't exist or is empty)
if not os.path.exists(DB_FILE):
    init_db()
else:
    init_db() # Run init_db again to ensure all columns/migrations are applied

# ==========================================
# 3. HELPERS (Including Audit Logging and AI)
# ==========================================
def calculate_risk_score(likelihood, impact):
    """Calculates risk score based on L x I (1=Low, 2=Medium, 3=High)."""
    scores = {"Low": 1, "Medium": 2, "High": 3}
    return scores.get(likelihood, 1) * scores.get(impact, 1)

def get_risk_color(score):
    """Determines risk severity based on score for UI coloring."""
    if score >= 7: return "red"
    elif score >= 4: return "orange"
    else: return "green"

# --- AUDIT TRAIL LOGGING ---
def log_action(user_email, action, details=""):
    """Records user actions into the audit_trail table."""
    run_query("INSERT INTO audit_trail (timestamp, user_email, action, details) VALUES (?, ?, ?, ?)", 
              (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user_email, action, details), is_write=True)

# --- AI AUTO-CLASSIFY (Simulation) ---
def ai_classify_risk(text):
    """Simulates an AI model classifying risk based on keywords."""
    text = text.lower()
    if any(x in text for x in ["phish", "email", "scam"]): return "DETECT", "High", "High"
    if any(x in text for x in ["malware", "virus", "patch"]): return "PROTECT", "Medium", "High"
    if any(x in text for x in ["backup", "loss"]): return "RECOVER", "Medium", "Medium"
    if any(x in text for x in ["asset", "inventory", "policy"]): return "IDENTIFY", "Low", "Medium"
    return "IDENTIFY", "Low", "Low"

def generate_pdf_report(title, content):
    """Generates a PDF report using ReportLab for download."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.2*inch)
    styles = getSampleStyleSheet()
    story = []
    
    # Joval Wines Logo Placeholder
    try:
        # Using a public image URL as a placeholder
        with urllib.request.urlopen("https://jovalwines.com.au/wp-content/uploads/2020/06/Joval-Wines-Logo.png") as r:
            img = Image(BytesIO(r.read()), width=2*inch, height=0.6*inch)
            story.append(img)
    except Exception:
        story.append(Paragraph("JOVAL WINES Risk Report", styles['Title']))

    story.append(Spacer(1, 12))
    story.append(Paragraph(title, styles['Heading1']))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%d %B %Y')}", styles['Normal']))
    story.append(Spacer(1, 20))

    if isinstance(content, pd.DataFrame):
        # Format for display in PDF (truncate long strings)
        content_str = content.astype(str).applymap(lambda x: (x[:50] + '...') if len(x) > 50 else x)
        data = [content.columns.tolist()] + content_str.values.tolist()
        
        # Define table style
        table_style = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1a1a1a')), 
            ('TEXTCOLOR',(0,0),(-1,0),colors.white), 
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey), 
            ('FONTSIZE', (0,0), (-1,-1), 8),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ])
        table = Table(data, repeatRows=1)
        table.setStyle(table_style)
        story.append(table)
        
    doc.build(story)
    buffer.seek(0)
    return buffer

def send_email(to_email, subject, body):
    """
    Simulated email function. Requires valid SMTP setup to work in a real environment.
    (Note: This function is currently a simulation as credentials are not provided.)
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        # --- Simulation Block ---
        # st.success(f"Email SIMULATION sent to {to_email} with subject: {subject}")
        # print(f"Email SIMULATION: To={to_email}, Subject={subject}, Body={body}")
        # --- End Simulation Block ---
        
        # Uncomment the block below and replace SENDER_PASSWORD with a valid app password 
        # for a real email solution (e.g., Gmail App Passwords)
        # server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        # server.starttls()
        # server.login(SENDER_EMAIL, SENDER_PASSWORD) 
        # server.send_message(msg)
        # server.quit()
        return True
    except Exception: 
        # print(f"Email failed: {e}")
        return False
    
# ==========================================
# 4. UI SETUP
# ==========================================
st.set_page_config(page_title="Joval Risk Portal", layout="wide")
st.markdown("""
<style>
    /* Custom CSS for a professional look */
    .header {background: #1a1a1a; color: white; padding: 2rem; text-align: center; border-radius: 12px 12px 0 0;}
    .metric-card {background: white; padding: 1.5rem; border-radius: 12px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border: 1px solid #f0f0f0;}
    .clickable-risk {cursor: pointer; padding: 0.75rem; border-radius: 8px; margin: 0.25rem 0; border: 1px solid #ddd; transition: background-color 0.2s;}
    .clickable-risk:hover {background-color: #f0f0f0 !important;}
    .approval-badge {background: #e6f7ff; padding: 0.25rem 0.5rem; border-radius: 12px; font-size: 0.8rem;}
    /* NIST hints style */
    .nist-hint {background-color: #f0f8ff; border-left: 5px solid #007bff; padding: 10px; border-radius: 4px; margin-bottom: 15px;}
</style>""", unsafe_allow_html=True)
st.markdown('<div class="header"><h1>JOVAL WINES</h1><p>NIST-Aligned Risk Management Portal</p></div>', unsafe_allow_html=True)

# ==========================================
# 5. LOGIN
# ==========================================
if "user" not in st.session_state:
    with st.sidebar:
        st.markdown("### Login")
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.button("Login"):
            hashed = hashlib.sha256(password.encode()).hexdigest()
            users = run_query("SELECT * FROM users WHERE username=? AND password=?", (username, hashed))
            if not users.empty:
                st.session_state.user = users.iloc[0].to_list()
                log_action(st.session_state.user[2], "LOGIN", "User logged in successfully.")
                st.rerun()
            else:
                st.error("Invalid credentials")
    st.stop()

user = st.session_state.user
comp_df = run_query("SELECT name FROM companies WHERE id=?", (user[5],))
user_company_name = comp_df.iloc[0]['name'] if not comp_df.empty else "Unknown"
# Default password for resets: 'Reset2025'
DEFAULT_PASSWORD_HASH = hashlib.sha256("Reset2025".encode()).hexdigest()

# ==========================================
# 6. SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("### User & Company")
    st.markdown(f"**{user[1]}** • {user_company_name} ({user[4]})")
    st.markdown("---")
    
    st.markdown("### Navigation")
    pages = ["Dashboard", "Log a new Risk", "Evidence Vault", "Vendor Management", "Reports"]
    if user[4] == "Approver": pages.insert(1, "My Approvals")
    if user[4] == "Admin": pages += ["Audit Trail", "Admin Panel"]
    
    for p in pages:
        if st.button(p, key=f"nav_{p}", use_container_width=True):
            st.session_state.page = p
            st.rerun()
            
    st.markdown("---")
    st.markdown("### External Links")
    # FIX 1: Updated Playbook Tracker link
    st.markdown("[**Playbook Tracker App**](https://joval-wines-nist-playbook-tracker.streamlit.app/)")
    
    if st.button("Logout", key="nav_logout", use_container_width=True):
        log_action(user[2], "LOGOUT")
        st.session_state.clear()
        st.rerun()

page = st.session_state.get("page", "Dashboard")

# ==========================================
# 7. DASHBOARD
# ==========================================
if page == "Dashboard":
    st.markdown("## Company Risk Dashboard")
    view_company_id = user[5]
    
    # Admin can select which company's dashboard to view
    if user[4] == "Admin":
        all_comps = run_query("SELECT id, name FROM companies")
        comp_opts = all_comps['name'].tolist()
        try: def_idx = comp_opts.index(user_company_name)
        except: def_idx = 0
        view_name = st.selectbox("Viewing Dashboard For:", comp_opts, index=def_idx)
        view_company_id = all_comps[all_comps['name'] == view_name].iloc[0]['id']
    else:
        st.info(f"Viewing risks for **{user_company_name}**.")
    
    risks = run_query("SELECT * FROM risks WHERE company_id=?", (view_company_id,))
    
    col1, col2, col3 = st.columns(3)
    
    if not risks.empty:
        high_risks = risks[risks['risk_score'] >= 7]
        medium_risks = risks[(risks['risk_score'] >= 4) & (risks['risk_score'] < 7)]
        pending_risks = risks[risks['status'] == 'Pending Approval']
    else:
        high_risks, medium_risks, pending_risks = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    with col1:
        st.markdown(f'<div class="metric-card" style="border-color: red;"><h2>{len(high_risks)}</h2><p>High Risks</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card" style="border-color: orange;"><h2>{len(medium_risks)}</h2><p>Medium Risks</p></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card" style="border-color: #007bff;"><h2>{len(pending_risks)}</h2><p>Pending Approvals</p></div>', unsafe_allow_html=True)
    
    st.write("---")
    st.markdown("### Risk Register Summary")
    
    if risks.empty:
        st.info("No risks logged for this company.")
    else:
        sorted_risks = risks.sort_values(by='risk_score', ascending=False)
        for _, r in sorted_risks.iterrows():
            color = get_risk_color(r['risk_score'])
            bg = "#ffebeb" if color == "red" else "#fff7eb" if color == "orange" else "#ebfff0"
            btn_label = f"**{r['title']}** | Score: **{r['risk_score']}** | Status: **{r['status']}**"
            
            # Using st.button inside a markdown div for custom styling and click handling
            if st.button(btn_label, key=f"r_{r['id']}", use_container_width=True):
                st.session_state.selected_risk = r['id']
                st.session_state.page = "Risk Detail"
                st.rerun()
            
            # Additional details below the button
            st.markdown(f'<div class="clickable-risk" style="background:{bg};"><small>NIST: {r["category"]} | Approver: {r["approver_email"]} | Submitted: {r["submitted_date"]}</small></div>', unsafe_allow_html=True)

# ==========================================
# 8. LOG RISK (with AI and NIST Guidance)
# ==========================================
elif page == "Log a new Risk":
    st.markdown("## Log a New Risk")
    st.info("Aligning risks with the NIST CSF (Identify, Protect, Detect, Respond, Recover) ensures comprehensive coverage.")
    
    comps = run_query("SELECT id, name FROM companies")
    try: def_idx = comps['name'].tolist().index(user_company_name)
    except: def_idx = 0
    c_name = st.selectbox("Company", comps['name'].tolist(), index=def_idx)
    c_id = comps[comps['name'] == c_name].iloc[0]['id']
    
    # FIX 2: Approver list is dynamic and functional
    approvers = run_query("SELECT email, username FROM users WHERE role='Approver' AND company_id=?", (c_id,))
    if approvers.empty:
        st.error(f"No Approver found for {c_name}. Risks cannot be submitted for approval.")
        st.stop()
    
    approver_usernames = approvers['username'].tolist()
    
    with st.form("new_risk"):
        col_a, col_b = st.columns([3,1])
        with col_a: title = st.text_input("Title")
        with col_b: 
            st.write("")
            st.write("")
            run_ai = st.form_submit_button("✨ Auto-Classify (NIST/Score)")
            
        desc = st.text_area("Description")
        
        # AI Defaults
        d_cat, d_lik, d_imp = "IDENTIFY", "Low", "Low"
        if run_ai and (title or desc):
            d_cat, d_lik, d_imp = ai_classify_risk(title + " " + desc)
            st.success(f"AI Suggestion Applied: Category={d_cat}, Likelihood={d_lik}, Impact={d_imp}")
            
        cat = st.selectbox("NIST CSF Category", list(NIST_FUNCTIONS.keys()), index=list(NIST_FUNCTIONS.keys()).index(d_cat))
        
        # NIST Guidance Hint (Added for compliance context)
        st.markdown(f'<div class="nist-hint">**{cat}:** {NIST_FUNCTIONS.get(cat, "No specific guidance.")}</div>', unsafe_allow_html=True)
        
        col_lik, col_imp = st.columns(2)
        with col_lik:
            lik = st.selectbox("Likelihood", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(d_lik))
        with col_imp:
            imp = st.selectbox("Impact", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(d_imp))
        
        selected_approver_name = st.selectbox("Assign to Approver", approver_usernames)
        approver_email = approvers[approvers['username'] == selected_approver_name].iloc[0]['email']

        if st.form_submit_button("Submit Risk for Approval"):
            if not title or not desc:
                st.error("Required fields missing.")
            else:
                score = calculate_risk_score(lik, imp)
                run_query("""INSERT INTO risks (company_id, title, description, category, likelihood, impact, status, submitted_by, submitted_date, risk_score, approver_email, workflow_step) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                          (c_id, title, desc, cat, lik, imp, "Pending Approval", user[2], datetime.now().strftime("%Y-%m-%d"), score, approver_email, "awaiting"), is_write=True)
                
                # Send email notification to approver
                send_email(approver_email, f"NEW Risk Awaiting Approval: {title}", 
                           f"Dear {selected_approver_name},\n\nA new risk titled '{title}' has been submitted and requires your urgent review and approval.\n\nThank you,\nJoval Risk Portal")
                
                st.success(f"Risk Logged and sent to {selected_approver_name} for approval.")
                log_action(user[2], "RISK SUBMITTED", f"{title} (Approver: {approver_email})")
                st.session_state.page = "Dashboard"
                st.rerun()

# ==========================================
# 9. APPROVER WORKFLOW (FIX 7, 8)
# ==========================================
elif page == "My Approvals" and user[4] == "Approver":
    st.markdown("## My Assigned Risks Awaiting Approval")
    
    # Query to fetch risks assigned to the current approver's email (FIX 7: Shows assigned risks)
    pend = run_query("SELECT id, title, risk_score, submitted_date, submitted_by FROM risks WHERE approver_email=? AND status='Pending Approval'", (user[2],))
    
    if pend.empty:
        st.success("You have no pending risks requiring your approval at this time. All caught up!")
    else:
        st.warning(f"🔔 You have **{len(pend)}** urgent risk(s) awaiting review. Please take action.")
        for _, r in pend.iterrows():
            if st.button(f"**Review: {r['title']}** (Score: {r['risk_score']}, Submitted by: {r['submitted_by']})", key=r['id'], use_container_width=True):
                st.session_state.selected_risk = r['id']
                st.session_state.page = "Risk Detail"
                st.rerun()

# ==========================================
# 10. RISK DETAIL (with Remediation & Jira)
# ==========================================
elif page == "Risk Detail" and "selected_risk" in st.session_state:
    rid = st.session_state.selected_risk
    risk = run_query("SELECT * FROM risks WHERE id=?", (rid,)).iloc[0]
    
    st.markdown(f"## Risk: {risk['title']}")
    
    col_s, col_j = st.columns(2)
    with col_s:
        st.markdown(f"NIST Category: **{risk['category']}** | Current Score: **{risk['risk_score']}** | Status: **{risk['status']}**")
    with col_j:
        # Jira Sync Button
        if not risk['jira_ticket_id']:
            if st.button("🔗 Sync to Jira"):
                # Simulate Jira integration
                time.sleep(0.5) 
                ticket = f"JIRA-{random.randint(1000,9999)}"
                run_query("UPDATE risks SET jira_ticket_id=? WHERE id=?", (ticket, rid), is_write=True)
                log_action(user[2], "JIRA SYNC", f"Risk {rid} linked to {ticket}")
                st.success(f"Created and linked to Jira Ticket: {ticket}")
                st.rerun()
        else:
            st.info(f"Jira Ticket: **{risk['jira_ticket_id']}**")

    t1, t2, t3 = st.tabs(["Details & Approval", "Remediation Plan", "Evidence Upload"])
    
    with t1:
        with st.form("edit_details"):
            c1, c2 = st.columns(2)
            new_title = st.text_input("Title", risk['title'])
            new_desc = st.text_area("Description", risk['description'])
            with c1:
                new_cat = st.selectbox("NIST CSF Category", list(NIST_FUNCTIONS.keys()), index=list(NIST_FUNCTIONS.keys()).index(risk['category']))
                new_lik = st.selectbox("Likelihood", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(risk['likelihood']))
            with c2:
                new_imp = st.selectbox("Impact", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(risk['impact']))
                
                # Approval/Status workflow control
                new_status = risk['status']
                if user[4] in ["Admin", "Approver"]:
                    st.markdown("---")
                    st.markdown("#### Approval/Status Control")
                    new_status = st.selectbox("Status", ["Pending Approval", "Approved", "Rejected", "Mitigated"], index=["Pending Approval", "Approved", "Rejected", "Mitigated"].index(risk['status']))
                    new_notes = st.text_area("Approver Notes", risk['approver_notes'] or "")
                else:
                    st.info(f"Current Status: {risk['status']} (Contact Approver: {risk['approver_email']})")
                    new_notes = risk['approver_notes']
                    
            if st.form_submit_button("Save Changes"):
                new_score = calculate_risk_score(new_lik, new_imp)
                approved_by = risk['approved_by']
                approved_date = risk['approved_date']
                
                # Logic for approval action
                is_approval_action = (risk['status'] == 'Pending Approval' and new_status in ['Approved', 'Rejected'] and user[4] in ["Admin", "Approver"])
                
                if is_approval_action:
                    approved_by = user[1]
                    approved_date = datetime.now().strftime("%Y-%m-%d")
                    st.success(f"Risk officially {new_status} by {approved_by}.")

                run_query("""UPDATE risks SET title=?, description=?, category=?, likelihood=?, impact=?, status=?, risk_score=?, approver_notes=?, approved_by=?, approved_date=? WHERE id=?""", 
                          (new_title, new_desc, new_cat, new_lik, new_imp, new_status, new_notes, approved_by, approved_date, rid), is_write=True)
                log_action(user[2], "RISK UPDATED", f"Risk {rid}. Status: {new_status}")
                st.success("Risk Details Updated")
                st.rerun()

    with t2:
        st.markdown("### Remediation and Mitigation Plan (NIST RESPOND & RECOVER)")
        with st.form("rem_plan"):
            plan = st.text_area("Action Plan", risk['remediation_plan'] or "")
            own = st.text_input("Remediation Owner", risk['remediation_owner'] or "")
            date = st.text_input("Target Date (YYYY-MM-DD)", risk['remediation_date'] or "")
            if st.form_submit_button("Update Remediation Plan"):
                run_query("UPDATE risks SET remediation_plan=?, remediation_owner=?, remediation_date=? WHERE id=?", (plan, own, date, rid), is_write=True)
                log_action(user[2], "REMEDIATION PLAN UPDATED", f"Risk {rid}. Owner: {own}")
                st.success("Remediation Plan Updated")
                st.rerun()

    with t3:
        st.markdown("### Upload Supporting Evidence")
        uploaded = st.file_uploader("Upload supporting document/screenshot")
        if uploaded:
            run_query("INSERT INTO evidence (risk_id, company_id, file_name, upload_date, uploaded_by, file_data) VALUES (?, ?, ?, ?, ?, ?)", 
                      (rid, risk['company_id'], uploaded.name, datetime.now().strftime("%Y-%m-%d"), user[1], uploaded.getvalue()), is_write=True)
            log_action(user[2], "EVIDENCE UPLOADED", f"Risk {rid}. File: {uploaded.name}")
            st.success("File uploaded and secured in the vault.")
            st.rerun()
            
        st.markdown("### Existing Evidence")
        files = run_query("SELECT id, file_name, uploaded_by, upload_date, file_data FROM evidence WHERE risk_id=?", (rid,))
        if files.empty:
            st.info("No evidence files uploaded for this specific risk.")
        else:
            for _, f in files.iterrows():
                st.download_button(f"Download {f['file_name']} (by {f['uploaded_by']} on {f['upload_date']})", f['file_data'], f['file_name'], key=f"d_{f['id']}", use_container_width=True)

    if st.button("← Back to Dashboard"):
        del st.session_state.selected_risk
        st.session_state.page = "Dashboard"
        st.rerun()

# ==========================================
# 11. VENDOR MANAGEMENT (NIST & Scoring)
# ==========================================
elif page == "Vendor Management":
    st.markdown("## Vendor Management (NIST CSF Aligned)")
    st.info("Assessing third-party risk based on NIST CSF functions (IDENTIFY, PROTECT, DETECT, RESPOND).")
    t1, t2, t3 = st.tabs(["Vendors List", "Question Template", "Risk Assessment"])
    
    with t1:
        with st.expander("Add New Vendor"):
            with st.form("add_v"):
                n = st.text_input("Vendor Name")
                e = st.text_input("Contact Email")
                if st.form_submit_button("Add Vendor"):
                    run_query("INSERT INTO vendors (name, contact_email, risk_level, company_id, vendor_score) VALUES (?, ?, ?, ?, ?)", (n, e, "Pending", user[5], 0), is_write=True)
                    log_action(user[2], "VENDOR ADDED", n)
                    st.success(f"Vendor {n} added.")
                    st.rerun()
        vendors = run_query("SELECT id, name, contact_email, risk_level, vendor_score, last_assessment_date FROM vendors WHERE company_id=?", (user[5],))
        st.dataframe(vendors.rename(columns={'vendor_score': 'Score', 'risk_level': 'Level'}), use_container_width=True)
        
    with t2:
        st.markdown("### Vendor Assessment Questions (NIST Aligned)")
        qs = run_query("SELECT id, question, weight FROM vendor_questions")
        
        # Max score is dynamic based on current weights
        max_s = qs['weight'].sum() if not qs.empty else 0
        st.info(f"Total Max Possible Score: {max_s}")

        edited = st.data_editor(qs, num_rows="dynamic", column_config={"question": st.column_config.TextColumn("Question (NIST Alignment)")}, use_container_width=True)
        
        if st.button("Save Question Template"):
            # Clear existing questions before saving new set for the company (assuming Company ID 1 is the master template)
            run_query("DELETE FROM vendor_questions WHERE company_id=?", (1,), is_write=True)
            for _, r in edited.iterrows():
                # Ensure weight is a positive integer
                weight = int(r.get('weight', 0)) if str(r.get('weight')).isdigit() and int(r.get('weight', 0)) > 0 else 1 
                if r['question']:
                    run_query("INSERT INTO vendor_questions (question, weight, company_id) VALUES (?, ?, ?)", (r['question'], weight, 1), is_write=True)
            log_action(user[2], "VENDOR TEMPLATE EDITED")
            st.success("Vendor assessment template saved.")
            st.rerun()

    with t3:
        v_list = run_query("SELECT id, name FROM vendors WHERE company_id=?", (user[5],))
        if v_list.empty:
            st.warning("No vendors to score. Please add a vendor first.")
        else:
            sel_v = st.selectbox("Select Vendor to Assess Risk", v_list['name'].tolist())
            vid = v_list[v_list['name'] == sel_v].iloc[0]['id']
            qs = run_query("SELECT * FROM vendor_questions")
            
            with st.form("score"):
                current_score = 0
                max_s = qs['weight'].sum()
                st.markdown(f"#### Security Questionnaire for {sel_v} (Max Score: {max_s})")
                
                # Store answers temporarily
                answers = {}
                for _, q in qs.iterrows():
                    # Radio button value is 'Yes' or 'No'
                    answers[q['id']] = st.radio(f"**{q['question']}** (Weight: {q['weight']})", ["Yes", "No"], key=q['id'], horizontal=True)
                    if answers[q['id']] == "Yes":
                        current_score += q['weight']
                
                if st.form_submit_button("Calculate and Save Score"):
                    # Compliance Logic: High Risk if score is below 50%
                    lvl = "High" if current_score < (max_s * 0.5) else "Low"
                    run_query("UPDATE vendors SET vendor_score=?, risk_level=?, last_assessment_date=? WHERE id=?", (current_score, lvl, datetime.now().strftime("%Y-%m-%d"), vid), is_write=True)
                    log_action(user[2], "VENDOR SCORED", f"Vendor {sel_v}: {current_score}/{max_s}. Level: {lvl}")
                    st.success(f"Vendor Risk Assessed. Score: **{current_score}/{max_s}**. Risk Level: **{lvl}**")
                    st.rerun()

# ==========================================
# 12. EVIDENCE VAULT (FIX 5)
# ==========================================
elif page == "Evidence Vault":
    st.markdown("## Company Evidence Vault")
    st.info("Central, searchable repository for all supporting documents across all risks.")
    
    # FIX 5: Retrieve all evidence for the user's company
    files = run_query("SELECT id, risk_id, file_name, uploaded_by, upload_date, file_data FROM evidence WHERE company_id=? ORDER BY upload_date DESC", (user[5],))
    
    if files.empty:
        st.warning("The Evidence Vault is currently empty for your company.")
    else:
        st.markdown(f"### {len(files)} File(s) in Vault")
        
        # Display as a table with download buttons
        files_display = files[['risk_id', 'file_name', 'uploaded_by', 'upload_date']].copy()
        files_display.columns = ['Risk ID', 'File Name', 'Uploaded By', 'Upload Date']
        
        # Display data with a column for download button
        for index, row in files.iterrows():
            col1, col2, col3, col4, col5 = st.columns([1, 4, 2, 2, 2])
            col1.write(f"{row['risk_id']}")
            col2.write(row['file_name'])
            col3.write(row['uploaded_by'])
            col4.write(row['upload_date'])
            
            # Download button
            col5.download_button("Download", row['file_data'], row['file_name'], key=f"ev_{row['id']}", use_container_width=True)

# ==========================================
# 13. REPORTS (NIST Board Report - FIX 4)
# ==========================================
elif page == "Reports":
    st.markdown("## Risk Reporting & Analytics")
    target_id = user[5]
    
    # Admin selector for company reports
    if user[4] == "Admin":
        comps = run_query("SELECT id, name FROM companies")
        c_view = st.selectbox("Generate Report For:", comps['name'].tolist())
        target_id = comps[comps['name'] == c_view].iloc[0]['id']

    t1, t2, t3, t4, t5 = st.tabs(["Heatmap", "NIST Alignment (Board)", "Status Breakdown", "High Risks", "Audit Trail Report"])
    
    with t1:
        st.markdown("### Risk Heatmap Distribution")
        df = run_query("SELECT risk_score FROM risks WHERE company_id=?", (target_id,))
        
        if df.empty:
             st.warning("No risks to generate heatmap. Log some risks first.")
        else:
            high = len(df[df['risk_score']>=7])
            med = len(df[(df['risk_score']>=4) & (df['risk_score']<7)])
            low = len(df[df['risk_score']<4])
            
            fig = go.Figure(data=[go.Bar(x=['High (7-9)', 'Medium (4-6)', 'Low (1-3)'], y=[high, med, low], 
                                         marker_color=['red', 'orange', 'green'])])
            fig.update_layout(title_text='Risk Count by Score Severity', height=400)
            st.plotly_chart(fig, use_container_width=True)
        
    with t2:
        # FIX 4: NIST Board Level Alignment Report
        st.markdown("### NIST CSF Board Level Alignment Report")
        st.info("Shows the distribution of logged risks across the five core NIST CSF functions. This highlights where the company's biggest exposures lie from a strategic perspective.")
        cat = run_query("SELECT category, count(*) as count FROM risks WHERE company_id=? GROUP BY category", (target_id,))
        
        if cat.empty:
            st.warning("No risks to report on for NIST alignment.")
        else:
            # Display Data
            st.dataframe(cat.rename(columns={'category': 'NIST CSF Function'}), use_container_width=True)
            
            # Pie Chart Visualization
            fig = px.pie(cat, values='count', names='category', title='Risk Distribution by NIST CSF Function', height=400)
            st.plotly_chart(fig, use_container_width=True)
            
            # PDF Download for Board
            pdf_data = generate_pdf_report(f"NIST CSF Alignment Report for {user_company_name}", cat)
            st.download_button("Download NIST Board Report (PDF)", pdf_data, "nist_alignment_report.pdf", use_container_width=True)

        
    with t3:
        st.markdown("### Risk Status Breakdown")
        stat = run_query("SELECT status, count(*) as count FROM risks WHERE company_id=? GROUP BY status", (target_id,))
        if not stat.empty: 
            fig = px.pie(stat, values='count', names='status', title='Current Risk Status', height=400)
            st.plotly_chart(fig, use_container_width=True)

    with t4:
        st.markdown("### High Priority Risks (Score >= 7)")
        high_risks = run_query("SELECT title, risk_score, status, remediation_owner, remediation_date, jira_ticket_id FROM risks WHERE company_id=? AND risk_score >= 7 ORDER BY risk_score DESC", (target_id,))
        if high_risks.empty:
            st.info("No High Risks currently open.")
        else:
            st.dataframe(high_risks, use_container_width=True)
            pdf = generate_pdf_report("High Risks Priority List", high_risks)
            st.download_button("Download High Risks PDF", pdf, "high_risks_priority.pdf", use_container_width=True)
            
    with t5:
        st.markdown("### Audit Trail Report (Last 100 Actions)")
        audit_data = run_query("SELECT timestamp, user_email, action, details FROM audit_trail ORDER BY id DESC LIMIT 100")
        if audit_data.empty:
            st.info("No audit data available.")
        else:
            st.dataframe(audit_data, use_container_width=True)
            pdf = generate_pdf_report("Last 100 Actions Audit Log", audit_data)
            st.download_button("Download Audit Log PDF", pdf, "audit_log_report.pdf", use_container_width=True)


# ==========================================
# 14. ADMIN PANEL (FIX 3: Delete/Reset PW)
# ==========================================
elif page == "Admin Panel" and user[4] == "Admin":
    st.markdown("## Global Admin Panel")
    st.warning("Use caution when modifying users or company data.")
    comps = run_query("SELECT * FROM companies")
    
    t1, t2 = st.tabs(["Manage Users", "Company Management"])
    
    with t1:
        st.markdown("### Create New User")
        with st.expander("Create New User", expanded=False):
            with st.form("new_u"):
                u = st.text_input("Username", key="new_u_name")
                e = st.text_input("Email", key="new_u_email")
                p = st.text_input("Temporary Password", type="password", key="new_u_pass")
                r = st.selectbox("Role", ["Admin", "Approver", "User"], key="new_u_role")
                c = st.selectbox("Company", comps['name'].tolist(), key="new_u_comp")
                if st.form_submit_button("Create User"):
                    cid = comps[comps['name'] == c].iloc[0]['id']
                    h = hashlib.sha256(p.encode()).hexdigest()
                    try:
                        run_query("INSERT INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", (u, e, h, r, cid), is_write=True)
                        log_action(user[2], "USER CREATED", u)
                        st.success(f"User {u} created successfully.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Username or Email already exists.")


        st.markdown("### Manage Existing Users (Edit, Delete, Reset Password)")
        users = run_query("SELECT id, username, email, role, company_id FROM users")
        
        # Merge company name into user data for better display
        comp_map = comps.set_index('id')['name'].to_dict()
        users['company_name'] = users['company_id'].map(comp_map)
        
        st.dataframe(users[['username', 'email', 'role', 'company_name']], use_container_width=True)

        st.markdown("---")

        # --- Iterating for actions (Delete/Reset PW) ---
        st.markdown("#### User Actions")
        
        if "edit_user_id" not in st.session_state:
            st.session_state.edit_user_id = None
            
        for _, row in users.iterrows():
            col_u, col_r, col_d, col_e = st.columns([5, 2, 2, 2])
            col_u.markdown(f"**{row['username']}** ({row['company_name']})")
            
            # FIX 3: Reset Password Button
            if col_r.button("Reset PW", key=f"rpw_{row['id']}", help="Reset password to default 'Reset2025'", use_container_width=True):
                if row['id'] == user[0]:
                    st.error("Cannot reset your own password here.")
                else:
                    run_query("UPDATE users SET password=? WHERE id=?", (DEFAULT_PASSWORD_HASH, row['id']), is_write=True)
                    log_action(user[2], "PASSWORD RESET", f"User ID {row['id']} password reset to default.")
                    st.success(f"Password for {row['username']} reset to default.")
                    
            # FIX 3: Delete User Button
            if col_d.button("Delete", key=f"del_{row['id']}", help="Permanently delete this user", use_container_width=True):
                if row['id'] == user[0]:
                    st.error("Cannot delete your own account.")
                else:
                    run_query("DELETE FROM users WHERE id=?", (row['id'],), is_write=True)
                    log_action(user[2], "USER DELETED", f"Deleted User {row['username']} (ID: {row['id']}).")
                    st.success(f"User {row['username']} deleted.")
                    st.rerun()

            # Edit Button
            if col_e.button("Edit", key=f"ed_{row['id']}", use_container_width=True):
                st.session_state.edit_user_id = row['id']
                st.session_state.edit_user_data = row.to_dict()
                st.rerun()
        
        # Edit Form logic
        if st.session_state.edit_user_id:
            ed = st.session_state.edit_user_data
            st.markdown(f"#### Edit User: {ed['username']}")
            with st.form("ed_form"):
                new_u = st.text_input("Username", ed['username'])
                new_e = st.text_input("Email", ed['email'])
                new_r = st.selectbox("Role", ["Admin", "Approver", "User"], index=["Admin", "Approver", "User"].index(ed['role']))
                
                # Cannot change company in the edit panel, but display it
                st.text(f"Company: {ed['company_name']} (Cannot modify company here)")

                col_save, col_cancel = st.columns(2)
                
                if col_save.form_submit_button("Save Changes"):
                    run_query("UPDATE users SET username=?, email=?, role=? WHERE id=?", (new_u, new_e, new_r, ed['id']), is_write=True)
                    log_action(user[2], "USER EDITED", f"ID {ed['id']} updated. New Role: {new_r}")
                    st.success("User details saved.")
                    st.session_state.edit_user_id = None
                    st.session_state.edit_user_data = None
                    st.rerun()
                
                if col_cancel.form_submit_button("Cancel"):
                    st.session_state.edit_user_id = None
                    st.session_state.edit_user_data = None
                    st.rerun()


    with t2:
        st.markdown("### Manage Companies")
        # Simple company listing
        for _, c in comps.iterrows():
            st.text(f"ID: {c['id']} | Name: {c['name']}")
            
        with st.expander("Add New Company"):
            with st.form("new_c"):
                c_name = st.text_input("New Company Name")
                if st.form_submit_button("Add Company"):
                    try:
                        run_query("INSERT INTO companies (name) VALUES (?)", (c_name,), is_write=True)
                        log_action(user[2], "COMPANY ADDED", c_name)
                        st.success(f"Company {c_name} added.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Company name already exists.")

# ==========================================
# 15. AUDIT TRAIL
# ==========================================
elif page == "Audit Trail" and user[4] == "Admin":
    st.markdown("## Full System Audit Trail")
    st.info("A comprehensive, uneditable log of all major user actions for compliance and security review.")
    st.dataframe(run_query("SELECT * FROM audit_trail ORDER BY id DESC"), use_container_width=True)

st.markdown("---\n© 2025 Joval Wines - NIST Compliant Risk Platform (v51.0)")
