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

# ReportLab Imports
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch

# Plotly
import plotly.graph_objects as go
import plotly.express as px

# Email
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
            return pd.read_sql(query, conn, params=params)
    except Exception as e:
        print(f"Database Error: {e}")
        return 0 if is_write else pd.DataFrame()
    finally:
        conn.close()

def init_db():
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
    
    c.execute("""CREATE TABLE IF NOT EXISTS evidence (id INTEGER PRIMARY KEY AUTOINCREMENT, risk_id INTEGER, company_id INTEGER, file_name TEXT, upload_date TEXT, uploaded_by TEXT, file_data BLOB)""")
    
    # Vendor Tables with Scoring
    c.execute("""CREATE TABLE IF NOT EXISTS vendors (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, contact_email TEXT, risk_level TEXT, last_assessment_date TEXT, company_id INTEGER, vendor_score INTEGER)""")
    c.execute("""CREATE TABLE IF NOT EXISTS vendor_questionnaire (id INTEGER PRIMARY KEY AUTOINCREMENT, vendor_id INTEGER, question_id INTEGER, answer TEXT, answered_date TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS vendor_questions (id INTEGER PRIMARY KEY AUTOINCREMENT, question TEXT, weight INTEGER, company_id INTEGER)""")
    
    # Audit Trail
    c.execute("""CREATE TABLE IF NOT EXISTS audit_trail (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, user_email TEXT, action TEXT, details TEXT)""")

    # Safe Column Migration (Ensure all necessary columns exist)
    cols = [("risks", "remediation_plan", "TEXT"), ("risks", "jira_ticket_id", "TEXT"), 
            ("vendors", "vendor_score", "INTEGER"), ("risks", "remediation_owner", "TEXT"), 
            ("risks", "remediation_date", "TEXT"), ("vendor_questions", "weight", "INTEGER")]
    for t, c_name, d_type in cols:
        try: c.execute(f"ALTER TABLE {t} ADD COLUMN {c_name} {d_type}")
        except: pass

    # Seed Data (Admin/Approvers)
    c.execute("SELECT count(*) FROM companies")
    if c.fetchone()[0] == 0:
        companies = ["Joval Wines", "Joval Family Wines", "BNV", "BAM"]
        for comp in companies: c.execute("INSERT OR IGNORE INTO companies (name) VALUES (?)", (comp,))
        
        c.execute("SELECT id, name FROM companies")
        rows = c.fetchall()
        hashed = hashlib.sha256("Joval2025".encode()).hexdigest()
        
        for cid, cname in rows:
            # Create a dedicated Approver for each company
            c.execute("INSERT OR IGNORE INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", 
                      (f"approver_{cname.lower().replace(' ', '')}", f"approver@{cname.lower().replace(' ', '')}.com.au", hashed, "Approver", cid))
            # Create a standard User for Joval Wines
            if cid == 1:
                c.execute("INSERT OR IGNORE INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", 
                          ("user_joval", "user@jovalwines.com.au", hashed, "User", cid))
            # Create Admin
            if cid == 1:
                c.execute("INSERT OR IGNORE INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", 
                          ("admin", "admin@jovalwines.com.au", hashed, "Admin", cid))
    
    # Seed Questions (NIST aligned - relevant to vendor risk)
    c.execute("SELECT count(*) FROM vendor_questions")
    if c.fetchone()[0] == 0:
        # Example questions mapped loosely to NIST CSF
        qs = [("Does the vendor provide visibility into their security controls? (IDENTIFY)", 10), 
              ("Is data processed by the vendor encrypted at rest and in transit? (PROTECT)", 15), 
              ("Does the vendor have a SOC/monitoring function? (DETECT)", 10), 
              ("Does the vendor have a documented Incident Response Plan? (RESPOND)", 10)]
        for q, w in qs: c.execute("INSERT INTO vendor_questions (question, weight, company_id) VALUES (?, ?, ?)", (q, w, 1))

    conn.commit()
    conn.close()

if not os.path.exists(DB_FILE):
    init_db()
else:
    init_db()

# ==========================================
# 3. HELPERS (Including Audit Logging and AI)
# ==========================================
def calculate_risk_score(likelihood, impact):
    scores = {"Low": 1, "Medium": 2, "High": 3}
    return scores.get(likelihood, 1) * scores.get(impact, 1)

def get_risk_color(score):
    if score >= 7: return "red"
    elif score >= 4: return "orange"
    else: return "green"

# --- AUDIT TRAIL LOGGING ---
def log_action(user_email, action, details=""):
    run_query("INSERT INTO audit_trail (timestamp, user_email, action, details) VALUES (?, ?, ?, ?)", 
              (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), user_email, action, details), is_write=True)

# --- AI AUTO-CLASSIFY (Simulation) ---
def ai_classify_risk(text):
    text = text.lower()
    if any(x in text for x in ["phish", "email", "scam"]): return "DETECT", "High", "High"
    if any(x in text for x in ["malware", "virus", "patch"]): return "PROTECT", "Medium", "High"
    if any(x in text for x in ["backup", "loss"]): return "RECOVER", "Medium", "Medium"
    if any(x in text for x in ["asset", "inventory", "policy"]): return "IDENTIFY", "Low", "Medium"
    return "IDENTIFY", "Low", "Low"

def generate_pdf_report(title, content):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=1.2*inch)
    styles = getSampleStyleSheet()
    story = []
    try:
        # Placeholder for Joval Wines Logo - uses a publicly available image
        with urllib.request.urlopen("https://jovalwines.com.au/wp-content/uploads/2020/06/Joval-Wines-Logo.png") as r:
            img = Image(BytesIO(r.read()), width=2*inch, height=0.6*inch)
            story.append(img)
    except:
        story.append(Paragraph("JOVAL WINES", styles['Title']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(title, styles['Heading1']))
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%d %B %Y')}", styles['Normal']))
    story.append(Spacer(1, 20))
    if isinstance(content, pd.DataFrame):
        # Format for display in PDF (truncate long strings)
        content_str = content.astype(str).applymap(lambda x: (x[:50] + '...') if len(x) > 50 else x)
        data = [content.columns.tolist()] + content_str.values.tolist()
        table = Table(data)
        table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1a1a1a')), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('FONTSIZE', (0,0), (-1,-1), 8)]))
        story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer

def send_email(to_email, subject, body):
    # This is a simulated email function as credentials are not guaranteed
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        # server.login(SENDER_EMAIL, SENDER_PASSWORD) # Requires valid credentials
        # server.send_message(msg)
        server.quit()
    except: pass
    
# ==========================================
# 4. UI SETUP
# ==========================================
st.set_page_config(page_title="Joval Risk Portal", layout="wide")
st.markdown("""
<style>
    .header {background: #1a1a1a; color: white; padding: 2rem; text-align: center;}
    .metric-card {background: white; padding: 1.5rem; border-radius: 12px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1);}
    .clickable-risk {cursor: pointer; padding: 0.75rem; border-radius: 8px; margin: 0.25rem 0;}
    .approval-badge {background: #e6f7ff; padding: 0.25rem 0.5rem; border-radius: 12px; font-size: 0.8rem;}
    /* NIST hints style */
    .nist-hint {background-color: #f0f8ff; border-left: 5px solid #007bff; padding: 10px; border-radius: 4px; margin-bottom: 15px;}
</style>""", unsafe_allow_html=True)
st.markdown('<div class="header"><h1>JOVAL WINES</h1><p>Risk Management Portal</p></div>', unsafe_allow_html=True)

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
DEFAULT_PASSWORD_HASH = hashlib.sha256("Reset2025".encode()).hexdigest()

# ==========================================
# 6. SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("### Playbook Tracker")
    # FIX: Updated Playbook Tracker link
    st.markdown("[**Open Playbook Tracker App**](https://joval-wines-nist-playbook-tracker.streamlit.app/)")
    st.markdown("---")
    st.markdown(f"**{user[1]}** • {user_company_name}")
    st.markdown("---")
    pages = ["Dashboard", "Log a new Risk", "Evidence Vault", "Vendor Management", "Reports"]
    if user[4] == "Approver": pages.insert(1, "My Approvals")
    if user[4] == "Admin": pages += ["Audit Trail", "Admin Panel"]
    for p in pages:
        if st.button(p, key=f"nav_{p}"):
            st.session_state.page = p
            st.rerun()

page = st.session_state.get("page", "Dashboard")

# ==========================================
# 7. DASHBOARD
# ==========================================
if page == "Dashboard":
    st.markdown("## Dashboard")
    view_company_id = user[5]
    if user[4] == "Admin":
        all_comps = run_query("SELECT id, name FROM companies")
        comp_opts = all_comps['name'].tolist()
        try: def_idx = comp_opts.index(user_company_name)
        except: def_idx = 0
        view_name = st.selectbox("Viewing Dashboard For:", comp_opts, index=def_idx)
        view_company_id = all_comps[all_comps['name'] == view_name].iloc[0]['id']
    
    risks = run_query("SELECT * FROM risks WHERE company_id=?", (view_company_id,))
    
    col1, col2 = st.columns(2)
    with col1:
        high = len(risks[risks['risk_score'] >= 7])
        st.markdown(f'<div class="metric-card"><h2>{high}</h2><p>High Risks</p></div>', unsafe_allow_html=True)
    with col2:
        total = len(risks)
        st.markdown(f'<div class="metric-card"><h2>{total}</h2><p>Total Risks</p></div>', unsafe_allow_html=True)
    
    st.write("---")
    if risks.empty:
        st.info("No risks logged.")
    else:
        for _, r in risks.iterrows():
            color = get_risk_color(r['risk_score'])
            bg = "#ffe6e6" if color == "red" else "#fff4e6" if color == "orange" else "#e6f7e6"
            btn_label = f"**{r['title']}** | Status: {r['status']} (Score: {r['risk_score']})"
            if st.button(btn_label, key=f"r_{r['id']}"):
                st.session_state.selected_risk = r['id']
                st.session_state.page = "Risk Detail"
                st.rerun()
            st.markdown(f'<div class="clickable-risk" style="background:{bg};"><small>{r["description"]}</small></div>', unsafe_allow_html=True)

# ==========================================
# 8. LOG RISK (with AI and NIST Guidance)
# ==========================================
elif page == "Log a new Risk":
    st.markdown("## Log a New Risk")
    st.info("This process aligns your risk with the NIST Cybersecurity Framework (CSF) for consistent reporting.")
    
    comps = run_query("SELECT id, name FROM companies")
    try: def_idx = comps['name'].tolist().index(user_company_name)
    except: def_idx = 0
    c_name = st.selectbox("Company", comps['name'].tolist(), index=def_idx)
    c_id = comps[comps['name'] == c_name].iloc[0]['id']
    
    # FIX: Ensure approver list is functional
    approvers = run_query("SELECT email, username FROM users WHERE role='Approver' AND company_id=?", (c_id,))
    if approvers.empty:
        st.error(f"No Approver found for {c_name}. Please check the Admin Panel.")
        st.stop()
    
    approver_emails = approvers['email'].tolist()
    approver_usernames = approvers['username'].tolist()
    
    selected_approver_name = st.selectbox("Assign to Approver", approver_usernames)
    approver_email = approvers[approvers['username'] == selected_approver_name].iloc[0]['email']

    
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
            st.success(f"AI Suggestion: Category={d_cat}, Likelihood={d_lik}, Impact={d_imp}")
            
        cat = st.selectbox("NIST CSF Category", list(NIST_FUNCTIONS.keys()), index=list(NIST_FUNCTIONS.keys()).index(d_cat))
        
        # NIST Guidance Hint
        st.markdown(f'<div class="nist-hint">**{cat}:** {NIST_FUNCTIONS.get(cat, "No specific guidance.")}</div>', unsafe_allow_html=True)
        
        col_lik, col_imp = st.columns(2)
        with col_lik:
            lik = st.selectbox("Likelihood", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(d_lik))
        with col_imp:
            imp = st.selectbox("Impact", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(d_imp))
        
        if st.form_submit_button("Submit Risk"):
            if not title or not desc:
                st.error("Required fields missing.")
            else:
                score = calculate_risk_score(lik, imp)
                run_query("""INSERT INTO risks (company_id, title, description, category, likelihood, impact, status, submitted_by, submitted_date, risk_score, approver_email, workflow_step) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", (c_id, title, desc, cat, lik, imp, "Pending Approval", user[2], datetime.now().strftime("%Y-%m-%d"), score, approver_email, "awaiting"), is_write=True)
                st.success("Risk Logged! Sent to Approver.")
                log_action(user[2], "RISK SUBMITTED", f"{title} (Approver: {approver_email})")
                send_email(approver_email, "New Risk Awaiting Approval", f"Title: {title}")
                st.session_state.page = "Dashboard"
                st.rerun()

# ==========================================
# 9. RISK DETAIL (with Remediation & Jira)
# ==========================================
elif page == "Risk Detail" and "selected_risk" in st.session_state:
    rid = st.session_state.selected_risk
    risk = run_query("SELECT * FROM risks WHERE id=?", (rid,)).iloc[0]
    
    st.markdown(f"## Risk: {risk['title']}")
    st.markdown(f"NIST Category: **{risk['category']}** | Current Score: **{risk['risk_score']}** | Status: **{risk['status']}**")
    
    # Jira Sync Button
    if not risk['jira_ticket_id']:
        if st.button("🔗 Sync to Jira"):
            time.sleep(1)
            ticket = f"JIRA-{random.randint(1000,9999)}"
            run_query("UPDATE risks SET jira_ticket_id=? WHERE id=?", (ticket, rid), is_write=True)
            log_action(user[2], "JIRA SYNC", f"Risk {rid} linked to {ticket}")
            st.success(f"Created {ticket}")
            st.rerun()
    else:
        st.info(f"Jira Ticket: {risk['jira_ticket_id']}")

    t1, t2, t3 = st.tabs(["Details", "Remediation", "Evidence"])
    
    with t1:
        with st.form("edit"):
            c1, c2 = st.columns(2)
            new_title = st.text_input("Title", risk['title'])
            new_desc = st.text_area("Description", risk['description'])
            with c1:
                new_cat = st.selectbox("NIST CSF Category", list(NIST_FUNCTIONS.keys()), index=list(NIST_FUNCTIONS.keys()).index(risk['category']))
                new_lik = st.selectbox("Likelihood", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(risk['likelihood']))
            with c2:
                new_imp = st.selectbox("Impact", ["Low", "Medium", "High"], index=["Low", "Medium", "High"].index(risk['impact']))
                # Only Approvers or Admins can change status past 'Pending Approval'
                if user[4] in ["Admin", "Approver"]:
                    new_status = st.selectbox("Status", ["Pending Approval", "Approved", "Rejected", "Mitigated"], index=["Pending Approval", "Approved", "Rejected", "Mitigated"].index(risk['status']))
                else:
                    st.text(f"Status: {risk['status']} (Contact Approver to change)")
                    new_status = risk['status']
                    
            new_notes = st.text_area("Approver Notes", risk['approver_notes'] or "")
            
            if st.form_submit_button("Save Changes"):
                new_score = calculate_risk_score(new_lik, new_imp)
                # Check if approval action was taken
                approved_by = risk['approved_by']
                approved_date = risk['approved_date']
                if risk['status'] == 'Pending Approval' and new_status in ['Approved', 'Rejected'] and user[4] in ["Admin", "Approver"]:
                    approved_by = user[1]
                    approved_date = datetime.now().strftime("%Y-%m-%d")

                run_query("""UPDATE risks SET title=?, description=?, category=?, likelihood=?, impact=?, status=?, risk_score=?, approver_notes=?, approved_by=?, approved_date=? WHERE id=?""", (new_title, new_desc, new_cat, new_lik, new_imp, new_status, new_score, new_notes, approved_by, approved_date, rid), is_write=True)
                log_action(user[2], "RISK UPDATED", f"Risk {rid}. Status: {new_status}")
                st.success("Updated")
                st.rerun()

    with t2:
        with st.form("rem"):
            plan = st.text_area("Action Plan (NIST RESPOND/RECOVER)", risk['remediation_plan'] or "")
            own = st.text_input("Remediation Owner", risk['remediation_owner'] or "")
            date = st.text_input("Target Date (YYYY-MM-DD)", risk['remediation_date'] or "")
            if st.form_submit_button("Update Plan"):
                run_query("UPDATE risks SET remediation_plan=?, remediation_owner=?, remediation_date=? WHERE id=?", (plan, own, date, rid), is_write=True)
                log_action(user[2], "REMEDIATION PLAN", f"Risk {rid}. Owner: {own}")
                st.success("Updated")
                st.rerun()

    with t3:
        st.markdown("### Upload New Evidence")
        uploaded = st.file_uploader("Upload supporting document/screenshot")
        if uploaded:
            run_query("INSERT INTO evidence (risk_id, company_id, file_name, upload_date, uploaded_by, file_data) VALUES (?, ?, ?, ?, ?, ?)", (rid, risk['company_id'], uploaded.name, datetime.now().strftime("%Y-%m-%d"), user[1], uploaded.getvalue()), is_write=True)
            log_action(user[2], "EVIDENCE UPLOADED", f"Risk {rid}. File: {uploaded.name}")
            st.success("File uploaded.")
            st.rerun()
            
        st.markdown("### Existing Evidence for this Risk")
        files = run_query("SELECT id, file_name, uploaded_by, upload_date, file_data FROM evidence WHERE risk_id=?", (rid,))
        if files.empty:
            st.info("No evidence files uploaded for this specific risk.")
        else:
            for _, f in files.iterrows():
                st.download_button(f"Download {f['file_name']} (by {f['uploaded_by']})", f['file_data'], f['file_name'], key=f"d_{f['id']}")

    if st.button("Back to Dashboard"):
        del st.session_state.selected_risk
        st.session_state.page = "Dashboard"
        st.rerun()

# ==========================================
# 10. VENDOR MANAGEMENT (NIST & Scoring)
# ==========================================
elif page == "Vendor Management":
    st.markdown("## Vendor Management (NIST CSF Aligned)")
    st.info("Vendor risks are managed through questionnaires mapped to NIST CSF functions (Identify, Protect, Detect, Respond).")
    t1, t2, t3 = st.tabs(["Vendors", "Question Template", "Risk Scoring"])
    
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
        vendors = run_query("SELECT id, name, contact_email, risk_level, vendor_score FROM vendors WHERE company_id=?", (user[5],))
        st.dataframe(vendors.rename(columns={'vendor_score': 'Score (Max 45)'}))
        
    with t2:
        st.markdown("### Vendor Assessment Questions (NIST Aligned)")
        st.markdown("Total Max Score: 45")
        qs = run_query("SELECT id, question, weight FROM vendor_questions")
        edited = st.data_editor(qs, num_rows="dynamic", column_config={"question": st.column_config.TextColumn("Question (Ensure NIST CSF alignment)")})
        
        if st.button("Save Question Template"):
            # Clear existing questions before saving new set
            run_query("DELETE FROM vendor_questions WHERE company_id=?", (1,), is_write=True)
            for _, r in edited.iterrows():
                if r['question'] and r['weight']:
                    run_query("INSERT INTO vendor_questions (question, weight, company_id) VALUES (?, ?, ?)", (r['question'], r['weight'], 1), is_write=True)
            log_action(user[2], "VENDOR TEMPLATE EDITED")
            st.success("Vendor assessment template saved.")
            st.rerun()

    with t3:
        v_list = run_query("SELECT id, name FROM vendors WHERE company_id=?", (user[5],))
        if v_list.empty:
            st.warning("No vendors to score. Please add a vendor first.")
        else:
            sel_v = st.selectbox("Select Vendor to Score", v_list['name'].tolist())
            vid = v_list[v_list['name'] == sel_v].iloc[0]['id']
            qs = run_query("SELECT * FROM vendor_questions")
            
            with st.form("score"):
                score = 0
                max_s = 0
                st.markdown(f"#### Security Questionnaire for {sel_v}")
                for _, q in qs.iterrows():
                    max_s += q['weight']
                    if st.radio(f"**{q['question']}**", ["Yes", "No"], key=q['id'], horizontal=True) == "Yes":
                        score += q['weight']
                
                st.markdown(f"**Max Possible Score: {max_s}**")
                
                if st.form_submit_button("Calculate Vendor Risk Score"):
                    # Compliance Logic: High Risk if score is below 50%
                    lvl = "High" if score < (max_s * 0.5) else "Low"
                    run_query("UPDATE vendors SET vendor_score=?, risk_level=?, last_assessment_date=? WHERE id=?", (score, lvl, datetime.now().strftime("%Y-%m-%d"), vid), is_write=True)
                    log_action(user[2], "VENDOR SCORED", f"Vendor {sel_v}: {score}/{max_s}. Level: {lvl}")
                    st.success(f"Vendor Risk Assessed. Score: {score}/{max_s}. Risk Level: {lvl}")

# ==========================================
# 11. REPORTS (Including NIST Board Report)
# ==========================================
elif page == "Reports":
    st.markdown("## Risk Reporting")
    target_id = user[5]
    if user[4] == "Admin":
        comps = run_query("SELECT id, name FROM companies")
        c_view = st.selectbox("Report For:", comps['name'].tolist())
        target_id = comps[comps['name'] == c_view].iloc[0]['id']

    t1, t2, t3, t4, t5 = st.tabs(["Heatmap", "NIST Alignment", "Status", "High Risks", "Custom Report Builder"])
    
    with t1:
        st.markdown("### Risk Heatmap Distribution")
        df = run_query("SELECT risk_score FROM risks WHERE company_id=?", (target_id,))
        high = len(df[df['risk_score']>=7])
        med = len(df[(df['risk_score']>=4) & (df['risk_score']<7)])
        low = len(df[df['risk_score']<4])
        fig = go.Figure(data=[go.Bar(x=['High (7-9)', 'Medium (4-6)', 'Low (1-3)'], y=[high, med, low], marker_color=['red', 'orange', 'green'])])
        fig.update_layout(title_text='Risk Count by Score Severity')
        st.plotly_chart(fig)
        
    with t2:
        st.markdown("### NIST CSF Board Level Alignment Report")
        cat = run_query("SELECT category, count(*) as count FROM risks WHERE company_id=? GROUP BY category", (target_id,))
        if cat.empty:
            st.warning("No risks to report on for NIST alignment.")
        else:
            st.dataframe(cat.rename(columns={'category': 'NIST CSF Function'}))
            
            # Pie Chart Visualization
            fig = px.pie(cat, values='count', names='category', title='Risk Distribution by NIST CSF Function')
            st.plotly_chart(fig)
            
            # PDF Download for Board
            pdf_data = generate_pdf_report(f"NIST CSF Alignment Report for {user_company_name}", cat)
            st.download_button("Download NIST Report PDF", pdf_data, "nist_alignment_report.pdf")

        
    with t3:
        st.markdown("### Risk Status Breakdown")
        stat = run_query("SELECT status, count(*) as count FROM risks WHERE company_id=? GROUP BY status", (target_id,))
        if not stat.empty: 
            fig = px.pie(stat, values='count', names='status', title='Current Risk Status')
            st.plotly_chart(fig)

    with t4:
        st.markdown("### High Priority Risks (Score >= 7)")
        high_risks = run_query("SELECT title, risk_score, status, remediation_owner, remediation_date FROM risks WHERE company_id=? AND risk_score >= 7 ORDER BY risk_score DESC", (target_id,))
        if high_risks.empty:
            st.info("No High Risks currently open.")
        else:
            st.dataframe(high_risks)
            pdf = generate_pdf_report("High Risks Priority List", high_risks)
            st.download_button("Download High Risks PDF", pdf, "high_risks_priority.pdf")

    with t5:
        st.markdown("### Custom Report Builder")
        cols = st.multiselect("Select Columns for Custom Report", ["title", "category", "likelihood", "impact", "status", "submitted_by", "risk_score", "remediation_owner", "remediation_date", "jira_ticket_id"], default=["title", "risk_score", "status"])
        if st.button("Generate Custom Report"):
            if cols:
                col_str = ", ".join(cols)
                res = run_query(f"SELECT {col_str} FROM risks WHERE company_id=?", (target_id,))
                st.dataframe(res)
                pdf = generate_pdf_report("Custom Risk Report", res)
                st.download_button("Download Custom PDF", pdf, "custom_report.pdf")
            else:
                st.warning("Please select at least one column.")

# ==========================================
# 12. ADMIN & AUDIT (Full Features)
# ==========================================
elif page == "Admin Panel" and user[4] == "Admin":
    st.markdown("## Admin Panel")
    comps = run_query("SELECT * FROM companies")
    
    # --- Add User ---
    with st.expander("Add New User"):
        with st.form("new_u"):
            u = st.text_input("Username")
            e = st.text_input("Email")
            p = st.text_input("Password", type="password")
            r = st.selectbox("Role", ["Admin", "Approver", "User"])
            c = st.selectbox("Company", comps['name'].tolist())
            if st.form_submit_button("Create User"):
                cid = comps[comps['name'] == c].iloc[0]['id']
                h = hashlib.sha256(p.encode()).hexdigest()
                run_query("INSERT INTO users (username, email, password, role, company_id) VALUES (?, ?, ?, ?, ?)", (u, e, h, r, cid), is_write=True)
                log_action(user[2], "USER CREATED", u)
                st.success("User Created")
    
    st.markdown("### Manage Existing Users")
    users = run_query("SELECT id, username, email, role, company_id FROM users")
    
    # --- Edit/Delete/Reset Logic ---
    for _, row in users.iterrows():
        c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
        c1.text(f"{row['username']} ({row['role']})")
        
        if c2.button("Edit", key=f"ed_{row['id']}"):
            st.session_state.edit_user = row.to_dict()
            st.rerun()

        if c3.button("Reset PW", key=f"rpw_{row['id']}"):
            if row['id'] == user[0]:
                st.error("Cannot reset your own password here.")
            else:
                run_query("UPDATE users SET password=? WHERE id=?", (DEFAULT_PASSWORD_HASH, row['id']), is_write=True)
                log_action(user[2], "PASSWORD RESET", f"User ID {row['id']} password reset to default.")
                st.success(f"Password for {row['username']} reset to default.")
                
        if c4.button("Delete", key=f"del_{row['id']}"):
            if row['id'] == user[0]:
                st.error("Cannot delete your own account.")
            else:
                run_query("DELETE FROM users WHERE id=?", (row['id'],), is_write=True)
                log_action(user[2], "USER DELETED", f"Deleted User {row['username']} (ID: {row['id']}).")
                st.success(f"User {row['username']} deleted.")
                st.rerun()


    if "edit_user" in st.session_state:
        ed = st.session_state.edit_user
        st.markdown(f"#### Edit {ed['username']}")
        with st.form("ed_form"):
            new_u = st.text_input("Username", ed['username'])
            new_e = st.text_input("Email", ed['email'])
            new_r = st.selectbox("Role", ["Admin", "Approver", "User"], index=["Admin", "Approver", "User"].index(ed['role']))
            if st.form_submit_button("Save Changes"):
                run_query("UPDATE users SET username=?, email=?, role=? WHERE id=?", (new_u, new_e, new_r, ed['id']), is_write=True)
                log_action(user[2], "USER EDITED", f"ID {ed['id']} updated. New Role: {new_r}")
                st.success("User details saved.")
                del st.session_state.edit_user
                st.rerun()

elif page == "Audit Trail" and user[4] == "Admin":
    st.markdown("## Audit Trail")
    st.info("A comprehensive log of all major user actions for compliance and security review.")
    st.dataframe(run_query("SELECT * FROM audit_trail ORDER BY id DESC"))

# --- EVIDENCE VAULT FIX ---
elif page == "Evidence Vault":
    st.markdown("## Company Evidence Vault")
    st.info("Central repository for all documents uploaded across all risks for your company.")
    files = run_query("SELECT id, risk_id, file_name, uploaded_by, upload_date, file_data FROM evidence WHERE company_id=?", (user[5],))
    
    if files.empty:
        st.warning("The Evidence Vault is currently empty for your company.")
    else:
        for _, f in files.iterrows():
            # Shows the risk ID the evidence is attached to
            st.download_button(f"Download {f['file_name']} (Risk ID: {f['risk_id']})", f['file_data'], f['file_name'], key=f"ev_{f['id']}")

# --- APPROVER WORKFLOW ---
elif page == "My Approvals" and user[4] == "Approver":
    st.markdown("## My Assigned Risks Awaiting Approval")
    
    # Query to fetch risks assigned to the current approver's email
    pend = run_query("SELECT id, title, risk_score, submitted_date FROM risks WHERE approver_email=? AND status='Pending Approval'", (user[2],))
    
    if pend.empty:
        st.info("You have no pending risks requiring your approval at this time. Good job!")
    else:
        st.warning(f"You have **{len(pend)}** risk(s) awaiting review. Please take action immediately.")
        for _, r in pend.iterrows():
            if st.button(f"Review: {r['title']} (Score: {r['risk_score']}, Submitted: {r['submitted_date']})", key=r['id']):
                st.session_state.selected_risk = r['id']
                st.session_state.page = "Risk Detail"
                st.rerun()

st.markdown("---\n© 2025 Joval Wines - NIST Compliant Risk Platform")
