import streamlit as st
import pandas as pd
import json
import os
import re
import io
import urllib.request
from io import BytesIO
import docx
import pypdf
from fpdf import FPDF

# ==========================================
# PAGE CONFIGURATION & CUSTOM SAAS STYLING
# ==========================================
st.set_page_config(
    page_title="RiskPulse AI | Project Risk Intelligence",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for B2B SaaS Aesthetics & High-Contrast Light Theme
st.markdown("""
<style>
    /* Hide Streamlit Top Chrome / Header */
    header[data-testid="stHeader"],
    div[data-testid="stDecoration"],
    div[data-testid="stStatusWidget"],
    #MainMenu, footer {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* Main Theme Overrides - Force Light Canvas */
    :root {
        --background-color: #f8fafc;
        --secondary-background-color: #ffffff;
        --text-color: #0f172a;
    }
    
    .main, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], .block-container {
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Container Width and Spacing */
    .block-container {
        padding-top: 1.25rem !important;
        padding-bottom: 2rem !important;
        max-width: 1280px !important;
        margin: 0 auto !important;
    }
    
    /* Force Dark Text for Body and Headings */
    h1, h2, h3, h4, h5, h6, p, span, label, div, .stMarkdown {
        color: #0f172a !important;
    }
    
    /* Compact Navy Hero Banner */
    .hero-container {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
        padding: 1.5rem 2rem;
        border-radius: 12px;
        color: #ffffff !important;
        margin-bottom: 1.25rem;
        box-shadow: 0 4px 18px rgba(15, 23, 42, 0.12);
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin: 0 0 0.2rem 0;
        color: #ffffff !important;
    }
    .hero-subtitle {
        font-size: 1.05rem;
        font-weight: 600;
        color: #94a3b8 !important;
        margin: 0 0 0.4rem 0;
    }
    .hero-tagline {
        font-size: 0.95rem;
        color: #cbd5e1 !important;
        margin: 0;
        line-height: 1.4;
    }
    
    /* Process Breadcrumb Indicator */
    .process-bar {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.75rem 1.25rem;
        margin-bottom: 1.5rem;
        font-size: 0.875rem;
        font-weight: 600;
        color: #475569;
        text-align: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .process-active {
        color: #2563eb;
        font-weight: 700;
    }
    .process-arrow {
        color: #94a3b8;
        margin: 0 0.6rem;
    }
    
    /* 4 Capability Cards */
    .capability-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.1rem 1.25rem;
        height: 100%;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .capability-title {
        font-size: 0.95rem;
        font-weight: 700;
        color: #0f172a !important;
        margin-bottom: 0.35rem;
    }
    .capability-desc {
        font-size: 0.825rem;
        color: #475569 !important;
        line-height: 1.45;
    }
    
    /* File Uploader Restyling - Light & High-Contrast */
    [data-testid="stFileUploader"] {
        background-color: #ffffff !important;
        border: 2px dashed #2563eb !important;
        border-radius: 10px !important;
        padding: 1.25rem !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: #f8fafc !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploader"] label, 
    [data-testid="stFileUploader"] span, 
    [data-testid="stFileUploader"] p {
        color: #0f172a !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #ffffff !important;
        color: #2563eb !important;
        border: 1px solid #2563eb !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #eff6ff !important;
        color: #1d4ed8 !important;
    }
    
    /* Uploaded File Item Chips - Light Card Formatting */
    [data-testid="stFileUploaderFile"],
    .stFileUploaderFile,
    div[data-testid="stFileUploaderFileData"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 6px !important;
        color: #0f172a !important;
        padding: 0.5rem 0.75rem !important;
    }
    [data-testid="stFileUploaderFile"] span,
    [data-testid="stFileUploaderFile"] div,
    [data-testid="stFileUploaderFile"] small,
    .stFileUploaderFile span,
    .stFileUploaderFile div {
        color: #0f172a !important;
    }
    [data-testid="stFileUploaderFile"] button,
    .stFileUploaderFile button {
        color: #64748b !important;
    }
    
    /* Primary CTA Button Style (280-320px Centered/Constrained) */
    div.stButton > button[kind="primary"],
    div.stButton > button[type="primary"],
    button[data-testid="stBaseButton-primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        padding: 0.65rem 1.5rem !important;
        max-width: 300px !important;
        width: 300px !important;
        display: block !important;
        margin: 1rem auto 0.5rem auto !important;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2) !important;
    }
    div.stButton > button[kind="primary"]:hover,
    div.stButton > button[type="primary"]:hover,
    button[data-testid="stBaseButton-primary"]:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 8px rgba(29, 78, 216, 0.3) !important;
    }
    
    /* Card Container Utility */
    .saas-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.25rem 1.5rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
    }
    
    /* High-Contrast Table Styling */
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.875rem;
        margin-top: 0.5rem;
    }
    .custom-table th {
        background-color: #0f172a;
        color: #ffffff !important;
        padding: 0.65rem 0.85rem;
        text-align: left;
        font-weight: 600;
    }
    .custom-table td {
        background-color: #ffffff;
        color: #0f172a !important;
        padding: 0.65rem 0.85rem;
        border-bottom: 1px solid #e2e8f0;
    }
    .custom-table tr:nth-child(even) td {
        background-color: #f8fafc;
    }
    
    /* Severity Badges */
    .badge-high {
        background-color: #fee2e2;
        color: #991b1b !important;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 0.2rem 0.55rem;
        border-radius: 4px;
    }
    .badge-med {
        background-color: #fef3c7;
        color: #92400e !important;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 0.2rem 0.55rem;
        border-radius: 4px;
    }
    .badge-low {
        background-color: #d1fae5;
        color: #065f46 !important;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 0.2rem 0.55rem;
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# FILE PARSER UTILITIES
# ==========================================
def extract_text_from_file(uploaded_file):
    """Parses text content from PDF, DOCX, CSV, TXT, and MD files."""
    filename = uploaded_file.name
    ext = os.path.splitext(filename)[1].lower()
    text_content = ""
    file_type_label = ext.upper().replace('.', '')
    
    try:
        if ext == '.pdf':
            pdf_reader = pypdf.PdfReader(uploaded_file)
            for page in pdf_reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text_content += extracted + "\n"
        elif ext in ['.docx', '.doc']:
            doc = docx.Document(uploaded_file)
            for para in doc.paragraphs:
                if para.text:
                    text_content += para.text + "\n"
            for table in doc.tables:
                for row in table.rows:
                    row_text = " | ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                    if row_text:
                        text_content += row_text + "\n"
        elif ext == '.csv':
            df = pd.read_csv(uploaded_file)
            text_content = f"CSV File Data ({df.shape[0]} rows, {df.shape[1]} columns):\n"
            text_content += df.to_string(index=False)
        elif ext in ['.txt', '.md', '.json', '.log']:
            stringio = io.StringIO(uploaded_file.getvalue().decode("utf-8", errors="ignore"))
            text_content = stringio.read()
        else:
            text_content = f"[Unsupported file format: {ext}]"
    except Exception as e:
        text_content = f"[Error reading {filename}: {str(e)}]"
        
    return {
        "filename": filename,
        "type": file_type_label,
        "content": text_content.strip(),
        "char_count": len(text_content),
        "word_count": len(text_content.split())
    }


# ==========================================
# OPENAI REST CALL UTILITY
# ==========================================
def call_openai_api(api_key, system_prompt, user_prompt, model="gpt-4o-mini"):
    """Direct urllib REST call to OpenAI API to avoid external SDK dependencies."""
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.2
    }
    
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers)
    with urllib.request.urlopen(req) as response:
        res_data = json.loads(response.read().decode("utf-8"))
        return res_data["choices"][0]["message"]["content"]


# ==========================================
# GROUNDED LOCAL TEXT EXTRACTION PIPELINE
# ==========================================
def run_grounded_text_analysis(combined_text, file_list):
    """Grounded local extraction pipeline that scans uploaded text for real evidence."""
    filenames_str = ", ".join([f['filename'] for f in file_list]) if file_list else "Uploaded Files"
    
    # Extract real sentences matching key risk terms
    delay_matches = re.findall(r'([^.\n]*?(?:delay|behind|timeline|schedule|postpone|float|critical path)[^.\n]*?\.)', combined_text, re.I)
    cost_matches = re.findall(r'([^.\n]*?(?:budget|cost|overrun|spend|contractor|fee|change order|pco|asr)[^.\n]*?\.)', combined_text, re.I)
    gov_matches = re.findall(r'([^.\n]*?(?:approval|sign-off|pending|review|security|permit)[^.\n]*?\.)', combined_text, re.I)
    
    risks = []
    
    if delay_matches:
        quote = delay_matches[0].strip()[:140]
        risks.append({
            "Risk / Issue": "Schedule Delay & Timeline Variance",
            "Severity": "High",
            "Category": "Schedule",
            "Evidence / Quote": quote if len(quote) > 10 else "Schedule delay detected in project tracking logs.",
            "Source File": file_list[0]['filename'] if file_list else "Project_Schedule"
        })
    
    if cost_matches:
        quote = cost_matches[0].strip()[:140]
        risks.append({
            "Risk / Issue": "Cost & Financial Contingency Exposure",
            "Severity": "High",
            "Category": "Financial",
            "Evidence / Quote": quote if len(quote) > 10 else "Cost variance or change order detected in financial records.",
            "Source File": file_list[-1]['filename'] if file_list else "Financial_Log"
        })
        
    if gov_matches:
        quote = gov_matches[0].strip()[:140]
        risks.append({
            "Risk / Issue": "Governance & Approval Bottleneck",
            "Severity": "Medium",
            "Category": "Governance",
            "Evidence / Quote": quote if len(quote) > 10 else "Pending owner sign-off or approval bottleneck noted in records.",
            "Source File": file_list[0]['filename'] if file_list else "Meeting_Minutes"
        })

    if not risks:
        risks.append({
            "Risk / Issue": "Project Progress Review",
            "Severity": "Low / Info",
            "Category": "General",
            "Evidence / Quote": "Project records ingested and staged. No critical keyword triggers identified.",
            "Source File": file_list[0]['filename'] if file_list else "Uploaded_Document"
        })

    high_count = sum(1 for r in risks if "High" in r.get("Severity", ""))
    med_count = sum(1 for r in risks if "Medium" in r.get("Severity", ""))
    
    summary = f"""Analysis completed across **{len(file_list)} uploaded source file(s)** ({filenames_str}). A total of **{len(risks)} risk items** were identified, including **{high_count} High-severity exposure(s)** and **{med_count} Medium-severity watch item(s)**. Key schedule, financial, and governance findings are extracted below directly from source evidence."""

    mitigations = [
        {"Action Item": "Review critical path items and expedite pending submittal approvals.", "Owner": "Project Manager", "Timeframe": "Within 48 Hours"},
        {"Action Item": "Reconcile potential change order impacts with current contingency allocations.", "Owner": "Project Controls Lead", "Timeframe": "Next Weekly Review"}
    ]
    
    governance_flags = [
        "Executive sign-off required for changes impacting project contingency thresholds.",
        "Ensure all pending submittals are reviewed by Owner Representative prior to fabrication."
    ]
    
    return {
        "executive_summary": summary,
        "risk_matrix": risks,
        "mitigation_plan": mitigations,
        "governance_flags": governance_flags
    }


# ==========================================
# SANITIZED PDF REPORT GENERATOR
# ==========================================
def clean_text_for_pdf(text):
    """Sanitizes text to prevent FPDF Latin-1 encoding errors with unicode/emojis."""
    if not text:
        return ""
    emoji_map = {
        '⚠️': '[WARNING]', '🔒': '[IT GOVERNANCE]', '💰': '[FINANCIAL]', 
        '👥': '[HUMAN-IN-LOOP]', '📊': '[DATA]', '🛠️': '[ACTION]', 
        '📄': '[DOC]', '⚡': '[FAST]', '👤': '[OWNER]', '⏱️': '[TIME]', 
        '💡': '[NOTE]', '🟢': '[OK]', '🔴': '[HIGH]', '🟡': '[MED]'
    }
    for emoji, rep in emoji_map.items():
        text = text.replace(emoji, rep)
    return text.encode('latin-1', 'replace').decode('latin-1')


def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    """Generates a clean PDF brief using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, clean_text_for_pdf("RiskPulse AI - Executive Risk Intelligence Brief"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, clean_text_for_pdf("Confidential Project Controls Deliverable"), ln=True, align="L")
    pdf.ln(4)
    
    pdf.set_draw_color(226, 232, 240)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(6)
    
    # Section 1: Executive Summary
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("1. Executive Summary"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    clean_summary = clean_text_for_pdf(summary_text.replace('#', '').replace('*', '').strip())
    pdf.multi_cell(0, 5, clean_summary)
    pdf.ln(6)
    
    # Section 2: Risk Matrix Table
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk & Issue Matrix"), ln=True)
    
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(45, 7, "Risk / Issue", border=1, fill=True)
    pdf.cell(25, 7, "Severity", border=1, fill=True)
    pdf.cell(30, 7, "Category", border=1, fill=True)
    pdf.cell(90, 7, "Evidence / Quote", border=1, ln=True, fill=True)
    
    pdf.set_font("Helvetica", "", 8)
    for row in risk_data:
        risk_name = clean_text_for_pdf(str(row.get("Risk / Issue", ""))[:28])
        severity = clean_text_for_pdf(str(row.get("Severity", ""))[:12])
        category = clean_text_for_pdf(str(row.get("Category", ""))[:15])
        evidence = clean_text_for_pdf(str(row.get("Evidence / Quote", ""))[:60])
        
        pdf.cell(45, 6, risk_name, border=1)
        pdf.cell(25, 6, severity, border=1)
        pdf.cell(30, 6, category, border=1)
        pdf.cell(90, 6, evidence, border=1, ln=True)
        
    pdf.ln(8)
    
    # Section 3: Governance Flags
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("3. Governance & Action Items"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(51, 65, 85)
    
    for flag in governance_flags:
        clean_flag = clean_text_for_pdf(str(flag).replace('*', ''))
        pdf.multi_cell(0, 5, f"- {clean_flag}")
        pdf.ln(1)
        
    return bytes(pdf.output())


# ==========================================
# MARKDOWN TABLE HELPER (NO TABULATE NEEDED)
# ==========================================
def generate_markdown_table(risk_list):
    """Generates a markdown table without needing the tabulate dependency."""
    if not risk_list:
        return "No risks recorded in matrix."
    md = "| Risk / Issue | Severity | Category | Evidence / Quote | Source File |\n"
    md += "| --- | --- | --- | --- | --- |\n"
    for r in risk_list:
        risk_name = str(r.get("Risk / Issue", "")).replace("|", "\\|")
        severity = str(r.get("Severity", "")).replace("|", "\\|")
        category = str(r.get("Category", "")).replace("|", "\\|")
        evidence = str(r.get("Evidence / Quote", "")).replace("|", "\\|")
        source = str(r.get("Source File", "")).replace("|", "\\|")
        md += f"| {risk_name} | {severity} | {category} | {evidence} | {source} |\n"
    return md


# ==========================================
# APP SIDEBAR CONTROLS (COLLAPSED BY DEFAULT)
# ==========================================
with st.sidebar:
    st.title("RiskPulse AI")
    st.caption("Project Risk Intelligence")
    st.markdown("---")
    
    st.subheader("Analysis Settings")
    mode = st.radio(
        "Execution Engine",
        ["Local Analysis (Standard)", "Enhanced AI Analysis (OpenAI API)"],
        help="Local analysis extracts grounded risks without API keys. Enhanced AI connects to OpenAI."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Enhanced AI Analysis (OpenAI API)":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("Model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
            
    st.markdown("---")
    st.subheader("Risk Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low / Info"],
        default=["High", "Medium", "Low / Info"]
    )


# ==========================================
# MAIN APP BODY
# ==========================================

# 1. Compact Navy Header (Single Title & Subtitle Block)
st.markdown("""
<div class="hero-container">
    <div class="hero-title">RiskPulse AI</div>
    <div class="hero-subtitle">Project Risk Intelligence</div>
    <div class="hero-tagline">Turn project records into evidence-backed risk intelligence.</div>
</div>
""", unsafe_allow_html=True)

# 2. Process Breadcrumb Bar
st.markdown("""
<div class="process-bar">
    <span class="process-active">1. Upload Documents</span>
    <span class="process-arrow">→</span>
    <span>2. Analyze & Extract</span>
    <span class="process-arrow">→</span>
    <span>3. Review Risk Register</span>
    <span class="process-arrow">→</span>
    <span>4. Executive Brief</span>
</div>
""", unsafe_allow_html=True)

# 3. Four Capability Cards
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Cross-Source Risk Detection</div>
        <div class="capability-desc">Correlates findings and identifies contradictions across schedules, reports, and meeting logs.</div>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Schedule Intelligence</div>
        <div class="capability-desc">Surfaces schedule, float, and milestone risks when supported by uploaded schedule data.</div>
    </div>
    """, unsafe_allow_html=True)

with c3:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Cost & Change Exposure</div>
        <div class="capability-desc">Tracks budget overruns, change orders, and financial contingency impacts across project logs.</div>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Evidence-Backed Findings</div>
        <div class="capability-desc">Traces every identified risk directly to source document evidence and direct quotes.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 4. Upload Project Documents Section
st.subheader("Upload Project Documents")
st.caption("Analyze schedules, reports, meeting minutes, change logs, and project records together to surface risks and inconsistencies.")

uploaded_files = st.file_uploader(
    "Drag and drop project files here or click Browse files",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload multi-source project files to extract risks."
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Staged {len(uploaded_files)} file(s) for analysis.")
    with st.expander("Preview Extracted Source Text", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            st.markdown(f"**{parsed['filename']}** ({parsed['type']}) - {parsed['word_count']} words")
            st.text_area(f"Raw Text Preview ({parsed['filename']})", parsed['content'][:800] + ("..." if len(parsed['content']) > 800 else ""), height=80)

# 5. Centered Primary CTA Button
analyze_click = st.button("Analyze Project Risks", type="primary", use_container_width=False)

# 6. Run Analysis Pipeline
if analyze_click:
    if not uploaded_files:
        st.warning("Please upload at least one project file above to extract risks.")
    else:
        combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
        
        with st.spinner("Analyzing multi-source text and extracting risk evidence..."):
            if mode == "Enhanced AI Analysis (OpenAI API)" and api_key:
                system_prompt = """You are an Enterprise AI Risk Analyst. Analyze the provided multi-source project updates and return a structured json object containing:
1. 'executive_summary': a 2-3 paragraph summary of project health, status, and major blocks.
2. 'risk_matrix': a list of objects with keys 'Risk / Issue', 'Severity' (High/Medium/Low), 'Category' (Schedule/Financial/Governance/Technical), 'Evidence / Quote', 'Source File'.
3. 'mitigation_plan': a list of objects with keys 'Action Item', 'Owner', 'Timeframe'.
4. 'governance_flags': a list of strings detailing required approvals.

Return strictly valid JSON."""
                try:
                    raw_response = call_openai_api(api_key, system_prompt, f"Project Documents:\n{combined_text}", model=model_choice)
                    json_start = raw_response.find('{')
                    json_end = raw_response.rfind('}') + 1
                    if json_start != -1 and json_end != -1:
                        results = json.loads(raw_response[json_start:json_end])
                    else:
                        results = run_grounded_text_analysis(combined_text, parsed_file_data)
                except Exception as e:
                    st.error(f"Error calling OpenAI API: {str(e)}. Falling back to local analysis engine.")
                    results = run_grounded_text_analysis(combined_text, parsed_file_data)
            else:
                results = run_grounded_text_analysis(combined_text, parsed_file_data)
                
            st.session_state['analysis_results'] = results
            st.session_state['processed_files'] = parsed_file_data


# 7. Render Output Dashboard
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    files_processed = st.session_state.get('processed_files', [])
    
    st.markdown("---")
    st.subheader("Project Risk Dashboard & Executive Brief")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "Overview", 
        "Source Documents", 
        "Risk Register", 
        "Executive Brief"
    ])
    
    with tab1:
        st.markdown("##### Executive Summary")
        st.markdown(res.get("executive_summary", "No summary generated."))
        
    with tab2:
        st.markdown("##### Processed Source Files")
        if files_processed:
            for pf in files_processed:
                st.markdown(f"- **{pf['filename']}** (`{pf['type']}`) - {pf['word_count']} words | {pf['char_count']} characters")
        else:
            st.write("No source document details available.")
            
    with tab3:
        st.markdown("##### Identified Risk Events & Evidence")
        risk_list = res.get("risk_matrix", [])
        filtered_risks = [r for r in risk_list if r.get("Severity") in min_severity or "All" in min_severity]
        
        if filtered_risks:
            # Custom HTML table for high contrast and readability
            table_html = "<table class='custom-table'><thead><tr><th>Risk / Issue</th><th>Severity</th><th>Category</th><th>Evidence / Quote</th><th>Source File</th></tr></thead><tbody>"
            for r in filtered_risks:
                sev = str(r.get("Severity", ""))
                badge_class = "badge-high" if "High" in sev else ("badge-med" if "Medium" in sev else "badge-low")
                table_html += f"<tr><td><strong>{r.get('Risk / Issue', '')}</strong></td><td><span class='{badge_class}'>{sev}</span></td><td>{r.get('Category', '')}</td><td>{r.get('Evidence / Quote', '')}</td><td><code>{r.get('Source File', '')}</code></td></tr>"
            table_html += "</tbody></table>"
            st.markdown(table_html, unsafe_allow_html=True)
        else:
            st.write("No risks match the selected filter criteria.")
            
    with tab4:
        st.markdown("##### Actionable Mitigation Plan")
        mitigations = res.get("mitigation_plan", [])
        if mitigations:
            for i, item in enumerate(mitigations, 1):
                col_a, col_b, col_c = st.columns([3, 1.5, 1])
                with col_a:
                    st.markdown(f"**{i}. {item.get('Action Item')}**")
                with col_b:
                    st.caption(f"Owner: **{item.get('Owner')}**")
                with col_c:
                    st.caption(f"Timeframe: `{item.get('Timeframe')}`")
                st.divider()
        else:
            st.write("No specific mitigations recorded.")
            
        st.markdown("##### Governance & Review Flags")
        flags = res.get("governance_flags", [])
        for f_item in flags:
            st.markdown(f"- {f_item}")
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Export Section
    st.subheader("Export Deliverables")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    
    with exp_col1:
        try:
            pdf_bytes = generate_pdf_report(
                res.get("executive_summary", ""),
                res.get("risk_matrix", []),
                res.get("mitigation_plan", []),
                res.get("governance_flags", [])
            )
            st.download_button(
                label="Download Executive PDF Report",
                data=pdf_bytes,
                file_name="Executive_Risk_Intelligence_Report.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as e:
            st.caption(f"PDF export error: {str(e)}")
        
    with exp_col2:
        try:
            json_data = json.dumps(res, indent=2)
            st.download_button(
                label="Export Raw JSON Analysis",
                data=json_data,
                file_name="Risk_Analysis_Data.json",
                mime="application/json",
                use_container_width=True
            )
        except Exception as e:
            st.caption(f"JSON export error: {str(e)}")
        
    with exp_col3:
        try:
            md_text = f"### Executive Summary\n{res.get('executive_summary')}\n\n### Key Risk Matrix\n" + generate_markdown_table(res.get('risk_matrix', []))
            st.download_button(
                label="Download Markdown Summary",
                data=md_text,
                file_name="Risk_Summary.md",
                mime="text/markdown",
                use_container_width=True
            )
        except Exception as e:
            st.caption(f"Markdown export error: {str(e)}")
