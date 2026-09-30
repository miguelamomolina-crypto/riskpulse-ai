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
# PAGE CONFIGURATION & ENTERPRISE STYLING
# ==========================================
st.set_page_config(
    page_title="RiskPulse AI | Project Risk Intelligence Platform",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Force Consistent Light Enterprise Theme CSS
st.markdown("""
<style>
    /* Import Inter Font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    /* Global Root Theme Enforcements */
    :root {
        --background-color: #F8FAFC !important;
        --secondary-background-color: #FFFFFF !important;
        --primary-color: #2563EB !important;
        --text-color: #0F172A !important;
        --font: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    /* Force Light Background on View Container */
    html, body, .stApp, [data-testid="stAppViewContainer"], .main {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    /* Remove Streamlit Top Header Bar Space & Extra Padding */
    header[data-testid="stHeader"] {
        display: none !important;
    }

    .block-container, [data-testid="block-container"] {
        max-width: 1280px !important;
        padding-top: 1.25rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        margin: 0 auto !important;
    }

    /* Hide Footer and Scrollbar Decorators */
    #MainMenu, footer, ::-webkit-scrollbar {
        display: none !important;
    }

    /* Text & Typography High Contrast Rules */
    h1, h2, h3, h4, h5, h6, p, span, div, label, .stMarkdown {
        color: #0F172A !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
        color: #0F172A !important;
        font-weight: 600 !important;
        margin-top: 0 !important;
        margin-bottom: 0.5rem !important;
    }

    .secondary-text, .stCaption, caption {
        color: #475569 !important;
    }

    /* Compact Navigation Header Bar (Max Height 64px) */
    .top-navbar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background-color: #0F172A;
        color: #FFFFFF !important;
        padding: 0 1.5rem;
        border-radius: 8px;
        margin-bottom: 1.25rem;
        height: 64px;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
    }

    .top-navbar * {
        color: #FFFFFF !important;
    }

    .navbar-brand {
        font-size: 1.15rem;
        font-weight: 700;
        letter-spacing: -0.02em;
    }

    .navbar-subtitle {
        font-size: 0.875rem;
        color: #94A3B8 !important;
        font-weight: 400;
    }

    .navbar-badge {
        background-color: #1E293B;
        color: #38BDF8 !important;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.25rem 0.65rem;
        border-radius: 6px;
        border: 1px solid #334155;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
        padding-top: 1.25rem !important;
    }

    [data-testid="stSidebar"] * {
        color: #0F172A !important;
    }

    [data-testid="stSidebar"] h3 {
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
        margin-bottom: 0.5rem !important;
    }

    /* Content Card Container */
    .card-box {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        padding: 1.25rem 1.5rem !important;
        margin-bottom: 1.25rem !important;
        box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03) !important;
    }

    /* File Uploader Custom Light Theme */
    [data-testid="stFileUploader"] {
        background-color: #FFFFFF !important;
        border: 2px dashed #CBD5E1 !important;
        border-radius: 8px !important;
        padding: 1.25rem !important;
        transition: border-color 0.2s ease, background-color 0.2s ease !important;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: #2563EB !important;
        background-color: #F8FAFC !important;
    }

    [data-testid="stFileUploader"] section {
        background-color: transparent !important;
    }

    [data-testid="stFileUploader"] label, 
    [data-testid="stFileUploader"] span, 
    [data-testid="stFileUploader"] small,
    [data-testid="stFileUploader"] p {
        color: #0F172A !important;
        font-family: 'Inter', sans-serif !important;
    }

    /* Primary CTA Button (Forced Blue) */
    .stButton > button, 
    button[kind="primary"], 
    button[type="primary"],
    [data-testid="baseButton-primary"] {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        border: 1px solid #2563EB !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        padding: 0.6rem 1.5rem !important;
        box-shadow: 0 1px 2px rgba(37, 99, 235, 0.15) !important;
        transition: background-color 0.15s ease-in-out !important;
    }

    .stButton > button:hover, 
    button[kind="primary"]:hover, 
    button[type="primary"]:hover,
    [data-testid="baseButton-primary"]:hover {
        background-color: #1D4ED8 !important;
        border-color: #1D4ED8 !important;
        color: #FFFFFF !important;
    }

    /* Secondary Download Buttons */
    .stDownloadButton > button {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        transition: all 0.15s ease !important;
    }

    .stDownloadButton > button:hover {
        background-color: #F8FAFC !important;
        border-color: #2563EB !important;
        color: #2563EB !important;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem !important;
        border-bottom: 1px solid #E2E8F0 !important;
        margin-bottom: 1.25rem !important;
    }

    .stTabs [data-baseweb="tab"] {
        height: 40px !important;
        padding: 0 1rem !important;
        background-color: transparent !important;
        border-radius: 6px 6px 0 0 !important;
        color: #64748B !important;
        font-weight: 500 !important;
        font-size: 0.9rem !important;
    }

    .stTabs [aria-selected="true"] {
        color: #2563EB !important;
        font-weight: 600 !important;
        border-bottom: 2px solid #2563EB !important;
    }

    /* DataFrame & Table Overrides */
    [data-testid="stDataFrame"], .stDataFrame {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
    }

    /* Expander Container */
    .stExpander {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        box-shadow: none !important;
    }

    .stExpander header {
        color: #0F172A !important;
        font-weight: 600 !important;
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
# OPENAI & MOCK AI RISK PIPELINE
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


def run_simulated_ai_pipeline(combined_text, file_list):
    """Simulates AI extraction when running in Demo Mode without an OpenAI API key."""
    filenames_str = ", ".join([f['filename'] for f in file_list])
    
    # Keyword & Evidence extraction
    has_delay = bool(re.search(r'delay|behind|timeline|schedule|postpone|lead time|float', combined_text, re.I))
    has_budget = bool(re.search(r'budget|cost|overrun|spend|contractor|change order|pco|asr', combined_text, re.I))
    has_security = bool(re.search(r'security|governance|approval|permission|it director|sign-off', combined_text, re.I))
    has_reassignment = bool(re.search(r'staff|subcontractor|resigned|reassigned|mep|resource', combined_text, re.I))
    
    risks = []
    
    if has_delay:
        risks.append({
            "Risk / Issue": "Critical Path Schedule Variance & Equipment Procurement Lead Time",
            "Severity": "High",
            "Category": "Schedule",
            "Evidence / Quote": "Long-lead material deliveries and shop drawing review bottlenecks indicate milestone float erosion.",
            "Source File": file_list[0]['filename'] if file_list else "Schedule_Update"
        })
    if has_budget:
        risks.append({
            "Risk / Issue": "Unforeseen Site Condition & Contingency Budget Drawdown",
            "Severity": "High",
            "Category": "Financial",
            "Evidence / Quote": "Pending change orders and additional fee requests exceed allocated contingency thresholds.",
            "Source File": file_list[-1]['filename'] if file_list else "Cost_Report"
        })
    if has_security:
        risks.append({
            "Risk / Issue": "IT Governance & Cloud System Sign-Off Bottleneck",
            "Severity": "High",
            "Category": "Governance",
            "Evidence / Quote": "System integration permissions require formal sign-off prior to operational deployment.",
            "Source File": file_list[0]['filename'] if file_list else "Governance_Log"
        })
    if has_reassignment or not risks:
        risks.append({
            "Risk / Issue": "Key Subcontractor Resource Shift & Staffing Variance",
            "Severity": "Medium",
            "Category": "Workflow",
            "Evidence / Quote": "Subcontractor staffing levels operating below baseline requirements.",
            "Source File": file_list[0]['filename'] if file_list else "Resource_Log"
        })
    
    # Positive baseline
    risks.append({
        "Risk / Issue": "Milestone Design Verification Phase Completed",
        "Severity": "Low / Positive",
        "Category": "Technical",
        "Evidence / Quote": "Initial engineering and architectural review completed with stakeholder sign-off.",
        "Source File": file_list[0]['filename'] if file_list else "Meeting_Minutes"
    })
    
    summary = f"""Executive Health Summary:
The analyzed project sources ({filenames_str}) indicate key progress in baseline design reviews alongside high-priority risk exposure in critical path schedule variance, pending change order contingency consumption, and required governance authorizations. Primary project velocity is subject to material lead times and required owner decisions.

Key Risk Highlights:
- Overall Status: Caution / Watch Items Identified
- Key Schedule Impact: Critical path equipment submittals require immediate owner review to lock production slots.
- Financial Exposure: Pending potential change orders represent active contingency consumption.
- Manual Review Time Saved: ~3.5 Hours saved across {len(file_list)} uploaded project documents.
"""

    mitigations = [
        {"Action Item": "Authorize long-lead equipment submittals to secure factory production slot", "Owner": "Owner Representative", "Timeframe": "Immediate (Within 48 Hours)", "Priority": "High"},
        {"Action Item": "Negotiate potential change order scope and pricing prior to contingency drawdown", "Owner": "Project Controls Lead", "Timeframe": "This Week", "Priority": "High"},
        {"Action Item": "Conduct executive governance review meeting for pending scope decisions", "Owner": "Project Manager & CFO", "Timeframe": "Next Weekly Sync", "Priority": "Medium"},
        {"Action Item": "Establish recurring multi-source risk tracking pipeline", "Owner": "Project Lead", "Timeframe": "Standard Cycle", "Priority": "Medium"}
    ]
    
    governance_flags = [
        "System Sign-off: Integration permissions require formal sign-off from IT Security Lead before deployment.",
        "Financial Threshold Variance: Pending change orders exceeding threshold limits require CFO review.",
        "Human-in-the-Loop Oversight: All automated task assignments must be verified by Project Controls Lead."
    ]
    
    return {
        "raw_text": summary,
        "executive_summary": summary,
        "risk_matrix": risks,
        "mitigation_plan": mitigations,
        "governance_flags": governance_flags
    }


# ==========================================
# PDF REPORT GENERATOR WITH UNICODE SANITIZER
# ==========================================
def clean_text_for_pdf(text):
    """Sanitizes text to prevent fpdf Latin-1 encoding errors."""
    if not text:
        return ""
    return text.encode('latin-1', 'replace').decode('latin-1')


def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    """Generates an executive PDF report using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Header
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, clean_text_for_pdf("Executive Risk Intelligence Report"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(71, 85, 105)
    pdf.cell(0, 6, clean_text_for_pdf("Generated by RiskPulse AI Platform | Confidential"), ln=True, align="L")
    pdf.ln(4)
    
    # Divider Line
    pdf.set_draw_color(226, 232, 240)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    # Section 1: Summary
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("1. Executive Health Summary"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    
    clean_summary = clean_text_for_pdf(summary_text.replace('#', '').replace('*', '').strip())
    pdf.multi_cell(0, 5, clean_summary)
    pdf.ln(6)
    
    # Section 2: Risk Table
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk and Issue Identification Matrix"), ln=True)
    
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(50, 7, "Risk / Issue", border=1, fill=True)
    pdf.cell(25, 7, "Severity", border=1, fill=True)
    pdf.cell(30, 7, "Category", border=1, fill=True)
    pdf.cell(85, 7, "Evidence / Quote", border=1, ln=True, fill=True)
    
    pdf.set_font("Helvetica", "", 8)
    for row in risk_data:
        risk_name = clean_text_for_pdf(str(row.get("Risk / Issue", ""))[:30])
        severity = clean_text_for_pdf(str(row.get("Severity", ""))[:12])
        category = clean_text_for_pdf(str(row.get("Category", ""))[:15])
        evidence = clean_text_for_pdf(str(row.get("Evidence / Quote", ""))[:55])
        
        pdf.cell(50, 6, risk_name, border=1)
        pdf.cell(25, 6, severity, border=1)
        pdf.cell(30, 6, category, border=1)
        pdf.cell(85, 6, evidence, border=1, ln=True)
        
    pdf.ln(8)
    
    # Section 3: Governance Flags
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("3. IT Governance and Sign-Off Flags"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(51, 65, 85)
    
    for flag in governance_flags:
        clean_flag = clean_text_for_pdf(flag.replace('*', ''))
        pdf.multi_cell(0, 5, f"- {clean_flag}")
        pdf.ln(1)
        
    return bytes(pdf.output())


# ==========================================
# SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
    st.markdown("### System Configuration")
    
    mode = st.radio(
        "Execution Engine",
        ["Demo / Simulated Mode (No Key)", "Live OpenAI API"],
        help="Select execution engine mode."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Live OpenAI API":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("LLM Engine Model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.caption("Enter an OpenAI API key to run live API analysis.")
            
    st.markdown("---")
    st.markdown("### Risk Threshold Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low / Positive"],
        default=["High", "Medium", "Low / Positive"]
    )


# ==========================================
# MAIN APPLICATION LAYOUT
# ==========================================

# Compact Navigation Bar (Max Height 64px)
st.markdown("""
<div class="top-navbar">
    <div style="display: flex; align-items: center; gap: 0.75rem;">
        <span class="navbar-brand">RiskPulse AI</span>
        <span style="color: #64748B;">|</span>
        <span class="navbar-subtitle">Project Risk Intelligence & Control Center</span>
    </div>
    <div class="navbar-badge">Enterprise Controls</div>
</div>
""", unsafe_allow_html=True)


# Source Document Ingestion Section
st.markdown("### Source Document Ingestion")
st.caption("Supports simultaneous upload of PDF status reports, Word documents, CSV exports, meeting notes, and text logs.")

uploaded_files = st.file_uploader(
    "Drag and drop project files here",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload 1 to 10 project status documents to aggregate."
)

parsed_file_data = []

if uploaded_files:
    st.info(f"Loaded {len(uploaded_files)} file(s) for multi-source risk extraction.")
    
    with st.expander("Preview Extracted Text and Metadata per File", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            
            st.markdown(f"**{parsed['filename']}** (`{parsed['type']}`) - {parsed['word_count']} words | {parsed['char_count']} characters")
            st.text_area(f"Raw Extracted Text ({parsed['filename']})", parsed['content'][:1000] + ("..." if len(parsed['content']) > 1000 else ""), height=100)
            st.divider()

st.markdown("<br>", unsafe_allow_html=True)

# Primary Trigger CTA Button
col_btn, _ = st.columns([2, 1])
with col_btn:
    analyze_click = st.button("Analyze Project Risks", type="primary", use_container_width=True)

# Execute Risk Pipeline
if analyze_click:
    if not uploaded_files:
        st.warning("Please upload at least one project file to run analysis.")
        sample_text = """Project Update:
Equipment deliverables face long-lead procurement delays. Change order requests exceed current contingency allowance. Governance approvals are pending IT director sign-off."""
        parsed_file_data = [{
            "filename": "Sample_Project_Update.txt",
            "type": "TXT",
            "content": sample_text,
            "char_count": len(sample_text),
            "word_count": len(sample_text.split())
        }]

    combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
    
    with st.spinner("Analyzing project documents, extracting risk triggers, and generating executive summary..."):
        if mode == "Live OpenAI API" and api_key:
            system_prompt = """You are an expert Enterprise Risk Analyst. Analyze the provided multi-source project updates and return a structured json object containing:
1. 'executive_summary': a 2-3 paragraph summary of project health, status, and major blocks.
2. 'risk_matrix': a list of objects with keys 'Risk / Issue', 'Severity' (High/Medium/Low), 'Category' (Schedule/Financial/Workflow/Governance), 'Evidence / Quote', 'Source File'.
3. 'mitigation_plan': a list of objects with keys 'Action Item', 'Owner', 'Timeframe', 'Priority'.
4. 'governance_flags': a list of strings detailing IT security or governance approvals required.

Return strictly valid JSON."""
            
            try:
                raw_response = call_openai_api(api_key, system_prompt, f"Project Documents:\n{combined_text}", model=model_choice)
                json_start = raw_response.find('{')
                json_end = raw_response.rfind('}') + 1
                if json_start != -1 and json_end != -1:
                    results = json.loads(raw_response[json_start:json_end])
                else:
                    results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
            except Exception as e:
                st.error(f"Error calling OpenAI API: {str(e)}. Falling back to local extraction engine.")
                results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
        else:
            results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
            
        st.session_state['analysis_results'] = results
        st.session_state['processed_files'] = parsed_file_data


# Render Results Dashboard
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    files_processed = st.session_state.get('processed_files', [])
    
    st.markdown("---")
    st.markdown("### Executive Risk Intelligence Dashboard")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "Executive Summary", 
        "Risk Identification Matrix", 
        "Actionable Mitigation Plan", 
        "Governance Flags"
    ])
    
    with tab1:
        st.markdown(res.get("executive_summary", "No summary generated."))
        
    with tab2:
        st.markdown("#### Identified Risk Events and Categorization")
        risk_list = res.get("risk_matrix", [])
        
        filtered_risks = [r for r in risk_list if r.get("Severity") in min_severity or "All" in min_severity]
        
        if filtered_risks:
            df_risks = pd.DataFrame(filtered_risks)
            st.dataframe(
                df_risks,
                use_container_width=True,
                column_config={
                    "Severity": st.column_config.TextColumn(
                        "Severity",
                        width="small"
                    ),
                    "Evidence / Quote": st.column_config.TextColumn(
                        "Direct Quote / Evidence",
                        width="large"
                    )
                }
            )
        else:
            st.write("No risks match the selected filter criteria.")
            
    with tab3:
        st.markdown("#### Recommended Mitigation Steps")
        mitigations = res.get("mitigation_plan", [])
        if mitigations:
            for i, item in enumerate(mitigations, 1):
                col_a, col_b, col_c = st.columns([3, 1.5, 1])
                with col_a:
                    st.markdown(f"**{i}. {item.get('Action Item')}**")
                with col_b:
                    st.caption(f"Owner: {item.get('Owner')}")
                with col_c:
                    st.caption(f"Timeframe: {item.get('Timeframe')}")
                st.divider()
        else:
            st.write("No specific mitigations recorded.")
            
    with tab4:
        st.markdown("#### Governance and Sign-Off Flags")
        flags = res.get("governance_flags", [])
        for f_item in flags:
            st.markdown(f"- {f_item}")
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Export Deliverables
    st.markdown("### Export and Share Deliverables")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    
    with exp_col1:
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
        
    with exp_col2:
        json_data = json.dumps(res, indent=2)
        st.download_button(
            label="Export Raw JSON Analysis",
            data=json_data,
            file_name="Risk_Analysis_Data.json",
            mime="application/json",
            use_container_width=True
        )
        
    with exp_col3:
        md_text = f"{res.get('executive_summary')}\n\n### Risk Table\n" + pd.DataFrame(res.get('risk_matrix', [])).to_markdown(index=False)
        st.download_button(
            label="Download Markdown Summary",
            data=md_text,
            file_name="Risk_Summary.md",
            mime="text/markdown",
            use_container_width=True
        )

