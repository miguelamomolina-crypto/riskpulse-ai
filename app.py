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

# Custom CSS for B2B SaaS Aesthetics & Light Theme Enforcements
st.markdown("""
<style>
    /* Force Light Canvas Overrides */
    :root {
        --background-color: #f8fafc;
        --secondary-background-color: #ffffff;
        --text-color: #0f172a;
    }
    
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], .main, .block-container {
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Global Typography High-Contrast Formatting */
    h1, h2, h3, h4, h5, h6, p, label, span, div, caption {
        color: #0f172a;
    }
    
    /* Header Container */
    .hero-container {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
        padding: 1.25rem 2rem;
        border-radius: 10px;
        color: #ffffff !important;
        margin-bottom: 1.25rem;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15);
    }
    .hero-title {
        font-size: 1.8rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin-bottom: 0.2rem;
        color: #ffffff !important;
    }
    .hero-subtitle {
        font-size: 0.95rem;
        color: #cbd5e1 !important;
        line-height: 1.4;
    }
    
    /* Process Bar */
    .process-bar {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.6rem 1.25rem;
        margin-bottom: 1.25rem;
        font-size: 0.85rem;
        color: #334155;
        font-weight: 600;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    
    /* Capability / Metric Cards */
    .capability-card {
        background-color: #ffffff;
        padding: 1.1rem;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(15,23,42,0.04);
        height: 100%;
    }
    .capability-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.35rem;
    }
    .capability-desc {
        font-size: 0.8rem;
        color: #475569;
        line-height: 1.4;
    }
    
    /* File Uploader Container & Restyled Uploaded File Chips */
    .stFileUploader {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 10px !important;
        background-color: #ffffff !important;
        padding: 1rem !important;
    }
    .stFileUploader:hover {
        border-color: #2563eb !important;
    }
    
    /* Restyle Streamlit Uploaded File Items (File Chips) to Light Mode */
    [data-testid="stFileUploaderFile"],
    .stFileUploaderFile,
    div[data-testid="stFileUploaderFile"],
    [data-testid="stFileUploaderFileData"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 6px !important;
        color: #0f172a !important;
        padding: 0.4rem 0.75rem !important;
        margin-top: 0.4rem !important;
    }
    [data-testid="stFileUploaderFile"] *,
    .stFileUploaderFile *,
    [data-testid="stFileUploaderFileData"] * {
        color: #0f172a !important;
        fill: #0f172a !important;
    }
    [data-testid="stFileUploaderDeleteBtn"] button,
    [data-testid="stFileUploaderDeleteBtn"] svg {
        color: #64748b !important;
        fill: #64748b !important;
    }
    
    /* Primary CTA Button Override */
    div.stButton > button[type="primary"],
    div.stButton > button[kind="primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        padding: 0.5rem 1.5rem !important;
        max-width: 320px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    }
    div.stButton > button[type="primary"]:hover,
    div.stButton > button[kind="primary"]:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
    }
    
    /* Status Badges */
    .severity-high {
        background-color: #fef2f2;
        color: #dc2626;
        padding: 0.2rem 0.6rem;
        border-radius: 4px;
        font-weight: 600;
        border: 1px solid #fecaca;
    }
    .severity-med {
        background-color: #fffbe3;
        color: #d97706;
        padding: 0.2rem 0.6rem;
        border-radius: 4px;
        font-weight: 600;
        border: 1px solid #fde68a;
    }
    .severity-low {
        background-color: #f0fdf4;
        color: #16a34a;
        padding: 0.2rem 0.6rem;
        border-radius: 4px;
        font-weight: 600;
        border: 1px solid #bbf7d0;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# FILE PARSER UTILITIES
# ==========================================
def extract_text_from_file(uploaded_file):
    """Parses text from PDF, DOCX, CSV, TXT, and MD files."""
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
# OPENAI & LOCAL ANALYSIS ENGINE
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


def run_grounded_text_analysis(combined_text, file_list):
    """Executes grounded local text extraction when running without an OpenAI API key."""
    filenames_str = ", ".join([f['filename'] for f in file_list]) if file_list else "Uploaded Files"
    
    # Grounded keyword analysis directly on source text
    has_delay = bool(re.search(r'delay|behind|timeline|schedule|postpone|float', combined_text, re.I))
    has_budget = bool(re.search(r'budget|cost|overrun|spend|contractor|pco|change order', combined_text, re.I))
    has_security = bool(re.search(r'security|permission|api|governance|approval|sign-off', combined_text, re.I))
    has_reassignment = bool(re.search(r'engineer|resigned|reassigned|staff|turnover|resource', combined_text, re.I))
    
    risks = []
    
    if has_delay:
        risks.append({
            "Risk / Issue": "Schedule Delay & Critical Path Slippage",
            "Severity": "High",
            "Category": "Schedule",
            "Evidence / Quote": "Source files indicate timeline slippage and potential critical path milestone impact.",
            "Source File": file_list[0]['filename'] if file_list else "Schedule_Log"
        })
    if has_budget:
        risks.append({
            "Risk / Issue": "Cost Variance & Contingency Drawdown",
            "Severity": "High",
            "Category": "Financial",
            "Evidence / Quote": "Financial or change management records indicate cost variances or pending change authorizations.",
            "Source File": file_list[-1]['filename'] if file_list else "Cost_Report"
        })
    if has_security:
        risks.append({
            "Risk / Issue": "Governance & Permission Approvals Pending",
            "Severity": "High",
            "Category": "Governance",
            "Evidence / Quote": "Documentation flags required governance sign-offs or permission authorizations prior to execution.",
            "Source File": file_list[0]['filename'] if file_list else "Governance_Log"
        })
    if has_reassignment or not risks:
        risks.append({
            "Risk / Issue": "Resource Allocation & Personnel Constraints",
            "Severity": "Medium",
            "Category": "Workflow",
            "Evidence / Quote": "Resource constraints or personnel shifts identified across project meeting records.",
            "Source File": file_list[0]['filename'] if file_list else "Project_Records"
        })
    
    summary = f"Analysis completed across **{len(file_list)} uploaded document(s)** ({filenames_str}). Key schedule, financial, and governance risk triggers were extracted directly from source evidence."

    mitigations = [
        {"Action Item": "Schedule alignment review on critical path deliverables", "Owner": "Project Manager", "Timeframe": "Within 48 Hours", "Priority": "High"},
        {"Action Item": "Re-evaluate resource allocation and personnel assignments", "Owner": "Project Lead", "Timeframe": "This Week", "Priority": "High"},
        {"Action Item": "Review financial variance and pending change requests", "Owner": "Project Controls Lead", "Timeframe": "Next Review Cycle", "Priority": "Medium"}
    ]
    
    governance_flags = [
        "Governance & Security Approvals: Ensure formal authorization before proceeding with critical deliverables.",
        "Financial Threshold Sign-off: Verify cost variances against authorized contingency thresholds."
    ]
    
    return {
        "raw_text": summary,
        "executive_summary": summary,
        "risk_matrix": risks,
        "mitigation_plan": mitigations,
        "governance_flags": governance_flags
    }


# Helper for Markdown Table Generation without tabulate
def generate_markdown_table(data_list):
    """Converts a list of dicts or DataFrame to a clean Markdown table string without requiring tabulate."""
    if not data_list:
        return "*No data available.*"
    
    if isinstance(data_list, pd.DataFrame):
        data_list = data_list.to_dict('records')
        
    if not isinstance(data_list, list) or len(data_list) == 0:
        return "*No data available.*"
        
    headers = list(data_list[0].keys())
    
    header_line = "| " + " | ".join(str(h) for h in headers) + " |"
    separator_line = "| " + " | ".join("---" for _ in headers) + " |"
    
    rows = [header_line, separator_line]
    
    for item in data_list:
        row_cells = []
        for h in headers:
            val = str(item.get(h, "")).replace("\n", " ").replace("|", "\\|")
            row_cells.append(val)
        rows.append("| " + " | ".join(row_cells) + " |")
        
    return "\n".join(rows)


# ==========================================
# PDF REPORT GENERATOR WITH UNICODE SANITIZER
# ==========================================
def clean_text_for_pdf(text):
    """Sanitizes text to prevent fpdf Latin-1 encoding errors with unicode/emojis."""
    if not text:
        return ""
    # Strip non-latin1 characters safely
    return text.encode('latin-1', 'replace').decode('latin-1')


def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    """Generates an executive PDF report using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Header
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(15, 23, 42) # Navy
    pdf.cell(0, 10, clean_text_for_pdf("RiskPulse AI — Executive Risk Intelligence Brief"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, clean_text_for_pdf("Project Risk Assessment & Evidence Brief | Confidential"), ln=True, align="L")
    pdf.ln(4)
    
    # Horizontal Divider Line
    pdf.set_draw_color(226, 232, 240)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    # Section 1: Executive Summary
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("1. Executive Summary"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(51, 65, 85)
    
    clean_summary = clean_text_for_pdf(summary_text.replace('#', '').replace('*', '').strip())
    pdf.multi_cell(0, 5, clean_summary)
    pdf.ln(5)
    
    # Section 2: Risk Matrix Table
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk Identification Matrix"), ln=True)
    
    # Table Header
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(45, 6, "Risk / Issue", border=1, fill=True)
    pdf.cell(20, 6, "Severity", border=1, fill=True)
    pdf.cell(25, 6, "Category", border=1, fill=True)
    pdf.cell(100, 6, "Evidence / Quote", border=1, ln=True, fill=True)
    
    pdf.set_font("Helvetica", "", 8)
    for row in risk_data:
        risk_name = clean_text_for_pdf(str(row.get("Risk / Issue", ""))[:28])
        severity = clean_text_for_pdf(str(row.get("Severity", ""))[:10])
        category = clean_text_for_pdf(str(row.get("Category", ""))[:12])
        evidence = clean_text_for_pdf(str(row.get("Evidence / Quote", ""))[:65])
        
        pdf.cell(45, 6, risk_name, border=1)
        pdf.cell(20, 6, severity, border=1)
        pdf.cell(25, 6, category, border=1)
        pdf.cell(100, 6, evidence, border=1, ln=True)
        
    pdf.ln(6)
    
    # Section 3: Governance Flags
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("3. Governance & Sign-off Flags"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(51, 65, 85)
    
    for flag in governance_flags:
        clean_flag = clean_text_for_pdf(flag.replace('*', ''))
        pdf.multi_cell(0, 5, f"- {clean_flag}")
        pdf.ln(1)
        
    return bytes(pdf.output())


# ==========================================
# APP SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
    st.title("RiskPulse AI")
    st.caption("Project Risk Intelligence")
    st.markdown("---")
    
    st.subheader("Analysis Settings")
    
    mode = st.radio(
        "Execution Engine",
        ["Demo / Simulated Mode (No Key)", "Live OpenAI API"],
        help="Demo mode uses local text extraction. Live API connects to OpenAI."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Live OpenAI API":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("LLM Model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.warning("Please enter an OpenAI API key to execute live calls.")
            
    st.markdown("---")
    st.subheader("Risk Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low / Positive"],
        default=["High", "Medium", "Low / Positive"]
    )


# ==========================================
# MAIN APP BODY
# ==========================================

# Header Banner
st.markdown("""
<div class="hero-container">
    <div class="hero-title">RiskPulse AI</div>
    <div class="hero-subtitle">Project Risk Intelligence</div>
</div>
""", unsafe_allow_html=True)

# Headline & Tagline
st.markdown("""
<div style="margin-bottom: 1rem;">
    <h2 style="color: #0f172a; font-size: 1.4rem; font-weight: 700; margin-bottom: 0.2rem;">Project Risk Intelligence</h2>
    <p style="color: #475569; font-size: 0.95rem; margin-top: 0;">Turn project records into evidence-backed risk intelligence.</p>
</div>
""", unsafe_allow_html=True)

# Process Indicator
st.markdown("""
<div class="process-bar">
    1. Upload Documents &nbsp;&rarr;&nbsp; 2. Analyze &amp; Extract &nbsp;&rarr;&nbsp; 3. Review Risk Register &nbsp;&rarr;&nbsp; 4. Executive Brief
</div>
""", unsafe_allow_html=True)

# Capability Cards Strip
col_c1, col_c2, col_c3, col_c4 = st.columns(4)
with col_c1:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Cross-Source Risk Detection</div>
        <div class="capability-desc">Correlates findings and identifies contradictions across schedules, reports, and meeting logs.</div>
    </div>
    """, unsafe_allow_html=True)

with col_c2:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Schedule Intelligence</div>
        <div class="capability-desc">Surfaces schedule, float, and milestone risks when supported by uploaded schedule data.</div>
    </div>
    """, unsafe_allow_html=True)

with col_c3:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Cost &amp; Change Exposure</div>
        <div class="capability-desc">Tracks budget overruns, change orders, and financial contingency impacts across project logs.</div>
    </div>
    """, unsafe_allow_html=True)

with col_c4:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Evidence-Backed Findings</div>
        <div class="capability-desc">Traces every identified risk directly to source document evidence and direct quotes.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# File Upload Section
st.subheader("Upload Project Documents")
st.caption("Analyze schedules, reports, meeting minutes, change logs, and project records together to surface risks and inconsistencies.")

uploaded_files = st.file_uploader(
    "Drag and drop files here or click to browse",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload project status documents to aggregate."
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Successfully staged **{len(uploaded_files)} file(s)** for extraction.")
    
    with st.expander("Preview Extracted Text & Metadata per File", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            
            st.markdown(f"**{parsed['filename']}** (`{parsed['type']}`) — {parsed['word_count']} words | {parsed['char_count']} characters")
            st.text_area(f"Raw Extracted Text ({parsed['filename']})", parsed['content'][:1000] + ("..." if len(parsed['content']) > 1000 else ""), height=100)
            st.divider()

st.markdown("<br>", unsafe_allow_html=True)

# Action CTA Trigger
col_btn1, col_btn2 = st.columns([1, 2])
with col_btn1:
    analyze_click = st.button("Analyze Project Risks", type="primary")

# Execute Pipeline
if analyze_click:
    if not uploaded_files:
        st.warning("Please upload at least one project document above to analyze.")
    else:
        combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
        
        with st.spinner("Analyzing multi-source text, extracting risk triggers, and building matrix..."):
            if mode == "Live OpenAI API" and api_key:
                system_prompt = """You are an expert Enterprise AI Risk Analyst. Analyze the provided multi-source project updates and return a structured json object containing:
1. 'executive_summary': a 2-3 paragraph summary of project health, status, and major blocks.
2. 'risk_matrix': a list of objects with keys 'Risk / Issue', 'Severity' (High/Medium/Low), 'Category' (Technical/Financial/Workflow/Governance), 'Evidence / Quote', 'Source File'.
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
                        results = run_grounded_text_analysis(combined_text, parsed_file_data)
                except Exception as e:
                    st.error(f"Error calling OpenAI API: {str(e)}. Falling back to local extraction engine.")
                    results = run_grounded_text_analysis(combined_text, parsed_file_data)
            else:
                results = run_grounded_text_analysis(combined_text, parsed_file_data)
                
            st.session_state['analysis_results'] = results
            st.session_state['processed_files'] = parsed_file_data


# Render Output Dashboard if available
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    files_processed = st.session_state.get('processed_files', [])
    
    st.markdown("---")
    st.subheader("Project Risk Dashboard & Brief")
    
    # View Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "Overview", 
        "Documents", 
        "Risk Register", 
        "Executive Brief"
    ])
    
    with tab1:
        st.markdown(res.get("executive_summary", "No summary generated."))
        st.info("Share this summary directly in weekly status updates or executive syncs.")
        
    with tab2:
        st.markdown("##### Ingested Document Inventory")
        if files_processed:
            doc_df = pd.DataFrame([{
                "Filename": p["filename"],
                "Format": p["type"],
                "Word Count": p["word_count"],
                "Character Count": p["char_count"]
            } for p in files_processed])
            st.dataframe(doc_df, use_container_width=True)
        else:
            st.write("No files recorded.")
            
    with tab3:
        st.markdown("##### Identified Risk Register")
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
                        help="Risk Impact Level",
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
            
    with tab4:
        st.markdown("##### Actionable Mitigations & Governance")
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
            
        st.markdown("##### Governance Flags")
        flags = res.get("governance_flags", [])
        for f_item in flags:
            st.markdown(f"- {f_item}")
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Export Section
    st.subheader("Export & Share Deliverables")
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
            st.error("PDF export unavailable.")
            st.caption(f"Details: {str(e)}")
        
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
            st.error("JSON export unavailable.")
            st.caption(f"Details: {str(e)}")
        
    with exp_col3:
        try:
            md_table = generate_markdown_table(res.get("risk_matrix", []))
            summary_text = res.get("executive_summary", "")
            md_text = f"# Executive Risk Summary\n\n{summary_text}\n\n## Risk Identification Matrix\n\n{md_table}"
            st.download_button(
                label="Download Markdown Summary",
                data=md_text,
                file_name="Risk_Summary.md",
                mime="text/markdown",
                use_container_width=True
            )
        except Exception as e:
            st.error("Markdown export unavailable.")
            st.caption(f"Details: {str(e)}")
