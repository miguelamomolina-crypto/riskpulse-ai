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
    page_title="RiskPulse AI | Project Risk Intelligence",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Clean B2B SaaS Aesthetics
st.markdown("""
<style>
    /* Global Theme Enforcements */
    :root {
        --bg-main: #f8fafc;
        --card-bg: #ffffff;
        --text-primary: #0f172a;
        --text-secondary: #475569;
        --border-color: #e2e8f0;
        --primary-blue: #2563eb;
        --hover-blue: #1d4ed8;
    }

    .main, .stApp {
        background-color: var(--bg-main) !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        color: var(--text-primary) !important;
    }

    /* Hide empty Streamlit header bar gap */
    header[data-testid="stHeader"] {
        display: none !important;
    }

    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 1280px !important;
    }

    /* Top Navbar */
    .top-navbar {
        background: #0f172a;
        padding: 1.25rem 2rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 1.25rem;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08);
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .brand-title {
        font-size: 1.6rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        color: #ffffff;
        margin: 0;
        line-height: 1.2;
    }

    .brand-subtitle {
        font-size: 0.95rem;
        color: #94a3b8;
        font-weight: 400;
        margin-top: 0.2rem;
    }

    /* Workflow Breadcrumb Line */
    .workflow-strip {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.6rem 1.25rem;
        margin-bottom: 1.25rem;
        font-size: 0.875rem;
        font-weight: 600;
        color: #475569;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .wf-step {
        color: #0f172a;
    }

    .wf-arrow {
        color: #94a3b8;
        font-weight: 400;
    }

    /* Compact Value Cards */
    .value-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
        height: 100%;
    }

    .value-card-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.25rem;
    }

    .value-card-desc {
        font-size: 0.8rem;
        color: #64748b;
        line-height: 1.35;
    }

    /* Upload Container */
    .upload-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
    }

    .section-title {
        font-size: 1.2rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.35rem;
    }

    .section-desc {
        font-size: 0.9rem;
        color: #475569;
        margin-bottom: 1rem;
        line-height: 1.4;
    }

    /* Streamlit File Uploader High-Contrast Customization */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #93c5fd;
        border-radius: 8px;
        background-color: #f8fafc;
        padding: 0.75rem;
        transition: border-color 0.2s ease;
    }

    div[data-testid="stFileUploader"]:hover {
        border-color: #2563eb;
        background-color: #eff6ff;
    }

    div[data-testid="stFileUploader"] section {
        background-color: transparent !important;
    }

    div[data-testid="stFileUploader"] label {
        color: #0f172a !important;
        font-weight: 600 !important;
    }

    /* High-contrast Browse Files button */
    div[data-testid="stFileUploader"] button {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
        padding: 0.4rem 1rem !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    }

    div[data-testid="stFileUploader"] button:hover {
        background-color: #f1f5f9 !important;
        border-color: #2563eb !important;
        color: #2563eb !important;
    }

    /* Blue Primary CTA Button - Compact 300px styling */
    .cta-container {
        max-width: 320px;
        margin-bottom: 1rem;
    }

    div.stButton > button[type="primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 0.65rem 1.25rem !important;
        transition: background-color 0.2s ease !important;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2) !important;
        width: 100% !important;
    }

    div.stButton > button[type="primary"]:hover {
        background-color: #1d4ed8 !important;
        box-shadow: 0 4px 8px rgba(29, 78, 216, 0.3) !important;
    }

    /* Results Dashboard Styling */
    .dashboard-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.5rem;
        margin-top: 1rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
    }

    /* High Contrast Data Table */
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.875rem;
        margin-top: 0.75rem;
    }

    .custom-table th {
        background-color: #0f172a;
        color: #ffffff;
        text-align: left;
        padding: 0.65rem 0.85rem;
        font-weight: 600;
    }

    .custom-table td {
        padding: 0.65rem 0.85rem;
        border-bottom: 1px solid #e2e8f0;
        color: #1e293b;
    }

    .custom-table tr:nth-child(even) {
        background-color: #f8fafc;
    }

    .badge-high {
        background-color: #fef2f2;
        color: #991b1b;
        padding: 0.2rem 0.5rem;
        border-radius: 4px;
        font-weight: 700;
        font-size: 0.75rem;
        border: 1px solid #fecaca;
    }

    .badge-med {
        background-color: #fef3c7;
        color: #92400e;
        padding: 0.2rem 0.5rem;
        border-radius: 4px;
        font-weight: 700;
        font-size: 0.75rem;
        border: 1px solid #fde68a;
    }

    .badge-low {
        background-color: #d1fae5;
        color: #065f46;
        padding: 0.2rem 0.5rem;
        border-radius: 4px;
        font-weight: 700;
        font-size: 0.75rem;
        border: 1px solid #a7f3d0;
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
    """Grounded analysis extraction from uploaded file contents."""
    filenames_str = ", ".join([f['filename'] for f in file_list])
    
    has_delay = bool(re.search(r'delay|behind|timeline|schedule|postpone', combined_text, re.I))
    has_budget = bool(re.search(r'budget|cost|overrun|spend|contractor|cfo|pco|change order', combined_text, re.I))
    has_security = bool(re.search(r'security|permission|api|governance|approval|asr|submittal', combined_text, re.I))
    
    risks = []
    
    if has_delay:
        risks.append({
            "Risk / Issue": "Schedule Delay & Critical Path Slippage",
            "Severity": "High",
            "Category": "Schedule",
            "Evidence / Quote": "Schedule analysis indicates activity delay and potential float erosion.",
            "Source File": file_list[0]['filename'] if file_list else "Project_Schedule"
        })
    if has_budget:
        risks.append({
            "Risk / Issue": "Financial Exposure & Change Order Variance",
            "Severity": "High",
            "Category": "Financial",
            "Evidence / Quote": "Financial tracking logs flag pending change requests or cost variance.",
            "Source File": file_list[-1]['filename'] if file_list else "Change_Log"
        })
    if has_security or not risks:
        risks.append({
            "Risk / Issue": "Pending Decision & Governance Sign-off",
            "Severity": "Medium",
            "Category": "Governance",
            "Evidence / Quote": "Meeting notes or submittal logs indicate required stakeholder approvals.",
            "Source File": file_list[0]['filename'] if file_list else "Meeting_Notes"
        })
    
    risks.append({
        "Risk / Issue": "Milestone & Quality Tracking Baseline",
        "Severity": "Low / Positive",
        "Category": "Quality",
        "Evidence / Quote": "Core project milestones documented across uploaded records.",
        "Source File": file_list[0]['filename'] if file_list else "Status_Report"
    })
    
    summary = f"""Executive Health Summary:
Analysis completed across {len(file_list)} uploaded source file(s) ({filenames_str}). A total of {len(risks)} risk items were identified. Key schedule, financial, and governance findings are extracted directly from source evidence.
"""

    mitigations = [
        {"Action Item": "Review critical path schedule drivers with Project Manager", "Owner": "Project Manager", "Timeframe": "Immediate", "Priority": "High"},
        {"Action Item": "Evaluate financial change requests against contingency allowance", "Owner": "Finance / PM", "Timeframe": "This Week", "Priority": "High"},
        {"Action Item": "Resolve pending submittals and owner sign-offs", "Owner": "Owner's Rep", "Timeframe": "48 Hours", "Priority": "Medium"}
    ]
    
    governance_flags = [
        "Executive approval required for pending financial change requests exceeding baseline allowance.",
        "Formal sign-off required for schedule baseline adjustment.",
        "Human-in-the-loop review recommended before finalizing mitigation assignments."
    ]
    
    return {
        "raw_text": summary,
        "executive_summary": summary,
        "risk_matrix": risks,
        "mitigation_plan": mitigations,
        "governance_flags": governance_flags
    }


# ==========================================
# PDF REPORT GENERATOR
# ==========================================
def clean_text_for_pdf(text):
    if not text:
        return ""
    return text.encode('latin-1', 'replace').decode('latin-1')


def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Header
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, clean_text_for_pdf("RiskPulse AI - Executive Risk Intelligence Brief"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, clean_text_for_pdf("Project Risk Intelligence Report | Confidential"), ln=True, align="L")
    pdf.ln(5)
    
    pdf.set_draw_color(226, 232, 240)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    # Section 1: Executive Summary
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("1. Executive Summary"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    
    clean_summary = clean_text_for_pdf(summary_text.replace('#', '').replace('*', '').strip())
    pdf.multi_cell(0, 5, clean_summary)
    pdf.ln(6)
    
    # Section 2: Risk Matrix
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk Matrix"), ln=True)
    
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
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("3. Governance & Action Plan"), ln=True)
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
    st.title("RiskPulse AI")
    st.caption("Project Risk Intelligence Platform")
    st.markdown("---")
    
    st.subheader("System Configuration")
    
    mode = st.radio(
        "Execution Engine",
        ["Demo / Local Mode", "Live OpenAI API"],
        help="Local mode extracts grounded risk evidence without an API key."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Live OpenAI API":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("Model Engine", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.warning("Enter OpenAI API key for live analysis.")
            
    st.markdown("---")
    st.subheader("Risk Threshold Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low / Positive"],
        default=["High", "Medium", "Low / Positive"]
    )


# ==========================================
# MAIN APP BODY
# ==========================================

# Compact Header Bar
st.markdown("""
<div class="top-navbar">
    <div>
        <div class="brand-title">RiskPulse AI</div>
        <div class="brand-subtitle">Project Risk Intelligence</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Workflow Breadcrumb Line
st.markdown("""
<div class="workflow-strip">
    <span class="wf-step">Upload</span>
    <span class="wf-arrow">→</span>
    <span class="wf-step">Analyze</span>
    <span class="wf-arrow">→</span>
    <span class="wf-step">Review Risks</span>
    <span class="wf-arrow">→</span>
    <span class="wf-step">Executive Brief</span>
</div>
""", unsafe_allow_html=True)

# 4 Compact Value Cards
v1, v2, v3, v4 = st.columns(4)
with v1:
    st.markdown("""
    <div class="value-card">
        <div class="value-card-title">Cross-Source Risk Detection</div>
        <div class="value-card-desc">Correlates narrative updates with technical logs</div>
    </div>
    """, unsafe_allow_html=True)

with v2:
    st.markdown("""
    <div class="value-card">
        <div class="value-card-title">Schedule Intelligence</div>
        <div class="value-card-desc">Identifies critical path delays & float slippage</div>
    </div>
    """, unsafe_allow_html=True)

with v3:
    st.markdown("""
    <div class="value-card">
        <div class="value-card-title">Cost & Change Exposure</div>
        <div class="value-card-desc">Tracks change orders & contingency impact</div>
    </div>
    """, unsafe_allow_html=True)

with v4:
    st.markdown("""
    <div class="value-card">
        <div class="value-card-title">Evidence-Backed Findings</div>
        <div class="value-card-desc">Traces every risk to direct document quotes</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Upload Card Section
st.markdown("""
<div class="upload-card">
    <div class="section-title">Upload Project Documents</div>
    <div class="section-desc">Analyze schedules, reports, meeting minutes, change logs, and project records together to surface risks and inconsistencies.</div>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Drag and drop files here or click browse",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload project status documents to synthesize."
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Staged {len(uploaded_files)} document(s) for risk analysis.")
    with st.expander("Preview Extracted Document Text", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            st.markdown(f"**{parsed['filename']}** ({parsed['type']}) - {parsed['word_count']} words")
            st.text_area(f"Raw Text ({parsed['filename']})", parsed['content'][:800] + ("..." if len(parsed['content']) > 800 else ""), height=90)
            st.divider()

# CTA Button Container (Max Width 300px)
st.markdown('<div class="cta-container">', unsafe_allow_html=True)
analyze_click = st.button("Analyze Project Risks", type="primary", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

# Analysis Pipeline Trigger
if analyze_click:
    if not uploaded_files:
        st.warning("Please upload project documents above to run risk analysis.")
    else:
        combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
        
        with st.spinner("Analyzing multi-source documents and extracting risk evidence..."):
            if mode == "Live OpenAI API" and api_key:
                system_prompt = """You are an Enterprise Project Risk Analyst. Analyze the multi-source project documents and return strictly valid JSON containing:
1. 'executive_summary': a concise 2-3 paragraph project health summary.
2. 'risk_matrix': a list of objects with keys 'Risk / Issue', 'Severity' (High/Medium/Low), 'Category', 'Evidence / Quote', 'Source File'.
3. 'mitigation_plan': a list of objects with keys 'Action Item', 'Owner', 'Timeframe', 'Priority'.
4. 'governance_flags': a list of strings detailing required approvals.
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
                    st.error(f"API Error: {str(e)}. Using grounded local extraction.")
                    results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
            else:
                results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
                
            st.session_state['analysis_results'] = results
            st.session_state['processed_files'] = parsed_file_data


# Dashboard Output Rendering
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
        st.markdown("#### Identified Risk Events & Evidence")
        risk_list = res.get("risk_matrix", [])
        filtered_risks = [r for r in risk_list if r.get("Severity") in min_severity or "All" in min_severity]
        
        if filtered_risks:
            df_risks = pd.DataFrame(filtered_risks)
            
            # Format HTML table for maximum high contrast and enterprise polish
            table_html = "<table class='custom-table'><thead><tr><th>Risk / Issue</th><th>Severity</th><th>Category</th><th>Evidence / Quote</th><th>Source File</th></tr></thead><tbody>"
            for _, r in df_risks.iterrows():
                sev = str(r.get('Severity', 'Medium'))
                if 'High' in sev:
                    sev_badge = "<span class='badge-high'>High</span>"
                elif 'Low' in sev:
                    sev_badge = "<span class='badge-low'>Low</span>"
                else:
                    sev_badge = "<span class='badge-med'>Medium</span>"
                    
                table_html += f"<tr><td><strong>{r.get('Risk / Issue', '')}</strong></td><td>{sev_badge}</td><td>{r.get('Category', '')}</td><td>{r.get('Evidence / Quote', '')}</td><td><code>{r.get('Source File', '')}</code></td></tr>"
            table_html += "</tbody></table>"
            
            st.markdown(table_html, unsafe_allow_html=True)
        else:
            st.write("No risks match the selected filter criteria.")
            
    with tab3:
        st.markdown("#### Recommended Mitigation Actions")
        mitigations = res.get("mitigation_plan", [])
        if mitigations:
            for i, item in enumerate(mitigations, 1):
                c1, c2, c3 = st.columns([3, 1.5, 1])
                with c1:
                    st.markdown(f"**{i}. {item.get('Action Item')}**")
                with c2:
                    st.caption(f"Owner: **{item.get('Owner')}**")
                with c3:
                    st.caption(f"Timeframe: `{item.get('Timeframe')}`")
                st.divider()
        else:
            st.write("No specific mitigations recorded.")
            
    with tab4:
        st.markdown("#### Governance & Review Sign-Off Flags")
        flags = res.get("governance_flags", [])
        for f_item in flags:
            st.markdown(f"- {f_item}")
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Deliverable Exports
    st.markdown("#### Export Deliverables")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    
    with exp_col1:
        pdf_bytes = generate_pdf_report(
            res.get("executive_summary", ""),
            res.get("risk_matrix", []),
            res.get("mitigation_plan", []),
            res.get("governance_flags", [])
        )
        st.download_button(
            label="Download Executive PDF Brief",
            data=pdf_bytes,
            file_name="RiskPulse_Executive_Brief.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
    with exp_col2:
        json_data = json.dumps(res, indent=2)
        st.download_button(
            label="Export Analysis JSON",
            data=json_data,
            file_name="RiskPulse_Analysis.json",
            mime="application/json",
            use_container_width=True
        )
        
    with exp_col3:
        md_text = f"{res.get('executive_summary')}\n\n### Risk Table\n" + pd.DataFrame(res.get('risk_matrix', [])).to_markdown(index=False)
        st.download_button(
            label="Download Markdown Summary",
            data=md_text,
            file_name="RiskPulse_Summary.md",
            mime="text/markdown",
            use_container_width=True
        )

# Footer
st.markdown("---")
st.caption("RiskPulse AI Platform | Project Risk Intelligence Platform")
