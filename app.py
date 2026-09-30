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
# PAGE CONFIGURATION & SAAS STYLING
# ==========================================
st.set_page_config(
    page_title="RiskPulse AI | Project Risk Intelligence",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Commercial Enterprise Aesthetics
st.markdown("""
<style>
    /* Global Canvas */
    .main {
        background-color: #f8fafc;
    }
    .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0f172a;
    }
    
    /* Hide empty Streamlit header wrapper */
    header[data-testid="stHeader"] {
        display: none !important;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 1280px;
    }

    /* Top SaaS Header Bar */
    .top-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background-color: #0f172a;
        color: #ffffff;
        padding: 0.85rem 1.5rem;
        border-radius: 8px;
        margin-bottom: 1.25rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }
    .brand-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #ffffff;
        letter-spacing: -0.01em;
    }
    .brand-subtitle {
        font-size: 0.85rem;
        color: #94a3b8;
        margin-left: 0.75rem;
        font-weight: 400;
    }

    /* Hero Headline & Subtitle */
    .page-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.25rem;
        letter-spacing: -0.01em;
    }
    .page-description {
        font-size: 0.95rem;
        color: #475569;
        margin-bottom: 1.25rem;
    }

    /* Process Indicator Bar */
    .process-indicator {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        background-color: #ffffff;
        padding: 0.6rem 1rem;
        border-radius: 6px;
        border: 1px solid #e2e8f0;
        margin-bottom: 1.25rem;
        font-size: 0.825rem;
        color: #64748b;
    }
    .process-step-active {
        font-weight: 600;
        color: #2563eb;
    }
    .process-step {
        font-weight: 500;
        color: #475569;
    }
    .process-divider {
        color: #cbd5e1;
    }

    /* Capability Cards */
    .capability-card {
        background: #ffffff;
        padding: 1.1rem 1.25rem;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
        height: 100%;
    }
    .capability-title {
        font-size: 0.9rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.4rem;
    }
    .capability-desc {
        font-size: 0.825rem;
        color: #475569;
        line-height: 1.4;
    }

    /* Button Styling Override - Permanent Blue #2563EB */
    div.stButton > button[kind="primary"],
    div.stButton > button {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border-color: #2563eb !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
        border-radius: 6px !important;
        padding: 0.55rem 1.5rem !important;
        box-shadow: 0 1px 2px rgba(37, 99, 235, 0.2) !important;
    }
    div.stButton > button:hover {
        background-color: #1d4ed8 !important;
        border-color: #1d4ed8 !important;
    }

    /* File Uploader Container */
    .stFileUploader {
        background-color: #ffffff;
        border: 1px dashed #cbd5e1;
        border-radius: 8px;
        padding: 0.75rem;
    }

    /* Risk Table Styling */
    .risk-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
        background: #ffffff;
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #e2e8f0;
    }
    .risk-table th {
        background-color: #0f172a;
        color: #ffffff;
        font-weight: 600;
        text-align: left;
        padding: 0.65rem 0.85rem;
    }
    .risk-table td {
        padding: 0.65rem 0.85rem;
        border-bottom: 1px solid #f1f5f9;
        color: #1e293b;
    }
    .badge-high {
        background-color: #fee2e2;
        color: #991b1b;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.15rem 0.5rem;
        border-radius: 4px;
    }
    .badge-med {
        background-color: #fef3c7;
        color: #92400e;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.15rem 0.5rem;
        border-radius: 4px;
    }
    .badge-low {
        background-color: #d1fae5;
        color: #065f46;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.15rem 0.5rem;
        border-radius: 4px;
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
# OPENAI & GROUNDED LOCAL AI RISK PIPELINE
# ==========================================
def call_openai_api(api_key, system_prompt, user_prompt, model="gpt-4o-mini"):
    """Direct urllib REST call to OpenAI API."""
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
    """Analyzes text from uploaded source files to extract evidence-backed risks."""
    risks = []
    
    for f in file_list:
        fname = f['filename']
        content = f['content']
        
        # Scan for delay/schedule keywords
        delay_matches = re.findall(r'([^.!\n]*?(?:delay|behind|schedule|critical path|milestone|float)[^.!\n]*?[.!\n])', content, re.I)
        if delay_matches:
            risks.append({
                "Risk / Issue": f"Schedule / Timeline Exposure in {fname}",
                "Severity": "High",
                "Category": "Schedule",
                "Evidence / Quote": delay_matches[0].strip()[:140],
                "Source File": fname
            })
            
        # Scan for cost/financial keywords
        cost_matches = re.findall(r'([^.!\n]*?(?:cost|budget|overrun|change order|pco|fee|contingency|\$)[^.!\n]*?[.!\n])', content, re.I)
        if cost_matches:
            risks.append({
                "Risk / Issue": f"Cost / Financial Exposure in {fname}",
                "Severity": "High",
                "Category": "Financial",
                "Evidence / Quote": cost_matches[0].strip()[:140],
                "Source File": fname
            })
            
        # Scan for governance/decision keywords
        gov_matches = re.findall(r'([^.!\n]*?(?:approval|decision|pending|sign-off|rfi|submittal|owner)[^.!\n]*?[.!\n])', content, re.I)
        if gov_matches and not delay_matches and not cost_matches:
            risks.append({
                "Risk / Issue": f"Governance / Approval Item in {fname}",
                "Severity": "Medium",
                "Category": "Governance",
                "Evidence / Quote": gov_matches[0].strip()[:140],
                "Source File": fname
            })

    if not risks:
        risks.append({
            "Risk / Issue": "Document Review Completed",
            "Severity": "Low / Positive",
            "Category": "Operational",
            "Evidence / Quote": f"Successfully parsed {len(file_list)} file(s). No immediate high-risk keyword triggers detected.",
            "Source File": file_list[0]['filename'] if file_list else "Uploaded Files"
        })

    filenames_str = ", ".join([f['filename'] for f in file_list])
    high_count = sum(1 for r in risks if "High" in r["Severity"])
    med_count = sum(1 for r in risks if "Medium" in r["Severity"])

    summary = f"""Executive Health Summary: Analysis completed across {len(file_list)} uploaded source file(s) ({filenames_str}). A total of {len(risks)} actionable items were identified, including {high_count} High-severity exposure(s) and {med_count} Medium-severity watch item(s). Key schedule, financial, and governance findings are extracted below directly from source evidence."""

    mitigations = [
        {"Action Item": f"Review and validate identified risk triggers in {risks[0]['Source File']}", "Owner": "Project Manager", "Timeframe": "Within 48 Hours", "Priority": "High"},
        {"Action Item": "Establish formal resolution path for open approvals and cost exposures", "Owner": "Project Controls Lead", "Timeframe": "Next Weekly Sync", "Priority": "Medium"}
    ]

    governance_flags = [
        f"Verification Sign-off: {len(risks)} risk items extracted from {len(file_list)} uploaded file(s) require review by the designated Project Manager prior to executive dispatch."
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
    """Sanitizes text for FPDF Latin-1 encoding."""
    if not text:
        return ""
    return text.encode('latin-1', 'replace').decode('latin-1')


def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    """Generates an executive PDF report."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Header
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, clean_text_for_pdf("RiskPulse AI — Executive Risk Brief"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, clean_text_for_pdf("Generated from Uploaded Source Documents | Confidential"), ln=True, align="L")
    pdf.ln(4)
    
    pdf.set_draw_color(226, 232, 240)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(6)
    
    # Section 1: Executive Summary
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("1. Executive Summary"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(0, 5, clean_text_for_pdf(summary_text))
    pdf.ln(6)
    
    # Section 2: Risk Matrix Table
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk & Issue Matrix"), ln=True)
    
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
        
    pdf.ln(6)
    
    # Section 3: Governance Flags
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("3. Human Review & Verification Flags"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(51, 65, 85)
    
    for flag in governance_flags:
        pdf.multi_cell(0, 5, f"- {clean_text_for_pdf(flag)}")
        pdf.ln(1)
        
    return bytes(pdf.output())


# ==========================================
# SIDEBAR CONTROLS
# ==========================================
with st.sidebar:
    st.markdown("### System Settings")
    
    mode = st.radio(
        "Execution Engine",
        ["Local Text Analysis Engine", "Live OpenAI API"],
        help="Local engine extracts grounded text evidence from uploaded files. Live API connects to OpenAI."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Live OpenAI API":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("LLM Model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.caption("Please enter an API key to run live model queries.")
            
    st.markdown("---")
    st.markdown("### Risk Thresholds")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low / Positive"],
        default=["High", "Medium", "Low / Positive"]
    )


# ==========================================
# MAIN APPLICATION LAYOUT
# ==========================================

# 1. Compact Top SaaS Header
st.markdown("""
<div class="top-header">
    <div>
        <span class="brand-title">RiskPulse AI</span>
        <span class="brand-subtitle">Project Risk Intelligence</span>
    </div>
</div>
""", unsafe_allow_html=True)

# 2. Main Page Headline & Subtitle
st.markdown('<div class="page-title">Project Risk Intelligence</div>', unsafe_allow_html=True)
st.markdown('<div class="page-description">Turn project records into evidence-backed risk intelligence.</div>', unsafe_allow_html=True)

# 3. Process Indicator Bar
st.markdown("""
<div class="process-indicator">
    <span class="process-step-active">1. Upload Documents</span>
    <span class="process-divider">→</span>
    <span class="process-step">2. Analyze & Extract</span>
    <span class="process-divider">→</span>
    <span class="process-step">3. Review Risk Register</span>
    <span class="process-divider">→</span>
    <span class="process-step">4. Executive Brief</span>
</div>
""", unsafe_allow_html=True)

# 4. Capability Cards (4 Columns)
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Cross-Source Risk Detection</div>
        <div class="capability-desc">Correlates narrative status updates, technical logs, and meeting minutes to surface discrepancies.</div>
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
        <div class="capability-desc">Tracks budget overruns, pending change orders, and contingency usage from project logs.</div>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Evidence-Backed Findings</div>
        <div class="capability-desc">Traces every extracted risk to direct quotes and verified source document references.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 5. File Upload Section
st.markdown("### Upload Project Documents")
st.caption("Analyze schedules, reports, meeting minutes, change logs, and project records together to surface risks and inconsistencies.")

uploaded_files = st.file_uploader(
    "Drag and drop files here or click to browse",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload project status documents to analyze.",
    label_visibility="collapsed"
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Staged {len(uploaded_files)} file(s) for risk analysis.")
    with st.expander("View Staged Document Metadata", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            st.markdown(f"**{parsed['filename']}** (`{parsed['type']}`) — {parsed['word_count']} words | {parsed['char_count']} characters")

st.markdown("<br>", unsafe_allow_html=True)

# 6. Primary Action Button
btn_col, _ = st.columns([1, 2])
with btn_col:
    analyze_click = st.button("Analyze Project Risks", type="primary")

# 7. Analysis Execution
if analyze_click:
    if not uploaded_files:
        st.warning("Please upload at least one project document above to analyze risks.")
    else:
        combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
        
        with st.status("Analyzing Project Documents...", expanded=True) as status:
            st.write("Parsing document text and extracting metadata...")
            st.write("Scanning for schedule, cost, and governance risk triggers...")
            
            if mode == "Live OpenAI API" and api_key:
                system_prompt = """You are an Enterprise Risk Analyst. Analyze the provided project documents and return a JSON object with:
1. 'executive_summary': a 2-3 paragraph summary of project health and key risk exposure.
2. 'risk_matrix': a list of objects with keys 'Risk / Issue', 'Severity' (High/Medium/Low), 'Category', 'Evidence / Quote', 'Source File'.
3. 'mitigation_plan': a list of objects with keys 'Action Item', 'Owner', 'Timeframe', 'Priority'.
4. 'governance_flags': a list of strings detailing required human sign-offs.

Return valid JSON."""
                try:
                    raw_response = call_openai_api(api_key, system_prompt, f"Project Documents:\n{combined_text}", model=model_choice)
                    json_start = raw_response.find('{')
                    json_end = raw_response.rfind('}') + 1
                    if json_start != -1 and json_end != -1:
                        results = json.loads(raw_response[json_start:json_end])
                    else:
                        results = run_grounded_text_analysis(combined_text, parsed_file_data)
                except Exception as e:
                    st.error(f"OpenAI API call encountered an error: {str(e)}. Executing grounded text engine.")
                    results = run_grounded_text_analysis(combined_text, parsed_file_data)
            else:
                results = run_grounded_text_analysis(combined_text, parsed_file_data)
                
            status.update(label="Analysis Complete", state="complete", expanded=False)
            
        st.session_state['analysis_results'] = results
        st.session_state['processed_files'] = parsed_file_data


# 8. Render Results Dashboard
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    
    st.markdown("---")
    st.markdown("## Project Controls & Risk Register")
    
    # Functional View Tabs
    tab_overview, tab_docs, tab_risks, tab_brief = st.tabs([
        "Overview", 
        "Source Documents", 
        "Risk Register", 
        "Executive Brief"
    ])
    
    with tab_overview:
        st.markdown("### Executive Summary")
        st.write(res.get("executive_summary", "No summary generated."))
        
        st.markdown("### Recommended Action Plan")
        mitigations = res.get("mitigation_plan", [])
        if mitigations:
            for item in mitigations:
                st.markdown(f"• **{item.get('Action Item')}** — Assigned to **{item.get('Owner')}** (`{item.get('Timeframe')}`)")
        else:
            st.write("No specific mitigations recorded.")
            
    with tab_docs:
        st.markdown("### Source Documents Analyzed")
        files_proc = st.session_state.get('processed_files', [])
        if files_proc:
            df_files = pd.DataFrame([{
                "Filename": f["filename"],
                "Format": f["type"],
                "Word Count": f["word_count"],
                "Character Count": f["char_count"]
            } for f in files_proc])
            st.dataframe(df_files, use_container_width=True, hide_index=True)
        else:
            st.write("No source document details available.")
            
    with tab_risks:
        st.markdown("### Identified Risks & Issues")
        risk_list = res.get("risk_matrix", [])
        filtered_risks = [r for r in risk_list if any(sev in r.get("Severity", "") for sev in min_severity) or not min_severity]
        
        if filtered_risks:
            df_risks = pd.DataFrame(filtered_risks)
            st.dataframe(
                df_risks,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Evidence / Quote": st.column_config.TextColumn("Direct Evidence / Quote", width="large")
                }
            )
        else:
            st.write("No risks match the selected threshold filter.")
            
    with tab_brief:
        st.markdown("### Governance Flags & PDF Export")
        flags = res.get("governance_flags", [])
        for f_item in flags:
            st.markdown(f"- {f_item}")
            
        st.markdown("<br>", unsafe_allow_html=True)
        pdf_bytes = generate_pdf_report(
            res.get("executive_summary", ""),
            res.get("risk_matrix", []),
            res.get("mitigation_plan", []),
            res.get("governance_flags", [])
        )
        st.download_button(
            label="Download Executive Brief (PDF)",
            data=pdf_bytes,
            file_name="RiskPulse_Executive_Brief.pdf",
            mime="application/pdf"
        )
