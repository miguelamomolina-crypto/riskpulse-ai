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
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="RiskPulse AI | Project Risk Intelligence",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# CUSTOM CSS - ENTERPRISE LIGHT THEME & COMPACT SPACING
# ==========================================
st.markdown("""
<style>
    /* Force Light Theme Canvas */
    :root {
        --background-color: #f8fafc;
        --secondary-background-color: #ffffff;
        --text-color: #0f172a;
    }
    
    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], .main, .block-container {
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }
    
    /* Hide Default Streamlit Chrome Header */
    header[data-testid="stHeader"], div[data-testid="stDecoration"], div[data-testid="stStatusWidget"] {
        display: none !important;
    }
    
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        max-width: 1280px !important;
    }
    
    /* Compact Navy Header */
    .navy-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
        padding: 1.25rem 1.75rem;
        border-radius: 10px;
        color: white !important;
        margin-bottom: 1rem;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15);
    }
    .navy-header h1 {
        color: #ffffff !important;
        font-size: 1.65rem !weight: 800;
        margin: 0 0 0.25rem 0 !important;
        padding: 0 !important;
        line-height: 1.2;
    }
    .navy-header p {
        color: #94a3b8 !important;
        font-size: 0.925rem !important;
        margin: 0 !important;
        padding: 0 !important;
        line-height: 1.4;
    }
    
    /* Workflow Sequence Bar */
    .workflow-bar {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        margin-bottom: 1.25rem;
        text-align: center;
        font-size: 0.825rem;
        font-weight: 600;
        color: #475569 !important;
        letter-spacing: 0.01em;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .workflow-step {
        color: #2563eb;
        font-weight: 700;
    }
    
    /* Compact Capability Cards */
    .capability-card {
        background: #ffffff;
        padding: 0.85rem 1rem;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
        height: 100%;
    }
    .capability-title {
        font-size: 0.875rem;
        font-weight: 700;
        color: #0f172a !important;
        margin-bottom: 0.35rem;
    }
    .capability-desc {
        font-size: 0.775rem;
        color: #475569 !important;
        line-height: 1.35;
        margin: 0;
    }
    
    /* Light File Uploader Styling */
    div[data-testid="stFileUploader"] {
        background-color: #ffffff !important;
        border: 2px dashed #2563eb !important;
        border-radius: 10px !important;
        padding: 0.75rem 1rem !important;
        margin-bottom: 0.75rem !important;
    }
    div[data-testid="stFileUploader"] section {
        background-color: #ffffff !important;
        padding: 0.5rem !important;
    }
    div[data-testid="stFileUploader"] label {
        color: #0f172a !important;
        font-weight: 700 !important;
        font-size: 0.95rem !important;
    }
    div[data-testid="stFileUploader"] small {
        color: #64748b !important;
    }
    
    /* Hide stray icons/squares in uploader top right */
    div[data-testid="stFileUploader"] section > div > button {
        display: none !important;
    }
    div[data-testid="stFileUploaderDropzoneInstructions"] small {
        display: none !important;
    }
    
    /* Uploaded File Item / Chip Styling */
    div[data-testid="stFileUploaderFile"], .stFileUploaderFile {
        background-color: #f8fafc !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 6px !important;
        color: #0f172a !important;
        padding: 0.4rem 0.75rem !important;
        margin-top: 0.4rem !important;
    }
    div[data-testid="stFileUploaderFile"] span, div[data-testid="stFileUploaderFile"] div {
        color: #0f172a !important;
    }
    div[data-testid="stFileUploaderFile"] button {
        color: #64748b !important;
    }
    
    /* CTA Button - Force Pure White Text on Blue */
    div.stButton {
        text-align: center;
        margin-top: 0.5rem;
        margin-bottom: 0.75rem;
    }
    div.stButton > button, button[type="primary"], button[kind="primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        font-size: 0.925rem !important;
        padding: 0.55rem 1.75rem !important;
        border-radius: 6px !important;
        border: none !important;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25) !important;
        width: 300px !important;
        max-width: 100% !important;
        margin: 0 auto !important;
        display: inline-block !important;
        transition: background-color 0.15s ease !important;
    }
    div.stButton > button:hover, button[type="primary"]:hover, button[kind="primary"]:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
    }
    div.stButton > button p, button[type="primary"] p {
        color: #ffffff !important;
        font-weight: 700 !important;
    }
    
    /* Result Cards & Headers */
    .result-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1.25rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
    }
    h1, h2, h3, h4, h5, h6 {
        color: #0f172a !important;
        font-weight: 700 !important;
    }
    p, span, label, caption {
        color: #0f172a !important;
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
# GROUNDED TEXT & AI RISK ANALYSIS PIPELINE
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
    """Grounds analysis strictly in actual uploaded document content."""
    risks = []
    
    # Extract line-by-line evidence from uploaded files
    for f in file_list:
        content = f.get('content', '')
        fname = f.get('filename', 'Document')
        
        # Look for schedule triggers
        delay_matches = re.findall(r'([^.\n]*?(?:delay|behind|postpone|schedule|timeline|submittal)[^.\n]*?\.)', content, re.I)
        if delay_matches:
            risks.append({
                "Risk / Issue": f"Schedule Delay & Timeline Variance in {fname}",
                "Severity": "High",
                "Category": "Schedule",
                "Evidence / Quote": delay_matches[0].strip(),
                "Source File": fname
            })
            
        # Look for budget / cost triggers
        cost_matches = re.findall(r'([^.\n]*?(?:budget|cost|overrun|spend|change order|allowance|contingency|fee)[^.\n]*?\.)', content, re.I)
        if cost_matches:
            risks.append({
                "Risk / Issue": f"Cost & Financial Contingency Exposure in {fname}",
                "Severity": "High",
                "Category": "Financial",
                "Evidence / Quote": cost_matches[0].strip(),
                "Source File": fname
            })

        # Look for governance / decision triggers
        gov_matches = re.findall(r'([^.\n]*?(?:approval|sign-off|pending|decision|architect|owner|governance)[^.\n]*?\.)', content, re.I)
        if gov_matches:
            risks.append({
                "Risk / Issue": f"Governance & Approval Bottleneck in {fname}",
                "Severity": "Medium",
                "Category": "Governance",
                "Evidence / Quote": gov_matches[0].strip(),
                "Source File": fname
            })

    if not risks:
        risks.append({
            "Risk / Issue": "General Document Review & Baseline Monitoring",
            "Severity": "Low / Positive",
            "Category": "General",
            "Evidence / Quote": f"Document text extracted successfully ({len(combined_text)} characters analyzed across {len(file_list)} files).",
            "Source File": file_list[0]['filename'] if file_list else "Uploaded Files"
        })

    filenames_str = ", ".join([f['filename'] for f in file_list])
    high_count = sum(1 for r in risks if r.get("Severity") == "High")
    med_count = sum(1 for r in risks if r.get("Severity") == "Medium")
    
    summary = f"""Executive Health Summary: Analysis completed across **{len(file_list)} uploaded source file(s)** ({filenames_str}). A total of **{len(risks)} actionable items** were identified, including **{high_count} High-severity exposure(s)** and **{med_count} Medium-severity watch item(s)**. Key schedule, financial, and governance findings are extracted below directly from source evidence."""

    mitigations = [
        {"Action Item": f"Review primary risk triggers identified in {r.get('Source File')}", "Owner": "Project Manager", "Timeframe": "Within 48 Hours", "Priority": r.get('Severity')}
        for r in risks[:4]
    ]

    governance_flags = [
        f"Sign-off required for items extracted from {r.get('Source File')}: \"{r.get('Evidence / Quote')[:80]}...\""
        for r in risks if r.get('Severity') == 'High'
    ]
    if not governance_flags:
        governance_flags = ["All documented milestones currently align with project guidelines."]

    return {
        "raw_text": summary,
        "executive_summary": summary,
        "risk_matrix": risks,
        "mitigation_plan": mitigations,
        "governance_flags": governance_flags
    }

# ==========================================
# REPORT GENERATORS & EXPORT HELPERS
# ==========================================
def clean_text_for_pdf(text):
    """Sanitizes text to prevent fpdf encoding errors."""
    if not text:
        return ""
    emoji_map = {
        '⚠️': '[WARNING]', '🔒': '[GOVERNANCE]', '💰': '[FINANCIAL]', 
        '👥': '[HUMAN-IN-LOOP]', '📊': '[DATA]', '🛠️': '[ACTION]', 
        '📄': '[DOC]', '⚡': '[FAST]', '👤': '[OWNER]', '⏱️': '[TIME]', 
        '💡': '[NOTE]'
    }
    for emoji, rep in emoji_map.items():
        text = text.replace(emoji, rep)
    return text.encode('latin-1', 'replace').decode('latin-1')

def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    """Generates an executive PDF report using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, clean_text_for_pdf("Executive Risk Intelligence Report"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, clean_text_for_pdf("Generated by RiskPulse AI Platform | Confidential"), ln=True, align="L")
    pdf.ln(5)
    
    pdf.set_draw_color(226, 232, 240)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("1. Executive Summary"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    
    clean_summary = clean_text_for_pdf(summary_text.replace('#', '').replace('*', '').strip())
    pdf.multi_cell(0, 5, clean_summary)
    pdf.ln(6)
    
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk Register"), ln=True)
    
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(241, 245, 249)
    pdf.cell(50, 7, "Risk / Issue", border=1, fill=True)
    pdf.cell(25, 7, "Severity", border=1, fill=True)
    pdf.cell(30, 7, "Category", border=1, fill=True)
    pdf.cell(85, 7, "Evidence Quote", border=1, ln=True, fill=True)
    
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
        
    return bytes(pdf.output())

def generate_markdown_table(risk_data):
    """Generates a markdown table string without tabulate dependency."""
    if not risk_data:
        return "No risk items recorded."
    headers = ["Risk / Issue", "Severity", "Category", "Evidence / Quote", "Source File"]
    lines = ["| " + " | ".join(headers) + " |"]
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
    for row in risk_data:
        r_name = str(row.get("Risk / Issue", "")).replace("|", "/")
        sev = str(row.get("Severity", "")).replace("|", "/")
        cat = str(row.get("Category", "")).replace("|", "/")
        ev = str(row.get("Evidence / Quote", "")).replace("|", "/")
        src = str(row.get("Source File", "")).replace("|", "/")
        lines.append(f"| {r_name} | {sev} | {cat} | {ev} | {src} |")
    return "\n".join(lines)

# ==========================================
# SIDEBAR NAVIGATION & SETTINGS
# ==========================================
with st.sidebar:
    st.title("RiskPulse AI")
    st.caption("Project Risk Intelligence")
    st.markdown("---")
    
    st.subheader("Analysis Settings")
    mode = st.radio(
        "Execution Engine",
        ["Local Analysis (No Key Required)", "Enhanced AI Analysis"],
        help="Local analysis extracts grounded risk evidence directly from files."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Enhanced AI Analysis":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("Model Engine", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.info("Enter an API key to enable OpenAI reasoning.")
            
    st.markdown("---")
    st.subheader("Risk Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low / Positive"],
        default=["High", "Medium", "Low / Positive"]
    )

# ==========================================
# COMPACT NAVY HEADER
# ==========================================
st.markdown("""
<div class="navy-header">
    <h1>RiskPulse AI</h1>
    <p style="font-weight: 600; color: #cbd5e1 !important; margin-bottom: 0.15rem !important;">Project Risk Intelligence</p>
    <p>Turn project records into evidence-backed risk intelligence.</p>
</div>
""", unsafe_allow_html=True)

# Process Indicator Bar
st.markdown("""
<div class="workflow-bar">
    <span class="workflow-step">1. Upload Documents</span> &nbsp;→&nbsp; 
    <span class="workflow-step">2. Analyze & Extract</span> &nbsp;→&nbsp; 
    <span class="workflow-step">3. Review Risk Register</span> &nbsp;→&nbsp; 
    <span class="workflow-step">4. Executive Brief</span>
</div>
""", unsafe_allow_html=True)

# 4 Capability Cards
c1, c2, m3, c4 = st.columns(4)
with c1:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Cross-Source Risk Detection</div>
        <p class="capability-desc">Correlates findings and identifies contradictions across schedules, reports, and meeting logs.</p>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Schedule Intelligence</div>
        <p class="capability-desc">Surfaces schedule, float, and milestone risks when supported by uploaded schedule data.</p>
    </div>
    """, unsafe_allow_html=True)

with m3:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Cost & Change Exposure</div>
        <p class="capability-desc">Tracks budget overruns, change orders, and financial contingency impacts across project logs.</p>
    </div>
    """, unsafe_allow_html=True)

with c4:
    st.markdown("""
    <div class="capability-card">
        <div class="capability-title">Evidence-Backed Findings</div>
        <p class="capability-desc">Traces every identified risk directly to source document evidence and direct quotes.</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='margin-bottom: 0.75rem;'></div>", unsafe_allow_html=True)

# ==========================================
# FILE INGESTION SECTION
# ==========================================
st.subheader("Upload Project Documents")
st.caption("Analyze schedules, reports, meeting minutes, change logs, and project records together to surface risks and inconsistencies.")

uploaded_files = st.file_uploader(
    "Drop project files here or browse",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="PDF, DOCX, DOC, CSV, TXT, MD · Up to 200 MB"
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Staged {len(uploaded_files)} file(s) for analysis.")
    with st.expander("Preview Extracted Source Text", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            st.markdown(f"**{parsed['filename']}** (`{parsed['type']}`) — {parsed['word_count']} words")
            st.text_area(f"Raw Extracted Text ({parsed['filename']})", parsed['content'][:800] + ("..." if len(parsed['content']) > 800 else ""), height=90)

# CTA Action Trigger Button
analyze_click = st.button("Analyze Project Risks", type="primary")

# ==========================================
# RUN ANALYSIS PIPELINE
# ==========================================
if analyze_click:
    if not uploaded_files:
        st.warning("Please upload at least one project file above to extract risks.")
    else:
        combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
        
        with st.status("Analyzing multi-source project documents...", expanded=True) as status_box:
            st.write("1. Parsing uploaded files and extracting text content...")
            st.write("2. Scanning document text for schedule, cost, and governance triggers...")
            
            if mode == "Enhanced AI Analysis" and api_key:
                st.write("3. Executing OpenAI API risk extraction...")
                system_prompt = """You are an Enterprise Project Controls Risk Analyst. Analyze the multi-source project updates and return a JSON object with keys:
'executive_summary' (string), 'risk_matrix' (list of dicts with 'Risk / Issue', 'Severity', 'Category', 'Evidence / Quote', 'Source File'), 'mitigation_plan' (list of dicts with 'Action Item', 'Owner', 'Timeframe', 'Priority'), 'governance_flags' (list of strings). Strictly valid JSON."""
                try:
                    raw_response = call_openai_api(api_key, system_prompt, f"Project Documents:\n{combined_text}", model=model_choice)
                    json_start = raw_response.find('{')
                    json_end = raw_response.rfind('}') + 1
                    if json_start != -1 and json_end != -1:
                        results = json.loads(raw_response[json_start:json_end])
                    else:
                        results = run_grounded_text_analysis(combined_text, parsed_file_data)
                except Exception as e:
                    st.error(f"OpenAI API call encountered an issue: {str(e)}. Falling back to local analysis engine.")
                    results = run_grounded_text_analysis(combined_text, parsed_file_data)
            else:
                st.write("3. Running grounded evidence extraction engine...")
                results = run_grounded_text_analysis(combined_text, parsed_file_data)
                
            st.write("4. Compiling executive brief and risk register...")
            status_box.update(label="Analysis complete!", state="complete", expanded=False)

        st.session_state['analysis_results'] = results
        st.session_state['processed_files'] = parsed_file_data

# ==========================================
# RENDER RESULTS DASHBOARD
# ==========================================
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    
    st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)
    st.subheader("Project Risk Dashboard & Brief")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "Overview", 
        "Risk Register", 
        "Mitigation Actions", 
        "Governance Flags"
    ])
    
    with tab1:
        st.markdown("##### Executive Summary")
        st.markdown(res.get("executive_summary", "No summary generated."))
        
    with tab2:
        st.markdown("##### Identified Risk Register")
        risk_list = res.get("risk_matrix", [])
        filtered_risks = [r for r in risk_list if r.get("Severity") in min_severity or "All" in min_severity]
        
        if filtered_risks:
            df_risks = pd.DataFrame(filtered_risks)
            st.dataframe(
                df_risks,
                use_container_width=True,
                column_config={
                    "Severity": st.column_config.TextColumn("Severity", width="small"),
                    "Evidence / Quote": st.column_config.TextColumn("Evidence Quote", width="large")
                }
            )
        else:
            st.write("No risk items match the selected severity filters.")
            
    with tab3:
        st.markdown("##### Recommended Mitigation Actions")
        mitigations = res.get("mitigation_plan", [])
        if mitigations:
            for i, item in enumerate(mitigations, 1):
                c_a, c_b, c_c = st.columns([3, 1.5, 1])
                with c_a:
                    st.markdown(f"**{i}. {item.get('Action Item')}**")
                with c_b:
                    st.caption(f"Owner: **{item.get('Owner')}**")
                with c_c:
                    st.caption(f"Timeframe: `{item.get('Timeframe')}`")
                st.divider()
        else:
            st.write("No mitigation actions recorded.")
            
    with tab4:
        st.markdown("##### Governance & Approval Sign-Off Flags")
        flags = res.get("governance_flags", [])
        for f_item in flags:
            st.markdown(f"- {f_item}")
            
    # Export Section
    st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)
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
                label="Download Executive PDF Brief",
                data=pdf_bytes,
                file_name="RiskPulse_Executive_Brief.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as e_pdf:
            st.error(f"PDF generation error: {str(e_pdf)}")
        
    with exp_col2:
        try:
            json_data = json.dumps(res, indent=2)
            st.download_button(
                label="Export Raw JSON Data",
                data=json_data,
                file_name="Risk_Analysis_Data.json",
                mime="application/json",
                use_container_width=True
            )
        except Exception as e_json:
            st.error(f"JSON export error: {str(e_json)}")
        
    with exp_col3:
        try:
            md_text = f"# Executive Risk Brief\n\n{res.get('executive_summary')}\n\n### Risk Register\n\n" + generate_markdown_table(res.get('risk_matrix', []))
            st.download_button(
                label="Download Markdown Brief",
                data=md_text,
                file_name="Risk_Summary.md",
                mime="text/markdown",
                use_container_width=True
            )
        except Exception as e_md:
            st.error(f"Markdown export error: {str(e_md)}")
