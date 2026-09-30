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
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium B2B SaaS Aesthetics (Linear / Vercel / Ramp / ACC Style)
st.markdown("""
<style>
    /* Theme Overrides */
    .main {
        background-color: #f8fafc;
    }
    .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0f172a;
    }
    
    /* Top Header Bar */
    .app-header {
        background-color: #0f172a;
        padding: 1.25rem 2rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 1.5rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08);
    }
    .app-header-title {
        font-size: 1.5rem;
        font-weight: 700;
        letter-spacing: -0.02em;
        margin: 0;
        color: #ffffff;
    }
    .app-header-subtitle {
        font-size: 0.9rem;
        color: #94a3b8;
        margin-top: 0.25rem;
    }
    
    /* Executive Metric Cards */
    .metric-card {
        background: #ffffff;
        padding: 1rem 1.25rem;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        text-align: left;
    }
    .metric-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0f172a;
        margin-top: 0.2rem;
    }
    .metric-label {
        font-size: 0.78rem;
        color: #64748b;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }

    /* Container Cards */
    .content-card {
        background: #ffffff;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
        margin-bottom: 1.5rem;
    }
    
    /* Custom Dropzone Styling */
    .stFileUploader {
        border: 2px dashed #94a3b8;
        border-radius: 10px;
        background-color: #ffffff;
        padding: 0.75rem;
    }
    
    /* Table Styling */
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.88rem;
        margin-top: 0.5rem;
    }
    .custom-table th {
        background-color: #0f172a;
        color: #ffffff;
        text-align: left;
        padding: 0.65rem 0.85rem;
        font-weight: 600;
        font-size: 0.80rem;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    .custom-table td {
        padding: 0.65rem 0.85rem;
        border-bottom: 1px solid #e2e8f0;
        color: #1e293b;
        vertical-align: top;
    }
    .custom-table tr:hover {
        background-color: #f8fafc;
    }
    
    /* Severity Badges */
    .badge-critical {
        background-color: #fef2f2;
        color: #991b1b;
        padding: 0.2rem 0.55rem;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
        border: 1px solid #fecaca;
        display: inline-block;
    }
    .badge-warning {
        background-color: #fef3c7;
        color: #92400e;
        padding: 0.2rem 0.55rem;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
        border: 1px solid #fde68a;
        display: inline-block;
    }
    .badge-success {
        background-color: #d1fae5;
        color: #065f46;
        padding: 0.2rem 0.55rem;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
        border: 1px solid #a7f3d0;
        display: inline-block;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# FILE PARSER UTILITIES
# ==========================================
def extract_text_from_file(uploaded_file):
    """Parses text from PDF, DOCX, CSV, TXT, and MD files safely."""
    filename = uploaded_file.name
    ext = os.path.splitext(filename)[1].lower()
    text_content = ""
    file_type_label = ext.upper().replace('.', '')
    
    try:
        if ext == '.pdf':
            uploaded_file.seek(0)
            pdf_reader = pypdf.PdfReader(BytesIO(uploaded_file.read()))
            for page in pdf_reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text_content += extracted + "\n"
        elif ext in ['.docx', '.doc']:
            uploaded_file.seek(0)
            doc = docx.Document(BytesIO(uploaded_file.read()))
            for para in doc.paragraphs:
                if para.text:
                    text_content += para.text + "\n"
            for table in doc.tables:
                for row in table.rows:
                    row_text = " | ".join([cell.text.strip() for cell in row.cells if cell.text.strip()])
                    if row_text:
                        text_content += row_text + "\n"
        elif ext == '.csv':
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file)
            text_content = f"CSV File Data ({df.shape[0]} rows, {df.shape[1]} columns):\n"
            text_content += df.to_string(index=False)
        elif ext in ['.txt', '.md', '.json', '.log']:
            uploaded_file.seek(0)
            text_content = uploaded_file.read().decode("utf-8", errors="ignore")
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
    """Direct urllib REST call to OpenAI API to avoid external SDK version mismatches."""
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
# GROUNDED TEXT ANALYSIS PIPELINE (LOCAL)
# ==========================================
def run_local_text_analysis(file_list):
    """
    Performs deterministic, grounded analysis strictly on uploaded files.
    Extracts real sentence evidence without fabricating unstated claims.
    """
    risks = []
    mitigations = []
    governance_flags = []
    
    # Keyword categories
    delay_kw = re.compile(r'\b(delay|behind|critical path|float|postpone|timeline|late|lead time|milestone impact)\b', re.I)
    cost_kw = re.compile(r'\b(budget|cost|overrun|spend|change order|pco|allowance|contingency|variance|exposure|fee|asr)\b', re.I)
    gov_kw = re.compile(r'\b(approval|pending|sign-off|unapproved|governance|decision|holding|authorization|signature)\b', re.I)
    quality_kw = re.compile(r'\b(submittal|rfi|clearance|specification|drawing|inspection|defect|obstruction)\b', re.I)

    combined_summary_bullets = []

    for f_item in file_list:
        filename = f_item['filename']
        content = f_item['content']
        lines = [line.strip() for line in content.split('\n') if len(line.strip()) > 15]
        
        for line in lines:
            # Check for schedule delay risks
            if delay_kw.search(line):
                risks.append({
                    "Risk / Issue": f"Schedule / Milestone Risk in {filename}",
                    "Severity": "High",
                    "Category": "Schedule",
                    "Evidence / Quote": line[:150],
                    "Source File": filename
                })
                mitigations.append({
                    "Action Item": f"Review schedule variance referenced in {filename}: '{line[:60]}...'",
                    "Owner": "Project Manager / Owner Rep",
                    "Timeframe": "Immediate",
                    "Priority": "High"
                })
            # Check for cost & financial risks
            elif cost_kw.search(line):
                risks.append({
                    "Risk / Issue": f"Cost / Financial Exposure in {filename}",
                    "Severity": "High",
                    "Category": "Financial",
                    "Evidence / Quote": line[:150],
                    "Source File": filename
                })
                mitigations.append({
                    "Action Item": f"Evaluate cost impact / change order in {filename}: '{line[:60]}...'",
                    "Owner": "Finance / Cost Manager",
                    "Timeframe": "This Week",
                    "Priority": "High"
                })
            # Check for governance / approval bottlenecks
            elif gov_kw.search(line):
                risks.append({
                    "Risk / Issue": f"Governance / Approval Bottleneck in {filename}",
                    "Severity": "Medium",
                    "Category": "Governance",
                    "Evidence / Quote": line[:150],
                    "Source File": filename
                })
                governance_flags.append(f"**Approval Required ({filename}):** {line[:120]}")
            # Check for submittal / RFI / technical issues
            elif quality_kw.search(line) and len(risks) < 8:
                risks.append({
                    "Risk / Issue": f"Technical / Submittal Item in {filename}",
                    "Severity": "Medium",
                    "Category": "Technical",
                    "Evidence / Quote": line[:150],
                    "Source File": filename
                })

    # Deduplicate risks based on evidence text
    unique_risks = []
    seen_evidence = set()
    for r in risks:
        key = r["Evidence / Quote"][:50].lower()
        if key not in seen_evidence:
            seen_evidence.add(key)
            unique_risks.append(r)

    # Fallback if text is clean or short
    if not unique_risks:
        unique_risks.append({
            "Risk / Issue": "Document Ingested — No Critical Risk Triggers Found",
            "Severity": "Low",
            "Category": "General",
            "Evidence / Quote": f"Text from {len(file_list)} document(s) parsed cleanly without schedule or budget risk keywords.",
            "Source File": file_list[0]['filename'] if file_list else "Uploaded_Doc"
        })

    # Build executive summary
    doc_names = ", ".join([f["filename"] for f in file_list])
    high_count = sum(1 for r in unique_risks if r["Severity"] == "High")
    med_count = sum(1 for r in unique_risks if r["Severity"] == "Medium")
    
    exec_summary = (
        f"Analysis completed across **{len(file_list)} uploaded source file(s)** ({doc_names}). "
        f"A total of **{len(unique_risks)} actionable items** were identified, including "
        f"**{high_count} High-severity exposure(s)** and **{med_count} Medium-severity watch item(s)**. "
        f"Key schedule, financial, and governance findings are extracted below directly from source evidence."
    )

    if not mitigations:
        mitigations.append({
            "Action Item": "Monitor project progress against baseline deliverables.",
            "Owner": "Project Lead",
            "Timeframe": "Weekly Review",
            "Priority": "Low"
        })

    if not governance_flags:
        governance_flags.append("No explicit unapproved governance bottlenecks detected in uploaded text.")

    return {
        "executive_summary": exec_summary,
        "risk_matrix": unique_risks[:12],
        "mitigation_plan": mitigations[:6],
        "governance_flags": governance_flags[:5]
    }


# ==========================================
# DETERMINISTIC EVM & CPM DATA ANALYZER
# ==========================================
def analyze_evm_cpm_data(file_list):
    """
    Extracts deterministic EVM and CPM metrics ONLY if actual structured data exists in CSV files.
    Returns None if structured EVM/CPM data is unavailable.
    """
    evm_metrics = None
    cpm_metrics = None

    for f_item in file_list:
        if f_item['type'] == 'CSV':
            try:
                # Read CSV
                f_bytes = BytesIO(f_item['content'].encode('utf-8'))
                df = pd.read_csv(f_bytes)
                cols_lower = [c.lower().strip() for c in df.columns]

                # Check for EVM columns
                pv_col = next((c for c in df.columns if 'planned' in c.lower() or c.lower() == 'pv'), None)
                ev_col = next((c for c in df.columns if 'earned' in c.lower() or c.lower() == 'ev'), None)
                ac_col = next((c for c in df.columns if 'actual' in c.lower() or c.lower() == 'ac'), None)
                bac_col = next((c for c in df.columns if 'budget' in c.lower() or c.lower() == 'bac'), None)

                if pv_col and ev_col and ac_col:
                    pv_sum = pd.to_numeric(df[pv_col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce').sum()
                    ev_sum = pd.to_numeric(df[ev_col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce').sum()
                    ac_sum = pd.to_numeric(df[ac_col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce').sum()
                    bac_sum = pd.to_numeric(df[bac_col].astype(str).str.replace('$', '').str.replace(',', ''), errors='coerce').sum() if bac_col else 0

                    cpi = round(ev_sum / ac_sum, 2) if ac_sum > 0 else 0.0
                    spi = round(ev_sum / pv_sum, 2) if pv_sum > 0 else 0.0
                    cv = ev_sum - ac_sum
                    sv = ev_sum - pv_sum

                    evm_metrics = {
                        "Source File": f_item['filename'],
                        "Planned Value (PV)": f"${pv_sum:,.2f}",
                        "Earned Value (EV)": f"${ev_sum:,.2f}",
                        "Actual Cost (AC)": f"${ac_sum:,.2f}",
                        "Budget at Completion (BAC)": f"${bac_sum:,.2f}" if bac_sum > 0 else "N/A",
                        "Cost Variance (CV)": f"${cv:,.2f}",
                        "Schedule Variance (SV)": f"${sv:,.2f}",
                        "CPI": cpi,
                        "SPI": spi
                    }

                # Check for CPM / Float columns
                float_col = next((c for c in df.columns if 'float' in c.lower() or 'slack' in c.lower()), None)
                cp_col = next((c for c in df.columns if 'critical' in c.lower() or 'path' in c.lower()), None)

                if float_col or cp_col:
                    float_vals = pd.to_numeric(df[float_col].astype(str).str.replace('days', '').str.strip(), errors='coerce') if float_col else pd.Series()
                    min_float = float_vals.min() if not float_vals.empty else None
                    cp_count = (df[cp_col].astype(str).str.upper().str.contains('YES|TRUE|1')).sum() if cp_col else 0

                    cpm_metrics = {
                        "Source File": f_item['filename'],
                        "Min Total Float (Days)": f"{min_float} Days" if min_float is not None else "N/A",
                        "Critical Path Tasks Count": cp_count,
                        "Total Tracked Activities": len(df)
                    }

            except Exception:
                continue

    return evm_metrics, cpm_metrics


# ==========================================
# PDF REPORT GENERATOR WITH UNICODE SANITIZER
# ==========================================
def clean_text_for_pdf(text):
    """Sanitizes text to prevent fpdf Latin-1 encoding errors with unicode/emojis."""
    if not text:
        return ""
    emoji_map = {
        '⚠️': '[WARNING]', '🔒': '[IT GOVERNANCE]', '💰': '[FINANCIAL]', 
        '👥': '[HUMAN-IN-LOOP]', '📊': '[DATA]', '🛠️': '[ACTION]', 
        '📄': '[DOC]', '⚡': '[FAST]', '👤': '[OWNER]', '⏱️': '[TIME]', 
        '💡': '[NOTE]', '🟢': '[OK]', '🔴': '[CRITICAL]', '🟡': '[WATCH]'
    }
    for emoji, rep in emoji_map.items():
        text = text.replace(emoji, rep)
    return text.encode('latin-1', 'replace').decode('latin-1')


def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    """Generates an executive PDF report using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Header
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, clean_text_for_pdf("RiskPulse AI — Executive Risk Intelligence Report"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, clean_text_for_pdf("Grounded Multi-Source Analysis | Confidential"), ln=True, align="L")
    pdf.ln(4)
    
    # Divider
    pdf.set_draw_color(226, 232, 240)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(6)
    
    # Section 1: Executive Summary
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 7, clean_text_for_pdf("1. Executive Health Summary"), ln=True)
    pdf.set_font("Helvetica", "", 9.5)
    pdf.set_text_color(51, 65, 85)
    
    clean_summary = clean_text_for_pdf(summary_text.replace('#', '').replace('*', '').strip())
    pdf.multi_cell(0, 5, clean_summary)
    pdf.ln(6)
    
    # Section 2: Risk Matrix Table
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 7, clean_text_for_pdf("2. Key Risk & Issue Identification Matrix"), ln=True)
    
    pdf.set_font("Helvetica", "B", 8.5)
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
    pdf.cell(0, 7, clean_text_for_pdf("3. Governance & Decision Items"), ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(51, 65, 85)
    
    for flag in governance_flags:
        clean_flag = clean_text_for_pdf(str(flag).replace('*', ''))
        pdf.multi_cell(0, 5, f"- {clean_flag}")
        pdf.ln(1)
        
    return bytes(pdf.output())


# ==========================================
# SIDEBAR & SYSTEM NAVIGATION
# ==========================================
with st.sidebar:
    st.markdown("### 🛡️ RiskPulse AI")
    st.caption("Project Risk Intelligence Platform")
    st.markdown("---")
    
    st.subheader("⚙️ System Configuration")
    
    mode = st.radio(
        "Execution Engine",
        ["Local Grounded Analysis (No Key)", "Live OpenAI API"],
        help="Local analysis extracts risk statements directly from uploaded files without external API dependency."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Live OpenAI API":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("LLM Engine Model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.warning("Please enter your OpenAI API key to execute live calls.")
            
    st.markdown("---")
    st.subheader("🎯 Risk Threshold Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low"],
        default=["High", "Medium", "Low"]
    )


# ==========================================
# MAIN APP HEADER
# ==========================================
st.markdown("""
<div class="app-header">
    <div>
        <div class="app-header-title">RiskPulse AI — Project Risk Intelligence</div>
        <div class="app-header-subtitle">Turn fragmented project records into evidence-backed risk intelligence</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ==========================================
# FILE INGESTION SECTION
# ==========================================
st.markdown('<div class="content-card">', unsafe_allow_html=True)
st.subheader("📁 Source Document Ingestion")
st.caption("Upload project artifacts for grounded cross-analysis (PDF, DOCX, DOC, CSV, TXT, MD · Up to 200 MB).")

uploaded_files = st.file_uploader(
    "Drop project files here",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload project status reports, meeting notes, change logs, and schedule exports."
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Loaded **{len(uploaded_files)} file(s)** for analysis.")
    for f in uploaded_files:
        parsed = extract_text_from_file(f)
        parsed_file_data.append(parsed)
    
    with st.expander("🔍 Inspect Extracted Text & Metadata per File", expanded=False):
        for p in parsed_file_data:
            st.markdown(f"**📄 {p['filename']}** (`{p['type']}`) — {p['word_count']} words | {p['char_count']} chars")
            st.text_area(f"Raw Text Preview ({p['filename']})", p['content'][:600] + ("..." if len(p['content']) > 600 else ""), height=80)
st.markdown('</div>', unsafe_allow_html=True)


# ==========================================
# EXECUTION BUTTON & PIPELINE
# ==========================================
col_btn, _ = st.columns([2, 1])
with col_btn:
    analyze_click = st.button("Analyze Project Risks", type="primary", use_container_width=True)

if analyze_click:
    if not uploaded_files:
        st.warning("⚠️ Please upload at least one project file above to execute risk analysis.")
    else:
        with st.status("Executing Grounded Risk Extraction Pipeline...", expanded=True) as status:
            st.write("1. Parsing uploaded documents & extracting text...")
            combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
            
            st.write("2. Scanning text for schedule, financial, and governance risk triggers...")
            
            if mode == "Live OpenAI API" and api_key:
                st.write("3. Calling OpenAI API for structured risk synthesis...")
                system_prompt = """You are an expert Enterprise AI Risk Analyst. Analyze the provided multi-source project updates and return a structured json object containing:
1. 'executive_summary': a 2-3 paragraph summary of project health, status, and major blocks.
2. 'risk_matrix': a list of objects with keys 'Risk / Issue', 'Severity' (High/Medium/Low), 'Category' (Technical/Financial/Workflow/Governance), 'Evidence / Quote', 'Source File'.
3. 'mitigation_plan': a list of objects with keys 'Action Item', 'Owner', 'Timeframe', 'Priority'.
4. 'governance_flags': a list of strings detailing IT security or governance approvals required.

Return strictly valid JSON grounded ONLY in the provided text. Do not invent unstated facts."""
                
                try:
                    raw_response = call_openai_api(api_key, system_prompt, f"Project Documents:\n{combined_text}", model=model_choice)
                    json_start = raw_response.find('{')
                    json_end = raw_response.rfind('}') + 1
                    if json_start != -1 and json_end != -1:
                        results = json.loads(raw_response[json_start:json_end])
                    else:
                        results = run_local_text_analysis(parsed_file_data)
                except Exception as e:
                    st.error(f"Error calling OpenAI API: {str(e)}. Falling back to local analysis engine.")
                    results = run_local_text_analysis(parsed_file_data)
            else:
                st.write("3. Running local grounded text analysis engine...")
                results = run_local_text_analysis(parsed_file_data)

            st.write("4. Generating executive report & action items...")
            status.update(label="Analysis Complete", state="complete", expanded=False)

        st.session_state['analysis_results'] = results
        st.session_state['processed_files'] = parsed_file_data


# ==========================================
# DASHBOARD RESULTS VIEW
# ==========================================
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    files_processed = st.session_state.get('processed_files', [])
    risk_list = res.get("risk_matrix", [])
    
    # Filter by selected severities
    filtered_risks = [r for r in risk_list if r.get("Severity") in min_severity or "All" in min_severity]
    high_count = sum(1 for r in risk_list if r.get("Severity") == "High")
    med_count = sum(1 for r in risk_list if r.get("Severity") == "Medium")

    st.markdown("---")
    
    # Top Metrics Strip (Dynamic from actual analysis)
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Files Analyzed</div><div class="metric-value">{len(files_processed)}</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Active Risks</div><div class="metric-value">{len(risk_list)}</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="metric-card"><div class="metric-label">High Severity</div><div class="metric-value" style="color:#991b1b;">{high_count}</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="metric-card"><div class="metric-label">Medium Severity</div><div class="metric-value" style="color:#92400e;">{med_count}</div></div>', unsafe_allow_html=True)
    with k5:
        mitigations_count = len(res.get("mitigation_plan", []))
        st.markdown(f'<div class="metric-card"><div class="metric-label">Open Actions</div><div class="metric-value">{mitigations_count}</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Executive Summary Card
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.subheader("📋 Executive Health Summary")
    st.markdown(res.get("executive_summary", "No summary generated."))
    st.markdown('</div>', unsafe_allow_html=True)

    # Risk & Issue Register Table Card
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.subheader("🚨 Risk & Issue Register")
    st.caption("Evidence-backed risk statements extracted directly from source documents.")
    
    if filtered_risks:
        table_html = '<table class="custom-table">'
        table_html += '<thead><tr><th>Risk / Issue</th><th>Severity</th><th>Category</th><th>Source File</th><th>Evidence Quote</th></tr></thead><tbody>'
        
        for r in filtered_risks:
            sev = str(r.get("Severity", "Low"))
            badge_class = "badge-critical" if "High" in sev else ("badge-warning" if "Medium" in sev else "badge-success")
            
            table_html += f'<tr>'
            table_html += f'<td><strong>{r.get("Risk / Issue", "")}</strong></td>'
            table_html += f'<td><span class="{badge_class}">{sev}</span></td>'
            table_html += f'<td>{r.get("Category", "General")}</td>'
            table_html += f'<td><code>{r.get("Source File", "Doc")}</code></td>'
            table_html += f'<td>"{r.get("Evidence / Quote", "")}"</td>'
            table_html += f'</tr>'
        table_html += '</tbody></table>'
        st.markdown(table_html, unsafe_allow_html=True)
    else:
        st.info("No risks match the selected severity filter.")
    st.markdown('</div>', unsafe_allow_html=True)

    # EVM / CPM Analytics Card (Dynamic or Graceful Empty State)
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.subheader("📈 Earned Value (EVM) & Critical Path (CPM) Analytics")
    
    evm_res, cpm_res = analyze_evm_cpm_data(files_processed)
    
    if evm_res or cpm_res:
        e_col, c_col = st.columns(2)
        with e_col:
            if evm_res:
                st.markdown(f"**EVM Analysis ({evm_res['Source File']})**")
                st.write(f"- Planned Value (PV): `{evm_res['Planned Value (PV)']}`")
                st.write(f"- Earned Value (EV): `{evm_res['Earned Value (EV)']}`")
                st.write(f"- Actual Cost (AC): `{evm_res['Actual Cost (AC)']}`")
                st.write(f"- CPI: `{evm_res['CPI']}` | SPI: `{evm_res['SPI']}`")
            else:
                st.info("EVM metrics unavailable in uploaded files.")
        with c_col:
            if cpm_res:
                st.markdown(f"**CPM Schedule Logic ({cpm_res['Source File']})**")
                st.write(f"- Min Total Float: `{cpm_res['Min Total Float (Days)']}`")
                st.write(f"- Critical Path Tasks: `{cpm_res['Critical Path Tasks Count']}`")
                st.write(f"- Total Activities: `{cpm_res['Total Tracked Activities']}`")
            else:
                st.info("CPM schedule logic unavailable in uploaded files.")
    else:
        st.info("ℹ️ **EVM / CPM Numerical Analysis Unavailable:** Structured EVM columns (PV, EV, AC) and schedule logic fields (Total Float, Critical Path) were not found in the uploaded documents. Analysis is limited to text risk extraction.")
    st.markdown('</div>', unsafe_allow_html=True)

    # Actionable Mitigation & Governance Card
    a_col, g_col = st.columns(2)
    with a_col:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        st.subheader("🛠️ Recommended Actions")
        mitigations = res.get("mitigation_plan", [])
        if mitigations:
            for item in mitigations:
                st.markdown(f"**Action:** {item.get('Action Item')}")
                st.caption(f"👤 Owner: **{item.get('Owner')}** | ⏱️ Timeframe: `{item.get('Timeframe')}`")
                st.divider()
        else:
            st.write("No specific actions generated.")
        st.markdown('</div>', unsafe_allow_html=True)

    with g_col:
        st.markdown('<div class="content-card">', unsafe_allow_html=True)
        st.subheader("🔒 Governance & Decision Flags")
        flags = res.get("governance_flags", [])
        if flags:
            for f_item in flags:
                st.markdown(f"- {f_item}")
        else:
            st.write("No governance flags identified.")
        st.markdown('</div>', unsafe_allow_html=True)

    # Export Section
    st.markdown('<div class="content-card">', unsafe_allow_html=True)
    st.subheader("📥 Export Deliverables")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    
    with exp_col1:
        pdf_bytes = generate_pdf_report(
            res.get("executive_summary", ""),
            res.get("risk_matrix", []),
            res.get("mitigation_plan", []),
            res.get("governance_flags", [])
        )
        st.download_button(
            label="📄 Download Executive Brief (PDF)",
            data=pdf_bytes,
            file_name="RiskPulse_Executive_Report.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
    with exp_col2:
        json_data = json.dumps(res, indent=2)
        st.download_button(
            label="💾 Download Raw Analysis (JSON)",
            data=json_data,
            file_name="RiskPulse_Analysis_Data.json",
            mime="application/json",
            use_container_width=True
        )
        
    with exp_col3:
        md_text = f"{res.get('executive_summary')}\n\n### Risk Register\n" + pd.DataFrame(res.get('risk_matrix', [])).to_markdown(index=False)
        st.download_button(
            label="📋 Download Markdown Summary",
            data=md_text,
            file_name="RiskPulse_Summary.md",
            mime="text/markdown",
            use_container_width=True
        )
    st.markdown('</div>', unsafe_allow_html=True)
