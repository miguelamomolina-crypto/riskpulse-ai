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
    page_title="RiskPulse AI | Project Risk Intelligence Platform",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for B2B SaaS Aesthetics - Target Specific Selectors Only
st.markdown("""
<style>
    /* Main Theme Base */
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"], .main, .block-container {
        background-color: #f8fafc !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    /* Hide Streamlit Top Chrome Header */
    header[data-testid="stHeader"], div[data-testid="stDecoration"], div[data-testid="stStatusWidget"] {
        display: none !important;
    }
    .block-container {
        padding-top: 1.25rem !important;
        padding-bottom: 2rem !important;
        max-width: 1280px !important;
    }

    /* Navy Header Container */
    .hero-container {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
        padding: 1.25rem 2rem;
        border-radius: 12px;
        color: #ffffff;
        margin-bottom: 1.25rem;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
    }
    .hero-title {
        font-size: 1.65rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin-bottom: 0.25rem;
        color: #ffffff !important;
    }
    .hero-subtitle {
        font-size: 1rem;
        font-weight: 600;
        color: #cbd5e1 !important;
        margin-bottom: 0.25rem;
    }
    .hero-tagline {
        font-size: 0.9rem;
        color: #94a3b8 !important;
        line-height: 1.4;
    }

    /* Process Bar */
    .process-bar {
        display: flex;
        justify-content: space-between;
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.75rem 1.25rem;
        margin-bottom: 1.25rem;
    }
    .process-step {
        font-size: 0.85rem;
        font-weight: 600;
        color: #475569;
    }
    .process-step-active {
        color: #2563eb;
    }

    /* Capability Cards */
    .capability-card {
        background: #ffffff;
        padding: 0.85rem 1rem;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        height: 100%;
    }
    .capability-title {
        font-size: 0.875rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.25rem;
    }
    .capability-desc {
        font-size: 0.775rem;
        color: #475569;
        line-height: 1.35;
    }

    /* Upload Box Enhancement */
    [data-testid="stFileUploader"] {
        border: 2px dashed #2563eb !important;
        border-radius: 12px !important;
        background-color: #ffffff !important;
        padding: 1.25rem !important;
    }
    [data-testid="stFileUploader"] section {
        background-color: #ffffff !important;
    }
    [data-testid="stFileUploader"] label, [data-testid="stFileUploader"] span, [data-testid="stFileUploader"] p {
        color: #0f172a !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #ffffff !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    [data-testid="stFileUploader"] button:hover {
        border-color: #2563eb !important;
        color: #1d4ed8 !important;
    }
    [data-testid="stFileUploaderFile"] {
        background-color: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 6px !important;
        color: #0f172a !important;
    }
    [data-testid="stFileUploaderFile"] span {
        color: #0f172a !important;
    }
    [data-testid="stTooltipIcon"] {
        display: none !important;
    }

    /* Primary CTA Button */
    div.stButton > button[kind="primary"], div.stButton > button[type="primary"] {
        background-color: #2563eb !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        padding: 0.6rem 1.5rem !important;
        width: 320px !important;
        margin: 0 auto !important;
        display: block !important;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2) !important;
    }
    div.stButton > button[kind="primary"]:hover, div.stButton > button[type="primary"]:hover {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
    }

    /* Headings and Typography */
    h1, h2, h3, h4, h5, h6, p, label, span {
        color: #0f172a;
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
        "temperature": 0.1
    }
    
    req = urllib.request.Request(url, data=json.dumps(payload).encode("utf-8"), headers=headers)
    with urllib.request.urlopen(req, timeout=30) as response:
        res_data = json.loads(response.read().decode("utf-8"))
        return res_data["choices"][0]["message"]["content"]


# ==========================================
# STRICT GROUNDED LOCAL ANALYSIS PIPELINE
# ==========================================
def run_grounded_local_analysis(combined_text, file_list):
    """
    Analyzes document text sentence by sentence using strict risk keywords.
    Requires explicit adverse/risk language (delayed, late, behind, overrun, pending approval, etc.).
    Neutral mentions (e.g., 'schedule reviewed', 'budget is $2M') are NOT classified as risks.
    """
    risks = []
    
    # Regex patterns for explicit adverse risk language
    adverse_schedule_pattern = re.compile(r'\b(delayed|delay|late|behind|negative float|missed|blocked|postponed|slippage)\b', re.I)
    adverse_cost_pattern = re.compile(r'\b(over budget|overrun|cost variance|exceeded|overage|unforeseen cost|unapproved fee)\b', re.I)
    adverse_gov_pattern = re.compile(r'\b(pending approval|rejected|unresolved|not approved|lacks sign-off|lacks approval|concerns raised)\b', re.I)
    
    for f_info in file_list:
        fname = f_info["filename"]
        content = f_info["content"]
        
        # Split into sentences
        sentences = [s.strip() for s in re.split(r'[.\n]+', content) if len(s.strip()) > 15]
        
        for stmt in sentences:
            # Check for adverse schedule language
            if adverse_schedule_pattern.search(stmt):
                severity = "High" if re.search(r'\b(critical path|critical|severe|major)\b', stmt, re.I) else "Not available"
                risks.append({
                    "Risk / Issue": "Schedule Delay & Timeline Variance",
                    "Severity": severity,
                    "Category": "Schedule",
                    "Evidence / Quote": stmt[:250],
                    "Source File": fname
                })
            # Check for adverse cost language
            elif adverse_cost_pattern.search(stmt):
                severity = "High" if re.search(r'\b(critical|severe|major|budget overrun)\b', stmt, re.I) else "Not available"
                risks.append({
                    "Risk / Issue": "Cost & Budget Variance Exposure",
                    "Severity": severity,
                    "Category": "Financial",
                    "Evidence / Quote": stmt[:250],
                    "Source File": fname
                })
            # Check for adverse governance language
            elif adverse_gov_pattern.search(stmt):
                severity = "Medium" if re.search(r'\b(pending|unresolved)\b', stmt, re.I) else "Not available"
                risks.append({
                    "Risk / Issue": "Governance & Approval Bottleneck",
                    "Severity": severity,
                    "Category": "Governance",
                    "Evidence / Quote": stmt[:250],
                    "Source File": fname
                })

    # Deduplicate risks based on quote
    unique_risks = []
    seen_quotes = set()
    for r in risks:
        if r["Evidence / Quote"] not in seen_quotes:
            seen_quotes.add(r["Evidence / Quote"])
            unique_risks.append(r)

    # If no supported risks were detected
    if not unique_risks:
        file_names_str = ", ".join([f["filename"] for f in file_list]) if file_list else "uploaded documents"
        summary = f"Analysis completed across **{len(file_list)} uploaded source file(s)** ({file_names_str}). No explicit supported risks were detected in the provided document text."
        return {
            "raw_text": summary,
            "executive_summary": summary,
            "risk_matrix": [],
            "mitigation_plan": [],
            "governance_flags": []
        }

    # Summary when risks ARE detected
    file_names_str = ", ".join([f["filename"] for f in file_list])
    summary = f"Analysis completed across **{len(file_list)} uploaded source file(s)** ({file_names_str}). A total of **{len(unique_risks)} supported risk item(s)** were extracted directly from document text evidence."

    return {
        "raw_text": summary,
        "executive_summary": summary,
        "risk_matrix": unique_risks,
        "mitigation_plan": [],
        "governance_flags": []
    }


# ==========================================
# DEPENDENCY-FREE MARKDOWN TABLE GENERATOR
# ==========================================
def generate_markdown_table(data_list):
    """Converts list of dictionaries into a clean Markdown table string without tabulate dependency."""
    if not data_list:
        return "*No structured findings recorded.*"
    
    headers = list(data_list[0].keys())
    header_row = "| " + " | ".join(headers) + " |"
    separator_row = "| " + " | ".join(["---"] * len(headers)) + " |"
    
    data_rows = []
    for item in data_list:
        row_str = "| " + " | ".join([str(item.get(h, "")).replace("\n", " ").replace("|", "\\|") for h in headers]) + " |"
        data_rows.append(row_str)
        
    return "\n".join([header_row, separator_row] + data_rows)


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
    
    # Header
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, clean_text_for_pdf("Executive Risk Intelligence Report"), ln=True, align="L")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, clean_text_for_pdf("Generated by RiskPulse AI Platform | Confidential"), ln=True, align="L")
    pdf.ln(5)
    
    # Horizontal Divider Line
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
    
    # Section 2: Risk Matrix Table
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk Matrix"), ln=True)
    
    if risk_data:
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
    else:
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, "No explicit supported risks detected.", ln=True)
        
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
        ["Local Analysis", "Enhanced AI Analysis"],
        help="Local Analysis extracts grounded sentence evidence without API keys. Enhanced AI Analysis connects to OpenAI."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Enhanced AI Analysis":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("LLM Model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.warning("Please enter an OpenAI API key for Enhanced AI Analysis.")
            
    st.markdown("---")
    st.subheader("Risk Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Not available", "Low / Positive"],
        default=["High", "Medium", "Not available", "Low / Positive"]
    )


# ==========================================
# MAIN APP BODY & HEADER
# ==========================================

# Top Navy Header
st.markdown("""
<div class="hero-container">
    <div class="hero-title">RiskPulse AI</div>
    <div class="hero-subtitle">Project Risk Intelligence</div>
    <div class="hero-tagline">Turn project records into evidence-backed risk intelligence.</div>
</div>
""", unsafe_allow_html=True)

# Process Bar
st.markdown("""
<div class="process-bar">
    <span class="process-step process-step-active">1. Upload Documents</span>
    <span class="process-step">→ 2. Analyze & Extract</span>
    <span class="process-step">→ 3. Review Risk Register</span>
    <span class="process-step">→ 4. Executive Brief</span>
</div>
""", unsafe_allow_html=True)

# Four Capability Cards
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

# File Ingestion Section
st.subheader("Upload Project Documents")
st.caption("Analyze schedules, reports, meeting minutes, change logs, and project records together to surface risks and inconsistencies.")

uploaded_files = st.file_uploader(
    "Drop project files here or browse",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Supports PDF, DOCX, DOC, CSV, TXT, MD up to 200 MB"
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Staged {len(uploaded_files)} file(s) for risk analysis.")
    
    with st.expander("Preview Staged Documents", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            
            st.markdown(f"**{parsed['filename']}** (`{parsed['type']}`) — {parsed['word_count']} words")
            st.text_area(f"Source Text Snippet ({parsed['filename']})", parsed['content'][:800] + ("..." if len(parsed['content']) > 800 else ""), height=90)
            st.divider()

st.markdown("<br>", unsafe_allow_html=True)

# Primary CTA Button
analyze_click = st.button("Analyze Project Risks", type="primary", use_container_width=False)

# Run Analysis Pipeline
if analyze_click:
    if not uploaded_files:
        st.warning("Please upload at least one project document above to run analysis.")
    else:
        # Extract parsed data if not already done
        if not parsed_file_data:
            for f in uploaded_files:
                parsed_file_data.append(extract_text_from_file(f))

        combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
        valid_filenames = set([p['filename'] for p in parsed_file_data])
        
        with st.spinner("Analyzing uploaded source text and extracting evidence-backed risks..."):
            if mode == "Enhanced AI Analysis" and api_key:
                system_prompt = """You are an Enterprise AI Risk Analyst. Analyze the provided multi-source project updates and return a structured JSON object.

CRITICAL GROUNDING RULES:
1. Every risk must be strictly supported by the uploaded source text.
2. 'Evidence / Quote' MUST be an exact or near-exact source passage directly from the provided text.
3. 'Source File' MUST be the actual filename provided in the source text.
4. Never invent or hallucinate owners, dates, costs, percentages, schedule impacts, severity, project status, or mitigation deadlines.
5. Use 'Not available' when a field or severity cannot be supported by explicit source text.
6. Do NOT infer a risk from neutral mentions of schedule, cost, budget, approval, or governance (e.g., 'budget is $2M' or 'schedule reviewed' are neutral facts, NOT risks).
7. Return strictly valid JSON with keys: 'executive_summary', 'risk_matrix', 'mitigation_plan', 'governance_flags'. Each risk object in 'risk_matrix' must have keys: 'Risk / Issue', 'Severity', 'Category', 'Evidence / Quote', 'Source File'."""
                
                try:
                    raw_response = call_openai_api(api_key, system_prompt, f"Project Documents:\n{combined_text}", model=model_choice)
                    json_start = raw_response.find('{')
                    json_end = raw_response.rfind('}') + 1
                    if json_start != -1 and json_end != -1:
                        results = json.loads(raw_response[json_start:json_end])
                        
                        # Filter out invalid AI findings lacking evidence or valid filename
                        valid_risks = []
                        for r in results.get("risk_matrix", []):
                            quote = str(r.get("Evidence / Quote", "")).strip()
                            s_file = str(r.get("Source File", "")).strip()
                            if quote and quote.lower() != "not available" and (s_file in valid_filenames or not valid_filenames):
                                valid_risks.append(r)
                        results["risk_matrix"] = valid_risks
                    else:
                        results = run_grounded_local_analysis(combined_text, parsed_file_data)
                except Exception as e:
                    st.error(f"Enhanced AI Analysis encountered an error: {str(e)}. Falling back to Local Analysis.")
                    results = run_grounded_local_analysis(combined_text, parsed_file_data)
            else:
                results = run_grounded_local_analysis(combined_text, parsed_file_data)
                
            st.session_state['analysis_results'] = results
            st.session_state['processed_files'] = parsed_file_data


# Render Output Dashboard if available
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
        st.markdown(res.get("executive_summary", "No summary generated."))
        
    with tab2:
        st.markdown("##### Processed Source Files")
        if files_processed:
            for pf in files_processed:
                st.markdown(f"- **{pf['filename']}** ({pf['type']}) — {pf['word_count']} words | {pf['char_count']} characters")
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
                        width="small"
                    ),
                    "Evidence / Quote": st.column_config.TextColumn(
                        "Source Evidence / Quote",
                        width="large"
                    )
                }
            )
        else:
            st.write("No supported risks match the selected filter criteria.")
            
    with tab4:
        st.markdown("##### Executive Summary & Governance Notes")
        st.markdown(res.get("executive_summary", ""))
        flags = res.get("governance_flags", [])
        if flags:
            st.markdown("###### Governance & Sign-off Notes")
            for f_item in flags:
                st.markdown(f"- {f_item}")
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Export Deliverables Section
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
                label="Download Executive PDF Brief",
                data=pdf_bytes,
                file_name="Executive_Risk_Brief.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        except Exception as err:
            st.error(f"PDF generation unavailable: {str(err)}")
        
    with exp_col2:
        try:
            json_data = json.dumps(res, indent=2)
            st.download_button(
                label="Export JSON Analysis",
                data=json_data,
                file_name="Risk_Analysis_Data.json",
                mime="application/json",
                use_container_width=True
            )
        except Exception as err:
            st.error(f"JSON export unavailable: {str(err)}")
        
    with exp_col3:
        try:
            md_text = f"# Executive Risk Summary\n\n{res.get('executive_summary', '')}\n\n## Risk Register\n\n" + generate_markdown_table(res.get('risk_matrix', []))
            st.download_button(
                label="Download Markdown Brief",
                data=md_text,
                file_name="Risk_Summary.md",
                mime="text/markdown",
                use_container_width=True
            )
        except Exception as err:
            st.error(f"Markdown export unavailable: {str(err)}")
