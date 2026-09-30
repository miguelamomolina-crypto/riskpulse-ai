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
    initial_sidebar_state="collapsed"
)

# Custom CSS for Clean Light B2B SaaS Aesthetics
st.markdown("""
<style>
:root {
    --background-color: #f8fafc;
    --secondary-background-color: #ffffff;
    --text-color: #0f172a;
    --primary-color: #2563eb;
}

html, body, .stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="stMainBlockContainer"],
.main,
.block-container {
    background: #f8fafc !important;
    color: #0f172a !important;
}

.stApp {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
}

[data-testid="stMarkdownContainer"],
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] span,
.stCaption,
label {
    color: #0f172a;
}

.hero-container {
    background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
    padding: 1.15rem 1.5rem;
    border-radius: 10px;
    color: #ffffff !important;
    margin-bottom: 1rem;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.10);
}
.hero-container * { color: inherit; }
.hero-title {
    font-size: 1.55rem;
    font-weight: 800;
    letter-spacing: -0.025em;
    color: #ffffff !important;
    margin-bottom: 0.2rem;
    display: flex;
    align-items: baseline;
    gap: 0.7rem;
}
.hero-subtitle {
    font-size: 0.92rem;
    color: #cbd5e1 !important;
    line-height: 1.4;
}

.process-bar {
    background: #ffffff;
    padding: 0.75rem 1.1rem;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
    margin-bottom: 1.1rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.75rem;
    font-size: 0.84rem;
    font-weight: 600;
    color: #475569 !important;
}
.process-bar span { color: #475569 !important; }
.process-step-active { color: #2563eb !important; }
.process-arrow { color: #94a3b8 !important; }

.capability-card {
    background: #ffffff;
    padding: 1rem 1.1rem;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
    text-align: left;
    min-height: 112px;
}
.capability-title {
    font-size: 0.92rem;
    font-weight: 700;
    color: #0f172a !important;
    margin-bottom: 0.35rem;
}
.capability-desc {
    font-size: 0.8rem;
    color: #64748b !important;
    line-height: 1.45;
}

[data-testid="stFileUploader"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    padding: 0.35rem !important;
}
[data-testid="stFileUploaderDropzone"] {
    background: #f8fafc !important;
    border: 1.5px dashed #93c5fd !important;
    border-radius: 8px !important;
    color: #0f172a !important;
}
[data-testid="stFileUploaderDropzone"] * {
    color: #0f172a !important;
}
[data-testid="stFileUploaderDropzone"] button {
    background: #ffffff !important;
    color: #0f172a !important;
    border: 1px solid #cbd5e1 !important;
}
[data-testid="stFileUploaderDropzone"] button:hover {
    border-color: #2563eb !important;
    color: #1d4ed8 !important;
}

div.stButton > button[kind="primary"],
div.stButton > button[data-testid="stBaseButton-primary"] {
    background-color: #2563eb !important;
    color: #ffffff !important;
    border-color: #2563eb !important;
    font-weight: 600 !important;
    border-radius: 7px !important;
}
div.stButton > button[kind="primary"]:hover,
div.stButton > button[data-testid="stBaseButton-primary"]:hover {
    background-color: #1d4ed8 !important;
    border-color: #1d4ed8 !important;
    color: #ffffff !important;
}

[data-testid="stSidebar"] {
    background: #ffffff !important;
}
[data-testid="stSidebar"] * {
    color: #0f172a;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""

<style>
/* Header contrast hotfix — keep all text inside navy banner readable */
.hero-container,
.hero-container [data-testid="stMarkdownContainer"],
.hero-container [data-testid="stMarkdownContainer"] *,
.hero-container .hero-title,
.hero-container .hero-title *,
.hero-container .hero-subtitle,
.hero-container .hero-subtitle * {
    color: #ffffff !important;
}

.hero-container .hero-title span:last-child,
.hero-container .hero-subtitle {
    color: #cbd5e1 !important;
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


def run_grounded_text_analysis(file_list):
    """Local fallback that extracts only evidence actually present in uploaded files."""
    category_rules = {
        "Schedule": [
            r"delay(?:ed|s)?", r"behind schedule", r"late", r"slippage", r"float",
            r"critical path", r"milestone", r"schedule", r"postpone(?:d)?"
        ],
        "Cost / Change": [
            r"budget", r"cost", r"overrun", r"change order", r"contingency",
            r"allowance", r"spend", r"variance"
        ],
        "Governance": [
            r"approval", r"sign[- ]?off", r"security", r"permission", r"governance",
            r"compliance", r"authorization"
        ],
        "Resource / Delivery": [
            r"resource", r"staff", r"reassign(?:ed|ment)?", r"resign(?:ed|ation)?",
            r"turnover", r"shortage", r"constraint", r"bottleneck"
        ],
        "Quality / Issue": [
            r"defect", r"failed", r"failure", r"rework", r"issue", r"problem",
            r"nonconformance", r"punch"
        ],
    }

    high_terms = re.compile(
        r"critical|severe|stop work|failed|failure|overrun|negative float|"
        r"behind schedule|delay(?:ed)?|blocked|breach|urgent", re.I
    )
    medium_terms = re.compile(
        r"risk|concern|pending|issue|variance|constraint|rework|slippage|"
        r"approval|change order|shortage", re.I
    )

    risks = []
    seen = set()

    for file_data in file_list:
        content = file_data.get("content", "")
        filename = file_data.get("filename", "Unknown source")

        candidates = re.split(r"(?<=[.!?])\s+|\n+", content)
        for raw in candidates:
            sentence = re.sub(r"\s+", " ", raw).strip()
            if len(sentence) < 20:
                continue

            matched_category = None
            matched_keyword = None
            for category, patterns in category_rules.items():
                for pattern in patterns:
                    if re.search(pattern, sentence, re.I):
                        matched_category = category
                        matched_keyword = pattern
                        break
                if matched_category:
                    break

            if not matched_category:
                continue

            key = (filename, sentence.lower())
            if key in seen:
                continue
            seen.add(key)

            if high_terms.search(sentence):
                severity = "High"
            elif medium_terms.search(sentence):
                severity = "Medium"
            else:
                severity = "Low / Positive"

            keyword_label = re.sub(r"[^A-Za-z /-]", "", matched_keyword or "").strip()
            title = f"{matched_category} signal"
            if keyword_label:
                title = f"{matched_category} signal: {keyword_label.title()}"

            risks.append({
                "Risk / Issue": title,
                "Severity": severity,
                "Category": matched_category,
                "Evidence / Quote": sentence[:500],
                "Source File": filename,
            })

    risks = risks[:25]

    high_count = sum(1 for r in risks if r["Severity"] == "High")
    medium_count = sum(1 for r in risks if r["Severity"] == "Medium")

    if risks:
        summary = (
            f"### Executive Health Summary\n"
            f"RiskPulse reviewed **{len(file_list)} uploaded source file(s)** and surfaced "
            f"**{len(risks)} evidence-backed signal(s)** using local text analysis. "
            f"The current set contains **{high_count} high-severity** and "
            f"**{medium_count} medium-severity** signal(s).\n\n"
            "This local fallback does not infer unsupported schedule, cost, or governance "
            "metrics. Each finding below is tied to text found in an uploaded source."
        )
    else:
        summary = (
            "### Executive Health Summary\n"
            f"RiskPulse reviewed **{len(file_list)} uploaded source file(s)**. "
            "No risk-trigger language was detected by the local rules-based analysis. "
            "This does not prove the project has no risks; it means the local fallback "
            "did not find supported text signals in the uploaded material."
        )

    mitigations = []
    for risk in risks[:8]:
        mitigations.append({
            "Action Item": f"Review and validate the evidence for: {risk['Risk / Issue']}",
            "Owner": "Project Team",
            "Timeframe": "Next project review",
            "Priority": risk["Severity"],
        })

    governance_flags = []
    for risk in [r for r in risks if r["Category"] == "Governance"][:5]:
        governance_flags.append(
            f"Review required: {risk['Evidence / Quote']} (Source: {risk['Source File']})"
        )

    return {
        "raw_text": summary,
        "executive_summary": summary,
        "risk_matrix": risks,
        "mitigation_plan": mitigations,
        "governance_flags": governance_flags,
    }


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
    
    # Strip any other non-latin1 characters
    return text.encode('latin-1', 'replace').decode('latin-1')


def generate_pdf_report(summary_text, risk_data, mitigation_data, governance_flags):
    """Generates an executive PDF report using fpdf2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Header
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(15, 23, 42) # Navy
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
    pdf.cell(0, 8, clean_text_for_pdf("1. Executive Health Summary"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    
    clean_summary = clean_text_for_pdf(summary_text.replace('#', '').replace('*', '').strip())
    pdf.multi_cell(0, 5, clean_summary)
    pdf.ln(6)
    
    # Section 2: Risk Matrix Table
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, clean_text_for_pdf("2. Key Risk & Issue Identification Matrix"), ln=True)
    
    # Table Header
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
    pdf.cell(0, 8, clean_text_for_pdf("3. IT Governance & Sign-off Flags"), ln=True)
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
    st.image("https://img.icons8.com/isometric-folders/100/shield.png", width=50)
    st.title("RiskPulse AI")
    st.caption("Project Risk Intelligence Platform")
    st.markdown("---")
    
    st.subheader("⚙️ System Configuration")
    
    mode = st.radio(
        "Execution Engine",
        ["Local Analysis", "Enhanced AI Analysis"],
        help="Local Analysis uses evidence-based rules. Enhanced AI Analysis uses the configured model."
    )
    
    api_key = ""
    model_choice = "gpt-4o-mini"
    if mode == "Enhanced AI Analysis":
        api_key = st.text_input("OpenAI API Key", type="password", placeholder="sk-...")
        model_choice = st.selectbox("LLM Engine Model", ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"])
        if not api_key:
            st.warning("Please enter your OpenAI API key to execute live calls.")
            
    st.markdown("---")
    st.subheader("🎯 Risk Threshold Filters")
    min_severity = st.multiselect(
        "Display Severities",
        ["High", "Medium", "Low / Positive"],
        default=["High", "Medium", "Low / Positive"]
    )


# ==========================================
# MAIN APP BODY
# ==========================================

# Product Header
st.markdown("""
<div class="hero-container">
    <div class="hero-title">
        <span>RiskPulse AI</span>
        <span style="font-size: 1rem; font-weight: 500; color: #94a3b8;">Project Risk Intelligence</span>
    </div>
    <div class="hero-subtitle">
        Turn project records into evidence-backed risk intelligence.
    </div>
</div>
""", unsafe_allow_html=True)

# Process Indicator Bar
st.markdown("""
<div class="process-bar">
    <span class="process-step-active">1. Upload Documents</span>
    <span class="process-arrow">→</span>
    <span>2. Analyze & Extract</span>
    <span class="process-arrow">→</span>
    <span>3. Review Risk Register</span>
    <span class="process-arrow">→</span>
    <span>4. Executive Brief</span>
</div>
""", unsafe_allow_html=True)

# 4 Capability Cards
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

# File Upload Section
st.subheader("Upload Project Documents")
st.caption("Analyze schedules, reports, meeting minutes, change logs, and project records together to surface risks and inconsistencies.")

uploaded_files = st.file_uploader(
    "Drag and drop files here or click to browse",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload project status documents to analyze."
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Successfully staged **{len(uploaded_files)} file(s)** for analysis.")
    
    with st.expander("🔍 Preview Extracted Text & Metadata per File", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            
            st.markdown(f"**📄 {parsed['filename']}** (`{parsed['type']}`) — {parsed['word_count']} words | {parsed['char_count']} characters")
            st.text_area(f"Raw Extracted Text ({parsed['filename']})", parsed['content'][:1000] + ("..." if len(parsed['content']) > 1000 else ""), height=100)
            st.divider()

st.markdown("<br>", unsafe_allow_html=True)

# Primary Action Trigger
col_btn1, col_btn2 = st.columns([1, 2])
with col_btn1:
    analyze_click = st.button("Analyze Project Risks", type="primary", use_container_width=True)

# Run Analysis Pipeline
if analyze_click:
    if not uploaded_files:
        st.warning("Upload at least one project document before running analysis.")
        st.stop()

    combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
    
    with st.spinner("Analyzing project records, extracting risk evidence, and compiling executive brief..."):
        if mode == "Enhanced AI Analysis" and api_key:
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
                    results = run_grounded_text_analysis(parsed_file_data)
            except Exception as e:
                st.error(f"Error calling OpenAI API: {str(e)}. Falling back to local synthesis engine.")
                results = run_grounded_text_analysis(parsed_file_data)
        else:
            results = run_grounded_text_analysis(parsed_file_data)
            
        st.session_state['analysis_results'] = results
        st.session_state['processed_files'] = parsed_file_data


# Render Output Dashboard if available
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    files_processed = st.session_state.get('processed_files', [])
    
    st.markdown("---")
    st.subheader("📊 Project Risk Dashboard & Brief")
    
    # Dashboard Views
    tab1, tab2, tab3, tab4 = st.tabs([
        "Overview", 
        "Risk Register", 
        "Action Plan", 
        "Governance Flags"
    ])
    
    with tab1:
        st.markdown(res.get("executive_summary", "No summary generated."))
        
    with tab2:
        st.markdown("##### Identified Risk Events & Evidence")
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
            
    with tab3:
        st.markdown("##### Recommended Actionable Mitigation Steps")
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
            
    with tab4:
        st.markdown("##### Human-in-the-Loop & Governance Validation")
        st.warning("⚠️ **Governance Notice:** The items below require explicit stakeholder sign-off before automated tasks are dispatched.")
        flags = res.get("governance_flags", [])
        for f_item in flags:
            st.markdown(f"- {f_item}")
            
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Export Section
    st.subheader("📥 Export & Share Deliverables")
    exp_col1, exp_col2, exp_col3 = st.columns(3)
    
    with exp_col1:
        pdf_bytes = generate_pdf_report(
            res.get("executive_summary", ""),
            res.get("risk_matrix", []),
            res.get("mitigation_plan", []),
            res.get("governance_flags", [])
        )
        st.download_button(
            label="📄 Download Executive PDF Report",
            data=pdf_bytes,
            file_name="Executive_Risk_Intelligence_Report.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
    with exp_col2:
        json_data = json.dumps(res, indent=2)
        st.download_button(
            label="💾 Export Raw JSON Analysis",
            data=json_data,
            file_name="Risk_Analysis_Data.json",
            mime="application/json",
            use_container_width=True
        )
        
    with exp_col3:
        md_text = f"{res.get('executive_summary')}\n\n### Risk Table\n" + pd.DataFrame(res.get('risk_matrix', [])).to_markdown(index=False)
        st.download_button(
            label="📋 Download Markdown Summary",
            data=md_text,
            file_name="Risk_Summary.md",
            mime="text/markdown",
            use_container_width=True
        )
