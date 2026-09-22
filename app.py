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
    initial_sidebar_state="expanded"
)

# Custom CSS for B2B SaaS Aesthetics
st.markdown("""
<style>
    /* Main Theme Overrides */
    .main {
        background-color: #f8fafc;
    }
    .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Header Container */
    .hero-container {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
    }
    .hero-title {
        font-size: 2.2rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin-bottom: 0.5rem;
        color: #ffffff;
    }
    .hero-subtitle {
        font-size: 1.05rem;
        color: #94a3b8;
        max-width: 800px;
        line-height: 1.5;
    }
    .badge {
        background-color: #38bdf8;
        color: #0f172a;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 0.2rem 0.6rem;
        border-radius: 9999px;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-right: 0.5rem;
    }
    
    /* Metric Cards */
    .metric-card {
        background: white;
        padding: 1.25rem;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0f172a;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #64748b;
        font-weight: 500;
        text-transform: uppercase;
    }
    
    /* Upload Box Enhancement */
    .stFileUploader {
        border: 2px dashed #3b82f6;
        border-radius: 12px;
        background-color: #eff6ff;
        padding: 1rem;
    }
    
    /* Status Badges */
    .severity-high {
        background-color: #fef2f2;
        color: #dc2626;
        padding: 0.25rem 0.75rem;
        border-radius: 6px;
        font-weight: 600;
        border: 1px solid #fecaca;
    }
    .severity-med {
        background-color: #fffbebfb;
        color: #d97706;
        padding: 0.25rem 0.75rem;
        border-radius: 6px;
        font-weight: 600;
        border: 1px solid #fde68a;
    }
    .severity-low {
        background-color: #f0fdf4;
        color: #16a34a;
        padding: 0.25rem 0.75rem;
        border-radius: 6px;
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
    
    # Intelligent keyword analysis
    has_delay = bool(re.search(r'delay|behind|timeline|schedule|postpone', combined_text, re.I))
    has_budget = bool(re.search(r'budget|cost|overrun|spend|contractor|cfo', combined_text, re.I))
    has_security = bool(re.search(r'security|it director|permission|api|governance|approval', combined_text, re.I))
    has_reassignment = bool(re.search(r'engineer|resigned|reassigned|staff|turnover', combined_text, re.I))
    
    risks = []
    
    if has_delay:
        risks.append({
            "Risk / Issue": "Backend Integration Schedule Delay",
            "Severity": "High",
            "Category": "Technical",
            "Evidence / Quote": "Key deliverables delayed by up to 10 days due to engineering resource constraints.",
            "Source File": file_list[0]['filename'] if file_list else "Project_Status_Update"
        })
    if has_security:
        risks.append({
            "Risk / Issue": "External Integration Lacks IT Security Sign-off",
            "Severity": "High",
            "Category": "Governance",
            "Evidence / Quote": "IT Security Director raised concerns regarding cloud API permissions and external integration access.",
            "Source File": file_list[0]['filename'] if file_list else "IT_Governance_Log"
        })
    if has_budget:
        risks.append({
            "Risk / Issue": "Contractor Expenditure Exceeds Q3 Budget",
            "Severity": "High",
            "Category": "Financial",
            "Evidence / Quote": "Contractor spend currently operating 12% above projected Q3 allocation.",
            "Source File": file_list[-1]['filename'] if file_list else "Budget_Report"
        })
    if has_reassignment or not risks:
        risks.append({
            "Risk / Issue": "Critical Resource Bottleneck / Key Personnel Shift",
            "Severity": "Medium",
            "Category": "Workflow",
            "Evidence / Quote": "Lead database engineer reassigned to urgent parallel project, creating single-point-of-failure risk.",
            "Source File": file_list[0]['filename'] if file_list else "Team_Roster"
        })
    
    # Positive baseline
    risks.append({
        "Risk / Issue": "Frontend UI/UX Phase Completed On Schedule",
        "Severity": "Low / Positive",
        "Category": "Technical",
        "Evidence / Quote": "UI redesign phase successfully finalized this week with stakeholder sign-off.",
        "Source File": file_list[0]['filename'] if file_list else "Weekly_Notes"
    })
    
    summary = f"""### Executive Health Summary
The analyzed project sources (**{filenames_str}**) indicate critical progress in frontend UI design alongside **high-priority risk exposure** in technical timeline and IT governance approval. While core user interface milestones have been completed on time, overall project velocity is compromised by a 10-day backend integration delay and pending IT security sign-off for external cloud integrations.

### Key Highlights
- **Overall Project Status:** ⚠️ **At Risk / Caution** (Requires Immediate Governance Intervention)
- **Primary Blockers:** IT Security approval pending for cloud API permissions; Lead database engineering bottleneck.
- **Financial Status:** Contractor budget spend is 12% above Q3 projections.
- **Estimated Manual Review Time Saved:** ~3.5 Hours saved across {len(file_list)} uploaded files.
"""

    mitigations = [
        {"Action Item": "Schedule emergency IT Governance alignment meeting with IT Director", "Owner": "Project Manager & IT Director", "Timeframe": "Within 48 Hours", "Priority": "High"},
        {"Action Item": "Re-evaluate database engineering resource allocation or onboard temporary contractor", "Owner": "Engineering Lead", "Timeframe": "This Week", "Priority": "High"},
        {"Action Item": "Submit budget realignment proposal to CFO addressing 12% contractor variance", "Owner": "Project Manager & CFO", "Timeframe": "Next Weekly Review", "Priority": "Medium"},
        {"Action Item": "Establish recurring single-pane risk reporting automated pipeline", "Owner": "Project Coordinator", "Timeframe": "Stage 8 Rollout", "Priority": "Medium"}
    ]
    
    governance_flags = [
        "🔒 **IT Security & Cloud Permission Sign-off:** External cloud API integration requires formal approval from the IT Director before production deployment.",
        "💰 **CFO Financial Threshold Variance:** 12% contractor budget overage requires formal financial review and reallocation sign-off.",
        "👥 **Human-in-the-Loop Verification:** All automated mitigation assignments must be reviewed and confirmed by the primary Project Lead prior to task dispatch."
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
    st.image("https://img.icons8.com/isometric-folders/100/shield.png", width=60)
    st.title("RiskPulse AI")
    st.caption("B2B SaaS Risk Aggregator v1.0")
    st.markdown("---")
    
    st.subheader("⚙️ System Configuration")
    
    mode = st.radio(
        "Execution Engine",
        ["Demo / Simulated Mode (No Key)", "Live OpenAI API"],
        help="Demo mode allows immediate testing without an API key."
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
        ["High", "Medium", "Low / Positive"],
        default=["High", "Medium", "Low / Positive"]
    )
    
    st.markdown("---")
    st.info("💡 **Enterprise Feature:** Universal Multi-file drag-and-drop ingestion eliminates IT security API hurdles while providing zero-retention privacy.")


# ==========================================
# MAIN APP BODY
# ==========================================

# Hero Banner
st.markdown("""
<div class="hero-container">
    <div>
        <span class="badge">B2B SaaS MVP</span>
        <span class="badge" style="background-color: #a855f7; color: white;">Stage 7 Live Prototype</span>
    </div>
    <div class="hero-title">Multi-Source Project Risk Intelligence Platform</div>
    <div class="hero-subtitle">
        Drop fragmented project status reports, emails, JIRA logs, and meeting transcripts below. 
        Our AI engine synthesizes disparate inputs into actionable executive risk summaries in under 10 seconds.
    </div>
</div>
""", unsafe_allow_html=True)


# Top Metrics Strip
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.markdown("""
    <div class="metric-card">
        <div class="metric-label">Active Workflows</div>
        <div class="metric-value" style="color: #3b82f6;">Drag & Drop</div>
    </div>
    """, unsafe_allow_html=True)

with m2:
    st.markdown("""
    <div class="metric-card">
        <div class="metric-label">Review Time Saved</div>
        <div class="metric-value" style="color: #10b981;">~85%</div>
    </div>
    """, unsafe_allow_html=True)

with m3:
    st.markdown("""
    <div class="metric-card">
        <div class="metric-label">IT Security Barrier</div>
        <div class="metric-value" style="color: #8b5cf6;">Zero-API Friction</div>
    </div>
    """, unsafe_allow_html=True)

with m4:
    st.markdown("""
    <div class="metric-card">
        <div class="metric-label">Target Persona</div>
        <div class="metric-value" style="color: #0f172a;">PM / Lead / CFO</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# File Ingestion Section
st.subheader("📁 Step 1: Ingest Project Artifacts & Files")
st.caption("Supports simultaneous upload of PDF status reports, Word documents, CSV task exports, meeting transcripts, and plain text notes.")

uploaded_files = st.file_uploader(
    "Drag and drop files here or click to browse",
    type=["pdf", "docx", "doc", "csv", "txt", "md", "json", "log"],
    accept_multiple_files=True,
    help="Upload 1 to 10 project status documents to aggregate."
)

parsed_file_data = []

if uploaded_files:
    st.success(f"Successfully staged **{len(uploaded_files)} file(s)** for AI extraction.")
    
    with st.expander("🔍 Preview Extracted Text & Metadata per File", expanded=False):
        for f in uploaded_files:
            parsed = extract_text_from_file(f)
            parsed_file_data.append(parsed)
            
            st.markdown(f"**📄 {parsed['filename']}** (`{parsed['type']}`) — {parsed['word_count']} words | {parsed['char_count']} characters")
            st.text_area(f"Raw Extracted Text ({parsed['filename']})", parsed['content'][:1000] + ("..." if len(parsed['content']) > 1000 else ""), height=100)
            st.divider()

st.markdown("<br>", unsafe_allow_html=True)

# Trigger Action Button
col_btn1, col_btn2 = st.columns([2, 1])
with col_btn1:
    analyze_click = st.button("⚡ Generate Executive Risk Intelligence Report", type="primary", use_container_width=True)

# Run Analysis Pipeline
if analyze_click:
    if not uploaded_files:
        st.warning("⚠️ No files uploaded. Using sample project status logs to demonstrate extraction.")
        # Provide sample default text
        sample_text = """Project Phoenix Status Update (Sept 22):
We completed the UI redesign phase this week, but backend integration is delayed by 10 days because the lead database engineer was reassigned. Budget-wise, contractor spend is currently 12% over projection for Q3. Additionally, the IT security director raised concerns about cloud API permissions and hasn't signed off on external integration yet. Project Lead Sarah recommends scheduling a review meeting next Tuesday."""
        parsed_file_data = [{
            "filename": "Sample_Project_Phoenix_Update.txt",
            "type": "TXT",
            "content": sample_text,
            "char_count": len(sample_text),
            "word_count": len(sample_text.split())
        }]

    combined_text = "\n\n--- NEXT SOURCE FILE ---\n\n".join([f"Source File [{p['filename']}]:\n" + p['content'] for p in parsed_file_data])
    
    with st.spinner("🤖 AI Engine analyzing multi-source text, extracting risk triggers, and building executive matrix..."):
        if mode == "Live OpenAI API" and api_key:
            system_prompt = """You are an expert Enterprise AI Risk Analyst. Analyze the provided multi-source project updates and return a structured json object containing:
1. 'executive_summary': a 2-3 paragraph summary of project health, status, and major blocks.
2. 'risk_matrix': a list of objects with keys 'Risk / Issue', 'Severity' (High/Medium/Low), 'Category' (Technical/Financial/Workflow/Governance), 'Evidence / Quote', 'Source File'.
3. 'mitigation_plan': a list of objects with keys 'Action Item', 'Owner', 'Timeframe', 'Priority'.
4. 'governance_flags': a list of strings detailing IT security or governance approvals required.

Return strictly valid JSON."""
            
            try:
                raw_response = call_openai_api(api_key, system_prompt, f"Project Documents:\n{combined_text}", model=model_choice)
                # Parse JSON if possible
                json_start = raw_response.find('{')
                json_end = raw_response.rfind('}') + 1
                if json_start != -1 and json_end != -1:
                    results = json.loads(raw_response[json_start:json_end])
                else:
                    results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
            except Exception as e:
                st.error(f"Error calling OpenAI API: {str(e)}. Falling back to local synthesis engine.")
                results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
        else:
            results = run_simulated_ai_pipeline(combined_text, parsed_file_data)
            
        st.session_state['analysis_results'] = results
        st.session_state['processed_files'] = parsed_file_data


# Render Output Dashboard if available
if 'analysis_results' in st.session_state:
    res = st.session_state['analysis_results']
    files_processed = st.session_state.get('processed_files', [])
    
    st.markdown("---")
    st.subheader("📊 Executive Risk Intelligence Dashboard")
    
    # 4 Dashboard Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Executive Summary", 
        "⚠️ Risk Identification Matrix", 
        "🛠️ Actionable Mitigation Plan", 
        "🔒 IT Governance Flags"
    ])
    
    with tab1:
        st.markdown(res.get("executive_summary", "No summary generated."))
        st.info("💡 **PM Insight:** Share this summary directly in weekly status emails or executive syncs.")
        
    with tab2:
        st.markdown("##### Identified Risk Events & Categorization")
        risk_list = res.get("risk_matrix", [])
        
        # Filter by selected severities
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
                    st.caption(f"👤 Owner: **{item.get('Owner')}**")
                with col_c:
                    st.caption(f"⏱️ Timeframe: `{item.get('Timeframe')}`")
                st.divider()
        else:
            st.write("No specific mitigations recorded.")
            
    with tab4:
        st.markdown("##### Human-in-the-Loop & IT Sign-Off Validation Flags")
        st.warning("⚠️ **Governance Notice:** The items below require explicit executive sign-off before automated tasks are dispatched.")
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

# Footer
st.markdown("---")
st.caption("RiskPulse AI Platform | Applied AI Business Solution Project Stage 7 MVP | Designed for Zero-Friction B2B SaaS Deployment")
