import streamlit as st
import pandas as pd
from io import BytesIO
import docx
import pypdf
from fpdf import FPDF

# ==========================================
# PAGE CONFIGURATION & APPLE-INSPIRED STYLING
# ==========================================
st.set_page_config(
    page_title="RiskPulse AI | Owner Risk Intelligence",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Minimalist Apple-like Light Aesthetic
st.markdown("""
<style>
    /* Global Clean Light Theme */
    .stApp {
        background-color: #fbfbfd;
        font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Helvetica Neue", Arial, sans-serif;
        color: #1d1d1f;
    }
    
    /* Header */
    .header-box {
        padding: 1.2rem 0 1rem 0;
        margin-bottom: 1.5rem;
        border-bottom: 1px solid #e5e5ea;
    }
    .header-title {
        font-size: 2.1rem;
        font-weight: 700;
        letter-spacing: -0.025em;
        color: #1d1d1f;
    }
    .header-sub {
        font-size: 1.05rem;
        color: #86868b;
        font-weight: 400;
        margin-top: 0.2rem;
    }

    /* Metric Badges */
    .stat-badge {
        background: #ffffff;
        border: 1px solid #e5e5ea;
        border-radius: 12px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02);
    }
    .stat-num {
        font-size: 1.4rem;
        font-weight: 600;
        color: #0066cc;
    }
    .stat-title {
        font-size: 0.78rem;
        color: #86868b;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-top: 0.2rem;
    }

    /* Section Cards */
    .section-title {
        font-size: 1.15rem;
        font-weight: 600;
        color: #1d1d1f;
        margin-bottom: 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR (MINIMALIST)
# ==========================================
with st.sidebar:
    st.title("RiskPulse AI")
    st.caption("Owner & Developer Intelligence")
    
    st.markdown("---")
    api_option = st.radio(
        "Mode",
        ["Demo / Simulated Mode", "Live OpenAI API"],
        help="Demo mode simulates AI risk extraction from uploaded documents."
    )
    
    openai_key = ""
    if api_option == "Live OpenAI API":
        openai_key = st.text_input("OpenAI API Key", type="password", help="Enter sk-...")
    
    st.markdown("---")
    severity_filter = st.multiselect(
        "Risk Severity",
        ["High Exposure", "Medium Exposure", "Low / Info"],
        default=["High Exposure", "Medium Exposure"]
    )
    
    st.markdown("---")
    st.caption("Designed for Capital Owners, Developers & Owner Reps. Zero P6 software complexity.")

# ==========================================
# CLEAN APPLE-STYLE HEADER
# ==========================================
st.markdown("""
<div class="header-box">
    <div class="header-title">RiskPulse AI</div>
    <div class="header-sub">Clear executive financial exposure and milestone intelligence for Capital Owners & Representatives.</div>
</div>
""", unsafe_allow_html=True)

# Minimal KPI Bar
c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('<div class="stat-badge"><div class="stat-num">Zero P6 Bloat</div><div class="stat-title">No Training Required</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="stat-badge"><div class="stat-num">~85% Time Saved</div><div class="stat-title">Instant Executive Synthesis</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="stat-badge"><div class="stat-num">1-Page Output</div><div class="stat-title">Investor & Board Ready</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# FILE INGESTION UTILITIES
# ==========================================
def extract_text_from_file(uploaded_file):
    text = ""
    file_type = uploaded_file.name.split('.')[-1].lower()
    
    try:
        if file_type == 'pdf':
            pdf_reader = pypdf.PdfReader(BytesIO(uploaded_file.read()))
            for page in pdf_reader.pages:
                text += page.extract_text() or ""
        elif file_type in ['docx', 'doc']:
            doc = docx.Document(BytesIO(uploaded_file.read()))
            for p in doc.paragraphs:
                text += p.text + "\n"
        elif file_type == 'csv':
            df = pd.read_csv(uploaded_file)
            text = df.to_string()
        elif file_type in ['txt', 'md', 'json', 'log']:
            text = uploaded_file.read().decode('utf-8', errors='ignore')
    except Exception as e:
        text = f"[Error reading {uploaded_file.name}: {str(e)}]"
    
    return text

# ==========================================
# UPLOAD ZONE (CLEAR OWNER-SPECIFIC PROMPT)
# ==========================================
st.markdown('<div class="section-title">Upload Project Files</div>', unsafe_allow_html=True)
st.caption("Drop your latest project updates below — Meeting Minutes, Schedules, RFIs, Submittals, ASRs, PRs, or GC Reports.")

uploaded_files = st.file_uploader(
    "Drag and drop files here",
    type=["pdf", "docx", "doc", "csv", "txt", "md"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

if uploaded_files:
    st.success(f"Loaded {len(uploaded_files)} file(s). Ready for analysis.")
    with st.expander("Preview Ingested Source Text"):
        for f in uploaded_files:
            f.seek(0)
            content = extract_text_from_file(f)
            st.markdown(f"**{f.name}** ({len(content)} characters)")
            st.text(content[:300] + ("..." if len(content) > 300 else ""))

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# AI SYNTHESIS ENGINE & PDF GENERATOR
# ==========================================
run_btn = st.button("Synthesize Executive Brief", type="primary", use_container_width=True)

def clean_pdf_text(text):
    if not isinstance(text, str):
        text = str(text)
    replacements = {
        '—': '-', '–': '-', '“': '"', '”': '"', 
        '‘': "'", '’': "'", '…': '...', '•': '*'
    }
    for orig, repl in replacements.items():
        text = text.replace(orig, repl)
    return text.encode('latin-1', 'ignore').decode('latin-1')

def generate_owner_report_pdf(summary_text, risks_df, mitigation_text):
    pdf = FPDF()
    pdf.add_page()
    
    # Header
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_text_color(29, 29, 31)
    pdf.cell(0, 10, clean_pdf_text("RiskPulse AI - Executive Owner Brief"))
    pdf.ln(8)
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(134, 134, 139)
    pdf.cell(0, 6, clean_pdf_text("Capital Project Executive Risk Summary"))
    pdf.ln(6)
    pdf.line(10, 26, 200, 26)
    pdf.ln(8)
    
    # Executive Health
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(29, 29, 31)
    pdf.cell(0, 8, clean_pdf_text("1. Executive Health & Milestone Snapshot"))
    pdf.ln(7)
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(0, 5, clean_pdf_text(summary_text))
    pdf.ln(6)
    
    # Risks
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(29, 29, 31)
    pdf.cell(0, 8, clean_pdf_text("2. Key Financial & Milestone Risks"))
    pdf.ln(7)
    
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(245, 245, 247)
    pdf.cell(35, 7, "Category", border=1, fill=True)
    pdf.cell(30, 7, "Severity", border=1, fill=True)
    pdf.cell(85, 7, "Risk Description", border=1, fill=True)
    pdf.cell(40, 7, "Cost/Time Exposure", border=1, fill=True)
    pdf.ln(7)
    
    pdf.set_font("Helvetica", "", 8)
    for _, row in risks_df.iterrows():
        pdf.cell(35, 6, clean_pdf_text(str(row["Category"])[:20]), border=1)
        pdf.cell(30, 6, clean_pdf_text(str(row["Severity"])), border=1)
        pdf.cell(85, 6, clean_pdf_text(str(row["Description"])[:50]), border=1)
        pdf.cell(40, 6, clean_pdf_text(str(row["Impact"])[:22]), border=1)
        pdf.ln(6)
        
    pdf.ln(6)
    
    # Actions
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(29, 29, 31)
    pdf.cell(0, 8, clean_pdf_text("3. Recommended Owner Action Items"))
    pdf.ln(7)
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(0, 5, clean_pdf_text(mitigation_text))
    
    return bytes(pdf.output())

if run_btn:
    if not uploaded_files:
        st.warning("Please upload project files above before running the synthesis.")
    else:
        with st.spinner("Analyzing project updates..."):
            
            # Combine text
            combined_text = ""
            for f in uploaded_files:
                f.seek(0)
                combined_text += f"\n--- File: {f.name} ---\n" + extract_text_from_file(f)
            
            # Simulated / Demo Synthesis
            exec_summary = (
                "Project remains on track for substantial completion, but long-lead switchgear submittal (SUB-014) "
                "and foundation change orders (PCO-004) create a cumulative 35-day float loss on the critical path. "
                "Contingency drawdown stands at 62%. Immediate Owner approval is required on SUB-014 to secure production."
            )
            
            risks_data = [
                {"Category": "Schedule Slippage", "Severity": "High Exposure", "Description": "Main Switchgear Submittal #14 pending owner execution.", "Impact": "+28 Days / Critical Path"},
                {"Category": "Financial / Cost", "Severity": "High Exposure", "Description": "PCO-004 Foundation Micropiles exceeds allowance.", "Impact": "+$145,000 Exposure"},
                {"Category": "Additional Services", "Severity": "Medium Exposure", "Description": "Architect ASR-003 canopy recalculations submitted.", "Impact": "+$18,200 Fee Request"},
                {"Category": "Owner Selection", "Severity": "Medium Exposure", "Description": "Lobby finish selection awaiting Owner sign-off.", "Impact": "Millwork Hold"}
            ]
            
            mitigation_plan = (
                "• Execute Switchgear Submittal #14 by Friday to prevent factory queue loss.\n"
                "• Direct Owner's Rep to negotiate PCO-004 micropile scope prior to contingency drawdown.\n"
                "• Approve ASR-003 canopy engineering fee ($18,200) and finalize lobby finish selection."
            )
            
            risks_df = pd.DataFrame(risks_data)
            filtered_df = risks_df[risks_df["Severity"].isin(severity_filter)]
            
            # Clean Output Display
            st.markdown("---")
            st.markdown('<div class="section-title">Executive Risk Brief</div>', unsafe_allow_html=True)
            
            col_a, col_b = st.columns(2)
            
            with col_a:
                st.markdown("##### 📋 Health & Milestone Snapshot")
                st.info(exec_summary)
                
                st.markdown("##### 🎯 Recommended Owner Actions")
                st.success(mitigation_plan)
                
            with col_b:
                st.markdown("##### 🚨 Key Financial & Schedule Exposure")
                st.dataframe(filtered_df, use_container_width=True, hide_index=True)
                
            st.markdown("---")
            pdf_bytes = generate_owner_report_pdf(exec_summary, filtered_df, mitigation_plan)
            
            st.download_button(
                label="Download Executive Brief (PDF)",
                data=pdf_bytes,
                file_name="RiskPulse_Executive_Brief.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# Footer
st.markdown("---")
st.caption("RiskPulse AI Platform | Applied AI Business Solution Project | Built for Capital Project Owners & Representatives")
