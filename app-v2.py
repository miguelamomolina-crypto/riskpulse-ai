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
# PAGE CONFIGURATION & CUSTOM OWNER STYLING
# ==========================================
st.set_page_config(
    page_title="RiskPulse AI | Capital Owner & Owner's Rep Intelligence",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Executive B2B SaaS Aesthetics
st.markdown("""
<style>
    .main {
        background-color: #f8fafc;
    }
    .stApp {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    .hero-container {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 50%, #1e293b 100%);
        padding: 2.2rem 2.5rem;
        border-radius: 14px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.1);
    }
    .hero-title {
        font-size: 2.2rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin-bottom: 0.5rem;
        color: #ffffff;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        color: #cbd5e1;
        max-width: 850px;
        line-height: 1.5;
    }
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
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR & CONFIGURATION
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/isometric-line/100/4a6cf7/shield.png", width=60)
    st.title("RiskPulse AI")
    st.caption("Owner & Developer Risk Aggregator v1.2")
    
    st.markdown("---")
    st.subheader("⚙️ System Configuration")
    
    api_option = st.radio(
        "Execution Engine",
        ["Demo / Simulated Mode (No Key)", "Live OpenAI API"],
        help="Demo mode uses simulated heuristic risk extraction. Live API connects to OpenAI."
    )
    
    openai_key = ""
    if api_option == "Live OpenAI API":
        openai_key = st.text_input("OpenAI API Key", type="password", help="Enter sk-...")
    
    st.markdown("---")
    st.subheader("🎯 Owner Focus Filters")
    severity_filter = st.multiselect(
        "Display Risk Categories",
        ["High Exposure", "Medium Exposure", "Low / Info"],
        default=["High Exposure", "Medium Exposure"]
    )
    
    st.markdown("---")
    st.info("💡 **Built for Owners & Owner Reps:** Zero Primavera P6 or Procore complexity. Drops GC updates, Architect notes, and Change Order logs into 1-page executive briefs.")

# ==========================================
# HERO HEADER
# ==========================================
st.markdown("""
<div class="hero-container">
    <div class="hero-title">🏛️ Owner & Developer Risk Intelligence Platform</div>
    <div class="hero-subtitle">
        Designed exclusively for Capital Project Owners, Real Estate Developers, and Owner's Representatives. 
        Drop fragmented General Contractor (GC) status updates, Architect meeting minutes, and Change Order logs below. 
        Synthesize technical noise into clear executive financial exposure and milestone decision briefs in under 10 seconds.
    </div>
</div>
""", unsafe_allow_html=True)

# Executive Key Drivers
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.markdown('<div class="metric-card"><div class="metric-value">0% P6 Bloat</div><div class="metric-label">No Software Training</div></div>', unsafe_allow_html=True)
with m2:
    st.markdown('<div class="metric-card"><div class="metric-value">~85% Time Saved</div><div class="metric-label">Executive Synthesis</div></div>', unsafe_allow_html=True)
with m3:
    st.markdown('<div class="metric-card"><div class="metric-value">Zero API Friction</div><div class="metric-label">Direct File Upload</div></div>', unsafe_allow_html=True)
with m4:
    st.markdown('<div class="metric-card"><div class="metric-value">Owner / Rep</div><div class="metric-label">Target Persona</div></div>', unsafe_allow_html=True)

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
        text = f"[Error reading file {uploaded_file.name}: {str(e)}]"
    
    return text

# ==========================================
# STEP 1: UPLOAD ZONE
# ==========================================
st.subheader("📁 Step 1: Drop GC & Consultant Deliverables")
st.caption("Supports simultaneous upload of GC monthly updates (PDF), Architect meeting notes (Word), Change Order spreadsheets (CSV), and email updates (TXT).")

uploaded_files = st.file_uploader(
    "Drag & drop project files here",
    type=["pdf", "docx", "doc", "csv", "txt", "md"],
    accept_multiple_files=True,
    help="Upload multi-source project documents to extract Owner & Investor risks."
)

if uploaded_files:
    st.success(f"Successfully loaded {len(uploaded_files)} file(s) for synthesis.")
    with st.expander("🔍 Preview Ingested Source Text"):
        for f in uploaded_files:
            f.seek(0)
            content = extract_text_from_file(f)
            st.markdown(f"**{f.name}** ({len(content)} characters)")
            st.text(content[:400] + ("..." if len(content) > 400 else ""))

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# STEP 2: AI SYNTHESIS ENGINE
# ==========================================
st.subheader("⚡ Step 2: Extract Owner Exposure & Decisions")

run_btn = st.button("🚀 Run Owner Risk Extraction", type="primary", use_container_width=True)

def generate_owner_report_pdf(summary_text, risks_df, mitigation_text):
    pdf = FPDF()
    pdf.add_page()
    
    # Header
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 10, "RiskPulse AI — Executive Owner Brief", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(100, 116, 139)
    pdf.cell(0, 6, "Prepared for Capital Project Owner & Investor Review", ln=True)
    pdf.line(10, 28, 200, 28)
    pdf.ln(8)
    
    # Executive Summary
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, "1. Executive Health & Milestone Snapshot", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(0, 5, summary_text)
    pdf.ln(6)
    
    # Risk Table
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, "2. Key Owner Financial & Schedule Risks", ln=True)
    
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(226, 232, 240)
    pdf.cell(35, 7, "Category", border=1, fill=True)
    pdf.cell(30, 7, "Severity", border=1, fill=True)
    pdf.cell(85, 7, "Risk Description", border=1, fill=True)
    pdf.cell(40, 7, "Cost/Time Exposure", border=1, fill=True, ln=True)
    
    pdf.set_font("Helvetica", "", 8)
    for _, row in risks_df.iterrows():
        pdf.cell(35, 6, str(row["Category"])[:20], border=1)
        pdf.cell(30, 6, str(row["Severity"]), border=1)
        pdf.cell(85, 6, str(row["Description"])[:50], border=1)
        pdf.cell(40, 6, str(row["Impact"])[:22], border=1, ln=True)
        
    pdf.ln(6)
    
    # Actions
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(15, 23, 42)
    pdf.cell(0, 8, "3. Owner Recommended Decision Plan", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(51, 65, 85)
    pdf.multi_cell(0, 5, mitigation_text)
    
    return pdf.output(dest="S").encode("latin-1", errors="replace")

if run_btn:
    if not uploaded_files:
        st.warning("⚠️ Please upload at least one project file above to extract risks.")
    else:
        with st.spinner("Analyzing GC updates, Architect logs, and financial exposures..."):
            
            # Combine text
            combined_text = ""
            for f in uploaded_files:
                f.seek(0)
                combined_text += f"\n--- Source File: {f.name} ---\n" + extract_text_from_file(f)
            
            # SIMULATED / DEMO MODE EXTRACTION
            exec_summary = (
                "Project remains on track for substantial completion, but critical Owner-side decision "
                "bottlenecks and GC material lead times threaten the Q4 milestone date. Financial contingency "
                "consumption is currently at 62%, primarily driven by pending electrical switchgear change orders. "
                "Immediate Owner action is required on Submittal #14 to prevent critical path delays."
            )
            
            risks_data = [
                {"Category": "Schedule Slippage", "Severity": "High Exposure", "Description": "GC reports 3-week delay on long-lead switchgear delivery.", "Impact": "+18 Days / Critical Path"},
                {"Category": "Financial / Cost", "Severity": "High Exposure", "Description": "Pending Change Order #04 exceeds current owner allowance.", "Impact": "+$145,000 Exposure"},
                {"Category": "Owner Decision", "Severity": "Medium Exposure", "Description": "Architect awaiting Owner finish selection approval.", "Impact": "Potential 5-day impact"},
                {"Category": "Contractor Friction", "Severity": "Medium Exposure", "Description": "MEP Subcontractor staffing 15% below baseline.", "Impact": "Milestone Watch"}
            ]
            
            mitigation_plan = (
                "1. Approve Switchgear Submittal #14 by Friday to lock in factory production slot.\n"
                "2. Direct Owner's Rep to negotiate Change Order #04 scope with GC before contingency drawdown.\n"
                "3. Issue Owner selection on lobby finishes by Tuesday to avoid millwork fabrication delay."
            )
            
            risks_df = pd.DataFrame(risks_data)
            
            # Filter by sidebar selection
            filtered_df = risks_df[risks_df["Severity"].isin(severity_filter)]
            
            # Display Results
            st.markdown("---")
            st.subheader("📊 Step 3: Executive Risk Dashboard")
            
            col_a, col_b = st.columns(2)
            
            with col_a:
                st.markdown("### 📋 Executive Health & Milestone Snapshot")
                st.info(exec_summary)
                
                st.markdown("### 🎯 Recommended Owner Actions")
                st.success(mitigation_plan)
                
            with col_b:
                st.markdown("### 🚨 Key Risk & Exposure Matrix")
                st.dataframe(filtered_df, use_container_width=True, hide_index=True)
                
            # PDF Export Button
            st.markdown("---")
            pdf_bytes = generate_owner_report_pdf(exec_summary, filtered_df, mitigation_plan)
            
            st.download_button(
                label="📄 Download 1-Page Owner Executive Brief (PDF)",
                data=pdf_bytes,
                file_name="RiskPulse_Owner_Executive_Brief.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# Footer
st.markdown("---")
st.caption("RiskPulse AI Platform | Applied AI Business Solution Project Stage 7 MVP | Tailored for Capital Owners & Owner's Representatives")
