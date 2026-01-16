import streamlit as st
from datetime import date
import io
import os
import re
import pandas as pd
from PIL import Image

# -----------------------------
# FILE & ASSETS SETUP
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "diagrams")

AWS_PN_LOGO = os.path.join(ASSETS_DIR, "aws partner logo.jpg")
ONETURE_LOGO = os.path.join(ASSETS_DIR, "oneture logo1.jpg")
AWS_ADV_LOGO = os.path.join(ASSETS_DIR, "aws advanced logo1.jpg")

SOW_DIAGRAM_MAP = {
    "L1 Support Bot POC SOW": os.path.join(ASSETS_DIR, "L1 Support Bot POC SOW.png"),
    "Ready Search POC Scope of Work Document": os.path.join(ASSETS_DIR, "Ready Search POC Scope of Work Document.png"),
    "AI based Image Enhancement POC SOW": os.path.join(ASSETS_DIR, "AI based Image Enhancement POC SOW.png"),
    "Beauty Advisor POC SOW": os.path.join(ASSETS_DIR, "Beauty Advisor POC SOW.png"),
    "AI based Image Inspection POC SOW": os.path.join(ASSETS_DIR, "AI based Image Inspection POC SOW.png"),
    "Gen AI for SOP POC SOW": os.path.join(ASSETS_DIR, "Gen AI for SOP POC SOW.png"),
    "Project Scope Document": os.path.join(ASSETS_DIR, "Project Scope Document.png"),
    "Gen AI Speech To Speech": os.path.join(ASSETS_DIR, "Gen AI Speech To Speech.png"),
    "PoC Scope Document": os.path.join(ASSETS_DIR, "PoC Scope Document.png")
}

SOW_COST_TABLE_MAP = {
    "L1 Support Bot POC SOW": {"poc_cost": "3,536.40 USD"},
    "Beauty Advisor POC SOW": {
        "poc_cost": "4,525.66 USD + 200 USD (Amazon Bedrock Cost) = 4,725.66",
        "prod_cost": "4,525.66 USD + 1,175.82 USD (Amazon Bedrock Cost) = 5,701.48"
    },
    "Ready Search POC Scope of Work Document": {"poc_cost": "2,641.40 USD"},
    "AI based Image Enhancement POC SOW": {"poc_cost": "2,814.34 USD"},
    "AI based Image Inspection POC SOW": {"poc_cost": "3,536.40 USD"},
    "Gen AI for SOP POC SOW": {"poc_cost": "2,110.30 USD"},
    "Project Scope Document": {"prod_cost": "2,993.60 USD"},
    "Gen AI Speech To Speech": {"prod_cost": "2,124.23 USD"},
    "PoC Scope Document": {"amazon_bedrock": "1,000 USD", "total": "$ 3,150"},
}

# -----------------------------
# STREAMLIT CONFIG
# -----------------------------
st.set_page_config(page_title="GenAI SOW Architect", layout="wide", page_icon="📄", initial_sidebar_state="expanded")

# Custom CSS
st.markdown("""
<style>
.block-container { padding-top: 1.5rem; }
.stTextArea textarea { border-radius: 10px; }
.stButton>button { border-radius: 8px; font-weight: 600; }
.stDataFrame { width: 100%; }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# SESSION STATE
# -----------------------------
if 'generated_sow' not in st.session_state:
    st.session_state.generated_sow = ""

if 'stakeholders' not in st.session_state:
    st.session_state.stakeholders = {
        "Partner": pd.DataFrame([{"Name": "Gaurav Kankaria", "Title": "Head of Analytics & ML", "Email": "gaurav.kankaria@oneture.com"}]),
        "Customer": pd.DataFrame([{"Name": "Cheten Dev", "Title": "Head of Product Design", "Email": "cheten.dev@nykaa.com"}]),
        "AWS": pd.DataFrame([{"Name": "Anubhav Sood", "Title": "AWS Account Executive", "Email": "anbhsood@amazon.com"}]),
        "Escalation": pd.DataFrame([
            {"Name": "Omkar Dhavalikar", "Title": "AI/ML Lead", "Email": "omkar.dhavalikar@oneture.com"},
            {"Name": "Gaurav Kankaria", "Title": "Head of Analytics and AIML", "Email": "gaurav.kankaria@oneture.com"}
        ])
    }

def clear_sow():
    st.session_state.generated_sow = ""

# -----------------------------
# SIDEBAR: PROJECT INTAKE
# -----------------------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/artificial-intelligence.png", width=60)
    st.title("SOW Architect")
    st.caption("Enterprise POC/MVP Engine")
    
    sow_type_options = list(SOW_DIAGRAM_MAP.keys())
    selected_sow_name = st.selectbox("Select SOW Type", sow_type_options)

    industry_options = ["Retail / E-commerce", "BFSI", "Manufacturing", "Telecom", "Healthcare", "Energy / Utilities", "Logistics", "Media", "Government", "Other (specify)"]
    industry_type = st.selectbox("Industry / Domain", industry_options)
    final_industry = st.text_input("Specify Industry", placeholder="Enter industry...") if industry_type == "Other (specify)" else industry_type

    duration = st.text_input("Timeline / Duration", "4 Weeks")
    
    if st.button("Reset All Fields", on_click=clear_sow):
        st.rerun()

# -----------------------------
# MAIN UI
# -----------------------------
st.title("🚀 GenAI Scope of Work Architect")

# Cover Page Branding
st.header("📸 Cover Page Branding")
customer_logo = st.file_uploader("Upload Customer Logo (Optional)", type=["png", "jpg", "jpeg"])
doc_date = st.date_input("Document Date", date.today())

# Objectives
st.header("2. Objectives & Stakeholders")
objective = st.text_area("Define the core business objective:", placeholder="e.g., Development of a Gen AI based PoC...", height=120)
outcomes = st.multiselect(
    "Select success metrics:", 
    ["Reduced Response Time", "Automated SOP Mapping", "Cost Savings", "Higher Accuracy", "Metadata Richness", "Revenue Growth", "Security Compliance", "Scalability", "Integration Feasibility"],
    default=["Higher Accuracy", "Cost Savings"]
)

st.subheader("👥 Project Stakeholders")
col1, col2 = st.columns(2)
with col1:
    st.markdown("**Partner Executive Sponsor**")
    st.session_state.stakeholders["Partner"] = st.data_editor(st.session_state.stakeholders["Partner"], num_rows="dynamic", key="ed_partner")
    st.markdown("**AWS Executive Sponsor**")
    st.session_state.stakeholders["AWS"] = st.data_editor(st.session_state.stakeholders["AWS"], num_rows="dynamic", key="ed_aws")

with col2:
    st.markdown("**Customer Executive Sponsor**")
    st.session_state.stakeholders["Customer"] = st.data_editor(st.session_state.stakeholders["Customer"], num_rows="dynamic", key="ed_customer")
    st.markdown("**Project Escalation Contacts**")
    st.session_state.stakeholders["Escalation"] = st.data_editor(st.session_state.stakeholders["Escalation"], num_rows="dynamic", key="ed_escalation")

# -----------------------------
# DOCX GENERATION LOGIC
# -----------------------------
def add_infra_cost_table(doc, sow_type_name):
    """Adds infra cost table below Architecture Diagram"""
    from docx.shared import Inches

    cost_data = SOW_COST_TABLE_MAP.get(sow_type_name)
    if not cost_data:
        return

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "System"
    hdr[1].text = "Infra Cost"
    hdr[2].text = "AWS Cost Calculator Link"

    aws_link = "https://calculator.aws/#/"

    if "poc_cost" in cost_data:
        row = table.add_row().cells
        row[0].text = "POC"
        row[1].text = cost_data["poc_cost"]
        row[2].text = aws_link
    if "prod_cost" in cost_data:
        row = table.add_row().cells
        row[0].text = "Production"
        row[1].text = cost_data["prod_cost"]
        row[2].text = aws_link

    doc.add_paragraph("")  # spacing

def create_docx_logic(text_content, branding_info, sow_type_name):
    """Generates final Word doc preserving structure exactly as user sees it"""
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # COVER PAGE
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_picture(AWS_PN_LOGO, width=Inches(1.6))
    doc.add_paragraph("\n"*3)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(branding_info['sow_name'])
    run.font.size = Pt(26)
    run.bold = True

    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_p.add_run("Scope of Work Document").font.size = Pt(14)

    doc.add_paragraph("\n"*4)

    # Logos row
    logo_table = doc.add_table(rows=1, cols=3)
    logo_table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Customer logo
    cell = logo_table.rows[0].cells[0]
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if branding_info.get("customer_logo_bytes"):
        cell.paragraphs[0].add_run().add_picture(io.BytesIO(branding_info["customer_logo_bytes"]), width=Inches(1.8))
    else:
        cell.paragraphs[0].add_run("Customer Logo").bold = True

    # Oneture logo
    cell = logo_table.rows[0].cells[1]
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.paragraphs[0].add_run().add_picture(ONETURE_LOGO, width=Inches(2.2))

    # AWS Advanced logo
    cell = logo_table.rows[0].cells[2]
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.paragraphs[0].add_run().add_picture(AWS_ADV_LOGO, width=Inches(1.8))

    doc.add_paragraph("\n"*3)

    # Date
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.add_run(branding_info["doc_date_str"]).bold = True

    doc.add_page_break()

    # --- CONTENT PROCESSING ---
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    lines = text_content.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            doc.add_paragraph("")
            continue
        clean_text = re.sub(r'^#+\s*', '', line)
        upper_text = clean_text.upper()

        # SECTION HEADINGS
        if line.startswith('#'):
            doc.add_heading(clean_text, level=1)
        elif line.startswith('##'):
            doc.add_heading(clean_text, level=2)
        elif line.startswith('###'):
            doc.add_heading(clean_text, level=3)
        else:
            doc.add_paragraph(clean_text)

        # SECTION 2.2 STAKEHOLDER TABLES
        if "2.2 PROJECT SPONSOR" in upper_text:
            for key in ["Partner", "Customer", "AWS", "Escalation"]:
                doc.add_heading(f"{key} Executive Sponsor" if key!="Escalation" else "Project Escalation Contacts", level=3)
                df = st.session_state.stakeholders[key]
                if not df.empty:
                    table = doc.add_table(rows=1, cols=len(df.columns))
                    table.style = "Table Grid"
                    hdr_cells = table.rows[0].cells
                    for idx, col_name in enumerate(df.columns):
                        hdr_cells[idx].text = col_name
                    for _, row in df.iterrows():
                        row_cells = table.add_row().cells
                        for idx, col_name in enumerate(df.columns):
                            row_cells[idx].text = str(row[col_name])

        # SECTION 4 ARCHITECTURE
        if "4 SOLUTION ARCHITECTURE" in upper_text:
            diagram_path = SOW_DIAGRAM_MAP.get(sow_type_name)
            if diagram_path and os.path.exists(diagram_path):
                doc.add_paragraph("")
                doc.add_picture(diagram_path, width=Inches(6))
                doc.add_paragraph(f"{sow_type_name} – Architecture Diagram")
                # Add infra cost table here
                add_infra_cost_table(doc, sow_type_name)

        i += 1

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# -----------------------------
# GENERATION BUTTON
# -----------------------------
if st.button("💾 Generate Microsoft Word Document"):
    branding_info = {
        "sow_name": selected_sow_name,
        "customer_logo_bytes": customer_logo.getvalue() if customer_logo else None,
        "doc_date_str": doc_date.strftime("%d %B %Y")
    }

    if not objective:
        st.warning("Enter Objective before generating SOW.")
    else:
        # Combine objective + prefilled markdown for rest
        prefilled_text = f"""1 TABLE OF CONTENTS
2 PROJECT OVERVIEW
2.1 OBJECTIVE
{objective}
2.2 PROJECT SPONSOR(S) / STAKEHOLDER(S) / PROJECT TEAM
2.3 ASSUMPTIONS & DEPENDENCIES
2.4 PoC Success Criteria
3 SCOPE OF WORK – TECHNICAL PROJECT PLAN
3.1 Phase 1: Infrastructure Setup
3.2 Phase 2: Create Core Workflows
3.3 Phase 3: Backend Components Implementation
3.4 Phase 4: Testing and Feedback
4 SOLUTION ARCHITECTURE / ARCHITECTURAL DIAGRAM
6 RESOURCES & COST ESTIMATES
"""
        st.session_state.generated_sow = prefilled_text

        docx_data = create_docx_logic(st.session_state.generated_sow, branding_info, selected_sow_name)
        st.download_button(
            label="📥 Download SOW (.docx)",
            data=docx_data,
            file_name=f"SOW_{selected_sow_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
