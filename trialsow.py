import streamlit as st
from datetime import date
import io
import re
import os
import pandas as pd

# --- FILE PATHING & DIAGRAM MAPPING ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "diagrams")

AWS_PN_LOGO = os.path.join(ASSETS_DIR, "aws partner logo.jpg")
ONETURE_LOGO = os.path.join(ASSETS_DIR, "oneture logo1.jpg")
AWS_ADV_LOGO = os.path.join(ASSETS_DIR, "aws advanced logo1.jpg")

SOW_DIAGRAM_MAP = {
    "L1 Support Bot POC SOW": os.path.join(BASE_DIR, "diagrams", "L1 Support Bot POC SOW.png"),
    "Ready Search POC Scope of Work Document": os.path.join(BASE_DIR, "diagrams", "Ready Search POC Scope of Work Document.png"),
    "AI based Image Enhancement POC SOW": os.path.join(BASE_DIR, "diagrams", "AI based Image Enhancement POC SOW.png"),
    "Beauty Advisor POC SOW": os.path.join(BASE_DIR, "diagrams", "Beauty Advisor POC SOW.png"),
    "AI based Image Inspection POC SOW": os.path.join(BASE_DIR, "diagrams", "AI based Image Inspection POC SOW.png"),
    "Gen AI for SOP POC SOW": os.path.join(BASE_DIR, "diagrams", "Gen AI for SOP POC SOW.png"),
    "Project Scope Document": os.path.join(BASE_DIR, "diagrams", "Project Scope Document.png"),
    "Gen AI Speech To Speech": os.path.join(BASE_DIR, "diagrams", "Gen AI Speech To Speech.png"),
    "PoC Scope Document": os.path.join(BASE_DIR, "diagrams", "PoC Scope Document.png")
}

# --- COST TABLE MAPPING ---
SOW_COST_TABLE_MAP = {
    "L1 Support Bot POC SOW": {"poc_cost": "3,536.40 USD"},
    "Beauty Advisor POC SOW": {
        "poc_cost": "4,525.66 USD + 200 USD (Amazon Bedrock Cost) = 4,725.66",
        "prod_cost": "4,525.66 USD + 1,175.82 USD (Amazon Bedrock Cost) = 5,701.48"
    },
    "Ready Search POC Scope of Work Document":{"poc_cost": "2,641.40 USD"},
    "AI based Image Enhancement POC SOW": {"poc_cost": "2,814.34 USD"},
    "AI based Image Inspection POC SOW": {"poc_cost": "3,536.40 USD"},
    "Gen AI for SOP POC SOW": {"poc_cost": "2,110.30 USD"},
    "Project Scope Document": {"prod_cost": "2,993.60 USD"},
    "Gen AI Speech To Speech": {"prod_cost": "2,124.23 USD"},
    "PoC Scope Document": {"amazon_bedrock": "1,000 USD", "total": "$ 3,150"}
}

# --- STREAMLIT CONFIG ---
st.set_page_config(page_title="GenAI SOW Architect", layout="wide", page_icon="📄", initial_sidebar_state="expanded")

# Custom CSS for styling
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stButton>button { border-radius: 8px; font-weight: 600; }
    .stTextArea textarea { border-radius: 10px; }
    .stTextInput input { border-radius: 8px; }
    .block-container { padding-top: 1.5rem; }
    .sow-preview { background-color: white; padding: 40px; border-radius: 12px; border: 1px solid #e2e8f0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.7; color: #1e293b; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); }
    h1, h2, h3 { color: #0f172a; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: 600; }
    [data-testid="stExpander"] { border: none; box-shadow: none; background: transparent; }
    .stakeholder-header { background-color: #f1f5f9; padding: 8px 12px; border-radius: 6px; margin-bottom: 10px; font-weight: bold; color: #334155; border-left: 4px solid #3b82f6; }
    </style>
""", unsafe_allow_html=True)

# --- FUNCTION TO ADD COST TABLE ---
def add_infra_cost_table(doc, sow_type_name):
    """Adds a formatted cost table in Word document per SOW"""
    from docx.shared import Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    cost_data = SOW_COST_TABLE_MAP.get(sow_type_name)
    if not cost_data:
        return

    doc.add_paragraph("")  # spacing
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
    if "amazon_bedrock" in cost_data:
        row = table.add_row().cells
        row[0].text = "Amazon Bedrock"
        row[1].text = cost_data["amazon_bedrock"]
        row[2].text = aws_link
    if "total" in cost_data:
        row = table.add_row().cells
        row[0].text = "Total Cost"
        row[1].text = cost_data["total"]
        row[2].text = aws_link

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("")

# --- FUNCTION TO CREATE DOCX ---
def create_docx_logic(text_content, branding_info, sow_type_name):
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import io

    doc = Document()

    # COVER PAGE
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_picture(AWS_PN_LOGO, width=Inches(1.6))

    doc.add_paragraph("\n" * 3)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(branding_info['sow_name'])
    run.font.size = Pt(26)
    run.bold = True

    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_p.add_run("Scope of Work Document").font.size = Pt(14)
    doc.add_paragraph("\n" * 4)

    # Logos
    logo_table = doc.add_table(rows=1, cols=3)
    logo_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell = logo_table.rows[0].cells[0]
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if branding_info.get("customer_logo_bytes"):
        cell.paragraphs[0].add_run().add_picture(io.BytesIO(branding_info["customer_logo_bytes"]), width=Inches(1.8))
    else:
        cell.paragraphs[0].add_run("Customer Logo").bold = True

    cell = logo_table.rows[0].cells[1]
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.paragraphs[0].add_run().add_picture(ONETURE_LOGO, width=Inches(2.2))

    cell = logo_table.rows[0].cells[2]
    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell.paragraphs[0].add_run().add_picture(AWS_ADV_LOGO, width=Inches(1.8))

    doc.add_paragraph("\n" * 3)
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_p.add_run(branding_info["doc_date_str"]).bold = True
    doc.add_page_break()

    # MAIN CONTENT
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)
    lines = text_content.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        clean_text = re.sub(r'^#+\s*', '', line)
        upper_text = clean_text.upper()

        if "4 SOLUTION ARCHITECTURE" in upper_text:
            doc.add_heading(clean_text, level=1)
            diagram_path = SOW_DIAGRAM_MAP.get(sow_type_name)
            if diagram_path and os.path.exists(diagram_path):
                try:
                    doc.add_picture(diagram_path, width=Inches(6.0))
                    p_cap = doc.add_paragraph(f"{sow_type_name} – Architecture Diagram")
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except:
                    doc.add_paragraph("[Architecture Diagram Missing]")
            add_infra_cost_table(doc, sow_type_name)
            i += 1
            continue

        if line.startswith('# '):
            doc.add_heading(clean_text, level=1)
        elif line.startswith('## '):
            doc.add_heading(clean_text, level=2)
        elif line.startswith('### '):
            doc.add_heading(clean_text, level=3)
        elif line.startswith('- ') or line.startswith('* '):
            doc.add_paragraph(clean_text[2:], style='List Bullet')
        else:
            doc.add_paragraph(clean_text)
        i += 1

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# --- INITIALIZE SESSION STATE ---
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

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/artificial-intelligence.png", width=60)
    st.title("SOW Architect")
    st.caption("Enterprise POC/MVP Engine")
    with st.expander("🔑 API Key", expanded=False):
        api_key = st.text_input("Gemini API Key", type="password")

# --- MAIN UI ---
st.title("🚀 GenAI Scope of Work Architect")
st.header("📸 Cover Page Branding")
customer_logo = st.file_uploader("Upload Customer Logo (Optional)", type=["png","jpg","jpeg"])
doc_date = st.date_input("Document Date", date.today())

# Objective & Stakeholders
st.header("2. Objectives & Stakeholders")
objective = st.text_area("Define the core business objective:", placeholder="e.g., Development of a Gen AI based WIMO Bot...", height=120)
outcomes = st.multiselect("Select success metrics:", ["Reduced Response Time","Automated SOP Mapping","Cost Savings","Higher Accuracy","Metadata Richness","Revenue Growth","Security Compliance","Scalability","Integration Feasibility"], default=["Higher Accuracy","Cost Savings"])

# Stakeholders
col_team1, col_team2 = st.columns(2)
with col_team1:
    st.markdown('<div class="stakeholder-header">Partner Executive Sponsor</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["Partner"] = st.data_editor(st.session_state.stakeholders["Partner"], num_rows="dynamic", use_container_width=True, key="ed_partner")
    st.markdown('<div class="stakeholder-header">AWS Executive Sponsor</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["AWS"] = st.data_editor(st.session_state.stakeholders["AWS"], num_rows="dynamic", use_container_width=True, key="ed_aws")
with col_team2:
    st.markdown('<div class="stakeholder-header">Customer Executive Sponsor</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["Customer"] = st.data_editor(st.session_state.stakeholders["Customer"], num_rows="dynamic", use_container_width=True, key="ed_customer")
    st.markdown('<div class="stakeholder-header">Project Escalation Contacts</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["Escalation"] = st.data_editor(st.session_state.stakeholders["Escalation"], num_rows="dynamic", use_container_width=True, key="ed_escalation")

# Generate SOW
selected_sow_name = st.selectbox("Select SOW Type", list(SOW_DIAGRAM_MAP.keys()))
if st.button("✨ Generate SOW Document", type="primary"):
    if not objective:
        st.error("⚠️ Business Objective is required.")
    else:
        st.session_state.generated_sow = f"# {selected_sow_name}\n\n## 2.1 Objective\n{objective}\n\n## 2.2 Stakeholders\n" \
            "### Partner Executive Sponsor\n" + st.session_state.stakeholders["Partner"].to_markdown(index=False) + "\n\n" \
            "### Customer Executive Sponsor\n" + st.session_state.stakeholders["Customer"].to_markdown(index=False) + "\n\n" \
            "### AWS Executive Sponsor\n" + st.session_state.stakeholders["AWS"].to_markdown(index=False) + "\n\n" \
            "### Project Escalation Contacts\n" + st.session_state.stakeholders["Escalation"].to_markdown(index=False) + "\n\n" \
            "3 SCOPE OF WORK\n4 SOLUTION ARCHITECTURE\nSpecifics to be discussed basis POC"

# Review & Export
if st.session_state.generated_sow:
    st.divider()
    st.header("3. Review & Export")
    tab_edit, tab_preview = st.tabs(["✍️ Document Editor", "📄 Visual Preview"])
    with tab_edit:
        st.session_state.generated_sow = st.text_area("Modify generated content:", value=st.session_state.generated_sow, height=700, key="sow_editor")
    with tab_preview:
        st.markdown(f'<div class="sow-preview">{st.session_state.generated_sow}</div>', unsafe_allow_html=True)

    if st.button("💾 Prepare Microsoft Word Document"):
        branding_info = {
            "sow_name": selected_sow_name,
            "customer_logo_bytes": customer_logo.getvalue() if customer_logo else None,
            "doc_date_str": doc_date.strftime("%d %B %Y")
        }
        docx_data = create_docx_logic(st.session_state.generated_sow, branding_info, selected_sow_name)
        st.download_button("📥 Download Now (.docx)", data=docx_data, file_name=f"SOW_{selected_sow_name.replace(' ','_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
