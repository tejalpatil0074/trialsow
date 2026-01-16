import streamlit as st
from datetime import date
import io
import re
import os
import requests
import pandas as pd

# --- FILE PATHING & DIAGRAM MAPPING ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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

# --- COST TABLES ---
SOW_COST_TABLE_MAP = {
    "L1 Support Bot POC SOW": {"poc_cost": "3,536.40 USD"},
    "Beauty Advisor POC SOW": {"poc_cost": "4,725.66 USD", "prod_cost": "5,701.48 USD"},
    "Ready Search POC Scope of Work Document": {"poc_cost": "2,641.40 USD"},
    "AI based Image Enhancement POC SOW": {"poc_cost": "2,814.34 USD"},
    "AI based Image Inspection POC SOW": {"poc_cost": "3,536.40 USD"},
    "Gen AI for SOP POC SOW": {"poc_cost": "2,110.30 USD"},
    "Project Scope Document": {"prod_cost": "2,993.60 USD"},
    "Gen AI Speech To Speech": {"prod_cost": "2,124.23 USD"},
    "PoC Scope Document": {"amazon_bedrock": "1,000 USD", "total": "$ 3,150"}
}

# --- CONFIGURATION ---
st.set_page_config(
    page_title="GenAI SOW Architect", 
    layout="wide", 
    page_icon="📄",
    initial_sidebar_state="expanded"
)

# --- STYLING ---
st.markdown("""
<style>
.stButton>button { border-radius: 8px; font-weight: 600; }
.stTextArea textarea { border-radius: 10px; }
.stTextInput input { border-radius: 8px; }
.stakeholder-header { 
    background-color: #f1f5f9; 
    padding: 8px 12px; 
    border-radius: 6px; 
    margin-bottom: 10px; 
    font-weight: bold;
    color: #334155;
    border-left: 4px solid #3b82f6;
}
.sow-preview {
    background-color: white;
    padding: 40px;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    line-height: 1.7;
    color: #1e293b;
    box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
}
</style>
""", unsafe_allow_html=True)

# --- DOCX GENERATION LOGIC ---
def create_docx_logic(text_content, branding_info, sow_type_name):
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # --- COVER PAGE ---
    doc.add_paragraph("\n"*3)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(branding_info['sow_name']); run.bold=True; run.font.size=Pt(26)
    sub = doc.add_paragraph(); sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run("Scope of Work Document").font.size = Pt(14)
    doc.add_paragraph("\n"*4)

    # Logos Table
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cells = table.rows[0].cells

    def insert_logo(cell, bytes_data, width, fallback):
        cell.paragraphs[0].text = ""
        p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if bytes_data:
            try:
                p.add_run().add_picture(io.BytesIO(bytes_data), width=Inches(width))
            except:
                p.add_run(fallback).bold=True
        else:
            p.add_run(fallback).bold=True

    insert_logo(cells[0], branding_info.get("customer_logo_bytes"), 1.8, "Customer Logo")
    insert_logo(cells[1], branding_info.get("oneture_logo_bytes"), 2.2, "ONETURE")
    insert_logo(cells[2], branding_info.get("aws_adv_logo_bytes"), 1.8, "AWS Advanced")

    # Date
    doc.add_paragraph("\n"*2)
    date_p = doc.add_paragraph(); date_p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    date_p.add_run(branding_info["doc_date_str"]).bold=True

    doc.add_page_break()

    # --- BODY ---
    lines = text_content.split("\n")
    i=0
    while i<len(lines):
        line = lines[i].strip()
        if not line:
            i+=1
            continue
        clean_text = re.sub(r'^#+\s*','',line).strip()
        upper_text = clean_text.upper()

        # Insert headings
        if line.startswith("#") or re.match(r'^\d+ ', line):
            doc.add_heading(clean_text, level=1)

            # Insert Solution Architecture Diagram + Cost Table
            if "4 SOLUTION ARCHITECTURE" in upper_text:
                diagram_path = SOW_DIAGRAM_MAP.get(sow_type_name)
                if diagram_path and os.path.exists(diagram_path):
                    doc.add_paragraph("")
                    doc.add_picture(diagram_path, width=Inches(6))
                    doc.add_paragraph(f"{sow_type_name} – Architecture Diagram")
                # Cost Table
                cost_data = SOW_COST_TABLE_MAP.get(sow_type_name, {})
                if cost_data:
                    table2 = doc.add_table(rows=1,cols=3)
                    table2.style='Table Grid'
                    hdr = table2.rows[0].cells
                    hdr[0].text='System'; hdr[1].text='Infra Cost'; hdr[2].text='AWS Cost Calculator Link'
                    aws_link = "https://calculator.aws/#/"
                    if 'poc_cost' in cost_data:
                        row=table2.add_row().cells
                        row[0].text='POC'; row[1].text=cost_data['poc_cost']; row[2].text=aws_link
                    if 'prod_cost' in cost_data:
                        row=table2.add_row().cells
                        row[0].text='Production'; row[1].text=cost_data['prod_cost']; row[2].text=aws_link
                doc.add_paragraph("")
        elif line.startswith("|"):
            # Parse table
            table_lines=[]
            while i<len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip()); i+=1
            if len(table_lines)>=3:
                headers=[c.strip() for c in table_lines[0].split("|") if c.strip()]
                tbl=doc.add_table(rows=1,cols=len(headers))
                tbl.style='Table Grid'
                for idx,h in enumerate(headers): tbl.rows[0].cells[idx].text=h
                for row_str in table_lines[1:]:
                    row_cells=tbl.add_row().cells
                    r_data=[c.strip() for c in row_str.split("|") if c.strip()]
                    for idx,c_text in enumerate(r_data):
                        if idx<len(row_cells): row_cells[idx].text=c_text
            doc.add_paragraph("")
            continue
        else:
            doc.add_paragraph(clean_text)
        i+=1
    bio=io.BytesIO(); doc.save(bio); bio.seek(0)
    return bio.getvalue()


# --- INITIALIZATION ---
if 'generated_sow' not in st.session_state: st.session_state.generated_sow = ""
if 'stakeholders' not in st.session_state:
    st.session_state.stakeholders = {
        "Partner": pd.DataFrame([{"Name":"Gaurav Kankaria","Title":"Head of Analytics & ML","Email":"gaurav.kankaria@oneture.com"}]),
        "Customer": pd.DataFrame([{"Name":"Cheten Dev","Title":"Head of Product Design","Email":"cheten.dev@nykaa.com"}]),
        "AWS": pd.DataFrame([{"Name":"Anubhav Sood","Title":"AWS Account Executive","Email":"anbhsood@amazon.com"}]),
        "Escalation": pd.DataFrame([{"Name":"Omkar Dhavalikar","Title":"AI/ML Lead","Email":"omkar.dhavalikar@oneture.com"},
                                    {"Name":"Gaurav Kankaria","Title":"Head of Analytics and AIML","Email":"gaurav.kankaria@oneture.com"}])
    }

def clear_sow(): st.session_state.generated_sow=""

# --- SIDEBAR: PROJECT INTAKE ---
with st.sidebar:
    st.title("SOW Architect")
    st.caption("Enterprise POC/MVP Engine")
    api_key = st.text_input("Gemini API Key", type="password")

    sow_type_options = list(SOW_DIAGRAM_MAP.keys())
    selected_sow_name = st.selectbox("Scope of Work Type", sow_type_options)
    industry_options = ["Retail / E-commerce","BFSI","Manufacturing","Telecom","Healthcare","Energy / Utilities","Logistics","Media","Government","Other (specify)"]
    industry_type = st.selectbox("Industry / Domain", industry_options)
    final_industry = st.text_input("Specify Industry", placeholder="Enter industry...") if industry_type=="Other (specify)" else industry_type
    duration = st.text_input("Timeline / Duration", "4 Weeks")

    if st.button("🗑️ Reset All Fields", on_click=clear_sow):
        st.rerun()

# --- COVER PAGE LOGOS ---
st.header("📸 Cover Page Branding")
col1,col2=st.columns(2)
with col1:
    aws_pn_logo=st.file_uploader("AWS Partner Network Logo", type=['png','jpg','jpeg'])
    customer_logo=st.file_uploader("Customer Logo", type=['png','jpg','jpeg'])
with col2:
    oneture_logo=st.file_uploader("Oneture Logo", type=['png','jpg','jpeg'])
    aws_adv_logo=st.file_uploader("AWS Advanced Logo", type=['png','jpg','jpeg'])
    doc_date=st.date_input("Document Date", date.today())

# --- OBJECTIVE & STAKEHOLDERS ---
st.header("2. Objectives & Stakeholders")
objective=st.text_area("Business Objective", height=120)
outcomes=st.multiselect("Select success metrics:", ["Reduced Response Time", "Automated SOP Mapping","Cost Savings","Higher Accuracy","Metadata Richness","Revenue Growth","Security Compliance","Scalability","Integration Feasibility"], default=["Higher Accuracy","Cost Savings"])

col_team1,col_team2=st.columns(2)
with col_team1:
    st.markdown('<div class="stakeholder-header">Partner Executive Sponsor</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["Partner"]=st.data_editor(st.session_state.stakeholders["Partner"], num_rows="dynamic", key="ed_partner")
    st.markdown('<div class="stakeholder-header">AWS Executive Sponsor</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["AWS"]=st.data_editor(st.session_state.stakeholders["AWS"], num_rows="dynamic", key="ed_aws")
with col_team2:
    st.markdown('<div class="stakeholder-header">Customer Executive Sponsor</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["Customer"]=st.data_editor(st.session_state.stakeholders["Customer"], num_rows="dynamic", key="ed_customer")
    st.markdown('<div class="stakeholder-header">Project Escalation Contacts</div>', unsafe_allow_html=True)
    st.session_state.stakeholders["Escalation"]=st.data_editor(st.session_state.stakeholders["Escalation"], num_rows="dynamic", key="ed_escalation")

# --- GENERATE LLM SOW ---
if st.button("✨ Generate SOW Document"):
    if not api_key: st.warning("Enter Gemini API Key"); st.stop()
    if not objective: st.warning("Enter Business Objective"); st.stop()

    prompt_text=f"""
Generate a COMPLETE formal enterprise Scope of Work (SOW) for {selected_sow_name} in {final_industry}.
Include sections: Table of Contents, Project Overview, Objective, Stakeholders (tables), Assumptions, PoC Success Criteria, Technical Plan, Solution Architecture, Resources & Cost Estimates.
PoC Success Criteria: {', '.join(outcomes)}
Timeline: {duration}
"""
    payload={"contents":[{"parts":[{"text":prompt_text}]}],"systemInstruction":{"parts":[{"text":"You are a senior Solutions Architect. Generate detailed SOW with numbered headings and stakeholder tables. No filler."}]}}
    try:
        res=requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key={api_key}", json=payload)
        if res.status_code==200:
            st.session_state.generated_sow=res.json()['candidates'][0]['content']['parts'][0]['text']
            st.balloons()
        else: st.error(f"API Error: {res.text}")
    except Exception as e: st.error(str(e))

# --- REVIEW & EXPORT ---
if st.session_state.generated_sow:
    st.header("3. Review & Export")
    tab_edit, tab_preview=st.tabs(["✍️ Edit","📄 Preview"])
    with tab_edit:
        st.session_state.generated_sow=st.text_area("Edit SOW", st.session_state.generated_sow, height=700)
    with tab_preview:
        st.markdown(f'<div class="sow-preview">{st.session_state.generated_sow}</div>', unsafe_allow_html=True)

    if st.button("💾 Download Word Document"):
        branding_info={
            "sow_name": selected_sow_name,
            "aws_pn_logo_bytes": aws_pn_logo.getvalue() if aws_pn_logo else None,
            "customer_logo_bytes": customer_logo.getvalue() if customer_logo else None,
            "oneture_logo_bytes": oneture_logo.getvalue() if oneture_logo else None,
            "aws_adv_logo_bytes": aws_adv_logo.getvalue() if aws_adv_logo else None,
            "doc_date_str": doc_date.strftime("%d %B %Y")
        }
        docx_data=create_docx_logic(st.session_state.generated_sow, branding_info, selected_sow_name)
        st.download_button("📥 Download SOW (.docx)", docx_data, f"SOW_{selected_sow_name.replace(' ','_')}.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document")
