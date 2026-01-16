import streamlit as st
from datetime import date
import io
import os
import re
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

# --- COST TABLE MAPPING ---
SOW_COST_TABLE_MAP = {
    "L1 Support Bot POC SOW": {"System": "POC", "Infra Cost": "3,536.40 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "Beauty Advisor POC SOW": {"System": "POC", "Infra Cost": "4,725.66 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "Ready Search POC Scope of Work Document": {"System": "POC", "Infra Cost": "2,641.40 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "AI based Image Enhancement POC SOW": {"System": "POC", "Infra Cost": "2,814.34 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "AI based Image Inspection POC SOW": {"System": "POC", "Infra Cost": "3,536.40 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "Gen AI for SOP POC SOW": {"System": "POC", "Infra Cost": "2,110.30 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "Project Scope Document": {"System": "Production", "Infra Cost": "2,993.60 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "Gen AI Speech To Speech": {"System": "Production", "Infra Cost": "2,124.23 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
    "PoC Scope Document": {"System": "POC", "Infra Cost": "1,000 USD", "AWS Cost Calculator Link": "https://calculator.aws/#/"},
}

# --- STREAMLIT PAGE CONFIG ---
st.set_page_config(page_title="GenAI SOW Architect", layout="wide", page_icon="📄")

# --- CACHED DOCX CREATION FUNCTION ---
def create_docx_logic(text_content, branding_info, sow_type_name):
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # --- COVER PAGE ---
    # AWS PN Logo always
    if branding_info.get('aws_pn_logo_bytes'):
        p_top = doc.add_paragraph()
        p_top.alignment = WD_ALIGN_PARAGRAPH.LEFT
        try:
            p_top.add_run().add_picture(io.BytesIO(branding_info['aws_pn_logo_bytes']), width=Inches(1.0))
        except:
            p_top.add_run("AWS PN").bold = True

    doc.add_paragraph("\n" * 3)

    # Title
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(branding_info['sow_name'])
    run.font.size = Pt(26)
    run.font.bold = True

    # Subtitle
    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle_p.add_run("Scope of Work Document")
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

    doc.add_paragraph("\n" * 4)

    # Logos table (Customer optional + Oneture + AWS Advanced)
    logo_table = doc.add_table(rows=1, cols=3)
    logo_table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def insert_logo(cell, bytes_data, width_val, fallback_text):
        cell.paragraphs[0].text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if bytes_data:
            try:
                p.add_run().add_picture(io.BytesIO(bytes_data), width=Inches(width_val))
            except:
                p.add_run(fallback_text).bold = True
        else:
            p.add_run(fallback_text).bold = True

    insert_logo(logo_table.rows[0].cells[0], branding_info.get('customer_logo_bytes'), 1.4, "[Customer Logo]")
    insert_logo(logo_table.rows[0].cells[1], branding_info.get('oneture_logo_bytes'), 2.2, "ONETURE")
    insert_logo(logo_table.rows[0].cells[2], branding_info.get('aws_adv_logo_bytes'), 1.3, "AWS Advanced")

    doc.add_paragraph("\n" * 4)
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = date_p.add_run(branding_info['doc_date_str'])
    run.font.size = Pt(12)
    run.font.bold = True

    doc.add_page_break()

    # --- PROCESS GENERATED CONTENT ---
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
        clean_text = re.sub(r'\*+|^#+\s*', '', line).strip()
        upper_text = clean_text.upper()

        # Insert Architecture Diagram + Cost Table
        if "4 SOLUTION ARCHITECTURE" in upper_text:
            doc.add_heading(clean_text, level=1)

            # Architecture diagram
            diagram_path = SOW_DIAGRAM_MAP.get(sow_type_name)
            if diagram_path and os.path.exists(diagram_path):
                doc.add_paragraph("")
                try:
                    doc.add_picture(diagram_path, width=Inches(6.0))
                    p_cap = doc.add_paragraph(f"{sow_type_name} – Architecture Diagram")
                    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except:
                    doc.add_paragraph("[Architecture Diagram Missing]")

            # Insert cost table automatically
            cost_info = SOW_COST_TABLE_MAP.get(sow_type_name)
            if cost_info:
                table = doc.add_table(rows=1, cols=len(cost_info))
                table.style = 'Table Grid'
                hdr_cells = table.rows[0].cells
                for idx, key in enumerate(cost_info.keys()):
                    hdr_cells[idx].text = key
                row_cells = table.add_row().cells
                for idx, key in enumerate(cost_info.keys()):
                    row_cells[idx].text = str(cost_info[key])

            i += 1
            continue

        # Stakeholder tables, TOC, headings preserved exactly
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
    return bio.getvalue()


# --- STREAMLIT UI ---
st.title("🚀 GenAI Scope of Work Architect")

# --- Branding Uploads ---
st.header("📸 Cover Page Branding")
col1, col2 = st.columns(2)
with col1:
    aws_pn_logo = st.file_uploader("AWS PN Logo (Always Included)", type=['png','jpg','jpeg'], key="aws_pn")
    customer_logo = st.file_uploader("Customer Logo (Optional)", type=['png','jpg','jpeg'], key="customer_logo")
with col2:
    oneture_logo = st.file_uploader("Oneture Logo (Always Included)", type=['png','jpg','jpeg'], key="oneture_logo")
    aws_adv_logo = st.file_uploader("AWS Advanced Logo (Always Included)", type=['png','jpg','jpeg'], key="aws_adv_logo")
    doc_date = st.date_input("Document Date", date.today())

# --- SOW Selection & Inputs ---
sow_options = list(SOW_DIAGRAM_MAP.keys())
selected_sow_name = st.selectbox("Select SOW Type:", sow_options)
objective = st.text_area("Business Objective:", height=120)
duration = st.text_input("Timeline/Duration:", "4 Weeks")

# --- Stakeholder Tables ---
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

st.subheader("👥 Project Stakeholders")
for key in st.session_state.stakeholders:
    st.markdown(f"**{key} Sponsors / Contacts**")
    st.session_state.stakeholders[key] = st.data_editor(
        st.session_state.stakeholders[key], num_rows="dynamic", use_container_width=True
    )

# --- LLM SOW Generation ---
api_key = st.text_input("Gemini API Key:", type="password")
if st.button("✨ Generate SOW Document"):
    if not api_key or not objective:
        st.warning("Enter API Key and Business Objective.")
    else:
        import requests

        prompt_text = f"""
        Generate COMPLETE SOW for {selected_sow_name} in professional tone.
        Objective: {objective}
        Timeline: {duration}
        Include: Stakeholder tables (Partner, Customer, AWS, Escalation),
        Sections 2.3, 2.4, 3 phases, 4 Solution Architecture, 6 Resources & Cost Estimates.
        Markdown only.
        """
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-preview-09-2025:generateContent?key={api_key}"
        payload = {
            "contents":[{"parts":[{"text": prompt_text}]}],
            "systemInstruction":{"parts":[{"text":"You are a senior Solutions Architect. Generate detailed SOW strictly following structure, stakeholder tables, phases, and cost table."}]}
        }
        with st.spinner("Generating SOW..."):
            try:
                res = requests.post(url, json=payload)
                if res.status_code == 200:
                    st.session_state.generated_sow = res.json()['candidates'][0]['content']['parts'][0]['text']
                    st.balloons()
                else:
                    st.error(f"API Error: {res.text}")
            except Exception as e:
                st.error(f"Error: {str(e)}")

# --- Review & Export ---
if 'generated_sow' in st.session_state and st.session_state.generated_sow:
    st.header("3. Review & Export")
    st.text_area("Edit SOW Content:", value=st.session_state.generated_sow, height=600)
    if st.button("💾 Download Word Document"):
        branding_info = {
            'sow_name': selected_sow_name,
            'aws_pn_logo_bytes': aws_pn_logo.getvalue() if aws_pn_logo else None,
            'customer_logo_bytes': customer_logo.getvalue() if customer_logo else None,
            'oneture_logo_bytes': oneture_logo.getvalue() if oneture_logo else None,
            'aws_adv_logo_bytes': aws_adv_logo.getvalue() if aws_adv_logo else None,
            'doc_date_str': doc_date.strftime("%d %B %Y")
        }
        docx_bytes = create_docx_logic(st.session_state.generated_sow, branding_info, selected_sow_name)
        st.download_button(
            "📥 Download SOW (.docx)", 
            data=docx_bytes, 
            file_name=f"SOW_{selected_sow_name.replace(' ','_')}.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
