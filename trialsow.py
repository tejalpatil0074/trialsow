import streamlit as st
from datetime import date
import io, re, os
import pandas as pd
import requests

# =====================================================
# PATHS & ASSETS
# =====================================================
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
    "PoC Scope Document": os.path.join(ASSETS_DIR, "PoC Scope Document.png"),
}

# =====================================================
# COST TABLE MAP
# =====================================================
SOW_COST_TABLE_MAP = {
    "L1 Support Bot POC SOW": {"poc_cost": "3,536.40 USD"},
    "Beauty Advisor POC SOW": {"poc_cost": "4,725.66 USD", "prod_cost": "5,701.48 USD"},
    "Ready Search POC Scope of Work Document": {"poc_cost": "2,641.40 USD"},
    "AI based Image Enhancement POC SOW": {"poc_cost": "2,814.34 USD"},
    "AI based Image Inspection POC SOW": {"poc_cost": "3,536.40 USD"},
    "Gen AI for SOP POC SOW": {"poc_cost": "2,110.30 USD"},
    "Project Scope Document": {"prod_cost": "2,993.60 USD"},
    "Gen AI Speech To Speech": {"prod_cost": "2,124.23 USD"},
    "PoC Scope Document": {"amazon_bedrock": "1,000 USD", "total": "3,150 USD"},
}

# =====================================================
# STREAMLIT CONFIG
# =====================================================
st.set_page_config("GenAI SOW Architect", layout="wide", page_icon="📄")

# =====================================================
# WORD – COST TABLE
# =====================================================
def add_infra_cost_table(doc, sow_type_name):
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    cost_data = SOW_COST_TABLE_MAP.get(sow_type_name)
    if not cost_data:
        return

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "System"
    hdr[1].text = "Infra Cost"
    hdr[2].text = "AWS Cost Calculator"

    aws_link = "https://calculator.aws/#/"

    if "poc_cost" in cost_data:
        r = table.add_row().cells
        r[0].text = "POC"
        r[1].text = cost_data["poc_cost"]
        r[2].text = aws_link

    if "prod_cost" in cost_data:
        r = table.add_row().cells
        r[0].text = "Production"
        r[1].text = cost_data["prod_cost"]
        r[2].text = aws_link

    if "amazon_bedrock" in cost_data:
        r = table.add_row().cells
        r[0].text = "Amazon Bedrock"
        r[1].text = cost_data["amazon_bedrock"]
        r[2].text = aws_link

    if "total" in cost_data:
        r = table.add_row().cells
        r[0].text = "Total"
        r[1].text = cost_data["total"]
        r[2].text = aws_link

    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# =====================================================
# WORD DOCUMENT GENERATION
# =====================================================
def create_docx_logic(text, branding, sow_name):
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # --- COVER ---
    doc.add_picture(AWS_PN_LOGO, width=Inches(1.6))
    doc.add_paragraph("")

    title = doc.add_paragraph(branding["sow_name"])
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].bold = True
    title.runs[0].font.size = Pt(26)

    doc.add_paragraph("")

    logos = doc.add_table(1, 3)

    if branding.get("customer_logo_bytes"):
        logos.rows[0].cells[0].paragraphs[0].add_run().add_picture(
            io.BytesIO(branding["customer_logo_bytes"]), width=Inches(1.8)
        )

    logos.rows[0].cells[1].paragraphs[0].add_run().add_picture(
        ONETURE_LOGO, width=Inches(2)
    )
    logos.rows[0].cells[2].paragraphs[0].add_run().add_picture(
        AWS_ADV_LOGO, width=Inches(1.8)
    )

    doc.add_page_break()

    # --- CONTENT ---
    for line in text.split("\n"):
        clean = re.sub(r'[#*]+', '', line).strip()
        upper = clean.upper()

        if "4 SOLUTION ARCHITECTURE" in upper:
            doc.add_heading(clean, 1)
            path = SOW_DIAGRAM_MAP.get(sow_name)
            if path and os.path.exists(path):
                doc.add_picture(path, width=Inches(6))
            continue

        if "6 RESOURCES & COST ESTIMATES" in upper:
            doc.add_heading(clean, 1)
            add_infra_cost_table(doc, sow_name)
            continue

        if line.startswith("#"):
            doc.add_heading(clean, 1)
        else:
            doc.add_paragraph(clean)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()

# =====================================================
# SIDEBAR
# =====================================================
st.sidebar.title("SOW Architect")
api_key = st.sidebar.text_input("Gemini API Key", type="password")

sow_options = list(SOW_COST_TABLE_MAP.keys())
selected_sow = st.sidebar.selectbox("SOW Type", sow_options)
selected_sow_name = selected_sow

st.sidebar.subheader("💰 Cost Preview")
cost = SOW_COST_TABLE_MAP.get(selected_sow)
if cost:
    df = pd.DataFrame(
        [{"System": k, "Cost": v, "AWS Link": "https://calculator.aws/#/"} for k, v in cost.items()]
    )
    st.sidebar.table(df)

# =====================================================
# MAIN UI
# =====================================================
st.title("GenAI Scope of Work Architect")

customer_logo = st.file_uploader("Upload Customer Logo (optional)", ["png", "jpg", "jpeg"])
doc_date = st.date_input("Document Date", date.today())
objective = st.text_area("Business Objective", height=120)
duration = st.text_input("Project Duration", "6–8 Weeks")
final_industry = "Enterprise"

# =====================================================
# STAKEHOLDERS (SAFE INIT)
# =====================================================
if "stakeholders" not in st.session_state:
    empty_df = pd.DataFrame(columns=["Name", "Role", "Email"])
    st.session_state.stakeholders = {
        "Partner": empty_df,
        "Customer": empty_df,
        "AWS": empty_df,
        "Escalation": empty_df,
    }

# =====================================================
# GENERATION
# =====================================================
if st.button("✨ Generate SOW Document", type="primary", use_container_width=True):
    if not api_key:
        st.warning("⚠️ Enter a Gemini API Key in the sidebar.")
    elif not objective:
        st.error("⚠️ Business Objective is required.")
    else:
        with st.spinner(f"Architecting {selected_sow_name}..."):

            url = (
                "https://generativelanguage.googleapis.com/v1beta/"
                f"models/gemini-2.5-flash-preview-09-2025:generateContent?key={api_key}"
            )

            def get_md(df):
                return df.to_markdown(index=False)

            prompt_text = f"""
Generate a COMPLETE formal enterprise Scope of Work (SOW) for {selected_sow_name} in {final_industry}.
STRICT PAGE & SECTION FLOW:
1 TABLE OF CONTENTS
2 PROJECT OVERVIEW
2.1 OBJECTIVE
2.2 PROJECT SPONSORS / STAKEHOLDERS
{get_md(st.session_state.stakeholders["Partner"])}
2.3 ASSUMPTIONS
2.4 PoC Success Criteria
3 SCOPE OF WORK
4 SOLUTION ARCHITECTURE / ARCHITECTURAL DIAGRAM
6 RESOURCES & COST ESTIMATES
"""

            payload = {
                "contents": [{"parts": [{"text": prompt_text}]}],
                "systemInstruction": {
                    "parts": [{
                        "text": "You are a senior Solutions Architect. Follow numbering strictly."
                    }]
                }
            }

            res = requests.post(url, json=payload, timeout=90)
            if res.status_code == 200:
                st.session_state.generated_sow = res.json()["candidates"][0]["content"]["parts"][0]["text"]
                st.balloons()
            else:
                st.error(res.text)

# =====================================================
# REVIEW & EXPORT
# =====================================================
if st.session_state.get("generated_sow"):
    st.header("Review & Export")

    st.session_state.generated_sow = st.text_area(
        "Edit Document",
        st.session_state.generated_sow,
        height=700
    )

    if st.button("💾 Prepare Microsoft Word Document"):
        branding_info = {
            "sow_name": selected_sow_name,
            "customer_logo_bytes": customer_logo.getvalue() if customer_logo else None,
        }

        docx_data = create_docx_logic(
            st.session_state.generated_sow,
            branding_info,
            selected_sow_name
        )

        st.download_button(
            "📥 Download Now (.docx)",
            docx_data,
            file_name=f"SOW_{selected_sow_name.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
