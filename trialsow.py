import streamlit as st
from datetime import date
import io
import re
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- 1. DATA MAPPING (Updated with your specific table info) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIAGRAM_DIR = os.path.join(BASE_DIR, "diagrams")

CASE_SPECIFIC_DATA = {
    "L1 Support Bot POC SOW": {"poc": "3,536.40 USD", "diag": "L1 Support Bot POC SOW.png", "link": "https://calculator.aws/#/estimate?id=l1-bot"},
    "Beauty Advisor POC SOW": {"poc": "4,725.66 USD", "prod": "5,701.48 USD", "diag": "Beauty Advisor POC SOW.png", "link": "https://calculator.aws/#/estimate?id=beauty-advisor"},
    "Ready Search POC Scope of Work Document": {"poc": "2,641.40 USD", "diag": "Ready Search POC Scope of Work Document.png", "link": "https://calculator.aws/#/estimate?id=ready-search"},
    "AI based Image Enhancement POC SOW": {"poc": "2,814.34 USD", "diag": "AI based Image Enhancement POC SOW.png", "link": "https://calculator.aws/#/estimate?id=image-enhancement"},
    "AI based Image Inspection POC SOW": {"poc": "3,536.40 USD", "diag": "AI based Image Inspection POC SOW.png", "link": "https://calculator.aws/#/estimate?id=image-inspection"},
    "Gen AI for SOP POC SOW": {"poc": "2,110.30 USD", "diag": "Gen AI for SOP POC SOW.png", "link": "https://calculator.aws/#/estimate?id=sop-gen"},
    "Project Scope Document": {"prod": "2,993.60 USD", "diag": "Project Scope Document.png", "link": "https://calculator.aws/#/estimate?id=project-scope"},
    "Gen AI Speech To Speech": {"prod": "2,124.23 USD", "diag": "Gen AI Speech To Speech.png", "link": "https://calculator.aws/#/estimate?id=speech-to-speech"},
    "PoC Scope Document": {"poc": "2,150 USD + 1,000 USD (Amazon Bedrock) = 3,150 USD", "diag": "PoC Scope Document.png", "link": "https://calculator.aws/#/estimate?id=poc-scope"}
}

# Custom CSS for an Enterprise UI
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stButton>button { border-radius: 8px; font-weight: 600; }
    .stTextArea textarea { border-radius: 10px; }
    .stTextInput input { border-radius: 8px; }
    .block-container { padding-top: 1.5rem; }
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
    h1, h2, h3 { color: #0f172a; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: 600; }
    [data-testid="stExpander"] { border: none; box-shadow: none; background: transparent; }
    .stakeholder-header { 
        background-color: #f1f5f9; 
        padding: 8px 12px; 
        border-radius: 6px; 
        margin-bottom: 10px; 
        font-weight: bold;
        color: #334155;
        border-left: 4px solid #3b82f6;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. DOCUMENT ENGINE ---
def create_docx_logic(text_content, branding_info, case_key):
    doc = Document()
    data = CASE_SPECIFIC_DATA.get(case_key, {})
    
    # --- COVER PAGE ---
    # Top Left: AWS Partner (Fixed)
    p_top = doc.add_paragraph()
    aws_pn = os.path.join(DIAGRAM_DIR, "aws_pn_logo.png")
    if os.path.exists(aws_pn): p_top.add_run().add_picture(aws_pn, width=Inches(1.0))
    
    doc.add_paragraph("\n" * 3)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(case_key)
    run.font.size, run.font.bold = Pt(26), True
    
    doc.add_paragraph("\n" * 4)
    
    # Logo Row (Customer, Oneture, AWS Advanced)
    logo_table = doc.add_table(rows=1, cols=3)
    logo_table.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if branding_info.get('cust_logo'):
        logo_table.rows[0].cells[0].paragraphs[0].add_run().add_picture(io.BytesIO(branding_info['cust_logo']), width=Inches(1.4))
    
    one_logo = os.path.join(DIAGRAM_DIR, "oneture_logo.png")
    if os.path.exists(one_logo):
        logo_table.rows[0].cells[1].paragraphs[0].add_run().add_picture(one_logo, width=Inches(2.2))
        
    aws_adv = os.path.join(DIAGRAM_DIR, "aws_adv_logo.png")
    if os.path.exists(aws_adv):
        logo_table.rows[0].cells[2].paragraphs[0].add_run().add_picture(aws_adv, width=Inches(1.3))

    doc.add_page_break() 
    
    # --- CONTENT PARSING ---
    lines = text_content.split('\n')
    for line in lines:
        clean = re.sub(r'\*+', '', line).strip()
        
        # Inject Architecture
        if "4 SOLUTION ARCHITECTURE" in clean.upper():
            doc.add_heading(clean, level=1)
            diag = os.path.join(DIAGRAM_DIR, data.get('diag', ''))
            if os.path.exists(diag): doc.add_picture(diag, width=Inches(5.8))
            continue
            
        # Inject Commercials Table (Your New Requirement)
        if "6 COMMERCIALS" in clean.upper():
            doc.add_heading("6 Commercials", level=1)
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = "System", "Infra cost", "AWS cost calculator link"
            
            if "poc" in data or "total" in data:
                row = table.add_row().cells
                row[0].text, row[1].text, row[2].text = "POC", data.get('poc', data.get('total')), data.get('link')
            
            if "prod" in data:
                row = table.add_row().cells
                row[0].text, row[1].text, row[2].text = "Production", data.get('prod'), data.get('link')
            continue

        if line.startswith('# '): doc.add_heading(clean, level=1)
        elif line.startswith('## '): doc.add_heading(clean, level=2)
        else: doc.add_paragraph(clean)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 3. UI LAYER ---
st.title("📄 GenAI SOW Architect")
with st.sidebar:
    api_key = st.text_input("Gemini API Key", type="password")
    case = st.selectbox("Use Case", list(CASE_SPECIFIC_DATA.keys()))
    cust_logo = st.file_uploader("Upload Customer Logo (Nykaa)", type=['png', 'jpg'])

if st.button("✨ Generate SOW"):
    # (LLM Call Logic remains same as previous versions)
    st.session_state.generated_sow = f"# 1 TABLE OF CONTENTS\n# 2 PROJECT OVERVIEW\n# 4 SOLUTION ARCHITECTURE\n# 6 COMMERCIALS"

if 'generated_sow' in st.session_state:
    if st.button("💾 Download Document"):
        info = {'cust_logo': cust_logo.getvalue() if cust_logo else None, 'date': date.today().strftime("%d %B %Y")}
        file = create_docx_logic(st.session_state.generated_sow, info, case)
        st.download_button("📥 Click here", data=file, file_name=f"{case}.docx")
