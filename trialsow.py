import streamlit as st
from datetime import date
import io
import re

# ======================================================
# ARCHITECTURE IMAGE MAPPING (ADDED)
# ======================================================
ARCHITECTURE_IMAGE_MAP = {
    "Beauty Advisor POC SOW": "assets/architectures/beauty_advisor.png",
    "Ready Search POC Scope of Work Document": "assets/architectures/ready_search.png",
    "AI based Image Inspection POC SOW": "assets/architectures/image_inspection.png",
    "AI based Image Enhancement POC SOW": "assets/architectures/image_enhancement.png",
    "L1 Support Bot POC SOW": "assets/architectures/l1_support_bot.png",
    "Poc Scope Document": "assets/architectures/poc_scope.png",
    "Gen AI Speech To Speech": "assets/architectures/speech_to_speech.png",
    "Project Scope Document": "assets/architectures/project_scope.png",
    "Gen AI for SOP POC SOW": "assets/architectures/genai_sop.png"
}

# ======================================================
# STREAMLIT CONFIG
# ======================================================
st.set_page_config(
    page_title="GenAI SOW Architect",
    layout="wide",
    page_icon="📄",
    initial_sidebar_state="expanded"
)

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


# ======================================================
# DOCX GENERATION LOGIC (UNCHANGED + ADDITION)
# ======================================================
def create_docx_logic(text_content, branding_info):
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()

    # ---------- COVER PAGE ----------
    doc.add_paragraph("\n" * 3)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(branding_info['solution_name'])
    run.bold = True
    run.font.size = Pt(26)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run("Scope of Work Document").font.size = Pt(14)

    doc.add_page_break()

    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    lines = text_content.split("\n")
    i = 0

    while i < len(lines):
        line = lines[i].strip()
        if not line:
            doc.add_paragraph("")
            i += 1
            continue

        clean_check = line.replace("#", "").upper()
        line_clean = line.replace("**", "").replace("*", "")

        # ======================================================
        # INSERT ARCHITECTURE IMAGE AFTER SCOPE OF WORK (ADDED)
        # ======================================================
        if "3 SCOPE OF WORK" in clean_check:
            doc.add_heading(line_clean, level=1)

            arch_path = branding_info.get("architecture_image_path")
            if arch_path:
                try:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(arch_path, width=Inches(5.8))

                    cap = doc.add_paragraph("Figure: High Level Solution Architecture")
                    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except:
                    doc.add_paragraph("Architecture diagram could not be loaded.")

            i += 1
            continue

        # ---------- STANDARD PARSING ----------
        if line.startswith("# "):
            doc.add_heading(line_clean[2:], level=1)
        elif line.startswith("## "):
            doc.add_heading(line_clean[3:], level=2)
        elif line.startswith("### "):
            doc.add_heading(line_clean[4:], level=3)
        elif line.startswith("- "):
            doc.add_paragraph(line_clean[2:], style="List Bullet")
        else:
            doc.add_paragraph(line_clean)

        i += 1

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ======================================================
# SESSION STATE
# ======================================================
if "generated_sow" not in st.session_state:
    st.session_state.generated_sow = ""

# ======================================================
# SIDEBAR
# ======================================================
with st.sidebar:
    st.title("SOW Architect")

    api_key = st.text_input("Gemini API Key", type="password")

    solution_options = list(ARCHITECTURE_IMAGE_MAP.keys())
    solution_type = st.selectbox("Solution Type", solution_options)
    final_solution = solution_type

# ======================================================
# MAIN UI
# ======================================================
st.title("🚀 GenAI Scope of Work Architect")

objective = st.text_area("Business Objective")

if st.button("Generate SOW"):
    st.session_state.generated_sow = f"""
# 1 TABLE OF CONTENTS

# 2 PROJECT OVERVIEW
## 2.1 OBJECTIVE
{objective}

# 3 SCOPE OF WORK – TECHNICAL PROJECT PLAN
Detailed scope content here...

# 4 SOLUTION ARCHITECTURE / ARCHITECTURAL DIAGRAM
"""

if st.session_state.generated_sow:
    st.text_area("Edit SOW", st.session_state.generated_sow, height=450)

    if st.button("Download DOCX"):
        branding_info = {
            "solution_name": final_solution,
            "architecture_image_path": ARCHITECTURE_IMAGE_MAP.get(final_solution)
        }

        docx = create_docx_logic(st.session_state.generated_sow, branding_info)

        st.download_button(
            "📥 Download Word Document",
            docx,
            file_name=f"SOW_{final_solution.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
