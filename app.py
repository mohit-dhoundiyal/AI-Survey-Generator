import io
import json
import os
import re

import streamlit as st
from docx import Document
from openpyxl import Workbook, load_workbook
from dotenv import load_dotenv

from google import genai
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from streamlit_oauth import OAuth2Component


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="AI Survey Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# MODERN STYLING (adaptive light & dark mode)
# ============================================================

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap');

        /* ---- Global Typography & Smooth Rendering ---- */
        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
        }

        /* ---- Remove white bar at the top (Streamlit header) ---- */
        header[data-testid="stHeader"] {
            background: transparent !important;
            color: var(--text-color) !important;
        }
        header[data-testid="stHeader"] * {
            color: var(--text-color) !important;
        }
        header[data-testid="stHeader"] svg {
            fill: var(--text-color) !important;
        }

        /* ---- Ambient adaptive background (bound to Streamlit's active theme) ---- */
        .stApp {
            background-color: var(--background-color) !important;
            background-image: 
                radial-gradient(circle at 10% 8%, rgba(124, 58, 237, 0.08) 0px, transparent 45%),
                radial-gradient(circle at 90% 92%, rgba(219, 39, 119, 0.07) 0px, transparent 45%) !important;
            background-attachment: fixed;
            color: var(--text-color) !important;
        }

        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 3.5rem;
            max-width: 1140px;
        }

        /* ---- Hero title banner ---- */
        .hero-banner {
            background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 35%, #BE185D 75%, #EA580C 100%);
            padding: clamp(1.4rem, 3.8vw, 2.25rem) clamp(1.3rem, 4vw, 2.5rem);
            border-radius: 20px;
            margin-bottom: 1.8rem;
            border: 1px solid rgba(255, 255, 255, 0.22);
            box-shadow: 0 16px 36px -6px rgba(109, 40, 217, 0.38), inset 0 1px 0 rgba(255, 255, 255, 0.3);
            position: relative;
            overflow: hidden;
        }
        .hero-banner::after {
            content: '';
            position: absolute;
            top: -45%;
            right: -10%;
            width: 320px;
            height: 320px;
            background: radial-gradient(circle, rgba(255, 255, 255, 0.18) 0%, transparent 70%);
            pointer-events: none;
        }
        .hero-badge {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 5px 14px;
            background: rgba(255, 255, 255, 0.18);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border: 1px solid rgba(255, 255, 255, 0.3);
            border-radius: 9999px;
            color: #FFFFFF !important;
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.05em;
            text-transform: uppercase;
            margin-bottom: 12px;
        }
        .hero-badge-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background-color: #34D399;
            box-shadow: 0 0 8px #34D399;
        }
        .hero-title {
            color: #FFFFFF !important;
            font-size: clamp(1.5rem, 4.5vw, 2.25rem) !important;
            font-weight: 800 !important;
            letter-spacing: -0.02em;
            margin: 0 0 10px 0 !important;
            line-height: 1.25 !important;
        }
        .hero-subtitle {
            color: rgba(255, 255, 255, 0.94) !important;
            font-size: clamp(0.92rem, 2.2vw, 1.05rem) !important;
            font-weight: 400 !important;
            margin: 0 !important;
            max-width: 780px;
            line-height: 1.55 !important;
        }

        /* ---- Headings ---- */
        h1, h2, h3, h4, [data-testid="stMarkdownContainer"] h2, [data-testid="stMarkdownContainer"] h3 {
            color: var(--text-color) !important;
            font-weight: 700 !important;
            letter-spacing: -0.015em;
        }
        [data-testid="stCaptionContainer"],
        [data-testid="stCaptionContainer"] p,
        [data-testid="stMarkdownContainer"] small {
            color: color-mix(in srgb, var(--text-color) 75%, transparent) !important;
        }

        /* ---- Bordered containers (Cards) ---- */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 16px !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 15%, transparent) !important;
            background: var(--secondary-background-color) !important;
            box-shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.12) !important;
            transition: border-color 0.2s ease, box-shadow 0.2s ease;
        }
        div[data-testid="stVerticalBlockBorderWrapper"]:hover {
            border-color: #7C3AED !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] > div {
            padding: clamp(1.1rem, 2.5vw, 1.5rem) !important;
        }

        /* ---- File uploader dropzone (Adaptive to dark/light) ---- */
        div[data-testid="stFileUploaderDropzone"] {
            border-radius: 14px !important;
            border: 2px dashed color-mix(in srgb, var(--text-color) 25%, transparent) !important;
            background: color-mix(in srgb, var(--text-color) 4%, var(--secondary-background-color)) !important;
            transition: all 0.2s ease;
            padding: 1.2rem 1rem !important;
        }
        div[data-testid="stFileUploaderDropzone"]:hover {
            border-color: #7C3AED !important;
            background: color-mix(in srgb, #7C3AED 8%, var(--secondary-background-color)) !important;
        }
        div[data-testid="stFileUploaderDropzone"] [data-testid="stMarkdownContainer"] p,
        div[data-testid="stFileUploaderDropzone"] span,
        div[data-testid="stFileUploaderDropzone"] small,
        div[data-testid="stFileUploaderDropzone"] div {
            color: var(--text-color) !important;
        }
        div[data-testid="stFileUploaderDropzone"] button {
            border-radius: 10px !important;
            background: color-mix(in srgb, var(--text-color) 10%, var(--secondary-background-color)) !important;
            color: var(--text-color) !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 22%, transparent) !important;
        }
        div[data-testid="stFileUploaderDropzoneInstructions"] * {
            color: var(--text-color) !important;
        }

        /* ---- File loaded badge ---- */
        .file-loaded-badge {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 8px 14px;
            border-radius: 10px;
            background: color-mix(in srgb, #10B981 18%, var(--secondary-background-color)) !important;
            color: #10B981 !important;
            border: 1px solid color-mix(in srgb, #10B981 35%, transparent) !important;
            font-size: 0.84rem;
            font-weight: 600;
            margin-top: 10px;
            word-break: break-all;
        }
        .file-loaded-badge span, .file-loaded-badge strong {
            color: #10B981 !important;
        }

        /* ---- Form Inputs, Textareas & Selects ---- */
        div[data-baseweb="input"], div[data-baseweb="textarea"], div[data-baseweb="select"] {
            border-radius: 10px !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 20%, transparent) !important;
            background-color: color-mix(in srgb, var(--text-color) 3%, var(--secondary-background-color)) !important;
            color: var(--text-color) !important;
            transition: border-color 0.2s ease, box-shadow 0.2s ease !important;
        }
        div[data-baseweb="input"]:focus-within, div[data-baseweb="textarea"]:focus-within, div[data-baseweb="select"]:focus-within {
            border-color: #7C3AED !important;
            box-shadow: 0 0 0 3px rgba(124, 58, 237, 0.25) !important;
        }
        div[data-baseweb="input"] input, div[data-baseweb="textarea"] textarea {
            color: var(--text-color) !important;
            background-color: transparent !important;
            font-family: inherit !important;
        }
        div[data-baseweb="select"] * {
            color: var(--text-color) !important;
        }
        label[data-testid="stWidgetLabel"] p {
            color: var(--text-color) !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
        }

        /* ---- Expanders (Question Cards) ---- */
        div[data-testid="stExpander"] {
            border: 1px solid color-mix(in srgb, var(--text-color) 15%, transparent) !important;
            border-radius: 14px !important;
            background: var(--secondary-background-color) !important;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.05);
            margin-bottom: 12px;
            overflow: hidden;
            transition: border-color 0.2s ease, box-shadow 0.2s ease;
        }
        div[data-testid="stExpander"]:hover {
            border-color: #7C3AED !important;
            box-shadow: 0 4px 16px rgba(124, 58, 237, 0.12);
        }
        div[data-testid="stExpander"] details summary {
            padding: 12px 16px !important;
            border-radius: 14px !important;
            background: transparent !important;
        }
        div[data-testid="stExpander"] details summary p {
            font-size: 0.98rem !important;
            font-weight: 600 !important;
            color: var(--text-color) !important;
        }
        div[data-testid="stExpander"] details summary svg {
            fill: var(--text-color) !important;
        }
        div[data-testid="stExpander"] details[open] summary {
            border-bottom: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent) !important;
            border-bottom-left-radius: 0 !important;
            border-bottom-right-radius: 0 !important;
        }

        /* ---- Buttons: Primary, Secondary, Download ---- */
        button[data-testid="baseButton-primary"] {
            border-radius: 12px !important;
            border: none !important;
            background: linear-gradient(135deg, #6366F1 0%, #7C3AED 50%, #9333EA 100%) !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            letter-spacing: 0.01em;
            padding: 0.65rem 1.4rem !important;
            min-height: 48px !important;
            box-shadow: 0 4px 18px -2px rgba(124, 58, 237, 0.45) !important;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }
        button[data-testid="baseButton-primary"]:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 8px 24px -2px rgba(124, 58, 237, 0.6) !important;
            filter: brightness(1.08);
        }
        button[data-testid="baseButton-primary"]:active {
            transform: translateY(0px) !important;
        }
        button[data-testid="baseButton-primary"] p,
        button[data-testid="baseButton-primary"] span,
        button[data-testid="baseButton-primary"] div {
            color: #FFFFFF !important;
            font-weight: 700 !important;
        }

        button[data-testid="baseButton-secondary"] {
            border-radius: 12px !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 18%, transparent) !important;
            background: color-mix(in srgb, var(--text-color) 4%, var(--secondary-background-color)) !important;
            color: var(--text-color) !important;
            font-weight: 600 !important;
            min-height: 44px !important;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04) !important;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }
        button[data-testid="baseButton-secondary"]:hover {
            border-color: #7C3AED !important;
            color: #7C3AED !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 4px 12px rgba(124, 58, 237, 0.15) !important;
        }
        button[data-testid="baseButton-secondary"]:active {
            transform: translateY(0px) !important;
        }
        button[data-testid="baseButton-secondary"] p,
        button[data-testid="baseButton-secondary"] span {
            color: inherit !important;
        }

        .stDownloadButton > button {
            border-radius: 12px !important;
            border: none !important;
            background: linear-gradient(135deg, #F59E0B 0%, #EA580C 100%) !important;
            color: #FFFFFF !important;
            font-weight: 700 !important;
            min-height: 48px !important;
            box-shadow: 0 4px 16px -2px rgba(234, 88, 12, 0.4) !important;
            transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        }
        .stDownloadButton > button:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 8px 24px -2px rgba(234, 88, 12, 0.55) !important;
            filter: brightness(1.08);
        }
        .stDownloadButton > button p,
        .stDownloadButton > button span {
            color: #FFFFFF !important;
            font-weight: 700 !important;
        }

        /* ---- Sidebar Styling ---- */
        section[data-testid="stSidebar"] {
            background: var(--secondary-background-color) !important;
            border-right: 1px solid color-mix(in srgb, var(--text-color) 12%, transparent) !important;
        }
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3,
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] strong,
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] b {
            color: var(--text-color) !important;
        }
        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        section[data-testid="stSidebar"] [data-testid="stCaptionContainer"],
        section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {
            color: color-mix(in srgb, var(--text-color) 78%, transparent) !important;
        }
        section[data-testid="stSidebar"] hr {
            border-color: color-mix(in srgb, var(--text-color) 14%, transparent) !important;
        }

        /* ---- Progress step pills ---- */
        .step-pill {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 14px;
            border-radius: 12px;
            font-size: 0.85rem;
            font-weight: 600;
            margin-bottom: 8px;
            transition: all 0.2s ease;
        }
        .step-done {
            background-color: color-mix(in srgb, #10B981 18%, var(--secondary-background-color)) !important;
            color: #10B981 !important;
            border: 1px solid color-mix(in srgb, #10B981 38%, transparent) !important;
        }
        .step-done span {
            color: #10B981 !important;
        }
        .step-current {
            background-color: color-mix(in srgb, #7C3AED 24%, var(--secondary-background-color)) !important;
            color: #A78BFA !important;
            border: 1px solid #7C3AED !important;
            box-shadow: 0 0 14px rgba(124, 58, 237, 0.25) !important;
        }
        .step-current span {
            color: #A78BFA !important;
        }
        .step-todo {
            background-color: color-mix(in srgb, var(--text-color) 6%, var(--secondary-background-color)) !important;
            color: color-mix(in srgb, var(--text-color) 75%, transparent) !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 14%, transparent) !important;
        }
        .step-todo span {
            color: color-mix(in srgb, var(--text-color) 75%, transparent) !important;
        }

        /* ---- Metric / Survey stats ---- */
        div[data-testid="stMetric"] {
            background: var(--secondary-background-color) !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 15%, transparent) !important;
            border-radius: 14px !important;
            padding: 12px 16px !important;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04) !important;
        }
        div[data-testid="stMetric"] label,
        div[data-testid="stMetric"] [data-testid="stMetricLabel"] p {
            color: color-mix(in srgb, var(--text-color) 75%, transparent) !important;
            font-size: 0.82rem !important;
            font-weight: 500 !important;
        }
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
            color: #7C3AED !important;
            font-weight: 800 !important;
            font-size: 1.6rem !important;
        }

        /* ---- Tags for question type ---- */
        .qtype-tag {
            display: inline-flex;
            align-items: center;
            padding: 4px 12px;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 700;
            letter-spacing: 0.02em;
            margin-left: 6px;
        }
        .qtype-mc {
            background-color: color-mix(in srgb, #3B82F6 18%, var(--secondary-background-color)) !important;
            color: #60A5FA !important;
            border: 1px solid color-mix(in srgb, #3B82F6 38%, transparent) !important;
        }
        .qtype-short {
            background-color: color-mix(in srgb, #EC4899 18%, var(--secondary-background-color)) !important;
            color: #F472B6 !important;
            border: 1px solid color-mix(in srgb, #EC4899 38%, transparent) !important;
        }
        .qtype-para {
            background-color: color-mix(in srgb, #F59E0B 18%, var(--secondary-background-color)) !important;
            color: #FBBF24 !important;
            border: 1px solid color-mix(in srgb, #F59E0B 38%, transparent) !important;
        }

        /* ---- Alerts & Status ---- */
        div[data-testid="stAlert"] {
            border-radius: 12px !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 15%, transparent) !important;
        }
        div[data-testid="stStatus"] {
            border-radius: 14px !important;
            border: 1px solid color-mix(in srgb, var(--text-color) 15%, transparent) !important;
            background: var(--secondary-background-color) !important;
        }

        /* ============================================================
           MOBILE RESPONSIVENESS (< 768px)
           ============================================================ */
        @media (max-width: 768px) {
            .block-container {
                padding-top: 1rem !important;
                padding-bottom: 2.5rem !important;
                padding-left: 0.85rem !important;
                padding-right: 0.85rem !important;
            }
            .hero-banner {
                padding: 1.3rem 1.1rem !important;
                border-radius: 16px !important;
                margin-bottom: 1.25rem !important;
            }
            .hero-title {
                font-size: 1.45rem !important;
            }
            .hero-subtitle {
                font-size: 0.88rem !important;
                line-height: 1.45 !important;
            }
            /* Stack columns on small viewports so fields don't get squished */
            div[data-testid="stHorizontalBlock"] {
                flex-wrap: wrap !important;
                gap: 0.85rem !important;
            }
            div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
                min-width: 100% !important;
                flex: 1 1 100% !important;
            }
            /* Mobile touch friendly buttons */
            button[data-testid="baseButton-primary"],
            button[data-testid="baseButton-secondary"],
            .stDownloadButton > button {
                width: 100% !important;
                min-height: 48px !important;
                margin-bottom: 0.5rem !important;
            }
            /* Prevent expander header overflow on phones */
            div[data-testid="stExpander"] details summary p {
                font-size: 0.9rem !important;
                white-space: nowrap !important;
                overflow: hidden !important;
                text-overflow: ellipsis !important;
                max-width: 80vw !important;
            }
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# CONSTANTS
# ============================================================

RATING_OPTIONS = [
    "Excellent",
    "Very Good",
    "Good",
    "Fair",
    "Poor"
]

FORMS_SCOPE = "https://www.googleapis.com/auth/forms.body"

FIXED_PARTICIPANT_QUESTIONS = [
    {
        "section": "Participant Information",
        "session_no": "",
        "session": "",
        "category": "Participant Information",
        "question": "Full Name",
        "question_type": "Short Answer",
        "options": [],
        "required": True
    },
    {
        "section": "Participant Information",
        "session_no": "",
        "session": "",
        "category": "Participant Information",
        "question": "Designation",
        "question_type": "Short Answer",
        "options": [],
        "required": True
    },
    {
        "section": "Participant Information",
        "session_no": "",
        "session": "",
        "category": "Participant Information",
        "question": "Organisation/Department",
        "question_type": "Short Answer",
        "options": [],
        "required": True
    }
]


# ============================================================
# GEMINI CLIENT
# ============================================================

def get_gemini_client():
    api_key = None

    # Streamlit Cloud / Streamlit secrets
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
    except Exception:
        pass

    # Local .env
    if not api_key:
        load_dotenv()
        api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY was not found. "
            "Add it to your .env file for local use."
        )

    return genai.Client(api_key=api_key)


# ============================================================
# DOCX EXTRACTION
# ============================================================

def extract_docx_text(uploaded_file):
    uploaded_file.seek(0)

    document = Document(uploaded_file)

    parts = []

    # Normal paragraphs
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()

        if text:
            parts.append(text)

    # Tables
    for table in document.tables:

        for row in table.rows:

            cells = []

            for cell in row.cells:
                cell_text = cell.text.strip()

                if cell_text:
                    cells.append(cell_text)

            if cells:
                parts.append(" | ".join(cells))

    return "\n".join(parts)


# ============================================================
# XLSX EXTRACTION
# ============================================================

def extract_xlsx_text(uploaded_file):
    uploaded_file.seek(0)

    workbook = load_workbook(
        uploaded_file,
        data_only=True
    )

    parts = []

    for worksheet in workbook.worksheets:

        parts.append(
            f"Sheet: {worksheet.title}"
        )

        for row in worksheet.iter_rows(
            values_only=True
        ):

            values = []

            for value in row:

                if value is not None:
                    values.append(
                        str(value).strip()
                    )

            if values:
                parts.append(
                    " | ".join(values)
                )

    return "\n".join(parts)


# ============================================================
# EXAMPLE FORM INSTRUCTION
# ============================================================

def build_example_instruction(example_text):

    if not example_text.strip():

        return """
No example feedback form was provided.

Use a professional workshop feedback style.
"""

    return f"""
An existing feedback form has been provided only as a
STYLE AND STRUCTURE REFERENCE.

Use it to understand:
- professional wording
- rating style
- feedback structure
- overall presentation

Do NOT copy participant information questions.
Do NOT copy questions blindly.
The workshop agenda is the PRIMARY source.

Example feedback form:

-------------------------
{example_text}
-------------------------
"""


# ============================================================
# GENERATE SURVEY USING GEMINI
# ============================================================

def generate_feedback_form(
    agenda_text,
    example_text=""
):

    client = get_gemini_client()

    example_instruction = build_example_instruction(
        example_text
    )

    prompt = f"""
You are an expert workshop feedback survey designer.

Create ONE complete feedback survey from the workshop agenda.

IMPORTANT RULES:

1. DO NOT generate these participant-information questions:
   - Full Name
   - Designation
   - Organisation/Department

   The application adds those separately.

2. Identify EVERY actual SESSION in the agenda.

3. Create EXACTLY ONE feedback question for EACH session.

4. Preserve the original session order.

5. Preserve the session number.

6. Preserve the session title/name.

7. Each session question must be specifically related to:
   - the session title
   - topics covered
   - activities
   - practical work
   - tools or concepts covered

8. Avoid generic questions such as:
   "How satisfied were you with the session?"

   Instead, ask about the actual learning/content of that session.

9. Every session question must be:
   Multiple Choice

10. Every session question must contain EXACTLY these options:

   Excellent
   Very Good
   Good
   Fair
   Poor

11. Every session question must be required.

12. Do NOT create more than one question per session.

13. Do NOT invent sessions.

14. After all session questions, create EXACTLY ONE overall workshop rating:

   Question:
   How would you rate the overall workshop?

   Type:
   Multiple Choice

   Options:
   Excellent
   Very Good
   Good
   Fair
   Poor

   Required:
   true

15. After the overall rating, create EXACTLY ONE suggestions question:

   Question:
   Please provide your suggestions for improving future workshops.

   Type:
   Paragraph

   Required:
   false

16. Return ONLY valid JSON.

Use this JSON structure:

{{
    "questions": [
        {{
            "section": "Day 1",
            "session_no": "1",
            "session": "Session title",
            "category": "Session Feedback",
            "question": "Specific session-based feedback question",
            "question_type": "Multiple Choice",
            "options": [
                "Excellent",
                "Very Good",
                "Good",
                "Fair",
                "Poor"
            ],
            "required": true
        }}
    ]
}}

Workshop agenda:

-------------------------
{agenda_text}
-------------------------

{example_instruction}
"""

    response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=prompt
    )

    response_text = response.text.strip()

    # Remove markdown fences if Gemini returns them
    response_text = re.sub(
        r"^```json\s*",
        "",
        response_text,
        flags=re.IGNORECASE
    )

    response_text = re.sub(
        r"^```\s*",
        "",
        response_text
    )

    response_text = re.sub(
        r"\s*```$",
        "",
        response_text
    )

    parsed = json.loads(response_text)

    return parsed.get("questions", [])


# ============================================================
# NORMALIZE QUESTIONS
# ============================================================

def normalize_questions(ai_questions):

    final_questions = []

    # --------------------------------------------------------
    # Add fixed participant questions exactly once
    # --------------------------------------------------------

    for fixed_question in FIXED_PARTICIPANT_QUESTIONS:

        final_questions.append(
            fixed_question.copy()
        )

    # --------------------------------------------------------
    # Add AI-generated questions
    # --------------------------------------------------------

    for item in ai_questions:

        if not isinstance(item, dict):
            continue

        question_text = str(
            item.get("question", "")
        ).strip()

        if not question_text:
            continue

        lower_question = question_text.lower()

        # Prevent AI from creating participant questions
        blocked_terms = [
            "full name",
            "designation",
            "organisation/department",
            "organization/department"
        ]

        if any(
            term in lower_question
            for term in blocked_terms
        ):
            continue

        question_type = str(
            item.get(
                "question_type",
                "Multiple Choice"
            )
        ).strip()

        if question_type == "Multiple Choice":

            options = RATING_OPTIONS.copy()

        else:

            raw_options = item.get(
                "options",
                []
            )

            if isinstance(raw_options, list):
                options = raw_options
            else:
                options = []

        normalized_item = {
            "section": str(
                item.get("section", "")
            ).strip(),

            "session_no": str(
                item.get("session_no", "")
            ).strip(),

            "session": str(
                item.get("session", "")
            ).strip(),

            "category": str(
                item.get(
                    "category",
                    "Session Feedback"
                )
            ).strip(),

            "question": question_text,

            "question_type": question_type,

            "options": options,

            "required": bool(
                item.get("required", True)
            )
        }

        final_questions.append(
            normalized_item
        )

    # --------------------------------------------------------
    # Overall workshop rating
    # --------------------------------------------------------

    overall_text = (
        "How would you rate the overall workshop?"
    )

    overall_exists = any(
        str(q.get("question", "")).strip().lower()
        == overall_text.lower()
        for q in final_questions
    )

    if not overall_exists:

        final_questions.append(
            {
                "section": "Overall Feedback",
                "session_no": "",
                "session": "",
                "category": "Overall Workshop Rating",
                "question": overall_text,
                "question_type": "Multiple Choice",
                "options": RATING_OPTIONS.copy(),
                "required": True
            }
        )

    # --------------------------------------------------------
    # Suggestions
    # --------------------------------------------------------

    suggestions_text = (
        "Please provide your suggestions for improving future workshops."
    )

    suggestions_exists = any(
        str(q.get("question", "")).strip().lower()
        == suggestions_text.lower()
        for q in final_questions
    )

    if not suggestions_exists:

        final_questions.append(
            {
                "section": "Overall Feedback",
                "session_no": "",
                "session": "",
                "category": "Suggestions",
                "question": suggestions_text,
                "question_type": "Paragraph",
                "options": [],
                "required": False
            }
        )

    return final_questions


# ============================================================
# EXCEL CREATION
# ============================================================

def create_excel(questions):

    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = "Feedback Form"

    headers = [
        "No.",
        "Section",
        "Session No.",
        "Session",
        "Category",
        "Question",
        "Question Type",
        "Options",
        "Required"
    ]

    worksheet.append(headers)

    for index, question in enumerate(
        questions,
        start=1
    ):

        options = question.get(
            "options",
            []
        )

        if isinstance(options, list):

            options_text = " | ".join(
                str(option)
                for option in options
            )

        else:

            options_text = str(options)

        worksheet.append(
            [
                index,
                question.get(
                    "section",
                    ""
                ),
                question.get(
                    "session_no",
                    ""
                ),
                question.get(
                    "session",
                    ""
                ),
                question.get(
                    "category",
                    ""
                ),
                question.get(
                    "question",
                    ""
                ),
                question.get(
                    "question_type",
                    ""
                ),
                options_text,
                question.get(
                    "required",
                    True
                )
            ]
        )

    # Header formatting
    for cell in worksheet[1]:

        cell.font = cell.font.copy(
            bold=True
        )

    worksheet.freeze_panes = "A2"

    column_widths = {
        "A": 8,
        "B": 22,
        "C": 15,
        "D": 40,
        "E": 28,
        "F": 70,
        "G": 20,
        "H": 50,
        "I": 12
    }

    for column, width in column_widths.items():

        worksheet.column_dimensions[
            column
        ].width = width

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return output.getvalue()


# ============================================================
# GOOGLE OAUTH COMPONENT
# ============================================================

def get_google_oauth_component():

    # IMPORTANT:
    # Reuses the SAME client credentials already present
    # inside the user's [auth] secrets section.

    client_id = st.secrets[
        "auth"
    ][
        "client_id"
    ]

    client_secret = st.secrets[
        "auth"
    ][
        "client_secret"
    ]

    authorize_url = (
        "https://accounts.google.com/o/oauth2/v2/auth"
    )

    token_url = (
        "https://oauth2.googleapis.com/token"
    )

    refresh_token_url = (
        "https://oauth2.googleapis.com/token"
    )

    revoke_token_url = None

    return OAuth2Component(
        client_id,
        client_secret,
        authorize_url,
        token_url,
        refresh_token_url,
        revoke_token_url
    )


# ============================================================
# GOOGLE FORM CREATION
# ============================================================

def create_google_form(
    title,
    questions,
    access_token
):

    credentials = Credentials(
        token=access_token,
        scopes=[FORMS_SCOPE]
    )

    service = build(
        "forms",
        "v1",
        credentials=credentials,
        cache_discovery=False
    )

    # --------------------------------------------------------
    # 1. Create blank form
    # --------------------------------------------------------

    form = service.forms().create(
        body={
            "info": {
                "title": title
            }
        }
    ).execute()

    form_id = form["formId"]

    # --------------------------------------------------------
    # 2. Add questions
    # --------------------------------------------------------

    requests = []

    for index, question in enumerate(
        questions
    ):

        question_text = str(
            question.get(
                "question",
                ""
            )
        ).strip()

        question_type = question.get(
            "question_type",
            "Multiple Choice"
        )

        required = bool(
            question.get(
                "required",
                True
            )
        )

        options = question.get(
            "options",
            []
        )

        # -----------------------------------------------
        # Short Answer
        # -----------------------------------------------

        if question_type == "Short Answer":

            requests.append(
                {
                    "createItem": {
                        "item": {
                            "title": question_text,
                            "questionItem": {
                                "question": {
                                    "required": required,
                                    "textQuestion": {}
                                }
                            }
                        },
                        "location": {
                            "index": index
                        }
                    }
                }
            )

        # -----------------------------------------------
        # Paragraph
        # -----------------------------------------------

        elif question_type == "Paragraph":

            requests.append(
                {
                    "createItem": {
                        "item": {
                            "title": question_text,
                            "questionItem": {
                                "question": {
                                    "required": required,
                                    "textQuestion": {
                                        "paragraph": True
                                    }
                                }
                            }
                        },
                        "location": {
                            "index": index
                        }
                    }
                }
            )

        # -----------------------------------------------
        # Multiple Choice
        # -----------------------------------------------

        else:

            choice_options = []

            for option in options:

                choice_options.append(
                    {
                        "value": str(option)
                    }
                )

            requests.append(
                {
                    "createItem": {
                        "item": {
                            "title": question_text,
                            "questionItem": {
                                "question": {
                                    "required": required,
                                    "choiceQuestion": {
                                        "type": "RADIO",
                                        "options": choice_options
                                    }
                                }
                            }
                        },
                        "location": {
                            "index": index
                        }
                    }
                }
            )

    # --------------------------------------------------------
    # 3. Send all questions
    # --------------------------------------------------------

    if requests:

        service.forms().batchUpdate(
            formId=form_id,
            body={
                "requests": requests
            }
        ).execute()

    return form_id


# ============================================================
# SESSION STATE
# ============================================================

if "questions" not in st.session_state:
    st.session_state.questions = []

if "generated" not in st.session_state:
    st.session_state.generated = False

if "google_form_token" not in st.session_state:
    st.session_state.google_form_token = None


# ============================================================
# SIDEBAR — STATUS / NAVIGATION (display-only, no new logic)
# ============================================================

with st.sidebar:

    st.markdown("## 📋 AI Survey Generator")
    st.caption("Workshop agenda → feedback survey → Excel / Google Form")

    st.divider()

    st.markdown("**Progress Tracker**")

    step1_done = st.session_state.generated
    step2_done = bool(st.session_state.questions) and step1_done
    step3_done = st.session_state.google_form_token is not None

    def pill(label, done, current=False):
        css_class = "step-done" if done else ("step-current" if current else "step-todo")
        icon = "✅" if done else ("➡️" if current else "◻️")
        st.markdown(
            f'<div class="step-pill {css_class}">{icon} <span>{label}</span></div>',
            unsafe_allow_html=True
        )

    pill("1. Upload agenda", step1_done, current=not step1_done)
    pill("2. Generate survey", step1_done, current=step1_done and not step2_done)
    pill("3. Review & edit", step2_done, current=step2_done and not step3_done)
    pill("4. Export (Excel / Forms)", step2_done)

    st.divider()

    if st.session_state.questions:
        st.metric("Total Questions", len(st.session_state.questions))

    st.divider()
    st.markdown(
        """
        <div style="font-size: 0.83rem; color: var(--text-color); line-height: 1.5; padding: 12px 14px; background: color-mix(in srgb, var(--text-color) 7%, var(--secondary-background-color)); border-radius: 12px; border: 1px solid color-mix(in srgb, var(--text-color) 14%, transparent);">
            💡 <strong>Pro-Tip:</strong> Uploading an existing feedback form (.xlsx) in Step 2 
            guides Gemini to mirror your team's preferred rating scale and tone.
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# HEADER & HERO BANNER
# ============================================================

st.markdown(
    """
    <div class="hero-banner">
        <div class="hero-badge">
            <span class="hero-badge-dot"></span>
            AI-Powered Survey Engine
        </div>
        <h1 class="hero-title">📋 AI Survey Generator</h1>
        <p class="hero-subtitle">
            Turn workshop agendas into comprehensive, tailored feedback surveys in seconds. 
            Review and fine-tune questions, then export directly to Excel or Google Forms ✨
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")


# ============================================================
# STEP 1 & 2 — INPUTS (Responsive side-by-side / stacked)
# ============================================================

input_col1, input_col2 = st.columns(2, gap="large")

with input_col1:

    with st.container(border=True):

        st.markdown("### 📌 1. Workshop Agenda")
        st.caption("Required · DOCX format")

        agenda_file = st.file_uploader(
            "Upload agenda DOCX",
            type=["docx"],
            key="agenda_file",
            help="Required. The agenda drives the session-by-session questions."
        )

        if agenda_file is not None:
            size_kb = round(len(agenda_file.getvalue()) / 1024, 1)
            st.markdown(
                f"""
                <div class="file-loaded-badge">
                    <span>✅</span>
                    <span><strong>{agenda_file.name}</strong> ({size_kb} KB loaded)</span>
                </div>
                """,
                unsafe_allow_html=True
            )

with input_col2:

    with st.container(border=True):

        st.markdown("### 💡 2. Reference Form")
        st.caption("Optional · XLSX format")

        example_file = st.file_uploader(
            "Upload an existing feedback XLSX for style/reference (optional)",
            type=["xlsx"],
            key="example_file",
            help="Optional. Used only as a style reference, not copied verbatim."
        )

        if example_file is not None:
            size_kb = round(len(example_file.getvalue()) / 1024, 1)
            st.markdown(
                f"""
                <div class="file-loaded-badge">
                    <span>✅</span>
                    <span><strong>{example_file.name}</strong> ({size_kb} KB loaded)</span>
                </div>
                """,
                unsafe_allow_html=True
            )


# ============================================================
# GENERATE SURVEY CTA BUTTON
# ============================================================

st.write("")

generate_clicked = st.button(
    "✨ Analyze Agenda & Generate Survey",
    type="primary",
    use_container_width=True
)

if generate_clicked:

    if agenda_file is None:

        st.error(
            "⚠️ Please upload the workshop agenda DOCX in Step 1 before generating."
        )

    else:

        try:

            with st.status(
                "🚀 Generating your feedback survey with AI...",
                expanded=True
            ) as status:

                st.write("📖 Reading agenda DOCX...")

                agenda_text = extract_docx_text(
                    agenda_file
                )

                example_text = ""

                if example_file is not None:

                    st.write("📖 Reading example feedback form...")

                    example_text = extract_xlsx_text(
                        example_file
                    )

                st.write("🤖 Asking Gemini to craft tailored questions...")

                ai_questions = generate_feedback_form(
                    agenda_text,
                    example_text
                )

                st.write("🧹 Normalizing and validating questions...")

                final_questions = normalize_questions(
                    ai_questions
                )

                st.session_state.questions = (
                    final_questions
                )

                st.session_state.generated = True

                status.update(
                    label=f"Survey generated successfully ({len(final_questions)} questions)!",
                    state="complete",
                    expanded=False
                )

            st.success(
                f"🎉 Survey generated successfully with **{len(final_questions)}** questions! "
                "Scroll down to Step 3 to review and customize."
            )

        except Exception as error:

            st.error(
                "Failed to generate survey."
            )

            st.exception(error)


# ============================================================
# STEP 3 — REVIEW & EDIT SURVEY
# ============================================================

if st.session_state.generated:

    st.divider()

    st.markdown("## 3. Review & Edit Survey")

    st.caption(
        "Expand any question to edit wording, question type, answer choices, or session metadata. "
        "Participant-information questions are added automatically and can be edited below."
    )

    questions = st.session_state.questions

    # --------------------------------------------------------
    # Questions Summary Stats
    # --------------------------------------------------------
    total_q = len(questions)
    mc_q = sum(1 for q in questions if q.get("question_type") == "Multiple Choice")
    sa_q = sum(1 for q in questions if q.get("question_type") == "Short Answer")
    para_q = sum(1 for q in questions if q.get("question_type") == "Paragraph")

    stat_c1, stat_c2, stat_c3, stat_c4 = st.columns(4)
    with stat_c1:
        st.metric("Total Questions", total_q)
    with stat_c2:
        st.metric("🔘 Multiple Choice", mc_q)
    with stat_c3:
        st.metric("✏️ Short Answer", sa_q)
    with stat_c4:
        st.metric("📝 Paragraph", para_q)

    st.write("")

    delete_index = None

    # --------------------------------------------------------
    # Existing questions inside modern expanders
    # --------------------------------------------------------

    for index, question in enumerate(
        questions
    ):

        preview = question.get("question", "").strip() or "(empty question)"
        type_label = question.get("question_type", "Multiple Choice")
        emoji_map = {
            "Multiple Choice": "🔘",
            "Short Answer": "✏️",
            "Paragraph": "📝"
        }
        tag_class_map = {
            "Multiple Choice": "qtype-mc",
            "Short Answer": "qtype-short",
            "Paragraph": "qtype-para"
        }
        emoji = emoji_map.get(type_label, "❓")
        expander_title = f"{emoji} Q{index + 1}: {preview}"

        with st.expander(expander_title, expanded=False):

            # Header row: Type tag & Delete button
            header_col1, header_col2 = st.columns(
                [3, 1]
            )

            with header_col1:
                tag_class = tag_class_map.get(type_label, "qtype-mc")
                st.markdown(
                    f'Question Type: <span class="qtype-tag {tag_class}">{type_label}</span>',
                    unsafe_allow_html=True
                )

            with header_col2:
                if st.button(
                    "🗑️ Delete Question",
                    key=f"delete_{index}",
                    use_container_width=True
                ):
                    delete_index = index

            # Question Text prompt
            question_text = st.text_area(
                "Question Text",
                value=question.get(
                    "question",
                    ""
                ),
                height=85,
                key=f"question_{index}"
            )

            # Question Type selector & Required toggle
            type_col1, type_col2 = st.columns(
                [3, 1]
            )

            with type_col1:
                question_types = [
                    "Short Answer",
                    "Paragraph",
                    "Multiple Choice"
                ]

                current_type = question.get(
                    "question_type",
                    "Multiple Choice"
                )

                if current_type not in question_types:
                    current_type = "Multiple Choice"

                question_type = st.selectbox(
                    "Question Type",
                    question_types,
                    index=question_types.index(
                        current_type
                    ),
                    key=f"type_{index}"
                )

            with type_col2:
                st.write("")
                st.write("")
                required = st.checkbox(
                    "Required field",
                    value=bool(
                        question.get(
                            "required",
                            True
                        )
                    ),
                    key=f"required_{index}"
                )

            # Multiple choice options
            if question_type == "Multiple Choice":

                existing_options = question.get(
                    "options",
                    []
                )

                options_text = st.text_area(
                    "Answer Choices (one per line)",
                    value="\n".join(
                        str(option)
                        for option in existing_options
                    ),
                    height=100,
                    key=f"options_{index}",
                    help="Enter each selectable choice on its own line."
                )

                options = [
                    line.strip()
                    for line in options_text.splitlines()
                    if line.strip()
                ]

            else:

                options = []

            # Metadata in a clean 2-column layout (responsive on mobile & desktop)
            meta_col1, meta_col2 = st.columns(2)

            with meta_col1:
                section = st.text_input(
                    "Section Name",
                    value=question.get(
                        "section",
                        ""
                    ),
                    key=f"section_{index}"
                )

                session_no = st.text_input(
                    "Session No.",
                    value=question.get(
                        "session_no",
                        ""
                    ),
                    key=f"session_no_{index}"
                )

            with meta_col2:
                session = st.text_input(
                    "Session Title",
                    value=question.get(
                        "session",
                        ""
                    ),
                    key=f"session_{index}"
                )

                category = st.text_input(
                    "Category",
                    value=question.get(
                        "category",
                        ""
                    ),
                    key=f"category_{index}"
                )

            # Update question state
            questions[index] = {
                "section": section,
                "session_no": session_no,
                "session": session,
                "category": category,
                "question": question_text,
                "question_type": question_type,
                "options": options,
                "required": required
            }

    # --------------------------------------------------------
    # Delete
    # --------------------------------------------------------

    if delete_index is not None:

        st.session_state.questions.pop(
            delete_index
        )

        st.rerun()

    # ========================================================
    # ADD QUESTION / SAVE CHANGES (Action Bar)
    # ========================================================

    st.divider()

    action_col1, action_col2 = st.columns(2)

    with action_col1:

        if st.button(
            "➕ Add Custom Question",
            use_container_width=True
        ):

            st.session_state.questions.append(
                {
                    "section": "Additional",
                    "session_no": "",
                    "session": "",
                    "category": "Additional Question",
                    "question": "",
                    "question_type": "Multiple Choice",
                    "options": RATING_OPTIONS.copy(),
                    "required": True
                }
            )

            st.rerun()

    with action_col2:

        if st.button(
            "💾 Save All Changes",
            type="primary",
            use_container_width=True
        ):

            st.success(
                "✅ All survey changes saved successfully."
            )

    # ========================================================
    # STEP 4 — EXCEL EXPORT
    # ========================================================

    st.divider()

    st.markdown("## 4. Download Survey")
    st.caption("Export your structured feedback survey to an Excel workbook formatted for offline distribution or team review.")

    with st.container(border=True):

        d_col1, d_col2 = st.columns([3, 2])

        with d_col1:
            st.markdown("### 📊 Excel (.xlsx) Export")
            st.write(
                "Contains all session questions, options, participant fields, "
                "and formatted column headers ready for spreadsheets."
            )

        with d_col2:
            excel_data = create_excel(
                st.session_state.questions
            )

            st.download_button(
                "⬇️ Download Excel Survey",
                data=excel_data,
                file_name="Feedback_Form.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True
            )

    # ========================================================
    # STEP 5 — GOOGLE FORM
    # ========================================================

    st.divider()

    st.markdown("## 5. Export to Google Forms")
    st.caption("Publish your approved survey straight into your Google account as an active, editable Google Form.")

    with st.container(border=True):

        try:

            oauth2 = get_google_oauth_component()

            # ----------------------------------------------------
            # Not authenticated yet
            # ----------------------------------------------------

            if st.session_state.google_form_token is None:

                st.markdown("### 🔗 Google Account Connection")
                st.write(
                    "Connect your Google account to authorize automatic Google Form generation."
                )

                auth_secrets = st.secrets.get("auth", {})
                redirect_uri = auth_secrets.get("redirect_uri") or (
                    "http://localhost:8501/"
                    "component/"
                    "streamlit_oauth.authorize_button"
                )

                result = oauth2.authorize_button(
                    name="🔗 Connect Google Account",
                    redirect_uri=redirect_uri,
                    scope=(
                        "openid email profile "
                        f"{FORMS_SCOPE}"
                    ),
                    extras_params={
                        "prompt": "consent",
                        "access_type": "offline"
                    },
                    pkce="S256",
                    key="google_forms_oauth",
                    use_container_width=True
                )

                with st.expander("ℹ️ Connection help & Redirect URI details", expanded=False):
                    st.markdown(
                        f"""
                        **Active Redirect URI sent to Google:**  
                        `{redirect_uri}`

                        > If you see **Google Error 400: redirect_uri_mismatch**, copy the exact URL above and add it to your [Google Cloud Console](https://console.cloud.google.com/apis/credentials) under **OAuth 2.0 Client IDs > Authorized redirect URIs**.
                        """
                    )

                if result and "token" in result:

                    st.session_state.google_form_token = (
                        result["token"]
                    )

                    st.rerun()

            # ----------------------------------------------------
            # Authenticated
            # ----------------------------------------------------

            else:

                auth_col1, auth_col2 = st.columns(
                    [3, 1]
                )

                with auth_col1:
                    st.success(
                        "✅ Google account connected successfully."
                    )

                with auth_col2:
                    if st.button(
                        "Disconnect",
                        use_container_width=True
                    ):

                        st.session_state.google_form_token = None

                        st.rerun()

                st.write("")

                # ------------------------------------------------
                # Create form button
                # ------------------------------------------------

                if st.button(
                    "🚀 Create Google Form Now",
                    type="primary",
                    use_container_width=True
                ):

                    try:

                        with st.spinner(
                            "Publishing questions to Google Forms API..."
                        ):

                            token = (
                                st.session_state
                                .google_form_token
                            )

                            access_token = token[
                                "access_token"
                            ]

                            form_id = create_google_form(
                                "Workshop Feedback Form",
                                st.session_state.questions,
                                access_token
                            )

                        st.success(
                            "🎉 Google Form created successfully!"
                        )

                        st.markdown(
                            f"""
                            <div style="background: var(--asg-pill-curr-bg); border: 1px solid var(--asg-pill-curr-border); border-radius: 12px; padding: 18px; margin-top: 10px;">
                                <h3 style="margin-top: 0; color: var(--asg-text-primary);">✅ Your Google Form is Live!</h3>
                                <p style="color: var(--asg-text-secondary); margin-bottom: 14px;">The form has been published to your Google Drive.</p>
                                <a href="https://docs.google.com/forms/d/{form_id}/edit" target="_blank" style="display: inline-block; background: #7C3AED; color: #FFFFFF; font-weight: 600; padding: 10px 20px; border-radius: 10px; text-decoration: none; margin-bottom: 12px;">
                                    ↗️ Open in Google Forms
                                </a>
                                <div style="font-size: 0.85rem; color: var(--asg-text-muted); margin-top: 8px;">
                                    Form ID: <code>{form_id}</code>
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )

                    except Exception as error:

                        st.error(
                            "❌ Failed to create Google Form."
                        )

                        st.exception(error)

        except KeyError:

            st.error(
                "Google OAuth credentials are missing "
                "from the [auth] section of secrets.toml."
            )

        except Exception as error:

            st.error(
                "Google Form authorization could not be loaded."
            )

            st.exception(error)