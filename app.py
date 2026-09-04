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
# LIGHT STYLING (visual polish only — no logic here)
# ============================================================

st.markdown(
    """
    <style>

        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');

        html, body, [class*="css"] {
            font-family: 'Poppins', sans-serif;
        }

        .stApp {
            background: linear-gradient(180deg, #F5F3FF 0%, #FDF4FF 45%, #FFF7ED 100%);
        }

        .block-container {
            padding-top: 2rem;
            padding-bottom: 3rem;
            max-width: 1100px;
        }

        /* ---- Hero title banner ---- */
        .hero-banner {
            background: linear-gradient(120deg, #7C3AED 0%, #DB2777 55%, #F97316 100%);
            padding: 28px 32px;
            border-radius: 18px;
            margin-bottom: 1.6rem;
            box-shadow: 0 10px 30px rgba(124, 58, 237, 0.25);
        }
        .hero-banner h1 {
            color: white !important;
            margin: 0 0 6px 0;
            font-weight: 700;
        }
        .hero-banner p {
            color: rgba(255,255,255,0.92);
            margin: 0;
            font-size: 1.02rem;
        }

        /* ---- Section headers get a colored accent bar ---- */
        h3 {
            border-left: 6px solid #A855F7;
            padding-left: 12px;
            border-radius: 3px;
        }

        /* ---- Cards / containers / expanders ---- */
        div[data-testid="stExpander"] {
            border: 1px solid #E9D5FF;
            border-radius: 14px;
            background: linear-gradient(135deg, #FFFFFF 0%, #FAF5FF 100%);
            box-shadow: 0 2px 10px rgba(168, 85, 247, 0.08);
            margin-bottom: 10px;
        }
        div[data-testid="stExpander"] details summary p {
            font-size: 1.02rem;
            font-weight: 600;
            color: #6D28D9;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 14px !important;
        }

        /* ---- Give bordered containers (upload cards) real breathing room ---- */
        div[data-testid="stVerticalBlock"] div[data-testid="stVerticalBlockBorderWrapper"] > div {
            padding: 22px 24px;
        }

        /* ---- Wider gutter between side-by-side columns ---- */
        div[data-testid="stHorizontalBlock"] {
            gap: 2.2rem;
        }

        /* ---- Buttons: colorful gradient ---- */
        .stButton>button {
            border-radius: 10px;
            border: none;
            background: linear-gradient(90deg, #7C3AED, #DB2777);
            color: white;
            font-weight: 600;
            transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .stButton>button:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 16px rgba(219, 39, 119, 0.35);
            color: white;
        }
        .stButton>button p {
            color: white !important;
        }

        .stDownloadButton>button {
            border-radius: 10px;
            border: none;
            background: linear-gradient(90deg, #F59E0B, #F97316);
            color: white;
            font-weight: 600;
        }
        .stDownloadButton>button:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 16px rgba(249, 115, 22, 0.35);
            color: white;
        }
        .stDownloadButton>button p {
            color: white !important;
        }

        /* ---- Sidebar ---- */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #4C1D95 0%, #831843 100%);
        }
        section[data-testid="stSidebar"] * {
            color: #F3E8FF !important;
        }
        section[data-testid="stSidebar"] hr {
            border-color: rgba(255,255,255,0.25);
        }

        /* ---- Progress step pills ---- */
        .step-pill {
            display: inline-block;
            padding: 3px 12px;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 700;
            margin-bottom: 8px;
        }
        .step-done { background-color: #34D399; color: #052E16; }
        .step-current { background-color: #FBBF24; color: #451A03; }
        .step-todo { background-color: rgba(255,255,255,0.15); color: #F3E8FF; }

        /* ---- Metric ---- */
        div[data-testid="stMetric"] {
            background: rgba(255,255,255,0.12);
            border-radius: 12px;
            padding: 10px;
        }

        /* ---- Tags for question type ---- */
        .qtype-tag {
            display: inline-block;
            padding: 2px 10px;
            border-radius: 999px;
            font-size: 0.72rem;
            font-weight: 700;
            margin-left: 8px;
        }
        .qtype-mc { background-color: #DBEAFE; color: #1E40AF; }
        .qtype-short { background-color: #FCE7F3; color: #9D174D; }
        .qtype-para { background-color: #FEF3C7; color: #92400E; }

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

    revoke_token_url = (
        "https://oauth2.googleapis.com/revoke"
    )

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

    st.markdown("**Progress**")

    step1_done = st.session_state.generated
    step2_done = bool(st.session_state.questions) and step1_done
    step3_done = st.session_state.google_form_token is not None

    def pill(label, done, current=False):
        css_class = "step-done" if done else ("step-current" if current else "step-todo")
        icon = "✅" if done else ("➡️" if current else "◻️")
        st.markdown(
            f'<span class="step-pill {css_class}">{icon} {label}</span>',
            unsafe_allow_html=True
        )

    pill("1. Upload agenda", step1_done, current=not step1_done)
    pill("2. Generate survey", step1_done, current=step1_done and not step2_done)
    pill("3. Review & edit", step2_done)
    pill("4. Export (Excel / Forms)", step2_done)

    st.divider()

    if st.session_state.questions:
        st.metric("Questions in survey", len(st.session_state.questions))

    st.divider()
    st.caption(
        "Tip: uploading an existing feedback form (XLSX) in step 2 "
        "helps Gemini match your usual tone and structure."
    )


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="hero-banner">
        <h1>📋 AI Survey Generator</h1>
        <p>Upload a workshop agenda and generate a complete feedback survey using Gemini AI ✨</p>
    </div>
    """,
    unsafe_allow_html=True
)

st.write("")


# ============================================================
# STEP 1 & 2 — INPUTS (grouped side-by-side for a tighter flow)
# ============================================================

input_col1, input_col2 = st.columns(2, gap="large")

with input_col1:

    with st.container(border=True):

        st.subheader("1. Upload Workshop Agenda")

        agenda_file = st.file_uploader(
            "Upload agenda DOCX",
            type=["docx"],
            key="agenda_file",
            help="Required. The agenda drives the session-by-session questions."
        )

        if agenda_file is not None:
            st.caption(f"📄 {agenda_file.name}")

with input_col2:

    with st.container(border=True):

        st.subheader("2. Optional Example Feedback Form")

        example_file = st.file_uploader(
            "Upload an existing feedback XLSX for style/reference (optional)",
            type=["xlsx"],
            key="example_file",
            help="Optional. Used only as a style reference, not copied verbatim."
        )

        if example_file is not None:
            st.caption(f"📄 {example_file.name}")


# ============================================================
# GENERATE SURVEY
# ============================================================

st.write("")

generate_clicked = st.button(
    "🔍 Analyze & Generate Survey",
    type="primary",
    use_container_width=True
)

if generate_clicked:

    if agenda_file is None:

        st.error(
            "Please upload the workshop agenda DOCX."
        )

    else:

        try:

            with st.status(
                "Generating your survey...",
                expanded=True
            ) as status:

                st.write("📖 Reading agenda...")

                agenda_text = extract_docx_text(
                    agenda_file
                )

                example_text = ""

                if example_file is not None:

                    st.write("📖 Reading example feedback form...")

                    example_text = extract_xlsx_text(
                        example_file
                    )

                st.write("🤖 Asking Gemini to draft the survey...")

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
                    label="Survey generated successfully",
                    state="complete",
                    expanded=False
                )

            st.success(
                f"Survey generated successfully with "
                f"{len(final_questions)} questions."
            )

        except Exception as error:

            st.error(
                "Failed to generate survey."
            )

            st.exception(error)


# ============================================================
# STEP 3 — REVIEW
# ============================================================

if st.session_state.generated:

    st.divider()

    st.subheader(
        "3. Review & Edit Survey"
    )

    st.caption(
        "Expand a question to edit it. Participant-information "
        "questions are added automatically and can still be edited here."
    )

    questions = st.session_state.questions

    delete_index = None

    # --------------------------------------------------------
    # Existing questions (now inside expanders for a cleaner list)
    # --------------------------------------------------------

    for index, question in enumerate(
        questions
    ):

        preview = question.get("question", "").strip() or "(empty question)"
        type_label = question.get("question_type", "")
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
        expander_title = f"{emoji} Q{index + 1} · {preview}"

        with st.expander(expander_title, expanded=False):

            header_col1, header_col2 = st.columns(
                [8, 2]
            )

            with header_col1:
                tag_class = tag_class_map.get(type_label, "qtype-mc")
                st.markdown(
                    f'Type: <span class="qtype-tag {tag_class}">{type_label}</span>',
                    unsafe_allow_html=True
                )

            with header_col2:

                if st.button(
                    "🗑️ Delete",
                    key=f"delete_{index}",
                    use_container_width=True
                ):

                    delete_index = index

            # ------------------------------------------------
            # Metadata
            # ------------------------------------------------

            left_col, right_col = st.columns(
                [2, 5]
            )

            with left_col:

                section = st.text_input(
                    "Section",
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

                session = st.text_input(
                    "Session",
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

            with right_col:

                question_text = st.text_area(
                    "Question",
                    value=question.get(
                        "question",
                        ""
                    ),
                    height=100,
                    key=f"question_{index}"
                )

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

            required = st.checkbox(
                "Required",
                value=bool(
                    question.get(
                        "required",
                        True
                    )
                ),
                key=f"required_{index}"
            )

            # ------------------------------------------------
            # Multiple choice options
            # ------------------------------------------------

            if question_type == "Multiple Choice":

                existing_options = question.get(
                    "options",
                    []
                )

                options_text = st.text_area(
                    "Options (one per line)",
                    value="\n".join(
                        str(option)
                        for option in existing_options
                    ),
                    height=110,
                    key=f"options_{index}"
                )

                options = [
                    line.strip()
                    for line in options_text.splitlines()
                    if line.strip()
                ]

            else:

                options = []

            # ------------------------------------------------
            # Update session state
            # ------------------------------------------------

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
    # ADD QUESTION / SAVE CHANGES (grouped as an action bar)
    # ========================================================

    st.divider()

    action_col1, action_col2 = st.columns(2)

    with action_col1:

        if st.button(
            "➕ Add Question",
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
            "💾 Save Changes",
            use_container_width=True
        ):

            st.success(
                "Survey changes saved."
            )

    # ========================================================
    # STEP 4 — EXCEL
    # ========================================================

    st.divider()

    st.subheader(
        "4. Download Survey"
    )

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

    st.subheader(
        "5. Create Google Form Automatically"
    )

    st.write(
        "Connect your Google account and create the "
        "approved survey directly in your own account."
    )

    try:

        oauth2 = get_google_oauth_component()

        # ----------------------------------------------------
        # Not authenticated yet
        # ----------------------------------------------------

        if st.session_state.google_form_token is None:

            result = oauth2.authorize_button(
                name="🔗 Connect Google Account",
                redirect_uri=(
                    "http://localhost:8501/"
                    "component/"
                    "streamlit_oauth.authorize_button"
                ),
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

            if result and "token" in result:

                st.session_state.google_form_token = (
                    result["token"]
                )

                st.rerun()

        # ----------------------------------------------------
        # Authenticated
        # ----------------------------------------------------

        else:

            st.success(
                "✅ Google account connected."
            )

            col1, col2 = st.columns(
                [3, 1]
            )

            with col2:

                if st.button(
                    "Disconnect"
                ):

                    st.session_state.google_form_token = None

                    st.rerun()

            # ------------------------------------------------
            # Create form
            # ------------------------------------------------

            if st.button(
                "🚀 Create Google Form",
                type="primary",
                use_container_width=True
            ):

                try:

                    with st.spinner(
                        "Creating Google Form..."
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
                        ### ✅ Your Google Form is ready

                        [Open Google Form](https://docs.google.com/forms/d/{form_id}/edit)
                        """
                    )

                    st.write(
                        "Form ID:"
                    )

                    st.code(
                        form_id,
                        language="text"
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