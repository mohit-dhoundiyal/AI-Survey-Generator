import os
import json
import re
from io import BytesIO

import streamlit as st
import openpyxl

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from dotenv import load_dotenv
from google import genai

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# =========================================================
# CONFIGURATION
# =========================================================

load_dotenv()

GEMINI_MODEL = "gemini-3.6-flash"


# =========================================================
# GEMINI SETUP
# =========================================================

api_key = os.getenv("GEMINI_API_KEY")

if not api_key:
    st.error(
        "GEMINI_API_KEY not found. "
        "Please check your .env file."
    )
    st.stop()

client = genai.Client(
    api_key=api_key
)


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="AI Survey Generator",
    page_icon="📋",
    layout="wide"
)


# =========================================================
# SESSION STATE
# =========================================================

if "questions" not in st.session_state:
    st.session_state.questions = []

if "files_analyzed" not in st.session_state:
    st.session_state.files_analyzed = False

if "agenda_text" not in st.session_state:
    st.session_state.agenda_text = ""

if "example_info" not in st.session_state:
    st.session_state.example_info = {
        "columns": [],
        "question_headers": [],
        "sample_values": {}
    }


# =========================================================
# DOCX READER
# =========================================================

def iter_docx_blocks(parent):
    """
    Read paragraphs and tables in the order
    they appear in the DOCX.
    """

    body = parent.element.body

    for child in body.iterchildren():

        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)

        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def extract_docx_text(uploaded_file):

    document = Document(uploaded_file)

    lines = []

    for block in iter_docx_blocks(document):

        # -------------------------------------------------
        # Paragraph
        # -------------------------------------------------

        if isinstance(block, Paragraph):

            text = block.text.strip()

            if text:
                lines.append(text)

        # -------------------------------------------------
        # Table
        # -------------------------------------------------

        elif isinstance(block, Table):

            for row in block.rows:

                cells = []

                for cell in row.cells:

                    cell_parts = []

                    for paragraph in cell.paragraphs:

                        text = paragraph.text.strip()

                        if text:
                            cell_parts.append(text)

                    cell_text = " ".join(cell_parts)

                    cell_text = re.sub(
                        r"\s+",
                        " ",
                        cell_text
                    )

                    cells.append(cell_text)

                if any(cells):

                    lines.append(
                        " | ".join(cells)
                    )

    return "\n".join(lines)


# =========================================================
# EXAMPLE XLSX READER
# =========================================================

def extract_example_form(uploaded_file):

    uploaded_file.seek(0)

    workbook = openpyxl.load_workbook(
        uploaded_file,
        data_only=True
    )

    result = {
        "columns": [],
        "question_headers": [],
        "sample_values": {}
    }

    for sheet_name in workbook.sheetnames:

        sheet = workbook[sheet_name]

        rows = list(
            sheet.iter_rows(
                values_only=True
            )
        )

        if not rows:
            continue

        headers = []

        for value in rows[0]:

            if value is not None:

                text = str(value).strip()

                if text:
                    headers.append(text)

        if not result["columns"]:
            result["columns"] = headers

        # -------------------------------------------------
        # Detect question-related columns
        # -------------------------------------------------

        for header in headers:

            lower_header = header.lower()

            if (
                lower_header.startswith("q")
                or "rate" in lower_header
                or "suggest" in lower_header
                or "feedback" in lower_header
            ):

                if header not in result["question_headers"]:
                    result["question_headers"].append(
                        header
                    )

        # -------------------------------------------------
        # Read sample answers
        # -------------------------------------------------

        for col_index, header in enumerate(
            rows[0]
        ):

            if not header:
                continue

            header = str(header).strip()

            values = []

            for row in rows[1:15]:

                if col_index >= len(row):
                    continue

                value = row[col_index]

                if value is not None:

                    value = str(value).strip()

                    if (
                        value
                        and value not in values
                    ):
                        values.append(value)

            if values:

                result["sample_values"][header] = (
                    values[:10]
                )

    return result


# =========================================================
# BUILD EXAMPLE-FORM INSTRUCTION
# =========================================================

def build_example_instruction(example_info):

    question_headers = example_info.get(
        "question_headers",
        []
    )

    sample_values = example_info.get(
        "sample_values",
        {}
    )

    # No example provided
    if not question_headers:

        return """
No example feedback form was provided.

Use the application's default professional survey style:

- Short Answer for participant information
- Multiple Choice for session evaluation
- Options:
  Excellent
  Very Good
  Good
  Fair
  Poor
- One overall workshop rating
- One optional paragraph suggestions question
"""

    question_style = "\n".join(
        f"- {question}"
        for question in question_headers
    )

    sample_values_json = json.dumps(
        sample_values,
        ensure_ascii=False,
        indent=2
    )

    return f"""
An example feedback form was provided.

Use it ONLY as a reference for:
- wording style
- response style
- professionalism
- survey structure

Example question styles:

{question_style}

Sample response values:

{sample_values_json}

Do NOT copy questions from the example.
The new questions must be based on the uploaded agenda.
"""


# =========================================================
# GEMINI - GENERATE SESSION QUESTIONS
# =========================================================

def generate_feedback_form(
    agenda_text,
    example_info
):

    example_instruction = (
        build_example_instruction(
            example_info
        )
    )

    prompt = f"""
You are an expert professional survey designer.

Your task is to create ONE complete workshop feedback
survey from the workshop agenda below.

========================================================
WORKSHOP AGENDA
========================================================

{agenda_text}

========================================================
EXAMPLE FEEDBACK FORM
========================================================

{example_instruction}

========================================================
FIXED PARTICIPANT INFORMATION
========================================================

The application will automatically add these three
questions BEFORE all other questions:

1. Full Name
2. Designation
3. Organisation/Department

DO NOT generate these questions.

========================================================
SESSION FEEDBACK
========================================================

Analyze the entire agenda carefully.

Identify EVERY actual session in the agenda.

For EVERY actual session:

- Create EXACTLY ONE feedback question.
- Keep the sessions in the same order as the agenda.
- Use the actual session number.
- Use the actual session title.
- Use the actual session details/topics.
- Use the speaker only if clearly provided.
- Do not merge separate sessions.
- Do not skip sessions.
- Do not invent sessions.
- Do not create separate questions for every subtopic.

The question should evaluate the usefulness and/or
quality of that specific session while clearly reflecting
what was actually covered.

BAD QUESTION:

"How satisfied were you with the session?"

GOOD QUESTION:

"How would you rate the usefulness and delivery of the
Data Preparation using Excel session, particularly its
coverage of data structuring, tables, sorting, filtering
and data validation?"

Every session question must be:

Type:
Multiple Choice

Options exactly:

Excellent
Very Good
Good
Fair
Poor

Required:
Yes

========================================================
OVERALL WORKSHOP
========================================================

Create exactly ONE overall question:

"How would you rate the overall workshop?"

Type:
Multiple Choice

Options:

Excellent
Very Good
Good
Fair
Poor

Required:
Yes

========================================================
SUGGESTIONS
========================================================

Create exactly ONE final question:

"Please provide your suggestions for improving future
workshops."

Type:
Paragraph

Required:
No

========================================================
FINAL ORDER
========================================================

The generated AI questions must appear in this order:

Session 1
Session 2
Session 3
...
Last Session
Overall Workshop
Suggestions

The application itself will insert the first three
participant questions before these.

Return ONLY JSON.
"""

    schema = {
        "type": "ARRAY",
        "items": {
            "type": "OBJECT",
            "properties": {

                "section": {
                    "type": "STRING",
                    "enum": [
                        "Session Feedback",
                        "Overall Feedback",
                        "Suggestions"
                    ]
                },

                "session_number": {
                    "type": "STRING"
                },

                "session_name": {
                    "type": "STRING"
                },

                "question": {
                    "type": "STRING"
                },

                "type": {
                    "type": "STRING",
                    "enum": [
                        "Multiple Choice",
                        "Paragraph"
                    ]
                },

                "options": {
                    "type": "ARRAY",
                    "items": {
                        "type": "STRING"
                    }
                },

                "category": {
                    "type": "STRING"
                },

                "required": {
                    "type": "BOOLEAN"
                }
            },

            "required": [
                "section",
                "session_number",
                "session_name",
                "question",
                "type",
                "options",
                "category",
                "required"
            ]
        }
    }

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=prompt,
        config={
            "response_mime_type": "application/json",
            "response_json_schema": schema
        }
    )

    generated = json.loads(
        response.text
    )

    return normalize_questions(
        generated
    )


# =========================================================
# NORMALIZE AI OUTPUT
# =========================================================

def normalize_questions(
    questions
):

    normalized = []

    # =====================================================
    # FIRST 3 QUESTIONS
    # Added ONLY here.
    # =====================================================

    normalized.append({

        "section":
            "Participant Information",

        "session_number":
            "",

        "session_name":
            "",

        "question":
            "Full Name",

        "type":
            "Short Answer",

        "options":
            [],

        "category":
            "Participant Information",

        "required":
            True

    })


    normalized.append({

        "section":
            "Participant Information",

        "session_number":
            "",

        "session_name":
            "",

        "question":
            "Designation",

        "type":
            "Short Answer",

        "options":
            [],

        "category":
            "Participant Information",

        "required":
            True

    })


    normalized.append({

        "section":
            "Participant Information",

        "session_number":
            "",

        "session_name":
            "",

        "question":
            "Organisation/Department",

        "type":
            "Short Answer",

        "options":
            [],

        "category":
            "Participant Information",

        "required":
            True

    })


    # =====================================================
    # ADD AI QUESTIONS
    # =====================================================

    for item in questions:

        if not isinstance(
            item,
            dict
        ):
            continue

        question = str(
            item.get(
                "question",
                ""
            )
        ).strip()

        if not question:
            continue

        # -------------------------------------------------
        # Prevent participant-question duplication
        # -------------------------------------------------

        if item.get("section") == (
            "Participant Information"
        ):
            continue


        # -------------------------------------------------
        # Prevent exact duplicate wording
        # -------------------------------------------------

        question_lower = (
            question.lower().strip()
        )

        existing = [

            q["question"]
            .lower()
            .strip()

            for q in normalized

        ]


        if question_lower in existing:
            continue


        # -------------------------------------------------
        # Question type and options
        # -------------------------------------------------

        question_type = item.get(
            "type",
            "Multiple Choice"
        )


        if question_type == (
            "Multiple Choice"
        ):

            options = [

                "Excellent",
                "Very Good",
                "Good",
                "Fair",
                "Poor"

            ]

        elif question_type == (
            "Paragraph"
        ):

            options = []

        else:

            question_type = (
                "Multiple Choice"
            )

            options = [

                "Excellent",
                "Very Good",
                "Good",
                "Fair",
                "Poor"

            ]


        normalized.append({

            "section":
                item.get(
                    "section",
                    "Session Feedback"
                ),

            "session_number":
                item.get(
                    "session_number",
                    ""
                ),

            "session_name":
                item.get(
                    "session_name",
                    ""
                ),

            "question":
                question,

            "type":
                question_type,

            "options":
                options,

            "category":
                item.get(
                    "category",
                    "General"
                ),

            "required":
                item.get(
                    "required",
                    True
                )

        })


    return normalized


# =========================================================
# CREATE FINAL EXCEL
# =========================================================

def create_excel(
    questions
):

    workbook = Workbook()

    ws = workbook.active

    ws.title = "Feedback Form"


    # =====================================================
    # HEADERS
    # =====================================================

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


    ws.append(
        headers
    )


    # =====================================================
    # HEADER STYLE
    # =====================================================

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )


    header_font = Font(
        color="FFFFFF",
        bold=True
    )


    border = Border(

        left=Side(
            style="thin",
            color="D9E1F2"
        ),

        right=Side(
            style="thin",
            color="D9E1F2"
        ),

        top=Side(
            style="thin",
            color="D9E1F2"
        ),

        bottom=Side(
            style="thin",
            color="D9E1F2"
        )

    )


    for cell in ws[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(

            horizontal="center",

            vertical="center"

        )

        cell.border = border


    # =====================================================
    # DATA
    # =====================================================

    for number, item in enumerate(
        questions,
        start=1
    ):

        options = item.get(
            "options",
            []
        )


        options_text = (

            " | ".join(
                options
            )

            if options

            else ""

        )


        ws.append([

            number,

            item.get(
                "section",
                ""
            ),

            item.get(
                "session_number",
                ""
            ),

            item.get(
                "session_name",
                ""
            ),

            item.get(
                "category",
                ""
            ),

            item.get(
                "question",
                ""
            ),

            item.get(
                "type",
                ""
            ),

            options_text,

            "Yes"

            if item.get(
                "required",
                True
            )

            else "No"

        ])


    # =====================================================
    # DATA STYLE
    # =====================================================

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row
    ):

        for cell in row:

            cell.border = border

            cell.alignment = Alignment(

                vertical="top",

                wrap_text=True

            )


    # =====================================================
    # COLUMN WIDTHS
    # =====================================================

    widths = {

        "A": 8,

        "B": 25,

        "C": 15,

        "D": 50,

        "E": 28,

        "F": 85,

        "G": 22,

        "H": 45,

        "I": 12

    }


    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width


    ws.freeze_panes = "A2"

    ws.auto_filter.ref = (
        ws.dimensions
    )


    # =====================================================
    # RETURN EXCEL BYTES
    # =====================================================

    output = BytesIO()

    workbook.save(
        output
    )

    output.seek(0)

    return output.getvalue()


# =========================================================
# GOOGLE FORMS SCRIPT GENERATOR
# =========================================================

def create_google_forms_script():

    script = r'''
/**
 * AI SURVEY GENERATOR
 *
 * EXPECTED SHEET:
 *     Feedback Form
 *
 * EXPECTED COLUMNS:
 *
 * No.
 * Section
 * Session No.
 * Session
 * Category
 * Question
 * Question Type
 * Options
 * Required
 *
 * HOW TO USE:
 *
 * 1. Download the Excel file from Streamlit.
 * 2. Import the Excel into Google Sheets.
 * 3. Open:
 *
 *      Extensions -> Apps Script
 *
 * 4. Delete the default code.
 * 5. Paste this entire script.
 * 6. Save.
 * 7. Run:
 *
 *      createGoogleFeedbackForm
 *
 * The first run will ask for permission.
 *
 * The script creates ONE Google Form.
 *
 * The Form responses are automatically linked
 * to the same Google Spreadsheet.
 */


/**
 * Create the Google Form.
 */
function createGoogleFeedbackForm() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      "Feedback Form"
    );


  // -------------------------------------------------------
  // Check Sheet
  // -------------------------------------------------------

  if (!sheet) {

    SpreadsheetApp.getUi().alert(
      'The sheet "Feedback Form" was not found.'
    );

    return;
  }


  // -------------------------------------------------------
  // Create Form
  // -------------------------------------------------------

  const formTitle =
    ss.getName() +
    " - Feedback Form";


  const form =
    FormApp.create(
      formTitle
    );


  form.setDescription(
    "Workshop feedback form generated from the approved survey."
  );


  form.setConfirmationMessage(
    "Thank you for submitting your feedback."
  );


  // -------------------------------------------------------
  // Connect responses to spreadsheet
  // -------------------------------------------------------

  form.setDestination(
    FormApp.DestinationType.SPREADSHEET,
    ss.getId()
  );


  // -------------------------------------------------------
  // Read sheet
  // -------------------------------------------------------

  const data =
    sheet.getDataRange().getValues();


  let currentSection = "";


  // -------------------------------------------------------
  // Process rows
  // -------------------------------------------------------

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    const section =
      String(
        row[1] || ""
      ).trim();


    const question =
      String(
        row[5] || ""
      ).trim();


    const questionType =
      String(
        row[6] || ""
      ).trim();


    const optionsText =
      String(
        row[7] || ""
      ).trim();


    const requiredText =
      String(
        row[8] || ""
      ).trim();


    if (!question) {
      continue;
    }


    const required =
      requiredText.toLowerCase() ===
      "yes";


    // -----------------------------------------------------
    // Section heading
    // -----------------------------------------------------

    if (
      section &&
      section !== currentSection
    ) {

      form
        .addSectionHeaderItem()
        .setTitle(section);

      currentSection = section;
    }


    // -----------------------------------------------------
    // Short Answer
    // -----------------------------------------------------

    if (
      questionType ===
      "Short Answer"
    ) {

      form
        .addTextItem()
        .setTitle(question)
        .setRequired(required);

    }


    // -----------------------------------------------------
    // Paragraph
    // -----------------------------------------------------

    else if (
      questionType ===
      "Paragraph"
    ) {

      form
        .addParagraphTextItem()
        .setTitle(question)
        .setRequired(required);

    }


    // -----------------------------------------------------
    // Multiple Choice
    // -----------------------------------------------------

    else if (
      questionType ===
      "Multiple Choice"
    ) {

      let options = [];


      if (optionsText) {

        options =
          optionsText
            .split("|")
            .map(
              function(option) {

                return option.trim();

              }
            )
            .filter(
              function(option) {

                return option.length > 0;

              }
            );
      }


      if (
        options.length > 0
      ) {

        form
          .addMultipleChoiceItem()
          .setTitle(question)
          .setChoiceValues(options)
          .setRequired(required);

      }

      else {

        form
          .addTextItem()
          .setTitle(question)
          .setRequired(required);

      }
    }

  }


  // -------------------------------------------------------
  // Enable Responses
  // -------------------------------------------------------

  form.setPublished(
    true
  );

  form.setAcceptingResponses(
    true
  );


  // -------------------------------------------------------
  // Create result sheet
  // -------------------------------------------------------

  let resultSheet =
    ss.getSheetByName(
      "Created Form"
    );


  if (!resultSheet) {

    resultSheet =
      ss.insertSheet(
        "Created Form"
      );

  }

  else {

    resultSheet.clear();

  }


  resultSheet.appendRow([

    "Form Title",

    "Edit URL",

    "Response URL",

    "Response Spreadsheet",

    "Created On"

  ]);


  resultSheet.appendRow([

    form.getTitle(),

    form.getEditUrl(),

    form.getPublishedUrl(),

    ss.getUrl(),

    new Date()

  ]);


  resultSheet.autoResizeColumns(
    1,
    5
  );


  // -------------------------------------------------------
  // Display response URL
  // -------------------------------------------------------

  SpreadsheetApp.getUi().alert(

    "Google Form created successfully!\n\n" +

    "Response URL:\n" +

    form.getPublishedUrl()

  );


  Logger.log(
    "Edit URL: " +
    form.getEditUrl()
  );


  Logger.log(
    "Response URL: " +
    form.getPublishedUrl()
  );

}


/**
 * Adds a custom menu whenever the sheet opens.
 */
function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu(
      "📋 Survey Form Generator"
    )
    .addItem(
      "Create Google Feedback Form",
      "createGoogleFeedbackForm"
    )
    .addToUi();

}
'''


    return script


# =========================================================
# APPLICATION UI
# =========================================================

st.title(
    "📋 AI Survey Generator"
)

st.write(
    "Upload a workshop agenda and optionally an example "
    "feedback form. The AI creates one survey question "
    "for every session in the agenda."
)

st.divider()


# =========================================================
# STEP 1 — UPLOAD
# =========================================================

st.subheader(
    "1️⃣ Upload Source Files"
)


col1, col2 = st.columns(2)


with col1:

    agenda_file = st.file_uploader(

        "📄 Workshop Agenda (.docx) *",

        type=["docx"],

        help=(
            "Required. Upload the workshop agenda "
            "containing the session details."
        )

    )


with col2:

    example_file = st.file_uploader(

        "📊 Example Feedback Form (.xlsx) — Optional",

        type=["xlsx"],

        help=(
            "Optional. Upload an existing feedback "
            "form to help AI understand your preferred "
            "wording and response style."
        )

    )


# =========================================================
# STEP 2 — ANALYZE
# =========================================================

if st.button(
    "🔍 Analyze Agenda",
    use_container_width=True
):

    if not agenda_file:

        st.warning(
            "Please upload the workshop agenda."
        )

        st.stop()


    with st.spinner(
        "Reading the workshop agenda..."
    ):

        try:

            agenda_file.seek(0)

            agenda_text = extract_docx_text(
                agenda_file
            )


            # -------------------------------------------------
            # Example is optional
            # -------------------------------------------------

            if example_file:

                with st.spinner(
                    "Reading the optional example feedback form..."
                ):

                    example_file.seek(0)

                    example_info = (
                        extract_example_form(
                            example_file
                        )
                    )

                example_message = (
                    "Example feedback form loaded."
                )

            else:

                example_info = {

                    "columns": [],

                    "question_headers": [],

                    "sample_values": {}

                }

                example_message = (
                    "No example form provided. "
                    "The default professional survey style "
                    "will be used."
                )


            st.session_state.agenda_text = (
                agenda_text
            )

            st.session_state.example_info = (
                example_info
            )

            st.session_state.files_analyzed = (
                True
            )

            st.success(
                "Agenda analyzed successfully. "
                + example_message
            )


        except Exception as e:

            st.error(
                "Could not read the uploaded files."
            )

            st.exception(e)


# =========================================================
# STEP 3 — PREVIEW
# =========================================================

if st.session_state.files_analyzed:

    st.divider()

    st.subheader(
        "2️⃣ Source Preview"
    )


    col1, col2 = st.columns(2)


    with col1:

        st.markdown(
            "### 📄 Workshop Agenda"
        )

        st.text_area(

            "Extracted Agenda",

            value=(
                st.session_state
                .agenda_text[:15000]
            ),

            height=400,

            disabled=True

        )


    with col2:

        st.markdown(
            "### 📊 Example Feedback Form"
        )


        info = (
            st.session_state
            .example_info
        )


        if info.get(
            "question_headers"
        ):

            st.success(
                "Example form is being used as a style reference."
            )


            st.write(
                "Detected columns:"
            )


            for column in info.get(
                "columns",
                []
            ):

                st.write(
                    f"• {column}"
                )


            st.write(
                "Detected question styles:"
            )


            for question in info.get(
                "question_headers",
                []
            ):

                st.write(
                    f"• {question[:150]}"
                )


        else:

            st.info(
                "No example form was uploaded. "
                "The app will use its default professional "
                "survey structure."
            )


# =========================================================
# STEP 4 — GENERATE
# =========================================================

if st.session_state.files_analyzed:

    st.divider()

    st.subheader(
        "3️⃣ Generate Feedback Form"
    )


    st.info(
        "The app will automatically add the first three "
        "participant fields and create one AI question "
        "for every session found in the agenda."
    )


    if st.button(

        "🤖 Generate Feedback Form",

        use_container_width=True

    ):

        with st.spinner(

            "Gemini is analyzing every session "
            "and generating the survey..."

        ):

            try:

                questions = (
                    generate_feedback_form(

                        st.session_state
                        .agenda_text,

                        st.session_state
                        .example_info

                    )
                )


                st.session_state.questions = (
                    questions
                )


                st.success(

                    f"Generated {len(questions)} "
                    "survey fields/questions."

                )


            except Exception as e:

                st.error(
                    "Gemini could not generate the survey."
                )

                st.exception(e)


# =========================================================
# STEP 5 — REVIEW / EDIT
# =========================================================

if st.session_state.questions:

    st.divider()

    st.subheader(
        "4️⃣ Review & Edit Your Feedback Form"
    )


    st.info(
        "Edit, save, delete or add questions before "
        "downloading the final files."
    )


    for i, item in enumerate(
        st.session_state.questions
    ):

        st.markdown(
            f"### Question {i + 1}"
        )


        # -------------------------------------------------
        # Question text
        # -------------------------------------------------

        edited_question = st.text_area(

            "Question",

            value=item.get(
                "question",
                ""
            ),

            key=f"question_{i}",

            height=90

        )


        # -------------------------------------------------
        # Section + Session
        # -------------------------------------------------

        col1, col2 = st.columns(2)


        with col1:

            sections = [

                "Participant Information",

                "Session Feedback",

                "Overall Feedback",

                "Suggestions"

            ]


            current_section = item.get(

                "section",

                "Session Feedback"

            )


            if current_section not in sections:

                current_section = (
                    "Session Feedback"
                )


            selected_section = st.selectbox(

                "Section",

                sections,

                index=sections.index(
                    current_section
                ),

                key=f"section_{i}"

            )


        with col2:

            session_name = st.text_input(

                "Session",

                value=item.get(
                    "session_name",
                    ""
                ),

                key=f"session_{i}"

            )


        # -------------------------------------------------
        # Session number + category
        # -------------------------------------------------

        col1, col2 = st.columns(2)


        with col1:

            session_number = st.text_input(

                "Session Number",

                value=item.get(
                    "session_number",
                    ""
                ),

                key=f"session_number_{i}"

            )


        with col2:

            category = st.text_input(

                "Category",

                value=item.get(
                    "category",
                    ""
                ),

                key=f"category_{i}"

            )


        # -------------------------------------------------
        # Question Type
        # -------------------------------------------------

        types = [

            "Short Answer",

            "Multiple Choice",

            "Paragraph"

        ]


        current_type = item.get(

            "type",

            "Multiple Choice"

        )


        if current_type not in types:

            current_type = (
                "Multiple Choice"
            )


        selected_type = st.selectbox(

            "Question Type",

            types,

            index=types.index(
                current_type
            ),

            key=f"type_{i}"

        )


        # -------------------------------------------------
        # Options
        # -------------------------------------------------

        if selected_type == (
            "Multiple Choice"
        ):

            default_options = item.get(
                "options",
                []
            )


            if not default_options:

                default_options = [

                    "Excellent",

                    "Very Good",

                    "Good",

                    "Fair",

                    "Poor"

                ]


            options_text = st.text_area(

                "Answer Options — one per line",

                value="\n".join(
                    default_options
                ),

                key=f"options_{i}",

                height=120

            )


            options = [

                option.strip()

                for option in (
                    options_text.split("\n")
                )

                if option.strip()

            ]

        else:

            options = []


        # -------------------------------------------------
        # Required
        # -------------------------------------------------

        required = st.checkbox(

            "Required",

            value=item.get(
                "required",
                True
            ),

            key=f"required_{i}"

        )


        # -------------------------------------------------
        # Save + Delete
        # -------------------------------------------------

        col1, col2 = st.columns(2)


        with col1:

            if st.button(

                "💾 Save",

                key=f"save_{i}",

                use_container_width=True

            ):

                st.session_state.questions[i] = {

                    "section":
                        selected_section,

                    "session_number":
                        session_number,

                    "session_name":
                        session_name,

                    "question":
                        edited_question,

                    "type":
                        selected_type,

                    "options":
                        options,

                    "category":
                        category,

                    "required":
                        required

                }


                st.success(
                    f"Question {i + 1} saved."
                )


        with col2:

            if st.button(

                "🗑️ Delete",

                key=f"delete_{i}",

                use_container_width=True

            ):

                st.session_state.questions.pop(
                    i
                )

                st.rerun()


        st.divider()


    # =====================================================
    # ADD QUESTION
    # =====================================================

    if st.button(

        "➕ Add Question",

        use_container_width=True

    ):

        st.session_state.questions.append({

            "section":
                "Session Feedback",

            "session_number":
                "",

            "session_name":
                "",

            "question":
                "Enter your question here...",

            "type":
                "Multiple Choice",

            "options": [

                "Excellent",

                "Very Good",

                "Good",

                "Fair",

                "Poor"

            ],

            "category":
                "General",

            "required":
                True

        })


        st.rerun()


# =========================================================
# STEP 6 — FINAL OUTPUTS
# =========================================================

if st.session_state.questions:

    st.divider()

    st.subheader(
        "5️⃣ Final Outputs"
    )


    st.write(
        "Download the approved Excel file and the "
        "Google Forms script."
    )


    # -----------------------------------------------------
    # EXCEL
    # -----------------------------------------------------

    excel_data = create_excel(

        st.session_state.questions

    )


    st.download_button(

        label="📥 Download Final Excel",

        data=excel_data,

        file_name="Feedback_Form.xlsx",

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        use_container_width=True

    )


    # -----------------------------------------------------
    # GOOGLE SCRIPT
    # -----------------------------------------------------

    script_data = (
        create_google_forms_script()
    )


    st.download_button(

        label="📜 Download Google Forms Script",

        data=script_data,

        file_name="Create_Google_Feedback_Form.gs",

        mime="text/plain",

        use_container_width=True

    )


    st.success(
        "Your final feedback package is ready."
    )


# =========================================================
# FOOTER
# =========================================================

st.divider()

st.caption(
    "AI Survey Generator • "
    "Agenda → Session-Based Questions → Excel → Google Form"
)