# 📋 AI Survey Generator

An AI-powered Streamlit application that turns a workshop/training agenda into a complete, editable feedback survey. The application uses Gemini to understand the agenda, generates session-specific feedback questions, and lets the user export the approved survey to Excel or create a Google Form directly in the Google account that authorizes the app.

## 🚀 Live Application

**Streamlit:** https://ai-survey-generator-73cdh7m55q2my7l5k5kksj.streamlit.app/

**GitHub:** https://github.com/mohit-dhoundiyal/AI-Survey-Generator

## ✨ Features

* Upload a workshop agenda in **DOCX** format.
* Optionally upload an existing feedback form in **XLSX** format as a style/reference file.
* Use **Gemini AI** to identify actual agenda sessions and generate **one specific feedback question per session**.
* Preserve the session number, title, and order.
* Automatically add three required participant fields:

  * Full Name
  * Designation
  * Organisation/Department
* Automatically add one overall workshop rating question.
* Automatically add one optional suggestions question.
* Use a five-point session rating scale:

  * Excellent
  * Very Good
  * Good
  * Fair
  * Poor
* Review and edit generated questions before export.
* Delete questions or add custom questions.
* Download a structured Excel file with a `Feedback Form` worksheet.
* Download a reusable Google Apps Script as a fallback workflow.
* Connect a user's Google account through OAuth 2.0 and create the Google Form directly in the authorizing user's account.

## 🧠 How It Works

```text
Workshop Agenda (DOCX)
          │
          ▼
     Agenda Parsing
          │
          ▼
       Gemini AI
          │
          ▼
 Session-specific Survey
          │
          ▼
     Review / Edit
       ┌──┴───────────────┐
       │                  │
       ▼                  ▼
   Excel Export      Google OAuth
                          │
                          ▼
                   Google Forms API
                          │
                          ▼
                  User's Google Form
```

## 📝 Survey Structure

The generated survey follows a consistent structure:

1. **Participant Information**

   * Full Name — Short Answer — Required
   * Designation — Short Answer — Required
   * Organisation/Department — Short Answer — Required
2. **Session Feedback**

   * Exactly one multiple-choice feedback question for each actual agenda session.
   * Options: Excellent, Very Good, Good, Fair, Poor.
   * Required.
3. **Overall Workshop**

   * How would you rate the overall workshop?
   * Multiple Choice.
   * Required.
4. **Suggestions**

   * Please provide your suggestions for improving future workshops.
   * Paragraph.
   * Optional.

## 📁 Project Structure

```text
AI-Survey-Generator/
├── app.py
├── requirements.txt
├── .gitignore
└── README.md
```

Local-only files that should **not** be committed:

```text
.env
.streamlit/secrets.toml
```

## 🛠️ Technologies Used

* **Python**
* **Streamlit** — web application UI
* **Google Gemini API** — agenda understanding and survey generation
* **python-docx** — DOCX agenda extraction
* **OpenPyXL** — Excel import/export
* **streamlit-oauth** — Google OAuth authorization flow
* **Google Forms API** — direct Google Form creation
* **Google API Python Client / google-auth** — Google API access

## ⚙️ Local Setup

### 1\. Clone the repository

```bash
git clone https://github.com/mohit-dhoundiyal/AI-Survey-Generator.git
cd AI-Survey-Generator
```

### 2\. Create and activate a virtual environment (recommended)

**Windows PowerShell:**

```powershell
python -m venv venv
.\\venv\\Scripts\\Activate.ps1
```

### 3\. Install dependencies

```bash
pip install -r requirements.txt
```

### 4\. Configure Gemini

Create a local `.env` file:

```env
GEMINI\_API\_KEY=your\_gemini\_api\_key
```

### 5\. Configure Google OAuth

Create:

```text
.streamlit/secrets.toml
```

Add the required OAuth configuration using the Google OAuth client credentials for the project. Keep all secrets out of GitHub.

For local OAuth testing, the callback used by the app is:

```text
http://localhost:8501/component/streamlit\_oauth.authorize\_button
```

This callback must also be registered as an authorized redirect URI for the Google OAuth client.

### 6\. Run the application

```bash
streamlit run app.py
```

Then open:

```text
http://localhost:8501
```

## 🔐 Google OAuth and Google Forms

The application uses OAuth so the person using the application can authorize access to Google Forms. The Google Form is created using the permissions granted by the account that completes the OAuth flow.

The Forms API permission used by the application is:

```text
https://www.googleapis.com/auth/forms.body
```

During development, a Google OAuth application in **Testing** mode may require users to be added as test users in Google Cloud. For public use, the OAuth consent configuration and any Google verification requirements applicable to the requested scopes must be completed.

## ☁️ Streamlit Community Cloud Deployment

1. Push the project to GitHub.
2. Open Streamlit Community Cloud.
3. Create or update the app from the GitHub repository.
4. Select `app.py` as the main file.
5. Add the required secrets in the Streamlit app's **Secrets** settings.
6. Deploy the app.

For the deployed OAuth flow, register the deployed callback URL with the Google OAuth client:

```text
https://YOUR-STREAMLIT-HOST/component/streamlit\_oauth.authorize\_button
```

Replace `YOUR-STREAMLIT-HOST` with the actual deployed Streamlit hostname.

## 📤 Excel Export

The generated Excel workbook contains a single worksheet named:

```text
Feedback Form
```

Columns:

```text
No.
Section
Session No.
Session
Category
Question
Question Type
Options
Required
```

## 🧩 Google Apps Script Fallback

The application also provides a downloadable Google Apps Script file. This can be used as a fallback when direct Google OAuth/Form creation is not available.

Typical workflow:

```text
Generate survey
      ↓
Download Excel
      ↓
Import Excel into Google Sheets
      ↓
Open Apps Script
      ↓
Paste the generated script
      ↓
Authorize with the user's Google account
      ↓
Create Google Form
```

The script creates the Form under the Google account that authorizes the script and can link responses to the same spreadsheet.

## 🔒 Security Notes

* Never commit `.env` or `.streamlit/secrets.toml` to GitHub.
* Never expose Gemini API keys, OAuth client secrets, cookie secrets, or access tokens in the source code.
* Use Streamlit Cloud Secrets for production credentials.
* Do not print or log OAuth access tokens.

## 🎯 Intended Use

The project is designed for workshops, training programs, conferences, capacity-building programs, and other learning events where feedback forms need to be generated quickly from an agenda.



