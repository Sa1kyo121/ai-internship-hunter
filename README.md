# AI Internship Hunter MVP

A simple Streamlit web tool for comparing a resume against an internship job description.

## What It Does

Upload a PDF or Word resume, then paste a job posting URL or job description.
You can also paste resume text manually.
For the cleanest tailored resume export, upload a Word `.docx` file. PDF upload works for text extraction, but exact layout editing is harder.

The tool returns:

- Match score out of 100
- Strong matching skills
- Missing keywords
- Resume improvement suggestions
- Tailored bullet points for the target role
- Tailored resume draft with Word and PDF downloads

Some job boards block automatic page extraction. If a URL cannot be read, paste the job description manually.

## Demo Mode

Demo Mode is turned on by default and does not require an OpenAI API key.
It uses keyword matching to generate a free resume-to-job analysis.

For stronger AI feedback, turn off Demo Mode in the sidebar and add an OpenAI API key.

## Tailored Resume Export

After analyzing a resume and job description, click **Generate Tailored Resume Draft**.
The app creates a resume-style tailored draft and lets you download it as `.docx` or `.pdf`.
Demo Mode uses a compact resume template and keeps changes conservative to avoid adding unnecessary pages.

When the uploaded resume is a Word `.docx`, the Word download preserves the original document formatting and applies conservative targeted bullet rewrites. PDF downloads still use the app's compact resume template.

## Setup

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Optional: add your OpenAI API key:

```bash
copy .env.example .env
```

Then edit `.env` and replace `your_openai_api_key_here` with your real API key.
You can also change `OPENAI_MODEL` if you want to use a different OpenAI model.

3. Run the web tool:

```bash
streamlit run app.py
```

## First Version Scope

This MVP intentionally avoids login, job search automation, dashboards, and databases.
The goal is one clear workflow:

Resume + Job Description -> AI match analysis -> better resume bullets.
