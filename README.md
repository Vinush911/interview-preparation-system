# PrepMate AI — Interview Preparation System

A simple Flask website that accepts a resume, textbook/class notes, PDF, DOCX, PPTX, TXT or Markdown text and uses the Groq API to create a role-focused interview preparation pack.

## Features
- Upload one or more source files or paste text.
- Choose target role, experience level and preparation focus.
- Generate source-grounded concepts, technical Q&A, scenario questions, HR questions, practical questions (when relevant), and a revision checklist.
- Save generated packs locally, reopen/delete them, copy content, or print/save as PDF.

## Setup
1. Create a free Groq API key: https://console.groq.com/keys
2. In this folder, copy `.env.example` to `.env`.
3. Open `.env` and replace the placeholder with your key:
   `GROQ_API_KEY=gsk_your_actual_key`
4. (Optional) Change `GROQ_MODEL` to a model currently available in your Groq account.
5. Install dependencies:
   ```bash
   python -m venv venv
   source venv/bin/activate        # macOS/Linux
   # venv\Scripts\activate       # Windows
   pip install -r requirements.txt
   ```
6. Start the app:
   ```bash
   python app.py
   ```
7. Open http://127.0.0.1:5000

## Notes
- Maximum upload size is 20 MB per request.
- Text extraction is supported for PDF, DOCX, PPTX, TXT and MD.
- Scanned/image-only PDFs may not contain extractable text; use an OCR/text PDF or paste the text.
- Generated packs are stored as JSON files in `data/` on your computer.
- Do not upload confidential resumes or personal data to an API unless you are comfortable sending that content to Groq.
