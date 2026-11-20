import os, re, json, uuid, base64
from datetime import datetime
from flask import Flask, request, jsonify, render_template, abort
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
import requests
from pypdf import PdfReader
from docx import Document
from pptx import Presentation

BASE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE, ".env"))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
DATA_DIR = os.path.join(BASE, "data")
os.makedirs(DATA_DIR, exist_ok=True)

GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = os.getenv("GROQ_MODEL", "openai/gpt-oss-120b")
ALLOWED = {".pdf", ".docx", ".pptx", ".txt", ".md"}

def call_groq(prompt):
    key = os.getenv("GROQ_API_KEY")
    if not key:
        raise RuntimeError("Missing GROQ_API_KEY. Add it to your .env file.")
    response = requests.post(
        GROQ_URL,
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
        json={
            "model": GROQ_MODEL,
            "messages": [
                {"role": "system", "content": (
                    "You are an expert interview coach. Create practical, accurate interview "
                    "preparation material grounded in the supplied source. Do not invent facts "
                    "about the candidate. Clearly label any general interview advice."
                )},
                {"role": "user", "content": prompt}
            ],
            "temperature": 0.35,
            "max_tokens": 5000
        },
        timeout=120
    )
    if not response.ok:
        try:
            detail = response.json().get("error", {}).get("message", response.text)
        except Exception:
            detail = response.text
        raise RuntimeError(f"Groq API error ({response.status_code}): {str(detail)[:500]}")
    return response.json()["choices"][0]["message"]["content"]

def extract_text(file_storage):
    filename = secure_filename(file_storage.filename or "upload")
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED:
        raise ValueError(f"Unsupported file type: {ext or 'unknown'}. Use PDF, DOCX, PPTX, TXT or MD.")
    if ext == ".pdf":
        reader = PdfReader(file_storage)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    if ext == ".docx":
        doc = Document(file_storage)
        return "\n".join(p.text for p in doc.paragraphs)
    if ext == ".pptx":
        prs = Presentation(file_storage)
        chunks = []
        for i, slide in enumerate(prs.slides, 1):
            chunks.append(f"Slide {i}")
            for shape in slide.shapes:
                if getattr(shape, "has_text_frame", False):
                    chunks.append(shape.text_frame.text)
                if getattr(shape, "has_table", False) and shape.has_table:
                    for row in shape.table.rows:
                        chunks.append(" | ".join(cell.text for cell in row.cells))
        return "\n".join(chunks)
    return file_storage.read().decode("utf-8", errors="ignore")

def save_pack(title, content):
    pack_id = uuid.uuid4().hex[:10]
    record = {
        "id": pack_id, "title": title, "content": content,
        "created": datetime.now().strftime("%d %b %Y, %I:%M %p")
    }
    with open(os.path.join(DATA_DIR, pack_id + ".json"), "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False)
    return record

def load_pack(pack_id):
    if not re.fullmatch(r"[a-f0-9]{10}", pack_id or ""):
        return None
    path = os.path.join(DATA_DIR, pack_id + ".json")
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    try:
        uploaded = request.files.getlist("files")
        pasted = request.form.get("text", "").strip()
        if not uploaded and not pasted:
            return jsonify(error="Upload a file or paste your study material first."), 400

        parts, names = [], []
        for f in uploaded:
            if not f.filename:
                continue
            names.append(secure_filename(f.filename))
            extracted = extract_text(f)
            if extracted.strip():
                parts.append(f"\n--- SOURCE: {f.filename} ---\n{extracted}")
        if pasted:
            parts.insert(0, "\n--- PASTED SOURCE ---\n" + pasted)

        source = "\n".join(parts).strip()
        if len(source) < 80:
            return jsonify(error="I couldn't extract enough text. If your PDF is scanned, try an OCR/text-based PDF or paste the text."), 400

        # Keep requests manageable for the free tier and tell the model if input was shortened.
        truncated = len(source) > 24000
        source = source[:24000]
        role = request.form.get("role", "").strip() or "General fresher-level role"
        level = request.form.get("level", "Fresher")
        focus = request.form.get("focus", "").strip() or "Balanced technical and HR preparation"

        prompt = f"""
Create a practical interview preparation pack using the source material below.

Candidate target role: {role}
Experience level: {level}
Preparation focus: {focus}

Use this exact Markdown structure:
# Interview Preparation Pack
## 1. Source Material Overview
- Summarise the topics and skills actually present in the source.
## 2. Important Concepts to Revise
- Explain key concepts simply, with examples only when supported or clearly marked as illustrative.
## 3. Technical Interview Questions
- Create 10 likely questions from the source. Give a concise model answer for each.
## 4. Scenario-Based Questions
- Create 5 practical/application questions with suggested answer points.
## 5. HR and Behavioural Questions
- Create 5 common questions and answer frameworks. Do not invent the candidate's personal experiences; use placeholders where personal details are needed.
## 6. Coding / Practical Questions
- Include up to 5 only if relevant to the source and target role. Otherwise say not applicable.
## 7. Quick Revision Checklist
- A concise checklist of what to revise before the interview.
## 8. Mock Interview
- Give 5 questions one at a time is not possible here, so list the questions and what a strong answer should cover.

Rules:
- Prioritise the uploaded material. Do not claim that a topic appears in the source if it does not.
- If the source is a resume, tailor questions to its projects, skills and claims. Flag claims the candidate should be ready to explain.
- If it is textbook/class notes, make the questions suitable for interviews for the target role.
- Use beginner-friendly, direct language. Avoid generic filler.
- Keep answers concise but useful.
- Mark any extra general interview advice as "General tip".
- If the source does not contain enough information for a section, say so rather than fabricating details.

SOURCE MATERIAL:
{source}
"""
        content = call_groq(prompt)
        title = (", ".join(names) if names else "Pasted material")[:90]
        record = save_pack(title, content)
        record["truncated"] = truncated
        return jsonify(id=record["id"], truncated=truncated)
    except ValueError as e:
        return jsonify(error=str(e)), 400
    except Exception as e:
        return jsonify(error=str(e)), 500

@app.route("/pack/<pack_id>")
def pack_page(pack_id):
    record = load_pack(pack_id)
    if not record:
        abort(404)
    return render_template("pack.html", pack=record)

@app.route("/history")
def history():
    items = []
    for filename in os.listdir(DATA_DIR):
        if filename.endswith(".json"):
            try:
                with open(os.path.join(DATA_DIR, filename), encoding="utf-8") as f:
                    item = json.load(f)
                items.append(item)
            except Exception:
                pass
    items.sort(key=lambda x: x.get("created", ""), reverse=True)
    return render_template("history.html", items=items)

@app.route("/delete/<pack_id>", methods=["POST"])
def delete(pack_id):
    record = load_pack(pack_id)
    if record:
        os.remove(os.path.join(DATA_DIR, pack_id + ".json"))
    return jsonify(ok=True)

@app.errorhandler(404)
def not_found(_):
    return render_template("index.html"), 404

if __name__ == "__main__":
    app.run(debug=True)
