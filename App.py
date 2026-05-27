from flask import Flask, request, jsonify, send_file, Response
from flask_cors import CORS
from openai import OpenAI
from dotenv import load_dotenv
import pandas as pd
import io
import json
import time
import threading
import uuid
import os
import subprocess
import tempfile

load_dotenv()

ENV_API_KEY = os.getenv("OPENAI_API_KEY", "")

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# =========================
# CONFIG
# =========================

MODEL = "gpt-4.1"

SYSTEM_PROMPT = """
You are an expert SEO content writer.

Rules:
- Write professional SEO optimized content
- Do NOT mention company name or website name
- Content should be generic and usable for backlinks
- Use headings and subheadings
- Do NOT use em dashes
- Use primary keyword naturally in title
- Use primary and secondary keywords naturally in first 4 paragraphs
- Bold all target keywords
- Content should be approximately 1000 words
- Use proper formatting
- Human sounding writing
- Avoid keyword stuffing
- Make content informative and readable
"""

# In-memory job store
jobs = {}

# =========================
# SERVE FRONTEND
# =========================

@app.route('/')
def index():
    return send_file('index.html')

# =========================
# HELPERS
# =========================

def generate_article(client, primary_keyword, secondary_keyword):
    user_prompt = f"""
Write a complete SEO optimized article.

Primary Keyword:
{primary_keyword}

Secondary Keyword:
{secondary_keyword}

Instructions:
Please provide 1000 words contents for each page.
Each content should be generic on the targeted keywords without using any company or website name.
Each article needs to be SEO optimized and suitable for backlink purposes.
Each article must have an SEO optimized title having the primary keyword in it.
Use the provided target keywords naturally within the first to fourth paragraphs and make them bold.
Do not use em dashes.
Use headings and sub headings and make the content relevant and professional.
Use prepositions in keywords so that they make sense.
"""
    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.7,
    )
    return response.choices[0].message.content


def run_generation_job(job_id, api_key, keywords_data):
    jobs[job_id]["status"] = "running"
    client = OpenAI(api_key=api_key)
    results = []

    total = len(keywords_data)
    for i, row in enumerate(keywords_data):
        if jobs[job_id].get("cancelled"):
            jobs[job_id]["status"] = "cancelled"
            return

        website = row.get("website", "")
        primary = row.get("primary_keyword", "")
        secondary = row.get("secondary_keyword", "")

        jobs[job_id]["progress"] = {
            "current": i + 1,
            "total": total,
            "current_keyword": primary
        }

        try:
            article = generate_article(client, primary, secondary)
            results.append({
                "website": website,
                "primary_keyword": primary,
                "secondary_keyword": secondary,
                "generated_article": article,
                "status": "success"
            })
        except Exception as e:
            results.append({
                "website": website,
                "primary_keyword": primary,
                "secondary_keyword": secondary,
                "generated_article": f"ERROR: {str(e)}",
                "status": "error"
            })

        time.sleep(1)

    jobs[job_id]["results"] = results
    jobs[job_id]["status"] = "completed"
    jobs[job_id]["progress"]["current"] = total


# =========================
# ROUTES
# =========================

@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "env_key_loaded": bool(ENV_API_KEY)})


@app.route("/api/generate", methods=["POST"])
def start_generation():
    data = request.get_json()
    # Frontend se aaye key ko use karo, nahi toh .env wali
    api_key = data.get("api_key", "").strip() or ENV_API_KEY
    keywords = data.get("keywords", [])

    if not api_key:
        return jsonify({"error": "API key missing — .env mein daalo ya frontend mein enter karo"}), 400
    if not keywords:
        return jsonify({"error": "keywords list is empty"}), 400

    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "status": "queued",
        "progress": {"current": 0, "total": len(keywords), "current_keyword": ""},
        "results": [],
        "cancelled": False
    }

    thread = threading.Thread(
        target=run_generation_job,
        args=(job_id, api_key, keywords),
        daemon=True
    )
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/api/job/<job_id>", methods=["GET"])
def get_job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "job_id": job_id,
        "status": job["status"],
        "progress": job["progress"],
        "result_count": len(job["results"])
    })


@app.route("/api/job/<job_id>/cancel", methods=["POST"])
def cancel_job(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    job["cancelled"] = True
    return jsonify({"message": "Cancellation requested"})


@app.route("/api/job/<job_id>/results", methods=["GET"])
def get_job_results(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "job_id": job_id,
        "status": job["status"],
        "results": job["results"]
    })


@app.route("/api/job/<job_id>/download/excel", methods=["GET"])
def download_excel(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    if not job["results"]:
        return jsonify({"error": "No results yet"}), 400

    df = pd.DataFrame(job["results"])
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"seo_articles_{job_id[:8]}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/api/job/<job_id>/download/word", methods=["GET"])
def download_word(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    if not job["results"]:
        return jsonify({"error": "No results yet"}), 400

    results = job["results"]

    # Windows compatible temp path
    tmp_dir = tempfile.gettempdir()
    output_docx = os.path.join(tmp_dir, f"seo_{job_id[:8]}.docx")

    js_lines = []
    js_lines.append("""
const { Document, Packer, Paragraph, TextRun, HeadingLevel, PageBreak } = require('docx');
const fs = require('fs');
const sections_children = [];
""")

    for i, r in enumerate(results):
        article_text = r.get("generated_article", "").replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
        primary = r.get("primary_keyword", "").replace("'", "\\'")
        secondary = r.get("secondary_keyword", "").replace("'", "\\'")
        website = r.get("website", "").replace("'", "\\'")

        if i > 0:
            js_lines.append("sections_children.push(new Paragraph({ children: [new PageBreak()] }));")

        js_lines.append(f"""
sections_children.push(new Paragraph({{
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({{ text: 'Article {i+1}: {primary}', bold: true }})]
}}));
sections_children.push(new Paragraph({{ children: [new TextRun({{ text: 'Website: ', bold: true }}), new TextRun('{website}')] }}));
sections_children.push(new Paragraph({{ children: [new TextRun({{ text: 'Primary Keyword: ', bold: true }}), new TextRun('{primary}')] }}));
sections_children.push(new Paragraph({{ children: [new TextRun({{ text: 'Secondary Keyword: ', bold: true }}), new TextRun('{secondary}')] }}));
sections_children.push(new Paragraph({{ children: [] }}));
""")

        lines = article_text.split("\\n")
        js_lines.append("const article_lines_{i} = {lines};".format(i=i, lines=json.dumps(lines)))
        js_lines.append(f"""
for (const line of article_lines_{i}) {{
  const trimmed = line.trim();
  if (trimmed.startsWith('## ')) {{
    sections_children.push(new Paragraph({{ heading: HeadingLevel.HEADING_2, children: [new TextRun(trimmed.replace(/^## /, ''))] }}));
  }} else if (trimmed.startsWith('# ')) {{
    sections_children.push(new Paragraph({{ heading: HeadingLevel.HEADING_1, children: [new TextRun(trimmed.replace(/^# /, ''))] }}));
  }} else if (trimmed === '') {{
    sections_children.push(new Paragraph({{ children: [] }}));
  }} else {{
    sections_children.push(new Paragraph({{ children: [new TextRun(trimmed.replace(/\\*\\*(.*?)\\*\\*/g, '$1'))] }}));
  }}
}}
""")

    # Windows path mein backslash escape karna zaroori hai
    escaped_output = output_docx.replace("\\", "\\\\")
    js_lines.append(f"""
const doc = new Document({{
  styles: {{
    default: {{ document: {{ run: {{ font: "Calibri", size: 24 }} }} }},
    paragraphStyles: [
      {{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: {{ size: 36, bold: true, font: "Calibri", color: "1F4E79" }},
        paragraph: {{ spacing: {{ before: 360, after: 120 }}, outlineLevel: 0 }} }},
      {{ id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: {{ size: 28, bold: true, font: "Calibri", color: "2E74B5" }},
        paragraph: {{ spacing: {{ before: 240, after: 80 }}, outlineLevel: 1 }} }}
    ]
  }},
  sections: [{{
    properties: {{ page: {{ size: {{ width: 12240, height: 15840 }}, margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }} }} }},
    children: sections_children
  }}]
}});

Packer.toBuffer(doc).then(buffer => {{
  fs.writeFileSync('{escaped_output}', buffer);
  console.log('DONE');
}});
""")

    js_script = "\n".join(js_lines)

    with tempfile.NamedTemporaryFile(mode='w', suffix='.js', delete=False, dir=tmp_dir) as f:
        f.write(js_script)
        script_path = f.name

    try:
        result = subprocess.run(
            ["node", script_path],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode != 0:
            return jsonify({"error": "DOCX generation failed", "detail": result.stderr}), 500

        return send_file(
            output_docx,
            as_attachment=True,
            download_name=f"seo_articles_{job_id[:8]}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    finally:
        try:
            os.unlink(script_path)
        except:
            pass


@app.route("/api/upload-excel", methods=["POST"])
def upload_excel():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    try:
        df = pd.read_excel(file)
        df.columns = [c.strip() for c in df.columns]

        keywords = []
        for _, row in df.iterrows():
            keywords.append({
                "website": str(row.get("Website", row.get("website", ""))).strip(),
                "primary_keyword": str(row.get("Primary Keywords", row.get("primary_keyword", ""))).strip(),
                "secondary_keyword": str(row.get("Secondary Keywords", row.get("secondary_keyword", ""))).strip(),
            })
        return jsonify({"keywords": keywords, "count": len(keywords)})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


if __name__ == "__main__":
    app.run(debug=True, port=5000)