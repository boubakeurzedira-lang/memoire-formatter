"""API HTTP appelée par Make.

POST /spec-text   (multipart: file=cahier .pdf/.docx)        -> {"text": "..."}
POST /format      (multipart: file=memoire.docx, rules=JSON) -> .docx formaté (+ en-tête X-Format-Report)
GET  /health
Authentification : en-tête  Authorization: Bearer <API_KEY>   (variable d'environnement API_KEY)
"""
import io
import json
import os
import subprocess
import tempfile

from docx import Document
from flask import Flask, Response, jsonify, request

from formatter import deep_merge, DEFAULT_RULES, format_docx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024
API_KEY = os.environ.get("API_KEY", "")


@app.before_request
def check_auth():
    if request.path == "/health":
        return None
    if not API_KEY:
        return jsonify(error="API_KEY non configurée sur le serveur"), 500
    if request.headers.get("Authorization", "") != "Bearer " + API_KEY:
        return jsonify(error="non autorisé"), 401
    return None


@app.get("/health")
def health():
    return jsonify(status="ok")


def _docx_text(data):
    doc = Document(io.BytesIO(data))
    parts = [p.text for p in doc.paragraphs if p.text.strip()]
    for t in doc.tables:
        for row in t.rows:
            parts.append(" | ".join(c.text.strip() for c in row.cells))
    return "\n".join(parts)


def _pdf_text(data):
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "in.pdf")
        with open(path, "wb") as f:
            f.write(data)
        out = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True, timeout=60)
    return out.stdout.decode("utf-8", "ignore")


@app.post("/spec-text")
def spec_text():
    f = request.files.get("file")
    if not f:
        return jsonify(error="champ 'file' manquant"), 400
    data = f.read()
    name = (f.filename or "").lower()
    try:
        if name.endswith(".docx") or data[:2] == b"PK":
            text = _docx_text(data)
        elif name.endswith(".pdf") or data[:4] == b"%PDF":
            text = _pdf_text(data)
        else:
            return jsonify(error="format non pris en charge (docx ou pdf)"), 415
    except Exception as e:  # fichier corrompu
        return jsonify(error="lecture impossible: %s" % e), 422
    return jsonify(text=text[:60000], chars=len(text), truncated=len(text) > 60000)


@app.get("/default-rules")
def default_rules():
    return jsonify(DEFAULT_RULES)


@app.post("/format")
def format_endpoint():
    f = request.files.get("file")
    if not f:
        return jsonify(error="champ 'file' manquant"), 400
    raw_rules = request.form.get("rules", "") or "{}"
    try:
        rules = json.loads(raw_rules)
        if not isinstance(rules, dict):
            raise ValueError("rules doit être un objet JSON")
    except ValueError as e:
        return jsonify(error="rules invalide: %s" % e), 400
    try:
        out, report = format_docx(f.read(), rules, with_pages=request.form.get("pages", "1") != "0")
    except Exception as e:
        return jsonify(error="échec de la mise en forme: %s" % e), 422
    summary = {k: report[k] for k in ("cover_detected", "has_chapters", "uncertain_headings",
                                       "toc_pages_estimated", "counts", "warnings")}
    resp = Response(out, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    resp.headers["Content-Disposition"] = 'attachment; filename="memoire_formate.docx"'
    resp.headers["X-Format-Report"] = json.dumps(summary, ensure_ascii=True)
    resp.headers["X-Outline-Count"] = str(len(report["outline"]))
    return resp


@app.post("/outline")
def outline():
    """Même traitement, mais renvoie le plan détecté (JSON) pour relecture."""
    f = request.files.get("file")
    if not f:
        return jsonify(error="champ 'file' manquant"), 400
    rules = json.loads(request.form.get("rules", "") or "{}")
    _, report = format_docx(f.read(), rules, with_pages=False)
    return jsonify(report)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
