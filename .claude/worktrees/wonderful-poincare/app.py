"""
HPM Systems Proposal Generator
Flask backend — serves UI and API on localhost:5000
Run with:  python app.py
"""

import json
import os
import subprocess
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from document_generator import generate_proposal

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
SETTINGS_FILE = BASE_DIR / "hpm_proposal" / "settings.json"
TEMPLATES_DIR = BASE_DIR / "hpm_proposal" / "templates"
DRAFTS_DIR = BASE_DIR / "hpm_proposal" / "drafts"
OUTPUT_DIR = BASE_DIR / "hpm_proposal" / "output"


# ---------------------------------------------------------------------------
# Settings helpers
# ---------------------------------------------------------------------------

def _load_settings():
    if SETTINGS_FILE.exists():
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def _save_settings(data):
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------------
# Page routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/settings-page")
def settings_page():
    return render_template("settings.html")


# ---------------------------------------------------------------------------
# Settings API
# ---------------------------------------------------------------------------

@app.route("/api/settings", methods=["GET"])
def get_settings():
    return jsonify(_load_settings())


@app.route("/api/settings", methods=["POST"])
def post_settings():
    data = request.json or {}
    # Extract and apply template overrides to their respective JSON files
    overrides = data.pop("_template_overrides", {})
    for template_id, overrides_data in overrides.items():
        _apply_template_overrides(template_id, overrides_data)
    _save_settings(data)
    return jsonify({"ok": True})


def _apply_template_overrides(template_id, overrides):
    """Write template default overrides back to the template JSON file."""
    registry_file = TEMPLATES_DIR / "templates.json"
    if not registry_file.exists():
        return
    with open(registry_file, encoding="utf-8") as f:
        registry = json.load(f)
    for filename in registry:
        tpl_file = TEMPLATES_DIR / filename
        if tpl_file.exists():
            with open(tpl_file, encoding="utf-8") as f:
                tpl = json.load(f)
            if tpl.get("template_id") == template_id:
                if "section_headings" in overrides and any(overrides["section_headings"]):
                    tpl["section_headings"] = overrides["section_headings"]
                if overrides.get("default_intro"):
                    tpl["default_intro"] = overrides["default_intro"]
                if overrides.get("default_closing"):
                    tpl["default_closing"] = overrides["default_closing"]
                if overrides.get("default_assumptions") is not None:
                    tpl["default_assumptions"] = overrides["default_assumptions"]
                with open(tpl_file, "w", encoding="utf-8") as f:
                    json.dump(tpl, f, indent=2, ensure_ascii=False)
                break


# ---------------------------------------------------------------------------
# Signature image upload
# ---------------------------------------------------------------------------

@app.route("/hpm_proposal/<path:filename>")
def serve_hpm_file(filename):
    """Serve files from hpm_proposal/ (e.g. signature images for preview)."""
    file_path = BASE_DIR / "hpm_proposal" / filename
    if file_path.exists() and file_path.suffix.lower() in (".png", ".jpg", ".jpeg"):
        return send_file(str(file_path))
    return "", 404


@app.route("/api/upload-signature", methods=["POST"])
def upload_signature():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "No file selected"}), 400
    filename = secure_filename(f.filename)
    dest = BASE_DIR / "hpm_proposal" / filename
    f.save(str(dest))
    rel_path = f"hpm_proposal/{filename}"
    return jsonify({"ok": True, "path": rel_path})


# ---------------------------------------------------------------------------
# Templates API
# ---------------------------------------------------------------------------

@app.route("/api/templates")
def list_templates():
    registry_file = TEMPLATES_DIR / "templates.json"
    if not registry_file.exists():
        return jsonify([])
    with open(registry_file, encoding="utf-8") as f:
        registry = json.load(f)
    templates = []
    for filename in registry:
        tpl_file = TEMPLATES_DIR / filename
        if tpl_file.exists():
            with open(tpl_file, encoding="utf-8") as f:
                tpl = json.load(f)
            templates.append({"id": tpl["template_id"], "name": tpl["template_name"]})
    return jsonify(templates)


@app.route("/api/template/<template_id>")
def get_template(template_id):
    registry_file = TEMPLATES_DIR / "templates.json"
    if not registry_file.exists():
        return jsonify({"error": "No templates found"}), 404
    with open(registry_file, encoding="utf-8") as f:
        registry = json.load(f)
    for filename in registry:
        tpl_file = TEMPLATES_DIR / filename
        if tpl_file.exists():
            with open(tpl_file, encoding="utf-8") as f:
                tpl = json.load(f)
            if tpl["template_id"] == template_id:
                # Substitute {user_email} placeholder in closing paragraph
                settings = _load_settings()
                user_email = settings.get("user", {}).get("email", "")
                if "default_closing" in tpl:
                    tpl["default_closing"] = tpl["default_closing"].replace(
                        "{user_email}", user_email
                    )
                return jsonify(tpl)
    return jsonify({"error": "Template not found"}), 404


# ---------------------------------------------------------------------------
# Document generation API
# ---------------------------------------------------------------------------

@app.route("/api/generate", methods=["POST"])
def generate():
    data = request.json
    settings = _load_settings()
    try:
        docx_bytes = generate_proposal(data, settings)
        proposal_num = (data.get("proposal_number") or "proposal").replace("/", "-")
        filename = f"{proposal_num}.docx"
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / filename
        with open(output_path, "wb") as f:
            f.write(docx_bytes)
        return send_file(
            output_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        app.logger.exception("generate failed")
        return jsonify({"error": str(e)}), 500


@app.route("/api/export-pdf", methods=["POST"])
def export_pdf():
    data = request.json
    settings = _load_settings()
    try:
        docx_bytes = generate_proposal(data, settings)
        proposal_num = (data.get("proposal_number") or "proposal").replace("/", "-")
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        docx_path = OUTPUT_DIR / f"{proposal_num}.docx"
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        # Locate LibreOffice
        lo_path = settings.get("output", {}).get("libreoffice_path", "").strip()
        if not lo_path:
            candidates = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                "/usr/bin/libreoffice",
                "/usr/bin/soffice",
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            ]
            for c in candidates:
                if os.path.exists(c):
                    lo_path = c
                    break

        if not lo_path:
            return jsonify(
                {
                    "error": (
                        "LibreOffice not found. Please download the .docx and "
                        "export to PDF from Word manually."
                    )
                }
            ), 404

        result = subprocess.run(
            [lo_path, "--headless", "--convert-to", "pdf", "--outdir", str(OUTPUT_DIR), str(docx_path)],
            capture_output=True,
            timeout=60,
        )
        pdf_path = OUTPUT_DIR / f"{proposal_num}.pdf"
        if not pdf_path.exists():
            stderr = result.stderr.decode(errors="replace")
            return jsonify(
                {
                    "error": (
                        f"PDF conversion failed. Please download the .docx and export "
                        f"to PDF from Word manually. Detail: {stderr[:200]}"
                    )
                }
            ), 500

        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=f"{proposal_num}.pdf",
            mimetype="application/pdf",
        )
    except subprocess.TimeoutExpired:
        return jsonify({"error": "PDF conversion timed out."}), 500
    except Exception as e:
        app.logger.exception("export-pdf failed")
        return jsonify({"error": str(e)}), 500


# ---------------------------------------------------------------------------
# Draft API
# ---------------------------------------------------------------------------

@app.route("/api/save-draft", methods=["POST"])
def save_draft():
    payload = request.json
    name = (payload.get("draft_name") or "").strip()
    if not name:
        name = f"draft_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    # Store draft_name inside payload for reference
    payload["draft_name"] = name
    DRAFTS_DIR.mkdir(parents=True, exist_ok=True)
    draft_file = DRAFTS_DIR / f"{name}.json"
    with open(draft_file, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)
    return jsonify({"ok": True, "name": name})


@app.route("/api/list-drafts")
def list_drafts():
    DRAFTS_DIR.mkdir(parents=True, exist_ok=True)
    drafts = []
    for f in sorted(DRAFTS_DIR.glob("*.json"), key=lambda x: x.stat().st_mtime, reverse=True):
        drafts.append(
            {
                "name": f.stem,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
            }
        )
    return jsonify(drafts)


@app.route("/api/load-draft/<name>")
def load_draft(name):
    draft_file = DRAFTS_DIR / f"{name}.json"
    if not draft_file.exists():
        return jsonify({"error": "Draft not found"}), 404
    with open(draft_file, encoding="utf-8") as f:
        return jsonify(json.load(f))


@app.route("/api/delete-draft/<name>", methods=["DELETE"])
def delete_draft(name):
    draft_file = DRAFTS_DIR / f"{name}.json"
    if draft_file.exists():
        draft_file.unlink()
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------

def _open_browser():
    import time
    time.sleep(1.5)
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    threading.Thread(target=_open_browser, daemon=True).start()
    print("\n  HPM Proposal Generator running at http://localhost:5000")
    print("  Press Ctrl+C to stop.\n")
    app.run(debug=False, host="localhost", port=5000)
