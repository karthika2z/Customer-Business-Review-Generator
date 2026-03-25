"""
app.py
Flask web UI for the CBR generator.

Upload flow:
  - User drag-and-drops individual CSV files (no ZIP needed)
  - Server processes entirely in a temporary directory (no persistent storage)
  - Generated PPTX is streamed back as a file download
  - Temp directory is deleted immediately after streaming

Template:
  - If env var TEMPLATE_BUCKET is set → downloaded from GCS at startup and cached in memory
  - Otherwise → loaded from the local path (template/CBR Template.pptx)

Designed for Google Cloud Run (stateless, ephemeral).
"""

import io
import logging
import os
import tempfile
import traceback
from datetime import date
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify

from data_loader import CustomerData, extract_excel_to_dir
from generate_cbr import build_presentation

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
CSV_EXT    = ".csv"

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ------------------------------------------------------------------
# Template loading (once at startup)
# ------------------------------------------------------------------
TEMPLATE_BUCKET = os.environ.get("TEMPLATE_BUCKET", "")
TEMPLATE_BLOB   = os.environ.get("TEMPLATE_BLOB",   "CBR Template.pptx")
LOCAL_TEMPLATE  = Path(__file__).parent / "template" / "CBR Template.pptx"

_template_bytes: bytes | None = None


def _load_template() -> bytes:
    """Return the template PPTX as bytes (GCS or local file)."""
    global _template_bytes
    if _template_bytes is not None:
        return _template_bytes

    if TEMPLATE_BUCKET:
        log.info("Downloading template from GCS: gs://%s/%s", TEMPLATE_BUCKET, TEMPLATE_BLOB)
        from google.cloud import storage  # imported lazily so local runs don't need the SDK
        client = storage.Client()
        blob   = client.bucket(TEMPLATE_BUCKET).blob(TEMPLATE_BLOB)
        _template_bytes = blob.download_as_bytes()
        log.info("Template loaded from GCS (%d bytes)", len(_template_bytes))
    elif LOCAL_TEMPLATE.exists():
        log.info("Loading template from local path: %s", LOCAL_TEMPLATE)
        _template_bytes = LOCAL_TEMPLATE.read_bytes()
    else:
        raise FileNotFoundError(
            f"No template found. Set TEMPLATE_BUCKET env var or place the file at {LOCAL_TEMPLATE}"
        )

    return _template_bytes


# ------------------------------------------------------------------
# Flask app
# ------------------------------------------------------------------
MAX_UPLOAD_MB = 100

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    uploaded = request.files.getlist("csvfiles")
    if not uploaded or all(f.filename == "" for f in uploaded):
        return jsonify({"error": "No files received."}), 400

    # Classify uploads: one Excel workbook OR one-or-more CSV files
    excel_files = [f for f in uploaded if Path(f.filename).suffix.lower() in EXCEL_EXTS]
    csv_files   = [f for f in uploaded if Path(f.filename).suffix.lower() == CSV_EXT]
    other_files = [f for f in uploaded
                   if Path(f.filename).suffix.lower() not in EXCEL_EXTS | {CSV_EXT}]

    if other_files:
        names = ", ".join(f.filename for f in other_files)
        return jsonify({"error": f"Only .csv or .xlsx files are accepted. Got: {names}"}), 400
    if excel_files and csv_files:
        return jsonify({"error": "Please upload either one Excel workbook or CSV files, not both."}), 400
    if len(excel_files) > 1:
        return jsonify({"error": "Please upload a single Excel workbook."}), 400

    # Process entirely in a temp directory — deleted when the block exits
    with tempfile.TemporaryDirectory(prefix="cbr_") as tmpdir:
        tmp = Path(tmpdir)

        if excel_files:
            # ── Extract sheets from Excel workbook into CSVs ─────────
            xlsx_file = excel_files[0]
            xlsx_path = tmp / Path(xlsx_file.filename).name
            xlsx_path.write_bytes(xlsx_file.read())
            try:
                extract_excel_to_dir(xlsx_path, tmp)
            except Exception as exc:
                log.exception("Excel extraction failed")
                return jsonify({"error": f"Failed to read Excel workbook: {exc}"}), 500
            xlsx_path.unlink()   # remove the .xlsx so CustomerData doesn't try to read it
            log.info("Extracted Excel workbook to CSVs in %s", tmp)
        else:
            # ── Save individual CSV files ────────────────────────────
            for f in csv_files:
                dest = tmp / Path(f.filename).name
                dest.write_bytes(f.read())
            log.info("Saved %d CSV files to %s", len(csv_files), tmp)

        # ── Load customer data ───────────────────────────────────────
        try:
            data = CustomerData(tmp)
        except Exception as exc:
            log.exception("CustomerData load failed")
            return jsonify({"error": f"Failed to read CSV data: {exc}"}), 500

        # ── Get template ─────────────────────────────────────────────
        try:
            template = _load_template()
        except Exception as exc:
            log.exception("Template load failed")
            return jsonify({"error": f"Could not load template: {exc}"}), 500

        # ── Generate PPTX ────────────────────────────────────────────
        try:
            out_path = tmp / "output.pptx"
            build_presentation(
                data=data,
                template_path=template,   # bytes → Presentation(BytesIO(...))
                output_path=str(out_path),
            )
            pptx_bytes = out_path.read_bytes()
        except Exception as exc:
            log.exception("Presentation generation failed")
            return jsonify({
                "error": f"Generation failed: {exc}",
                "detail": traceback.format_exc(),
            }), 500

    # ── Temp dir deleted — stream from memory ────────────────────────
    today     = date.today().strftime("%Y-%m-%d")
    safe_name = data.customer_name.replace(" ", "_").replace("/", "-")
    filename  = f"CBR - {safe_name} - {today}.pptx"
    mime      = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    log.info("Streaming %s (%d bytes)", filename, len(pptx_bytes))
    return send_file(
        io.BytesIO(pptx_bytes),
        mimetype=mime,
        as_attachment=True,
        download_name=filename,
    )


# ------------------------------------------------------------------
# Entry point (Cloud Run sets PORT env var to 8080)
# ------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
