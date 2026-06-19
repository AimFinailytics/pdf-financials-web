from __future__ import annotations

import json
import os
import sys
import time
import zipfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from financials_converter import cleanup_dir, convert_pdfs, create_processing_dir


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_ROOT = BASE_DIR / "uploads"
OUTPUT_ROOT = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {".pdf"}

# How long a generated Excel output is kept before being swept off disk.
# Long enough for a user to download right after converting; short enough that
# the ephemeral disk doesn't fill up over time.
OUTPUT_TTL_SECONDS = int(os.environ.get("OUTPUT_TTL_SECONDS", str(6 * 60 * 60)))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 250 * 1024 * 1024


def purge_old_outputs() -> None:
    """Delete output job folders older than OUTPUT_TTL_SECONDS so generated
    Excel files don't accumulate forever on the (small, ephemeral) disk."""
    if not OUTPUT_ROOT.exists():
        return
    cutoff = time.time() - OUTPUT_TTL_SECONDS
    for child in OUTPUT_ROOT.iterdir():
        try:
            if child.is_dir() and child.stat().st_mtime < cutoff:
                cleanup_dir(child)
        except OSError:
            pass


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/health")
def health():
    # Render (and any uptime monitor) pings this to confirm the app is live.
    return jsonify({"status": "ok"}), 200


@app.post("/api/convert")
def convert():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "Upload at least one PDF."}), 400

    purge_old_outputs()

    job_id = os.urandom(8).hex()
    upload_dir = create_processing_dir(UPLOAD_ROOT)
    output_dir = OUTPUT_ROOT / job_id
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_paths: list[Path] = []
    try:
        for file in files:
            original_name = file.filename or "uploaded.pdf"
            suffix = Path(original_name).suffix.lower()
            if suffix not in ALLOWED_EXTENSIONS:
                continue
            safe_name = secure_filename(original_name) or f"upload_{len(pdf_paths) + 1}.pdf"
            dest = upload_dir / safe_name
            file.save(dest)
            pdf_paths.append(dest)

        if not pdf_paths:
            cleanup_dir(upload_dir)
            cleanup_dir(output_dir)
            return jsonify({"error": "No valid PDF files were uploaded."}), 400

        result = convert_pdfs(pdf_paths, output_dir)

        downloadable = []
        for path in result.output_paths:
            downloadable.append(
                {
                    "name": path.name,
                    "url": f"/download/{job_id}/{path.name}",
                    "size": path.stat().st_size,
                }
            )

        if len(downloadable) > 1:
            zip_path = output_dir / "all_outputs.zip"
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
                for path in result.output_paths:
                    archive.write(path, arcname=path.name)
            downloadable.insert(
                0,
                {
                    "name": zip_path.name,
                    "url": f"/download/{job_id}/{zip_path.name}",
                    "size": zip_path.stat().st_size,
                },
            )

        manifest = {
            "job_id": job_id,
            "files": downloadable,
            "summaries": result.summaries,
            "skipped": result.skipped,
        }
        (output_dir / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        return jsonify(manifest)
    except Exception as exc:
        cleanup_dir(output_dir)
        return jsonify({"error": f"Conversion failed: {exc}"}), 500
    finally:
        cleanup_dir(upload_dir)


@app.get("/download/<job_id>/<filename>")
def download(job_id: str, filename: str):
    safe_job = secure_filename(job_id)
    safe_name = secure_filename(filename)
    path = OUTPUT_ROOT / safe_job / safe_name
    if not path.exists() or not path.is_file():
        return jsonify({"error": "File not found."}), 404
    return send_file(path, as_attachment=True, download_name=path.name)


if __name__ == "__main__":
    # Local/dev entrypoint. In production the app is served by gunicorn
    # (see Procfile), which imports `app` directly and never runs this block.
    port = int(os.environ.get("PORT", "5050"))
    debug = os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    host = os.environ.get("HOST", "0.0.0.0")
    app.run(host=host, port=port, debug=debug)
