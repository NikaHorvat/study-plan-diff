#!/usr/bin/env python3
"""Web interface backend for study-plan comparison."""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from compare_study_plans import run

BASE_DIR = Path(__file__).resolve().parent
ALLOWED_EXTENSIONS = {".xlsx"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80 MB


@app.errorhandler(413)
def payload_too_large(_exc):
    return jsonify({"error": "Uploaded files are too large. The combined upload must stay below 80 MB."}), 413


@app.errorhandler(500)
def internal_server_error(_exc):
    return jsonify({"error": "The server hit an internal error while processing the comparison."}), 500


def is_allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    return send_file(BASE_DIR / "website.html")


@app.post("/compare")
def compare_endpoint():
    old_file = request.files.get("old")
    new_file = request.files.get("new")

    if not old_file or not new_file:
        return jsonify({"error": "Both old and new .xlsx files are required."}), 400

    if not old_file.filename or not new_file.filename:
        return jsonify({"error": "Invalid uploaded file name."}), 400

    if not is_allowed(old_file.filename) or not is_allowed(new_file.filename):
        return jsonify({"error": "Only .xlsx files are accepted."}), 400

    with tempfile.TemporaryDirectory(prefix="study_plan_diff_") as tmp_dir:
        tmp_path = Path(tmp_dir)

        old_name = secure_filename(old_file.filename)
        new_name = secure_filename(new_file.filename)
        if not old_name:
            old_name = "old.xlsx"
        if not new_name:
            new_name = "new.xlsx"

        old_path = tmp_path / old_name
        new_path = tmp_path / new_name

        old_file.save(old_path)
        new_file.save(new_path)

        output_name = f"{Path(old_name).stem}_with_diffs.xlsx"
        output_path = tmp_path / output_name

        try:
            run(str(old_path), str(new_path), str(output_path))
        except Exception as exc:  # pragma: no cover
            app.logger.exception("Comparison failed")
            return jsonify({"error": f"Comparison failed: {exc}"}), 500

        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5050"))
    app.run(host="0.0.0.0", port=port, debug=True)

