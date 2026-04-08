import os
import glob
import shutil
import sqlite3
import subprocess
import tempfile
from datetime import datetime, timezone

from flask import Flask, render_template, request, send_file, jsonify
from html_to_pptx import html_to_pptx

app = Flask(__name__)

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.db")
ALLOWED_EXT = {".py", ".html", ".htm"}


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute(
        """CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            input_filename TEXT NOT NULL,
            pptx_filename TEXT,
            conversion_type TEXT NOT NULL DEFAULT 'py',
            status TEXT NOT NULL,
            error TEXT,
            created_at TEXT NOT NULL
        )"""
    )
    # Migrate old table if py_filename column exists
    try:
        conn.execute("SELECT py_filename FROM history LIMIT 1")
        conn.execute("ALTER TABLE history RENAME COLUMN py_filename TO input_filename")
        conn.execute(
            "ALTER TABLE history ADD COLUMN conversion_type TEXT NOT NULL DEFAULT 'py'"
        )
        conn.commit()
    except Exception:
        pass
    conn.close()


init_db()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/history")
def history():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM history ORDER BY id DESC LIMIT 50"
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "No file selected"}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ALLOWED_EXT:
        return jsonify({"error": f"Unsupported file type. Allowed: {', '.join(ALLOWED_EXT)}"}), 400

    input_name = file.filename
    conv_type = "html" if ext in (".html", ".htm") else "py"
    now = datetime.now(timezone.utc).isoformat()
    temp_dir = tempfile.mkdtemp()

    try:
        file_path = os.path.join(temp_dir, input_name)
        file.save(file_path)

        if conv_type == "html":
            pptx_path = html_to_pptx(file_path, temp_dir)
        else:
            result = subprocess.run(
                ["python", file_path],
                cwd=temp_dir,
                capture_output=True,
                text=True,
                timeout=30,
            )
            if result.returncode != 0:
                err = result.stderr or "Script failed with no output"
                _save_history(input_name, None, conv_type, "error", err, now)
                return jsonify({"error": err}), 400

            pptx_files = glob.glob(os.path.join(temp_dir, "*.pptx"))
            if not pptx_files:
                err = "Script ran successfully but no .pptx file was generated"
                _save_history(input_name, None, conv_type, "error", err, now)
                return jsonify({"error": err}), 400
            pptx_path = pptx_files[0]

        pptx_name = os.path.basename(pptx_path)
        _save_history(input_name, pptx_name, conv_type, "success", None, now)
        return send_file(pptx_path, as_attachment=True, download_name=pptx_name)

    except subprocess.TimeoutExpired:
        err = "Script timed out after 30 seconds"
        _save_history(input_name, None, conv_type, "error", err, now)
        return jsonify({"error": err}), 400
    except Exception as e:
        _save_history(input_name, None, conv_type, "error", str(e), now)
        return jsonify({"error": str(e)}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _save_history(input_name, pptx_name, conv_type, status, error, created_at):
    conn = get_db()
    conn.execute(
        "INSERT INTO history (input_filename, pptx_filename, conversion_type, status, error, created_at) VALUES (?, ?, ?, ?, ?, ?)",
        (input_name, pptx_name, conv_type, status, error, created_at),
    )
    conn.commit()
    conn.close()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
