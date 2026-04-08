import os
import glob
import shutil
import sqlite3
import subprocess
import tempfile
from datetime import datetime, timezone

from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute(
        """CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            py_filename TEXT NOT NULL,
            pptx_filename TEXT,
            status TEXT NOT NULL,
            error TEXT,
            created_at TEXT NOT NULL
        )"""
    )
    conn.commit()
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
    if not file.filename or not file.filename.endswith(".py"):
        return jsonify({"error": "Please upload a .py file"}), 400

    py_name = file.filename
    now = datetime.now(timezone.utc).isoformat()
    temp_dir = tempfile.mkdtemp()

    try:
        script_path = os.path.join(temp_dir, py_name)
        file.save(script_path)

        result = subprocess.run(
            ["python", script_path],
            cwd=temp_dir,
            capture_output=True,
            text=True,
            timeout=30,
        )

        if result.returncode != 0:
            err = result.stderr or "Script failed with no output"
            _save_history(py_name, None, "error", err, now)
            return jsonify({"error": err}), 400

        pptx_files = glob.glob(os.path.join(temp_dir, "*.pptx"))
        if not pptx_files:
            err = "Script ran successfully but no .pptx file was generated"
            _save_history(py_name, None, "error", err, now)
            return jsonify({"error": err}), 400

        pptx_path = pptx_files[0]
        pptx_name = os.path.basename(pptx_path)

        _save_history(py_name, pptx_name, "success", None, now)
        return send_file(pptx_path, as_attachment=True, download_name=pptx_name)

    except subprocess.TimeoutExpired:
        err = "Script timed out after 30 seconds"
        _save_history(py_name, None, "error", err, now)
        return jsonify({"error": err}), 400
    except Exception as e:
        _save_history(py_name, None, "error", str(e), now)
        return jsonify({"error": str(e)}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _save_history(py_name, pptx_name, status, error, created_at):
    conn = get_db()
    conn.execute(
        "INSERT INTO history (py_filename, pptx_filename, status, error, created_at) VALUES (?, ?, ?, ?, ?)",
        (py_name, pptx_name, status, error, created_at),
    )
    conn.commit()
    conn.close()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
