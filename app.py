# app.py

from pathlib import Path
import subprocess
import sys
import os
import webbrowser
from threading import Timer

from flask import Flask, request, render_template_string, send_from_directory, abort


app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "update_colorimetric_report.py"
OUTPUT_DIR = BASE_DIR / "output"

HOST = "127.0.0.1"
PORT = 5000
URL = f"http://{HOST}:{PORT}"


HTML_PAGE = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Colorimetric Report Generator</title>

    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 40px;
            background: #f7f7f7;
        }

        .container {
            max-width: 850px;
            margin: auto;
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0, 0, 0, 0.12);
        }

        h1 {
            margin-top: 0;
        }

        label {
            font-weight: bold;
            display: block;
            margin-bottom: 8px;
        }

        select, button {
            font-size: 16px;
            padding: 10px;
            margin-top: 8px;
        }

        button {
            cursor: pointer;
            background: #1f6feb;
            color: white;
            border: none;
            border-radius: 6px;
            padding: 11px 18px;
        }

        button:hover {
            background: #1557b0;
        }

        pre {
            background: #111;
            color: #e6e6e6;
            padding: 15px;
            border-radius: 8px;
            overflow-x: auto;
            white-space: pre-wrap;
        }

        .success {
            color: green;
            font-weight: bold;
        }

        .error {
            color: red;
            font-weight: bold;
        }

        .files {
            margin-top: 20px;
            padding: 15px;
            background: #f1f1f1;
            border-radius: 8px;
        }

        .files a {
            display: block;
            margin: 6px 0;
        }
    </style>
</head>

<body>
<div class="container">
    <h1>Colorimetric Report Generator</h1>

    <form method="post" action="/generate">
        <label for="method">Choose template method:</label>

        <select name="method" id="method" required>
            <option value="Bray">Bray</option>
            <option value="Olsen">Olsen</option>
        </select>

        <br><br>

        <button type="submit">Generate Reports</button>
    </form>

    {% if status %}
        <hr>

        {% if success %}
            <p class="success">{{ status }}</p>
        {% else %}
            <p class="error">{{ status }}</p>
        {% endif %}

        <h3>Terminal output</h3>
        <pre>{{ output }}</pre>
    {% endif %}

    {% if files %}
        <div class="files">
            <h3>Generated files</h3>

            {% for file in files %}
                <a href="/download/{{ method_lower }}/{{ file }}" target="_blank">
                    {{ file }}
                </a>
            {% endfor %}
        </div>
    {% endif %}
</div>
</body>
</html>
"""


def open_browser():
    """
    Open the Flask app automatically in the default browser.
    """
    webbrowser.open_new(URL)


def list_generated_files(method):
    """
    Lists generated XLSM files inside:

        output/bray/
        output/olsen/
    """
    method_lower = method.lower()
    method_dir = OUTPUT_DIR / method_lower

    if not method_dir.exists():
        return []

    files = sorted(
        [
            path.name
            for path in method_dir.glob("*.xlsm")
            if path.is_file()
        ]
    )

    return files


@app.route("/", methods=["GET"])
def index():
    return render_template_string(
        HTML_PAGE,
        status=None,
        success=None,
        output=None,
        files=[],
        method_lower=None,
    )


@app.route("/generate", methods=["POST"])
def generate():
    method = request.form.get("method", "").strip()

    if method not in ["Bray", "Olsen"]:
        return render_template_string(
            HTML_PAGE,
            status="Invalid method selected.",
            success=False,
            output="Method must be Bray or Olsen.",
            files=[],
            method_lower=None,
        )

    if not SCRIPT_PATH.exists():
        return render_template_string(
            HTML_PAGE,
            status="Script not found.",
            success=False,
            output=f"Could not find: {SCRIPT_PATH}",
            files=[],
            method_lower=method.lower(),
        )

    command = [
        sys.executable,
        str(SCRIPT_PATH),
        "--method",
        method,
    ]

    try:
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        result = subprocess.run(
            command,
            cwd=BASE_DIR,
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            check=False,
            env=env,
        )

        output = ""

        if result.stdout:
            output += result.stdout

        if result.stderr:
            output += "\n\nSTDERR:\n"
            output += result.stderr

        success = result.returncode == 0

        if success:
            status = f"{method} reports generated successfully."
        else:
            status = f"{method} report generation failed."

        files = list_generated_files(method)

        return render_template_string(
            HTML_PAGE,
            status=status,
            success=success,
            output=output,
            files=files,
            method_lower=method.lower(),
        )

    except Exception as e:
        return render_template_string(
            HTML_PAGE,
            status="Unexpected error while running the script.",
            success=False,
            output=str(e),
            files=[],
            method_lower=method.lower(),
        )


@app.route("/download/<method>/<filename>", methods=["GET"])
def download_file(method, filename):
    method = method.lower()

    if method not in ["bray", "olsen"]:
        abort(404)

    method_dir = OUTPUT_DIR / method
    file_path = method_dir / filename

    if not file_path.exists():
        abort(404)

    return send_from_directory(
        method_dir,
        filename,
        as_attachment=True,
    )


if __name__ == "__main__":
    Timer(1.0, open_browser).start()
    app.run(host=HOST, port=PORT, debug=False)
