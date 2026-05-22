# app.py
# Main INIA_SVNB browser app
#
# This root app controls:
#   1. Spectrometer CSV -> colorimetric report XLSM
#   2. pH database update
#   3. Add pH into generated XLSM reports
#
# Expected structure:
#
# INIA_SVNB/
# ├── app.py
# ├── update_colorimetric_report.py
# ├── input/
# ├── input.dat
# ├── output/
# ├── get_db/
# │   └── build_ph_database.py
# └── fill_ph/
#     └── fill_ph_from_dat.py


from pathlib import Path
import subprocess
import sys
import os
import shutil
import traceback
import webbrowser
from threading import Timer

from flask import Flask, request, render_template_string, send_from_directory, abort


# ============================================================
# BASIC PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"
INPUT_DAT = BASE_DIR / "input.dat"

OUTPUT_DIR = BASE_DIR / "output"

UPDATE_COLOR_SCRIPT = BASE_DIR / "update_colorimetric_report.py"

GET_DB_DIR = BASE_DIR / "get_db"
BUILD_PH_SCRIPT = GET_DB_DIR / "build_ph_database.py"

FILL_PH_DIR = BASE_DIR / "fill_ph"
FILL_PH_SCRIPT = FILL_PH_DIR / "fill_ph_from_dat.py"

HOST = "127.0.0.1"
PORT = 5000
URL = f"http://{HOST}:{PORT}"


# ============================================================
# FLASK APP
# ============================================================

app = Flask(__name__)


# ============================================================
# HTML
# ============================================================

HTML_PAGE = """
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>INIA SVNB Report Pipeline</title>

    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 35px;
            background: #f7f7f7;
        }

        .container {
            max-width: 1150px;
            margin: auto;
            background: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 2px 12px rgba(0, 0, 0, 0.12);
        }

        h1 {
            margin-top: 0;
        }

        h2 {
            margin-top: 30px;
            border-bottom: 1px solid #ddd;
            padding-bottom: 6px;
        }

        label {
            font-weight: bold;
            display: block;
            margin-top: 14px;
            margin-bottom: 6px;
        }

        input[type="text"],
        input[type="number"],
        select {
            width: 100%;
            box-sizing: border-box;
            font-size: 15px;
            padding: 9px;
            border: 1px solid #bbb;
            border-radius: 6px;
        }

        button {
            cursor: pointer;
            background: #1f6feb;
            color: white;
            border: none;
            border-radius: 6px;
            padding: 12px 18px;
            font-size: 15px;
            font-weight: bold;
            margin-top: 16px;
            margin-right: 8px;
        }

        button:hover {
            background: #1557b0;
        }

        .danger-button {
            background: #d93025;
        }

        .danger-button:hover {
            background: #a61b14;
        }

        .green-button {
            background: #2f9e44;
        }

        .green-button:hover {
            background: #237032;
        }

        .purple-button {
            background: #6f42c1;
        }

        .purple-button:hover {
            background: #553098;
        }

        pre {
            background: #111;
            color: #e6e6e6;
            padding: 15px;
            border-radius: 8px;
            overflow-x: auto;
            white-space: pre-wrap;
            max-height: 650px;
        }

        .success {
            color: green;
            font-weight: bold;
        }

        .error {
            color: red;
            font-weight: bold;
        }

        .note {
            background: #fff7d6;
            border-left: 5px solid #e6b800;
            padding: 12px;
            margin-bottom: 20px;
            color: #333;
        }

        .box {
            background: #f1f1f1;
            padding: 16px;
            border-radius: 8px;
            margin-top: 16px;
        }

        .small {
            color: #666;
            font-size: 13px;
            margin-top: 4px;
        }

        .row {
            display: flex;
            gap: 18px;
        }

        .col {
            flex: 1;
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

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 16px;
            font-size: 14px;
        }

        th, td {
            border: 1px solid #ccc;
            padding: 7px;
            text-align: left;
        }

        th {
            background: #eee;
        }
    </style>
</head>

<body>
<div class="container">

    <h1>INIA SVNB Report Pipeline</h1>

    <div class="note">
        This main app runs the full workflow from the root folder:
        <br>
        <b>spectrometer CSV → colorimetric XLSM → pH database → final XLSM with pH</b>
    </div>

    <form method="post">

        <h2>1. Spectrometer CSV input</h2>

        <label>Spectrometer CSV file path</label>
        <input
            type="text"
            name="spectrometer_csv"
            value="{{ values.spectrometer_csv }}"
            placeholder="G:\\Mi unidad\\LABSAF ILLPA\\Cuantificación_23_03_2026_12_52_11.csv"
        >
        <div class="small">
            The app will copy this file into <b>INIA_SVNB/input/</b> and write <b>input.dat</b>.
            If you leave it empty, the app will use the existing files already in <b>input/</b>.
        </div>

        <label>Colorimetric method</label>
        <select name="method">
            <option value="Bray" {% if values.method == "Bray" %}selected{% endif %}>Bray</option>
            <option value="Olsen" {% if values.method == "Olsen" %}selected{% endif %}>Olsen</option>
            <option value="Both" {% if values.method == "Both" %}selected{% endif %}>Both Bray and Olsen</option>
        </select>

        <h2>2. pH database settings</h2>

        <label>Folder with pH Excel files</label>
        <input
            type="text"
            name="ph_folder"
            value="{{ values.ph_folder }}"
            placeholder="G:\\Mi unidad\\LABSAF ILLPA\\1. Documentos Internos\\7.5 Registros Tecnicos\\2026\\SUELOS\\1.pH"
        >

        <div class="row">
            <div class="col">
                <label>pH SQLite output</label>
                <input
                    type="text"
                    name="ph_database_file"
                    value="{{ values.ph_database_file }}"
                >
            </div>

            <div class="col">
                <label>pH CSV output</label>
                <input
                    type="text"
                    name="ph_csv_file"
                    value="{{ values.ph_csv_file }}"
                >
            </div>
        </div>

        <div class="row">
            <div class="col">
                <label>pH filename filter</label>
                <input
                    type="text"
                    name="ph_file_filter"
                    value="{{ values.ph_file_filter }}"
                >
            </div>

            <div class="col">
                <label>pH sheet name</label>
                <input
                    type="text"
                    name="ph_sheet_name"
                    value="{{ values.ph_sheet_name }}"
                >
            </div>
        </div>

        <h2>3. Fill pH into generated reports</h2>

        <div class="row">
            <div class="col">
                <label>Generated report sheet name</label>
                <input
                    type="text"
                    name="fill_sheet_name"
                    value="{{ values.fill_sheet_name }}"
                >
            </div>

            <div class="col">
                <label>Code column</label>
                <input
                    type="text"
                    name="fill_code_col"
                    value="{{ values.fill_code_col }}"
                >
            </div>

            <div class="col">
                <label>pH output column</label>
                <input
                    type="text"
                    name="fill_output_col"
                    value="{{ values.fill_output_col }}"
                >
            </div>
        </div>

        <div class="row">
            <div class="col">
                <label>First row</label>
                <input
                    type="number"
                    name="fill_first_row"
                    value="{{ values.fill_first_row }}"
                >
            </div>

            <div class="col">
                <label>Block size</label>
                <input
                    type="number"
                    name="fill_block_size"
                    value="{{ values.fill_block_size }}"
                >
            </div>

            <div class="col">
                <label>Gap rows</label>
                <input
                    type="number"
                    name="fill_gap_rows"
                    value="{{ values.fill_gap_rows }}"
                >
            </div>
        </div>

        <h2>Actions</h2>

        <button name="action" value="generate_reports">
            1. Generate colorimetric reports
        </button>

        <button name="action" value="update_ph_db" class="danger-button">
            2. Update pH database
        </button>

        <button name="action" value="fill_ph" class="green-button">
            3. Add pH to generated reports
        </button>

        <button name="action" value="full_pipeline" class="purple-button">
            Run full pipeline
        </button>

    </form>

    {% if status %}
        <hr>
        {% if success %}
            <p class="success">{{ status }}</p>
        {% else %}
            <p class="error">{{ status }}</p>
        {% endif %}
    {% endif %}

    {% if output %}
        <h2>Terminal output</h2>
        <pre>{{ output }}</pre>
    {% endif %}

    {% if files %}
        <div class="files">
            <h2>Generated final files with pH</h2>

            {% for item in files %}
                <a href="/download/{{ item.method }}/{{ item.filename }}" target="_blank">
                    {{ item.method }}/with_ph/{{ item.filename }}
                </a>
            {% endfor %}
        </div>
    {% endif %}

</div>
</body>
</html>
"""


# ============================================================
# SMALL HELPERS
# ============================================================

def run_command(command, cwd):
    """
    Run a Python command and return output text plus success status.
    """
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    result = subprocess.run(
        command,
        cwd=cwd,
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

    return success, output


def copy_spectrometer_csv_to_input(csv_path_text):
    """
    Copy the selected spectrometer CSV to INIA_SVNB/input/
    and write INIA_SVNB/input.dat with that CSV filename.

    If csv_path_text is empty, this function does nothing.
    """
    if not csv_path_text.strip():
        return True, "No CSV path provided. Using existing input/ and input.dat.\n"

    source_csv = Path(csv_path_text.strip())

    if not source_csv.exists():
        return False, f"Spectrometer CSV not found:\n  {source_csv}\n"

    if source_csv.suffix.lower() != ".csv":
        return False, f"Selected file is not a CSV:\n  {source_csv}\n"

    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    destination_csv = INPUT_DIR / source_csv.name

    shutil.copy2(source_csv, destination_csv)

    with open(INPUT_DAT, "w", encoding="utf-8") as f:
        f.write(source_csv.name + "\n")

    msg = ""
    msg += "Copied spectrometer CSV:\n"
    msg += f"  From: {source_csv}\n"
    msg += f"  To:   {destination_csv}\n"
    msg += "Updated input.dat:\n"
    msg += f"  {INPUT_DAT}\n"
    msg += f"  Content: {source_csv.name}\n"

    return True, msg


def get_methods(method):
    """
    Convert selected method into list.
    """
    if method == "Both":
        return ["Bray", "Olsen"]

    return [method]


def generate_colorimetric_reports(method, spectrometer_csv):
    """
    Run update_colorimetric_report.py for Bray/Olsen/Both.
    """
    output = ""

    ok, msg = copy_spectrometer_csv_to_input(spectrometer_csv)
    output += msg

    if not ok:
        return False, output

    if not UPDATE_COLOR_SCRIPT.exists():
        return False, f"Script not found:\n  {UPDATE_COLOR_SCRIPT}\n"

    all_success = True

    for one_method in get_methods(method):
        output += "\n"
        output += "============================================================\n"
        output += f"GENERATING COLORIMETRIC REPORTS: {one_method}\n"
        output += "============================================================\n"

        command = [
            sys.executable,
            str(UPDATE_COLOR_SCRIPT),
            "--method",
            one_method,
        ]

        success, cmd_output = run_command(command, cwd=BASE_DIR)

        output += cmd_output

        if not success:
            all_success = False
            output += f"\nFAILED generating {one_method} reports.\n"
        else:
            output += f"\nOK: {one_method} reports generated.\n"

    return all_success, output


def update_ph_database(values):
    """
    Run get_db/build_ph_database.py with external arguments.
    """
    if not BUILD_PH_SCRIPT.exists():
        return False, f"Script not found:\n  {BUILD_PH_SCRIPT}\n"

    ph_folder = values["ph_folder"].strip()

    if not ph_folder:
        return False, "pH folder is empty. Please provide the folder containing pH Excel files.\n"

    command = [
        sys.executable,
        str(BUILD_PH_SCRIPT),

        "--ph-folder",
        ph_folder,

        "--database-file",
        values["ph_database_file"],

        "--csv-file",
        values["ph_csv_file"],

        "--file-name-filter",
        values["ph_file_filter"],

        "--sheet-name",
        values["ph_sheet_name"],

        "--quiet",
    ]

    success, output = run_command(command, cwd=GET_DB_DIR)

    return success, output


def make_input_for_ph_dat_from_generated_reports(method):
    """
    Create a temporary input_for_ph.dat using all generated XLSM files
    from output/bray or output/olsen.

    The fill_ph script will read this DAT file and process those XLSM files.
    """
    method_lower = method.lower()

    method_output_dir = OUTPUT_DIR / method_lower

    if not method_output_dir.exists():
        raise FileNotFoundError(
            f"Generated report folder not found: {method_output_dir}"
        )

    xlsm_files = sorted(
        [
            path
            for path in method_output_dir.glob("*.xlsm")
            if path.is_file()
        ]
    )

    if not xlsm_files:
        raise FileNotFoundError(
            f"No generated .xlsm files found in: {method_output_dir}"
        )

    dat_file = method_output_dir / "input_for_ph.dat"

    with open(dat_file, "w", encoding="utf-8") as f:
        for xlsm_file in xlsm_files:
            f.write(xlsm_file.name + "\n")

    return dat_file, method_output_dir, xlsm_files


def fill_ph_into_generated_reports(method, values):
    """
    Run fill_ph/fill_ph_from_dat.py over generated XLSM files.

    Input:
        output/bray/*.xlsm
        output/olsen/*.xlsm

    Output:
        output/bray/with_ph/*.xlsm
        output/olsen/with_ph/*.xlsm
    """
    if not FILL_PH_SCRIPT.exists():
        return False, f"Script not found:\n  {FILL_PH_SCRIPT}\n"

    ph_csv_file = Path(values["ph_csv_file"])

    if not ph_csv_file.exists():
        return False, f"pH CSV file not found:\n  {ph_csv_file}\n"

    all_success = True
    output = ""

    for one_method in get_methods(method):
        method_lower = one_method.lower()

        output += "\n"
        output += "============================================================\n"
        output += f"ADDING pH TO GENERATED REPORTS: {one_method}\n"
        output += "============================================================\n"

        try:
            dat_file, generated_dir, xlsm_files = make_input_for_ph_dat_from_generated_reports(one_method)

            with_ph_dir = generated_dir / "with_ph"
            with_ph_dir.mkdir(parents=True, exist_ok=True)

            output += f"Generated report folder:\n  {generated_dir}\n"
            output += f"Temporary input_for_ph.dat:\n  {dat_file}\n"
            output += f"Files to process: {len(xlsm_files)}\n"
            output += f"Final output folder:\n  {with_ph_dir}\n"

            command = [
                sys.executable,
                str(FILL_PH_SCRIPT),

                "--input-dat",
                str(dat_file),

                "--input-dir",
                str(generated_dir),

                "--ph-csv",
                str(ph_csv_file),

                "--output-dir",
                str(with_ph_dir),

                "--sheet-name",
                values["fill_sheet_name"],

                "--first-row",
                str(values["fill_first_row"]),

                "--block-size",
                str(values["fill_block_size"]),

                "--gap-rows",
                str(values["fill_gap_rows"]),

                "--code-col",
                values["fill_code_col"],

                "--output-col",
                values["fill_output_col"],

                "--quiet",
            ]

            success, cmd_output = run_command(command, cwd=FILL_PH_DIR)

            output += cmd_output

            if not success:
                all_success = False
                output += f"\nFAILED adding pH to {one_method} reports.\n"
            else:
                output += f"\nOK: pH added to {one_method} reports.\n"

        except Exception:
            all_success = False
            output += "\nERROR while preparing pH filling:\n"
            output += traceback.format_exc()

    return all_success, output


def list_final_with_ph_files():
    """
    List files in:
        output/bray/with_ph/
        output/olsen/with_ph/
    """
    results = []

    for method_lower in ["bray", "olsen"]:
        folder = OUTPUT_DIR / method_lower / "with_ph"

        if not folder.exists():
            continue

        for path in sorted(folder.glob("*.xlsm")):
            if path.is_file():
                results.append(
                    {
                        "method": method_lower,
                        "filename": path.name,
                    }
                )

    return results


def open_browser():
    webbrowser.open_new(URL)


def get_default_values():
    return {
        "spectrometer_csv": "",
        "method": "Bray",

        "ph_folder": r"G:\Mi unidad\LABSAF ILLPA\1. Documentos Internos\7.5 Registros Tecnicos\2026\SUELOS\1.pH",
        "ph_database_file": str(BASE_DIR / "ph_database_Ver03.sqlite"),
        "ph_csv_file": str(BASE_DIR / "ph_database_Ver03.csv"),
        "ph_file_filter": "Ver.03",
        "ph_sheet_name": "F-103",

        "fill_sheet_name": "P_DIS",
        "fill_code_col": "C",
        "fill_output_col": "E",
        "fill_first_row": "37",
        "fill_block_size": "20",
        "fill_gap_rows": "3",
    }


def get_values_from_form():
    defaults = get_default_values()

    values = {}

    for key, default in defaults.items():
        values[key] = request.form.get(key, default).strip()

    return values


# ============================================================
# ROUTES
# ============================================================

@app.route("/", methods=["GET", "POST"])
def index():
    status = None
    success = None
    output = ""
    files = list_final_with_ph_files()

    if request.method == "GET":
        values = get_default_values()

        return render_template_string(
            HTML_PAGE,
            values=values,
            status=status,
            success=success,
            output=output,
            files=files,
        )

    values = get_values_from_form()
    action = request.form.get("action", "").strip()

    try:
        if action == "generate_reports":
            success, output = generate_colorimetric_reports(
                method=values["method"],
                spectrometer_csv=values["spectrometer_csv"],
            )

            if success:
                status = "Colorimetric reports generated successfully."
            else:
                status = "Colorimetric report generation failed."

        elif action == "update_ph_db":
            success, output = update_ph_database(values)

            if success:
                status = "pH database updated successfully."
            else:
                status = "pH database update failed."

        elif action == "fill_ph":
            success, output = fill_ph_into_generated_reports(
                method=values["method"],
                values=values,
            )

            if success:
                status = "pH added to generated reports successfully."
            else:
                status = "Adding pH to generated reports failed."

        elif action == "full_pipeline":
            full_output = ""
            full_success = True

            s1, out1 = generate_colorimetric_reports(
                method=values["method"],
                spectrometer_csv=values["spectrometer_csv"],
            )

            full_output += "\n\n================ STEP 1: GENERATE REPORTS ================\n"
            full_output += out1

            if not s1:
                full_success = False

            s2, out2 = update_ph_database(values)

            full_output += "\n\n================ STEP 2: UPDATE pH DATABASE ================\n"
            full_output += out2

            if not s2:
                full_success = False

            s3, out3 = fill_ph_into_generated_reports(
                method=values["method"],
                values=values,
            )

            full_output += "\n\n================ STEP 3: ADD pH TO REPORTS ================\n"
            full_output += out3

            if not s3:
                full_success = False

            success = full_success
            output = full_output

            if success:
                status = "Full pipeline finished successfully."
            else:
                status = "Full pipeline finished with errors. Check the output."

        else:
            success = False
            status = "Invalid action."
            output = f"Unknown action: {action}"

    except Exception:
        success = False
        status = "Unexpected error."
        output = traceback.format_exc()

    files = list_final_with_ph_files()

    return render_template_string(
        HTML_PAGE,
        values=values,
        status=status,
        success=success,
        output=output,
        files=files,
    )


@app.route("/download/<method>/<filename>", methods=["GET"])
def download_file(method, filename):
    method = method.lower()

    if method not in ["bray", "olsen"]:
        abort(404)

    folder = OUTPUT_DIR / method / "with_ph"
    file_path = folder / filename

    if not file_path.exists():
        abort(404)

    return send_from_directory(
        folder,
        filename,
        as_attachment=True,
    )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    Timer(1.0, open_browser).start()

    app.run(
        host=HOST,
        port=PORT,
        debug=False,
    )
