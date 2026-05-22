# app.py

from flask import Flask, request, render_template_string
import traceback

from fill_ph_from_dat import fill_many_xlsm_from_dat


app = Flask(__name__)


HTML_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Fill pH into XLSM files</title>

    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f4f6f8;
            margin: 0;
            padding: 0;
        }

        .container {
            max-width: 1150px;
            margin: 30px auto;
            background: white;
            padding: 25px 35px;
            border-radius: 12px;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.12);
        }

        h1 {
            margin-top: 0;
            color: #222;
        }

        h2 {
            margin-top: 30px;
            color: #333;
            border-bottom: 1px solid #ddd;
            padding-bottom: 6px;
        }

        label {
            display: block;
            margin-top: 14px;
            font-weight: bold;
            color: #333;
        }

        input[type="text"],
        input[type="number"] {
            width: 100%;
            padding: 9px;
            margin-top: 5px;
            box-sizing: border-box;
            border: 1px solid #bbb;
            border-radius: 6px;
            font-size: 14px;
        }

        .row {
            display: flex;
            gap: 20px;
        }

        .col {
            flex: 1;
        }

        button {
            margin-top: 25px;
            padding: 12px 24px;
            background: #1f6feb;
            color: white;
            border: none;
            border-radius: 7px;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
        }

        button:hover {
            background: #1557b0;
        }

        .note {
            background: #fff7d6;
            border-left: 5px solid #e6b800;
            padding: 12px;
            margin-bottom: 20px;
            color: #333;
        }

        .success {
            background: #e7f7ec;
            border-left: 5px solid #2f9e44;
            padding: 12px;
            margin-top: 25px;
            color: #1b5e20;
        }

        .error {
            background: #fdecea;
            border-left: 5px solid #d93025;
            padding: 12px;
            margin-top: 25px;
            color: #8a1f11;
            white-space: pre-wrap;
            font-family: Consolas, monospace;
            font-size: 13px;
        }

        pre {
            background: #111827;
            color: #d1d5db;
            padding: 18px;
            border-radius: 8px;
            overflow-x: auto;
            white-space: pre-wrap;
            font-size: 13px;
            line-height: 1.45;
            max-height: 650px;
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
            background: #f0f0f0;
        }

        .small {
            font-size: 13px;
            color: #666;
            margin-top: 5px;
        }

        .checkbox-row {
            margin-top: 18px;
        }
    </style>
</head>

<body>
<div class="container">
    <h1>Fill pH into XLSM files</h1>

    <div class="note">
        This app reads <b>input_for_ph.dat</b>, fills pH values into every listed XLSM file,
        and saves the results in <b>with_ph</b> using the same original filename.
        The browser cannot safely select folder paths directly, so paste the Windows paths below.
    </div>

    <form method="POST">

        <h2>Main paths</h2>

        <label>input_for_ph.dat</label>
        <input
            type="text"
            name="input_dat"
            value="{{ values.input_dat }}"
            placeholder="G:\\Mi unidad\\LABSAF ILLPA\\input_for_ph.dat"
            required
        >
        <div class="small">This file should contain one XLSM filename per line.</div>

        <label>Input folder containing the XLSM files</label>
        <input
            type="text"
            name="input_dir"
            value="{{ values.input_dir }}"
            placeholder="G:\\Mi unidad\\LABSAF ILLPA"
            required
        >

        <label>pH CSV database</label>
        <input
            type="text"
            name="ph_csv"
            value="{{ values.ph_csv }}"
            placeholder="G:\\Mi unidad\\LABSAF ILLPA\\ph_database_Ver03.csv"
            required
        >

        <label>Output folder</label>
        <input
            type="text"
            name="output_dir"
            value="{{ values.output_dir }}"
            placeholder="G:\\Mi unidad\\LABSAF ILLPA\\with_ph"
            required
        >
        <div class="small">Outputs will be saved as: output folder / same input filename.</div>

        <h2>Workbook settings</h2>

        <div class="row">
            <div class="col">
                <label>Sheet name</label>
                <input
                    type="text"
                    name="sheet_name"
                    value="{{ values.sheet_name }}"
                    required
                >
            </div>

            <div class="col">
                <label>First row</label>
                <input
                    type="number"
                    name="first_row"
                    value="{{ values.first_row }}"
                    required
                >
            </div>

            <div class="col">
                <label>Block size</label>
                <input
                    type="number"
                    name="block_size"
                    value="{{ values.block_size }}"
                    required
                >
            </div>

            <div class="col">
                <label>Gap rows</label>
                <input
                    type="number"
                    name="gap_rows"
                    value="{{ values.gap_rows }}"
                    required
                >
            </div>
        </div>

        <div class="row">
            <div class="col">
                <label>Code column</label>
                <input
                    type="text"
                    name="code_col"
                    value="{{ values.code_col }}"
                    required
                >
            </div>

            <div class="col">
                <label>pH output column</label>
                <input
                    type="text"
                    name="output_col"
                    value="{{ values.output_col }}"
                    required
                >
            </div>
        </div>

        <div class="checkbox-row">
            <label>
                <input
                    type="checkbox"
                    name="verbose"
                    value="yes"
                    {% if values.verbose == "yes" %}checked{% endif %}
                >
                Show detailed row-by-row log
            </label>
        </div>

        <button type="submit">Fill pH values</button>
    </form>

    {% if result %}
        <div class="success">
            <strong>Finished.</strong><br>
            Total files listed: {{ result.total_files }}<br>
            Successful files: {{ result.successful_files }}<br>
            Failed files: {{ result.failed_files }}<br>
            Total pH written: {{ result.total_written }}<br>
            Total not found: {{ result.total_not_found }}<br>
            Total skipped: {{ result.total_skipped }}<br>
            Output directory: {{ result.output_dir }}
        </div>

        <h2>File results</h2>

        <table>
            <thead>
                <tr>
                    <th>Status</th>
                    <th>Input</th>
                    <th>Output</th>
                    <th>Written</th>
                    <th>Not found</th>
                    <th>Skipped</th>
                    <th>Error</th>
                </tr>
            </thead>
            <tbody>
                {% for item in result.file_results %}
                <tr>
                    <td>{{ item.status }}</td>
                    <td>{{ item.input }}</td>
                    <td>{{ item.output }}</td>
                    <td>{{ item.written }}</td>
                    <td>{{ item.not_found }}</td>
                    <td>{{ item.skipped }}</td>
                    <td>{{ item.error }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    {% endif %}

    {% if error %}
        <h2>Error</h2>
        <div class="error">{{ error }}</div>
    {% endif %}

    {% if log_output %}
        <h2>Log</h2>
        <pre>{{ log_output }}</pre>
    {% endif %}
</div>
</body>
</html>
"""


@app.route("/", methods=["GET", "POST"])
def index():
    log_output = ""
    result = None
    error = None

    default_values = {
        "input_dat": r"G:\Mi unidad\LABSAF ILLPA\input_for_ph.dat",
        "input_dir": r"G:\Mi unidad\LABSAF ILLPA",
        "ph_csv": r"G:\Mi unidad\LABSAF ILLPA\ph_database_Ver03.csv",
        "output_dir": r"G:\Mi unidad\LABSAF ILLPA\with_ph",
        "sheet_name": "P_DIS",
        "first_row": "37",
        "block_size": "20",
        "gap_rows": "3",
        "code_col": "C",
        "output_col": "E",
        "verbose": "yes",
    }

    values = default_values.copy()

    if request.method == "POST":
        values["input_dat"] = request.form.get("input_dat", "").strip()
        values["input_dir"] = request.form.get("input_dir", "").strip()
        values["ph_csv"] = request.form.get("ph_csv", "").strip()
        values["output_dir"] = request.form.get("output_dir", "").strip()
        values["sheet_name"] = request.form.get("sheet_name", "P_DIS").strip()
        values["first_row"] = request.form.get("first_row", "37").strip()
        values["block_size"] = request.form.get("block_size", "20").strip()
        values["gap_rows"] = request.form.get("gap_rows", "3").strip()
        values["code_col"] = request.form.get("code_col", "C").strip()
        values["output_col"] = request.form.get("output_col", "E").strip()
        values["verbose"] = request.form.get("verbose", "no")

        logs = []

        def browser_log(message):
            logs.append(str(message))

        try:
            result = fill_many_xlsm_from_dat(
                input_dat_file=values["input_dat"],
                input_dir=values["input_dir"],
                ph_csv=values["ph_csv"],
                output_dir=values["output_dir"],
                input_sheet_name=values["sheet_name"],
                first_row=int(values["first_row"]),
                block_size=int(values["block_size"]),
                gap_rows=int(values["gap_rows"]),
                input_code_col=values["code_col"],
                input_output_col=values["output_col"],
                verbose=(values["verbose"] == "yes"),
                log_callback=browser_log,
            )

        except Exception:
            error = traceback.format_exc()

        log_output = "\n".join(logs)

    return render_template_string(
        HTML_PAGE,
        values=values,
        log_output=log_output,
        result=result,
        error=error,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5001, debug=True)
