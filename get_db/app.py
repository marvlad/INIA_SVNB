# app.py

from flask import Flask, render_template, request
from build_ph_database import build_ph_database
import traceback

app = Flask(__name__)


@app.route("/", methods=["GET", "POST"])
def index():
    log_output = ""
    result = None
    error = None

    default_values = {
        "ph_folder": r"G:\Mi unidad\LABSAF ILLPA\1. Documentos Internos\7.5 Registros Tecnicos\2026\SUELOS\1.pH",
        "database_file": r"G:\Mi unidad\LABSAF ILLPA\ph_database_Ver03.sqlite",
        "csv_file": r"G:\Mi unidad\LABSAF ILLPA\ph_database_Ver03.csv",
        "file_name_filter": "Ver.03",
        "sheet_name": "F-103",
        "first_row": "27",
        "block_size": "21",
        "gap_rows": "2",
        "duplicate_col": "B",
        "code_col": "C",
        "ph_col": "H",
        "verbose": "yes",
    }

    values = default_values.copy()

    if request.method == "POST":
        values["ph_folder"] = request.form.get("ph_folder", "").strip()
        values["database_file"] = request.form.get("database_file", "").strip()
        values["csv_file"] = request.form.get("csv_file", "").strip()
        values["file_name_filter"] = request.form.get("file_name_filter", "Ver.03").strip()
        values["sheet_name"] = request.form.get("sheet_name", "F-103").strip()
        values["first_row"] = request.form.get("first_row", "27").strip()
        values["block_size"] = request.form.get("block_size", "21").strip()
        values["gap_rows"] = request.form.get("gap_rows", "2").strip()
        values["duplicate_col"] = request.form.get("duplicate_col", "B").strip()
        values["code_col"] = request.form.get("code_col", "C").strip()
        values["ph_col"] = request.form.get("ph_col", "H").strip()
        values["verbose"] = request.form.get("verbose", "no")

        logs = []

        def browser_log(message):
            logs.append(str(message))

        try:
            result = build_ph_database(
                ph_folder=values["ph_folder"],
                database_file=values["database_file"],
                csv_file=values["csv_file"],
                file_name_filter=values["file_name_filter"],
                ph_sheet_name=values["sheet_name"],
                ph_first_row=int(values["first_row"]),
                block_size=int(values["block_size"]),
                gap_rows=int(values["gap_rows"]),
                duplicate_col=values["duplicate_col"],
                code_col=values["code_col"],
                ph_value_col=values["ph_col"],
                verbose=(values["verbose"] == "yes"),
                log_callback=browser_log,
            )

        except Exception:
            error = traceback.format_exc()

        log_output = "\n".join(logs)

    return render_template(
        "index.html",
        values=values,
        log_output=log_output,
        result=result,
        error=error,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5001, debug=True)
