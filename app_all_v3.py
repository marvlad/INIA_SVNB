# all_app_v3.py
#
# Main INIA_SVNB browser app
#
# Workflow:
#   1. Read existing input.dat with spectrometer CSV names
#   2. Run update_colorimetric_report.py
#   3. Create input_for_ph.dat from generated Analizado_*.xlsm files
#   4. Update pH database with fixed names in the root directory
#   5. Fill pH into generated reports
#   6. Save final files in final_report/bray or final_report/olsen
#
# IMPORTANT:
#
#   output/bray/ and output/olsen/
#       contain temporary/intermediate reports WITHOUT pH.
#       These may be cleaned before a new run.
#
#   final_report/bray/ and final_report/olsen/
#       contain completed reports WITH pH.
#       These directories are NEVER cleaned by this application.


from pathlib import Path
import subprocess
import sys
import os
import traceback
import webbrowser
from threading import Timer

from flask import (
    Flask,
    request,
    render_template_string,
    send_from_directory,
    abort,
)


# ============================================================
# BASIC PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_DIR = BASE_DIR / "input"
INPUT_DAT = BASE_DIR / "input.dat"

OUTPUT_DIR = BASE_DIR / "output"
FINAL_REPORT_DIR = BASE_DIR / "final_report"

# Banner image:
# INIA_SVNB/banner/banner.jpeg
BANNER_DIR = BASE_DIR / "banner"

UPDATE_COLOR_SCRIPT = (
    BASE_DIR
    / "update_colorimetric_report.py"
)

GET_DB_DIR = (
    BASE_DIR
    / "get_db"
)

BUILD_PH_SCRIPT = (
    GET_DB_DIR
    / "build_ph_database.py"
)

FILL_PH_DIR = (
    BASE_DIR
    / "fill_ph"
)

FILL_PH_SCRIPT = (
    FILL_PH_DIR
    / "fill_ph_from_dat.py"
)

PH_DATABASE_FILE = (
    BASE_DIR
    / "ph_database_Ver03.sqlite"
)

PH_CSV_FILE = (
    BASE_DIR
    / "ph_database_Ver03.csv"
)

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

<meta
    name="viewport"
    content="width=device-width, initial-scale=1"
>

<title>PICASSO | INIA P Report Tool</title>


<style>

:root {
    --blue: #1f6feb;
    --blue-dark: #174ea6;

    --green: #2f9e44;
    --red: #d93025;
    --purple: #6f42c1;

    --background: #f4f7fb;
    --text: #263238;
    --muted: #667085;

    --border: #d7dee8;
}


/* ============================================================
   GENERAL PAGE
   ============================================================ */

* {
    box-sizing: border-box;
}


body {
    font-family:
        Arial,
        Helvetica,
        sans-serif;

    margin: 0;

    padding: 32px;

    background:
        var(--background);

    color:
        var(--text);
}


.container {
    max-width: 1150px;

    margin: auto;

    background:
        white;

    padding: 30px;

    border-radius: 18px;

    box-shadow:
        0 8px 30px
        rgba(20, 45, 75, 0.10);
}


/* ============================================================
   PICASSO BANNER
   ============================================================ */

.banner {
    position: relative;

    min-height: 270px;

    display: flex;

    align-items: center;

    overflow: hidden;

    margin:
        -6px
        -6px
        30px
        -6px;

    padding:
        40px
        45px;

    border-radius:
        20px;

    background-image:

        linear-gradient(
            90deg,
            rgba(5, 19, 38, 0.86) 0%,
            rgba(9, 38, 70, 0.72) 46%,
            rgba(9, 38, 70, 0.34) 100%
        ),

        url("/banner/banner.jpeg");

    background-size:
        cover;

    background-position:
        center;

    background-repeat:
        no-repeat;

    color:
        white;

    box-shadow:
        0 14px 32px
        rgba(0, 0, 0, 0.20);
}


.banner::after {
    content: "";

    position: absolute;

    left: 0;
    right: 0;
    bottom: 0;

    height: 50%;

    background:
        linear-gradient(
            to top,
            rgba(0, 0, 0, 0.30),
            transparent
        );

    pointer-events:
        none;
}


.banner-content {
    position: relative;

    z-index: 2;

    max-width: 820px;
}


.original-title {
    display: inline-block;

    margin-bottom: 12px;

    padding:
        7px
        12px;

    border:
        1px solid
        rgba(255, 255, 255, 0.35);

    border-radius:
        20px;

    background:
        rgba(255, 255, 255, 0.13);

    backdrop-filter:
        blur(5px);

    font-size:
        14px;

    font-weight:
        600;

    letter-spacing:
        0.3px;
}


.picasso-title {
    margin: 0;

    font-size:
        58px;

    line-height:
        1;

    font-weight:
        800;

    letter-spacing:
        6px;

    text-shadow:
        0 3px 12px
        rgba(0, 0, 0, 0.40);
}


.picasso-acronym {
    margin-top:
        17px;

    max-width:
        780px;

    font-size:
        19px;

    line-height:
        1.55;

    font-weight:
        400;

    color:
        rgba(255, 255, 255, 0.96);

    text-shadow:
        0 2px 8px
        rgba(0, 0, 0, 0.45);
}


.acro-letter {
    display:
        inline-block;

    font-size:
        24px;

    font-weight:
        800;

    color:
        #ffffff;

    margin-right:
        1px;
}


.banner-subtitle {
    margin-top:
        14px;

    font-size:
        14px;

    color:
        rgba(255, 255, 255, 0.82);
}


/* ============================================================
   HEADINGS
   ============================================================ */

h2 {
    margin-top:
        32px;

    border-bottom:
        1px solid var(--border);

    padding-bottom:
        7px;

    color:
        #24364b;
}


/* ============================================================
   FORM
   ============================================================ */

label {
    font-weight:
        bold;

    display:
        block;

    margin-top:
        14px;

    margin-bottom:
        6px;
}


input[type="text"],
input[type="number"],
select {
    width:
        100%;

    box-sizing:
        border-box;

    font-size:
        15px;

    padding:
        10px;

    border:
        1px solid #bbb;

    border-radius:
        7px;

    background:
        white;
}


input[type="text"]:focus,
input[type="number"]:focus,
select:focus {
    outline:
        none;

    border-color:
        #4389e8;

    box-shadow:
        0 0 0 3px
        rgba(31, 111, 235, 0.12);
}


/* ============================================================
   BUTTONS
   ============================================================ */

button {
    cursor:
        pointer;

    background:
        var(--blue);

    color:
        white;

    border:
        none;

    border-radius:
        7px;

    padding:
        12px
        18px;

    font-size:
        15px;

    font-weight:
        bold;

    margin-top:
        16px;

    margin-right:
        8px;

    transition:
        transform 0.15s ease,
        box-shadow 0.15s ease,
        background 0.15s ease;
}


button:hover {
    background:
        var(--blue-dark);

    transform:
        translateY(-1px);

    box-shadow:
        0 5px 12px
        rgba(0, 0, 0, 0.14);
}


.danger-button {
    background:
        var(--red);
}


.danger-button:hover {
    background:
        #a61b14;
}


.green-button {
    background:
        var(--green);
}


.green-button:hover {
    background:
        #237032;
}


.purple-button {
    background:
        var(--purple);
}


.purple-button:hover {
    background:
        #553098;
}


/* ============================================================
   INFORMATION BOXES
   ============================================================ */

.note {
    background:
        #fff8df;

    border-left:
        5px solid
        #e6b800;

    padding:
        14px 16px;

    margin-bottom:
        20px;

    color:
        #333;

    border-radius:
        0 8px 8px 0;
}


.small {
    color:
        var(--muted);

    font-size:
        13px;

    line-height:
        1.5;

    margin-top:
        6px;
}


.path-box {
    background:
        #f2f4f7;

    padding:
        11px;

    border-radius:
        7px;

    font-family:
        Consolas,
        monospace;

    font-size:
        13px;

    margin-top:
        8px;

    overflow-wrap:
        anywhere;
}


/* ============================================================
   ROWS / COLUMNS
   ============================================================ */

.row {
    display:
        flex;

    gap:
        18px;
}


.col {
    flex:
        1;
}


/* ============================================================
   STATUS
   ============================================================ */

.success {
    color:
        green;

    font-weight:
        bold;
}


.error {
    color:
        red;

    font-weight:
        bold;
}


/* ============================================================
   TERMINAL
   ============================================================ */

pre {
    background:
        #111;

    color:
        #e6e6e6;

    padding:
        15px;

    border-radius:
        8px;

    overflow-x:
        auto;

    white-space:
        pre-wrap;

    max-height:
        650px;
}


/* ============================================================
   FILES
   ============================================================ */

.files {
    margin-top:
        20px;

    padding:
        15px;

    background:
        #f1f1f1;

    border-radius:
        8px;
}


.files a {
    display:
        block;

    margin:
        7px 0;

    color:
        #1f6feb;

    text-decoration:
        none;
}


.files a:hover {
    text-decoration:
        underline;
}


/* ============================================================
   MOBILE
   ============================================================ */

@media (max-width: 760px) {

    body {
        padding:
            14px;
    }


    .container {
        padding:
            17px;
    }


    .banner {
        min-height:
            280px;

        padding:
            30px 24px;

        background-position:
            center;

        background-image:

            linear-gradient(
                rgba(5, 19, 38, 0.75),
                rgba(5, 19, 38, 0.75)
            ),

            url("/banner/banner.jpeg");
    }


    .picasso-title {
        font-size:
            40px;

        letter-spacing:
            4px;
    }


    .picasso-acronym {
        font-size:
            15px;
    }


    .acro-letter {
        font-size:
            19px;
    }


    .row {
        flex-direction:
            column;

        gap:
            0;
    }

}

</style>

</head>


<body>


<div class="container">


    <!-- ======================================================
         PICASSO BANNER
         ====================================================== -->

    <div class="banner">

        <div class="banner-content">


            <div class="original-title">

                INIA P Report Tool 😊

            </div>


            <h1 class="picasso-title">

                PICASSO

            </h1>


            <div class="picasso-acronym">

                <span class="acro-letter">P</span>
                INIA

                <span class="acro-letter">I</span>nspection
                &amp;

                <span class="acro-letter">C</span>olorimetric

                <span class="acro-letter">A</span>nalysis

                <span class="acro-letter">S</span>ystem
                for

                <span class="acro-letter">S</span>ummary

                <span class="acro-letter">O</span>utput

            </div>


            <div class="banner-subtitle">

                Automated phosphorus analysis and pH report workflow

            </div>


        </div>

    </div>


    <!-- ======================================================
         INFORMATION
         ====================================================== -->

    <div class="note">

        This main app runs the workflow from the root folder:

        <br>

        <b>
            input.dat spectrometer CSV list
            → colorimetric XLSM
            → pH database
            → final XLSM with pH
        </b>

        <br><br>

        Temporary reports inside
        <b>output/</b>
        are cleaned before a new run.

        <br>

        Existing reports inside
        <b>final_report/</b>
        are preserved.

        <br><br>

        Password-protected pH Excel files will use password
        <b>12</b>
        by default.

    </div>


    <form method="post">


        <!-- ==================================================
             1. SPECTROMETER CSV
             ================================================== -->

        <h2>
            1. Spectrometer CSV input
        </h2>


        <div class="path-box">

            Existing input DAT file:

            <br>

            {{ input_dat_path }}

        </div>


        <div class="small">

            This app uses the existing
            <b>input.dat</b>.

            That file must contain the spectrometer CSV
            filenames, one per line.

            The CSV files must be inside
            <b>input/</b>.

        </div>


        <label>
            Colorimetric method
        </label>


        <select name="method">

            <option
                value="Bray"
                {% if values.method == "Bray" %}selected{% endif %}
            >
                Bray
            </option>


            <option
                value="Olsen"
                {% if values.method == "Olsen" %}selected{% endif %}
            >
                Olsen
            </option>


            <option
                value="Both"
                {% if values.method == "Both" %}selected{% endif %}
            >
                Both Bray and Olsen
            </option>

        </select>


        <!-- ==================================================
             2. PH DATABASE
             ================================================== -->

        <h2>
            2. pH database settings
        </h2>


        <label>
            Folder with pH Excel files
        </label>


        <input
            type="text"
            name="ph_folder"
            value="{{ values.ph_folder }}"
            placeholder="G:\\Mi unidad\\LABSAF ILLPA\\1. Documentos Internos\\7.5 Registros Tecnicos\\2026\\SUELOS\\1.pH"
        >


        <div class="row">


            <div class="col">

                <label>
                    pH database output
                </label>

                <div class="path-box">

                    {{ ph_database_file }}

                </div>

            </div>


            <div class="col">

                <label>
                    pH CSV output
                </label>

                <div class="path-box">

                    {{ ph_csv_file }}

                </div>

            </div>


        </div>


        <div class="row">


            <div class="col">

                <label>
                    pH filename filter
                </label>

                <input
                    type="text"
                    name="ph_file_filter"
                    value="{{ values.ph_file_filter }}"
                >

            </div>


            <div class="col">

                <label>
                    pH sheet name
                </label>

                <input
                    type="text"
                    name="ph_sheet_name"
                    value="{{ values.ph_sheet_name }}"
                >

            </div>


        </div>


        <div class="row">


            <div class="col">

                <label>
                    Excel password for pH files
                </label>

                <input
                    type="text"
                    name="excel_password"
                    value="{{ values.excel_password }}"
                >

                <div class="small">

                    Default password: 12

                </div>

            </div>


        </div>


        <!-- ==================================================
             3. FILL PH
             ================================================== -->

        <h2>
            3. Fill pH into generated reports
        </h2>


        <div class="small">

            The app will automatically create
            <b>input_for_ph.dat</b>
            from the generated
            <b>Analizado_*.xlsm</b>
            files in
            <b>output/bray</b>
            or
            <b>output/olsen</b>.

            Final files will be written to
            <b>final_report/bray</b>
            or
            <b>final_report/olsen</b>.

            Existing final reports are not removed during
            temporary-output cleanup.

        </div>


        <div class="row">


            <div class="col">

                <label>
                    Generated report sheet name
                </label>

                <input
                    type="text"
                    name="fill_sheet_name"
                    value="{{ values.fill_sheet_name }}"
                >

            </div>


            <div class="col">

                <label>
                    Code column
                </label>

                <input
                    type="text"
                    name="fill_code_col"
                    value="{{ values.fill_code_col }}"
                >

            </div>


            <div class="col">

                <label>
                    pH output column
                </label>

                <input
                    type="text"
                    name="fill_output_col"
                    value="{{ values.fill_output_col }}"
                >

            </div>


        </div>


        <div class="row">


            <div class="col">

                <label>
                    First row
                </label>

                <input
                    type="number"
                    name="fill_first_row"
                    value="{{ values.fill_first_row }}"
                >

            </div>


            <div class="col">

                <label>
                    Block size
                </label>

                <input
                    type="number"
                    name="fill_block_size"
                    value="{{ values.fill_block_size }}"
                >

            </div>


            <div class="col">

                <label>
                    Gap rows
                </label>

                <input
                    type="number"
                    name="fill_gap_rows"
                    value="{{ values.fill_gap_rows }}"
                >

            </div>


        </div>


        <!-- ==================================================
             ACTIONS
             ================================================== -->

        <h2>
            Actions
        </h2>


        <button
            name="action"
            value="generate_reports"
        >
            1. Generate colorimetric reports
        </button>


        <button
            name="action"
            value="update_ph_db"
            class="danger-button"
        >
            2. Update pH database
        </button>


        <button
            name="action"
            value="fill_ph"
            class="green-button"
        >
            3. Add pH to generated reports
        </button>


        <button
            name="action"
            value="full_pipeline"
            class="purple-button"
        >
            Run full pipeline
        </button>


    </form>


    <!-- ======================================================
         STATUS
         ====================================================== -->

    {% if status %}

        <hr>


        {% if success %}

            <p class="success">

                {{ status }}

            </p>

        {% else %}

            <p class="error">

                {{ status }}

            </p>

        {% endif %}


    {% endif %}


    <!-- ======================================================
         TERMINAL OUTPUT
         ====================================================== -->

    {% if output %}

        <h2>
            Terminal output
        </h2>

        <pre>{{ output }}</pre>

    {% endif %}


    <!-- ======================================================
         FINAL FILES
         ====================================================== -->

    {% if files %}

        <div class="files">

            <h2>
                Final reports with pH
            </h2>


            {% for item in files %}

                <a
                    href="/download/{{ item.method }}/{{ item.filename }}"
                    target="_blank"
                >
                    final_report/{{ item.method }}/{{ item.filename }}
                </a>

            {% endfor %}


        </div>

    {% endif %}


</div>


</body>

</html>
"""


# ============================================================
# HELPERS
# ============================================================


def run_command(
    command,
    cwd,
):

    env = dict(
        os.environ
    )

    env[
        "PYTHONIOENCODING"
    ] = "utf-8"

    env[
        "PYTHONUTF8"
    ] = "1"

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

        output += (
            result.stdout
        )

    if result.stderr:

        output += (
            "\n\nSTDERR:\n"
        )

        output += (
            result.stderr
        )

    success = (
        result.returncode
        == 0
    )

    return (
        success,
        output,
    )


def get_methods(
    method,
):

    if method == "Both":

        return [
            "Bray",
            "Olsen",
        ]

    return [
        method
    ]


# ============================================================
# CLEAN TEMPORARY OUTPUTS
# ============================================================


def clean_previous_outputs(
    method,
):

    """
    Remove ONLY temporary/intermediate files from output/<method>.

    Removed:

        output/<method>/Analizado*.xlsm
        output/<method>/input_for_ph.dat

    NEVER touched:

        final_report/bray/
        final_report/olsen/

    final_report contains completed reports with pH and may
    contain manual analyst notes.

    Therefore this function must NEVER delete, clean, or modify
    anything inside final_report/.
    """

    method_lower = (
        method.lower()
    )

    method_output_dir = (
        OUTPUT_DIR
        / method_lower
    )

    method_output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    deleted_output_files = []

    # --------------------------------------------------------
    # Delete old temporary colorimetric reports WITHOUT pH.
    #
    # ONLY:
    #
    #   output/bray/Analizado*.xlsm
    #
    # or:
    #
    #   output/olsen/Analizado*.xlsm
    #
    # --------------------------------------------------------

    for path in (
        method_output_dir.glob(
            "Analizado*.xlsm"
        )
    ):

        if path.is_file():

            path.unlink()

            deleted_output_files.append(
                path.name
            )

    # --------------------------------------------------------
    # Delete old temporary DAT list.
    # --------------------------------------------------------

    input_for_ph = (
        method_output_dir
        / "input_for_ph.dat"
    )

    if (
        input_for_ph.exists()
    ):

        input_for_ph.unlink()

    # --------------------------------------------------------
    # ABSOLUTELY NO final_report cleanup here.
    # --------------------------------------------------------

    msg = ""

    msg += (
        f"Cleaning previous temporary "
        f"{method} output...\n"
    )

    msg += (
        "Old intermediate reports deleted: "
        f"{len(deleted_output_files)}\n"
    )

    msg += (
        "Old input_for_ph.dat removed "
        "if it existed.\n"
    )

    msg += (
        "final_report was NOT cleaned "
        "or deleted.\n"
    )

    return msg


# ============================================================
# INPUT.DAT CHECK
# ============================================================


def check_input_dat():

    output = ""

    if not INPUT_DAT.exists():

        return (
            False,
            (
                "input.dat was not found:\n"
                f"  {INPUT_DAT}\n"
            ),
        )

    if not INPUT_DIR.exists():

        return (
            False,
            (
                "input folder was not found:\n"
                f"  {INPUT_DIR}\n"
            ),
        )

    with open(
        INPUT_DAT,
        "r",
        encoding="utf-8-sig",
    ) as f:

        lines = [

            line.strip()

            for line in f

            if (
                line.strip()
                and not
                line.strip().startswith(
                    "#"
                )
            )
        ]

    if not lines:

        return (
            False,
            (
                "input.dat is empty:\n"
                f"  {INPUT_DAT}\n"
            ),
        )

    output += (
        "input.dat found:\n"
    )

    output += (
        f"  {INPUT_DAT}\n"
    )

    output += (
        f"CSV names listed: "
        f"{len(lines)}\n"
    )

    missing = []

    for csv_name in lines:

        csv_path = (
            INPUT_DIR
            / csv_name
        )

        if not csv_path.exists():

            missing.append(
                csv_name
            )

    if missing:

        output += (
            "\nMissing CSV files "
            "inside input/:\n"
        )

        for name in missing:

            output += (
                f"  - {name}\n"
            )

        return (
            False,
            output,
        )

    output += (
        "\nAll CSV files listed in input.dat "
        "were found inside input/.\n"
    )

    return (
        True,
        output,
    )


# ============================================================
# GENERATE COLORIMETRIC REPORTS
# ============================================================


def generate_colorimetric_reports(
    method,
):

    output = ""

    ok, msg = (
        check_input_dat()
    )

    output += msg

    if not ok:

        return (
            False,
            output,
        )

    if not UPDATE_COLOR_SCRIPT.exists():

        return (
            False,
            (
                "Script not found:\n"
                f"  {UPDATE_COLOR_SCRIPT}\n"
            ),
        )

    all_success = True

    for one_method in (
        get_methods(
            method
        )
    ):

        output += "\n"

        output += (
            "============================================================\n"
        )

        output += (
            f"PREPARING NEW RUN: "
            f"{one_method}\n"
        )

        output += (
            "============================================================\n"
        )

        # ----------------------------------------------------
        # Remove ONLY temporary output from previous run.
        #
        # final_report is NEVER cleaned.
        # ----------------------------------------------------

        try:

            clean_msg = (
                clean_previous_outputs(
                    one_method
                )
            )

            output += (
                clean_msg
            )

        except Exception:

            all_success = False

            output += (
                "\nERROR cleaning previous "
                "temporary output files:\n"
            )

            output += (
                traceback.format_exc()
            )

            # Do not generate into a directory that may
            # still contain stale temporary files.
            continue

        output += "\n"

        output += (
            "============================================================\n"
        )

        output += (
            f"GENERATING COLORIMETRIC REPORTS: "
            f"{one_method}\n"
        )

        output += (
            "============================================================\n"
        )

        command = [

            sys.executable,

            str(
                UPDATE_COLOR_SCRIPT
            ),

            "--method",

            one_method,
        ]

        (
            success,
            cmd_output,
        ) = run_command(

            command,

            cwd=BASE_DIR,
        )

        output += (
            cmd_output
        )

        if not success:

            all_success = False

            output += (
                f"\nFAILED generating "
                f"{one_method} reports.\n"
            )

            # Do not create input_for_ph.dat after
            # failed generation.
            continue

        output += (
            f"\nOK: "
            f"{one_method} reports generated.\n"
        )

        try:

            dat_msg = (
                create_input_for_ph_dat(
                    one_method
                )
            )

            output += "\n"

            output += (
                dat_msg
            )

        except Exception:

            all_success = False

            output += (
                "\nERROR creating "
                "input_for_ph.dat:\n"
            )

            output += (
                traceback.format_exc()
            )

    return (
        all_success,
        output,
    )


# ============================================================
# UPDATE PH DATABASE
# ============================================================


def update_ph_database(
    values,
):

    if not BUILD_PH_SCRIPT.exists():

        return (
            False,
            (
                "Script not found:\n"
                f"  {BUILD_PH_SCRIPT}\n"
            ),
        )

    ph_folder = (
        values[
            "ph_folder"
        ].strip()
    )

    if not ph_folder:

        return (
            False,
            (
                "pH folder is empty. "
                "Please provide the folder "
                "containing pH Excel files.\n"
            ),
        )

    command = [

        sys.executable,

        str(
            BUILD_PH_SCRIPT
        ),

        "--ph-folder",
        ph_folder,

        "--database-file",
        str(
            PH_DATABASE_FILE
        ),

        "--csv-file",
        str(
            PH_CSV_FILE
        ),

        "--file-name-filter",
        values[
            "ph_file_filter"
        ],

        "--sheet-name",
        values[
            "ph_sheet_name"
        ],

        "--excel-password",
        values[
            "excel_password"
        ],

        "--quiet",
    ]

    (
        success,
        output,
    ) = run_command(

        command,

        cwd=GET_DB_DIR,
    )

    return (
        success,
        output,
    )


# ============================================================
# CREATE INPUT_FOR_PH.DAT
# ============================================================


def create_input_for_ph_dat(
    method,
):

    """
    Create:

        output/bray/input_for_ph.dat
        output/olsen/input_for_ph.dat

    using generated:

        output/bray/Analizado_*.xlsm
        output/olsen/Analizado_*.xlsm
    """

    method_lower = (
        method.lower()
    )

    method_output_dir = (
        OUTPUT_DIR
        / method_lower
    )

    if not method_output_dir.exists():

        raise FileNotFoundError(
            "Generated output folder "
            "does not exist: "
            f"{method_output_dir}"
        )

    xlsm_files = sorted(
        [

            path

            for path in (
                method_output_dir.glob(
                    "Analizado*.xlsm"
                )
            )

            if path.is_file()
        ]
    )

    if not xlsm_files:

        raise FileNotFoundError(
            "No Analizado*.xlsm files "
            "found in: "
            f"{method_output_dir}"
        )

    dat_file = (
        method_output_dir
        / "input_for_ph.dat"
    )

    with open(
        dat_file,
        "w",
        encoding="utf-8",
    ) as f:

        for xlsm_file in (
            xlsm_files
        ):

            f.write(
                xlsm_file.name
                + "\n"
            )

    msg = ""

    msg += (
        "Created input_for_ph.dat:\n"
    )

    msg += (
        f"  {dat_file}\n"
    )

    msg += (
        f"Files listed: "
        f"{len(xlsm_files)}\n"
    )

    for xlsm_file in (
        xlsm_files
    ):

        msg += (
            f"  - {xlsm_file.name}\n"
        )

    return msg


# ============================================================
# FILL PH INTO GENERATED REPORTS
# ============================================================


def fill_ph_into_generated_reports(
    method,
    values,
):

    if not FILL_PH_SCRIPT.exists():

        return (
            False,
            (
                "Script not found:\n"
                f"  {FILL_PH_SCRIPT}\n"
            ),
        )

    if not PH_CSV_FILE.exists():

        return (
            False,
            (
                "pH CSV file not found:\n"
                f"  {PH_CSV_FILE}\n"
            ),
        )

    all_success = True

    output = ""

    for one_method in (
        get_methods(
            method
        )
    ):

        method_lower = (
            one_method.lower()
        )

        output += "\n"

        output += (
            "============================================================\n"
        )

        output += (
            f"ADDING pH TO GENERATED REPORTS: "
            f"{one_method}\n"
        )

        output += (
            "============================================================\n"
        )

        try:

            dat_msg = (
                create_input_for_ph_dat(
                    one_method
                )
            )

            output += (
                dat_msg
            )

            generated_dir = (
                OUTPUT_DIR
                / method_lower
            )

            dat_file = (
                generated_dir
                / "input_for_ph.dat"
            )

            # ------------------------------------------------
            # Final reports go here.
            #
            # IMPORTANT:
            # This function does NOT clean this directory.
            # Existing files are left in place.
            # ------------------------------------------------

            final_method_dir = (
                FINAL_REPORT_DIR
                / method_lower
            )

            final_method_dir.mkdir(
                parents=True,
                exist_ok=True,
            )

            output += "\n"

            output += (
                "Input generated reports folder:\n"
            )

            output += (
                f"  {generated_dir}\n"
            )

            output += (
                "Final report folder:\n"
            )

            output += (
                f"  {final_method_dir}\n"
            )

            output += (
                "Existing final reports are "
                "not cleaned.\n"
            )

            command = [

                sys.executable,

                str(
                    FILL_PH_SCRIPT
                ),

                "--input-dat",
                str(
                    dat_file
                ),

                "--input-dir",
                str(
                    generated_dir
                ),

                "--ph-csv",
                str(
                    PH_CSV_FILE
                ),

                "--output-dir",
                str(
                    final_method_dir
                ),

                "--sheet-name",
                values[
                    "fill_sheet_name"
                ],

                "--first-row",
                str(
                    values[
                        "fill_first_row"
                    ]
                ),

                "--block-size",
                str(
                    values[
                        "fill_block_size"
                    ]
                ),

                "--gap-rows",
                str(
                    values[
                        "fill_gap_rows"
                    ]
                ),

                "--code-col",
                values[
                    "fill_code_col"
                ],

                "--output-col",
                values[
                    "fill_output_col"
                ],

                "--quiet",
            ]

            (
                success,
                cmd_output,
            ) = run_command(

                command,

                cwd=FILL_PH_DIR,
            )

            output += (
                cmd_output
            )

            if not success:

                all_success = False

                output += (
                    f"\nFAILED adding pH "
                    f"to {one_method} reports.\n"
                )

            else:

                output += (
                    f"\nOK: pH added to "
                    f"{one_method} reports.\n"
                )

        except Exception:

            all_success = False

            output += (
                "\nERROR while adding pH:\n"
            )

            output += (
                traceback.format_exc()
            )

    return (
        all_success,
        output,
    )


# ============================================================
# LIST FINAL REPORT FILES
# ============================================================


def list_final_report_files():

    """
    Read-only listing of final reports.

    Nothing is deleted or modified here.
    """

    results = []

    for method_lower in [
        "bray",
        "olsen",
    ]:

        folder = (
            FINAL_REPORT_DIR
            / method_lower
        )

        if not folder.exists():

            continue

        for path in sorted(
            folder.glob(
                "*.xlsm"
            )
        ):

            if path.is_file():

                results.append(
                    {
                        "method":
                            method_lower,

                        "filename":
                            path.name,
                    }
                )

    return results


# ============================================================
# OPEN BROWSER
# ============================================================


def open_browser():

    webbrowser.open_new(
        URL
    )


# ============================================================
# DEFAULT VALUES
# ============================================================


def get_default_values():

    return {

        "method":
            "Bray",

        "ph_folder":
            (
                r"G:\Mi unidad\LABSAF ILLPA"
                r"\1. Documentos Internos"
                r"\7.5 Registros Tecnicos"
                r"\2026\SUELOS\1.pH"
            ),

        "ph_file_filter":
            "Ver.03",

        "ph_sheet_name":
            "F-103",

        "excel_password":
            "12",

        "fill_sheet_name":
            "P_DIS",

        "fill_code_col":
            "C",

        "fill_output_col":
            "E",

        "fill_first_row":
            "37",

        "fill_block_size":
            "20",

        "fill_gap_rows":
            "3",
    }


def get_values_from_form():

    defaults = (
        get_default_values()
    )

    values = {}

    for (
        key,
        default,
    ) in defaults.items():

        values[
            key
        ] = request.form.get(
            key,
            default,
        ).strip()

    return values


# ============================================================
# ROUTES
# ============================================================


# ------------------------------------------------------------
# Serve banner/banner.jpeg
# ------------------------------------------------------------

@app.route(
    "/banner/<path:filename>",
    methods=[
        "GET"
    ],
)
def banner_file(
    filename,
):

    return send_from_directory(
        BANNER_DIR,
        filename,
    )


# ------------------------------------------------------------
# Main page
# ------------------------------------------------------------

@app.route(
    "/",
    methods=[
        "GET",
        "POST",
    ],
)
def index():

    status = None

    success = None

    output = ""

    files = (
        list_final_report_files()
    )

    if request.method == "GET":

        values = (
            get_default_values()
        )

        return render_template_string(

            HTML_PAGE,

            values=
                values,

            status=
                status,

            success=
                success,

            output=
                output,

            files=
                files,

            input_dat_path=
                str(
                    INPUT_DAT
                ),

            ph_database_file=
                str(
                    PH_DATABASE_FILE
                ),

            ph_csv_file=
                str(
                    PH_CSV_FILE
                ),
        )

    values = (
        get_values_from_form()
    )

    action = request.form.get(
        "action",
        "",
    ).strip()

    try:

        # ----------------------------------------------------
        # GENERATE REPORTS
        # ----------------------------------------------------

        if (
            action
            == "generate_reports"
        ):

            (
                success,
                output,
            ) = generate_colorimetric_reports(

                method=
                    values[
                        "method"
                    ]
            )

            if success:

                status = (
                    "Colorimetric reports "
                    "generated successfully."
                )

            else:

                status = (
                    "Colorimetric report "
                    "generation failed."
                )


        # ----------------------------------------------------
        # UPDATE PH DATABASE
        # ----------------------------------------------------

        elif (
            action
            == "update_ph_db"
        ):

            (
                success,
                output,
            ) = update_ph_database(
                values
            )

            if success:

                status = (
                    "pH database updated "
                    "successfully."
                )

            else:

                status = (
                    "pH database update failed."
                )


        # ----------------------------------------------------
        # FILL PH
        # ----------------------------------------------------

        elif (
            action
            == "fill_ph"
        ):

            (
                success,
                output,
            ) = fill_ph_into_generated_reports(

                method=
                    values[
                        "method"
                    ],

                values=
                    values,
            )

            if success:

                status = (
                    "pH added to generated "
                    "reports successfully."
                )

            else:

                status = (
                    "Adding pH to generated "
                    "reports failed."
                )


        # ----------------------------------------------------
        # FULL PIPELINE
        # ----------------------------------------------------

        elif (
            action
            == "full_pipeline"
        ):

            full_output = ""

            full_success = True


            # =================================================
            # STEP 1
            # =================================================

            (
                s1,
                out1,
            ) = generate_colorimetric_reports(

                method=
                    values[
                        "method"
                    ]
            )

            full_output += (
                "\n\n"
                "================ "
                "STEP 1: GENERATE REPORTS "
                "================\n"
            )

            full_output += (
                out1
            )

            if not s1:

                full_success = False


            # =================================================
            # STEP 2
            # =================================================

            (
                s2,
                out2,
            ) = update_ph_database(
                values
            )

            full_output += (
                "\n\n"
                "================ "
                "STEP 2: UPDATE pH DATABASE "
                "================\n"
            )

            full_output += (
                out2
            )

            if not s2:

                full_success = False


            # =================================================
            # STEP 3
            # =================================================

            (
                s3,
                out3,
            ) = fill_ph_into_generated_reports(

                method=
                    values[
                        "method"
                    ],

                values=
                    values,
            )

            full_output += (
                "\n\n"
                "================ "
                "STEP 3: ADD pH TO REPORTS "
                "================\n"
            )

            full_output += (
                out3
            )

            if not s3:

                full_success = False


            success = (
                full_success
            )

            output = (
                full_output
            )

            if success:

                status = (
                    "Full pipeline finished "
                    "successfully."
                )

            else:

                status = (
                    "Full pipeline finished "
                    "with errors. "
                    "Check the output."
                )


        # ----------------------------------------------------
        # INVALID ACTION
        # ----------------------------------------------------

        else:

            success = False

            status = (
                "Invalid action."
            )

            output = (
                f"Unknown action: "
                f"{action}"
            )


    except Exception:

        success = False

        status = (
            "Unexpected error."
        )

        output = (
            traceback.format_exc()
        )


    files = (
        list_final_report_files()
    )


    return render_template_string(

        HTML_PAGE,

        values=
            values,

        status=
            status,

        success=
            success,

        output=
            output,

        files=
            files,

        input_dat_path=
            str(
                INPUT_DAT
            ),

        ph_database_file=
            str(
                PH_DATABASE_FILE
            ),

        ph_csv_file=
            str(
                PH_CSV_FILE
            ),
    )


# ------------------------------------------------------------
# Download final reports
# ------------------------------------------------------------

@app.route(
    "/download/<method>/<filename>",
    methods=[
        "GET"
    ],
)
def download_file(
    method,
    filename,
):

    method = (
        method.lower()
    )

    if method not in [
        "bray",
        "olsen",
    ]:

        abort(
            404
        )

    folder = (
        FINAL_REPORT_DIR
        / method
    )

    file_path = (
        folder
        / filename
    )

    if not file_path.exists():

        abort(
            404
        )

    return send_from_directory(

        folder,

        filename,

        as_attachment=True,
    )


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    INPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # Creating the directory if missing is harmless.
    #
    # No existing files are deleted here.
    # --------------------------------------------------------

    FINAL_REPORT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    Timer(
        1.0,
        open_browser,
    ).start()

    app.run(
        host=HOST,
        port=PORT,
        debug=False,
    )
