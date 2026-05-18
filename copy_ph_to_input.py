# build_ph_database_only.py

from pathlib import Path
import re
import csv
import sqlite3
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

PH_FOLDER = r"G:\Mi unidad\LABSAF ILLPA\1. Documentos Internos\7.5 Registros Tecnicos\2026\SUELOS\1.pH"

DATABASE_FILE = r"G:\Mi unidad\LABSAF ILLPA\ph_database.sqlite"

CSV_FILE = r"G:\Mi unidad\LABSAF ILLPA\ph_database.csv"

PH_SHEET_NAME = "F-103"

PH_FIRST_ROW = 27

PH_MARKER_COL = "B"   # D, D2, 1, 2, etc.
PH_CODE_COL = "C"     # SU code
PH_VALUE_COL = "H"    # pH value

VERBOSE = True


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def vprint(message):
    if VERBOSE:
        print(message)


def normalize_text(value):
    """
    Clean Excel text.

    It removes:
    - spaces
    - non-breaking spaces
    - leading Excel apostrophe, for example 'SU1149-ILL-26
    """
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    # Remove leading Excel apostrophe
    text = text.lstrip("'").strip()

    return text


def normalize_code(value):
    """
    Normalize SU code.
    """
    return normalize_text(value).upper()


def normalize_marker(value):
    """
    Normalize the marker from column B.

    Important:
    This keeps values like:
        D
        D2
        1
        2

    It only converts numeric 1.0 to 1.
    """
    text = normalize_text(value).upper()

    if text == "":
        return ""

    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except ValueError:
        pass

    return text


def extract_su_number(value):
    """
    Extract only the numeric SU number.

    Examples:
        SU1149-ILL-26   -> 1149
        'SU1149-ILL-26  -> 1149
        SU0079          -> 79
    """
    text = normalize_code(value)

    match = re.search(r"SU\s*0*(\d+)", text)

    if not match:
        return None

    return int(match.group(1))


def get_sheet(workbook, sheet_name):
    """
    Get sheet by name.
    """
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {workbook.sheetnames}"
        )

    return workbook[sheet_name]


# ============================================================
# SQLITE FUNCTIONS
# ============================================================

def create_database(db_path):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ph_data")

    cur.execute(
        """
        CREATE TABLE ph_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,

            su_code TEXT,
            ph_value TEXT,
            marker TEXT,

            su_number INTEGER,

            source_file TEXT,
            source_sheet TEXT,
            source_row INTEGER
        )
        """
    )

    cur.execute(
        """
        CREATE INDEX idx_ph_su_code
        ON ph_data (su_code)
        """
    )

    cur.execute(
        """
        CREATE INDEX idx_ph_su_number_marker
        ON ph_data (su_number, marker)
        """
    )

    cur.execute(
        """
        CREATE INDEX idx_ph_marker
        ON ph_data (marker)
        """
    )

    conn.commit()

    return conn


def insert_record(
    conn,
    su_code,
    ph_value,
    marker,
    su_number,
    source_file,
    source_sheet,
    source_row,
):
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO ph_data (
            su_code,
            ph_value,
            marker,
            su_number,
            source_file,
            source_sheet,
            source_row
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            su_code,
            str(ph_value),
            marker,
            su_number,
            source_file,
            source_sheet,
            source_row,
        ),
    )


# ============================================================
# MAIN DATABASE BUILDER
# ============================================================

def build_ph_database():
    ph_folder = Path(PH_FOLDER)
    db_path = Path(DATABASE_FILE)
    csv_path = Path(CSV_FILE)

    print("")
    print("============================================================")
    print("BUILDING pH DATABASE FROM ALL EXCEL FILES")
    print("============================================================")
    print(f"pH folder:")
    print(f"  {ph_folder}")
    print(f"SQLite database:")
    print(f"  {db_path}")
    print(f"CSV output:")
    print(f"  {csv_path}")
    print("============================================================")

    if not ph_folder.exists():
        raise FileNotFoundError(f"Folder not found: {ph_folder}")

    excel_files = sorted(
        list(ph_folder.glob("*.xlsx")) +
        list(ph_folder.glob("*.xlsm"))
    )

    print("")
    print(f"Excel files found: {len(excel_files)}")

    conn = create_database(db_path)

    csv_file = open(csv_path, "w", newline="", encoding="utf-8-sig")
    writer = csv.writer(csv_file)

    # You asked for the B values as the 3rd column.
    writer.writerow(
        [
            "SU_CODE",        # column 1
            "PH_VALUE",       # column 2
            "MARKER_B",       # column 3: B values, D, D2, 1, 2, etc.
            "SU_NUMBER",
            "SOURCE_FILE",
            "SOURCE_SHEET",
            "SOURCE_ROW",
        ]
    )

    total_inserted = 0
    total_skipped = 0

    for file_index, excel_file in enumerate(excel_files, start=1):
        print("")
        print("------------------------------------------------------------")
        print(f"[{file_index}/{len(excel_files)}] Opening file:")
        print(f"  {excel_file.name}")
        print("------------------------------------------------------------")

        try:
            wb = load_workbook(excel_file, data_only=True, read_only=True)

            if PH_SHEET_NAME not in wb.sheetnames:
                print(f"  SKIPPED: sheet '{PH_SHEET_NAME}' not found.")
                print(f"  Available sheets: {wb.sheetnames}")
                wb.close()
                continue

            ws = wb[PH_SHEET_NAME]

            print(f"  Sheet: {ws.title}")
            print(f"  Max row: {ws.max_row}")
            print(f"  Reading rows {PH_FIRST_ROW} to {ws.max_row}")
            print("")
            print("  Extracting:")
            print(f"    {PH_CODE_COL} = SU code")
            print(f"    {PH_VALUE_COL} = pH")
            print(f"    {PH_MARKER_COL} = marker / duplicate value")
            print("")

            file_inserted = 0
            file_skipped = 0

            for row in range(PH_FIRST_ROW, ws.max_row + 1):
                raw_marker = ws[f"{PH_MARKER_COL}{row}"].value
                raw_code = ws[f"{PH_CODE_COL}{row}"].value
                raw_ph = ws[f"{PH_VALUE_COL}{row}"].value

                marker = normalize_marker(raw_marker)
                su_code = normalize_code(raw_code)
                ph_value = normalize_text(raw_ph)
                su_number = extract_su_number(su_code)

                print(
                    f"  Row {row}: "
                    f"B={raw_marker!r} -> {marker!r}, "
                    f"C={raw_code!r} -> {su_code!r}, "
                    f"H={raw_ph!r} -> {ph_value!r}"
                )

                # Skip fully empty rows
                if marker == "" and su_code == "" and ph_value == "":
                    print("    skipped: empty row")
                    file_skipped += 1
                    continue

                # Skip rows without SU code
                if su_code == "":
                    print("    skipped: empty SU code")
                    file_skipped += 1
                    continue

                # Skip rows that do not contain SU number
                if su_number is None:
                    print("    skipped: no SU number found")
                    file_skipped += 1
                    continue

                # Skip rows without pH
                if ph_value == "":
                    print("    skipped: empty pH value")
                    file_skipped += 1
                    continue

                insert_record(
                    conn=conn,
                    su_code=su_code,
                    ph_value=ph_value,
                    marker=marker,
                    su_number=su_number,
                    source_file=excel_file.name,
                    source_sheet=ws.title,
                    source_row=row,
                )

                writer.writerow(
                    [
                        su_code,
                        ph_value,
                        marker,
                        su_number,
                        excel_file.name,
                        ws.title,
                        row,
                    ]
                )

                print(
                    f"    INSERTED: "
                    f"SU_CODE={su_code}, "
                    f"PH_VALUE={ph_value}, "
                    f"MARKER_B={marker}"
                )

                file_inserted += 1

            conn.commit()
            wb.close()

            print("")
            print(f"  File inserted rows: {file_inserted}")
            print(f"  File skipped rows:  {file_skipped}")

            total_inserted += file_inserted
            total_skipped += file_skipped

        except Exception as e:
            print("")
            print("  ERROR reading this file:")
            print(f"    {excel_file}")
            print("  Exception:")
            print(f"    {e}")

    csv_file.close()

    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ph_data")
    db_count = cur.fetchone()[0]

    conn.close()

    print("")
    print("============================================================")
    print("DATABASE BUILD FINISHED")
    print("============================================================")
    print(f"Total inserted rows: {total_inserted}")
    print(f"Total skipped rows:  {total_skipped}")
    print(f"Rows in SQLite DB:   {db_count}")
    print("")
    print("Created files:")
    print(f"  {db_path}")
    print(f"  {csv_path}")


if __name__ == "__main__":
    build_ph_database()
