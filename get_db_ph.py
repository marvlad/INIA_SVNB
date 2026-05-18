# build_ph_database_from_excel_files.py

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

# pH Excel tab
PH_SHEET_NAME = "F-103"

# pH file structure
PH_FIRST_ROW = 27

# Read:
#   27 to 47
#   50 to 70
#   73 to 93
#   96 to 116
#   ...
BLOCK_SIZE = 21
GAP_ROWS = 2

PH_DUPLICATE_COL = "B"
PH_CODE_COL = "C"
PH_VALUE_COL = "H"

VERBOSE = True


# ============================================================
# PRINT HELPER
# ============================================================

def vprint(message):
    if VERBOSE:
        print(message)


# ============================================================
# NORMALIZATION
# ============================================================

def normalize_text(value):
    """
    Clean Excel text.

    Handles:
        SU1149-ILL-26
        'SU1149-ILL-26
        spaces
        non-breaking spaces

    The apostrophe in front is ignored.
    """
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    # Ignore leading Excel apostrophe:
    #   'SU1149-ILL-26 -> SU1149-ILL-26
    text = text.lstrip("'").strip()

    return text


def normalize_code(value):
    """
    Normalize SU code from column C.
    """
    return normalize_text(value).upper()


def normalize_duplicate(value):
    """
    Normalize duplicate value from column B.

    Keeps:
        D
        D2
        1
        2
        20

    Converts:
        1.0 -> 1
        2.0 -> 2
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


def normalize_ph(value):
    """
    Normalize pH value from column H.

    Converts comma decimal to dot:
        7,5 -> 7.5
    """
    text = normalize_text(value)

    if text == "":
        return ""

    text = text.replace(",", ".")

    return text


def extract_su_number(value):
    """
    Extract numeric SU number.

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


# ============================================================
# ROW ITERATOR
# ============================================================

def iter_ph_rows(ws):
    """
    Generate rows in this pattern:

        27 to 47
        50 to 70
        73 to 93
        96 to 116
        ...

    Each block has 21 rows.
    Between blocks, 2 rows are skipped.
    """
    row = PH_FIRST_ROW

    while row <= ws.max_row:
        block_start = row
        block_end = row + BLOCK_SIZE - 1

        vprint("")
        vprint("------------------------------------------------------------")
        vprint(f"Reading pH block: rows {block_start} to {block_end}")
        vprint("------------------------------------------------------------")

        for r in range(block_start, min(block_end, ws.max_row) + 1):
            yield r

        row = block_end + GAP_ROWS + 1


# ============================================================
# DATABASE
# ============================================================

def create_database(db_file):
    db_path = Path(db_file)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS ph_data")

    cur.execute(
        """
        CREATE TABLE ph_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,

            duplicate TEXT,
            code TEXT,
            ph TEXT,

            su_number INTEGER,

            source_file TEXT,
            source_sheet TEXT,
            source_row INTEGER
        )
        """
    )

    cur.execute(
        """
        CREATE INDEX idx_ph_code
        ON ph_data (code)
        """
    )

    cur.execute(
        """
        CREATE INDEX idx_ph_su_number
        ON ph_data (su_number)
        """
    )

    cur.execute(
        """
        CREATE INDEX idx_ph_duplicate
        ON ph_data (duplicate)
        """
    )

    conn.commit()

    return conn


def insert_record(
    conn,
    duplicate,
    code,
    ph,
    su_number,
    source_file,
    source_sheet,
    source_row,
):
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO ph_data (
            duplicate,
            code,
            ph,
            su_number,
            source_file,
            source_sheet,
            source_row
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            duplicate,
            code,
            ph,
            su_number,
            source_file,
            source_sheet,
            source_row,
        ),
    )


# ============================================================
# MAIN BUILDER
# ============================================================

def build_ph_database():
    ph_folder = Path(PH_FOLDER)
    db_path = Path(DATABASE_FILE)
    csv_path = Path(CSV_FILE)

    print("")
    print("============================================================")
    print("BUILDING PH DATABASE FROM EXCEL FILES")
    print("============================================================")
    print("pH folder:")
    print(f"  {ph_folder}")
    print("SQLite database:")
    print(f"  {db_path}")
    print("CSV output:")
    print(f"  {csv_path}")
    print("pH sheet:")
    print(f"  {PH_SHEET_NAME}")
    print("Columns:")
    print(f"  {PH_DUPLICATE_COL} = duplicate")
    print(f"  {PH_CODE_COL} = code")
    print(f"  {PH_VALUE_COL} = ph")
    print("Rows:")
    print("  27:47, 50:70, 73:93, ...")
    print("")
    print("Stop rule:")
    print("  If column C has no SU/code, stop reading that file and move to next file.")
    print("============================================================")

    if not ph_folder.exists():
        raise FileNotFoundError(f"pH folder not found: {ph_folder}")

    excel_files = sorted(
        list(ph_folder.glob("*.xlsx")) +
        list(ph_folder.glob("*.xlsm"))
    )

    print("")
    print(f"Excel files found: {len(excel_files)}")

    conn = create_database(db_path)

    csv_handle = open(csv_path, "w", newline="", encoding="utf-8-sig")
    writer = csv.writer(csv_handle)

    writer.writerow(
        [
            "duplicate",
            "code",
            "ph",
            "su_number",
            "source_file",
            "source_sheet",
            "source_row",
        ]
    )

    total_inserted = 0
    total_skipped = 0
    total_files_stopped_by_empty_code = 0

    for file_index, excel_file in enumerate(excel_files, start=1):
        print("")
        print("============================================================")
        print(f"[{file_index}/{len(excel_files)}] ACCESSING EXCEL FILE")
        print("============================================================")
        print("File:")
        print(f"  {excel_file}")

        try:
            print("")
            print("Opening workbook...")
            wb = load_workbook(excel_file, data_only=True, read_only=True)

            print("Workbook opened.")
            print("Available sheets:")
            for sheet in wb.sheetnames:
                print(f"  - {sheet}")

            if PH_SHEET_NAME not in wb.sheetnames:
                print("")
                print(f"SKIPPED FILE: sheet '{PH_SHEET_NAME}' not found.")
                print(f"Available sheets: {wb.sheetnames}")
                wb.close()
                continue

            ws = wb[PH_SHEET_NAME]

            print("")
            print(f"Accessing sheet: {PH_SHEET_NAME}")
            print(f"Max row: {ws.max_row}")

            file_inserted = 0
            file_skipped = 0
            stopped_by_empty_code = False

            for row in iter_ph_rows(ws):
                duplicate_cell = f"{PH_DUPLICATE_COL}{row}"
                code_cell = f"{PH_CODE_COL}{row}"
                ph_cell = f"{PH_VALUE_COL}{row}"

                raw_duplicate = ws[duplicate_cell].value
                raw_code = ws[code_cell].value
                raw_ph = ws[ph_cell].value

                duplicate = normalize_duplicate(raw_duplicate)
                code = normalize_code(raw_code)
                ph = normalize_ph(raw_ph)
                su_number = extract_su_number(code)

                print("")
                print(f"Reading row {row}:")
                print(f"  {duplicate_cell} duplicate raw = {raw_duplicate!r} -> {duplicate!r}")
                print(f"  {code_cell} code raw      = {raw_code!r} -> {code!r}")
                print(f"  {ph_cell} ph raw          = {raw_ph!r} -> {ph!r}")
                print(f"  Extracted SU number       = {su_number}")

                # ------------------------------------------------------------
                # IMPORTANT NEW RULE
                # ------------------------------------------------------------
                # If there is no SU/code in column C, this means the file ended.
                # Stop this file and move on to the next Excel file.
                # ------------------------------------------------------------
                if code == "":
                    print("")
                    print("  STOP FILE:")
                    print(f"    No code/SU found in {code_cell}.")
                    print("    This means the data ended in this file.")
                    print("    Moving to the next pH Excel file.")
                    stopped_by_empty_code = True
                    total_files_stopped_by_empty_code += 1
                    break

                # If column C has text, but it does not contain SU, skip only this row.
                if su_number is None:
                    print("  SKIPPED ROW: code exists, but no SU number found.")
                    file_skipped += 1
                    continue

                # If pH is empty, skip only this row.
                if ph == "":
                    print("  SKIPPED ROW: empty pH.")
                    file_skipped += 1
                    continue

                insert_record(
                    conn=conn,
                    duplicate=duplicate,
                    code=code,
                    ph=ph,
                    su_number=su_number,
                    source_file=excel_file.name,
                    source_sheet=PH_SHEET_NAME,
                    source_row=row,
                )

                writer.writerow(
                    [
                        duplicate,
                        code,
                        ph,
                        su_number,
                        excel_file.name,
                        PH_SHEET_NAME,
                        row,
                    ]
                )

                print("  INSERTED INTO DATABASE AND CSV:")
                print(f"    duplicate = {duplicate}")
                print(f"    code      = {code}")
                print(f"    ph        = {ph}")
                print(f"    file      = {excel_file.name}")
                print(f"    row       = {row}")

                file_inserted += 1

            conn.commit()
            wb.close()

            total_inserted += file_inserted
            total_skipped += file_skipped

            print("")
            print("------------------------------------------------------------")
            print(f"Finished file: {excel_file.name}")
            print(f"Inserted rows: {file_inserted}")
            print(f"Skipped rows:  {file_skipped}")

            if stopped_by_empty_code:
                print("Stop reason: first empty code in column C")
            else:
                print("Stop reason: reached max row")

            print("------------------------------------------------------------")

        except Exception as e:
            print("")
            print("ERROR while reading file:")
            print(f"  {excel_file}")
            print("Exception:")
            print(f"  {e}")

    csv_handle.close()

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
    print(f"Files stopped by first empty code in column C: {total_files_stopped_by_empty_code}")
    print("")
    print("Created:")
    print(f"  {db_path}")
    print(f"  {csv_path}")


if __name__ == "__main__":
    build_ph_database()
