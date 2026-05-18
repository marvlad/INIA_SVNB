# build_ph_database_from_excel_files_Ver03_only.py

from pathlib import Path
import re
import csv
import sqlite3
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

PH_FOLDER = r"G:\Mi unidad\LABSAF ILLPA\1. Documentos Internos\7.5 Registros Tecnicos\2026\SUELOS\1.pH"

DATABASE_FILE = r"G:\Mi unidad\LABSAF ILLPA\ph_database_Ver03.sqlite"

CSV_FILE = r"G:\Mi unidad\LABSAF ILLPA\ph_database_Ver03.csv"

FILE_NAME_FILTER = "Ver.03"

PH_SHEET_NAME = "F-103"

PH_FIRST_ROW = 27

# Blocks:
#   27 to 47
#   50 to 70
#   73 to 93
#   ...
BLOCK_SIZE = 21
GAP_ROWS = 2

PH_DUPLICATE_COL = "B"
PH_CODE_COL = "C"
PH_VALUE_COL = "H"

VERBOSE = True


# ============================================================
# HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    # Ignore Excel apostrophe:
    # 'SU1149-ILL-26 -> SU1149-ILL-26
    text = text.lstrip("'").strip()

    return text


def normalize_code(value):
    return normalize_text(value).upper()


def normalize_duplicate(value):
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
    text = normalize_text(value)

    if text == "":
        return ""

    text = text.replace(",", ".")
    return text


def extract_su_number(value):
    text = normalize_code(value)

    match = re.search(r"SU\s*0*(\d+)", text)

    if not match:
        return None

    return int(match.group(1))


def get_ver03_excel_files(ph_folder):
    all_excel_files = sorted(
        list(ph_folder.glob("*.xlsx")) +
        list(ph_folder.glob("*.xlsm"))
    )

    filtered_files = [
        path for path in all_excel_files
        if FILE_NAME_FILTER.upper() in path.name.upper()
    ]

    print("")
    print("File filter:")
    print(f"  Only reading files containing: {FILE_NAME_FILTER}")
    print(f"  Total Excel files found: {len(all_excel_files)}")
    print(f"  Ver.03 files selected:   {len(filtered_files)}")

    print("")
    print("Selected files:")
    for path in filtered_files:
        print(f"  - {path.name}")

    return filtered_files


# ============================================================
# DATABASE
# ============================================================

def create_database(db_file):
    conn = sqlite3.connect(db_file)
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

    cur.execute("CREATE INDEX idx_ph_code ON ph_data (code)")
    cur.execute("CREATE INDEX idx_ph_su_number ON ph_data (su_number)")
    cur.execute("CREATE INDEX idx_ph_duplicate ON ph_data (duplicate)")

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
    print("BUILDING PH DATABASE FROM Ver.03 EXCEL FILES ONLY")
    print("============================================================")
    print("pH folder:")
    print(f"  {ph_folder}")
    print("SQLite database:")
    print(f"  {db_path}")
    print("CSV output:")
    print(f"  {csv_path}")
    print("pH sheet:")
    print(f"  {PH_SHEET_NAME}")
    print("Rows:")
    print("  27:47, 50:70, 73:93, ...")
    print("")
    print("Stop rule:")
    print("  If both column C code and column H pH are empty, stop that file.")
    print("============================================================")

    if not ph_folder.exists():
        raise FileNotFoundError(f"pH folder not found: {ph_folder}")

    excel_files = get_ver03_excel_files(ph_folder)

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
    total_files_stopped = 0

    for file_index, excel_file in enumerate(excel_files, start=1):
        print("")
        print("============================================================")
        print(f"[{file_index}/{len(excel_files)}] ACCESSING Ver.03 EXCEL FILE")
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
                wb.close()
                continue

            ws = wb[PH_SHEET_NAME]

            print("")
            print(f"Accessing sheet: {PH_SHEET_NAME}")
            print(f"Excel reported max row: {ws.max_row}")
            print("Important: max_row may be large because of formatting.")
            print("The script will stop when C and H are both empty.")

            file_inserted = 0
            file_skipped = 0
            stopped_by_empty_code_and_ph = False

            # ------------------------------------------------------------
            # SAFE BLOCK LOOP
            # ------------------------------------------------------------
            # Instead of looping until ws.max_row forever, we generate blocks:
            #   27:47
            #   50:70
            #   73:93
            # and stop immediately when C and H are both empty.
            # ------------------------------------------------------------

            block_start = PH_FIRST_ROW

            while True:
                block_end = block_start + BLOCK_SIZE - 1

                print("")
                print("------------------------------------------------------------")
                print(f"Reading pH block: rows {block_start} to {block_end}")
                print("------------------------------------------------------------")

                stop_this_file = False

                for row in range(block_start, block_end + 1):
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

                    # ----------------------------------------------------
                    # IMPORTANT STOP RULE
                    # ----------------------------------------------------
                    # If there is no code in C and no pH in H,
                    # the useful data ended in this file.
                    # ----------------------------------------------------
                    if code == "" and ph == "":
                        print("")
                        print("  STOP FILE:")
                        print(f"    {code_cell} is empty.")
                        print(f"    {ph_cell} is empty.")
                        print("    Column C and H are both empty.")
                        print("    Moving to the next Excel file.")
                        stop_this_file = True
                        stopped_by_empty_code_and_ph = True
                        total_files_stopped += 1
                        break

                    # If C has no SU but H has something, skip the row.
                    if code == "":
                        print("  SKIPPED ROW: empty code in column C, but pH column is not empty.")
                        file_skipped += 1
                        continue

                    if su_number is None:
                        print("  SKIPPED ROW: code exists, but no SU number found.")
                        file_skipped += 1
                        continue

                    if ph == "":
                        print("  SKIPPED ROW: empty pH in column H.")
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

                if stop_this_file:
                    break

                # Move to next block:
                # 27:47 -> next 50
                # 50:70 -> next 73
                block_start = block_end + GAP_ROWS + 1

                # Extra safety guard:
                # If block_start goes far beyond Excel's reported max row,
                # stop even if formatting caused weird behavior.
                if block_start > ws.max_row + BLOCK_SIZE + GAP_ROWS:
                    print("")
                    print("  SAFETY STOP:")
                    print("    Reached beyond Excel's reported max row.")
                    print("    Moving to next file.")
                    break

            conn.commit()
            wb.close()

            total_inserted += file_inserted
            total_skipped += file_skipped

            print("")
            print("------------------------------------------------------------")
            print(f"Finished file: {excel_file.name}")
            print(f"Inserted rows: {file_inserted}")
            print(f"Skipped rows:  {file_skipped}")

            if stopped_by_empty_code_and_ph:
                print("Stop reason: first row where C and H are both empty")
            else:
                print("Stop reason: safety stop or reached file end")

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
    print(f"Files stopped by empty C and H: {total_files_stopped}")
    print("")
    print("Created:")
    print(f"  {db_path}")
    print(f"  {csv_path}")


if __name__ == "__main__":
    build_ph_database()
