# build_ph_database.py

from pathlib import Path
from io import BytesIO
import re
import csv
import sqlite3
import argparse

import msoffcrypto
from openpyxl import load_workbook


# ============================================================
# PRINT / LOG HELPER
# ============================================================

def default_log(message):
    print(message)


# ============================================================
# EXCEL OPEN HELPER
# ============================================================

def open_excel_workbook(excel_file, password="12", data_only=True, read_only=True):
    """
    Open normal or password-protected Excel files.

    Normal .xlsx/.xlsm files are opened directly with openpyxl.
    Password-protected files are decrypted first using msoffcrypto,
    then opened with openpyxl.

    Password used by default: "12"
    """
    try:
        return load_workbook(
            excel_file,
            data_only=data_only,
            read_only=read_only,
        )

    except Exception as normal_error:
        decrypted_file = BytesIO()

        try:
            with open(excel_file, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted_file)

            decrypted_file.seek(0)

            return load_workbook(
                decrypted_file,
                data_only=data_only,
                read_only=read_only,
            )

        except Exception as password_error:
            raise RuntimeError(
                f"Could not open Excel file normally or with password '{password}'. "
                f"Normal error: {normal_error}. "
                f"Password error: {password_error}"
            )


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


def parse_ph_number(value):
    """
    Convert pH value to float.

    Valid:
        7
        7.1
        7,1
        8.25

    Invalid:
        ""
        "-"
        "---"
        "NO DATA"
        "abc"
        "S/D"
    """
    text = normalize_ph(value)

    if text == "":
        return None

    if set(text) <= {"-"}:
        return None

    try:
        return float(text)
    except ValueError:
        return None


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


def should_stop_file(raw_code, raw_ph):
    """
    Stop if ANY of these is true:

      1. Column C has no code.
      2. Column C has at least 3 '-' characters.
      3. Column C does not contain a valid SU number.
      4. Column H has no meaningful numeric pH value.
    """
    code = normalize_code(raw_code)
    ph_text = normalize_ph(raw_ph)

    su_number = extract_su_number(code)
    ph_number = parse_ph_number(raw_ph)

    if code == "":
        return True, "Column C is empty"

    if code.count("-") >= 3:
        return True, f"Column C has at least 3 '-' characters: {code!r}"

    if su_number is None:
        return True, f"Column C does not contain a valid SU code: {code!r}"

    if ph_number is None:
        return True, f"Column H does not contain a valid numeric pH value: {ph_text!r}"

    return False, ""


# ============================================================
# FILE FILTER
# ============================================================

def get_excel_files(ph_folder, file_name_filter, log):
    """
    Return only .xlsx/.xlsm files whose name contains file_name_filter.
    Case-insensitive.
    """
    all_excel_files = sorted(
        list(ph_folder.glob("*.xlsx")) +
        list(ph_folder.glob("*.xlsm"))
    )

    # Ignore Excel temporary lock files
    all_excel_files = [
        path for path in all_excel_files
        if not path.name.startswith("~$")
    ]

    filtered_files = [
        path for path in all_excel_files
        if file_name_filter.upper() in path.name.upper()
    ]

    log("")
    log("File filter:")
    log(f"  Only reading files containing: {file_name_filter}")
    log(f"  Total Excel files found: {len(all_excel_files)}")
    log(f"  Selected files:          {len(filtered_files)}")

    log("")
    log("Selected files:")
    for path in filtered_files:
        log(f"  - {path.name}")

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
            ph REAL,

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

def build_ph_database(
    ph_folder,
    database_file,
    csv_file,
    file_name_filter="Ver.03",
    ph_sheet_name="F-103",
    excel_password="12",
    ph_first_row=27,
    block_size=21,
    gap_rows=2,
    duplicate_col="B",
    code_col="C",
    ph_value_col="H",
    verbose=True,
    log_callback=None,
):
    """
    Main function.

    Can be used from:
      1. command line
      2. Flask app
      3. another Python script
    """

    log = log_callback if log_callback is not None else default_log

    ph_folder = Path(ph_folder)
    db_path = Path(database_file)
    csv_path = Path(csv_file)

    db_path.parent.mkdir(parents=True, exist_ok=True)
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    log("")
    log("============================================================")
    log("BUILDING PH DATABASE FROM EXCEL FILES")
    log("============================================================")
    log("pH folder:")
    log(f"  {ph_folder}")
    log("SQLite database:")
    log(f"  {db_path}")
    log("CSV output:")
    log(f"  {csv_path}")
    log("pH sheet:")
    log(f"  {ph_sheet_name}")
    log("Filename filter:")
    log(f"  {file_name_filter}")
    log("Excel password:")
    log(f"  {excel_password}")
    log("Columns:")
    log(f"  {duplicate_col} = duplicate")
    log(f"  {code_col} = code")
    log(f"  {ph_value_col} = ph")
    log("Rows:")
    log(f"  First row: {ph_first_row}")
    log(f"  Block size: {block_size}")
    log(f"  Gap rows: {gap_rows}")
    log("")
    log("Stop rules with OR logic:")
    log("  Stop current file if C is empty")
    log("  OR C has 3 or more '-' characters")
    log("  OR C does not contain a valid SU code")
    log("  OR H is not a meaningful numeric pH value")
    log("============================================================")

    if not ph_folder.exists():
        raise FileNotFoundError(f"pH folder not found: {ph_folder}")

    excel_files = get_excel_files(ph_folder, file_name_filter, log)

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
    total_files_stopped = 0
    total_errors = 0

    for file_index, excel_file in enumerate(excel_files, start=1):
        log("")
        log("============================================================")
        log(f"[{file_index}/{len(excel_files)}] ACCESSING EXCEL FILE")
        log("============================================================")
        log("File:")
        log(f"  {excel_file}")

        try:
            log("")
            log("Opening workbook...")
            wb = open_excel_workbook(
                excel_file,
                password=excel_password,
                data_only=True,
                read_only=True,
            )

            log("Workbook opened.")
            log("Available sheets:")
            for sheet in wb.sheetnames:
                log(f"  - {sheet}")

            if ph_sheet_name not in wb.sheetnames:
                log("")
                log(f"SKIPPED FILE: sheet '{ph_sheet_name}' not found.")
                wb.close()
                continue

            ws = wb[ph_sheet_name]

            log("")
            log(f"Accessing sheet: {ph_sheet_name}")
            log(f"Excel reported max row: {ws.max_row}")
            log("The script will stop when the first non-meaningful row is found.")

            file_inserted = 0
            stop_reason = ""

            block_start = ph_first_row

            while True:
                block_end = block_start + block_size - 1

                log("")
                log("------------------------------------------------------------")
                log(f"Reading pH block: rows {block_start} to {block_end}")
                log("------------------------------------------------------------")

                stop_this_file = False

                for row in range(block_start, block_end + 1):
                    duplicate_cell = f"{duplicate_col}{row}"
                    code_cell = f"{code_col}{row}"
                    ph_cell = f"{ph_value_col}{row}"

                    raw_duplicate = ws[duplicate_cell].value
                    raw_code = ws[code_cell].value
                    raw_ph = ws[ph_cell].value

                    duplicate = normalize_duplicate(raw_duplicate)
                    code = normalize_code(raw_code)
                    ph_text = normalize_ph(raw_ph)
                    ph_number = parse_ph_number(raw_ph)
                    su_number = extract_su_number(code)

                    if verbose:
                        log("")
                        log(f"Reading row {row}:")
                        log(f"  {duplicate_cell} duplicate raw = {raw_duplicate!r} -> {duplicate!r}")
                        log(f"  {code_cell} code raw      = {raw_code!r} -> {code!r}")
                        log(f"  {ph_cell} ph raw          = {raw_ph!r} -> {ph_text!r}")
                        log(f"  pH numeric                = {ph_number}")
                        log(f"  Dash count in code        = {code.count('-')}")
                        log(f"  Extracted SU number       = {su_number}")

                    stop, reason = should_stop_file(raw_code, raw_ph)

                    if stop:
                        log("")
                        log("  STOP FILE:")
                        log(f"    Reason: {reason}")
                        log("    Moving to the next Excel file.")
                        stop_this_file = True
                        stop_reason = reason
                        total_files_stopped += 1
                        break

                    insert_record(
                        conn=conn,
                        duplicate=duplicate,
                        code=code,
                        ph=ph_number,
                        su_number=su_number,
                        source_file=excel_file.name,
                        source_sheet=ph_sheet_name,
                        source_row=row,
                    )

                    writer.writerow(
                        [
                            duplicate,
                            code,
                            ph_number,
                            su_number,
                            excel_file.name,
                            ph_sheet_name,
                            row,
                        ]
                    )

                    if verbose:
                        log("  INSERTED INTO DATABASE AND CSV:")
                        log(f"    duplicate = {duplicate}")
                        log(f"    code      = {code}")
                        log(f"    ph        = {ph_number}")
                        log(f"    file      = {excel_file.name}")
                        log(f"    row       = {row}")

                    file_inserted += 1

                if stop_this_file:
                    break

                block_start = block_end + gap_rows + 1

                if block_start > ws.max_row + block_size + gap_rows:
                    log("")
                    log("  SAFETY STOP:")
                    log("    Reached beyond Excel's reported max row.")
                    log("    Moving to next file.")
                    stop_reason = "Safety stop beyond Excel max row"
                    break

            conn.commit()
            wb.close()

            total_inserted += file_inserted

            log("")
            log("------------------------------------------------------------")
            log(f"Finished file: {excel_file.name}")
            log(f"Inserted rows: {file_inserted}")
            log(f"Stop reason:   {stop_reason}")
            log("------------------------------------------------------------")

        except Exception as e:
            total_errors += 1
            log("")
            log("ERROR while reading file:")
            log(f"  {excel_file}")
            log("Exception:")
            log(f"  {e}")

    csv_handle.close()

    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM ph_data")
    db_count = cur.fetchone()[0]

    conn.close()

    log("")
    log("============================================================")
    log("DATABASE BUILD FINISHED")
    log("============================================================")
    log(f"Total inserted rows: {total_inserted}")
    log(f"Rows in SQLite DB:   {db_count}")
    log(f"Files stopped by OR stop rule: {total_files_stopped}")
    log(f"Files with errors: {total_errors}")
    log("")
    log("Created:")
    log(f"  {db_path}")
    log(f"  {csv_path}")

    return {
        "total_inserted": total_inserted,
        "db_count": db_count,
        "files_stopped": total_files_stopped,
        "errors": total_errors,
        "database_file": str(db_path),
        "csv_file": str(csv_path),
    }


# ============================================================
# COMMAND LINE ARGUMENTS
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Build pH SQLite database and CSV from Ver.03 Excel files."
    )

    parser.add_argument(
        "--ph-folder",
        required=True,
        help="Folder containing the pH Excel files."
    )

    parser.add_argument(
        "--database-file",
        required=True,
        help="Output SQLite database path."
    )

    parser.add_argument(
        "--csv-file",
        required=True,
        help="Output CSV file path."
    )

    parser.add_argument(
        "--file-name-filter",
        default="Ver.03",
        help="Only process Excel files whose names contain this text. Default: Ver.03"
    )

    parser.add_argument(
        "--sheet-name",
        default="F-103",
        help="Excel sheet name to read. Default: F-103"
    )

    parser.add_argument(
        "--excel-password",
        default="12",
        help="Password for protected Excel files. Default: 12"
    )

    parser.add_argument(
        "--first-row",
        type=int,
        default=27,
        help="First row to read. Default: 27"
    )

    parser.add_argument(
        "--block-size",
        type=int,
        default=21,
        help="Number of rows in each block. Default: 21"
    )

    parser.add_argument(
        "--gap-rows",
        type=int,
        default=2,
        help="Number of rows between blocks. Default: 2"
    )

    parser.add_argument(
        "--duplicate-col",
        default="B",
        help="Column for duplicate value. Default: B"
    )

    parser.add_argument(
        "--code-col",
        default="C",
        help="Column for SU code. Default: C"
    )

    parser.add_argument(
        "--ph-col",
        default="H",
        help="Column for pH value. Default: H"
    )

    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Reduce row-by-row logging."
    )

    return parser.parse_args()


def main():
    args = parse_args()

    build_ph_database(
        ph_folder=args.ph_folder,
        database_file=args.database_file,
        csv_file=args.csv_file,
        file_name_filter=args.file_name_filter,
        ph_sheet_name=args.sheet_name,
        excel_password=args.excel_password,
        ph_first_row=args.first_row,
        block_size=args.block_size,
        gap_rows=args.gap_rows,
        duplicate_col=args.duplicate_col,
        code_col=args.code_col,
        ph_value_col=args.ph_col,
        verbose=not args.quiet,
    )


if __name__ == "__main__":
    main()
