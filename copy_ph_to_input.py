# fill_input_xlsm_from_ph_csv.py

from pathlib import Path
import re
import csv
import shutil
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_XLSM = Path(
    r"G:\Mi unidad\LABSAF ILLPA\input.xlsm"
)

PH_CSV = Path(
    r"G:\Mi unidad\LABSAF ILLPA\ph_database_Ver03.csv"
)

OUTPUT_XLSM = Path(
    r"G:\Mi unidad\LABSAF ILLPA\input_WITH_PH_FROM_CSV.xlsm"
)

# Input workbook tab
INPUT_SHEET_NAME = "P_DIS"

# Input structure:
# Read:
#   C37:C56
#   C60:C79
#   C83:C102
#   C106:C125
#   ...
#
# Ignore D rows:
#   36, 59, 82, 105, ...
INPUT_FIRST_ROW = 37
BLOCK_SIZE = 20
GAP_ROWS = 3

INPUT_CODE_COL = "C"
INPUT_OUTPUT_COL = "E"

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
    Clean text.

    Handles:
        SU1149-ILL-26
        'SU1149-ILL-26
        spaces
        non-breaking spaces

    The leading apostrophe is ignored.
    """
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    # Ignore leading Excel apostrophe
    text = text.lstrip("'").strip()

    return text


def normalize_code(value):
    """
    Normalize SU code.
    """
    return normalize_text(value).upper()


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


def parse_ph(value):
    """
    Convert pH to float.

    Handles:
        7.5
        7,5
        "7.5"

    Invalid:
        ""
        "-"
        "---"
        "NO DATA"
    """
    text = normalize_text(value)

    if text == "":
        return None

    text = text.replace(",", ".")

    if set(text) <= {"-"}:
        return None

    try:
        return float(text)
    except ValueError:
        return None


# ============================================================
# CSV COLUMN HELPERS
# ============================================================

def find_column(fieldnames, possible_names):
    """
    Find a column using several possible names, case-insensitive.
    """
    normalized = {
        name.strip().lower(): name
        for name in fieldnames
    }

    for possible in possible_names:
        key = possible.strip().lower()
        if key in normalized:
            return normalized[key]

    return None


# ============================================================
# BUILD LOOKUP FROM CSV
# ============================================================

def build_ph_lookup_from_csv(csv_path):
    """
    Read ph_database_Ver03.csv and build:

        su_number -> best record

    If the same SU appears more than once, keep the largest pH.
    """
    print("")
    print("============================================================")
    print("READING pH CSV DATABASE")
    print("============================================================")
    print("CSV file:")
    print(f"  {csv_path}")

    if not csv_path.exists():
        raise FileNotFoundError(f"pH CSV not found: {csv_path}")

    lookup = {}

    total_rows = 0
    inserted_rows = 0
    skipped_rows = 0
    duplicate_replaced = 0
    duplicate_kept = 0

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)

        if reader.fieldnames is None:
            raise ValueError("CSV file has no header row.")

        print("")
        print("Detected CSV columns:")
        for col in reader.fieldnames:
            print(f"  - {col}")

        code_col = find_column(
            reader.fieldnames,
            ["code", "su_code", "SU_CODE", "codigo", "código"]
        )

        ph_col = find_column(
            reader.fieldnames,
            ["ph", "PH", "ph_value", "PH_VALUE"]
        )

        source_file_col = find_column(
            reader.fieldnames,
            ["source_file", "SOURCE_FILE"]
        )

        source_row_col = find_column(
            reader.fieldnames,
            ["source_row", "SOURCE_ROW"]
        )

        if code_col is None:
            raise ValueError(
                "Could not find code column in CSV. "
                "Expected one of: code, su_code, SU_CODE, codigo."
            )

        if ph_col is None:
            raise ValueError(
                "Could not find pH column in CSV. "
                "Expected one of: ph, PH, ph_value, PH_VALUE."
            )

        print("")
        print(f"Using code column: {code_col}")
        print(f"Using pH column:   {ph_col}")

        for csv_row_number, row in enumerate(reader, start=2):
            total_rows += 1

            raw_code = row.get(code_col, "")
            raw_ph = row.get(ph_col, "")

            code = normalize_code(raw_code)
            su_number = extract_su_number(code)
            ph = parse_ph(raw_ph)

            source_file = row.get(source_file_col, "") if source_file_col else ""
            source_row = row.get(source_row_col, "") if source_row_col else ""

            if VERBOSE:
                print("")
                print(f"CSV row {csv_row_number}:")
                print(f"  code raw = {raw_code!r} -> {code!r}")
                print(f"  SU number = {su_number}")
                print(f"  pH raw = {raw_ph!r} -> {ph}")
                print(f"  source file = {source_file}")
                print(f"  source row  = {source_row}")

            if code == "":
                print("  SKIPPED CSV ROW: empty code")
                skipped_rows += 1
                continue

            if su_number is None:
                print("  SKIPPED CSV ROW: no SU number")
                skipped_rows += 1
                continue

            if ph is None:
                print("  SKIPPED CSV ROW: invalid pH")
                skipped_rows += 1
                continue

            record = {
                "su_number": su_number,
                "code": code,
                "ph": ph,
                "source_file": source_file,
                "source_row": source_row,
                "csv_row": csv_row_number,
            }

            if su_number not in lookup:
                lookup[su_number] = record
                inserted_rows += 1

                if VERBOSE:
                    print("  INSERTED INTO LOOKUP")

            else:
                old_record = lookup[su_number]

                if ph > old_record["ph"]:
                    lookup[su_number] = record
                    duplicate_replaced += 1

                    print("")
                    print("  DUPLICATE SU FOUND:")
                    print(f"    SU{su_number}")
                    print(f"    Old pH: {old_record['ph']} from {old_record['source_file']}")
                    print(f"    New pH: {ph} from {source_file}")
                    print("    Selected new value because it is larger.")

                else:
                    duplicate_kept += 1

                    print("")
                    print("  DUPLICATE SU FOUND:")
                    print(f"    SU{su_number}")
                    print(f"    Existing pH: {old_record['ph']} from {old_record['source_file']}")
                    print(f"    New pH: {ph} from {source_file}")
                    print("    Kept existing value because it is larger or equal.")

    print("")
    print("============================================================")
    print("CSV LOOKUP BUILD FINISHED")
    print("============================================================")
    print(f"Total CSV rows read:     {total_rows}")
    print(f"Inserted unique SUs:     {inserted_rows}")
    print(f"Skipped CSV rows:        {skipped_rows}")
    print(f"Duplicates replaced:     {duplicate_replaced}")
    print(f"Duplicates kept:         {duplicate_kept}")
    print(f"Final lookup size:       {len(lookup)}")

    return lookup


# ============================================================
# INPUT ROW ITERATOR
# ============================================================

def iter_input_rows(ws):
    """
    Generate input rows:

        37 to 56
        60 to 79
        83 to 102
        106 to 125
        ...

    This ignores the D rows:
        36, 59, 82, 105, ...
    """
    row = INPUT_FIRST_ROW

    while row <= ws.max_row:
        block_start = row
        block_end = row + BLOCK_SIZE - 1

        print("")
        print("------------------------------------------------------------")
        print(f"Input block: rows {block_start} to {block_end}")
        print("------------------------------------------------------------")

        for r in range(block_start, min(block_end, ws.max_row) + 1):
            yield r

        row = block_end + GAP_ROWS + 1


# ============================================================
# SHEET HELPER
# ============================================================

def get_sheet(workbook, sheet_name):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {workbook.sheetnames}"
        )

    return workbook[sheet_name]


# ============================================================
# FILL INPUT XLSM
# ============================================================

def fill_input_xlsm_from_lookup(ph_lookup):
    print("")
    print("============================================================")
    print("FILLING INPUT XLSM FROM pH CSV LOOKUP")
    print("============================================================")
    print("Input XLSM:")
    print(f"  {INPUT_XLSM}")
    print("Input sheet:")
    print(f"  {INPUT_SHEET_NAME}")
    print("Output XLSM:")
    print(f"  {OUTPUT_XLSM}")
    print("Reading input rows:")
    print("  C37:C56, C60:C79, C83:C102, ...")
    print("Writing pH to:")
    print("  E37:E56, E60:E79, E83:E102, ...")
    print("============================================================")

    if not INPUT_XLSM.exists():
        raise FileNotFoundError(f"Input XLSM not found: {INPUT_XLSM}")

    if not INPUT_XLSM.parent.exists():
        raise FileNotFoundError(f"Input folder not found: {INPUT_XLSM.parent}")

    print("")
    print("Copying input XLSM to output XLSM...")
    shutil.copy2(INPUT_XLSM, OUTPUT_XLSM)
    print("Copy created:")
    print(f"  {OUTPUT_XLSM}")

    print("")
    print("Opening input workbook for reading values...")
    wb_values = load_workbook(INPUT_XLSM, data_only=True, keep_vba=True)
    ws_values = get_sheet(wb_values, INPUT_SHEET_NAME)

    print(f"Input sheet used: {ws_values.title}")
    print(f"Input max row: {ws_values.max_row}")

    print("")
    print("Opening output workbook for writing values...")
    wb_out = load_workbook(OUTPUT_XLSM, keep_vba=True)
    ws_out = get_sheet(wb_out, INPUT_SHEET_NAME)

    written = 0
    not_found = 0
    skipped = 0

    log = []

    for row in iter_input_rows(ws_values):
        input_cell = f"{INPUT_CODE_COL}{row}"
        output_cell = f"{INPUT_OUTPUT_COL}{row}"

        raw_code = ws_values[input_cell].value
        code = normalize_code(raw_code)
        su_number = extract_su_number(code)

        print("")
        print("============================================================")
        print(f"INPUT ROW {row}")
        print("============================================================")
        print(f"Reading {input_cell}:")
        print(f"  Raw code:        {raw_code!r}")
        print(f"  Normalized code: {code!r}")
        print(f"  SU number:       {su_number}")

        if code == "":
            print("  SKIPPED: empty input code")
            skipped += 1
            continue

        if su_number is None:
            print("  SKIPPED: no SU number found")
            skipped += 1

            log.append(
                {
                    "row": row,
                    "code": code,
                    "su_number": "",
                    "ph": "",
                    "status": "SKIPPED: no SU number",
                    "source_file": "",
                    "source_row": "",
                }
            )
            continue

        if su_number not in ph_lookup:
            print("  NOT FOUND in ph_database_Ver03.csv lookup")
            not_found += 1

            log.append(
                {
                    "row": row,
                    "code": code,
                    "su_number": su_number,
                    "ph": "",
                    "status": "NOT FOUND",
                    "source_file": "",
                    "source_row": "",
                }
            )
            continue

        record = ph_lookup[su_number]
        ph = record["ph"]

        print("  FOUND pH:")
        print(f"    pH:          {ph}")
        print(f"    CSV code:    {record['code']}")
        print(f"    Source file: {record['source_file']}")
        print(f"    Source row:  {record['source_row']}")
        print(f"    CSV row:     {record['csv_row']}")

        print("")
        print("  Writing to output workbook:")
        print(f"    Destination cell: {output_cell}")
        print(f"    Value: {ph}")

        ws_out[output_cell] = ph
        written += 1

        log.append(
            {
                "row": row,
                "code": code,
                "su_number": su_number,
                "ph": ph,
                "status": "OK",
                "source_file": record["source_file"],
                "source_row": record["source_row"],
            }
        )

    print("")
    print("Saving output workbook...")
    wb_out.save(OUTPUT_XLSM)

    wb_values.close()
    wb_out.close()

    print("")
    print("============================================================")
    print("FILLING FINISHED")
    print("============================================================")
    print(f"Output file created:")
    print(f"  {OUTPUT_XLSM}")
    print(f"Written rows: {written}")
    print(f"Not found:    {not_found}")
    print(f"Skipped:      {skipped}")

    print("")
    print("FINAL SUMMARY")
    print("============================================================")

    for item in log:
        print(
            f"Row {item['row']:>4} | "
            f"C={item['code']:<25} | "
            f"SU={str(item['su_number']):<8} | "
            f"pH={str(item['ph']):<8} | "
            f"{item['status']:<20} | "
            f"{item['source_file']} | "
            f"source row={item['source_row']}"
        )


# ============================================================
# MAIN
# ============================================================

def main():
    ph_lookup = build_ph_lookup_from_csv(PH_CSV)

    fill_input_xlsm_from_lookup(ph_lookup)


if __name__ == "__main__":
    main()
