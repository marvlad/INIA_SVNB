# debug_find_one_su_in_ph_file.py

from pathlib import Path
import re
import time
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_XLSM = r"G:\Mi unidad\LABSAF ILLPA\input.xlsm"

PH_FOLDER = r"G:\Mi unidad\LABSAF ILLPA\1. Documentos Internos\7.5 Registros Tecnicos\2026\SUELOS\1.pH"

# Input file tab
INPUT_SHEET_NAME = "P_DIS"

# pH file tab
PH_SHEET_NAME = "F-103"

# Choose which row from the input file you want to test first
TEST_INPUT_ROW = 36

INPUT_MARKER_COL = "B"
INPUT_CODE_COL = "C"

PH_FIRST_ROW = 27
PH_MARKER_COL = "B"
PH_CODE_COL = "C"
PH_VALUE_COL = "H"


# ============================================================
# NORMALIZATION FUNCTIONS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    # Ignore leading apostrophe, e.g. 'SU1149-ILL-26
    text = text.lstrip("'").strip()

    return text


def normalize_code(value):
    return normalize_text(value).upper()


def normalize_marker(value):
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
    text = normalize_text(value).upper()

    match = re.search(r"SU\s*0*(\d+)", text)

    if not match:
        return None

    return int(match.group(1))


def extract_su_ranges_from_filename(filename):
    """
    Examples:
        SU0668-0767.xlsx
        SU0597-0602- SU0608-0667.xlsx

    Returns:
        [(668, 767)]
        [(597, 602), (608, 667)]
    """
    text = normalize_text(filename).upper()

    pattern = r"SU\s*0*(\d+)\s*-\s*0*(\d+)"

    ranges = []

    for start, end in re.findall(pattern, text):
        start = int(start)
        end = int(end)

        if start <= end:
            ranges.append((start, end))
        else:
            ranges.append((end, start))

    return ranges


def codes_match(input_code, ph_code):
    """
    Match either exact normalized code or just SU number.

    This allows:
        SU1149-ILL-26
        'SU1149-ILL-26
        SU1149
    """
    input_code_norm = normalize_code(input_code)
    ph_code_norm = normalize_code(ph_code)

    if input_code_norm == "" or ph_code_norm == "":
        return False

    if input_code_norm == ph_code_norm:
        return True

    input_su = extract_su_number(input_code_norm)
    ph_su = extract_su_number(ph_code_norm)

    return input_su is not None and ph_su is not None and input_su == ph_su


# ============================================================
# STEP 1: READ ONE SU FROM INPUT FILE
# ============================================================

def read_su_from_input_file():
    input_path = Path(INPUT_XLSM)

    print("")
    print("============================================================")
    print("STEP 1: OPENING INPUT XLSM")
    print("============================================================")
    print(f"Input file:")
    print(f"  {input_path}")

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    print("")
    print("Opening workbook with openpyxl...")
    wb = load_workbook(input_path, data_only=True, keep_vba=True)

    print(f"Workbook opened.")
    print(f"Available sheets:")
    for sheet in wb.sheetnames:
        print(f"  - {sheet}")

    if INPUT_SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"Input sheet '{INPUT_SHEET_NAME}' not found. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb[INPUT_SHEET_NAME]

    print("")
    print(f"Accessing input sheet:")
    print(f"  {INPUT_SHEET_NAME}")

    marker_cell = f"{INPUT_MARKER_COL}{TEST_INPUT_ROW}"
    code_cell = f"{INPUT_CODE_COL}{TEST_INPUT_ROW}"

    raw_marker = ws[marker_cell].value
    raw_code = ws[code_cell].value

    marker = normalize_marker(raw_marker)
    code = normalize_code(raw_code)
    su_number = extract_su_number(code)

    print("")
    print("Reading cells from input file:")
    print(f"  {marker_cell} = {raw_marker!r}  -> normalized marker = {marker!r}")
    print(f"  {code_cell} = {raw_code!r}  -> normalized code = {code!r}")
    print(f"  Extracted SU number = {su_number}")

    wb.close()

    if code == "":
        raise ValueError(f"No SU code found in {code_cell}")

    if su_number is None:
        raise ValueError(f"Could not extract SU number from {code_cell}: {raw_code!r}")

    return {
        "row": TEST_INPUT_ROW,
        "marker": marker,
        "code": code,
        "su_number": su_number,
    }


# ============================================================
# STEP 2: WAIT 5 SECONDS AND FIND PROBABLE PH FILE
# ============================================================

def find_probable_ph_files(su_number):
    ph_folder = Path(PH_FOLDER)

    print("")
    print("============================================================")
    print("STEP 2: WAITING 5 SECONDS BEFORE LOOKING FOR pH FILE")
    print("============================================================")
    print(f"SU number from input file:")
    print(f"  SU{su_number}")

    print("")
    print("Waiting 5 seconds...")
    time.sleep(5)

    print("")
    print("Now checking pH folder:")
    print(f"  {ph_folder}")

    if not ph_folder.exists():
        raise FileNotFoundError(f"pH folder not found: {ph_folder}")

    excel_files = sorted(
        list(ph_folder.glob("*.xlsx")) +
        list(ph_folder.glob("*.xlsm"))
    )

    print("")
    print(f"Excel files found in pH folder: {len(excel_files)}")

    candidates = []

    print("")
    print("Checking filename ranges:")
    print("------------------------------------------------------------")

    for excel_file in excel_files:
        ranges = extract_su_ranges_from_filename(excel_file.name)

        print("")
        print(f"File:")
        print(f"  {excel_file.name}")

        if not ranges:
            print("  No SU range found in filename.")
            continue

        for start, end in ranges:
            print(f"  Range detected: SU{start} to SU{end}")

            if start <= su_number <= end:
                print(f"  RESULT: YES, SU{su_number} is probably in this file.")
                candidates.append(excel_file)
            else:
                print(f"  RESULT: NO, SU{su_number} is not in this range.")

    print("")
    print("------------------------------------------------------------")
    print(f"Candidate pH files found: {len(candidates)}")

    for candidate in candidates:
        print(f"  - {candidate.name}")

    return candidates


# ============================================================
# STEP 3: OPEN CANDIDATE PH FILE AND SEARCH INSIDE
# ============================================================

def search_su_inside_ph_file(ph_file, input_info):
    print("")
    print("============================================================")
    print("STEP 3: OPENING PROBABLE pH FILE")
    print("============================================================")
    print(f"pH file:")
    print(f"  {ph_file}")

    wb = load_workbook(ph_file, data_only=True, read_only=True)

    print("")
    print("pH workbook opened.")
    print("Available sheets:")
    for sheet in wb.sheetnames:
        print(f"  - {sheet}")

    if PH_SHEET_NAME not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"pH sheet '{PH_SHEET_NAME}' not found in file {ph_file.name}. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb[PH_SHEET_NAME]

    print("")
    print(f"Accessing pH sheet:")
    print(f"  {PH_SHEET_NAME}")

    print("")
    print("Searching inside pH file:")
    print(f"  Starting row: {PH_FIRST_ROW}")
    print(f"  Last row: {ws.max_row}")
    print(f"  Comparing input SU code: {input_info['code']}")
    print(f"  Comparing input SU number: {input_info['su_number']}")
    print(f"  Input marker from column B: {input_info['marker']}")
    print("")
    print("Rows checked:")
    print("------------------------------------------------------------")

    matches = []

    for row in range(PH_FIRST_ROW, ws.max_row + 1):
        raw_marker = ws[f"{PH_MARKER_COL}{row}"].value
        raw_code = ws[f"{PH_CODE_COL}{row}"].value
        raw_ph = ws[f"{PH_VALUE_COL}{row}"].value

        marker = normalize_marker(raw_marker)
        code = normalize_code(raw_code)
        su_number = extract_su_number(code)

        print(
            f"Row {row}: "
            f"B={raw_marker!r}->{marker!r}, "
            f"C={raw_code!r}->{code!r}, "
            f"SU={su_number}, "
            f"H={raw_ph!r}"
        )

        if code == "":
            print("  -> skipped: empty code")
            continue

        if codes_match(input_info["code"], code):
            print("  -> CODE MATCH FOUND")

            if marker == input_info["marker"]:
                print("  -> MARKER ALSO MATCHES")
                print(f"  -> pH value to use from H{row}: {raw_ph!r}")

                matches.append(
                    {
                        "row": row,
                        "marker": marker,
                        "code": code,
                        "ph": raw_ph,
                        "match_type": "code + marker",
                    }
                )
            else:
                print(
                    "  -> code matches, but marker does not match: "
                    f"input marker={input_info['marker']!r}, "
                    f"pH marker={marker!r}"
                )

    wb.close()

    print("")
    print("============================================================")
    print("SEARCH RESULT")
    print("============================================================")

    if not matches:
        print("No exact code + marker match found inside this pH file.")
        return None

    if len(matches) == 1:
        match = matches[0]

        print("One exact match found:")
        print(f"  pH file: {ph_file.name}")
        print(f"  Sheet: {PH_SHEET_NAME}")
        print(f"  Row: {match['row']}")
        print(f"  Marker B: {match['marker']}")
        print(f"  Code C: {match['code']}")
        print(f"  pH H: {match['ph']}")

        return match

    print("Multiple matches found:")
    for match in matches:
        print(match)

    return matches


# ============================================================
# MAIN
# ============================================================

def main():
    input_info = read_su_from_input_file()

    candidates = find_probable_ph_files(input_info["su_number"])

    if not candidates:
        print("")
        print("No probable pH file found by filename range.")
        return

    for candidate in candidates:
        result = search_su_inside_ph_file(candidate, input_info)

        if result:
            print("")
            print("Final result:")
            print(f"  Input row: {input_info['row']}")
            print(f"  Input code: {input_info['code']}")
            print(f"  Input marker: {input_info['marker']}")
            print(f"  Probable pH file: {candidate.name}")
            print(f"  pH value found: {result['ph'] if isinstance(result, dict) else result}")
            break


if __name__ == "__main__":
    main()
