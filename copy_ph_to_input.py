# copy_ph_to_input.py

from pathlib import Path
import re
import shutil
from openpyxl import load_workbook


# ------------------------------------------------------------
# CONFIGURATION
# ------------------------------------------------------------

INPUT_XLSM = r"G:\Mi unidad\LABSAF ILLPA\input.xlsm"

PH_FOLDER = r"G:\Mi unidad\LABSAF ILLPA\1. Documentos Internos\7.5 Registros Tecnicos\2026\SUELOS\1.pH"

OUTPUT_XLSM = r"G:\Mi unidad\LABSAF ILLPA\input_WITH_PH.xlsm"

# Sheet name in the input file.
# If None, it uses the active sheet.
INPUT_SHEET_NAME = None

# Sheet name in pH files.
# If None, it uses the active sheet.
PH_SHEET_NAME = None

# Input structure
INPUT_FIRST_ROW = 36
BLOCK_SIZE = 21          # D + 20 samples
GAP_ROWS = 2             # skip 2 rows between blocks
INPUT_CODE_COL = "C"
INPUT_MARKER_COL = "B"
INPUT_OUTPUT_COL = "E"

# pH file structure
PH_FIRST_ROW = 27
PH_MARKER_COL = "B"
PH_CODE_COL = "C"
PH_VALUE_COL = "H"


# ------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------

def normalize_text(value):
    """
    Convert Excel value to clean string.
    """
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_marker(value):
    """
    Normalize column B values.

    Examples:
        D -> D
        1 -> 1
        1.0 -> 1
        " 1 " -> 1
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
    Extract numeric SU number from something like:

        SU1149-ILL-26
        SU0079
        SU10747-10859

    Returns an integer, for example:
        SU1149-ILL-26 -> 1149
        SU0079 -> 79
    """
    text = normalize_text(value).upper()

    match = re.search(r"SU\s*0*(\d+)", text)
    if not match:
        return None

    return int(match.group(1))


def extract_all_su_numbers(text):
    """
    Extract all SU-like numbers from a filename.

    Example:
        "SU0597-0602- SU0608-0667.xlsx"

    Returns:
        [597, 602, 608, 667]
    """
    text = normalize_text(text).upper()

    numbers = []

    # Find explicit SU numbers
    for match in re.finditer(r"SU\s*0*(\d+)", text):
        numbers.append(int(match.group(1)))

        # After SU0597, the filename may continue with -0602
        # This captures the number immediately after the dash.
        after = text[match.end():]
        dash_match = re.match(r"\s*-\s*0*(\d+)", after)
        if dash_match:
            numbers.append(int(dash_match.group(1)))

    return numbers


def extract_su_ranges_from_filename(filename):
    """
    Extract SU ranges from filenames.

    Supports:
        SU0668-0767.xlsx
        SU0597-0602- SU0608-0667.xlsx
        SU1144-1175.xlsx
        SU10268-10269.xlsx

    Returns list of tuples:
        [(668, 767)]
        [(597, 602), (608, 667)]
    """
    text = normalize_text(filename).upper()

    ranges = []

    # Matches:
    # SU0668-0767
    # SU0597-0602
    # SU0608-0667
    pattern = r"SU\s*0*(\d+)\s*-\s*0*(\d+)"

    for start, end in re.findall(pattern, text):
        start = int(start)
        end = int(end)

        if start <= end:
            ranges.append((start, end))
        else:
            ranges.append((end, start))

    return ranges


def file_contains_su_number(path, su_number):
    """
    Check whether a pH filename contains a range that includes su_number.
    """
    ranges = extract_su_ranges_from_filename(path.name)

    for start, end in ranges:
        if start <= su_number <= end:
            return True

    return False


def find_candidate_ph_files(ph_folder, su_number):
    """
    Return all pH files whose filename range contains the SU number.
    """
    ph_folder = Path(ph_folder)

    candidates = []

    for path in ph_folder.glob("*.xlsx"):
        if file_contains_su_number(path, su_number):
            candidates.append(path)

    for path in ph_folder.glob("*.xlsm"):
        if file_contains_su_number(path, su_number):
            candidates.append(path)

    return sorted(candidates)


def get_sheet(workbook, sheet_name=None):
    """
    Return selected sheet or active sheet.
    """
    if sheet_name is None:
        return workbook.active

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}"
        )

    return workbook[sheet_name]


def codes_match(input_code, ph_code):
    """
    Match by exact normalized text or by SU number.

    This allows:
        SU1149-ILL-26
    to match cells that contain:
        SU1149
        SU1149-ILL-26
    """
    input_text = normalize_text(input_code).upper()
    ph_text = normalize_text(ph_code).upper()

    if input_text == "" or ph_text == "":
        return False

    if input_text == ph_text:
        return True

    input_su = extract_su_number(input_text)
    ph_su = extract_su_number(ph_text)

    if input_su is not None and ph_su is not None:
        return input_su == ph_su

    return False


def find_ph_value_in_file(ph_file, input_code, input_marker):
    """
    Open one pH Excel file and search from PH_FIRST_ROW down.

    Matching rule:
        1. Column C must match the SU code or SU number.
        2. Column B should match the input marker D, 1, 2, etc.

    If exact marker match is not found, it returns a code-only fallback
    only if the code appears once.
    """
    wb = load_workbook(ph_file, data_only=True, read_only=True)
    ws = get_sheet(wb, PH_SHEET_NAME)

    input_marker_norm = normalize_marker(input_marker)

    exact_matches = []
    code_only_matches = []

    for row in range(PH_FIRST_ROW, ws.max_row + 1):
        ph_marker = ws[f"{PH_MARKER_COL}{row}"].value
        ph_code = ws[f"{PH_CODE_COL}{row}"].value
        ph_value = ws[f"{PH_VALUE_COL}{row}"].value

        if not codes_match(input_code, ph_code):
            continue

        ph_marker_norm = normalize_marker(ph_marker)

        code_only_matches.append(
            {
                "row": row,
                "marker": ph_marker_norm,
                "code": ph_code,
                "value": ph_value,
            }
        )

        if ph_marker_norm == input_marker_norm:
            exact_matches.append(
                {
                    "row": row,
                    "marker": ph_marker_norm,
                    "code": ph_code,
                    "value": ph_value,
                }
            )

    wb.close()

    if len(exact_matches) == 1:
        return exact_matches[0]["value"], exact_matches[0]["row"], "exact marker + code match"

    if len(exact_matches) > 1:
        raise ValueError(
            f"Multiple exact matches in {ph_file.name} for code={input_code}, "
            f"marker={input_marker}. Rows: {[m['row'] for m in exact_matches]}"
        )

    if len(code_only_matches) == 1:
        return code_only_matches[0]["value"], code_only_matches[0]["row"], "code-only fallback"

    if len(code_only_matches) > 1:
        raise ValueError(
            f"Code found multiple times in {ph_file.name}, but marker did not match. "
            f"Input code={input_code}, input marker={input_marker}. "
            f"Candidate rows: {[(m['row'], m['marker']) for m in code_only_matches]}"
        )

    return None, None, "not found"


def iter_input_rows(ws):
    """
    Generate rows in this pattern:

        36 to 56
        59 to 79
        82 to 102
        ...

    Each block has 21 rows.
    Between blocks there are 2 skipped rows.
    """
    row = INPUT_FIRST_ROW

    while row <= ws.max_row:
        block_start = row
        block_end = row + BLOCK_SIZE - 1

        for r in range(block_start, min(block_end, ws.max_row) + 1):
            yield r

        row = block_end + GAP_ROWS + 1


# ------------------------------------------------------------
# MAIN FUNCTION
# ------------------------------------------------------------

def main():
    input_path = Path(INPUT_XLSM)
    output_path = Path(OUTPUT_XLSM)
    ph_folder = Path(PH_FOLDER)

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if not ph_folder.exists():
        raise FileNotFoundError(f"pH folder not found: {ph_folder}")

    # Copy original file first, so macros are preserved.
    shutil.copy2(input_path, output_path)

    # Read values from original file with data_only=True.
    # This is useful if column C has formulas.
    wb_values = load_workbook(input_path, data_only=True, keep_vba=True)
    ws_values = get_sheet(wb_values, INPUT_SHEET_NAME)

    # Open copied workbook for writing.
    # keep_vba=True is important for .xlsm files.
    wb_out = load_workbook(output_path, keep_vba=True)
    ws_out = get_sheet(wb_out, INPUT_SHEET_NAME)

    log = []

    for row in iter_input_rows(ws_values):
        input_marker = ws_values[f"{INPUT_MARKER_COL}{row}"].value
        input_code = ws_values[f"{INPUT_CODE_COL}{row}"].value

        marker_text = normalize_marker(input_marker)
        code_text = normalize_text(input_code)

        # Skip empty rows.
        if code_text == "":
            continue

        su_number = extract_su_number(code_text)

        if su_number is None:
            log.append(
                {
                    "input_row": row,
                    "marker": marker_text,
                    "code": code_text,
                    "status": "SKIPPED - no SU number found",
                    "file": "",
                    "ph_row": "",
                    "ph_value": "",
                }
            )
            continue

        candidates = find_candidate_ph_files(ph_folder, su_number)

        if not candidates:
            log.append(
                {
                    "input_row": row,
                    "marker": marker_text,
                    "code": code_text,
                    "status": "ERROR - no pH file range contains this SU",
                    "file": "",
                    "ph_row": "",
                    "ph_value": "",
                }
            )
            continue

        found_value = None
        found_file = None
        found_row = None
        found_status = None

        for ph_file in candidates:
            ph_value, ph_row, status = find_ph_value_in_file(
                ph_file=ph_file,
                input_code=code_text,
                input_marker=marker_text,
            )

            if ph_value is not None:
                found_value = ph_value
                found_file = ph_file
                found_row = ph_row
                found_status = status
                break

        if found_value is None:
            log.append(
                {
                    "input_row": row,
                    "marker": marker_text,
                    "code": code_text,
                    "status": "ERROR - pH value not found inside candidate file(s)",
                    "file": "; ".join(p.name for p in candidates),
                    "ph_row": "",
                    "ph_value": "",
                }
            )
            continue

        # Write pH value into column E of the output workbook.
        ws_out[f"{INPUT_OUTPUT_COL}{row}"] = found_value

        log.append(
            {
                "input_row": row,
                "marker": marker_text,
                "code": code_text,
                "status": found_status,
                "file": found_file.name,
                "ph_row": found_row,
                "ph_value": found_value,
            }
        )

    wb_out.save(output_path)

    wb_values.close()
    wb_out.close()

    print("\nFinished.")
    print(f"Output file created:")
    print(f"  {output_path}")

    print("\nSummary:")
    for item in log:
        print(
            f"Input row {item['input_row']:>4} | "
            f"B={item['marker']:<3} | "
            f"C={item['code']:<20} | "
            f"pH={str(item['ph_value']):<10} | "
            f"{item['status']} | "
            f"{item['file']}"
        )


if __name__ == "__main__":
    main()
