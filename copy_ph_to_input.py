# copy_ph_to_input_VERBOSE.py

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

INPUT_SHEET_NAME = None
PH_SHEET_NAME = None

INPUT_FIRST_ROW = 36
BLOCK_SIZE = 21
GAP_ROWS = 2

INPUT_MARKER_COL = "B"
INPUT_CODE_COL = "C"
INPUT_OUTPUT_COL = "E"

PH_FIRST_ROW = 27
PH_MARKER_COL = "B"
PH_CODE_COL = "C"
PH_VALUE_COL = "H"

VERBOSE = True


# ------------------------------------------------------------
# VERBOSE PRINT
# ------------------------------------------------------------

def vprint(message):
    if VERBOSE:
        print(message)


# ------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text


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
    text = normalize_text(filename).upper()

    ranges = []

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
    ranges = extract_su_ranges_from_filename(path.name)

    vprint(f"      Checking file range: {path.name}")

    if not ranges:
        vprint("        No SU range found in filename.")
        return False

    for start, end in ranges:
        vprint(f"        Range found: SU{start} to SU{end}")

        if start <= su_number <= end:
            vprint(f"        YES: SU{su_number} is inside SU{start}-{end}")
            return True
        else:
            vprint(f"        NO: SU{su_number} is not inside SU{start}-{end}")

    return False


def find_candidate_ph_files(ph_folder, su_number):
    ph_folder = Path(ph_folder)

    vprint("")
    vprint(f"    Looking for pH Excel file containing SU{su_number}")
    vprint(f"    Folder: {ph_folder}")

    candidates = []

    excel_files = list(ph_folder.glob("*.xlsx")) + list(ph_folder.glob("*.xlsm"))

    vprint(f"    Number of Excel files found in pH folder: {len(excel_files)}")

    for path in sorted(excel_files):
        if file_contains_su_number(path, su_number):
            candidates.append(path)

    vprint(f"    Candidate files found: {len(candidates)}")

    for candidate in candidates:
        vprint(f"      Candidate: {candidate.name}")

    return sorted(candidates)


def get_sheet(workbook, sheet_name=None):
    if sheet_name is None:
        ws = workbook.active
        vprint(f"    Using active sheet: {ws.title}")
        return ws

    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. Available sheets: {workbook.sheetnames}"
        )

    ws = workbook[sheet_name]
    vprint(f"    Using sheet: {ws.title}")
    return ws


def codes_match(input_code, ph_code):
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
    vprint("")
    vprint(f"    Opening pH Excel file:")
    vprint(f"      {ph_file}")

    wb = load_workbook(ph_file, data_only=True, read_only=True)
    ws = get_sheet(wb, PH_SHEET_NAME)

    input_marker_norm = normalize_marker(input_marker)
    input_code_norm = normalize_text(input_code)

    vprint(f"    Searching inside pH file from row {PH_FIRST_ROW} to {ws.max_row}")
    vprint(f"    Need to match:")
    vprint(f"      Input marker B = {input_marker_norm}")
    vprint(f"      Input code   C = {input_code_norm}")

    exact_matches = []
    code_only_matches = []

    for row in range(PH_FIRST_ROW, ws.max_row + 1):
        ph_marker = ws[f"{PH_MARKER_COL}{row}"].value
        ph_code = ws[f"{PH_CODE_COL}{row}"].value
        ph_value = ws[f"{PH_VALUE_COL}{row}"].value

        ph_marker_norm = normalize_marker(ph_marker)
        ph_code_norm = normalize_text(ph_code)

        vprint(
            f"      pH row {row}: "
            f"B={ph_marker_norm!r}, "
            f"C={ph_code_norm!r}, "
            f"H={ph_value!r}"
        )

        if not codes_match(input_code, ph_code):
            vprint("        Code does not match.")
            continue

        vprint("        Code matches.")

        code_only_matches.append(
            {
                "row": row,
                "marker": ph_marker_norm,
                "code": ph_code,
                "value": ph_value,
            }
        )

        if ph_marker_norm == input_marker_norm:
            vprint("        Marker also matches.")
            exact_matches.append(
                {
                    "row": row,
                    "marker": ph_marker_norm,
                    "code": ph_code,
                    "value": ph_value,
                }
            )
        else:
            vprint(
                f"        Marker does not match. "
                f"Input marker={input_marker_norm}, pH marker={ph_marker_norm}"
            )

    wb.close()

    if len(exact_matches) == 1:
        match = exact_matches[0]

        vprint("")
        vprint("    MATCH FOUND:")
        vprint(f"      File: {ph_file.name}")
        vprint(f"      pH row: {match['row']}")
        vprint(f"      Marker B: {match['marker']}")
        vprint(f"      Code C: {match['code']}")
        vprint(f"      pH value H: {match['value']}")

        return match["value"], match["row"], "exact marker + code match"

    if len(exact_matches) > 1:
        raise ValueError(
            f"Multiple exact matches in {ph_file.name} for code={input_code}, "
            f"marker={input_marker}. Rows: {[m['row'] for m in exact_matches]}"
        )

    if len(code_only_matches) == 1:
        match = code_only_matches[0]

        vprint("")
        vprint("    WARNING: code matched, but marker did not match.")
        vprint("    Using code-only fallback because code appears only once.")
        vprint(f"      File: {ph_file.name}")
        vprint(f"      pH row: {match['row']}")
        vprint(f"      Marker B: {match['marker']}")
        vprint(f"      Code C: {match['code']}")
        vprint(f"      pH value H: {match['value']}")

        return match["value"], match["row"], "code-only fallback"

    if len(code_only_matches) > 1:
        raise ValueError(
            f"Code found multiple times in {ph_file.name}, but marker did not match. "
            f"Input code={input_code}, input marker={input_marker}. "
            f"Candidate rows: {[(m['row'], m['marker']) for m in code_only_matches]}"
        )

    vprint("")
    vprint(f"    No match found in file: {ph_file.name}")

    return None, None, "not found"


def iter_input_rows(ws):
    row = INPUT_FIRST_ROW

    while row <= ws.max_row:
        block_start = row
        block_end = row + BLOCK_SIZE - 1

        vprint("")
        vprint("------------------------------------------------------------")
        vprint(f"Input block: rows {block_start} to {block_end}")
        vprint("------------------------------------------------------------")

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

    print("")
    print("============================================================")
    print("STARTING pH COPY PROCESS")
    print("============================================================")
    print(f"Input XLSM:")
    print(f"  {input_path}")
    print(f"pH folder:")
    print(f"  {ph_folder}")
    print(f"Output XLSM:")
    print(f"  {output_path}")
    print("============================================================")

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if not ph_folder.exists():
        raise FileNotFoundError(f"pH folder not found: {ph_folder}")

    print("")
    print("Copying original XLSM to output file...")
    shutil.copy2(input_path, output_path)
    print("Copy created successfully.")
    print(f"  {output_path}")

    print("")
    print("Opening input workbook for reading values...")
    wb_values = load_workbook(input_path, data_only=True, keep_vba=True)
    ws_values = get_sheet(wb_values, INPUT_SHEET_NAME)
    print(f"Input sheet used: {ws_values.title}")
    print(f"Input max row: {ws_values.max_row}")

    print("")
    print("Opening copied workbook for writing pH values...")
    wb_out = load_workbook(output_path, keep_vba=True)
    ws_out = get_sheet(wb_out, INPUT_SHEET_NAME)
    print(f"Output sheet used: {ws_out.title}")

    log = []

    for row in iter_input_rows(ws_values):
        print("")
        print("============================================================")
        print(f"READING INPUT ROW {row}")
        print("============================================================")

        input_marker = ws_values[f"{INPUT_MARKER_COL}{row}"].value
        input_code = ws_values[f"{INPUT_CODE_COL}{row}"].value

        marker_text = normalize_marker(input_marker)
        code_text = normalize_text(input_code)

        print(f"Input cell {INPUT_MARKER_COL}{row}: {input_marker!r} -> normalized: {marker_text!r}")
        print(f"Input cell {INPUT_CODE_COL}{row}: {input_code!r} -> normalized: {code_text!r}")

        if code_text == "":
            print("This row has no code in column C. Skipping.")
            continue

        su_number = extract_su_number(code_text)

        if su_number is None:
            print("Could not extract SU number. Skipping this row.")

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

        print(f"Extracted SU number: SU{su_number}")

        candidates = find_candidate_ph_files(ph_folder, su_number)

        if not candidates:
            print("")
            print("ERROR:")
            print(f"  No pH file was found containing SU{su_number}")

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

        print("")
        print("Now searching inside candidate pH file(s)...")

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
            print("")
            print("ERROR:")
            print("  The correct pH Excel file was found by filename range,")
            print("  but the code/marker was not found inside the file.")
            print(f"  Input row: {row}")
            print(f"  Marker B: {marker_text}")
            print(f"  Code C: {code_text}")
            print("  Candidate files:")
            for p in candidates:
                print(f"    {p.name}")

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

        print("")
        print("WRITING VALUE TO OUTPUT XLSM")
        print(f"  pH value from: {found_file.name}")
        print(f"  pH source row: {found_row}")
        print(f"  Source pH cell: {PH_VALUE_COL}{found_row}")
        print(f"  Destination cell: {INPUT_OUTPUT_COL}{row}")
        print(f"  Value written: {found_value}")

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

    print("")
    print("Saving output workbook...")
    wb_out.save(output_path)

    wb_values.close()
    wb_out.close()

    print("")
    print("============================================================")
    print("FINISHED")
    print("============================================================")
    print(f"Output file created:")
    print(f"  {output_path}")

    print("")
    print("FINAL SUMMARY")
    print("============================================================")

    for item in log:
        print(
            f"Input row {item['input_row']:>4} | "
            f"B={item['marker']:<5} | "
            f"C={item['code']:<25} | "
            f"pH={str(item['ph_value']):<12} | "
            f"pH row={str(item['ph_row']):<5} | "
            f"{item['status']} | "
            f"{item['file']}"
        )


if __name__ == "__main__":
    main()
