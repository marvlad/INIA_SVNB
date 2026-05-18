# split_ph_olsen_bray.py

import csv
from pathlib import Path


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_CSV = Path(
    r"G:\Mi unidad\LABSAF ILLPA\ph_database_Ver03.csv"
)

OLSEN_CSV = Path(
    r"G:\Mi unidad\LABSAF ILLPA\olsen.csv"
)

BRAY_CSV = Path(
    r"G:\Mi unidad\LABSAF ILLPA\bray.csv"
)

PH_LIMIT = 6.5


# ============================================================
# HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    return str(value).strip().replace(",", ".")


def parse_ph(value):
    text = normalize_text(value)

    if text == "":
        return None

    try:
        return float(text)
    except ValueError:
        return None


def find_column(fieldnames, possible_names):
    """
    Find a column even if the capitalization is different.

    Example:
        ph
        PH
        PH_VALUE
    """
    normalized = {
        name.strip().lower(): name
        for name in fieldnames
    }

    for name in possible_names:
        key = name.strip().lower()
        if key in normalized:
            return normalized[key]

    return None


# ============================================================
# MAIN
# ============================================================

def main():
    print("")
    print("============================================================")
    print("SPLITTING pH CSV INTO OLSEN AND BRAY")
    print("============================================================")
    print(f"Input CSV:")
    print(f"  {INPUT_CSV}")
    print(f"Olsen output, pH > {PH_LIMIT}:")
    print(f"  {OLSEN_CSV}")
    print(f"Bray output, pH < {PH_LIMIT}:")
    print(f"  {BRAY_CSV}")
    print("============================================================")

    if not INPUT_CSV.exists():
        raise FileNotFoundError(f"Input CSV not found: {INPUT_CSV}")

    olsen_rows = []
    bray_rows = []
    skipped_rows = []

    with open(INPUT_CSV, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)

        if reader.fieldnames is None:
            raise ValueError("Input CSV has no header row.")

        print("")
        print("Detected columns:")
        for col in reader.fieldnames:
            print(f"  - {col}")

        code_col = find_column(
            reader.fieldnames,
            ["code", "SU_CODE", "su_code", "codigo", "código"]
        )

        ph_col = find_column(
            reader.fieldnames,
            ["ph", "PH", "PH_VALUE", "ph_value"]
        )

        if code_col is None:
            raise ValueError(
                "Could not find the SU code column. "
                "Expected one of: code, SU_CODE, su_code, codigo."
            )

        if ph_col is None:
            raise ValueError(
                "Could not find the pH column. "
                "Expected one of: ph, PH, PH_VALUE, ph_value."
            )

        print("")
        print(f"Using code column: {code_col}")
        print(f"Using pH column:   {ph_col}")

        for row_number, row in enumerate(reader, start=2):
            code = normalize_text(row.get(code_col, ""))
            ph_raw = row.get(ph_col, "")
            ph = parse_ph(ph_raw)

            print("")
            print(f"Row {row_number}:")
            print(f"  code raw = {row.get(code_col, '')!r} -> {code!r}")
            print(f"  pH raw   = {ph_raw!r} -> {ph}")

            if code == "":
                print("  SKIPPED: empty SU code")
                skipped_rows.append((row_number, code, ph_raw, "empty code"))
                continue

            if ph is None:
                print("  SKIPPED: pH is not numeric")
                skipped_rows.append((row_number, code, ph_raw, "invalid pH"))
                continue

            output_row = {
                "code": code,
                "ph": ph,
            }

            if ph > PH_LIMIT:
                print(f"  SELECTED FOR OLSEN: pH {ph} > {PH_LIMIT}")
                olsen_rows.append(output_row)

            elif ph < PH_LIMIT:
                print(f"  SELECTED FOR BRAY: pH {ph} < {PH_LIMIT}")
                bray_rows.append(output_row)

            else:
                print(f"  SKIPPED: pH is exactly {PH_LIMIT}")
                skipped_rows.append((row_number, code, ph_raw, "pH exactly 6.5"))

    # ------------------------------------------------------------
    # Write Olsen CSV
    # ------------------------------------------------------------
    with open(OLSEN_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["code", "ph"])
        writer.writeheader()
        writer.writerows(olsen_rows)

    # ------------------------------------------------------------
    # Write Bray CSV
    # ------------------------------------------------------------
    with open(BRAY_CSV, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["code", "ph"])
        writer.writeheader()
        writer.writerows(bray_rows)

    print("")
    print("============================================================")
    print("FINISHED")
    print("============================================================")
    print(f"Olsen rows, pH > {PH_LIMIT}: {len(olsen_rows)}")
    print(f"Bray rows, pH < {PH_LIMIT}:  {len(bray_rows)}")
    print(f"Skipped rows:             {len(skipped_rows)}")
    print("")
    print("Created files:")
    print(f"  {OLSEN_CSV}")
    print(f"  {BRAY_CSV}")


if __name__ == "__main__":
    main()
