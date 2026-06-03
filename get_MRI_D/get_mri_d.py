from pathlib import Path
import re
import io
import csv
import sqlite3
import argparse
from datetime import datetime
from openpyxl import load_workbook


try:
    import msoffcrypto
except ImportError:
    msoffcrypto = None


# ============================================================
# NORMALIZATION
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip()
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    text = text.lstrip("'").strip()

    return text


def normalize_code(value):
    return normalize_text(value).upper()


def parse_number(value):
    text = normalize_text(value)

    if text == "":
        return None

    text = text.replace(",", ".")

    if set(text) <= {"-"}:
        return None

    try:
        return float(text)
    except Exception:
        return None


# ============================================================
# FILE DETECTION
# ============================================================

def detect_version(path):
    name = Path(path).name.upper()

    if "VER.02" in name or "VER 02" in name or "VER02" in name:
        return "Ver.02"

    if "VER.04" in name or "VER 04" in name or "VER04" in name:
        return "Ver.04"

    if "VER.05" in name or "VER 05" in name or "VER05" in name:
        return "Ver.05"

    return "Unknown"


def detect_method(path):
    name = Path(path).name.upper()

    if "OLSEN" in name:
        return "OLSEN"

    if "BRAY" in name or "KURTZ" in name:
        return "BRAY Y KURTZ"

    return "UNKNOWN"


def settings_for_version(version):
    """
    Corrected extraction rules.

    Ver.02:
      codigo in C
      tipo in D
      resultado in O
      date in E10
      password = 12

    Ver.04:
      codigo in C
      tipo in F
      resultado in Q
      date in T12
      no password

    Ver.05:
      codigo in C
      tipo in F
      resultado in Q
      date in T13
      no password
    """
    if version == "Ver.02":
        return {
            "codigo_col": "C",
            "tipo_col": "D",
            "resultado_col": "O",
            "date_cell": "E10",
            "password": "12",
        }

    if version == "Ver.04":
        return {
            "codigo_col": "C",
            "tipo_col": "F",
            "resultado_col": "Q",
            "date_cell": "T12",
            "password": None,
        }

    if version == "Ver.05":
        return {
            "codigo_col": "C",
            "tipo_col": "F",
            "resultado_col": "Q",
            "date_cell": "T13",
            "password": None,
        }

    return {
        "codigo_col": "C",
        "tipo_col": "F",
        "resultado_col": "Q",
        "date_cell": "T13",
        "password": None,
    }


def extract_date_from_filename(path):
    """
    Extract date like:
      13-05-26
      07-05-26
      31-03-26

    Returns yyyy-mm-dd if possible.
    """
    name = Path(path).name

    match = re.search(r"(\d{1,2})-(\d{1,2})-(\d{2,4})", name)

    if not match:
        return ""

    day = int(match.group(1))
    month = int(match.group(2))
    year = int(match.group(3))

    if year < 100:
        year += 2000

    try:
        return datetime(year, month, day).strftime("%Y-%m-%d")
    except Exception:
        return ""


def normalize_date(value, fallback=""):
    if value is None:
        return fallback

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    text = normalize_text(value)

    if text == "":
        return fallback

    for fmt in (
        "%m/%d/%y",
        "%m/%d/%Y",
        "%d/%m/%y",
        "%d/%m/%Y",
        "%d-%m-%y",
        "%d-%m-%Y",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass

    return text


# ============================================================
# WORKBOOK OPENING
# ============================================================

def decrypt_excel_with_password(excel_file, password):
    if msoffcrypto is None:
        raise RuntimeError(
            "msoffcrypto-tool is not installed. Install it with:\n"
            "  pip install msoffcrypto-tool"
        )

    decrypted = io.BytesIO()

    with open(excel_file, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)

    decrypted.seek(0)
    return decrypted


def open_workbook_for_reading(excel_file, password=None, verbose=True):
    excel_file = Path(excel_file)
    keep_vba = excel_file.suffix.lower() == ".xlsm"

    try:
        if verbose:
            print("  Trying openpyxl directly...")

        wb = load_workbook(
            excel_file,
            data_only=True,
            keep_vba=keep_vba,
            read_only=False,
        )

        if verbose:
            print("  Opened directly.")

        return wb

    except Exception as direct_error:
        if verbose:
            print(f"  Direct open failed: {direct_error}")

        if password is None:
            raise

    if password is not None:
        if verbose:
            print(f"  Trying password decryption with password={password!r}...")

        decrypted = decrypt_excel_with_password(excel_file, password)

        wb = load_workbook(
            decrypted,
            data_only=True,
            keep_vba=keep_vba,
            read_only=False,
        )

        if verbose:
            print("  Opened after password decryption.")

        return wb

    raise RuntimeError(f"Could not open workbook: {excel_file}")


# ============================================================
# SHEET SELECTION
# ============================================================

def get_sheet(wb, preferred_sheet="P_DIS"):
    if preferred_sheet in wb.sheetnames:
        return wb[preferred_sheet]

    return wb[wb.sheetnames[0]]


# ============================================================
# SQLITE DATABASE
# ============================================================

def create_database(sqlite_path):
    sqlite_path = Path(sqlite_path)
    sqlite_path.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(sqlite_path)
    cur = conn.cursor()

    cur.execute("DROP TABLE IF EXISTS colorimetric_results")

    cur.execute(
        """
        CREATE TABLE colorimetric_results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,

            codigo_muestra TEXT,
            tipo TEXT,
            categoria TEXT,
            resultado REAL,

            metodo TEXT,
            version TEXT,
            fecha TEXT,

            source_file TEXT,
            source_sheet TEXT,
            source_row INTEGER,

            codigo_col TEXT,
            tipo_col TEXT,
            resultado_col TEXT,

            duplicated_from_source_row INTEGER
        )
        """
    )

    cur.execute("CREATE INDEX idx_color_code ON colorimetric_results (codigo_muestra)")
    cur.execute("CREATE INDEX idx_color_tipo ON colorimetric_results (tipo)")
    cur.execute("CREATE INDEX idx_color_categoria ON colorimetric_results (categoria)")
    cur.execute("CREATE INDEX idx_color_metodo ON colorimetric_results (metodo)")
    cur.execute("CREATE INDEX idx_color_version ON colorimetric_results (version)")
    cur.execute("CREATE INDEX idx_color_fecha ON colorimetric_results (fecha)")
    cur.execute("CREATE INDEX idx_color_source ON colorimetric_results (source_file, source_row)")

    conn.commit()
    return conn


def insert_record(conn, record):
    cur = conn.cursor()

    cur.execute(
        """
        INSERT INTO colorimetric_results (
            codigo_muestra,
            tipo,
            categoria,
            resultado,

            metodo,
            version,
            fecha,

            source_file,
            source_sheet,
            source_row,

            codigo_col,
            tipo_col,
            resultado_col,

            duplicated_from_source_row
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record["codigo_muestra"],
            record["tipo"],
            record["categoria"],
            record["resultado"],

            record["metodo"],
            record["version"],
            record["fecha"],

            record["source_file"],
            record["source_sheet"],
            record["source_row"],

            record["codigo_col"],
            record["tipo_col"],
            record["resultado_col"],

            record["duplicated_from_source_row"],
        ),
    )


# ============================================================
# RECORD MAKER
# ============================================================

def make_record_from_row(
    ws,
    row,
    categoria,
    codigo_col,
    tipo_col,
    resultado_col,
    metodo,
    version,
    fecha,
    source_file,
    duplicated_from_source_row=None,
):
    codigo = normalize_code(ws[f"{codigo_col}{row}"].value)
    tipo = normalize_text(ws[f"{tipo_col}{row}"].value).upper()
    raw_resultado = ws[f"{resultado_col}{row}"].value
    resultado = parse_number(raw_resultado)

    if codigo == "":
        return None, "empty codigo"

    if tipo == "":
        return None, "empty tipo"

    if resultado is None:
        return None, f"invalid resultado: {raw_resultado!r}"

    record = {
        "codigo_muestra": codigo,
        "tipo": tipo,
        "categoria": categoria,
        "resultado": resultado,

        "metodo": metodo,
        "version": version,
        "fecha": fecha,

        "source_file": source_file,
        "source_sheet": ws.title,
        "source_row": row,

        "codigo_col": codigo_col,
        "tipo_col": tipo_col,
        "resultado_col": resultado_col,

        "duplicated_from_source_row": duplicated_from_source_row,
    }

    return record, ""


# ============================================================
# EXTRACTION
# ============================================================

def extract_records_from_workbook(
    excel_file,
    input_dir,
    sheet_name="P_DIS",
    verbose=True,
):
    excel_file = Path(excel_file)
    input_dir = Path(input_dir)

    version = detect_version(excel_file)
    metodo = detect_method(excel_file)
    settings = settings_for_version(version)

    codigo_col = settings["codigo_col"]
    tipo_col = settings["tipo_col"]
    resultado_col = settings["resultado_col"]
    date_cell = settings["date_cell"]
    password = settings["password"]

    records = []

    print("")
    print("============================================================")
    print("EXTRACTING FILE")
    print("============================================================")
    print(f"File:          {excel_file}")
    print(f"Version:       {version}")
    print(f"Method:        {metodo}")
    print(f"Sheet:         {sheet_name}")
    print(f"Codigo column: {codigo_col}")
    print(f"Tipo column:   {tipo_col}")
    print(f"Result column: {resultado_col}")
    print(f"Date cell:     {date_cell}")
    print(f"Password:      {password!r}")

    wb = open_workbook_for_reading(
        excel_file=excel_file,
        password=password,
        verbose=verbose,
    )

    ws = get_sheet(wb, preferred_sheet=sheet_name)

    fallback_date = extract_date_from_filename(excel_file)
    fecha = normalize_date(ws[date_cell].value, fallback=fallback_date)

    print(f"Date found:    {fecha}")
    print(f"Using sheet:   {ws.title}")
    print(f"Max row:       {ws.max_row}")

    try:
        source_file = str(excel_file.relative_to(input_dir))
    except Exception:
        source_file = str(excel_file)

    extracted_mri = 0
    extracted_d = 0
    extracted_d_below = 0
    skipped = 0

    for row in range(1, ws.max_row + 1):
        tipo = normalize_text(ws[f"{tipo_col}{row}"].value).upper()

        if tipo == "MRI":
            record, reason = make_record_from_row(
                ws=ws,
                row=row,
                categoria="MRI",
                codigo_col=codigo_col,
                tipo_col=tipo_col,
                resultado_col=resultado_col,
                metodo=metodo,
                version=version,
                fecha=fecha,
                source_file=source_file,
                duplicated_from_source_row=None,
            )

            if record is None:
                skipped += 1
                if verbose:
                    print(f"  SKIP MRI row={row}: {reason}")
            else:
                records.append(record)
                extracted_mri += 1
                if verbose:
                    print(
                        f"  FOUND MRI: row={row} "
                        f"codigo={record['codigo_muestra']} "
                        f"resultado={record['resultado']}"
                    )

        elif tipo == "D":
            record, reason = make_record_from_row(
                ws=ws,
                row=row,
                categoria="D",
                codigo_col=codigo_col,
                tipo_col=tipo_col,
                resultado_col=resultado_col,
                metodo=metodo,
                version=version,
                fecha=fecha,
                source_file=source_file,
                duplicated_from_source_row=None,
            )

            if record is None:
                skipped += 1
                if verbose:
                    print(f"  SKIP D row={row}: {reason}")
            else:
                records.append(record)
                extracted_d += 1
                if verbose:
                    print(
                        f"  FOUND D: row={row} "
                        f"codigo={record['codigo_muestra']} "
                        f"resultado={record['resultado']}"
                    )

            # Also extract row immediately below D
            below_row = row + 1

            if below_row <= ws.max_row:
                below_record, below_reason = make_record_from_row(
                    ws=ws,
                    row=below_row,
                    categoria="D_BELOW",
                    codigo_col=codigo_col,
                    tipo_col=tipo_col,
                    resultado_col=resultado_col,
                    metodo=metodo,
                    version=version,
                    fecha=fecha,
                    source_file=source_file,
                    duplicated_from_source_row=row,
                )

                if below_record is None:
                    skipped += 1
                    if verbose:
                        print(
                            f"  SKIP D_BELOW row={below_row}, "
                            f"from D row={row}: {below_reason}"
                        )
                else:
                    records.append(below_record)
                    extracted_d_below += 1
                    if verbose:
                        print(
                            f"  FOUND D_BELOW: row={below_row} "
                            f"codigo={below_record['codigo_muestra']} "
                            f"tipo={below_record['tipo']} "
                            f"resultado={below_record['resultado']} "
                            f"duplicated_from_row={row}"
                        )

    wb.close()

    print("")
    print("Finished file:")
    print(f"  Extracted MRI:      {extracted_mri}")
    print(f"  Extracted D:        {extracted_d}")
    print(f"  Extracted D_BELOW:  {extracted_d_below}")
    print(f"  Skipped:            {skipped}")
    print(f"  Total records:      {len(records)}")

    return records


# ============================================================
# FILE DISCOVERY
# ============================================================

def get_excel_files(input_dir, min_number=None, max_number=None):
    input_dir = Path(input_dir)

    files = sorted(
        list(input_dir.glob("*.xlsx")) +
        list(input_dir.glob("*.xlsm"))
    )

    selected = []

    for path in files:
        if path.name.startswith("~"):
            print(f"Skipping temporary/hidden file: {path.name}")
            continue

        version = detect_version(path)

        if version not in ("Ver.02", "Ver.04", "Ver.05"):
            continue

        if min_number is not None or max_number is not None:
            match = re.match(r"^(\d+)", path.name.strip())

            if not match:
                continue

            number = int(match.group(1))

            if min_number is not None and number < min_number:
                continue

            if max_number is not None and number > max_number:
                continue

        selected.append(path)

    print("")
    print("============================================================")
    print("INPUT FILES")
    print("============================================================")
    print(f"Input dir:      {input_dir}")
    print(f"Selected files: {len(selected)}")

    for path in selected:
        print(f"  - {path.name}")

    return selected


# ============================================================
# CSV EXPORT
# ============================================================

def export_csv(records, csv_path):
    csv_path = Path(csv_path)
    csv_path.parent.mkdir(parents=True, exist_ok=True)

    fieldnames = [
        "codigo_muestra",
        "tipo",
        "categoria",
        "resultado",
        "metodo",
        "version",
        "fecha",
        "source_file",
        "source_sheet",
        "source_row",
        "codigo_col",
        "tipo_col",
        "resultado_col",
        "duplicated_from_source_row",
    ]

    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()

        for record in records:
            writer.writerow(record)


# ============================================================
# MAIN
# ============================================================

def build_colorimetric_database(
    input_dir,
    sqlite_path,
    csv_path=None,
    sheet_name="P_DIS",
    min_number=None,
    max_number=None,
    verbose=True,
):
    input_dir = Path(input_dir)
    sqlite_path = Path(sqlite_path)

    print("")
    print("============================================================")
    print("BUILD COLORIMETRIC MRI/D/D_BELOW DATABASE")
    print("============================================================")
    print(f"Input dir:  {input_dir}")
    print(f"SQLite DB:  {sqlite_path}")
    print(f"CSV output: {csv_path}")
    print("")
    print("Rules:")
    print("  Ver.02: codigo=C, tipo=D, resultado=O, fecha=E10, password=12")
    print("  Ver.04: codigo=C, tipo=F, resultado=Q, fecha=T12")
    print("  Ver.05: codigo=C, tipo=F, resultado=Q, fecha=T13")
    print("  Extract rows where tipo is MRI")
    print("  Extract rows where tipo is D")
    print("  Also extract the row immediately below every D as D_BELOW")
    print("  Method comes from filename: OLSEN or BRAY Y KURTZ")
    print("============================================================")

    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir}")

    excel_files = get_excel_files(
        input_dir=input_dir,
        min_number=min_number,
        max_number=max_number,
    )

    conn = create_database(sqlite_path)

    all_records = []

    total_files = len(excel_files)
    successful_files = 0
    failed_files = 0

    for index, excel_file in enumerate(excel_files, start=1):
        print("")
        print("############################################################")
        print(f"PROCESSING {index}/{total_files}")
        print("############################################################")

        try:
            records = extract_records_from_workbook(
                excel_file=excel_file,
                input_dir=input_dir,
                sheet_name=sheet_name,
                verbose=verbose,
            )

            for record in records:
                insert_record(conn, record)

            all_records.extend(records)
            successful_files += 1

        except Exception as e:
            failed_files += 1
            print("")
            print("ERROR PROCESSING FILE:")
            print(f"  {excel_file}")
            print(f"  {e}")

    conn.commit()
    conn.close()

    if csv_path is not None:
        export_csv(all_records, csv_path)

    print("")
    print("============================================================")
    print("DATABASE BUILD FINISHED")
    print("============================================================")
    print(f"Files selected:     {total_files}")
    print(f"Successful files:   {successful_files}")
    print(f"Failed files:       {failed_files}")
    print(f"Total records:      {len(all_records)}")
    print(f"SQLite created:     {sqlite_path}")

    if csv_path is not None:
        print(f"CSV created:        {csv_path}")


# ============================================================
# ARGUMENTS
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Extract MRI, D, and D_BELOW rows from F-82 colorimetric Excel files into SQLite."
    )

    parser.add_argument(
        "--input-dir",
        default="/Users/mascenci/Downloads/Clientes",
        help="Folder containing colorimetric Excel files."
    )

    parser.add_argument(
        "--sqlite",
        default="database_colorimetric/colorimetric_mri_d.sqlite",
        help="Output SQLite database path."
    )

    parser.add_argument(
        "--csv",
        default="database_colorimetric/colorimetric_mri_d.csv",
        help="Optional output CSV path."
    )

    parser.add_argument(
        "--sheet-name",
        default="P_DIS",
        help="Sheet name to read. Default: P_DIS."
    )

    parser.add_argument(
        "--min-number",
        type=int,
        default=None,
        help="Optional first file number to process."
    )

    parser.add_argument(
        "--max-number",
        type=int,
        default=None,
        help="Optional last file number to process."
    )

    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Reduce row-level logging."
    )

    return parser.parse_args()


def main():
    args = parse_args()

    build_colorimetric_database(
        input_dir=args.input_dir,
        sqlite_path=args.sqlite,
        csv_path=args.csv,
        sheet_name=args.sheet_name,
        min_number=args.min_number,
        max_number=args.max_number,
        verbose=not args.quiet,
    )


if __name__ == "__main__":
    main()
