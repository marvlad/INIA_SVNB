from pathlib import Path
import sqlite3
import argparse


# ============================================================
# HELPERS
# ============================================================

def normalize_tipo(value):
    text = str(value).strip().upper()

    if text not in ("MRI", "D"):
        raise ValueError("tipo must be MRI or D")

    return text


def normalize_month(value):
    """
    Accept:
      5
      05
      2026-05

    Returns:
      month number as int, or YYYY-MM string.
    """
    text = str(value).strip()

    if text == "":
        raise ValueError("month cannot be empty")

    if "-" in text:
        parts = text.split("-")

        if len(parts) != 2:
            raise ValueError("month must be like 05 or 2026-05")

        year = int(parts[0])
        month = int(parts[1])

        if month < 1 or month > 12:
            raise ValueError("month must be between 1 and 12")

        return f"{year:04d}-{month:02d}"

    month = int(text)

    if month < 1 or month > 12:
        raise ValueError("month must be between 1 and 12")

    return month


# ============================================================
# QUERY DATABASE
# ============================================================

def query_colorimetric_db(
    sqlite_path,
    tipo,
    month,
    metodo=None,
    version=None,
    show_source=False,
):
    sqlite_path = Path(sqlite_path)

    if not sqlite_path.exists():
        raise FileNotFoundError(f"SQLite database not found: {sqlite_path}")

    tipo = normalize_tipo(tipo)
    month_filter = normalize_month(month)

    conn = sqlite3.connect(sqlite_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    where = ["tipo = ?"]
    params = [tipo]

    # Month filter
    if isinstance(month_filter, int):
        where.append("CAST(strftime('%m', fecha) AS INTEGER) = ?")
        params.append(month_filter)
    else:
        where.append("substr(fecha, 1, 7) = ?")
        params.append(month_filter)

    if metodo is not None:
        where.append("metodo = ?")
        params.append(metodo)

    if version is not None:
        where.append("version = ?")
        params.append(version)

    source_cols = """
        source_file,
        source_sheet,
        source_row
    """

    query = f"""
        SELECT
            codigo_muestra,
            resultado,
            tipo,
            metodo,
            version,
            fecha,
            {source_cols}
        FROM colorimetric_results
        WHERE {" AND ".join(where)}
        ORDER BY fecha, metodo, codigo_muestra
    """

    cur.execute(query, params)
    rows = cur.fetchall()

    conn.close()

    print("")
    print("============================================================")
    print("COLORIMETRIC QUERY")
    print("============================================================")
    print(f"SQLite: {sqlite_path}")
    print(f"Tipo:   {tipo}")
    print(f"Month:  {month}")
    if metodo:
        print(f"Method: {metodo}")
    if version:
        print(f"Version:{version}")
    print(f"Rows:   {len(rows)}")
    print("============================================================")
    print("")

    if show_source:
        print("codigo_muestra\tresultado\tfecha\tmetodo\tversion\tsource_file\tsource_sheet\tsource_row")
        for row in rows:
            print(
                f"{row['codigo_muestra']}\t"
                f"{row['resultado']}\t"
                f"{row['fecha']}\t"
                f"{row['metodo']}\t"
                f"{row['version']}\t"
                f"{row['source_file']}\t"
                f"{row['source_sheet']}\t"
                f"{row['source_row']}"
            )
    else:
        print("codigo_muestra\tresultado")
        for row in rows:
            print(f"{row['codigo_muestra']}\t{row['resultado']}")

    return rows


# ============================================================
# ARGUMENTS
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Query MRI or D colorimetric results by month."
    )

    parser.add_argument(
        "--sqlite",
        default="database_colorimetric/colorimetric_mri_d.sqlite",
        help="Path to colorimetric SQLite database."
    )

    parser.add_argument(
        "--tipo",
        required=True,
        choices=["MRI", "D", "mri", "d"],
        help="Type to print: MRI or D."
    )

    parser.add_argument(
        "--month",
        required=True,
        help="Month to filter. Examples: 5, 05, 2026-05."
    )

    parser.add_argument(
        "--metodo",
        default=None,
        choices=["OLSEN", "BRAY Y KURTZ"],
        help="Optional method filter: OLSEN or BRAY Y KURTZ."
    )

    parser.add_argument(
        "--version",
        default=None,
        choices=["Ver.02", "Ver.04", "Ver.05"],
        help="Optional version filter."
    )

    parser.add_argument(
        "--show-source",
        action="store_true",
        help="Also print source file, sheet, and row."
    )

    return parser.parse_args()


def main():
    args = parse_args()

    query_colorimetric_db(
        sqlite_path=args.sqlite,
        tipo=args.tipo,
        month=args.month,
        metodo=args.metodo,
        version=args.version,
        show_source=args.show_source,
    )


if __name__ == "__main__":
    main()
