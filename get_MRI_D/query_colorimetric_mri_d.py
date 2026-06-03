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


def build_month_where(month_filter, params):
    """
    Build SQL month filter.

    fecha is stored as text, usually yyyy-mm-dd.
    """
    if isinstance(month_filter, int):
        params.append(month_filter)
        return "CAST(strftime('%m', fecha) AS INTEGER) = ?"

    params.append(month_filter)
    return "substr(fecha, 1, 7) = ?"


# ============================================================
# QUERY MRI
# ============================================================

def query_mri(
    sqlite_path,
    month,
    metodo=None,
    version=None,
    show_source=False,
):
    month_filter = normalize_month(month)

    conn = sqlite3.connect(sqlite_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    where = ["categoria = 'MRI'"]
    params = []

    where.append(build_month_where(month_filter, params))

    if metodo is not None:
        where.append("metodo = ?")
        params.append(metodo)

    if version is not None:
        where.append("version = ?")
        params.append(version)

    query = f"""
        SELECT
            codigo_muestra,
            resultado,
            tipo,
            categoria,
            metodo,
            version,
            fecha,
            source_file,
            source_sheet,
            source_row
        FROM colorimetric_results
        WHERE {" AND ".join(where)}
        ORDER BY fecha, metodo, codigo_muestra
    """

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    print("")
    print("============================================================")
    print("COLORIMETRIC QUERY: MRI")
    print("============================================================")
    print(f"SQLite: {sqlite_path}")
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
# QUERY D + D_BELOW
# ============================================================

def query_d_with_below(
    sqlite_path,
    month,
    metodo=None,
    version=None,
    show_source=False,
):
    """
    For each D row, find the corresponding D_BELOW row using:

        D_BELOW.duplicated_from_source_row = D.source_row
        same source_file
        same metodo
        same version
        same fecha

    Then print:
        codigo_muestra_D
        resultado_D
        codigo_muestra_D_BELOW
        resultado_D_BELOW
        difference = D - D_BELOW
    """
    month_filter = normalize_month(month)

    conn = sqlite3.connect(sqlite_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    where = ["d.categoria = 'D'"]
    params = []

    if isinstance(month_filter, int):
        where.append("CAST(strftime('%m', d.fecha) AS INTEGER) = ?")
        params.append(month_filter)
    else:
        where.append("substr(d.fecha, 1, 7) = ?")
        params.append(month_filter)

    if metodo is not None:
        where.append("d.metodo = ?")
        params.append(metodo)

    if version is not None:
        where.append("d.version = ?")
        params.append(version)

    query = f"""
        SELECT
            d.codigo_muestra AS d_codigo_muestra,
            d.resultado AS d_resultado,
            d.tipo AS d_tipo,
            d.fecha AS fecha,
            d.metodo AS metodo,
            d.version AS version,
            d.source_file AS source_file,
            d.source_sheet AS source_sheet,
            d.source_row AS d_source_row,

            b.codigo_muestra AS below_codigo_muestra,
            b.resultado AS below_resultado,
            b.tipo AS below_tipo,
            b.source_row AS below_source_row,

            (d.resultado - b.resultado) AS difference
        FROM colorimetric_results d
        LEFT JOIN colorimetric_results b
          ON b.categoria = 'D_BELOW'
         AND b.source_file = d.source_file
         AND b.metodo = d.metodo
         AND b.version = d.version
         AND b.fecha = d.fecha
         AND b.duplicated_from_source_row = d.source_row
        WHERE {" AND ".join(where)}
        ORDER BY d.fecha, d.metodo, d.source_file, d.source_row
    """

    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()

    print("")
    print("============================================================")
    print("COLORIMETRIC QUERY: D + D_BELOW")
    print("============================================================")
    print(f"SQLite: {sqlite_path}")
    print(f"Month:  {month}")
    if metodo:
        print(f"Method: {metodo}")
    if version:
        print(f"Version:{version}")
    print(f"Rows:   {len(rows)}")
    print("Difference:")
    print("  difference = D_resultado - D_BELOW_resultado")
    print("============================================================")
    print("")

    if show_source:
        print(
            "D_codigo\tD_resultado\t"
            "D_BELOW_codigo\tD_BELOW_resultado\t"
            "difference\tfecha\tmetodo\tversion\t"
            "source_file\tD_row\tD_BELOW_row"
        )

        for row in rows:
            print(
                f"{row['d_codigo_muestra']}\t"
                f"{row['d_resultado']}\t"
                f"{row['below_codigo_muestra'] or ''}\t"
                f"{row['below_resultado'] if row['below_resultado'] is not None else ''}\t"
                f"{row['difference'] if row['difference'] is not None else ''}\t"
                f"{row['fecha']}\t"
                f"{row['metodo']}\t"
                f"{row['version']}\t"
                f"{row['source_file']}\t"
                f"{row['d_source_row']}\t"
                f"{row['below_source_row'] if row['below_source_row'] is not None else ''}"
            )

    else:
        print("D_codigo\tD_resultado\tD_BELOW_codigo\tD_BELOW_resultado\tdifference")

        for row in rows:
            print(
                f"{row['d_codigo_muestra']}\t"
                f"{row['d_resultado']}\t"
                f"{row['below_codigo_muestra'] or ''}\t"
                f"{row['below_resultado'] if row['below_resultado'] is not None else ''}\t"
                f"{row['difference'] if row['difference'] is not None else ''}"
            )

    return rows


# ============================================================
# MAIN QUERY
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

    if tipo == "MRI":
        return query_mri(
            sqlite_path=sqlite_path,
            month=month,
            metodo=metodo,
            version=version,
            show_source=show_source,
        )

    if tipo == "D":
        return query_d_with_below(
            sqlite_path=sqlite_path,
            month=month,
            metodo=metodo,
            version=version,
            show_source=show_source,
        )


# ============================================================
# ARGUMENTS
# ============================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description="Query MRI or D colorimetric results by month. For D, also shows D_BELOW and difference."
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
        help="Type to print: MRI or D. If D, also prints D_BELOW and difference."
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
