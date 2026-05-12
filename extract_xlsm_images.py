# extract_xlsm_images.py

from pathlib import Path
from zipfile import ZipFile
import shutil
import xml.etree.ElementTree as ET


INPUT_XLSM = "template/F-82 Reporte de Resultados Colorimetricos Ver.05 Bray.xlsm"
OUTPUT_DIR = "extracted_images_from_xlsm"


def safe_name(name):
    return name.replace("/", "_").replace("\\", "_")


def extract_all_excel_images(input_xlsm, output_dir):
    """
    Extracts all raw images from an Excel .xlsm/.xlsx file.

    Excel files are zip archives. Images are usually stored in:
        xl/media/
    """
    input_xlsm = Path(input_xlsm)
    output_dir = Path(output_dir)

    if not input_xlsm.exists():
        raise FileNotFoundError(f"File not found: {input_xlsm}")

    output_dir.mkdir(parents=True, exist_ok=True)

    media_dir = output_dir / "media"
    media_dir.mkdir(parents=True, exist_ok=True)

    extracted = []

    with ZipFile(input_xlsm, "r") as z:
        media_files = [
            name for name in z.namelist()
            if name.startswith("xl/media/")
        ]

        if not media_files:
            print("No images found in xl/media/")
            return []

        print(f"Found {len(media_files)} image file(s) in workbook.")

        for media_file in media_files:
            original_name = Path(media_file).name
            output_path = media_dir / original_name

            with z.open(media_file) as source, open(output_path, "wb") as target:
                shutil.copyfileobj(source, target)

            extracted.append(output_path)

            print(f"Extracted: {output_path}")

    return extracted


def extract_sheet_mapping(input_xlsm, output_dir):
    """
    Tries to create a simple manifest showing which images belong to which sheet.

    This is not perfect for every Excel file, but helps identify:
    - logo image
    - equation image
    - chart/image used in P_DIS
    """
    input_xlsm = Path(input_xlsm)
    output_dir = Path(output_dir)

    manifest_path = output_dir / "image_manifest.txt"

    with ZipFile(input_xlsm, "r") as z:
        all_files = z.namelist()

        drawing_files = [
            name for name in all_files
            if name.startswith("xl/drawings/drawing")
            and name.endswith(".xml")
        ]

        rel_files = [
            name for name in all_files
            if name.startswith("xl/drawings/_rels/")
            and name.endswith(".rels")
        ]

        with open(manifest_path, "w", encoding="utf-8") as f:
            f.write("Excel image manifest\n")
            f.write("====================\n\n")

            f.write("Raw images found in xl/media:\n")
            for name in all_files:
                if name.startswith("xl/media/"):
                    f.write(f"  {name}\n")

            f.write("\nDrawing files:\n")
            for drawing_file in drawing_files:
                f.write(f"\n{drawing_file}\n")
                f.write("-" * len(drawing_file) + "\n")

                try:
                    xml_text = z.read(drawing_file).decode("utf-8", errors="ignore")
                    f.write(xml_text[:5000])
                    f.write("\n")
                except Exception as e:
                    f.write(f"Could not read drawing XML: {e}\n")

            f.write("\nRelationship files:\n")
            for rel_file in rel_files:
                f.write(f"\n{rel_file}\n")
                f.write("-" * len(rel_file) + "\n")

                try:
                    xml_text = z.read(rel_file).decode("utf-8", errors="ignore")
                    f.write(xml_text)
                    f.write("\n")
                except Exception as e:
                    f.write(f"Could not read rel XML: {e}\n")

    print(f"\nManifest saved to: {manifest_path}")


def main():
    extract_all_excel_images(INPUT_XLSM, OUTPUT_DIR)
    extract_sheet_mapping(INPUT_XLSM, OUTPUT_DIR)

    print()
    print("Done.")
    print(f"Open this folder and inspect the images:")
    print(f"  {Path(OUTPUT_DIR).resolve() / 'media'}")
    print()
    print("After you identify the files, rename/copy them like:")
    print("  extracted_images/logo.png")
    print("  extracted_images/equation.png")


if __name__ == "__main__":
    main()