# Created Nov 24 to capture/audit the current Alt Text from DOCX files in preferred Excel format
# Columns match output excel file from Alt Text to Word Report
# Updated Nov 25 to use alttext_utils.py and helper folder

"""
Single DOCX → Excel (A–V headers). No API calls.
Fills in:
A: Thumbnail, B: Image Filename, C: Picture ID (what is eventually matched), D: File Name/Source,
E: Image Topic (blank)
F: Section/Module (derived from heading)
G: Page Number (guess based on paragraph index)
H: In Word? (auto "Yes")
I: In Excel? (blank)
J: Current Alt Text
K: Alt Text Count (len)
L: Notes (blank)
"""

import datetime
from pathlib import Path
from io import BytesIO

from docx import Document
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


from helpers import alttext_utils as altutils

# =========================
# CONFIG — EDIT THESE
# =========================
# 👉 Add a FOLDER to search for DOCX files
INPUT_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1050 Information/1050 Testing"
)

# Where Excel audit files will be saved
OUTPUT_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1050 Information/1050 Testing"
)

# Processes specific to this script.  Others stored in alttext_utils.py

def process_single_docx(docx_path: Path, output_folder: Path):
    print(f"\n📄 Processing: {docx_path.name}")
    doc = Document(str(docx_path))
    paragraphs = doc.paragraphs

    alt_map = altutils.parse_alt_text(docx_path)
    rels = altutils.parse_relationships(docx_path)
    rid_para_map = altutils.locate_all_image_instances(doc)
    docpr_map = altutils.extract_docpr_tags_with_rid(docx_path)

    print(f"  rels (images) found: {len(rels)}")
    print(f"  alt_map entries: {len(alt_map)}")
    print(f"  docpr entries: {len(docpr_map)}")

    base_name = docx_path.name

    wb = Workbook()
    ws = wb.active
    DEFAULT_THUMB_WIDTH_PX = 100
    ws.column_dimensions["A"].width = altutils._pixels_to_column_width(DEFAULT_THUMB_WIDTH_PX)
    ws.title = "Alt Text Audit"

    headers = [
        "Thumbnail",
        "Image Filename",
        "Picture ID",
        "File Name/Source",
        "Image Topic",
        "Section/Module",
        "Page Number",
        "In Word?",
        "In Excel?",
        "Current Alt Text",
        "Alt Text Count",
        "Notes",
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(bold=True)

    row = 2
    openpyxl_image_refs = []

    for rid, target in rels.items():
        img_bytes = altutils.load_image_bytes(docx_path, target)
        if not img_bytes:
            continue

        suffix = Path(target).suffix.lower()

        # ✅ Try to convert both .emf and .svg to PNG for thumbnails
        if suffix in (".emf", ".svg"):
            converted = altutils.convert_vector_to_png_bytes(
                img_bytes, suffix, docx_path.parent
            )
            img_bytes = converted if converted else None

        thumb_bytes = None
        thumb_size = (0, 0)

        if img_bytes:
            try:
                thumb_bytes, thumb_size = altutils.resize_png(img_bytes)
            except Exception:
                thumb_bytes = None

        para_indices = rid_para_map.get(rid, [-1])
        existing_alt = alt_map.get(rid, "").strip()

        for p_idx in para_indices:
            # Column A: Thumbnail
            if thumb_bytes and thumb_size and thumb_size[0] > 0 and thumb_size[1] > 0:
                excel_img = altutils.insert_png_bytes_in_cell(
                    ws,
                    row=row,
                    col=1,  # Column A
                    png_bytes=thumb_bytes,
                    width_px=thumb_size[0],
                    height_px=thumb_size[1],
                    adjust_row_col=True,
                    max_col_width_px=100,
                )
                if excel_img is not None:
                    # Keep a reference so images don't disappear on save
                    openpyxl_image_refs.append(excel_img)
            else:
                ws.cell(row=row, column=1, value="[No Thumbnail]")

            # Column B: Image Filename
            ws.cell(row=row, column=2, value=Path(target).name)
            # Column C: Picture ID
            ws.cell(row=row, column=3, value=docpr_map.get(rid, {}).get("picture_id", ""))
            # Column D: File Name/Source
            ws.cell(row=row, column=4, value=base_name)
            # Column E: Image Topic (blank)
            ws.cell(row=row, column=5, value="")
            # Column F: Section/Module
            ws.cell(row=row, column=6, value=altutils.extract_section_from_heading(paragraphs, p_idx))
            # Column G: Page Number
            ws.cell(row=row, column=7, value=altutils.guess_page_number(p_idx))
            # Column H: In Word?
            ws.cell(row=row, column=8, value="Yes")
            # Column I: In Excel?
            ws.cell(row=row, column=9, value="")
            # Column J: Current Alt Text
            ws.cell(row=row, column=10, value=existing_alt)
            # Column K: Alt Text Count (length of J)
            ws.cell(
                row=row,
                column=11,
                value=f"=LEN({get_column_letter(10)}{row})",  # J column
            )
            # Column L: Notes
            ws.cell(row=row, column=12, value="")

            row += 1

    # Autofit columns & wrap text
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(
            (len(str(cell.value)) for cell in col_cells if cell.value), default=0
        )
        for cell in col_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
    ws.column_dimensions["A"].width = 50

    output_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_xlsx = output_folder / f"{docx_path.stem}_Audit_{timestamp}.xlsx"
    wb.save(output_xlsx)
    print(f"✅ Excel saved to: {output_xlsx}")


def process_path(input_path: Path, output_folder: Path):
    """Handle either a single DOCX or a folder of DOCX files."""
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx":
            print(f"❌ Input file is not a .docx: {input_path}")
            return
        process_single_docx(input_path, output_folder)

    elif input_path.is_dir():
        docx_files = sorted(input_path.glob("*.docx"))
        if not docx_files:
            print(f"⚠️ No .docx files found in folder: {input_path}")
            return

        print(f"📂 Folder mode: found {len(docx_files)} .docx files in {input_path}")
        for docx_file in docx_files:
            process_single_docx(docx_file, output_folder)

    else:
        print(f"❌ INPUT_PATH does not exist: {input_path}")


if __name__ == "__main__":
    process_path(INPUT_PATH, OUTPUT_FOLDER)
