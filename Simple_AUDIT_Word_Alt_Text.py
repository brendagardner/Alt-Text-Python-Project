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
# 👉 Set this to either:
#    • a single .docx file, OR
INPUT_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/1035 Reviews"
    # or e.g.
    # "/Users/.../Math 1010 Lecture Notes Module 3_withAltText.docx"
)

# Where Excel audit files will be saved
OUTPUT_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information"
)





def process_single_docx(docx_path: Path, output_folder: Path):
    """Original behavior: one Excel per DOCX."""
    print(f"\n📄 Processing single file: {docx_path.name}")

    wb = Workbook()
    ws = wb.active
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
    image_refs: List[Tuple[BytesIO, ExcelImage]] = []

    row = append_docx_to_sheet(docx_path, ws, row, image_refs)

    # Autofit columns & wrap text
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(
            (len(str(cell.value)) for cell in col_cells if cell.value), default=0
        )
        for cell in col_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
    ws.column_dimensions["A"].width = 30

    output_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_xlsx = output_folder / f"{docx_path.stem}_Audit_{timestamp}.xlsx"
    wb.save(output_xlsx)
    print(f"✅ Excel saved to: {output_xlsx}")


def process_folder_to_single_excel(folder_path: Path, output_folder: Path):
    """Folder mode: combine all DOCX files into one Excel workbook."""
    docx_files = sorted(folder_path.glob("*.docx"))
    if not docx_files:
        print(f"⚠️ No .docx files found in folder: {folder_path}")
        return

    print(f"\n📂 Folder mode: combining {len(docx_files)} DOCX files into one Excel report")

    wb = Workbook()
    ws = wb.active
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
    image_refs: List[Tuple[BytesIO, ExcelImage]] = []

    for docx_path in docx_files:
        row = append_docx_to_sheet(docx_path, ws, row, image_refs)

    # Autofit columns & wrap text
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(
            (len(str(cell.value)) for cell in col_cells if cell.value), default=0
        )
        for cell in col_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
    ws.column_dimensions["A"].width = 30

    output_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_xlsx = output_folder / f"AltTextAudit_{timestamp}.xlsx"
    wb.save(output_xlsx)
    print(f"✅ Combined Excel saved to: {output_xlsx}")


def process_path(input_path: Path, output_folder: Path):
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx":
            print(f"❌ Input file is not a .docx: {input_path}")
            return
        process_single_docx(input_path, output_folder)

    elif input_path.is_dir():
        process_folder_to_single_excel(input_path, output_folder)

    else:
        print(f"❌ INPUT_PATH does not exist: {input_path}")


if __name__ == "__main__":
    process_path(INPUT_PATH, OUTPUT_FOLDER)
