# ====
# UNIFIED: MATCH BY PICTURE ID, APPLY ALT TEXT, AND BUILD QA REPORT
# Updates:
#  - Added Image Topic column (Excel column F -> report column E)
#  - Mark both "In Word?" and "In Excel?" as Yes when Excel alt text was found AND applied to Word
#  - Ensure alt text application is applied for every docPr (multiple images with same Picture ID)
#  - Capture Picture ID(s) and docPr title(s) for Word-only images (for Section)
#  - NEW: use python-docx to locate image paragraph indices and find the nearest preceding Heading 1 to populate Section
# Brenda Gardner — Updated Nov 5 2025 (revised per requested changes)
# ====

import zipfile
import os
import tempfile
import subprocess
from io import BytesIO
from pathlib import Path
from typing import Union, Tuple
from collections import defaultdict

import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree
from PIL import Image as PILImage

from docx import Document  # NEW: for heading-based Section extraction

# ====
# CONFIG — EDIT THESE
# ====
DOCX_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1010 Information/1010 Lecture Notes Chapters 1-5/Math 1010 Lecture Notes Module [...]"
)
EXCEL_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1010 Information/1010 Test Files/Mod 2 Push Alt Text.xlsx")

OUTPUT_DOCX  = DOCX_PATH.with_name(DOCX_PATH.stem + "_withAltText.docx")
REPORT_PATH  = DOCX_PATH.with_name(DOCX_PATH.stem + "_AltText_Report.xlsx")

# Excel column mapping (letters).
PICTURE_ID_COLUMN = column_index_from_string("C")  # Excel Picture ID
ALT_TEXT_COLUMN   = column_index_from_string("T")  # Excel Alt Text ("P" or "T" per your sheet)
FILE_NAME_COLUMN  = column_index_from_string("J")
SECTION_COLUMN    = column_index_from_string("K")
PAGE_NUMBER_COLUMN= column_index_from_string("L")
NOTES_COLUMN      = column_index_from_string("V")
TOPIC_COLUMN      = column_index_from_string("F")  # NEW: Image Topic in Excel column F

# Thumbnails
LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"  # adjust if needed
IMAGE_WIDTH = 150  # px

# ====
# HELPERS
# ====
def _norm_pid(val) -> str:
    """Normalize Picture IDs so 42 == 42.0 == ' 42 '."""
    if val is None or str(val).strip() == "":
        return ""
    s = str(val).strip()
    try:
        return str(int(float(s)))
    except ValueError:
        return s

def safe_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()

def load_image_bytes(docx_path: Path, target: str) -> Union[bytes, None]:
    """Extract raw image bytes from inside DOCX zip."""
    try:
        with zipfile.ZipFile(docx_path) as z:
            candidate = "word/" + target
            if candidate in z.namelist():
                return z.read(candidate)
    except Exception as e:
        print(f"⚠ Error reading {target}: {e}")
    return None

def convert_vector_to_png_bytes(input_bytes: bytes, suffix: str, temp_dir: Path) -> Union[bytes, None]:
    """Convert vector (EMF/WMF/SVG) to PNG via LibreOffice."""
    temp_in = temp_dir / f"vector_input{suffix}"
    temp_in.write_bytes(input_bytes)
    try:
        subprocess.run(
            [LIBREOFFICE_PATH, "--headless", "--convert-to", "png", "--outdir", str(temp_dir), str(temp_in)],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        png_path = temp_dir / (temp_in.stem + ".png")
        if png_path.exists():
            data = png_path.read_bytes()
            png_path.unlink(missing_ok=True)
            temp_in.unlink(missing_ok=True)
            return data
    except Exception as e:
        print(f"⚠ LibreOffice conversion failed for {suffix}: {e}")
    return None

def resize_png(png_bytes: bytes, width: int = IMAGE_WIDTH) -> Tuple[bytes, int]:
    pil = PILImage.open(BytesIO(png_bytes))
    aspect = pil.height / pil.width if pil.width else 1.0
    new_h = max(1, int(width * aspect))
    out = BytesIO()
    pil.resize((width, new_h), PILImage.Resampling.LANCZOS).save(out, format="PNG", dpi=(300, 300))
    out.seek(0)
    return out.getvalue(), new_h

def insert_image_from_docx(ws, docx_path: Path, target_filename: str, row: int, tmp_dir: Path):
    """Extract image from DOCX and embed in Excel."""
    try:
        img_bytes = load_image_bytes(docx_path, f"media/{target_filename}")
        if not img_bytes:
            ws.cell(row=row, column=1, value=f"(missing: {target_filename})")
            return

        suffix = Path(target_filename).suffix.lower()
        if suffix in [".emf", ".wmf", ".svg"]:
            converted = convert_vector_to_png_bytes(img_bytes, suffix, tmp_dir)
            if converted:
                img_bytes = converted
            else:
                ws.cell(row=row, column=1, value=f"(vector file: {target_filename} — not rendered)")
                ws.row_dimensions[row].height = 80
                return

        # Save the resized image as a temporary file (avoid BytesIO closure)
        resized_bytes, new_height = resize_png(img_bytes, width=IMAGE_WIDTH)
        tmp_png_path = tmp_dir / f"{Path(target_filename).stem}_temp.png"
        with open(tmp_png_path, "wb") as f:
            f.write(resized_bytes)

        xl_img = XLImage(str(tmp_png_path))  # use path, not BytesIO
        xl_img.width = IMAGE_WIDTH
        xl_img.height = new_height
        ws.add_image(xl_img, f"A{row}")
        ws.row_dimensions[row].height = new_height * 0.75

    except Exception as e:
        ws.cell(row=row, column=1, value=f"(error: {target_filename}: {e})")
        print(f"⚠ Could not insert {target_filename}: {e}")

def adjust_row_height(ws, row: int, img_height: int = None, min_height: int = 50):
    """Adjust row height based on image height or keep a consistent minimum."""
    if img_height and img_height > min_height:
        ws.row_dimensions[row].height = img_height * 0.75
    else:
        ws.row_dimensions[row].height = min_height

def solid_fill(argb: str) -> PatternFill:
    return PatternFill(patternType="solid", fgColor=argb)

# --- Functions borrowed/adjusted from WORD_AI_OUTPUT_FILE.py to locate image paragraph indices
def read_paragraph_texts(doc: Document):
    return [p.text for p in doc.paragraphs]

def extract_section_from_heading(paragraphs, p_idx: int) -> str:
    """Walk backwards from paragraph index p_idx and return nearest Heading 1 text (or empty)."""
    if p_idx is None or p_idx < 0 or p_idx >= len(paragraphs):
        return ""
    for i in range(p_idx, -1, -1):
        para = paragraphs[i]
        style_name = para.style.name if para.style else ""
        if style_name == "Heading 1":
            text = para.text.strip()
            # try to extract a section-like prefix like "1.2" etc.
            import re
            m = re.match(r"^([A-Za-z]*\s*\d+(\.\d+)*)(\s|$)", text)
            if m:
                return m.group(1)
            return text
    return ""

def locate_all_image_instances(doc: Document):
    """
    Return mapping: rid -> list of paragraph indices where that image instance occurs.
    This inspects runs in paragraphs, table cells, and tables recursively.
    """
    rid_to_paras = {}
    for i, p in enumerate(doc.paragraphs):
        for run in p.runs:
            blips = run._element.xpath(
                ".//*[local-name()='blip' and namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main']"
            )
            for blip in blips:
                rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if rid:
                    rid_to_paras.setdefault(rid, []).append(i)

    def inspect_cell(cell, row_idx_hint):
        for p in cell.paragraphs:
            for run in p.runs:
                blips = run._element.xpath(
                    ".//*[local-name()='blip' and namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main']"
                )
                for blip in blips:
                    rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    if rid:
                        rid_to_paras.setdefault(rid, []).append(row_idx_hint)
        for t in cell.tables:
            inspect_table(t, row_idx_hint)

    def inspect_table(table, row_idx_hint):
        for row in table.rows:
            for cell in row.cells:
                inspect_cell(cell, row_idx_hint)

    for tbl in doc.tables:
        parent_idx = -1
