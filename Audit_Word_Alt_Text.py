# Created Nov 24 to capture/audit the current Alt Text from DOCX files in preferred Excel format
#
# This script:
# Handles either a single DOCX or a folder of DOCX files and creates an Excel "Alt Text Audit" file
# One row per *visible* image (no orphans, no duplicates) - FIX THIS FOR DUPLICATES
# Populates columns A–AC with audit data; AI columns S–V and suggested columns W–AD are headers only for later use
# Page # uses build_rid_page_map (lastRenderedPageBreak) — same method as WORD_to_AI.py
# Decorative images now noted as "decorative" in Column D

# to Fix - have file also run .emf images for user to know and fix
# reuse page numbers and image topic from .xlsx for new audit file
# put into a single Excel file for all DOCX files in a folder
# also find .svg images and convert to .png

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
# 👉 Add a FOLDER or FILE to search for DOCX files
INPUT_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/0920 Information/920 2025 Lecture Notes/Module 6/BAIN_Module 6_Accessible_Lecture Notes_withAltText.docx"
)

# Where Excel audit files will be saved
OUTPUT_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/0920 Information/920 2025 Lecture Notes/Module 6"
)

GRADE_LEVEL    = "6-8"           # e.g. "12-14" (Column H)
COURSE_NAME    = "Foundational Math"  # e.g. "College Algebra" (Column I)
ZAMZAR_API_KEY = None              # set to os.getenv("ZAMZAR") if you want
                                   # high-quality EMF→PNG via Zamzar;
                                   # leave None to skip EMF thumbnails

# Set to True to combine all DOCX files in a folder into ONE Excel workbook.
# Set to False (default) to produce one Excel file per DOCX.
COMBINE_INTO_SINGLE_EXCEL = False

# =============================================================================
# Internal helpers
# =============================================================================

def _setup_workbook():
    """Create a new workbook with the master 30-column header row."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Alt Text Audit"
    ws.column_dimensions["A"].width = altutils._pixels_to_column_width(100)

    headers = [
        "Image Thumbnail",                               # A  (1)
        "Image Filename",                                # B  (2)
        "Picture ID",                                    # C  (3)
        "Image URL",                                     # D  (4) — blank for Word sources
        "Image Size",                                    # E  (5)
        "Image Topic",                                   # F  (6) — blank for audit
        "Image Context",                                 # G  (7)
        "Page #",                                        # H  (8)
        "Section/Module/Figure #",                       # I  (9)
        "Filename (Hyperlinked) or Webpage Name (URL)",  # J  (10)
        "Grade Level",                                   # K  (11)
        "Course Name",                                   # L  (12)
        "Original Source Author",                        # M  (13)
        "Copyright",                                     # N  (14) — blank
        "Date of Audit",                                 # O  (15) — auto-filled
        "Original Alt Text",                             # P  (16)
        "Original Alt Text Character Count",             # Q  (17) — =LEN(P)
        "Original Alt Text Notes",                       # R  (18) — blank
        "A.I. Short Alt Text",                           # S  (19) — blank for audit
        "A.I. Med Alt Text",                             # T  (20) — blank for audit
        "A.I. Long Alt Text",                            # U  (21) — blank for audit
        "A.I. Notes",                                    # V  (22) — blank for audit
        "New/Suggested Alt Text",                        # W  (23) — blank
        "New/Suggested Alt Text Count",                  # X  (24) — =LEN(W)
        "New/Suggested Alt Text Notes",                  # Y  (25) — blank
        "Status - Alt Text Completed?\n(Original/Suggested/N)",  # Z  (26) — blank
        "Reviewer Notes",                                # AA (27) — blank
        "In Word?",                                      # AB (28)
        "In Excel?",                                     # AC (29) — blank
        "Date of Completion",                            # AD (30) — blank
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(bold=True)

    return wb, ws


def _set_column_widths_and_wrap(ws):
    """Apply standard column widths and top-aligned text wrap to all cells."""
    ws.column_dimensions['B'].width = 30   # Image Filename
    ws.column_dimensions['C'].width = 18   # Picture ID
    ws.column_dimensions['D'].width = 30   # Image URL
    ws.column_dimensions['E'].width = 14   # Image Size
    ws.column_dimensions['F'].width = 25   # Image Topic
    ws.column_dimensions['G'].width = 40   # Image Context
    ws.column_dimensions['H'].width = 8    # Page #
    ws.column_dimensions['I'].width = 25   # Section/Module/Figure #
    ws.column_dimensions['J'].width = 35   # Filename (Hyperlinked) or Webpage Name
    ws.column_dimensions['K'].width = 12   # Grade Level
    ws.column_dimensions['L'].width = 20   # Course Name
    ws.column_dimensions['M'].width = 20   # Original Source Author
    ws.column_dimensions['N'].width = 20   # Copyright
    ws.column_dimensions['O'].width = 16   # Date of Audit
    ws.column_dimensions['P'].width = 50   # Original Alt Text
    ws.column_dimensions['Q'].width = 16   # Original Alt Text Character Count
    ws.column_dimensions['R'].width = 40   # Original Alt Text Notes
    ws.column_dimensions['S'].width = 50   # A.I. Short Alt Text
    ws.column_dimensions['T'].width = 50   # A.I. Med Alt Text
    ws.column_dimensions['U'].width = 50   # A.I. Long Alt Text
    ws.column_dimensions['V'].width = 40   # A.I. Notes
    ws.column_dimensions['W'].width = 50   # New/Suggested Alt Text
    ws.column_dimensions['X'].width = 18   # New/Suggested Alt Text Count
    ws.column_dimensions['Y'].width = 40   # New/Suggested Alt Text Notes
    ws.column_dimensions['Z'].width = 20   # Status - Alt Text Completed?
    ws.column_dimensions['AA'].width = 30  # Reviewer Notes
    ws.column_dimensions['AB'].width = 12  # In Word?
    ws.column_dimensions['AC'].width = 12  # In Excel?
    ws.column_dimensions['AD'].width = 18  # Date of Completion

    for col_cells in ws.columns:
        for cell in col_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")


def _append_docx_rows(docx_path: Path, ws, start_row: int, openpyxl_image_refs: list) -> int:
    """
    Process one DOCX file and append one row per visible image to ws.
    Returns the next available row number after all rows are written.
    """
    print(f"\n📄 Processing: {docx_path.name}")
    doc = Document(str(docx_path))
    paragraphs = doc.paragraphs
    doc_author = getattr(doc.core_properties, "author", "") or ""

    paragraph_texts = altutils.read_paragraph_texts(doc)
    alt_map      = altutils.parse_alt_text(docx_path)
    rels         = altutils.parse_relationships(docx_path)
    rid_para_map = altutils.locate_all_image_instances(doc)
    docpr_map    = altutils.extract_docpr_tags_with_rid(docx_path)
    rid_page_map      = altutils.build_rid_page_map(docx_path)
    rid_figure_map    = altutils.build_rid_figure_map(docx_path)
    rid_decorative    = altutils.build_rid_decorative_map(docx_path)

    print(f"  rels (images) found: {len(rels)}")
    print(f"  alt_map entries: {len(alt_map)}")
    print(f"  docpr entries: {len(docpr_map)}")

    base_name = docx_path.name
    today_str = datetime.datetime.now().strftime("%m/%d/%Y")

    row = start_row
    seen_images = set()  # deduplication scoped to this DOCX

    for rid, target in rels.items():
        # Skip rels with no docPr entry (orphan/unseen images)
        if rid not in docpr_map:
            continue

        para_indices = rid_para_map.get(rid) or []
        if not para_indices:
            continue

        # Only process each media target once per DOCX
        image_key = (docx_path.name, target)
        if image_key in seen_images:
            continue
        seen_images.add(image_key)

        img_bytes = altutils.load_image_bytes(docx_path, target)
        if not img_bytes:
            continue

        suffix = Path(target).suffix.lower()

        if suffix == ".wmf":
            print(f"  ⏭️  Skipping WMF file: {target}")
            continue

        if suffix in (".emf", ".svg"):
            converted = altutils.convert_vector_to_png_bytes(
                img_bytes, suffix, docx_path.parent, ZAMZAR_API_KEY
            )
            if not converted:
                print(f"  ⚠️  Could not convert {Path(target).name} to PNG — no thumbnail.")
            img_bytes = converted if converted else None

        if not img_bytes:
            continue

        p_idx = para_indices[0]
        existing_alt = alt_map.get(rid, "").strip()
        context = altutils.build_context_snippet(paragraph_texts, p_idx, window=1)

        thumb_bytes = None
        thumb_size = (0, 0)
        try:
            thumb_bytes, thumb_size = altutils.resize_png(img_bytes)
        except Exception:
            thumb_bytes = None

        # Column A (1): Image Thumbnail
        if thumb_bytes and thumb_size and thumb_size[0] > 0 and thumb_size[1] > 0:
            excel_img = altutils.insert_png_bytes_in_cell(
                ws, row=row, col=1,
                png_bytes=thumb_bytes,
                width_px=thumb_size[0], height_px=thumb_size[1],
                adjust_row_col=True, max_col_width_px=100,
            )
            if excel_img is not None:
                openpyxl_image_refs.append(excel_img)
        else:
            ws.cell(row=row, column=1, value="[No Thumbnail]")

        # Column B (2): Image Filename
        ws.cell(row=row, column=2, value=Path(target).name)

        # Column C (3): Picture ID
        ws.cell(row=row, column=3, value=docpr_map.get(rid, {}).get("picture_id", ""))

        # Column D (4): "decorative" if the image is marked as decorative in Word; blank otherwise
        if rid in rid_decorative:
            ws.cell(row=row, column=4, value="decorative")

        # Column E (5): Image Size
        if thumb_size and thumb_size[0] and thumb_size[1]:
            ws.cell(row=row, column=5, value=f"{thumb_size[0]} x {thumb_size[1]} px")
        else:
            ws.cell(row=row, column=5, value="")

        # Column F (6): Image Topic — blank for audit (no AI call)
        ws.cell(row=row, column=6, value="")

        # Column G (7): Image Context
        ws.cell(row=row, column=7, value=context)

        # Column H (8): Page #
        ws.cell(row=row, column=8, value=rid_page_map.get(rid, ""))

        # Column I (9): Section/Module/Figure #
        figure_label = rid_figure_map.get(rid, "")
        if not figure_label:
            figure_label = altutils.extract_section_from_heading(paragraphs, p_idx)
        ws.cell(row=row, column=9, value=figure_label)

        # Column J (10): Filename (Hyperlinked)
        j_cell = ws.cell(row=row, column=10, value=base_name)
        j_cell.hyperlink = str(docx_path)
        j_cell.style = "Hyperlink"

        # Column K (11): Grade Level
        ws.cell(row=row, column=11, value=GRADE_LEVEL)

        # Column L (12): Course Name
        ws.cell(row=row, column=12, value=COURSE_NAME)

        # Column M (13): Original Source Author
        ws.cell(row=row, column=13, value=doc_author)

        # Column N (14): Copyright — blank

        # Column O (15): Date of Audit
        ws.cell(row=row, column=15, value=today_str)

        # Column P (16): Original Alt Text
        ws.cell(row=row, column=16, value=existing_alt)

        # Column Q (17): Original Alt Text Character Count — formula =LEN(P)
        ws.cell(row=row, column=17, value=f"=LEN({get_column_letter(16)}{row})")

        # Columns R–V (18–22): Notes / A.I. columns — blank for audit

        # Column W (23): New/Suggested Alt Text — blank

        # Column X (24): New/Suggested Alt Text Count — formula =LEN(W)
        ws.cell(row=row, column=24, value=f"=LEN({get_column_letter(23)}{row})")

        # Columns Y–AA (25–27): Notes / Status / Reviewer Notes — blank

        # Column AB (28): In Word?
        ws.cell(row=row, column=28, value="Yes")

        # Columns AC–AD (29–30): In Excel? / Date of Completion — blank

        row += 1

    return row  # next available row for the next DOCX (or unused if single-file mode)


# =============================================================================
# Public entry points
# =============================================================================

def process_single_docx(docx_path: Path, output_folder: Path):
    """Process one DOCX and save as its own Excel file."""
    wb, ws = _setup_workbook()
    openpyxl_image_refs = []
    _append_docx_rows(docx_path, ws, start_row=2, openpyxl_image_refs=openpyxl_image_refs)
    _set_column_widths_and_wrap(ws)

    output_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_xlsx = output_folder / f"{docx_path.stem}_Audit_{timestamp}.xlsx"
    wb.save(output_xlsx)
    print(f"✅ Excel saved to: {output_xlsx}")


def process_folder_to_single_excel(folder_path: Path, output_folder: Path):
    """Process all DOCX files in a folder and combine into one Excel workbook."""
    docx_files = sorted(folder_path.glob("*.docx"))
    if not docx_files:
        print(f"⚠️ No .docx files found in folder: {folder_path}")
        return

    print(f"\n📂 Combined mode: merging {len(docx_files)} DOCX file(s) into one Excel workbook")
    wb, ws = _setup_workbook()
    openpyxl_image_refs = []
    current_row = 2

    for docx_file in docx_files:
        current_row = _append_docx_rows(
            docx_file, ws, start_row=current_row,
            openpyxl_image_refs=openpyxl_image_refs
        )

    _set_column_widths_and_wrap(ws)
    output_folder.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_xlsx = output_folder / f"AltTextAudit_combined_{timestamp}.xlsx"
    wb.save(output_xlsx)
    print(f"✅ Combined Excel ({len(docx_files)} files) saved to: {output_xlsx}")


def process_path(input_path: Path, output_folder: Path):
    """Handle either a single DOCX or a folder of DOCX files."""
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx":
            print(f"❌ Input file is not a .docx: {input_path}")
            return
        process_single_docx(input_path, output_folder)

    elif input_path.is_dir():
        if COMBINE_INTO_SINGLE_EXCEL:
            process_folder_to_single_excel(input_path, output_folder)
        else:
            docx_files = sorted(input_path.glob("*.docx"))
            if not docx_files:
                print(f"⚠️ No .docx files found in folder: {input_path}")
                return
            print(f"📂 Folder mode: found {len(docx_files)} .docx file(s) — one Excel per DOCX")
            for docx_file in docx_files:
                process_single_docx(docx_file, output_folder)

    else:
        print(f"❌ INPUT_PATH does not exist: {input_path}")


if __name__ == "__main__":
    process_path(INPUT_PATH, OUTPUT_FOLDER)
