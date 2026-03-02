# Created Nov 24 to capture/audit the current Alt Text from DOCX files in preferred Excel format
#
# This script:
# - Handles either a single DOCX or a folder of DOCX files
# - Creates an Excel "Alt Text Audit" file per DOCX
# - One row per *visible* image (no orphans, no duplicates)
# - Populates columns A–O; P–X are headers only for later use

# to Fix - have file also run .emf images for user to know and fix
# reuse page numbers and image topic from .xlsx for new audit file
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
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1050 Information/1050 Sections/Module 1"
)

# Where Excel audit files will be saved
OUTPUT_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1050 Information/1050 Sections/Module 1"
)

GRADE_LEVEL = "10-14"  # e.g. "12-14" (Column H)
COURSE_NAME = "College Algebra"  # e.g. "College Algebra" (Column I)

def process_single_docx(docx_path: Path, output_folder: Path):
    print(f"\n📄 Processing: {docx_path.name}")
    doc = Document(str(docx_path))
    paragraphs = doc.paragraphs
    doc_author = getattr(doc.core_properties, "author", "") or ""

    # Used for building context snippets (like AI script)
    paragraph_texts = altutils.read_paragraph_texts(doc)

    alt_map = altutils.parse_alt_text(docx_path)
    rels = altutils.parse_relationships(docx_path)
    rid_para_map = altutils.locate_all_image_instances(doc)
    docpr_map = altutils.extract_docpr_tags_with_rid(docx_path)

    print(f"  rels (images) found: {len(rels)}")
    print(f"  alt_map entries: {len(alt_map)}")
    print(f"  docpr entries: {len(docpr_map)}")

    base_name = docx_path.name
    today_str = datetime.datetime.now().strftime("%m-%d-%Y") # Today's date in MM-DD-YYYY format for Column U

    wb = Workbook()
    ws = wb.active
    DEFAULT_THUMB_WIDTH_PX = 100
    ws.column_dimensions["A"].width = altutils._pixels_to_column_width(DEFAULT_THUMB_WIDTH_PX)
    ws.title = "Alt Text Audit"

    # Master headers A–X (1–24)
    headers = [
        "Image Thumbnail",                               # A
        "Image Filename",                                # B
        "Picture ID",                                    # C
        "File Name",                                     # D
        "Image Size",                                    # E
        "Image Topic",                                   # F
        "Image Context",                                 # G
        "Grade Level",                                   # H
        "Course Name",                                   # I
        "File Name/Source (Hyperlink, if possible)",     # J
        "Section/Module",                                # K
        "Page #",                                        # L
        "Document Author, if known",                     # M
        "Original Alt Text",                             # N
        "Original Alt Text Character Count",             # O
        "Original Alt Text Notes",                       # P
        "New/Suggested Alt Text",                        # Q
        "New/Suggested Alt Text Count",                  # R
        "New/Suggested Alt Text Notes",                  # S
        "Alt Text Completed? (Original/Suggested/N)",    # T
        "Date of Audit",                                  # U
        "TBD",                                           # V
        "Approved by",                                   # W
        "Reviewer Notes",                                # X
    ]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = Font(bold=True)

    row = 2
    openpyxl_image_refs = []

    # ✅ Match AI script behavior: avoid orphan/duplicate images
    seen_images = set()

    for rid, target in rels.items():
        # 1️⃣ Skip rels that don't have a docPr entry (orphan/unseen images)
        if rid not in docpr_map:
            continue

        # 2️⃣ Get paragraph indices for this rid; skip if none
        para_indices = rid_para_map.get(rid) or []
        if not para_indices:
            continue

        # 3️⃣ Only process each media target once per DOCX
        image_key = (docx_path.name, target)
        if image_key in seen_images:
            continue
        seen_images.add(image_key)

        # Load image bytes
        img_bytes = altutils.load_image_bytes(docx_path, target)
        if not img_bytes:
            continue

        suffix = Path(target).suffix.lower()

        # ✅ Convert .emf / .svg to PNG for consistent handling
        if suffix in (".emf", ".svg"):
            converted = altutils.convert_vector_to_png_bytes(
                img_bytes, suffix, docx_path.parent
            )
            img_bytes = converted if converted else None

        if not img_bytes:
            continue

        # Use the *first* occurrence for context/page/section (like AI script)
        p_idx = para_indices[0]
        existing_alt = alt_map.get(rid, "").strip()

        # --- Build context snippet + topic (no API call) ---
        context = altutils.build_context_snippet(paragraph_texts, p_idx, window=1)
        topic = ""
        try:
            # Prefer topic from alt text if present; fall back to context
            if existing_alt:
                topic = altutils.extract_topic_from_altTexts(existing_alt) or ""
            if not topic:
                topic = altutils.extract_topic_from_altTexts(context) or ""
        except Exception:
            topic = ""

        # Thumbnail from image bytes (thumb_bytes, thumb_size)
        thumb_bytes = None
        thumb_size = (0, 0)
        try:
            thumb_bytes, thumb_size = altutils.resize_png(img_bytes)
        except Exception:
            thumb_bytes = None

        # Column A: Image Thumbnail
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

        # Column B (Image Filename)
        ws.cell(row=row, column=2, value=Path(target).name)

        # Column C (Picture ID)
        ws.cell(
            row=row,
            column=3,
            value=docpr_map.get(rid, {}).get("picture_id", ""),
        )

        # Column D (File Name)
        ws.cell(row=row, column=4, value=base_name)

        # Column E (Image Size) — use thumbnail size (same behavior style as AI script)
        if thumb_size and thumb_size[0] and thumb_size[1]:
            ws.cell(
                row=row,
                column=5,
                value=f"{thumb_size[0]} x {thumb_size[1]} px",
            )
        else:
            ws.cell(row=row, column=5, value="")

        # Column F (Image Topic) — from extracted topic
        ws.cell(row=row, column=6, value=topic)

        # Column G (Image Context) — context snippet used
        ws.cell(row=row, column=7, value=context)

        # Column H (Grade Level) — from CONFIG
        ws.cell(row=row, column=8, value=GRADE_LEVEL)

        # Column I (Course Name) — from CONFIG
        ws.cell(row=row, column=9, value=COURSE_NAME)

        # Column J (File Name/Source, with hyperlink)
        file_path_str = str(docx_path)
        j_cell = ws.cell(row=row, column=10, value=file_path_str)
        j_cell.hyperlink = file_path_str
        j_cell.style = "Hyperlink"

        # Column K (Section/Module)
        ws.cell(
            row=row,
            column=11,
            value=altutils.extract_section_from_heading(paragraphs, p_idx),
        )

        # Column L (Page #)
        ws.cell(
            row=row,
            column=12,
            value=altutils.guess_page_number(p_idx),
        )

        # Column M (Document Author, if known)
        ws.cell(row=row, column=13, value=doc_author)

        # Column N (Original Alt Text)
        ws.cell(row=row, column=14, value=existing_alt)

        # Column O (Original Alt Text Character Count) = LEN(N row)
        ws.cell(
            row=row,
            column=15,
            value=f"=LEN({get_column_letter(14)}{row})",
        )
        # Column U (21): Date of Audit → today's date
        ws.cell(row=row, column=21, value=today_str)

        # Columns P–T, V–X remain blank for now.

        row += 1

    # Autofit columns & wrap text (simple version)
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
