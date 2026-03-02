# ====
# UNIFIED: MATCH BY PICTURE ID, APPLY ALT TEXT, AND BUILD QA REPORT
# What this script does:
# 1) Reads Excel (Picture ID, #Column T for alt text to apply)
# 2) Opens DOCX, snapshots ORIGINAL alt text (pre-change)
# 3) Reviews Word for other Equations that are not in Equation Editor
# 4) Applies Excel alt text by Picture ID (supports "N/A" and "Mark as Decorative")
# 5) Builds a QA report with: original columns (no additional review)
# Note: OUTPUT EXCEL FILE: Use Visual Basic script to put images in cell

# to Fix:  
# REPORT TO INCLUDE ALL MASTER COLUMNS
# Not applying alt text to duplicate images
# Images in output file need to be as large as original Excel file

# Brenda Gardner — Updated Nov 5 2025 (revised per requested changes)
# Dec 10 -removed .wmf files in Word report. 
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
from helpers import alttext_utils as utils

# ====
# CONFIG — EDIT THESE
# ====
DOCX_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/2023 Textbook/Cryptography_8_images.docx"
)

EXCEL_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/Small Test Word to AI/103x_Cryptography_AltText_Apply.xlsx"
)
OUTPUT_DOCX  = DOCX_PATH.with_name(DOCX_PATH.stem + "_withAltText.docx")
REPORT_PATH  = DOCX_PATH.with_name(DOCX_PATH.stem + "_AltText_Report.xlsx")

# Excel column mapping (letters).
PICTURE_ID_COLUMN = column_index_from_string("C")  # Excel Picture ID
ALT_TEXT_COLUMN   = column_index_from_string("T")  # Use Excel Suggested Alt Text ("T")
FILE_NAME_COLUMN  = column_index_from_string("J")
SECTION_COLUMN    = column_index_from_string("K")
PAGE_NUMBER_COLUMN= column_index_from_string("L")
NOTES_COLUMN      = column_index_from_string("V")
TOPIC_COLUMN      = column_index_from_string("F")  # NEW: Image Topic in Excel column F

# Thumbnails
ZAMZAR_API_KEY = os.getenv("ZAMZAR")   # set env var ZAMZAR=your_key; used for EMF→PNG
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
            converted = utils.convert_vector_to_png_bytes(img_bytes, suffix, tmp_dir, ZAMZAR_API_KEY)
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
        for idx, p in enumerate(doc.paragraphs):
            # check if the table element is within this paragraph's parent children
            if tbl._element in list(p._element.getparent().getchildren()):
                parent_idx = idx
                break
        if parent_idx == -1:
            parent_idx = len(doc.paragraphs) - 1
        inspect_table(tbl, parent_idx)

    return rid_to_paras

# ====
# STEP 1 — READ EXCEL (keyed by Picture ID)
# ====
wb_src = openpyxl.load_workbook(EXCEL_PATH)
ws_src = wb_src.active

alt_by_pid = {}    # picture_id -> record
alt_by_filename = {}    # filename   -> record (fallback)

records_loaded = 0
for row in ws_src.iter_rows(min_row=2):
    picture_id = _norm_pid(row[PICTURE_ID_COLUMN - 1].value)
    alt_text   = safe_str(row[ALT_TEXT_COLUMN   - 1].value)
    image_name = safe_str(row[1].value)    # column B in the source Excel (1-based index)
    file_name  = safe_str(row[FILE_NAME_COLUMN  - 1].value)
    section    = safe_str(row[SECTION_COLUMN    - 1].value)
    page_num   = safe_str(row[PAGE_NUMBER_COLUMN- 1].value)
    notes      = safe_str(row[NOTES_COLUMN    - 1].value)
    topic      = safe_str(row[TOPIC_COLUMN    - 1].value)  # NEW: topic from Excel column F

    if not (picture_id or file_name or alt_text):
        continue

    rec = {
        "picture_id": picture_id,
        "image_name": image_name,
        "alt_text": alt_text,
        "file_name": file_name,
        "section": section,
        "page_number": page_num,
        "notes": notes,
        "topic": topic,
    }

    if picture_id:
        alt_by_pid[picture_id] = rec
        records_loaded += 1
    if file_name:
        alt_by_filename[file_name] = rec

print(f"✅ Loaded {records_loaded} Excel records keyed by Picture ID; fallback filename keys: {len(alt_by_filename)}")

# Keep for report checks by filename presence
alt_text_map = alt_by_filename

# ====
# Prepare python-docx context for heading lookup (NEW)
# ====
try:
    doc_for_headings = Document(str(DOCX_PATH))
    paragraph_objs = doc_for_headings.paragraphs
    rid_para_map = locate_all_image_instances(doc_for_headings)
    print(f"📄 Parsed document paragraphs: {len(paragraph_objs)}; found {len(rid_para_map)} image instances (rIds) via python-docx")
except Exception as e:
    print(f"⚠ Could not parse DOCX with python-docx for heading extraction: {e}")
    doc_for_headings = None
    paragraph_objs = []
    rid_para_map = {}

# ====
# STEP 2 — OPEN DOCX, SNAPSHOT ORIGINAL ALT TEXTS, APPLY UPDATES
# ====
with tempfile.TemporaryDirectory() as tmpdir:
    tmp_path = Path(tmpdir)
    with zipfile.ZipFile(DOCX_PATH, "r") as zip_ref:
        zip_ref.extractall(tmp_path)

    document_xml_path = tmp_path / "word" / "document.xml"
    rels_path    = tmp_path / "word" / "_rels" / "document.xml.rels"

    ns = {
        "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    tree    = etree.parse(str(document_xml_path))
    root    = tree.getroot()
    rels_tree = etree.parse(str(rels_path))
    rels_root = rels_tree.getroot()

    # rId -> filename
    rels_map = {}
    for rel in rels_root.findall(".//", namespaces=ns):
        rid = rel.get("Id")
        target = rel.get("Target")
        if rid and target and target.startswith("media/"):
            rels_map[rid] = target.split("/")[-1]

    print(f"🖼 Found {len(rels_map)} image relationships in DOCX")

    word_images    = set()
    word_action_log    = {}
    original_alt_text_map = {}  # filename -> list of ORIGINAL descr (pre-change) - handles duplicates
    matched_filenames    = set()  # filenames that got a valid Excel match
    applied_filenames    = set()  # filenames where an Excel alt text was applied (Alt Text Added / Marked as Decorative)

    # NEW: filename -> list of picture_ids, and filename -> list of docPr titles to capture multiple occurrences
    filename_to_pids = defaultdict(list)
    filename_to_titles = defaultdict(list)

    # --- SNAPSHOT ORIGINAL ALT TEXTS (before any changes)
    for docPr in root.findall(".//wp:docPr", ns):
        drawing = docPr.getparent()
        blip = drawing.find(".//a:blip", ns)
        if blip is None:
            continue
        rid = blip.get(f"{{{ns['r']}}}embed")
        if not rid:
            continue
        filename = rels_map.get(rid)
        if not filename:
            continue

        # Add to word_images set
        word_images.add(filename)

        # Capture original alt text (handle duplicates)
        original_descr = docPr.get("descr", "")
        if filename not in original_alt_text_map:
            original_alt_text_map[filename] = []
        original_alt_text_map[filename].append(original_descr)

        # Capture picture_id(s) and titles per filename (preserve duplicates)
        picture_id_raw = docPr.get("id")
        picture_id = _norm_pid(picture_id_raw)
        if picture_id:
            filename_to_pids[filename].append(picture_id)

        title = docPr.get("title", "")
        if title:
            filename_to_titles[filename].append(title)

        # NEW: attempt to derive section from preceding Heading 1 using python-docx paragraph index if available
        try:
            if rid_para_map and rid in rid_para_map and doc_for_headings:
                para_indices = rid_para_map.get(rid, [])
                if para_indices:
                    p_idx = para_indices[0]
                    section_guess = extract_section_from_heading(paragraph_objs, p_idx)
                    if section_guess:
                        filename_to_titles[filename].append(section_guess)
        except Exception:
            # non-fatal
            pass

    # --- APPLY UPDATES BY PICTURE ID (iterate every docPr element so repeated PictureIDs are all updated)
    updated = 0
    for docPr in root.findall(".//wp:docPr", ns):
        drawing = docPr.getparent()
        blip = drawing.find(".//a:blip", ns)
        rid = blip.get(f"{{{ns['r']}}}embed") if blip is not None else None
        filename = rels_map.get(rid) if rid else None
        if not filename:
            continue

        # Picture ID from Word
        picture_id_raw = docPr.get("id")
        picture_id = _norm_pid(picture_id_raw)

        excel_match = alt_by_pid.get(picture_id)
        if not excel_match:
            # do NOT log mismatch here; we decide after the loop
            continue

        # — we have a match —
        alt_text = (excel_match.get("alt_text") or "").strip()

        if alt_text.lower() == "n/a":
            word_action_log[filename] = "Modify this image"
            matched_filenames.add(filename)
            continue

        if alt_text.lower() == "mark as decorative":
            # Proper decorative toggle
            def mark_decorative(docPr_el):
                # Clear legacy attrs
                docPr_el.attrib.pop("descr", None)
                docPr_el.attrib.pop("title", None)
                docPr_el.set("descr", "")
                docPr_el.set("title", "")

                # Add a11y:decorative ext (best-effort)
                # Namespaces
                A_NS = ns["a"]
                A11Y_NS = "http://schemas.microsoft.com/office/drawing/2018/a11y"
                etree.register_namespace("a11y", A11Y_NS)

                extLst = docPr_el.find(f".//{{{A_NS}}}extLst")
                if extLst is None:
                    extLst = etree.Element(f"{{{A_NS}}}extLst")
                    docPr_el.append(extLst)

                ext = None
                for candidate in extLst.findall(f".//{{{A_NS}}}ext"):
                    if candidate.find(f".//{{{A11Y_NS}}}decorative") is not None:
                        ext = candidate
                        break
                if ext is None:
                    ext = etree.Element(f"{{{A_NS}}}ext")
                    ext.set("uri", "{0000-0000-0000-0000-0000}")  # placeholder URI
                    extLst.append(ext)

                decorative = ext.find(f".//{{{A11Y_NS}}}decorative")
                if decorative is None:
                    decorative = etree.Element(f"{{{A11Y_NS}}}decorative")
                    ext.append(decorative)
                decorative.set("val", "1")

            mark_decorative(docPr)
            word_action_log[filename] = "Marked as Decorative"
            matched_filenames.add(filename)
            applied_filenames.add(filename)
            updated += 1
            continue

        # Standard alt text path
        if alt_text:
            docPr.attrib.pop("title", None)  # remove any stale title
            docPr.set("descr", alt_text)  # write new alt text
            word_action_log[filename] = "Alt Text Added"
            matched_filenames.add(filename)
            applied_filenames.add(filename)
            updated += 1
        else:
            # matched PID but empty value: leave as no-change
            word_action_log[filename] = "No Change"
            matched_filenames.add(filename)

    # ✅ FIX: Changed "Image ID Mismatch" to "Not in Excel"
    for fname in word_images:
        if fname not in matched_filenames and word_action_log.get(fname) is None:
            word_action_log[fname] = "Not in Excel"

    # Save updated Word file
    tree.write(str(document_xml_path), xml_declaration=True, encoding="utf-8", standalone="yes")
    with zipfile.ZipFile(OUTPUT_DOCX, "w", zipfile.ZIP_DEFLATED) as docx_zip:
        for folder_name, _, filenames in os.walk(tmp_path):
            for fn in filenames:
                fp = Path(folder_name) / fn
                arcname = fp.relative_to(tmp_path)
                docx_zip.write(fp, arcname)

print(f"\n🎉 Saved new Word file with alt text to:\n{OUTPUT_DOCX}")

# Re-scan output docx for media filenames for the report
with zipfile.ZipFile(OUTPUT_DOCX, "r") as zip_ref:
    media_files = {
        Path(f).name
        for f in zip_ref.namelist()
        if f.startswith("word/media/") and not f.endswith("/")
    }
print(f"🖼 Found {len(media_files)} images in Word media folder for QA report")

# ====
# PREP — COMBINE IMAGE SOURCES FOR REPORT
# ====
excel_images = set(alt_text_map.keys())

# Combine all known images:
#    - those in the Word file (referenced pics)
#    - those extracted from media/
#    - those listed in Excel metadata
all_images = sorted(word_images.union(media_files).union(excel_images))

print(f"🧩 Preparing QA report for {len(all_images)} total images")
print(f"   • {len(word_images)} from Word (drawings in document.xml)")
print(f"   • {len(media_files)} from media folder")
print(f"   • {len(excel_images)} from Excel metadata")

# ====
# STEP 3 – BUILD EXCEL QA REPORT  (FINAL CLEAN VERSION)
# ====
GREEN_FONT = "FF008000"  # dark green
ORANGE_FONT = "FFCC8400"  # orange-gold

wb = Workbook()
ws = wb.active
ws.title = "AltText QA Report"

# --- Header setup ---
headers = [
    "Thumbnail",           # A
    "Image Name",          # B
    "Picture ID",          # C
    "Original File",       # D
    "Image Topic",         # E (NEW)
    "Section/Module",      # F
    "Page #",              # G
    "In Word?",            # H
    "In Excel?",           # I
    "Word Alt Text",       # J
    "Applied/New Alt Text (Excel)", # K
    "Word Action",         # L
    "Status",              # M
    "Notes"                # N
]

col_widths = [50, 30, 25, 25, 20, 12, 10, 10, 10, 70, 40, 20, 18, 40]

for col, h in enumerate(headers, start=1):
    c = ws.cell(row=1, column=col, value=h)
    c.font = Font(bold=True)
    ws.column_dimensions[get_column_letter(col)].width = col_widths[col - 1]

# --- Set default alignments for later cells ---
wrap_alignment = Alignment(wrap_text=True, vertical="top")
shrink_wrap_alignment = Alignment(wrap_text=True, shrink_to_fit=True, vertical="top")

with tempfile.TemporaryDirectory() as tmp_dir:
    tmp_dir = Path(tmp_dir)
    print(f"📁 Temporary image folder: {tmp_dir}")

    new_alt_applied_count = 0
    row_num = 2

    for name in all_images:
        # Try to find Excel record by: 1) filename, 2) Picture ID from Word
        rec = alt_text_map.get(name, {})
        if not rec and name in word_images:
            # Get Picture ID(s) from Word for this filename
            pids = filename_to_pids.get(name, [])
            if pids:
                # Prefer first matching Excel record by PID
                for pid in pids:
                    maybe = alt_by_pid.get(pid)
                    if maybe:
                        rec = maybe
                        break

        file_name    = rec.get("file_name", "")
        section      = rec.get("section", "")
        page_num     = rec.get("page_number", "")
        notes        = rec.get("notes", "")
        excel_alt_val = (rec.get("alt_text") or "").strip()
        topic_val     = (rec.get("topic") or "").strip()  # NEW: topic from Excel

        # --- Original Word alt text (Column J) - handles duplicates ---
        alt_texts = original_alt_text_map.get(name, [])
        word_alt_text = " | ".join(alt_texts) if alt_texts else ""

        # Always show Excel alt text in Column K
        applied_alt_val = excel_alt_val
        if excel_alt_val and excel_alt_val != word_alt_text:
            new_alt_applied_count += 1

        # --- Membership flags ---
        in_word  = "Yes" if name in word_images else ""
        in_excel = "Yes" if name in excel_images else ""
        # If this image had an Excel alt applied to it in Word, mark both Yes
        if name in applied_filenames:
            in_word = "Yes"
            in_excel = "Yes"

        # --- Insert thumbnail or fallback text ---
        if name in media_files:
            try:
                insert_image_from_docx(ws, OUTPUT_DOCX, name, row_num, tmp_dir)
                adjust_row_height(ws, row_num, ws.row_dimensions[row_num].height)
            except Exception as e:
                ws.cell(row=row_num, column=1, value=f"(no image: {e})")
                adjust_row_height(ws, row_num)
        else:
            ws.cell(row=row_num, column=1, value="(Excel-only — no thumbnail)")
            adjust_row_height(ws, row_num, None, min_height=50)

        # --- Picture ID(s) for this image (Column C) ---
        picture_id = rec.get("picture_id", "")
        if not picture_id:
            pids_for_name = filename_to_pids.get(name, [])
            if pids_for_name:
                picture_id = ", ".join(pids_for_name)

        # If still empty, keep as ""
        picture_id = picture_id or ""

        # --- If no section from Excel, try docPr title(s) or heading-based section captured earlier ---
        if not section:
            titles = filename_to_titles.get(name, [])
            if titles:
                section = " | ".join(titles)

        # --- Write base metadata (Columns B–F) ---
        ws.cell(row=row_num, column=2, value=name)           # Image Name (B)
        ws.cell(row=row_num, column=3, value=picture_id)     # Picture ID (C)
        ws.cell(row=row_num, column=4, value=file_name)      # Original File (D)
        ws.cell(row=row_num, column=5, value=topic_val)      # Image Topic (E) NEW
        ws.cell(row=row_num, column=6, value=section)        # Section/Module (F)
        ws.cell(row=row_num, column=7, value=page_num)       # Page # (G)

        # --- In Word / In Excel flags (H,I) ---
        ws.cell(row=row_num, column=8, value=in_word)
        ws.cell(row=row_num, column=9, value=in_excel)

        # --- Alt Text columns (J,K) ---
        ws.cell(row=row_num, column=10, value=word_alt_text).alignment = shrink_wrap_alignment
        cell_new = ws.cell(row=row_num, column=11, value=applied_alt_val)
        cell_new.alignment = shrink_wrap_alignment

        # --- Word Action + Status determination (L,M) ---
        word_action = word_action_log.get(name, "No Change")
        ws.cell(row=row_num, column=12, value=word_action)

        # Determine status
        if word_action in ("Alt Text Added", "Marked as Decorative"):
            status = "OK"
        elif word_action in ("Not in Excel", "Modify this image", "No Change"):
            status = "Review"
        elif name in excel_images and name not in word_images:
            status = "Excel-only"
        elif name in word_images and name not in excel_images:
            status = "Word-Only"
        else:
            status = "Review"

        s_cell = ws.cell(row=row_num, column=13, value=status)
        if status == "OK":
            s_cell.font = Font(bold=True, color=GREEN_FONT)
        elif status == "Review":
            s_cell.font = Font(bold=True, color=ORANGE_FONT)
        else:
            s_cell.font = Font(bold=True)

        ws.cell(row=row_num, column=14, value=notes).alignment = wrap_alignment
        row_num += 1

    # ====
    # SUMMARY SHEET — built inside temp block (so thumbnails persist)
    # ====
    summary_ws = wb.create_sheet("Summary")
    summary_ws["A1"] = "Alt Text Report Summary"
    summary_ws["A1"].font = Font(bold=True, size=14)

    summary_ws["A3"], summary_ws["B3"] = "Total Word Media Files", len(media_files)
    summary_ws["A4"], summary_ws["B4"] = "Total Excel Alt Text Entries", len(excel_images)
    summary_ws["A5"], summary_ws["B5"] = "Matched Word + Excel", len(media_files.intersection(excel_images))
    summary_ws["A6"], summary_ws["B6"] = "Word-only Images", len(media_files - excel_images)
    summary_ws["A7"], summary_ws["B7"] = "Excel-only Entries", len(excel_images - media_files)
    summary_ws["A8"], summary_ws["B8"] = "New Alt Texts Applied", new_alt_applied_count

    summary_ws["B8"].font = Font(bold=True, color=GREEN_FONT)

    summary_ws["A10"] = "Status Counts"
    summary_ws["A10"].font = Font(bold=True, size=12)

    status_counts = {}
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=13).value
        if val:
            status_counts[val] = status_counts.get(val, 0) + 1

    row = 12
    for status, count in sorted(status_counts.items()):
        c_status = summary_ws.cell(row=row, column=1, value=status)
        c_count  = summary_ws.cell(row=row, column=2, value=count)
        if status == "OK":
            c_status.font = Font(color=GREEN_FONT, bold=True)
            c_count.font  = Font(color=GREEN_FONT, bold=True)
        elif status == "Review":
            c_status.font = Font(color=ORANGE_FONT, bold=True)
            c_count.font  = Font(color=ORANGE_FONT, bold=True)
        row += 1

    summary_ws.column_dimensions["A"].width = 35
    summary_ws.column_dimensions["B"].width = 15
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    # ✅ Save workbook before temp dir is deleted
    wb.save(REPORT_PATH)
    print(f"✅ Excel QA report saved to:\n{REPORT_PATH}")
