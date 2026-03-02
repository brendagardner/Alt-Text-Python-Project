# Applies alt text from a SINGLE Excel file to MULTIPLE Word documents.
# Uses Column I (source file name) to group Excel rows by document,
# and matches images by Picture ID (Column C) — Primary Key in Excel.
#
# Created by Brenda Gardner, March 2026
#
# HOW IT WORKS:
#   1. Reads the Excel file and groups rows by source file name (Column I)
#   2. Scans WORD_FOLDER for all .docx files
#   3. Uses NAME_MAPPING to connect Excel source names to actual Word filenames
#   4. For each matched Word doc: applies alt text by Picture ID, saves output
#   5. For unmatched Word docs with images: reports "No Alt Text in Excel File"
#   6. For unmatched Word docs with NO images: skips silently
#   7. Saves a consolidated QA report covering all documents
#
# MANUALLY REVIEW:
#   - Edit the paths in CONFIG
#   - Update NAME_MAPPING when file names change (see instructions there)

import zipfile
import os
import sys
import tempfile
import difflib
from io import BytesIO
from pathlib import Path
from collections import defaultdict
from typing import Union
import datetime

import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree
from PIL import Image as PILImage
from docx import Document

# Import shared utilities (alttext_utils.py must be in the same folder)
sys.path.insert(0, str(Path(__file__).parent))
from helpers import alttext_utils as utils

# ====================================================================
# CONFIG — EDIT THIS SECTION BEFORE RUNNING
# ====================================================================

EXCEL_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Shared Alt Text Files/1060 Alt Text Files/1060_Mod1_AltText_REVIEW.xlsx"
)

WORD_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1060 Information/2026 Trigonometry Chap 1, pre-alt text"
)

# Output folder for Word docs with alt text applied (created automatically)
OUTPUT_FOLDER = WORD_FOLDER / "with_alt_text"

# QA report file (saved to the Word folder)
REPORT_PATH = WORD_FOLDER / "1060_Mod1_AltText_QA_Report.xlsx"

# ── NAME MAPPING ──────────────────────────────────────────────────────────────
# This table connects the ORIGINAL source file names listed in Excel Column I
# to the ACTUAL/NEW Word document filenames in WORD_FOLDER.
#
# Left side  = exactly as it appears in Column I of the Excel file
# Right side = the actual filename in the Word folder
#
# WHEN NAMES CHANGE IN FUTURE PROJECTS:
#   - Open the Excel file and look at Column I to see what names are listed there
#   - Look in the Word folder to see the actual filenames
#   - Update the right-hand values below to match the new actual filenames
#   - The left-hand values should match Excel Column I exactly (copy & paste to be safe)
#
# TIP: If the script prints "FUZZY MATCH SUGGESTIONS" at the top of its output,
#      those are likely correct matches you can just paste in here.
# ─────────────────────────────────────────────────────────────────────────────
NAME_MAPPING = {
    "Exercises 1.2 OER Trigonometry 2023.docx": "Exercises 1.2 OER Trigonometry 2026.docx",
    "Exercises 1.6 OER Trigonometry 2023.docx": "Exercises 1.6 OER Trigonometry 2026.docx",
    "Intro Chapter 1 OER Trigonometry 2023.docx": "Intro Chapter 1 OER Trigonometry 2026.docx",
    "Section 1.1 OER Trigonometry 2023.docx": "Section 1.1 OER Trigonometry 2026.docx",
    "Section 1.2 OER Trigonometry 2023.docx": "Section 1.2 OER Trigonometry 2026.docx",
    "Section 1.3 OER Trigonometry 2023.docx": "Section 1.3 OER Trigonometry 2026.docx",
    "Section 1.4 OER Trigonometry 2023.docx": "Section 1.4 OER Trigonometry 2026.docx",
    "Section 1.6 OER Trigonometry 2023.docx": "Section 1.6 OER Trigonometry 2026.docx",
}

# Excel column letters — adjust only if your spreadsheet layout changes
PICTURE_ID_COLUMN  = column_index_from_string("C")   # Picture ID (key for matching)
ALT_TEXT_COLUMN    = column_index_from_string("T")   # Suggested Alt Text to apply
FILE_NAME_COLUMN   = column_index_from_string("I")   # Source document name
SECTION_COLUMN     = column_index_from_string("J")   # Section / Module
PAGE_NUMBER_COLUMN = column_index_from_string("K")   # Page number
NOTES_COLUMN       = column_index_from_string("V")   # Notes
TOPIC_COLUMN       = column_index_from_string("E")   # Image topic

LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


# ====================================================================
# HELPER FUNCTIONS  (no editing needed below this line)
# ====================================================================

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
    return "" if val is None else str(val).strip()


def _get_ns():
    return {
        "w":  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a":  "http://schemas.openxmlformats.org/drawingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "r":  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }


def _build_rels_map(rels_root, ns) -> dict:
    """Build rId -> image filename map from relationship XML."""
    rels_map = {}
    for rel in rels_root.findall(".//", namespaces=ns):
        rid = rel.get("Id")
        target = rel.get("Target", "")
        if rid and target.startswith("media/"):
            rels_map[rid] = target.split("/")[-1]
    return rels_map


def scan_word_doc_for_images(docx_path: Path) -> list:
    """Return list of (picture_id, image_filename, original_alt_text) for every image."""
    ns = _get_ns()
    images = []
    try:
        with zipfile.ZipFile(docx_path) as z:
            rels_xml = z.read("word/_rels/document.xml.rels")
            doc_xml  = z.read("word/document.xml")
    except Exception as e:
        print(f"    ⚠ Could not open {docx_path.name}: {e}")
        return images

    rels_map = _build_rels_map(etree.fromstring(rels_xml), ns)
    doc_root = etree.fromstring(doc_xml)

    for docPr in doc_root.findall(".//wp:docPr", ns):
        drawing = docPr.getparent()
        blip = drawing.find(".//a:blip", ns)
        if blip is None:
            continue
        rid = blip.get(f"{{{ns['r']}}}embed")
        if not rid:
            continue
        filename = rels_map.get(rid, "")
        if filename:
            images.append((
                _norm_pid(docPr.get("id")),
                filename,
                docPr.get("descr", ""),
            ))
    return images


def _mark_decorative(docPr_el, ns):
    """Set the Word 'Mark as Decorative' flag on a docPr element."""
    A_NS    = ns["a"]
    A11Y_NS = "http://schemas.microsoft.com/office/drawing/2018/a11y"
    etree.register_namespace("a11y", A11Y_NS)
    docPr_el.attrib.pop("descr", None)
    docPr_el.attrib.pop("title", None)
    docPr_el.set("descr", "")
    docPr_el.set("title", "")
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
        ext.set("uri", "{0000-0000-0000-0000-0000}")
        extLst.append(ext)
    decorative = ext.find(f".//{{{A11Y_NS}}}decorative")
    if decorative is None:
        decorative = etree.Element(f"{{{A11Y_NS}}}decorative")
        ext.append(decorative)
    decorative.set("val", "1")


def apply_alt_text_to_doc(docx_path: Path, alt_by_pid: dict, output_path: Path):
    """
    Apply alt text from alt_by_pid (picture_id -> record) to docx_path.
    Saves result to output_path.
    Returns: (word_action_log, original_alt_text_map, word_images, filename_to_pids)
    """
    ns = _get_ns()
    word_action_log       = {}
    original_alt_text_map = defaultdict(list)
    word_images           = set()
    matched_filenames     = set()
    applied_filenames     = set()
    filename_to_pids      = defaultdict(list)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        with zipfile.ZipFile(docx_path, "r") as zip_ref:
            zip_ref.extractall(tmp_path)

        document_xml_path = tmp_path / "word" / "document.xml"
        rels_path          = tmp_path / "word" / "_rels" / "document.xml.rels"

        tree      = etree.parse(str(document_xml_path))
        root      = tree.getroot()
        rels_tree = etree.parse(str(rels_path))
        rels_map  = _build_rels_map(rels_tree.getroot(), ns)

        # ── Snapshot original alt texts (before any changes) ──────────────
        for docPr in root.findall(".//wp:docPr", ns):
            drawing = docPr.getparent()
            blip    = drawing.find(".//a:blip", ns)
            if blip is None:
                continue
            rid      = blip.get(f"{{{ns['r']}}}embed")
            filename = rels_map.get(rid) if rid else None
            if not filename:
                continue
            word_images.add(filename)
            original_alt_text_map[filename].append(docPr.get("descr", ""))
            pid = _norm_pid(docPr.get("id"))
            if pid:
                filename_to_pids[filename].append(pid)

        # ── Apply alt text by Picture ID ───────────────────────────────────
        updated = 0
        for docPr in root.findall(".//wp:docPr", ns):
            drawing  = docPr.getparent()
            blip     = drawing.find(".//a:blip", ns)
            rid      = blip.get(f"{{{ns['r']}}}embed") if blip is not None else None
            filename = rels_map.get(rid) if rid else None
            if not filename:
                continue

            pid          = _norm_pid(docPr.get("id"))
            excel_match  = alt_by_pid.get(pid)
            if not excel_match:
                continue

            alt_text = (excel_match.get("alt_text") or "").strip()

            if alt_text.lower() == "n/a":
                word_action_log[filename] = "Modify This Image"
                matched_filenames.add(filename)
                continue

            if alt_text.lower() == "mark as decorative":
                _mark_decorative(docPr, ns)
                word_action_log[filename] = "Marked as Decorative"
                matched_filenames.add(filename)
                applied_filenames.add(filename)
                updated += 1
                continue

            if alt_text:
                docPr.attrib.pop("title", None)
                docPr.set("descr", alt_text)
                word_action_log[filename] = "Alt Text Added"
                matched_filenames.add(filename)
                applied_filenames.add(filename)
                updated += 1
            else:
                word_action_log[filename] = "No Change"
                matched_filenames.add(filename)

        # Mark images present in Word but not in Excel
        for fname in word_images:
            if fname not in matched_filenames and word_action_log.get(fname) is None:
                word_action_log[fname] = "Not in Excel"

        # ── Save updated Word file ─────────────────────────────────────────
        output_path.parent.mkdir(parents=True, exist_ok=True)
        tree.write(
            str(document_xml_path),
            xml_declaration=True, encoding="utf-8", standalone="yes"
        )
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as docx_zip:
            for folder_name, _, filenames in os.walk(tmp_path):
                for fn in filenames:
                    fp = Path(folder_name) / fn
                    docx_zip.write(fp, fp.relative_to(tmp_path))

    print(f"    ✅ Saved: {output_path.name}  ({updated} alt texts applied)")
    return word_action_log, dict(original_alt_text_map), word_images, filename_to_pids


def insert_image_from_docx(ws, docx_path: Path, target_filename: str, row: int, tmp_dir: Path):
    """Extract an image from inside a DOCX and embed it as a thumbnail in Excel."""
    try:
        img_bytes = utils.load_image_bytes(docx_path, f"media/{target_filename}")
        if not img_bytes:
            ws.cell(row=row, column=1, value=f"(missing: {target_filename})")
            ws.row_dimensions[row].height = 50
            return

        suffix = Path(target_filename).suffix.lower()
        if suffix in (".emf", ".wmf", ".svg"):
            converted = utils.convert_vector_to_png_bytes(img_bytes, suffix, tmp_dir)
            if converted:
                img_bytes = converted
            else:
                ws.cell(row=row, column=1, value=f"(vector — not rendered: {target_filename})")
                ws.row_dimensions[row].height = 80
                return

        height = 0
        try:
            pil_img = PILImage.open(BytesIO(img_bytes))
            height  = pil_img.size[1]
        except Exception:
            pass

        xl_img = XLImage(BytesIO(img_bytes))
        ws.add_image(xl_img, f"A{row}")
        scale = getattr(utils, "ROW_HEIGHT_SCALE", 0.75)
        ws.row_dimensions[row].height = int(height * scale) if height else 80

    except Exception as e:
        ws.cell(row=row, column=1, value=f"(error: {e})")
        ws.row_dimensions[row].height = 50


def suggest_fuzzy_matches(excel_source_names: list, actual_word_files: list) -> dict:
    """
    Compare Excel source names to actual Word filenames and suggest close matches.
    These are printed as hints — nothing is auto-applied.
    """
    suggestions = {}
    actual_lower = {f.lower(): f for f in actual_word_files}
    for source_name in excel_source_names:
        if source_name in NAME_MAPPING:
            continue   # already handled
        candidates = difflib.get_close_matches(
            source_name.lower(), actual_lower.keys(), n=1, cutoff=0.5
        )
        if candidates:
            suggestions[source_name] = actual_lower[candidates[0]]
    return suggestions


# ====================================================================
# STEP 1 — READ EXCEL  (group records by source file name, Column I)
# ====================================================================

print(f"\n📊 Loading Excel: {EXCEL_PATH.name}")
wb_src = openpyxl.load_workbook(EXCEL_PATH)
ws_src = wb_src.active

# excel_groups: source_name (Column I) -> {picture_id -> record}
excel_groups = defaultdict(dict)

for row in ws_src.iter_rows(min_row=2):
    picture_id = _norm_pid(row[PICTURE_ID_COLUMN - 1].value)
    alt_text   = safe_str(row[ALT_TEXT_COLUMN   - 1].value)
    file_name  = safe_str(row[FILE_NAME_COLUMN  - 1].value)
    section    = safe_str(row[SECTION_COLUMN    - 1].value)
    page_num   = safe_str(row[PAGE_NUMBER_COLUMN - 1].value)
    notes      = safe_str(row[NOTES_COLUMN      - 1].value)
    topic      = safe_str(row[TOPIC_COLUMN      - 1].value)
    image_name = safe_str(row[1].value)   # Column B

    if not (picture_id or file_name or alt_text):
        continue

    rec = {
        "picture_id":  picture_id,
        "image_name":  image_name,
        "alt_text":    alt_text,
        "file_name":   file_name,
        "section":     section,
        "page_number": page_num,
        "notes":       notes,
        "topic":       topic,
    }

    if picture_id and file_name:
        excel_groups[file_name][picture_id] = rec

total_records = sum(len(v) for v in excel_groups.values())
print(f"  ✅ Found {len(excel_groups)} source files, {total_records} total records")
for src_name, recs in sorted(excel_groups.items()):
    print(f"     • {src_name}  ({len(recs)} images)")


# ====================================================================
# STEP 2 — SCAN WORD FOLDER
# ====================================================================

word_docs = sorted(d for d in WORD_FOLDER.glob("*.docx") if not d.name.startswith("~$"))
print(f"\n📁 Found {len(word_docs)} Word docs in: {WORD_FOLDER.name}")

# Reverse mapping: actual word filename -> Excel source name
reverse_mapping = {v: k for k, v in NAME_MAPPING.items()}

# Check for unmapped Excel names and suggest fuzzy matches
actual_word_names = [d.name for d in word_docs]
suggestions = suggest_fuzzy_matches(list(excel_groups.keys()), actual_word_names)

if suggestions:
    print()
    print("💡 FUZZY MATCH SUGGESTIONS (not auto-applied):")
    print("   If these look correct, copy them into the NAME_MAPPING section above.")
    for src, suggested in suggestions.items():
        print(f"   Excel Column I: '{src}'")
        print(f"   Likely match:   '{suggested}'")
        print()


# ====================================================================
# STEP 3 — PROCESS EACH WORD DOCUMENT
# ====================================================================

OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# Master list of rows for the QA report
report_rows = []

for docx_path in word_docs:
    doc_name = docx_path.name
    print(f"\n📄 {doc_name}")

    # ── Look up this Word doc in the NAME_MAPPING ─────────────────────────
    excel_source_name = reverse_mapping.get(doc_name)
    alt_by_pid        = excel_groups.get(excel_source_name, {}) if excel_source_name else {}

    # ── No Excel data for this Word doc ───────────────────────────────────
    if not alt_by_pid:
        print(f"  ⚠ No NAME_MAPPING entry — scanning for images...")
        images_in_doc = scan_word_doc_for_images(docx_path)

        if not images_in_doc:
            # No images at all → skip silently (no rows added to report)
            print(f"  ✔ No images found — skipping")
            continue

        # Images present but no alt text data → log in report
        print(f"  ⚠ Found {len(images_in_doc)} image(s) with no alt text in Excel")
        for pic_id, img_filename, orig_alt in images_in_doc:
            report_rows.append({
                "word_doc":       doc_name,
                "image_filename": img_filename,
                "picture_id":     pic_id,
                "file_name":      excel_source_name or "",
                "topic":          "",
                "section":        "",
                "page_number":    "",
                "original_alt":   orig_alt,
                "applied_alt":    "",
                "notes":          "",
                "status":         "No Alt Text in Excel File",
                "source_docx":    docx_path,
                "output_docx":    None,
            })
        continue

    # ── Excel data found → apply alt text ─────────────────────────────────
    print(f"  🔗 Excel source: '{excel_source_name}'  ({len(alt_by_pid)} records)")
    output_path = OUTPUT_FOLDER / doc_name

    word_action_log, orig_alt_map, word_images, fname_to_pids = \
        apply_alt_text_to_doc(docx_path, alt_by_pid, output_path)

    # ── Collect rows for QA report ────────────────────────────────────────
    for img_filename in sorted(word_images):
        # Find the matching Excel record via Picture ID
        pids = fname_to_pids.get(img_filename, [])
        rec  = {}
        for pid in pids:
            maybe = alt_by_pid.get(pid)
            if maybe:
                rec = maybe
                break

        action = word_action_log.get(img_filename, "No Change")

        if action == "Alt Text Added":
            status = "OK"
        elif action == "Marked as Decorative":
            status = "OK — Marked as Decorative"
        elif action == "Modify This Image":
            status = "Review — Modify This Image"
        elif action == "Not in Excel":
            status = "No Alt Text in Excel File"
        else:
            status = "Review"

        orig_alts = orig_alt_map.get(img_filename, [])
        report_rows.append({
            "word_doc":       doc_name,
            "image_filename": img_filename,
            "picture_id":     rec.get("picture_id") or (", ".join(pids) if pids else ""),
            "file_name":      rec.get("file_name", excel_source_name or ""),
            "topic":          rec.get("topic", ""),
            "section":        rec.get("section", ""),
            "page_number":    rec.get("page_number", ""),
            "original_alt":   " | ".join(orig_alts),
            "applied_alt":    rec.get("alt_text", ""),
            "notes":          rec.get("notes", ""),
            "status":         status,
            "source_docx":    docx_path,
            "output_docx":    output_path,
        })

print(f"\n✅ All Word documents processed.")
print(f"   QA report will cover {len(report_rows)} images across "
      f"{len(set(r['word_doc'] for r in report_rows))} documents.")


# ====================================================================
# STEP 4 — BUILD CONSOLIDATED QA REPORT
# ====================================================================

print(f"\n📊 Building QA report...")

GREEN_FONT  = "FF008000"
ORANGE_FONT = "FFCC8400"
RED_FONT    = "FFCC0000"

wb = Workbook()
ws = wb.active
ws.title = "AltText QA Report"

# ── Column headers ────────────────────────────────────────────────────────────
headers = [
    "Image Thumbnail",                    # A
    "Word Document",                      # B  ← which Word file
    "Image Filename",                     # C
    "Picture ID",                         # D
    "In Excel?",                          # E
    "Excel Source File (Column I)",       # F
    "Image Topic",                        # G
    "Section / Module",                   # H
    "Page",                               # I
    "Original Alt Text (pre-change)",     # J
    "Orig. Char. Count",                  # K
    "Applied Alt Text",                   # L
    "Appl. Char. Count",                  # M
    "Notes",                              # N
    "Status",                             # O
    "Date Run",                           # P
    "Reviewer Notes",                     # Q  ← blank for manual entry
]
col_widths = [50, 38, 26, 10, 14, 38, 20, 18, 8, 60, 10, 60, 10, 35, 28, 14, 30]

for col, h in enumerate(headers, start=1):
    c = ws.cell(row=1, column=col, value=h)
    c.font = Font(bold=True)
    ws.column_dimensions[get_column_letter(col)].width = col_widths[col - 1]

wrap_align = Alignment(wrap_text=True, vertical="top")
run_date   = datetime.date.today()

# Sort by Word doc name, then numerically by Picture ID
def _sort_key(r):
    pid = r.get("picture_id", "")
    try:
        return (r["word_doc"], int(pid))
    except (ValueError, TypeError):
        return (r["word_doc"], 0)

report_rows.sort(key=_sort_key)

# ── Write report rows ─────────────────────────────────────────────────────────
with tempfile.TemporaryDirectory() as tmp_dir:
    tmp_dir = Path(tmp_dir)

    for row_num, row_data in enumerate(report_rows, start=2):
        word_doc    = row_data["word_doc"]
        img_fname   = row_data["image_filename"]
        picture_id  = row_data["picture_id"]
        status      = row_data["status"]
        orig_alt    = row_data["original_alt"]
        applied_alt = row_data["applied_alt"]
        source_docx = row_data["source_docx"]
        output_docx = row_data["output_docx"]

        in_excel_label = (
            "No — Alt Text Missing"
            if status == "No Alt Text in Excel File"
            else "Yes"
        )

        # A: Image thumbnail (use output doc if it exists, else source)
        thumb_docx = (
            output_docx if (output_docx and output_docx.exists()) else source_docx
        )
        if thumb_docx and thumb_docx.exists():
            try:
                insert_image_from_docx(ws, thumb_docx, img_fname, row_num, tmp_dir)
            except Exception as e:
                ws.cell(row=row_num, column=1, value=f"(no thumbnail: {e})")
                ws.row_dimensions[row_num].height = 50
        else:
            ws.cell(row=row_num, column=1, value="(source file not found)")
            ws.row_dimensions[row_num].height = 50

        ws.cell(row=row_num, column=2,  value=word_doc)
        ws.cell(row=row_num, column=3,  value=img_fname)
        ws.cell(row=row_num, column=4,  value=picture_id)
        ws.cell(row=row_num, column=5,  value=in_excel_label)
        ws.cell(row=row_num, column=6,  value=row_data["file_name"])
        ws.cell(row=row_num, column=7,  value=row_data["topic"])
        ws.cell(row=row_num, column=8,  value=row_data["section"])
        ws.cell(row=row_num, column=9,  value=row_data["page_number"])

        c_orig = ws.cell(row=row_num, column=10, value=orig_alt)
        c_orig.alignment = wrap_align
        ws.cell(row=row_num, column=11, value=len(orig_alt) if orig_alt else 0)

        c_appl = ws.cell(row=row_num, column=12, value=applied_alt)
        c_appl.alignment = wrap_align
        ws.cell(row=row_num, column=13, value=len(applied_alt) if applied_alt else 0)

        ws.cell(row=row_num, column=14, value=row_data["notes"]).alignment = wrap_align

        # O: Status with color coding
        s_cell = ws.cell(row=row_num, column=15, value=status)
        if "OK" in status:
            s_cell.font = Font(bold=True, color=GREEN_FONT)
        elif status == "No Alt Text in Excel File":
            s_cell.font = Font(bold=True, color=RED_FONT)
        else:
            s_cell.font = Font(bold=True, color=ORANGE_FONT)

        date_cell = ws.cell(row=row_num, column=16, value=run_date)
        date_cell.number_format = "mm-dd-yyyy"
        # Column Q (Reviewer Notes) left blank for manual entry

    # ── Summary sheet ─────────────────────────────────────────────────────────
    sw = wb.create_sheet("Summary")
    sw["A1"] = "Alt Text Application Summary"
    sw["A1"].font = Font(bold=True, size=14)

    total_ok      = sum(1 for r in report_rows if "OK" in r["status"])
    total_missing = sum(1 for r in report_rows if r["status"] == "No Alt Text in Excel File")
    total_review  = sum(1 for r in report_rows
                        if "OK" not in r["status"] and r["status"] != "No Alt Text in Excel File")

    sw["A3"] = "Total images in report"
    sw["B3"] = len(report_rows)
    sw["A4"] = "Alt text applied (OK)"
    sw["B4"] = total_ok
    sw["B4"].font = Font(color=GREEN_FONT, bold=True)
    sw["A5"] = "Images missing Excel alt text"
    sw["B5"] = total_missing
    sw["B5"].font = Font(color=RED_FONT, bold=True)
    sw["A6"] = "Needs review"
    sw["B6"] = total_review
    sw["B6"].font = Font(color=ORANGE_FONT, bold=True)

    sw["A8"] = "Results by Word Document"
    sw["A8"].font = Font(bold=True, size=12)
    sw["B8"] = "Applied"
    sw["C8"] = "Total Images"
    sw["D8"] = "Missing Alt Text"
    sw["A8"].font = sw["B8"].font = sw["C8"].font = sw["D8"].font = Font(bold=True)

    processed_docs = sorted(set(r["word_doc"] for r in report_rows))
    for i, doc in enumerate(processed_docs, start=9):
        sw.cell(row=i, column=1, value=doc)
        doc_rows  = [r for r in report_rows if r["word_doc"] == doc]
        ok_count  = sum(1 for r in doc_rows if "OK" in r["status"])
        miss_count = sum(1 for r in doc_rows if r["status"] == "No Alt Text in Excel File")
        sw.cell(row=i, column=2, value=ok_count).font = Font(color=GREEN_FONT)
        sw.cell(row=i, column=3, value=len(doc_rows))
        if miss_count:
            mc = sw.cell(row=i, column=4, value=miss_count)
            mc.font = Font(color=RED_FONT)

    sw.column_dimensions["A"].width = 42
    sw.column_dimensions["B"].width = 12
    sw.column_dimensions["C"].width = 14
    sw.column_dimensions["D"].width = 18

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes    = "B2"

    wb.save(REPORT_PATH)

print(f"\n✅ QA report saved to:\n   {REPORT_PATH}")
print(f"\n🎉 Done! Processed Word files saved to:\n   {OUTPUT_FOLDER}")
