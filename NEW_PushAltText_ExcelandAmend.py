# Applies alt text from a SINGLE Excel file to MULTIPLE Word documents,
# Writes the results BACK to the same Excel file
#
# Created by Brenda Gardner, March 24, 2026
#
# KEY DIFFERENCES from PushAltText_Excel_to_MultipleWord.py:
#   - Results are written back to the ORIGINAL Excel file (Columns Z, AB, AC, AD)
#   - Rows where Column Z already has a value are SKIPPED (not overwritten)
#   - Word images with no matching Excel row get a NEW ROW appended to Excel
#   - Output Word files are saved as <OriginalName>_withAltText.docx
#   - A timestamped backup of the Excel is created before any changes
#
# HOW IT WORKS:
#   1. Creates a timestamped backup of the Excel file
#   2. Reads the Excel and groups rows by source file name (Column J)
#      — skips rows where Column Z already has a status (already processed)
#   3. Scans WORD_FOLDER for all .docx files
#   4. Uses NAME_MAPPING to connect Excel source names to actual Word filenames
#   5. For each matched Word doc: applies alt text by Picture ID, saves output
#      as <original>_withAltText.docx
#   6. Writes status back to Excel: Column Z (Status), AB (In Word?),
#      AC (In Excel?), AD (Date of Completion)
#   7. Appends new rows for Word images that have no matching row in Excel
#
# TO CUSTOMIZE:
#   - Update NAME_MAPPING when file names change (see instructions there)
#   - Update EXCEL_PATH and WORD_FOLDER paths as needed


import zipfile
import os
import sys
import shutil
import tempfile
import difflib
from collections import defaultdict
from pathlib import Path
import datetime

import re
from openpyxl.utils import column_index_from_string, get_column_letter
from lxml import etree
from docx import Document

sys.path.insert(0, str(Path(__file__).parent))
from helpers import alttext_utils as utils

# ====================================================================
# CONFIG — EDIT THIS SECTION BEFORE RUNNING
# ====================================================================

EXCEL_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/0920 Information/920 2025 Lecture Notes/Module 6/0920_Mod6_AltText_ADDED.xlsx"
)

WORD_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/0920 Information/920 2025 Lecture Notes/Module 6"
)

# Output folder for Word docs with alt text applied (created automatically)
OUTPUT_FOLDER = WORD_FOLDER / "0920_Mod6_AltText_ADDED"

# Output Excel — original is never modified; results are written to this new file
EXCEL_OUTPUT_PATH = EXCEL_PATH.parent / (EXCEL_PATH.stem + "_DONE" + EXCEL_PATH.suffix)

# ── NAME MAPPING ──────────────────────────────────────────────────────────────
# Only needed when Excel Column J names DIFFER from the actual Word filenames.
# ─────────────────────────────────────────────────────────────────────────────
NAME_MAPPING = {
    # Add entries here only if Column J names differ from Word filenames.
    # "Section 2.1 OER Trigonometry 2023.docx": "Section 2.1 OER Trigonometry 2026.docx",
    # Leave empty (as shown) when names match exactly.
}

# Excel column letters — matches Master 30-column schema
PICTURE_ID_COLUMN  = column_index_from_string("C")   # col  3/C — Picture ID
TOPIC_COLUMN       = column_index_from_string("F")   # col  6/F — Image Topic
CONTEXT_COLUMN     = column_index_from_string("G")   # col  7/G — Image Context  ← written for unfound images
PAGE_NUMBER_COLUMN = column_index_from_string("H")   # col  8/H — Page #
SECTION_COLUMN     = column_index_from_string("I")   # col  9/I — Section/Module/Figure #  ← written for unfound images
FILE_NAME_COLUMN   = column_index_from_string("J")   # col 10/I — Filename (Hyperlinked) or Webpage Name  (old col I — Source document name)
ALT_TEXT_COLUMN    = column_index_from_string("W")   # col 23/W — New/Suggested Alt Text                  (old col T — Suggested Alt Text)
NOTES_COLUMN       = column_index_from_string("Y")   # col 25/Y — New/Suggested Alt Text Notes            (old col V — Notes)

# Columns written back into the Excel file
STATUS_COLUMN   = column_index_from_string("Z")   # col 26 — Status
IN_WORD_COLUMN  = column_index_from_string("AB")  # col 28 — In Word?
IN_EXCEL_COLUMN = column_index_from_string("AC")  # col 29 — In Excel?
DATE_COLUMN     = column_index_from_string("AD")  # col 30 — Date of Completion

# (Font color coding is not used — cell values are written directly into the
#  zip XML to preserve in-cell images; openpyxl styling is not applied.)


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


def _para_text(para_el, ns) -> str:
    """Concatenate all text runs inside a paragraph element."""
    return "".join(t.text or "" for t in para_el.findall(".//w:t", ns)).strip()


def scan_word_doc_for_images(docx_path: Path) -> list:
    """
    Return a list of dicts for every image in the document:
        picture_id, image_filename, original_alt, context, figure_ref, page_num

    Uses the same utils functions as WORD_to_AI.py for accurate results:
      - utils.build_rid_page_map()    → page label like "CA2-04" (Column H)
      - utils.build_rid_figure_map()  → figure number like "2.1.1" (Column I)
      - utils.build_context_snippet() → surrounding paragraph text (Column G)
      - utils.extract_docpr_tags_with_rid() → bridges rId <-> picture_id/filename
      - utils.locate_all_image_instances()  → rId -> paragraph indices
      - utils.read_paragraph_texts()        → plain text of every paragraph
    """
    images = []
    try:
        doc = Document(str(docx_path))
    except Exception as e:
        print(f"    ⚠ Could not open {docx_path.name}: {e}")
        return images

    try:
        paragraph_texts = utils.read_paragraph_texts(doc)
        rid_para_map    = utils.locate_all_image_instances(doc)       # {rId: [para_indices]}
        rid_page_map    = utils.build_rid_page_map(docx_path)         # {rId: "CA2-04"}
        rid_figure_map  = utils.build_rid_figure_map(docx_path)       # {rId: "2.1.1"}
        docpr_map       = utils.extract_docpr_tags_with_rid(docx_path) # {rId: {picture_id, image_file}}
    except Exception as e:
        print(f"    ⚠ Could not extract image data from {docx_path.name}: {e}")
        return images

    for rid, info in docpr_map.items():
        pid      = _norm_pid(info.get("picture_id", ""))
        img_file = info.get("image_file", "")

        page_num     = rid_page_map.get(rid, "")
        figure_label = rid_figure_map.get(rid, "")

        para_indices = rid_para_map.get(rid, [])
        p_idx        = para_indices[0] if para_indices else -1
        context      = (utils.build_context_snippet(paragraph_texts, p_idx, window=1)
                        if p_idx >= 0 else "")

        images.append({
            "picture_id":     pid,
            "image_filename": img_file,
            "original_alt":   "",   # not needed downstream; alt text comes from Excel
            "context":        context,
            "figure_ref":     figure_label,
            "page_num":       page_num,
        })
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
    Saves result to output_path (with _withAltText.docx suffix).
    Returns: (word_action_log, original_alt_text_map, word_images, filename_to_pids)
    """
    ns = _get_ns()
    word_action_log       = {}
    original_alt_text_map = defaultdict(list)
    word_images           = set()
    matched_filenames     = set()
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

            pid         = _norm_pid(docPr.get("id"))
            excel_match = alt_by_pid.get(pid)
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
                updated += 1
                continue

            if alt_text:
                docPr.attrib.pop("title", None)
                docPr.set("descr", alt_text)
                word_action_log[filename] = "Alt Text Added"
                matched_filenames.add(filename)
                updated += 1
            else:
                word_action_log[filename] = "No Change"
                matched_filenames.add(filename)

        # Mark images present in Word but not in Excel
        for fname in word_images:
            if fname not in matched_filenames and word_action_log.get(fname) is None:
                word_action_log[fname] = "Not in Excel"

        # ── Save updated Word file with _withAltText suffix ────────────────
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


def suggest_fuzzy_matches(excel_source_names: list, actual_word_files: list) -> dict:
    """
    Compare Excel source names to actual Word filenames and suggest close matches.
    These are printed as hints — nothing is auto-applied.
    """
    suggestions = {}
    actual_lower = {f.lower(): f for f in actual_word_files}
    for source_name in excel_source_names:
        if source_name in NAME_MAPPING:
            continue
        candidates = difflib.get_close_matches(
            source_name.lower(), actual_lower.keys(), n=1, cutoff=0.5
        )
        if candidates:
            suggestions[source_name] = actual_lower[candidates[0]]
    return suggestions


# ── Excel zip write-back helpers ──────────────────────────────────────────────
# These work directly on the zip XML so that openpyxl never gets a chance to
# strip the in-cell images (xl/media/*), richData, and metadata.xml.

def _excel_date_serial(date: datetime.date) -> int:
    """Convert a Python date to an Excel serial number (days since 1899-12-30)."""
    return (date - datetime.date(1899, 12, 30)).days


def _col_num_from_ref(ref: str) -> int:
    """Return 1-based column index from a cell reference like 'AB12'."""
    m = re.match(r"([A-Za-z]+)", ref)
    if not m:
        return 0
    n = 0
    for ch in m.group(1).upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _get_or_create_cell(row_el, row_num: int, col_idx: int, XLSX_NS: str):
    """Return the <c> element for col_idx in this row, creating it in order if absent."""
    col_letter = get_column_letter(col_idx)
    ref = f"{col_letter}{row_num}"
    for c_el in row_el.findall(f"{{{XLSX_NS}}}c"):
        if c_el.get("r") == ref:
            return c_el
    # Create and insert in ascending column order
    new_c = etree.Element(f"{{{XLSX_NS}}}c")
    new_c.set("r", ref)
    children = list(row_el)
    insert_at = len(children)
    for i, child in enumerate(children):
        if _col_num_from_ref(child.get("r", "A")) > col_idx:
            insert_at = i
            break
    row_el.insert(insert_at, new_c)
    return new_c


def _set_cell(row_el, row_num: int, col_idx: int,
              val_type: str, value,
              XLSX_NS: str, date_style_idx: int, text_style_idx: int):
    """
    Write a value into a cell without touching sharedStrings.xml.
    val_type: 'str' -> inline string;  'date' -> Excel serial number with date style.
    """
    c_el = _get_or_create_cell(row_el, row_num, col_idx, XLSX_NS)
    # Clear any existing child elements
    for child in list(c_el):
        c_el.remove(child)

    if val_type == "date":
        c_el.attrib.pop("t", None)
        c_el.set("s", str(date_style_idx))
        v_el = etree.SubElement(c_el, f"{{{XLSX_NS}}}v")
        v_el.text = str(_excel_date_serial(value))
    else:
        # Inline string — avoids any changes to sharedStrings.xml
        c_el.set("t", "inlineStr")
        if "s" not in c_el.attrib:
            c_el.set("s", str(text_style_idx))
        is_el = etree.SubElement(c_el, f"{{{XLSX_NS}}}is")
        t_el  = etree.SubElement(is_el, f"{{{XLSX_NS}}}t")
        t_el.text = str(value)


def write_back_to_excel_preserving_images(
    excel_path: Path,
    row_updates: dict,    # {row_num: {col_idx: (val_type, value)}}
    new_rows_data: list,  # [{col_idx: (val_type, value)}, ...]
    date_style_idx: int = 6,   # style index with date numFmt (same as Column O)
    text_style_idx: int = 3,   # style index for regular data cells
    output_path: Path = None,  # if given, writes here instead of overwriting excel_path
):
    """
    Patch sheet1.xml inside the Excel zip and write everything else back unchanged.
    This is the ONLY safe way to modify this workbook — openpyxl.save() strips
    all in-cell images (xl/media/*), xl/richData/, and xl/metadata.xml.
    Reads from excel_path; writes to output_path (or back to excel_path if not set).
    """
    XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    actual_output = output_path if output_path else excel_path
    tmp_path = actual_output.with_suffix(".tmp.xlsx")

    with zipfile.ZipFile(excel_path, "r") as zin:
        sheet_bytes = zin.read("xl/worksheets/sheet1.xml")
        root        = etree.fromstring(sheet_bytes)
        sheet_data  = root.find(f"{{{XLSX_NS}}}sheetData")

        # Build row_num -> <row> element map
        row_map = {int(r.get("r", 0)): r
                   for r in sheet_data.findall(f"{{{XLSX_NS}}}row")}

        # ── Update existing rows ───────────────────────────────────────────
        for row_num, col_updates in row_updates.items():
            row_el = row_map.get(row_num)
            if row_el is None:
                continue
            for col_idx, (val_type, value) in col_updates.items():
                _set_cell(row_el, row_num, col_idx, val_type, value,
                          XLSX_NS, date_style_idx, text_style_idx)

        # ── Append new rows ────────────────────────────────────────────────
        if new_rows_data:
            next_row_num = (max(row_map.keys()) + 1) if row_map else 2
            for row_data in new_rows_data:
                row_el = etree.SubElement(sheet_data, f"{{{XLSX_NS}}}row")
                row_el.set("r", str(next_row_num))
                for col_idx, (val_type, value) in sorted(row_data.items()):
                    _set_cell(row_el, next_row_num, col_idx, val_type, value,
                              XLSX_NS, date_style_idx, text_style_idx)
                next_row_num += 1

        modified_sheet = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # Write a new zip: modified sheet1.xml + everything else untouched
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/worksheets/sheet1.xml":
                    zout.writestr(item, modified_sheet)
                else:
                    zout.writestr(item, zin.read(item.filename))

    # Atomically place output file
    shutil.move(str(tmp_path), str(actual_output))


def _read_excel_data(excel_path: Path) -> list:
    """
    Read all data rows (row 2 onward) from sheet1.xml directly from the xlsx zip.
    Returns a list of dicts, one per row:  {col_idx: cell_value, ..., 'row_num': N}

    Handles:
      - shared strings  (t="s")
      - inline strings  (t="inlineStr")
      - plain numeric / date values
    No openpyxl.load_workbook() call — avoids the IndexError that newer
    Excel files trigger when openpyxl tries to bind style IDs it doesn't know.
    """
    XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

    with zipfile.ZipFile(excel_path, "r") as z:
        sheet_bytes = z.read("xl/worksheets/sheet1.xml")

        # Load shared strings table (may not exist in all workbooks)
        shared_strings = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_root = etree.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_root.findall(f"{{{XLSX_NS}}}si"):
                # Concatenate all <t> runs inside this shared-string item
                text = "".join(t.text or "" for t in si.findall(f".//{{{XLSX_NS}}}t"))
                shared_strings.append(text)

    root       = etree.fromstring(sheet_bytes)
    sheet_data = root.find(f"{{{XLSX_NS}}}sheetData")

    rows_data = []
    for row_el in sheet_data.findall(f"{{{XLSX_NS}}}row"):
        row_num = int(row_el.get("r", 0))
        if row_num < 2:          # skip the header row
            continue

        row_dict = {"row_num": row_num}

        for c_el in row_el.findall(f"{{{XLSX_NS}}}c"):
            ref    = c_el.get("r", "")
            t_attr = c_el.get("t", "")  # cell type attribute

            col_match = re.match(r"([A-Za-z]+)", ref)
            if not col_match:
                continue
            col_idx = column_index_from_string(col_match.group(1).upper())

            value = ""
            if t_attr == "s":
                # Shared string index
                v_el = c_el.find(f"{{{XLSX_NS}}}v")
                if v_el is not None and v_el.text is not None:
                    idx = int(v_el.text)
                    if 0 <= idx < len(shared_strings):
                        value = shared_strings[idx]
            elif t_attr == "inlineStr":
                is_el = c_el.find(f"{{{XLSX_NS}}}is")
                if is_el is not None:
                    value = "".join(
                        t.text or ""
                        for t in is_el.findall(f".//{{{XLSX_NS}}}t")
                    )
            else:
                # Numeric, date, boolean, formula result, or empty
                v_el = c_el.find(f"{{{XLSX_NS}}}v")
                if v_el is not None and v_el.text is not None:
                    value = v_el.text

            row_dict[col_idx] = value

        rows_data.append(row_dict)

    return rows_data


# ====================================================================
# STEP 0 — ANNOUNCE OUTPUT PATHS
# (No backup needed — the original Excel is read-only; results go to _DONE)
# ====================================================================

print(f"\n📂 Input Excel:   {EXCEL_PATH.name}")
print(f"   Output Excel:  {EXCEL_OUTPUT_PATH.name}  ← original is never modified")


# ====================================================================
# STEP 1 — READ EXCEL  (group records by source file name, Column J)
# ====================================================================

print(f"\n📊 Loading Excel: {EXCEL_PATH.name}")
# Read directly from the xlsx zip XML — no openpyxl.load_workbook() call.
# This avoids the IndexError that newer Excel files (with rich-data style IDs)
# trigger in openpyxl, and produces a file you can open and edit normally.
rows_data = _read_excel_data(EXCEL_PATH)

# excel_groups: source_name (Column J) -> {picture_id -> record}
excel_groups = defaultdict(dict)

# row_num_by_key: (source_name, picture_id) -> excel row number (for writing back)
row_num_by_key = {}

# already_done: (source_name, picture_id) pairs where Column Z is non-empty
# — these rows are skipped so existing statuses are never overwritten
already_done = set()

for row_dict in rows_data:
    row_num    = row_dict["row_num"]
    picture_id = _norm_pid(row_dict.get(PICTURE_ID_COLUMN, ""))
    alt_text   = safe_str(row_dict.get(ALT_TEXT_COLUMN, ""))
    file_name  = safe_str(row_dict.get(FILE_NAME_COLUMN, ""))
    section    = safe_str(row_dict.get(SECTION_COLUMN, ""))
    page_num   = safe_str(row_dict.get(PAGE_NUMBER_COLUMN, ""))
    notes      = safe_str(row_dict.get(NOTES_COLUMN, ""))
    topic      = safe_str(row_dict.get(TOPIC_COLUMN, ""))
    image_name = safe_str(row_dict.get(2, ""))              # Column B — image filename
    status_val = safe_str(row_dict.get(STATUS_COLUMN, ""))  # Column Z — existing status

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
        "row_num":     row_num,
    }

    if picture_id and file_name:
        excel_groups[file_name][picture_id] = rec
        key = (file_name, picture_id)
        row_num_by_key[key] = row_num
        if status_val:
            already_done.add(key)

total_records = sum(len(v) for v in excel_groups.values())
print(f"  ✅ Found {len(excel_groups)} source files, {total_records} total records")
print(f"  ℹ️  {len(already_done)} row(s) already have a status in Column Z — will skip")
for src_name, recs in sorted(excel_groups.items()):
    print(f"     • {src_name}  ({len(recs)} images)")

# Global picture_id → (file_name, picture_id) key, built across ALL source files.
# Used in Step 3 so a Word image whose picture_id already exists somewhere in
# the Excel never gets an extra new row — we update the existing row instead.
global_pid_to_key = {}
for (file_name, pid) in row_num_by_key:
    if pid and pid not in global_pid_to_key:
        global_pid_to_key[pid] = (file_name, pid)


# ====================================================================
# STEP 2 — SCAN WORD FOLDER
# ====================================================================

word_docs = sorted(d for d in WORD_FOLDER.glob("*.docx") if not d.name.startswith("~$"))
print(f"\n📁 Found {len(word_docs)} Word docs in: {WORD_FOLDER.name}")

# Reverse mapping: actual Word filename -> Excel source name (Column J value)
reverse_mapping = {v: k for k, v in NAME_MAPPING.items()}

# Check for unmapped Excel names and suggest fuzzy matches
actual_word_names = [d.name for d in word_docs]
suggestions = suggest_fuzzy_matches(list(excel_groups.keys()), actual_word_names)

if suggestions:
    print()
    print("💡 FUZZY MATCH SUGGESTIONS (not auto-applied):")
    print("   If these look correct, copy them into the NAME_MAPPING section above.")
    for src, suggested in suggestions.items():
        print(f"   Excel Column J: '{src}'")
        print(f"   Likely match:   '{suggested}'")
        print()


# ====================================================================
# STEP 3 — PROCESS EACH WORD DOCUMENT
# ====================================================================

OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

run_date = datetime.date.today()

# Collect write-back results keyed by (excel_source_name, picture_id)
results_by_key = {}

# New rows to append for Word images that have no matching Excel row
new_rows = []

for docx_path in word_docs:
    doc_name = docx_path.name
    print(f"\n📄 {doc_name}")

    # ── Look up this Word doc in the NAME_MAPPING ──────────────────────────
    excel_source_name = reverse_mapping.get(doc_name)
    # If no NAME_MAPPING entry, try a direct match: Word filename == Column J name.
    # This handles the common case where files haven't been renamed at all.
    if not excel_source_name and doc_name in excel_groups:
        excel_source_name = doc_name
    alt_by_pid = excel_groups.get(excel_source_name, {}) if excel_source_name else {}

    # ── No Excel data for this Word doc ───────────────────────────────────
    if not alt_by_pid:
        print(f"  ⚠ No NAME_MAPPING entry — scanning for images...")
        images_in_doc = scan_word_doc_for_images(docx_path)

        if not images_in_doc:
            print(f"  ✔ No images found — skipping")
            continue

        print(f"  ⚠ Found {len(images_in_doc)} image(s) with no alt text in Excel "
              f"→ will append new rows to Excel")
        for img in images_in_doc:
            new_rows.append({
                "image_filename": img["image_filename"],
                "picture_id":     img["picture_id"],
                "file_name":      doc_name,
                "context":        img["context"],
                "figure_ref":     img["figure_ref"],
                "page_num":       img.get("page_num", ""),
                "status":         "No Alt Text in Excel File",
                "in_word":        "Yes",
                "in_excel":       "No",
                # no date — alt text has not yet been applied
            })
        continue

    # ── Excel data found → apply alt text to Word doc ─────────────────────
    print(f"  🔗 Excel source: '{excel_source_name}'  ({len(alt_by_pid)} records)")

    # Scan doc for context BEFORE applying alt text so "Not in Excel" images
    # get their surrounding paragraph text and figure reference populated.
    context_by_pid = {img["picture_id"]: img
                      for img in scan_word_doc_for_images(docx_path)}

    # Build output path: insert _withAltText before the .docx extension
    output_name = docx_path.stem + "_withAltText.docx"
    output_path = OUTPUT_FOLDER / output_name

    word_action_log, orig_alt_map, word_images, fname_to_pids = \
        apply_alt_text_to_doc(docx_path, alt_by_pid, output_path)

    # ── Map results back to Excel rows ────────────────────────────────────
    for img_filename in sorted(word_images):
        pids     = fname_to_pids.get(img_filename, [])
        rec      = {}
        pid_used = ""
        for pid in pids:
            maybe = alt_by_pid.get(pid)
            if maybe:
                rec      = maybe
                pid_used = pid
                break

        action = word_action_log.get(img_filename, "No Change")

        if action == "Alt Text Added":
            status  = "OK"
            in_word = "Yes"
        elif action == "Marked as Decorative":
            status  = "OK — Marked as Decorative"
            in_word = "Yes"
        elif action == "Modify This Image":
            status  = "Review — Modify This Image"
            in_word = "No"
        elif action == "Not in Excel":
            # Check whether this picture_id already exists anywhere in the Excel.
            # If it does, write a review status to that existing row — no new row.
            # Only append a new row when the picture_id is truly absent from Excel.
            existing_key = None
            for pid in pids:
                if pid in global_pid_to_key:
                    existing_key = global_pid_to_key[pid]
                    break

            if existing_key:
                # Picture ID found in Excel — update that row's Column Z instead
                results_by_key[existing_key] = {
                    "status":   "Review Original Alt Text - No New Alt Text Applied",
                    "in_word":  "Yes",
                    "in_excel": "Yes",
                    "date":     None,   # no date — no alt text was applied
                }
            else:
                # Truly not in Excel at all — gather context and append a new row
                ctx = {}
                for pid in pids:
                    if pid in context_by_pid:
                        ctx = context_by_pid[pid]
                        break
                new_rows.append({
                    "image_filename": img_filename,
                    "picture_id":     ", ".join(pids) if pids else "",
                    "file_name":      excel_source_name,
                    "context":        ctx.get("context", ""),
                    "figure_ref":     ctx.get("figure_ref", ""),
                    "page_num":       ctx.get("page_num", ""),
                    "status":         "No Alt Text in Excel File",
                    "in_word":        "Yes",
                    "in_excel":       "No",
                    # no date — alt text has not yet been applied
                })
            continue
        elif action == "No Change":
            status  = "Review Original Alt Text - No New Alt Text Applied"
            in_word = "No"
        else:
            status  = "Review"
            in_word = "No"

        file_name_key = rec.get("file_name", excel_source_name or "")
        key = (file_name_key, pid_used)

        results_by_key[key] = {
            "status":   status,
            "in_word":  in_word,
            "in_excel": "Yes",
            "date":     run_date,
        }

print(f"\n✅ All Word documents processed.")
print(f"   Results collected for {len(results_by_key)} images  |  "
      f"{len(new_rows)} new rows to append")


# ====================================================================
# STEP 4 — WRITE RESULTS BACK TO THE ORIGINAL EXCEL FILE
#
# IMPORTANT: We do NOT use openpyxl to save here. openpyxl.save() strips
# all in-cell images (xl/media/*), xl/richData/, and xl/metadata.xml from
# the zip — making every image in Column A disappear.
#
# Instead, write_back_to_excel_preserving_images() edits sheet1.xml
# directly inside the zip and copies every other file through unchanged.
# ====================================================================

print(f"\n📊 Writing results back to Excel...")

updated_count  = 0
skipped_count  = 0
appended_count = 0

# Build the dict of row-level updates for existing rows
xl_row_updates = {}

for key, result in results_by_key.items():
    if key in already_done:
        skipped_count += 1
        continue
    row_num = row_num_by_key.get(key)
    if not row_num:
        continue
    col_updates = {
        STATUS_COLUMN:   ("str",  result["status"]),
        IN_WORD_COLUMN:  ("str",  result["in_word"]),
        IN_EXCEL_COLUMN: ("str",  result["in_excel"]),
    }
    # Column AD (Date of Completion) is only stamped when alt text was actually
    # applied — i.e. status is "OK" or "OK — Marked as Decorative".
    # "Review —..." rows leave the date cell blank.
    if result["status"].startswith("OK"):
        col_updates[DATE_COLUMN] = ("date", result["date"])
    xl_row_updates[row_num] = col_updates
    updated_count += 1

# Build new-row data for Word images that had no Excel row
xl_new_rows = []
for nr in new_rows:
    xl_new_rows.append({
        2:                   ("str",  nr["image_filename"]),
        PICTURE_ID_COLUMN:   ("str",  nr["picture_id"]),
        CONTEXT_COLUMN:      ("str",  nr.get("context", "")),    # Col G — surrounding text
        PAGE_NUMBER_COLUMN:  ("str",  nr.get("page_num", "")),   # Col H — page number
        FILE_NAME_COLUMN:    ("str",  nr["file_name"]),
        SECTION_COLUMN:      ("str",  nr.get("figure_ref", "")), # Col I — figure ref
        STATUS_COLUMN:       ("str",  nr["status"]),
        IN_WORD_COLUMN:      ("str",  nr["in_word"]),
        IN_EXCEL_COLUMN:     ("str",  nr["in_excel"]),
        # DATE_COLUMN intentionally omitted — no alt text was applied to these rows
    })
    appended_count += 1

# Read from original, write to _DONE — all images and rich data are preserved
write_back_to_excel_preserving_images(
    EXCEL_PATH,
    xl_row_updates,
    xl_new_rows,
    date_style_idx=6,   # numFmtId=16 ("d-mmm") — same style used in Column O
    text_style_idx=3,   # standard data-cell style used throughout the sheet
    output_path=EXCEL_OUTPUT_PATH,
)

print(f"\n✅ Excel saved: {EXCEL_OUTPUT_PATH.name}")
print(f"   • {updated_count} row(s) updated  (Columns Z, AB, AC; Column AD only when alt text applied)")
print(f"   • {skipped_count} row(s) skipped  (Column Z already had a status)")
print(f"   • {appended_count} new row(s) appended  (Word images not in Excel)")
print(f"\n🎉 Done!")
print(f"   Word files saved to: {OUTPUT_FOLDER}")
print(f"   Excel output:        {EXCEL_OUTPUT_PATH.name}")
