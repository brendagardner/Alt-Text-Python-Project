# ====
# UNIFIED: MATCH BY PICTURE ID, APPLY ALT TEXT, AND BUILD QA REPORT
# ====
# What this script does:
# 1) Reads Excel (Picture ID, Alt Text, metadata)
# 2) Opens DOCX, snapshots ORIGINAL alt text (pre-change)
# 3) Applies Excel alt text by Picture ID (supports "N/A" and "Mark as Decorative")
# 4) Rebuilds a QA report with:
#    - Column I  : Word Alt Text (ORIGINAL, pre-change)
#    - Column J  : Applied/New Alt Text (Excel) — only if applied/changed
#    - Word Action and simplified Status: OK / Review (green/orange font)
#    - Excel-only rows: placeholder thumb + ImageName/PictureID from Excel

# To FINISH:
# 1. .emf files are not supported yet.
# 2. Still using Visual Basic to put images in cell?
# 3. FIX Output report -
#   a. Add Topic for easier search?
#   b. Verify applying to multiple images (if image repeats in Word Document)
#
# Brenda Gardner — Final merged version (Picture ID pipeline)  w/Claude help
# ====

import os
import zipfile
import tempfile
import subprocess
from io import BytesIO
from pathlib import Path
from typing import Union, Tuple

import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, column_index_from_string
from lxml import etree
from PIL import Image as PILImage

# ====
# CONFIG — EDIT THESE
# ====
DOCX_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1010 Information/1010 Lecture Notes Chapters 1-5/Math 1010 Lecture Notes Module 2.docx")
EXCEL_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1010 Information/1010 Test Files/Mod 2 Push Alt Text.xlsx")

OUTPUT_DOCX = DOCX_PATH.with_name(DOCX_PATH.stem + "_AltText.docx")
REPORT_PATH = DOCX_PATH.with_name(DOCX_PATH.stem + "_AltText_Report.xlsx")

# Excel column mapping (letters). Change ALT_TEXT_COLUMN to "P" if your alt text is in P.
PICTURE_ID_COLUMN = column_index_from_string("C")  # Excel Picture ID
ALT_TEXT_COLUMN = column_index_from_string("T")  # Excel Alt Text ("P" or "T" per your sheet)
FILE_NAME_COLUMN = column_index_from_string("J")
SECTION_COLUMN = column_index_from_string("K")
PAGE_NUMBER_COLUMN = column_index_from_string("L")
NOTES_COLUMN = column_index_from_string("V")

# Thumbnails
LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"  # adjust if needed
IMAGE_WIDTH = 150  # px


# ====
# HELPERS
# ====
def load_alt_text_excel(excel_path: Path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # Build column index map by header names (first row)
    headers = {cell.value.strip(): cell.column for cell in ws[1] if cell.value}

    def get_col(name, default_letter):
        """Get column index by header name or fallback letter."""
        if name in headers:
            return headers[name]
        return column_index_from_string(default_letter)

    # dynamic mapping
    PID_COL = get_col("Picture ID", "C")
    ALT_COL = get_col("Alt Text", "T")
    FILE_COL = get_col("Original File", "J")
    SECTION_COL = get_col("Section/Module", "K")
    PAGE_COL = get_col("Page #", "L")
    NOTES_COL = get_col("Notes", "V")

    alt_text_map = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        picture_id = str(row[PID_COL - 1]).strip() if row[PID_COL - 1] else ""
        alt_text = str(row[ALT_COL - 1]).strip() if row[ALT_COL - 1] else ""
        file_name = str(row[FILE_COL - 1]).strip() if row[FILE_COL - 1] else ""
        section = str(row[SECTION_COL - 1]).strip() if row[SECTION_COL - 1] else ""
        page_num = str(row[PAGE_COL - 1]).strip() if row[PAGE_COL - 1] else ""
        notes = str(row[NOTES_COL - 1]).strip() if row[NOTES_COL - 1] else ""

        # Use either Picture ID or File name as key
        if picture_id or file_name:
            key = picture_id or file_name
            alt_text_map[key] = {
                "picture_id": picture_id,
                "alt_text": alt_text,
                "file_name": file_name,
                "section": section,
                "page_number": page_num,
                "notes": notes,
            }

    print(f"✅ Loaded {len(alt_text_map)} Excel rows with metadata columns C–V mapped dynamically.")
    return alt_text_map


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
        ws.cell(row=row, column=1, value=f"(error: {e})")
        print(f"⚠ Could not insert {target_filename}: {e}")


def adjust_row_height(ws, row: int, img_height: int = None, min_height: int = 50):
    """Adjust row height based on image height or keep a consistent minimum."""
    if img_height and img_height > min_height:
        ws.row_dimensions[row].height = img_height * 0.75
    else:
        ws.row_dimensions[row].height = min_height


def solid_fill(argb: str) -> PatternFill:
    return PatternFill(patternType="solid", fgColor=argb)


# ====
# STEP 1 — READ EXCEL (keyed by Picture ID)
# ====
wb_src = openpyxl.load_workbook(EXCEL_PATH)
ws_src = wb_src.active

alt_by_pid = {}  # picture_id -> record
alt_by_filename = {}  # filename   -> record (fallback)

records_loaded = 0
for row in ws_src.iter_rows(min_row=2):
    picture_id = _norm_pid(row[PICTURE_ID_COLUMN - 1].value)
    alt_text = safe_str(row[ALT_TEXT_COLUMN - 1].value)
    image_name = safe_str(row[1].value)  # column B in the source Excel (1-based index)
    file_name = safe_str(row[FILE_NAME_COLUMN - 1].value)
    section = safe_str(row[SECTION_COLUMN - 1].value)
    page_num = safe_str(row[PAGE_NUMBER_COLUMN - 1].value)
    notes = safe_str(row[NOTES_COLUMN - 1].value)

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
# STEP 2 — OPEN DOCX, SNAPSHOT ORIGINAL ALT TEXTS, APPLY UPDATES
# ====
with tempfile.TemporaryDirectory() as tmpdir:
    tmp_path = Path(tmpdir)
    with zipfile.ZipFile(DOCX_PATH, "r") as zip_ref:
        zip_ref.extractall(tmp_path)

    document_xml_path = tmp_path / "word" / "document.xml"
    rels_path = tmp_path / "word" / "_rels" / "document.xml.rels"

    ns = {
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }

    tree = etree.parse(str(document_xml_path))
    root = tree.getroot()
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

    word_images = set()
    word_action_log = {}
    original_alt_text_map = {}  # filename -> list of alt texts
    matched_filenames = set()  # filenames that got a valid Excel match

    # --- SNAPSHOT ORIGINAL ALT TEXTS (before any changes)
    word_pid_to_filename = {}  # picture_id -> filename
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

        # Capture original alt text
        original_descr = docPr.get("descr", "")
        if filename not in original_alt_text_map:
            original_alt_text_map[filename] = []
        original_alt_text_map[filename].append(original_descr)

        picture_id_raw = docPr.get("id")
        picture_id = _norm_pid(picture_id_raw)
        if picture_id:
            word_pid_to_filename[picture_id] = filename

    # --- APPLY UPDATES BY PICTURE ID
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
                # Ensure prefix registered (lxml handles URIs; prefix is cosmetic)
                etree.register_namespace("a11y", A11Y_NS)

                extLst = docPr_el.find(f".//{{{A_NS}}}extLst")
                if extLst is None:
                    extLst = etree.Element(f"{{{A_NS}}}extLst")
                    # docPr must contain extLst directly (sibling of cNvPr)
                    docPr_el.append(extLst)

                # MSO usually wraps vendor extensions in <a:ext uri="...">
                # A concrete, stable URI isn't publicly documented; we can use a
                # deterministic GUID string to host the a11y child.
                # Word accepts unknown URIs as vendor extensions.
                ext = None
                for candidate in extLst.findall(f".//{{{A_NS}}}ext"):
                    # reuse an existing ext if it already contains an a11y:decorative
                    if candidate.find(f".//{{{A11Y_NS}}}decorative") is not None:
                        ext = candidate
                        break
                if ext is None:
                    ext = etree.Element(f"{{{A_NS}}}ext")
                    ext.set("uri", "{0000-0000-0000-0000-0000}")  # placeholder URI
                    extLst.append(ext)

                # Add/ensure a11y:decorative val="1"
                decorative = ext.find(f".//{{{A11Y_NS}}}decorative")
                if decorative is None:
                    decorative = etree.Element(f"{{{A11Y_NS}}}decorative")
                    ext.append(decorative)
                decorative.set("val", "1")


            mark_decorative(docPr)
            word_action_log[filename] = "Marked as Decorative"
            matched_filenames.add(filename)
            updated += 1
            continue

        # Standard alt text path
        if alt_text:
            docPr.attrib.pop("title", None)  # remove any stale title
            docPr.set("descr", alt_text)  # write new alt text
            word_action_log[filename] = "Alt Text Added"
            matched_filenames.add(filename)
            updated += 1
        else:
            # matched PID but empty value: leave as no-change
            word_action_log[filename] = "No Change"
            matched_filenames.add(filename)

    for fname in word_images:
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

# ✅ Combine all known images:
#    - those in the Word file
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
    "Thumbnail",
    "Image Name",
    "Picture ID",
    "Original File",
    "Section/Module",
    "Page #",
    "In Word?",
    "In Excel?",
    "Word Alt Text",
    "Applied/New Alt Text (Excel)",
    "Word Action",
    "Status",
    "Notes"
]

col_widths = [50, 30, 25, 25, 12, 10, 10, 10, 70, 40, 20, 18, 40]

for col, h in enumerate(headers, start=1):
    c = ws.cell(row=1, column=col, value=h)
    c.font = Font(bold=True)
    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = col_widths[col - 1]

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
            # Get Picture ID from Word for this filename
            for pid, fname in word_pid_to_filename.items():
                if fname == name:
                    rec = alt_by_pid.get(pid, {})
                    break

        file_name = rec.get("file_name", "")
        section = rec.get("section", "")
        page_num = rec.get("page_number", "")
        notes = rec.get("notes", "")
        excel_alt_val = (rec.get("alt_text") or "").strip()

        # --- Original Word alt text (Column I) ---
        alt_texts = original_alt_text_map.get(name, [])
        word_alt_text = " | ".join(alt_texts) if alt_texts else ""

        # Always show Excel alt text in Column J
        applied_alt_val = excel_alt_val
        if excel_alt_val and excel_alt_val != word_alt_text:
            new_alt_applied_count += 1

        # --- Membership flags ---
        in_word = "Yes" if name in word_images else ""
        in_excel = "Yes" if name in excel_images else ""
        if name in word_images and name in excel_images:
            in_word = in_excel = "Yes"

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

        # --- Picture ID from Excel record ---
        picture_id = rec.get("picture_id", "")

        # --- Write base metadata (Columns B–F) ---
        ws.cell(row=row_num, column=2, value=name)
        ws.cell(row=row_num, column=3, value=picture_id)
        ws.cell(row=row_num, column=4, value=file_name)
        ws.cell(row=row_num, column=5, value=section)
        ws.cell(row=row_num, column=6, value=page_num)

        # --- In Word / In Excel flags ---
        ws.cell(row=row_num, column=7, value=in_word)
        ws.cell(row=row_num, column=8, value=in_excel)

        # --- Alt Text columns ---
        ws.cell(row=row_num, column=9, value=word_alt_text).alignment = shrink_wrap_alignment
        cell_new = ws.cell(row=row_num, column=10, value=applied_alt_val)
        cell_new.alignment = shrink_wrap_alignment

        # --- Word Action + Status determination ---
        word_action = word_action_log.get(name, "No Change")
        ws.cell(row=row_num, column=11, value=word_action)

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

        s_cell = ws.cell(row=row_num, column=12, value=status)
        if status == "OK":
            s_cell.font = Font(bold=True, color=GREEN_FONT)
        elif status == "Review":
            s_cell.font = Font(bold=True, color=ORANGE_FONT)
        else:
            s_cell.font = Font(bold=True)

        ws.cell(row=row_num, column=13, value=notes).alignment = wrap_alignment
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
        val = ws.cell(row=r, column=12).value
        if val:
            status_counts[val] = status_counts.get(val, 0) + 1

    row = 12
    for status, count in sorted(status_counts.items()):
        c_status = summary_ws.cell(row=row, column=1, value=status)
        c_count = summary_ws.cell(row=row, column=2, value=count)
        if status == "OK":
            c_status.font = Font(color=GREEN_FONT, bold=True)
            c_count.font = Font(color=GREEN_FONT, bold=True)
        elif status == "Review":
            c_status.font = Font(color=ORANGE_FONT, bold=True)
            c_count.font = Font(color=ORANGE_FONT, bold=True)
        row += 1

    summary_ws.column_dimensions["A"].width = 35
    summary_ws.column_dimensions["B"].width = 15
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    # ✅ Save workbook before temp dir is deleted
    wb.save(REPORT_PATH)
    print(f"✅ Excel QA report saved to:\n{REPORT_PATH}")