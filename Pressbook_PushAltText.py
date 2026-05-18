# Pushes Alt Text in an Excel file to Pressbooks page
# Created by Brenda Gardner, May 17, 2026
#
# TO CUSTOMIZE:
# Step 1. Update configurations
# Step 2. Set ABACUS_API_KEY as an environment variable, or paste directly below.
# Step 3. Update EXCEL_PATH
#
#WORKFLOW
#  Step 1 → PressBooks_Scrape_and_AI.py → Excel file for content experts
#  Step 2 → Expert fills Column W (New/Suggested Alt Text) in Excel
#  Step 3 → Run THIS script → alt text pushed to Pressbooks live site
"""
PressBooks_Push_AltText.py
===========================
Script 2 of 2 in the Pressbooks Accessibility Workflow.

WHAT THIS SCRIPT DOES
----------------------
1. Reads the Excel file produced by PressBooks_Scrape_and_AI.py.
2. For every row where Column W (New/Suggested Alt Text) is filled in
   by a content expert, it locates the matching image in the Pressbooks
   media library (matched by filename from Column B / URL from Column C).
3. Updates the alt text on Pressbooks via the WordPress REST API.
4. Writes the result (OK / Skipped / Error) back into Column Y (Status)
   and Column AA (Date of Completion) in a NEW output Excel file so the
   original review file is never overwritten.

COLUMN W EMPTY BEHAVIOUR
-------------------------
Every row always receives a status in Column Z so the reviewer can see
exactly what happened to each image:
  Column V filled in  → alt text is pushed; Column Y = "OK" or "Error: ..."
  Column V empty      → Column Y = "Skipped - check original alt text is not empty"
  Already up-to-date  → Column Y = "Skipped — already up-to-date"
  Not found in library→ Column Y = "Not Found in Media Library — update manually"

WORKFLOW
--------
  Step 1 → PressBooks_Scrape_and_AI.py → Excel file for content experts
  Step 2 → Expert fills Column W (New/Suggested Alt Text) in Excel
  Step 3 → Run THIS script → alt text pushed to Pressbooks live site

HOW THE PRESSBOOKS MATCH WORKS
-------------------------------
The script finds the correct Pressbooks media item using the image URL
from Column D.  It extracts the filename (e.g. "figure-3-4.png") and
searches the Pressbooks media library:

    GET /wp-json/wp/v2/media?search=figure-3-4&per_page=5

If exactly one result is returned whose source_url contains that filename,
the script updates it:

    POST /wp-json/wp/v2/media/{id}
    Body: {"alt_text": "..."}

If the image cannot be found (e.g. it was uploaded externally or uses a
non-standard URL), the row is skipped with a "Not Found in Media Library"
status and a note so the content team can update it manually.

CREDENTIALS
-----------
Create an Application Password in Pressbooks:
  Users → Your Profile → Application Passwords → Add New
Name it something like "AltTextScript" and copy the generated key.
Set it as PB_APP_PASSWORD below or as an environment variable.

REQUIREMENTS
------------
  pip install requests openpyxl lxml
"""
# ============================================================
# IMPORTS
# ============================================================

# Standard-library imports are placed here so os.getenv() is available
# inside the CONFIG block below.
import base64
import datetime
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# ============================================================
# CONFIG — EDIT THIS SECTION BEFORE RUNNING
# ============================================================

# --- Pressbooks credentials ---
PB_SITE = "https://slcc.pressbooks.pub/collegebiology1"
PB_USERNAME     = "brenda.gardner@slcc.edu"
PB_APP_PASSWORD = os.getenv("PB_APP_PASSWORD", "") # ← paste your Application Password here, or set
                       # the PB_APP_PASSWORD environment variable

# --- Excel file produced by PressBooks_Scrape_and_AI.py ---
INPUT_EXCEL_PATH = "/Users/brendagardner/Desktop/CollegeBiology_Ch1_AltText_Review_20260517_1828.xlsx"  # ← REQUIRED: full path to the Excel file
# Example:  INPUT_EXCEL_PATH = "/Users/brendagardner/Desktop/CollegeBiology_AltText_Review_20260517_0900.xlsx"

# --- Output ---
# A new Excel file is written with _DONE appended — the original is never modified.
# Leave OUTPUT_EXCEL_PATH blank to auto-generate in the same folder as the input.
OUTPUT_EXCEL_PATH = ""

# --- Column numbers (1-based) — must match PressBooks_Scrape_and_AI.py layout ---
COL_IMAGE_URL     = 3   # C — Image URL (no Picture ID column in this layout)
COL_FILENAME      = 2   # B — Image Filename
COL_CURRENT_ALT   = 15  # O — Original / Current Alt Text
COL_SUGGESTED_ALT = 22  # V — New / Suggested Alt Text (filled by reviewer)
COL_STATUS        = 25  # Y — Status written back by this script
COL_DATE_DONE     = 27  # AA — Date of Completion written back by this script

# If True, rows where the suggested alt text already matches what is live
# on Pressbooks are skipped (nothing to do).
SKIP_IF_UNCHANGED = True

# --- Safety: dry-run mode ---
DRY_RUN = False # When DRY_RUN = True, the script prints what it WOULD do but does NOT make any changes to Pressbooks.

try:
    import requests
    import openpyxl
    from lxml import etree
except ImportError as e:
    sys.exit(
        f"Missing dependency: {e}\n"
        "Run:  pip install requests openpyxl lxml"
    )

# ============================================================
# PRESSBOOKS / WORDPRESS REST API HELPERS
# ============================================================

def _auth_header() -> Dict[str, str]:
    """Return a Basic-auth header using the Application Password."""
    token = base64.b64encode(
        f"{PB_USERNAME}:{PB_APP_PASSWORD}".encode("utf-8")
    ).decode("utf-8")
    return {"Authorization": f"Basic {token}"}


def _api_get(endpoint: str, params: dict = None) -> Optional[dict]:
    """GET from the Pressbooks/WordPress REST API.  Returns parsed JSON or None."""
    url = f"{PB_SITE.rstrip('/')}/wp-json/wp/v2/{endpoint}"
    try:
        resp = requests.get(url, headers=_auth_header(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print(f"    ❌ GET {endpoint} failed: {e}")
        return None


def _api_post(endpoint: str, body: dict) -> Tuple[bool, str]:
    """
    POST to the Pressbooks/WordPress REST API.
    Returns (success: bool, message: str).
    """
    url = f"{PB_SITE.rstrip('/')}/wp-json/wp/v2/{endpoint}"
    try:
        resp = requests.post(
            url,
            headers={**_auth_header(), "Content-Type": "application/json"},
            json=body,
            timeout=30,
        )
        resp.raise_for_status()
        return True, "OK"
    except requests.HTTPError as e:
        body_text = e.response.text[:300] if e.response else ""
        return False, f"HTTP {e.response.status_code}: {body_text}"
    except Exception as e:
        return False, str(e)


# WordPress appends a size suffix to resized image filenames, e.g.
# "CarlWoeseTree-300x163.png" is a resized copy of "CarlWoeseTree.png".
# The media library stores the original name, so we must strip the suffix
# before searching or every lookup will return zero results.
_WP_SIZE_SUFFIX_RE = re.compile(r"-\d+x\d+$")


def _strip_wp_size_suffix(stem: str) -> str:
    """Remove WordPress auto-generated size suffix, e.g. -300x163 or -1024x428."""
    return _WP_SIZE_SUFFIX_RE.sub("", stem)


def find_media_by_filename(filename: str) -> Optional[int]:
    raw_stem = re.sub(r"-\d+x\d+$", "", Path(filename).stem)
    slug     = re.sub(r"[^a-z0-9]+", "-", raw_stem.lower()).strip("-")

    # Attempt 1: exact slug match
    results = _api_get("media", params={"slug": slug, "per_page": 5})
    if results and isinstance(results, list) and len(results) == 1:
        return results[0]["id"]

    # Attempt 2: scan all media and match source_url
    page = 1
    while True:
        items = _api_get("media", params={"per_page": 100, "page": page})
        if not items or not isinstance(items, list):
            break
        for item in items:
            if raw_stem.lower() in item.get("source_url", "").lower():
                return item["id"]
        if len(items) < 100:
            break
        page += 1

    return None


def find_media_by_url(image_url: str) -> Optional[int]:
    """
    Try to find a media item by filename extracted from its URL.
    Falls back to a second filename-only search if the first fails.
    """
    if not image_url:
        return None
    filename = Path(image_url.split("?")[0]).name   # strip query strings
    if not filename:
        return None
    return find_media_by_filename(filename)


def get_current_alt_text(media_id: int) -> str:
    """Fetch the current alt_text for a media item from the REST API."""
    result = _api_get(f"media/{media_id}")
    if result and isinstance(result, dict):
        return (result.get("alt_text") or "").strip()
    return ""


def update_alt_text(media_id: int, alt_text: str) -> Tuple[bool, str]:
    """
    Update the alt_text for a media item via the WordPress REST API.
    Returns (success, message).
    """
    if DRY_RUN:
        return True, "[DRY RUN] would update"
    return _api_post(f"media/{media_id}", {"alt_text": alt_text})

# ============================================================
# EXCEL WRITE-BACK
# Edits sheet1.xml directly inside the zip so that in-cell images
# (Column A thumbnails) are preserved — openpyxl.save() strips them.
# ============================================================

_XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _excel_serial(d: datetime.date) -> int:
    """Convert a Python date to an Excel serial number."""
    return (d - datetime.date(1899, 12, 30)).days


def _col_idx_from_ref(ref: str) -> int:
    """Return 1-based column index from a cell reference like 'Z12'."""
    m = re.match(r"([A-Za-z]+)", ref)
    if not m:
        return 0
    n = 0
    for ch in m.group(1).upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _col_letter(col: int) -> str:
    """1-based column index → Excel column letter(s)."""
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _get_or_create_cell(row_el, row_num: int, col_idx: int) -> object:
    """Return (or create) the <c> element for col_idx inside row_el."""
    ref = f"{_col_letter(col_idx)}{row_num}"
    for c in row_el.findall(f"{{{_XLSX_NS}}}c"):
        if c.get("r") == ref:
            return c
    new_c = etree.Element(f"{{{_XLSX_NS}}}c")
    new_c.set("r", ref)
    # Insert in column order
    children = list(row_el)
    ins = len(children)
    for i, ch in enumerate(children):
        if _col_idx_from_ref(ch.get("r", "A")) > col_idx:
            ins = i
            break
    row_el.insert(ins, new_c)
    return new_c


def _set_cell_str(row_el, row_num: int, col_idx: int, value: str,
                  text_style: int = 3) -> None:
    """Write a string value as an inline string cell (no sharedStrings change)."""
    c = _get_or_create_cell(row_el, row_num, col_idx)
    for child in list(c):
        c.remove(child)
    c.set("t", "inlineStr")
    c.set("s", str(text_style))
    is_el = etree.SubElement(c, f"{{{_XLSX_NS}}}is")
    t_el  = etree.SubElement(is_el, f"{{{_XLSX_NS}}}t")
    t_el.text = str(value)


def _set_cell_date(row_el, row_num: int, col_idx: int,
                   date: datetime.date, date_style: int = 6) -> None:
    """Write a date as an Excel serial number with date format style."""
    c = _get_or_create_cell(row_el, row_num, col_idx)
    for child in list(c):
        c.remove(child)
    c.attrib.pop("t", None)
    c.set("s", str(date_style))
    v = etree.SubElement(c, f"{{{_XLSX_NS}}}v")
    v.text = str(_excel_serial(date))


def write_results_to_excel(
    input_path: Path,
    output_path: Path,
    row_updates: Dict[int, Dict],  # {excel_row_num: {"status": str, "date": date|None}}
) -> None:
    """
    Write a new Excel file to output_path by reading input_path in a single
    pass and copying every zip entry byte-for-byte — except sheet1.xml, which
    gets the status/date cells patched in.

    This preserves all in-cell images (xl/media/*, xl/richData/, xl/metadata.xml)
    because those entries are never touched — they flow straight from the source
    zip to the destination zip unchanged.  openpyxl.save() is never called, so
    it cannot strip the rich-data embedding.

    The input file is never modified.
    """
    tmp_path = output_path.with_suffix(".tmp.xlsx")

    # Read from the original, write to a fresh tmp zip in one pass
    with zipfile.ZipFile(input_path, "r") as zin:
        sheet_bytes = zin.read("xl/worksheets/sheet1.xml")
        root        = etree.fromstring(sheet_bytes)
        sheet_data  = root.find(f"{{{_XLSX_NS}}}sheetData")

        row_map = {
            int(r.get("r", 0)): r
            for r in sheet_data.findall(f"{{{_XLSX_NS}}}row")
        }

        for row_num, update in row_updates.items():
            row_el = row_map.get(row_num)
            if row_el is None:
                continue
            status = update.get("status", "")
            date   = update.get("date")
            if status:
                _set_cell_str(row_el, row_num, COL_STATUS, status)
            if date:
                _set_cell_date(row_el, row_num, COL_DATE_DONE, date)

        modified_sheet = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        # Copy every entry from the source zip; swap in the patched sheet1.xml
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "xl/worksheets/sheet1.xml":
                    zout.writestr(item, modified_sheet)
                else:
                    zout.writestr(item, zin.read(item.filename))

    # Atomically place the finished file at the requested output path
    shutil.move(str(tmp_path), str(output_path))

# ============================================================
# EXCEL READER
# Reads directly from the xlsx zip XML (no openpyxl.load_workbook call)
# so that files with rich-data in-cell images don't raise IndexError.
# ============================================================

def read_excel_rows(path: Path) -> List[Dict]:
    """
    Read all data rows (row ≥ 2) from sheet1.xml.
    Returns a list of dicts: {col_idx (int): cell_value (str), 'row_num': int}.
    """
    with zipfile.ZipFile(path, "r") as z:
        sheet_bytes = z.read("xl/worksheets/sheet1.xml")

        # Build shared strings table
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_root = etree.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_root.findall(f"{{{_XLSX_NS}}}si"):
                text = "".join(t.text or "" for t in si.findall(f".//{{{_XLSX_NS}}}t"))
                shared_strings.append(text)

    root       = etree.fromstring(sheet_bytes)
    sheet_data = root.find(f"{{{_XLSX_NS}}}sheetData")
    rows_data  = []

    for row_el in sheet_data.findall(f"{{{_XLSX_NS}}}row"):
        row_num = int(row_el.get("r", 0))
        if row_num < 2:
            continue

        row_dict: Dict = {"row_num": row_num}

        for c_el in row_el.findall(f"{{{_XLSX_NS}}}c"):
            ref    = c_el.get("r", "")
            t_attr = c_el.get("t", "")
            m      = re.match(r"([A-Za-z]+)", ref)
            if not m:
                continue
            col_idx = _col_idx_from_ref(ref)

            value = ""
            if t_attr == "s":
                v = c_el.find(f"{{{_XLSX_NS}}}v")
                if v is not None and v.text is not None:
                    idx = int(v.text)
                    value = shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
            elif t_attr == "inlineStr":
                is_el = c_el.find(f"{{{_XLSX_NS}}}is")
                if is_el is not None:
                    value = "".join(t.text or "" for t in is_el.findall(f".//{{{_XLSX_NS}}}t"))
            else:
                v = c_el.find(f"{{{_XLSX_NS}}}v")
                if v is not None and v.text:
                    value = v.text

            row_dict[col_idx] = str(value).strip()

        rows_data.append(row_dict)

    return rows_data

# ============================================================
# MAIN
# ============================================================

def main():
    # --- Validate config ---
    if not INPUT_EXCEL_PATH:
        sys.exit(
            "❌  INPUT_EXCEL_PATH is not set.\n"
            "    Edit the CONFIG section at the top of this script and try again."
        )
    input_path = Path(INPUT_EXCEL_PATH).expanduser()
    if not input_path.exists():
        sys.exit(f"❌  Excel file not found: {input_path}")

    if not PB_APP_PASSWORD:
        sys.exit(
            "❌  PB_APP_PASSWORD is not set.\n"
            "    Create an Application Password in Pressbooks and paste it in the\n"
            "    CONFIG section, or set the PB_APP_PASSWORD environment variable."
        )

    output_path = (
        Path(OUTPUT_EXCEL_PATH).expanduser()
        if OUTPUT_EXCEL_PATH
        else input_path.with_name(input_path.stem + "_DONE" + input_path.suffix)
    )

    print(f"\n📊 Input Excel:   {input_path.name}")
    print(f"   Output Excel:  {output_path.name}")
    print(f"   Pressbooks:    {PB_SITE}")
    print(f"   DRY RUN:       {DRY_RUN}")
    if DRY_RUN:
        print("   ⚠️  Dry-run mode is ON — no changes will be made to Pressbooks.")
    print()

    # --- Read Excel ---
    print("📋 Reading Excel …")
    rows_data = read_excel_rows(input_path)
    print(f"   {len(rows_data)} data row(s) found.")

    # --- Process each row ---
    row_updates: Dict[int, Dict] = {}
    today = datetime.date.today()

    counts = {"updated": 0, "skipped_empty": 0, "skipped_unchanged": 0,
              "not_found": 0, "error": 0}

    for row in rows_data:
        row_num       = row["row_num"]
        image_url     = row.get(COL_IMAGE_URL,    "")
        filename      = row.get(COL_FILENAME,      "")
        current_alt   = row.get(COL_CURRENT_ALT,  "")
        suggested_alt = row.get(COL_SUGGESTED_ALT, "").strip()

        # --- Column V is empty: always write a status so the reviewer
        #     can see every row was visited, regardless of outcome ---
        if not suggested_alt:
            row_updates[row_num] = {"status": "Skipped - check original alt text is not empty"}
            counts["skipped_empty"] += 1
            continue

        # Skip rows where suggested == current (nothing to do)
        if SKIP_IF_UNCHANGED and suggested_alt == current_alt.strip():
            row_updates[row_num] = {"status": "Skipped — alt text unchanged"}
            counts["skipped_unchanged"] += 1
            continue

        print(f"  Row {row_num}: {filename}")

        # --- Find media item in Pressbooks ---
        media_id = find_media_by_url(image_url)
        if media_id is None and filename:
            media_id = find_media_by_filename(filename)

        if media_id is None:
            print(f"    ❌ Not found in media library.")
            row_updates[row_num] = {
                "status": "Not Found in Media Library — update manually",
            }
            counts["not_found"] += 1
            continue

        # --- Optionally verify current alt text live on Pressbooks ---
        if SKIP_IF_UNCHANGED:
            live_alt = get_current_alt_text(media_id)
            if live_alt == suggested_alt:
                print(f"    ⏭  Already up-to-date on Pressbooks.")
                row_updates[row_num] = {"status": "Skipped — already up-to-date"}
                counts["skipped_unchanged"] += 1
                continue

        # --- Update alt text ---
        print(f"    → New alt text: {suggested_alt[:80]}{'…' if len(suggested_alt) > 80 else ''}")
        success, msg = update_alt_text(media_id, suggested_alt)

        if success:
            status = "OK" if not DRY_RUN else "OK (Dry Run)"
            print(f"    ✅ {msg}")
            row_updates[row_num] = {"status": status, "date": today}
            counts["updated"] += 1
        else:
            print(f"    ❌ Error: {msg}")
            row_updates[row_num] = {"status": f"Error: {msg[:120]}"}
            counts["error"] += 1

    # --- Write results back to Excel ---
    print("\n📝 Writing results to Excel …")
    write_results_to_excel(input_path, output_path, row_updates)

    # --- Summary ---
    print(f"\n✅  Done!")
    print(f"   Updated on Pressbooks : {counts['updated']}")
    print(f"   Skipped (empty V)     : {counts['skipped_empty']}")
    print(f"   Skipped (unchanged)   : {counts['skipped_unchanged']}")
    print(f"   Not found in library  : {counts['not_found']}")
    print(f"   Errors                : {counts['error']}")
    print(f"\n   Results Excel saved to: {output_path}")

    if DRY_RUN:
        print(
            "\n   ⚠️  This was a DRY RUN.  To apply changes for real:"
            "\n       Set DRY_RUN = False at the top of the script and run again."
        )
    if counts["not_found"]:
        print(
            f"\n   ℹ️  {counts['not_found']} image(s) were not found in the Pressbooks"
            "\n      media library.  These are marked 'Not Found' in Column Z."
            "\n      You can update their alt text manually in the Pressbooks"
            "\n      media library under: Dashboard → Media → Library."
        )


if __name__ == "__main__":
    main()
