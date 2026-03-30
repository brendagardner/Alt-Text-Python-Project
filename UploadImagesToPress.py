#!/usr/bin/env python3
"""
UploadImagesToPress.py
======================
Upload images from a Word-to-Pressbooks workflow to the Pressbooks media
library, set alt text in one step, and rewrite the _PB.html file so every
<img> tag has the live Pressbooks URL and the correct alt attribute.

WHAT THIS SCRIPT DOES
---------------------
1. Reads the approved alt text from your Excel review file.
   Supports two Excel formats (set MATCH_MODE and ALT_TEXT_COLUMN):

   • MATCH_MODE = "pid"  — WORD_to_AI.py format
       Column C = Picture ID,  Column I = source .docx name,
       Column J = Section/Module,  Column T = approved alt text
       Use this for multi-chapter projects where image filenames repeat
       across Word files (e.g. image1.emf appears in Section 1.1 AND 1.2).
       Match key:  (source_docx_name, picture_id)

   • MATCH_MODE = "filename"  — Audit_Word_Alt_Text.py format
       Column B = Image Filename,  Column K = Section/Module,
       Column Q = approved alt text  (or whatever ALT_TEXT_COLUMN is set to)
       Use this for single-chapter projects where every image filename is
       unique.  No Word XML lookup needed; INPUT_DOCX may be left blank.
       Match key:  image filename stem  (e.g. "image5")

2. For MATCH_MODE="pid" only: reads the Word (.docx) file to build a
   Picture-ID → media-filename map, used to resolve the absolute image
   paths that pandoc embeds in the _PB.html output.

3. For each image found in the HTML:
   a. Locates the image file on disk.
      Handles both absolute paths (WordToPressbooks.py output) and
      relative paths (plain pandoc / externally produced HTML).
   b. Converts EMF files to PNG via Zamzar (300 DPI, high quality).
      WMF files are skipped — they rarely contain meaningful content.
      PNG/JPG/GIF are used as-is.
   c. Determines the Pressbooks media-library filename:
      • MATCH_MODE="filename" → keeps the original filename (already unique).
      • MATCH_MODE="pid"      → applies smart naming from Column J:
          If Column J value is unique across the Excel → use it: Ex1.2.1.png
          If Column J has duplicates → append Picture ID:  1.2_5.png
   d. Uploads the image to the Pressbooks REST API with alt text set
      atomically — no separate "open and type" step.
   e. Records the returned Pressbooks URL.

4. Rewrites the _PB.html file, replacing every local <img src="..."> with
   the live Pressbooks URL and the approved alt text already filled in.
   The updated HTML is ready to paste directly into the Pressbooks chapter
   editor.

REQUIREMENTS
------------
    pip install requests beautifulsoup4 openpyxl python-docx lxml Pillow

CREDENTIALS
-----------
Fill in the settings block below.  Never share your Application Password.
Create one in Pressbooks: Users → Your Profile → Application Passwords.
Name it something like "ImageMedia".

USAGE
-----
    python3 UploadImagesToPress.py

Or set the CONFIG block below and run from PyCharm with the green Run button.

Brenda Gardner — Created March 2026
"""

import os
import re
import base64
import tempfile
from collections import Counter
from io import BytesIO
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import openpyxl
from helpers import alttext_utils as utils

# ============================================================
# CONFIG — FILL THESE IN BEFORE RUNNING
# ============================================================

# --- Pressbooks credentials ---
PB_SITE         = "https://slcc.pressbooks.pub/quantitativereasoning"          # your Pressbooks site URL
PB_USERNAME     = "brenda.gardner@slcc.edu"              # your Pressbooks username
#PB_APP_PASSWORD = os.getenv("PB_APP_PASSWORD", "")
PB_APP_PASSWORD = "TUnW Ob0c KnOK g8Y7 dMOG ozYC" # ← set env var PB_APP_PASSWORD
                                                          #   or paste key directly here
                                                          #   (Application Password name: ImageMedia)

# --- Input files (fill in before running) ---
INPUT_DOCX  = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/Cryptography Files/19. Cryptography_withAltText.docx")   # ← path to the source .docx file
                         #   Required for MATCH_MODE="pid"; optional for "filename"
                         #   e.g. Path("/Users/brendagardner/.../Section 1.2.docx")
EXCEL_PATH  = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/Cryptography Files/103x_Cryptography_AltText_Apply.xlsx")   # ← path to the Excel review file
                         #   e.g. Path("/Users/brendagardner/.../1060_Mod1_AltText_REVIEW.xlsx")
HTML_PATH   = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/103x HTML Version/Cryptography_V1.html")   # ← path to the HTML file (absolute or relative img paths both work)
                         #   e.g. Path("/Users/brendagardner/.../Section_1.2_PB.html")

MEDIA_BASE_DIR  = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/Cryptography Files")

# --- Zamzar API key (for EMF → PNG conversion) ---
ZAMZAR_API_KEY = os.getenv("ZAMZAR", "")

# --- Site ID for this Pressbooks book (found in your media URLs after /sites/) ---
# Look at any image URL in your existing Pressbooks chapter, e.g.:
#   https://slcc.pressbooks.pub/app/uploads/sites/86/2026/02/image.png
#                                                           ^^
PB_SITE_ID  = "86"

# --- Excel format / matching mode ---
MATCH_MODE      = "filename"   # "pid"      — multi-chapter project (WORD_to_AI.py format)
                           #              key = (source_docx_name, picture_id)
                           # "filename" — single-chapter / unique filenames
                           #              (Audit_Word_Alt_Text.py format)
                           #              key = image filename stem

ALT_TEXT_COLUMN = "F"     # Column letter containing the APPROVED alt text
                           # "T" = WORD_to_AI.py format  (Column T, index 19)
                           # "Q" = Audit_Word_Alt_Text.py format  (Column Q, index 16)

SECTION_COLUMN  = "D"     # Column letter containing Section/Module name (used for
                           # smart Pressbooks filenames in MATCH_MODE="pid")
                           # "J" = WORD_to_AI.py format
                           # "K" = Audit_Word_Alt_Text.py format

# ============================================================
# EXCEL HELPERS
# ============================================================

def _col_letter_to_idx(letter: str) -> int:
    """Convert a column letter (A, B, ..., Z, AA, ...) to a 0-based index."""
    letter = letter.upper().strip()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1


def load_alt_text_lookup(
    excel_path: Path,
    match_mode: str,
    alt_text_column: str,
    section_column: str,
) -> tuple:
    """
    Build two lookup structures from the Excel review file.

    For MATCH_MODE="pid":
        key   = (source_docx_name, str(picture_id))
        value = {"alt_text": str, "section": str}
        Reads: Column C (Picture ID), Column I (source file), Column J (section),
               and whichever column ALT_TEXT_COLUMN points to.

    For MATCH_MODE="filename":
        key   = image filename stem  (e.g. "image5")
        value = {"alt_text": str, "section": str}
        Reads: Column B (Image Filename), the section column, and alt text column.

    Returns:
        lookup   — dict as described above
        j_counts — Counter of section values (used for smart naming in "pid" mode)
    """
    wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    ws = wb.active

    alt_idx     = _col_letter_to_idx(alt_text_column)   # e.g. Q→16, T→19
    section_idx = _col_letter_to_idx(section_column)    # e.g. J→9, K→10

    lookup   = {}
    j_values = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue

        # Safe column reads (row may be shorter than expected)
        def _cell(idx):
            return str(row[idx]).strip() if idx < len(row) and row[idx] is not None else ""

        alt_text = _cell(alt_idx)
        section  = _cell(section_idx)

        if match_mode == "filename":
            # Column B (index 1) holds the image filename (e.g. "image5.png")
            filename = _cell(1)
            stem = Path(filename).stem if filename else ""
            if stem:
                lookup[stem] = {"alt_text": alt_text, "section": section}
                if section:
                    j_values.append(section)

        else:  # "pid" mode
            # Column C (index 2) = Picture ID, Column I (index 8) = source docx name
            picture_id  = _cell(2)
            source_file = _cell(8)

            # Normalise Picture ID (Excel may store as float: "2.0" → "2")
            try:
                picture_id = str(int(float(picture_id)))
            except (ValueError, TypeError):
                pass

            if source_file and picture_id:
                key = (source_file, picture_id)
                lookup[key] = {"alt_text": alt_text, "section": section}
                if section:
                    j_values.append(section)

    j_counts = Counter(j_values)
    return lookup, j_counts


def smart_filename(section: str, picture_id: str, j_counts: Counter) -> str:
    """
    Return a clean filename for the Pressbooks media library (MATCH_MODE="pid").

    Rule:
      • If the Section/Module value is unique across the Excel → use as-is:
            Ex1.2.1  →  Ex1.2.1.png
      • If it appears on multiple rows → append the Picture ID:
            1.2  + pid 5  →  1.2_5.png
    The result is always unique.
    """
    if not section:
        return f"img_pid{picture_id}.png"
    base = re.sub(r'[^\w.\-]', '_', section)   # safe for web filenames
    if j_counts.get(section, 0) <= 1:
        return f"{base}.png"
    return f"{base}_{picture_id}.png"


# ============================================================
# WORD XML HELPERS  (used only in MATCH_MODE="pid")
# ============================================================

def build_media_to_pid_map(docx_path: Path) -> dict:
    """
    Return {media_filename_stem: picture_id} for every image in the docx.

    Uses alttext_utils functions that extract Picture IDs from the Word XML.
    The stem (filename without extension) is used so .emf and .png variants
    both resolve to the same Picture ID.

    Example:  {"image1": "2", "image24": "12", ...}
    """
    rels  = utils.parse_relationships(docx_path)          # rId → 'media/image1.emf'
    docpr = utils.extract_docpr_tags_with_rid(docx_path)  # rId → {picture_id, ...}

    stem_to_pid = {}
    for rid, target in rels.items():
        pid_info = docpr.get(rid, {})
        pid = pid_info.get("picture_id", "")
        try:
            pid = str(int(float(pid)))
        except (ValueError, TypeError):
            pass
        if pid:
            stem = Path(target).stem   # e.g. 'image1'
            stem_to_pid[stem] = pid

    return stem_to_pid


# ============================================================
# PRESSBOOKS (WORDPRESS REST API) HELPERS
# ============================================================

def _pb_auth_header() -> dict:
    """Basic auth header using Application Password."""
    token = base64.b64encode(
        f"{PB_USERNAME}:{PB_APP_PASSWORD}".encode("utf-8")
    ).decode("utf-8")
    return {"Authorization": f"Basic {token}"}


def upload_image_to_pressbooks(
    img_bytes: bytes,
    filename: str,
    alt_text: str,
) -> "str | None":
    """
    Upload an image to the Pressbooks media library and set alt text.

    Returns the public URL of the uploaded image, or None on failure.

    Pressbooks is built on WordPress; we use the standard WP REST API:
      POST /wp-json/wp/v2/media
    The alt_text is set in the same request via a follow-up PATCH.
    """
    media_url = f"{PB_SITE}/wp-json/wp/v2/media"
    headers = {
        **_pb_auth_header(),
        "Content-Disposition": f'attachment; filename="{filename}"',
    }

    # Detect MIME type from extension
    ext = Path(filename).suffix.lower()
    mime_map = {
        ".png":  "image/png",
        ".jpg":  "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif":  "image/gif",
        ".svg":  "image/svg+xml",
    }
    content_type = mime_map.get(ext, "image/png")

    try:
        # Step 1 — upload the file
        resp = requests.post(
            media_url,
            headers={**headers, "Content-Type": content_type},
            data=img_bytes,
            timeout=120,
        )
        resp.raise_for_status()
        media_id  = resp.json()["id"]
        image_url = resp.json()["source_url"]

        # Step 2 — set alt text via POST (WordPress REST API uses POST for updates)
        patch_resp = requests.post(
            f"{media_url}/{media_id}",
            headers={**_pb_auth_header(), "Content-Type": "application/json"},
            json={"alt_text": alt_text},
            timeout=30,
        )
        patch_resp.raise_for_status()

        print(f"  ✅ Uploaded: {filename}  →  {image_url}")
        return image_url

    except requests.HTTPError as e:
        print(f"  ❌ HTTP error uploading {filename}: {e.response.status_code} {e.response.text[:200]}")
        return None
    except Exception as e:
        print(f"  ❌ Error uploading {filename}: {e}")
        return None


# ============================================================
# HTML REWRITER
# ============================================================

def rewrite_html(html_path: Path, src_to_info: dict) -> Path:
    """
    Rewrite <img src="..."> tags in the HTML file.

    src_to_info maps the original src value (as it appears in the HTML,
    whether absolute or relative) to a dict:
        {"url": "https://slcc.pressbooks.pub/...", "alt": "..."}

    Saves the updated HTML to a new file with '_PB' appended before the
    extension (e.g. '17_Orem_Pandoc.html' → '17_Orem_Pandoc_PB.html').
    The original file is NOT modified.

    Returns the path of the new file.
    """
    html = html_path.read_text(encoding="utf-8")
    soup = BeautifulSoup(html, "html.parser")

    replaced = 0
    skipped  = 0

    for img in soup.find_all("img"):
        src = img.get("src", "")
        if src in src_to_info:
            info = src_to_info[src]
            img["src"] = info["url"]
            img["alt"] = info["alt"]
            # Clean up pandoc's inline width/height styles — Pressbooks handles sizing
            if img.get("style"):
                del img["style"]
            replaced += 1
        else:
            skipped += 1

    # Build output path:  myfile.html  →  myfile_PB.html
    # resolve() ensures the output file lands in the right folder even if
    # html_path was passed as a relative or bare filename.
    out_path = html_path.resolve().with_stem(html_path.stem + "_PB")
    out_path.write_text(str(soup), encoding="utf-8")
    print(f"\n✅ HTML rewritten: {replaced} image(s) updated, {skipped} skipped.")
    print(f"   Original:  {html_path.name}")
    print(f"   Pressbooks-ready: {out_path.name}")
    return out_path


# ============================================================
# MAIN
# ============================================================

def main():
    # --- Validate config ---
    errors = []

    if MATCH_MODE not in ("pid", "filename"):
        errors.append(f"MATCH_MODE must be 'pid' or 'filename', got: {MATCH_MODE!r}")

    if MATCH_MODE == "pid":
        if not INPUT_DOCX or not INPUT_DOCX.exists():
            errors.append(
                f"INPUT_DOCX not found: {INPUT_DOCX}\n"
                f"  (INPUT_DOCX is required for MATCH_MODE='pid')"
            )

    if not EXCEL_PATH or not EXCEL_PATH.exists():
        errors.append(f"EXCEL_PATH not found: {EXCEL_PATH}")
    if not HTML_PATH or not HTML_PATH.exists():
        errors.append(f"HTML_PATH not found: {HTML_PATH}")
    if not PB_APP_PASSWORD:
        errors.append(
            "PB_APP_PASSWORD is not set.\n"
            "  Option A: set the environment variable PB_APP_PASSWORD=your_key\n"
            "  Option B: paste the key directly into the PB_APP_PASSWORD line above."
        )
    if errors:
        for e in errors:
            print(f"❌ {e}")
        return

    print(f"📊 Alt text Excel:    {EXCEL_PATH.name}")
    print(f"🌐 HTML file:         {HTML_PATH.name}")
    print(f"🔗 Pressbooks site:   {PB_SITE}")
    print(f"📌 Match mode:        {MATCH_MODE}  (alt text column: {ALT_TEXT_COLUMN})")
    if MATCH_MODE == "pid":
        print(f"📄 Source Word file:  {INPUT_DOCX.name}")
    print()

    # 1. Load alt text lookup from Excel
    print("📋 Loading alt text lookup from Excel...")
    lookup, j_counts = load_alt_text_lookup(
        EXCEL_PATH, MATCH_MODE, ALT_TEXT_COLUMN, SECTION_COLUMN
    )
    print(f"   {len(lookup)} image(s) found in Excel.")

    # 2. Build Picture-ID → media filename map (pid mode only)
    stem_to_pid = {}
    docx_name   = ""
    if MATCH_MODE == "pid":
        print("📎 Reading Word XML for Picture ID → filename map...")
        stem_to_pid = build_media_to_pid_map(INPUT_DOCX)
        docx_name   = INPUT_DOCX.name
        print(f"   {len(stem_to_pid)} image(s) found in Word file.")

    # 3. Parse the HTML to find all <img> tags
    # Determine the base directory for resolving relative image paths.
    # Always resolve() to convert relative paths → absolute, so media_root
    # is never just "." even if HTML_PATH was given as a bare filename.
    if MEDIA_BASE_DIR and MEDIA_BASE_DIR.resolve().exists():
        media_root = MEDIA_BASE_DIR.resolve()
    else:
        media_root = HTML_PATH.resolve().parent
    print(f"📁 Looking for images in: {media_root}")
    print("🔍 Parsing HTML for <img> tags...")
    html = HTML_PATH.read_text(encoding="utf-8")
    soup = BeautifulSoup(html, "html.parser")
    img_tags = soup.find_all("img")
    print(f"   {len(img_tags)} <img> tag(s) found.")
    print()

    # 4. Process each image
    src_to_info = {}   # original src value → {url, alt}

    with tempfile.TemporaryDirectory() as tmp_str:
        tmp_dir = Path(tmp_str)

        for img in img_tags:
            src = img.get("src", "")
            if not src:
                continue

            # Resolve path — handle both absolute and relative src values
            src_path = Path(src)
            if src_path.is_absolute():
                img_path = src_path
            else:
                img_path = media_root / src_path

            if not img_path.exists():
                print(f"  ⚠️  File not found on disk, skipping: {src}")
                print(f"       (looked in: {img_path})")
                continue

            stem   = img_path.stem           # e.g. 'image5'
            suffix = img_path.suffix.lower() # e.g. '.png' or '.emf'

            # --- Skip WMF (low quality, rarely meaningful) ---
            if suffix == ".wmf":
                print(f"  ⏭️  Skipping WMF: {img_path.name}")
                continue

            # --- Look up alt text ---
            if MATCH_MODE == "filename":
                info = lookup.get(stem)
                if not info:
                    print(f"  ⚠️  No Excel row for filename '{stem}', skipping.")
                    continue
                alt_text = info["alt_text"]
                section  = info["section"]
                # Keep original filename — it is already unique in this mode
                pb_filename = img_path.name if suffix != ".emf" else f"{stem}.png"

            else:  # MATCH_MODE == "pid"
                pid = stem_to_pid.get(stem)
                if not pid:
                    print(f"  ⚠️  No Picture ID found for '{stem}', skipping.")
                    continue
                key = (docx_name, pid)
                info = lookup.get(key)
                if not info:
                    print(f"  ⚠️  No Excel row for ({docx_name}, PID={pid}), skipping.")
                    continue
                alt_text = info["alt_text"]
                section  = info["section"]
                # Apply smart naming rule using Column J/K values
                pb_filename = smart_filename(section, pid, j_counts)

            if not alt_text:
                mode_hint = f"filename='{stem}'" if MATCH_MODE == "filename" else f"PID={pid if MATCH_MODE == 'pid' else '?'}"
                print(f"  ⚠️  Alt text column {ALT_TEXT_COLUMN} is empty for {mode_hint} ({section}). Skipping upload.")
                continue

            print(f"  🖼  {img_path.name}  →  {pb_filename}")

            # --- Load image bytes ---
            img_bytes = img_path.read_bytes()

            # --- Convert vector formats ---
            if suffix == ".emf":
                if not ZAMZAR_API_KEY:
                    print(f"     ⚠️  No ZAMZAR key — cannot convert EMF. Set env var ZAMZAR=your_key.")
                    continue
                converted = utils.convert_emf_via_zamzar(img_bytes, img_path.name, ZAMZAR_API_KEY)
                if not converted:
                    print(f"     ❌ Zamzar conversion failed, skipping.")
                    continue
                img_bytes = converted

            elif suffix == ".svg":
                converted = utils.convert_svg_to_png_bytes(img_bytes)
                if not converted:
                    print(f"     ❌ SVG conversion failed, skipping.")
                    continue
                img_bytes = converted

            # --- Upload to Pressbooks ---
            url = upload_image_to_pressbooks(img_bytes, pb_filename, alt_text)
            if url:
                src_to_info[src] = {"url": url, "alt": alt_text}

    # 5. Rewrite the HTML → saves as _PB.html, original untouched
    if src_to_info:
        print()
        print(f"✏️  Writing Pressbooks-ready HTML ({len(src_to_info)} image(s) updated)...")
        pb_html_path = rewrite_html(HTML_PATH, src_to_info)
    else:
        print("\n⚠️  No images were successfully uploaded — HTML not written.")
        return

    print()
    print("Done! Next step: open the _PB.html file, copy the content inside")
    print(f"  the <body> tags, and paste into the Pressbooks chapter editor (Text/HTML mode).")
    print(f"  File: {pb_html_path}")


if __name__ == "__main__":
    main()
