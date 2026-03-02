#!/usr/bin/env python3
"""
UploadImagesToPress.py
======================
Upload images from a Word-to-Pressbooks workflow to the Pressbooks media
library, set alt text in one step, and rewrite the _PB.html file so every
<img> tag has the live Pressbooks URL and the correct alt attribute.

WHAT THIS SCRIPT DOES
---------------------
1. Reads the approved alt text from your Excel review file (Column T).
   Matches each image using:  Source filename (Column I)  +  Picture ID (Column C)
   — this combination is guaranteed unique even when filenames repeat across
   Word documents (e.g. image1.emf appears in Section 1.1 AND Exercises 1.2).

2. Reads the Word (.docx) file to build a Picture-ID → media-filename map.
   This is used to resolve the absolute image paths that pandoc embeds in
   the _PB.html output.

3. For each image found in the HTML:
   a. Locates the image file on disk (absolute path from pandoc output).
   b. Converts EMF files to PNG via Zamzar (300 DPI, high quality).
      WMF files are skipped — they rarely contain meaningful math diagrams.
      PNG/JPG/GIF are used as-is.
   c. Applies a smart filename for the Pressbooks media library:
      • If Column J (Section/Module) is unique → use it directly: Ex1.2.1.png
      • If Column J has duplicates → append Picture ID:  1.2_5.png
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

Or set INPUT_DOCX / EXCEL_PATH / HTML_PATH in the CONFIG block and run
from PyCharm with the green Run button.

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
PB_SITE       = "https://slcc.pressbooks.pub"          # your Pressbooks site URL
PB_USERNAME   = "brenda.gardner@slcc.edu"              # your Pressbooks username
PB_APP_PASSWORD = os.getenv("PB_APP_PASSWORD", "")     # ← set env var PB_APP_PASSWORD
                                                        #   or paste key directly here
                                                        #   (Application Password name: ImageMedia)

# --- Input files (fill in before running) ---
INPUT_DOCX  = Path("")   # ← path to the source .docx file, e.g.:
                         #   Path("/Users/brendagardner/.../Section 1.2.docx")
EXCEL_PATH  = Path("")   # ← path to the _AltText_REVIEW.xlsx file
HTML_PATH   = Path("")   # ← path to the _PB.html output from WordToPressbooks.py

# --- Zamzar API key (for EMF → PNG conversion) ---
ZAMZAR_API_KEY = os.getenv("ZAMZAR", "")

# --- Site ID for this Pressbooks book (found in your media URLs after /sites/) ---
# Look at any image URL in your existing HTML, e.g.:
#   https://slcc.pressbooks.pub/app/uploads/sites/86/2026/02/image.png
#                                                           ^^
PB_SITE_ID  = "86"

# ============================================================
# EXCEL HELPERS
# ============================================================

def load_alt_text_lookup(excel_path: Path) -> dict:
    """
    Build two lookup structures from the Excel review file.

    Returns:
        lookup  — dict keyed by (source_docx_name, str(picture_id))
                  value: {"alt_text": str, "section": str}
        j_counts — Counter of Column J values (used for smart naming)
    """
    wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    ws = wb.active

    lookup   = {}
    j_values = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        # Column layout (0-indexed):
        # A=0  B=1(filename)  C=2(picture_id)  I=8(source_file)
        # J=9(section)        T=19(approved alt text)
        source_file = str(row[8]).strip()  if row[8] else ""
        picture_id  = str(row[2]).strip()  if row[2] else ""
        section     = str(row[9]).strip()  if row[9] else ""
        alt_text    = str(row[19]).strip() if row[19] else ""

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
    Return a clean filename for the Pressbooks media library.

    Rule:
      • If the Section/Module value (Column J) is unique across the whole
        Excel file → use it as-is:           Ex1.2.1  →  Ex1.2.1.png
      • If it appears on multiple rows → append the Picture ID:
                                            1.2  + pid 5  →  1.2_5.png
    The result is always unique (since source_file+PID is unique, and PID
    differentiates images that share the same Section label).
    """
    if not section:
        return f"img_pid{picture_id}.png"
    base = re.sub(r'[^\w.\-]', '_', section)   # safe for web filenames
    if j_counts.get(section, 0) <= 1:
        return f"{base}.png"
    return f"{base}_{picture_id}.png"


# ============================================================
# WORD XML HELPERS
# ============================================================

def build_media_to_pid_map(docx_path: Path) -> dict:
    """
    Return {media_filename_stem: picture_id} for every image in the docx.

    Uses alttext_utils functions that extract Picture IDs from the Word XML.
    The stem (filename without extension) is used so .emf and .png variants
    both resolve to the same Picture ID.

    Example:  {"image1": "2", "image24": "12", ...}
    """
    rels    = utils.parse_relationships(docx_path)        # rId → 'media/image1.emf'
    docpr   = utils.extract_docpr_tags_with_rid(docx_path) # rId → {picture_id, ...}

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
) -> str | None:
    """
    Upload an image to the Pressbooks media library and set alt text.

    Returns the public URL of the uploaded image, or None on failure.

    Pressbooks is built on WordPress; we use the standard WP REST API:
      POST /wp-json/wp/v2/media
    The alt_text is set in the same request via a form field.
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

        # Step 2 — set alt text via PATCH
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

def rewrite_html(html_path: Path, src_to_info: dict) -> None:
    """
    Rewrite <img src="..."> tags in the HTML file.

    src_to_info maps the original absolute src path (as it appears in the
    pandoc-generated HTML) to a dict:
        {"url": "https://slcc.pressbooks.pub/...", "alt": "..."}

    Saves the updated HTML to the same path (overwrites in place).
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

    html_path.write_text(str(soup), encoding="utf-8")
    print(f"\n✅ HTML rewritten: {replaced} image(s) updated, {skipped} skipped.")
    print(f"   Saved: {html_path}")


# ============================================================
# MAIN
# ============================================================

def main():
    # --- Validate config ---
    errors = []
    if not INPUT_DOCX or not INPUT_DOCX.exists():
        errors.append(f"INPUT_DOCX not found: {INPUT_DOCX}")
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

    print(f"📄 Source Word file:  {INPUT_DOCX.name}")
    print(f"📊 Alt text Excel:    {EXCEL_PATH.name}")
    print(f"🌐 HTML file:         {HTML_PATH.name}")
    print(f"🔗 Pressbooks site:   {PB_SITE}")
    print()

    # 1. Load alt text lookup from Excel
    print("📋 Loading alt text lookup from Excel...")
    lookup, j_counts = load_alt_text_lookup(EXCEL_PATH)
    print(f"   {len(lookup)} image(s) found in Excel.")

    # 2. Build Picture-ID → media filename map from the Word file
    print("📎 Reading Word XML for Picture ID → filename map...")
    stem_to_pid = build_media_to_pid_map(INPUT_DOCX)
    print(f"   {len(stem_to_pid)} image(s) found in Word file.")

    # 3. Parse the HTML to find all <img> tags
    print("🔍 Parsing HTML for <img> tags...")
    html = HTML_PATH.read_text(encoding="utf-8")
    soup = BeautifulSoup(html, "html.parser")
    img_tags = soup.find_all("img")
    print(f"   {len(img_tags)} <img> tag(s) found.")
    print()

    # 4. Process each image
    src_to_info = {}   # original src → {url, alt}
    docx_name   = INPUT_DOCX.name

    with tempfile.TemporaryDirectory() as tmp_str:
        tmp_dir = Path(tmp_str)

        for img in img_tags:
            src = img.get("src", "")
            if not src:
                continue

            img_path = Path(src)
            if not img_path.exists():
                print(f"  ⚠️  File not found on disk, skipping: {src}")
                continue

            stem   = img_path.stem                   # e.g. 'image1'
            suffix = img_path.suffix.lower()         # e.g. '.emf'

            # --- Skip WMF (low quality, rarely meaningful) ---
            if suffix == ".wmf":
                print(f"  ⏭️  Skipping WMF: {img_path.name}")
                continue

            # --- Look up Picture ID ---
            pid = stem_to_pid.get(stem)
            if not pid:
                print(f"  ⚠️  No Picture ID found for '{stem}', skipping.")
                continue

            # --- Look up alt text ---
            key = (docx_name, pid)
            info = lookup.get(key)
            if not info:
                print(f"  ⚠️  No Excel row for ({docx_name}, PID={pid}), skipping.")
                continue

            alt_text = info["alt_text"]
            section  = info["section"]

            if not alt_text:
                print(f"  ⚠️  Column T is empty for PID={pid} ({section}). Skipping upload.")
                continue

            # --- Determine clean Pressbooks filename ---
            pb_filename = smart_filename(section, pid, j_counts)

            print(f"  🖼  {img_path.name}  (PID={pid}, section={section})  →  {pb_filename}")

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

    # 5. Rewrite the HTML
    if src_to_info:
        print()
        print(f"✏️  Rewriting HTML ({len(src_to_info)} image(s) updated)...")
        rewrite_html(HTML_PATH, src_to_info)
    else:
        print("\n⚠️  No images were successfully uploaded — HTML not modified.")

    print()
    print("Done! Next step: paste the updated HTML body into the Pressbooks chapter editor.")


if __name__ == "__main__":
    main()
