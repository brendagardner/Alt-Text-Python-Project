# THE PURPOSE OF THIS SCRIPT IS TO Build Alt Text Pipeline: DOCX → Images → Zamzar (.EMF) → Abacus (multi-threaded) → Rich Excel ===
# Provides a single Excel file with all the Alt Text data.

# NOTES BEFORE RUNNING:
# 1. SET API KEY CONFIGURATIONS - BuildAPIVisionKey for both Zamzar and Abacus
# 2. REVIEW FOLDER, OUTPUT FILE NAME AND VISION MODEL

# Brenda Gardner — Created Nov 4 2025
# updated Feb 17 to use Zamzar API for EMF to PNG conversion
# Updated Feb 18 to add converted images to a sub folder specified, put figure numer in column J and fixed page numbers
# Updated Feb 19 to input "Ex*.*.*" in Column J for example figures, added output folder for images

import os
import re
import json
import base64
import shutil
import zipfile
import subprocess
import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Union, Tuple

import requests

from docx import Document
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from helpers import alttext_utils as utils
from lxml import etree

# =========================
# CONFIG — EDIT THESE
# =========================
# INPUT_PATH can be EITHER:
#   • a folder containing .docx files, OR
#   • a single .docx file
INPUT_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/0920 Information/920 2025 Lecture Notes/BAIN_Mod4_Accessible_Lecture Notes.docx"
)
OUTPUT_EXCEL_PREFIX = "0920_BAIN_Mod4_AltText"
OUTPUT_EXCEL_PATH = Path(f"{OUTPUT_EXCEL_PREFIX}.xlsx")
OUTPUT_WORD_PATH = Path(f"{OUTPUT_EXCEL_PREFIX}.docx")
DEFAULT_GRADE_BAND = "6-8"
COURSE_TOPIC = "Foundational Math"

# PS file output folder
#PS_OUTPUT_FOLDER = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1060 Information/OER Trigonometry Textbook 2023/Module 1/Sections 1.2-1.6/Sections 1.2-1.6 Files")
#PS_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# PNG file output folder (High Quality for Web)
PNG_OUTPUT_FOLDER = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/0920 Information/920 2025 Lecture Notes/920 Images/Module 4 Images")
PNG_OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

# API KEYS
ZAMZAR_API_KEY = os.getenv('ZAMZAR')
if not ZAMZAR_API_KEY:
    print("⚠️ Warning: ZAMZAR API key not found. EMF conversion will be skipped.")

ZAMZAR_BASE_URL = "https://api.zamzar.com"
ZAMZAR_JOBS_URL = f"{ZAMZAR_BASE_URL}/v1/jobs"

LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
ABACUS_API_KEY = os.getenv("ABACUS_API_KEY")
ROUTE_LLM_ENDPOINT = "https://routellm.abacus.ai/v1/chat/completions"
VISION_MODEL = "route-llm" #meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8 is cheapest
MAX_THREADS = 3
API_TIMEOUT_SEC = 180

if not ABACUS_API_KEY:
    raise EnvironmentError("❌ Missing ABACUS_API_KEY environment variable.")

# shared helpers found in alttext_utils.py

def build_prompt(grade_band: str, course_topic: str, context: str) -> str:
    """
    Build the prompt for Abacus using the shared STYLE_GUIDE_RULES from utils.
    """
    return f"{utils.STYLE_GUIDE_RULES}\nContext snippet: {context}\nReturn ONLY JSON."


def call_abacus(img_bytes: bytes, prompt: str) -> Dict[str, str]:
    """
    Call the Abacus vision endpoint and return the parsed JSON dict.
    """
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    payload = {
        "model": VISION_MODEL,
        "messages": [
            {
                "role": "system",
                "content": "You are an accessibility expert returning ONLY JSON alt text.",
            },
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ],
            },
        ],
    }
    headers = {
        "Authorization": f"Bearer {ABACUS_API_KEY}",
        "Content-Type": "application/json",
    }
    r = requests.post(ROUTE_LLM_ENDPOINT, headers=headers, json=payload, timeout=API_TIMEOUT_SEC)
    r.raise_for_status()
    content = r.json().get("choices", [{}])[0].get("message", {}).get("content", "")

    try:
        js = json.loads(content.strip("`").replace("json", "").strip())
    except Exception:
        # Fallback on parse error
        return {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}

    # Ensure topic key exists
    if isinstance(js, dict) and "topic" not in js:
        js["topic"] = ""
    return js


def convert_emf_via_zamzar(img_bytes: bytes, original_filename: str) -> Tuple[bytes, bytes]:
    """
    """
    # Convert EMF to High-Quality PNG using Zamzar API.
    # Saves the PNG to PNG_OUTPUT_FOLDER.
    """
    Convert EMF to PNG using Zamzar API.
    (.ps conversion is currently commented out to save credits)
    Returns: (png_bytes, None) or (None, None) on failure
    """
    if not ZAMZAR_API_KEY:
        print("⚠️ Skipping Zamzar conversion: no API key")
        return None, None

    try:
        # Step 1: Upload EMF and request conversion to PNG with High DPI
        files = {'source_file': (original_filename, BytesIO(img_bytes), 'image/x-emf')}

        # Added export_options for high quality (300 DPI)
        data = {
            'target_format': 'png',
            'export_options': 'dpi=300'
        }

        response = requests.post(
            ZAMZAR_JOBS_URL,
            auth=(ZAMZAR_API_KEY, ''),
            files=files,
            data=data,
            timeout=60
        )
        response.raise_for_status()
        job_png = response.json()
        job_id_png = job_png['id']

        # --- PS CONVERSION OPTION (DISABLED TO SAVE CREDITS) ---
        # Step 2: Upload EMF and request conversion to PS
        # files = {'source_file': (original_filename, BytesIO(img_bytes), 'image/x-emf')}
        # data = {'target_format': 'ps'}
        # response = requests.post(
        #     ZAMZAR_JOBS_URL,
        #     auth=(ZAMZAR_API_KEY, ''),
        #     files=files,
        #     data=data,
        #     timeout=60
        # )
        # response.raise_for_status()
        # job_ps = response.json()
        # job_id_ps = job_ps['id']
        # -------------------------------------------------------

        # Step 3: Poll for PNG completion
        png_bytes = None
        for _ in range(30):
            status_response = requests.get(
                f"{ZAMZAR_BASE_URL}/v1/jobs/{job_id_png}",
                auth=(ZAMZAR_API_KEY, ''),
                timeout=30
            )
            status_response.raise_for_status()
            status = status_response.json()

            if status['status'] == 'successful':
                file_id = status['target_files'][0]['id']
                download_response = requests.get(
                    f"{ZAMZAR_BASE_URL}/v1/files/{file_id}/content",
                    auth=(ZAMZAR_API_KEY, ''),
                    timeout=60
                )
                download_response.raise_for_status()
                png_bytes = download_response.content

                # --- NEW: Save High-Quality PNG to Folder ---
                png_filename = Path(original_filename).stem + ".png"
                png_path = PNG_OUTPUT_FOLDER / png_filename
                with open(png_path, 'wb') as f:
                    f.write(png_bytes)
                print(f"✅ Saved High-Quality PNG: {png_path}")
                # --------------------------------------------
                break

            elif status['status'] == 'failed':
                print(f"❌ Zamzar PNG conversion failed for {original_filename}")
                break

            import time
            time.sleep(10)

        # --- PS POLLING OPTION (DISABLED TO SAVE CREDITS) ---
        # ps_bytes = None
        # for _ in range(30):
        #     status_response = requests.get(
        #         f"{ZAMZAR_BASE_URL}/v1/jobs/{job_id_ps}",
        #         auth=(ZAMZAR_API_KEY, ''),
        #         timeout=30
        #     )
        #     status_response.raise_for_status()
        #     status = status_response.json()
        #     if status['status'] == 'successful':
        #         file_id = status['target_files'][0]['id']
        #         download_response = requests.get(
        #             f"{ZAMZAR_BASE_URL}/v1/files/{file_id}/content",
        #             auth=(ZAMZAR_API_KEY, ''),
        #             timeout=60
        #         )
        #         download_response.raise_for_status()
        #         ps_bytes = download_response.content
        #         ps_filename = Path(original_filename).stem + ".ps"
        #         ps_path = PS_OUTPUT_FOLDER / ps_filename
        #         with open(ps_path, 'wb') as f:
        #             f.write(ps_bytes)
        #         print(f"✅ Saved PS file: {ps_path}")
        #         break
        #     elif status['status'] == 'failed':
        #         print(f"❌ Zamzar PS conversion failed for {original_filename}")
        #         break
        #     import time
        #     time.sleep(10)
        # -------------------------------------------------------

        return png_bytes, None # Returning None for ps_bytes since it's disabled

    except Exception as e:
        print(f"❌ Zamzar conversion error for {original_filename}: {e}")
        return None, None

def worker_process(
    docx_path,
    paragraphs,
    paragraph_texts,
    rid_para_map,
    pic_id_map,
    alt_map,
    rels,
    docpr_map,
    grade_band,
    course_topic,
    doc_author,
    rid_page_map,
    rid_figure_map,
):
    rows = []

    # Ensure we only call the API once per underlying media file
    seen_images = set()

    for rid, target in rels.items():
        # 1️⃣ Skip any rel that does NOT have a wp:docPr entry
        #    → this cuts out orphan images like your extra media/image2.png
        if rid not in docpr_map:
            # print(f"Skipping {rid} -> {target}: no docPr (orphan rel)")
            continue

        # 2️⃣ Get paragraph indices for this rid
        para_indices = rid_para_map.get(rid) or []
        if not para_indices:
            # print(f"Skipping {rid} -> {target}: no paragraph indices")
            continue

        # 3️⃣ Only process each media target once per DOCX
        image_key = (docx_path.name, target)
        if image_key in seen_images:
            # print(f"Skipping {rid} -> {target}: already processed")
            continue
        seen_images.add(image_key)

        img_bytes = utils.load_image_bytes(docx_path, target)
        if not img_bytes:
            # print(f"Skipping {rid} -> {target}: no img_bytes")
            continue

        suffix = Path(target).suffix.lower()

        if suffix == ".emf":
            # EMF files: convert to PNG via Zamzar API
            print(f"🔄 Converting EMF via Zamzar: {target}")
            png_bytes, ps_bytes = convert_emf_via_zamzar(img_bytes, os.path.basename(target))
            if png_bytes:
                img_bytes = png_bytes
                print(f"✅ EMF converted to PNG via Zamzar")
            else:
                print(f"⚠️ Zamzar conversion failed, skipping {target}")
                continue
        elif suffix == ".svg":
            # SVG files: convert to PNG via CairoSVG / LibreOffice
            print(f"🔄 Converting SVG: {target}")
            converted = utils.convert_vector_to_png_bytes(img_bytes, suffix, docx_path.parent)
            if converted:
                img_bytes = converted
                print(f"✅ SVG converted to PNG")
            else:
                print(f"⚠️ SVG conversion failed, skipping {target}")
                continue
        elif suffix in (".wmf",):
            # WMF: skip — low quality and rarely meaningful
            print(f"⏭️ Skipping WMF file: {target}")
            continue
        elif suffix in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp"):
            # Standard raster formats: use as-is
            print(f"📷 Processing raster image: {target}")
        else:
            print(f"⏭️ Skipping unsupported format: {target}")
            continue

        if not img_bytes:
            continue

        # Use the *first* occurrence for context
        p_idx = para_indices[0]

        existing_alt = alt_map.get(rid, "").strip()

        # Figure label from the nearest Caption paragraph (forward-walk map)
        figure_label = rid_figure_map.get(rid, "")
        if not figure_label:
            # Fallback: search backwards in the flat paragraph list
            figure_label = utils.extract_figure_label_before_image(paragraphs, p_idx)

        # Page number from lastRenderedPageBreak map (accurate per-image pages)
        page_num = rid_page_map.get(rid, "")

        context = utils.build_context_snippet(paragraph_texts, p_idx, window=1)
        prompt = build_prompt(grade_band, course_topic, context)

        try:
            ai_json = call_abacus(img_bytes, prompt)
        except Exception as e:
            print(f"⚠️ Abacus API error on {target}: {e}")
            ai_json = {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}

        short_alt = utils.clean_text(ai_json.get("short", ""), 250)
        standard_alt = utils.clean_text(ai_json.get("standard", ""), 600)
        long_alt = utils.clean_text(ai_json.get("long", ""), 2000)
        notes = utils.clean_text(ai_json.get("notes", ""), 1000)
        topic = utils.clean_text(ai_json.get("topic", ""), 100)

        pid_info = docpr_map.get(rid, {})
        pic_id = pid_info.get("picture_id", "")
        image_name = pid_info.get("image_file", os.path.basename(target))

        row = {
            "thumbnail_bytes": img_bytes,
            "image_name": image_name,
            "picture_id": pic_id,
            "file_name_source": docx_path.name,
            "image_topic": topic,
            "image_context": context,
            "grade_level": grade_band,
            "course_topic": course_topic,
            "section": figure_label,
            "page_number": page_num,  # Use header page number
            "doc_author": doc_author,
            "in_word": "Yes",
            "in_excel": "",
            "current_alt_text": existing_alt,
            "ai_short": short_alt,
            "ai_standard": standard_alt,
            "ai_long": long_alt,
            "ai_notes": notes,
        }
        rows.append(row)

    return rows


def main():
    # --- Resolve input: single file or folder ---
    print(f"🔍 Checking input path: {INPUT_PATH}")

    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"❌ INPUT_PATH does not exist: {INPUT_PATH}")

    if INPUT_PATH.is_dir():
        # Folder mode: all .docx files in the folder
        docx_files = sorted(INPUT_PATH.glob("*.docx"))
        output_dir = INPUT_PATH
        print(f"📂 Folder mode. Found {len(docx_files)} .docx file(s) in {INPUT_PATH}")
    elif INPUT_PATH.is_file():
        # Single-file mode: just that file
        if INPUT_PATH.suffix.lower() != ".docx":
            raise ValueError(f"❌ INPUT_PATH is a file but not a .docx: {INPUT_PATH}")
        docx_files = [INPUT_PATH]
        output_dir = INPUT_PATH.parent
        print(f"📄 Single-file mode. Processing one DOCX: {INPUT_PATH.name}")
    else:
        raise ValueError(f"❌ INPUT_PATH is neither a file nor a folder: {INPUT_PATH}")

    if not docx_files:
        print("⚠️ No .docx files found to process.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Alt Text AI Output"

    # === New headers A–X ===
    headers = [
        "Image Thumbnail",                         # A
        "Image Filename",                          # B
        "Picture ID",                              # C
        "Image Size",                              # D
        "Image Topic",                             # E
        "Image Context",                           # F
        "Grade Level",                             # G
        "Course Topic",                            # H
        "File Name/Source",                        # I
        "Section/Module",                          # J
        "Page",                                    # K
        "Doc Author",                              # L
        "Current Alt Text",                        # M
        "Current Alt Text Character Count",        # N
        "Current Alt Text Notes",                  # O (blank)
        "A.I. Short Alt Text",                     # P
        "A.I. Med Alt Text",                       # Q
        "A.I. Long Alt Text",                      # R
        "A.I. Notes",                              # S
        "Suggested Alt Text",                      # T (blank)
        "Suggested Character Count",               # U (LEN(T))
        "Suggested Alt Text Notes",                # V (blank)
        "Reviewer Approved",                       # W (blank)
        "Reviewer notes",                          # X (blank)
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)

    row = 2
    image_refs: List[Tuple[BytesIO, ExcelImage]] = []

    for docx_path in docx_files:
        print(f"\n📄 Processing: {docx_path.name}")
        doc = Document(str(docx_path))

        # NEW
        paragraphs = doc.paragraphs
        paragraph_texts = utils.read_paragraph_texts(doc)

        # Doc author for this document
        doc_author = ""
        try:
            doc_author = doc.core_properties.author or ""
        except Exception:
            doc_author = ""

        pic_id_map = utils.parse_picture_ids(docx_path)
        alt_map = utils.parse_alt_text(docx_path)
        rels = utils.parse_relationships(docx_path)
        rid_para_map = utils.locate_all_image_instances(doc)
        docpr_map = utils.extract_docpr_tags_with_rid(docx_path)

        # Build per-image page number map (uses lastRenderedPageBreak)
        rid_page_map   = utils.build_rid_page_map(docx_path)
        # Build per-image figure label map (forward caption scan)
        rid_figure_map = utils.build_rid_figure_map(docx_path)

        print("docpr_map keys:", list(docpr_map.keys()))
        print("rid_page_map sample:", dict(list(rid_page_map.items())[:5]))
        print("rid_figure_map sample:", dict(list(rid_figure_map.items())[:5]))

        rows = worker_process(
            docx_path=docx_path,
            paragraphs=paragraphs,
            paragraph_texts=paragraph_texts,
            rid_para_map=rid_para_map,
            pic_id_map=pic_id_map,
            alt_map=alt_map,
            rels=rels,
            docpr_map=docpr_map,
            grade_band=DEFAULT_GRADE_BAND,
            course_topic=COURSE_TOPIC,
            doc_author=doc_author,
            rid_page_map=rid_page_map,
            rid_figure_map=rid_figure_map,
        )

        for loc_index, r in enumerate(rows, start=1):
            thumb_bytes = r.get("thumbnail_bytes", None)

            # Default values
            width = height = 0

            if thumb_bytes:
                img_stream = BytesIO(thumb_bytes)
                excel_img = ExcelImage(img_stream)
                image_refs.append((img_stream, excel_img))

                try:
                    pil_img = PILImage.open(BytesIO(thumb_bytes))
                    width, height = pil_img.size
                except Exception:
                    width, height = (0, 0)

                ws.add_image(excel_img, f"A{row}")
                ws.row_dimensions[row].height = int(height * utils.ROW_HEIGHT_SCALE)
            else:
                ws.cell(row=row, column=1, value="[No Thumbnail]")

            # Column B: Image Filename
            raw_name = r.get("image_name", "")
            clean_name = Path(raw_name).name  # strips any folder prefix
            ws.cell(row=row, column=2, value=clean_name)

            # Column C: Picture ID
            ws.cell(row=row, column=3, value=r.get("picture_id", ""))

            # Column D: Image Size (e.g., "800 x 600")
            size_str = f"{width} x {height}" if width and height else ""
            ws.cell(row=row, column=4, value=size_str)

            # Column E: Image Topic
            ws.cell(row=row, column=5, value=r.get("image_topic", ""))

            # Column F: Image Context (snippet used for prompt)
            ws.cell(row=row, column=6, value=r.get("image_context", ""))

            # Column G: Grade Level
            ws.cell(row=row, column=7, value=r.get("grade_level", DEFAULT_GRADE_BAND))

            # Column H: Course Topic
            ws.cell(row=row, column=8, value=r.get("course_topic", COURSE_TOPIC))

            # Column I: File Name/Source
            ws.cell(row=row, column=9, value=r.get("file_name_source", ""))

            # Column J: Section/Module
            ws.cell(row=row, column=10, value=r.get("section", ""))

            # Column K: Page
            ws.cell(row=row, column=11, value=r.get("page_number", ""))

            # Column L: Doc Author
            ws.cell(row=row, column=12, value=r.get("doc_author", ""))

            # Column M: Current Alt Text
            current_alt = r.get("current_alt_text", "")
            ws.cell(row=row, column=13, value=current_alt)

            # Column N: Current Alt Text Character Count (formula LEN(Mrow))
            ws.cell(row=row, column=14, value=f"=LEN(M{row})")

            # Column O: Current Alt Text Notes (blank per spec)
            # (no value set)

            # Column P: A.I. Short Alt Text
            ws.cell(row=row, column=16, value=r.get("ai_short", ""))

            # Column Q: A.I. Med Alt Text
            ws.cell(row=row, column=17, value=r.get("ai_standard", ""))

            # Column R: A.I. Long Alt Text
            ws.cell(row=row, column=18, value=r.get("ai_long", ""))

            # Column S: A.I. Notes
            ws.cell(row=row, column=19, value=r.get("ai_notes", ""))

            # Column T: Suggested Alt Text (blank per spec)
            # (no initial value)

            # Column U: Suggested Character Count (formula LEN(Trow))
            ws.cell(row=row, column=21, value=f"=LEN(T{row})")

            # Column V: Suggested Alt Text Notes (blank per spec)
            # Column W: Reviewer Approved (blank per spec)
            # Column X: Reviewer notes (blank per spec)
            # (no values set)

            row += 1

    # === Column sizing & wrapping (similar to Website script) ===

    # Make thumbnail column wide in pixels (converted to Excel width)
    ws.column_dimensions['A'].width = utils._pixels_to_column_width(600)

    # Base widths for the rest (these are starting points; wrap_and_autofit can tweak)
    ws.column_dimensions['B'].width = 30  # Image Filename
    ws.column_dimensions['C'].width = 18  # Picture ID
    ws.column_dimensions['D'].width = 14  # Image Size
    ws.column_dimensions['E'].width = 25  # Image Topic
    ws.column_dimensions['F'].width = 40  # Image Context
    ws.column_dimensions['G'].width = 12  # Grade Level
    ws.column_dimensions['H'].width = 20  # Course Topic
    ws.column_dimensions['I'].width = 30  # File Name/Source
    ws.column_dimensions['J'].width = 25  # Section/Module
    ws.column_dimensions['K'].width = 8  # Page
    ws.column_dimensions['L'].width = 20  # Doc Author
    ws.column_dimensions['M'].width = 50  # Current Alt Text
    ws.column_dimensions['N'].width = 16  # Current Alt Count
    ws.column_dimensions['O'].width = 40  # Current Alt Notes
    ws.column_dimensions['P'].width = 50  # AI Short
    ws.column_dimensions['Q'].width = 50  # AI Med
    ws.column_dimensions['R'].width = 50  # AI Long
    ws.column_dimensions['S'].width = 40  # AI Notes
    ws.column_dimensions['T'].width = 50  # Suggested Alt Text
    ws.column_dimensions['U'].width = 18  # Suggested Count
    ws.column_dimensions['V'].width = 40  # Suggested Notes
    ws.column_dimensions['W'].width = 18  # Reviewer Approved
    ws.column_dimensions['X'].width = 30  # Reviewer notes

    # Columns that should wrap and be auto-fitted (similar idea to website script)
    text_columns = [
        "B",  # Image Filename
        "C",  # Picture ID
        "D",  # Image Size
        "E",  # Image Topic
        "F",  # Image Context
        "G",  # Grade Level
        "H",  # Course Topic
        "I",  # File Name/Source
        "J",  # Section/Module
        "K",  # Page
        "L",  # Doc Author
        "M",  # Current Alt Text
        "O",  # Current Alt Text Notes
        "P",  # AI Short
        "Q",  # AI Med
        "R",  # AI Long
        "S",  # AI Notes
        "T",  # Suggested Alt Text
        "V",  # Suggested Notes
        "W",  # Reviewer Approved
        "X",  # Reviewer notes
    ]

    # Use the same helper as the website script
    utils.wrap_and_autofit_columns(
        ws,
        text_columns=text_columns,
        max_col_width_px=240,
    )

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    out = output_dir / f"{OUTPUT_EXCEL_PREFIX}_{timestamp}.xlsx"
    wb.save(out)
    print(f"\n✅ Completed — Excel saved to:\n{out}")


if __name__ == "__main__":
    main()
