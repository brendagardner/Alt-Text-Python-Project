# THE PURPOSE OF THIS SCRIPT IS TO Build Alt Text Pipeline: DOCX → Images → Abacus (multi-threaded) → Rich Excel ===

# NOTES BEFORE RUNNING:
# 1. SET API KEY CONFIGURATIONS - BuildAPIVisionKey
# 2. REVIEW FOLDER, OUTPUT FILE NAME AND VISION MODEL

# Brenda Gardner — Created Nov 4 2025
# Dec 10 @12:20 - Updated column layout A–X per new Master File

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
# Folder where the Word DOCX files are located
FOLDER_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/Small Test Word to AI"
)
OUTPUT_EXCEL_PREFIX = "103x_Voting_Theory_AI"
DEFAULT_GRADE_BAND = "12-14"
COURSE_TOPIC = "Quantitative Reasoning"

# Optional to configure.
LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

ABACUS_API_KEY = os.getenv("ABACUS_API_KEY")
ROUTE_LLM_ENDPOINT = "https://routellm.abacus.ai/v1/chat/completions"
VISION_MODEL = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"
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


def worker_process(
    docx_path: Path,
    paragraphs,
    paragraph_texts: List[str],
    rid_para_map: Dict[str, List[int]],
    pic_id_map: Dict[str, str],
    alt_map: Dict[str, str],
    rels: Dict[str, str],
    docpr_map: Dict[str, Dict[str, str]],
    grade_band: str,
    course_topic: str,
    doc_author: str,
) -> List[Dict[str, Union[str, int]]]:
    rows: List[Dict[str, Union[str, int]]] = []

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
        if suffix in (".emf", ".svg"):
            converted = utils.convert_vector_to_png_bytes(
                img_bytes, suffix, docx_path.parent
            )
            img_bytes = converted if converted else None
        if not img_bytes:
            # print(f"Skipping {rid} -> {target}: conversion failed")
            continue

        # Use the *first* occurrence for context
        p_idx = para_indices[0]

        existing_alt = alt_map.get(rid, "").strip()
        section = utils.extract_section_from_heading(paragraphs, p_idx)
        page_num = utils.guess_page_number(p_idx)
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
            "section": section,
            "page_number": page_num or "",
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
    print(f"📂 Scanning folder: {FOLDER_PATH}")
    if not FOLDER_PATH.exists():
        raise FileNotFoundError(f"❌ Folder does not exist: {FOLDER_PATH}")

    docx_files = sorted(FOLDER_PATH.glob("*.docx"))
    if not docx_files:
        print("⚠️ No .docx files found.")
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
        paragraphs = doc.paragraphs  # NEW
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

        print("docpr_map keys:", list(docpr_map.keys()))

        rows = worker_process(
            docx_path=docx_path,
            paragraphs=paragraphs,  # NEW
            paragraph_texts=paragraph_texts,
            rid_para_map=rid_para_map,
            pic_id_map=pic_id_map,
            alt_map=alt_map,
            rels=rels,
            docpr_map=docpr_map,
            grade_band=DEFAULT_GRADE_BAND,
            course_topic=COURSE_TOPIC,
            doc_author=doc_author,
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

    out = FOLDER_PATH / f"{OUTPUT_EXCEL_PREFIX}_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    wb.save(out)
    print(f"\n✅ Completed — Excel saved to:\n{out}")


if __name__ == "__main__":
    main()
