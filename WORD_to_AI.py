# THE PURPOSE OF THIS SCRIPT IS TO Build Alt Text Pipeline: DOCX → Images → Abacus (multi-threaded) → Rich Excel ===

# NOTES BEFORE RUNNING:
# 1. SET API KEY CONFIGURATIONS - BuildAPIVisionKey
# 2. REVIEW FOLDER, OUTPUT FILE NAME AND VISION MODEL

# Brenda Gardner — Nov 4 2025 Updated - Added LocationIndex in Column D.
# Nov 26 - Updated Style guide, .emf file conversion working better

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
FOLDER_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1050 Information/Module 1 Sections/Section 1.1  OER College Algebra Spr2021_Images/"
)
OUTPUT_EXCEL_PREFIX = "1050_1.1_AI_AltText"
DEFAULT_GRADE_BAND = "10-12"
COURSE_TOPIC = "Intermediate Algebra"

LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

ABACUS_API_KEY = os.getenv("ABACUS_API_KEY")
ROUTE_LLM_ENDPOINT = "https://routellm.abacus.ai/v1/chat/completions"
VISION_MODEL = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"
MAX_THREADS = 3
API_TIMEOUT_SEC = 180

if not ABACUS_API_KEY:
    raise EnvironmentError("❌ Missing ABACUS_API_KEY environment variable.")

#shared helpers now found in alttext_utils.py

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
    paragraphs,                      # NEW
    paragraph_texts: List[str],
    rid_para_map: Dict[str, List[int]],
    pic_id_map: Dict[str, str],
    alt_map: Dict[str, str],
    rels: Dict[str, str],
    docpr_map: Dict[str, Dict[str, str]],
    grade_band: str,
    course_topic: str,
) -> List[Dict[str, Union[str, int]]]:
    rows: List[Dict[str, Union[str, int]]] = []

    for rid, target in rels.items():
        img_bytes = utils.load_image_bytes(docx_path, target)
        if not img_bytes:
            continue

        suffix = Path(target).suffix.lower()
        if suffix in (".emf", ".svg"):
            converted = utils.convert_vector_to_png_bytes(
                img_bytes, suffix, docx_path.parent
            )
            img_bytes = converted if converted else None
        if not img_bytes:
            continue

        para_indices = rid_para_map.get(rid, [-1])
        existing_alt = alt_map.get(rid, "").strip()

        for p_idx in para_indices:
            # ✅ Use real paragraph objects for heading/section
            section = utils.extract_section_from_heading(paragraphs, p_idx)
            page_num = utils.guess_page_number(p_idx)

            # ✅ Use text list for context snippet
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
                "section": section,
                "page_number": page_num or "",
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

    headers = [
        "Thumbnail",          # A
        "Image Filename",     # B
        "Picture ID",         # C
        "LocationIndex",      # D
        "File Name/Source",   # E
        "Image Topic",        # F
        "Section/Module",     # G
        "Page Number",        # H
        "In Word?",           # I
        "In Excel?",          # J
        "Current Alt Text",   # K
        "AI Short Alt Text",  # L
        "AI Standard Alt",    # M
        "AI Long Alt Text",   # N
        "AI Notes",           # O
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

        pic_id_map = utils.parse_picture_ids(docx_path)
        alt_map = utils.parse_alt_text(docx_path)
        rels = utils.parse_relationships(docx_path)
        rid_para_map = utils.locate_all_image_instances(doc)
        docpr_map = utils.extract_docpr_tags_with_rid(docx_path)

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
        )

        for loc_index, r in enumerate(rows, start=1):
            thumb_bytes = r.pop("thumbnail_bytes", None)
            if thumb_bytes:
                img_stream = BytesIO(thumb_bytes)
                excel_img = ExcelImage(img_stream)
                image_refs.append((img_stream, excel_img))

                try:
                    pil_img = PILImage.open(BytesIO(thumb_bytes))
                    w, h = pil_img.size
                except Exception:
                    w, h = (0, 0)

                ws.add_image(excel_img, f"A{row}")
                ws.row_dimensions[row].height = int(h * utils.ROW_HEIGHT_SCALE)
            else:
                ws.cell(row=row, column=1, value="[No Thumbnail]")

            # Column B: Image Filename
            raw_name = r.get("image_name", "")
            clean_name = Path(raw_name).name  # strips any folder prefix
            ws.cell(row=row, column=2, value=clean_name)
            ws.cell(row=row, column=3, value=r.get("picture_id", "")) #C
            ws.cell(row=row, column=4, value=loc_index)  # D - LocationIndex (within this doc)
            ws.cell(row=row, column=5, value=r.get("file_name_source", ""))
            ws.cell(row=row, column=6, value=r.get("image_topic", ""))
            ws.cell(row=row, column=7, value=r.get("section", ""))
            ws.cell(row=row, column=8, value=r.get("page_number", ""))
            ws.cell(row=row, column=9, value=r.get("in_word", "Yes"))
            ws.cell(row=row, column=10, value=r.get("in_excel", ""))
            ws.cell(row=row, column=11, value=r.get("current_alt_text", ""))
            ws.cell(row=row, column=12, value=r.get("ai_short", ""))
            ws.cell(row=row, column=13, value=r.get("ai_standard", ""))
            ws.cell(row=row, column=14, value=r.get("ai_long", ""))
            ws.cell(row=row, column=15, value=r.get("ai_notes", ""))

            row += 1

    # Autofit columns & wrap text
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_length = 0
        for cell in col_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        adjusted_width = min(max_length + 2, 60)
        ws.column_dimensions[col_letter].width = adjusted_width
        ws.column_dimensions["A"].width = 35

    out = FOLDER_PATH / f"{OUTPUT_EXCEL_PREFIX}_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    wb.save(out)
    print(f"\n✅ Completed — Excel saved to:\n{out}")


if __name__ == "__main__":
    main()
