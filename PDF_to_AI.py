"""
PDF_to_AI.py — VECTOR GRAPHICS EDITION

Pipeline:
PDF → Vector Graphics (EMF/drawings) → Render as PNG → Abacus Vision → Excel

Extracts vector drawings (not just raster images) from PDFs.
Uses the shared helpers in helpers/alttext_utils.py

Note:  NOT WORKING WELL - PULLS TABLES AND TEXTBOXES AS WELL AS GRAPHICS
"""

import os
import base64
import json
import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple, Union, Optional

import fitz  # PyMuPDF
import requests
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font

from helpers import alttext_utils as utils

# =========================
# CONFIG — EDIT THESE
# =========================

INPUT_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1060 Information/Sample Pages from math-1060-oer-trigonometry-textbook-version-2.pdf"
)

OUTPUT_EXCEL_PREFIX = "Sample_PDF_VectorGraphics_AltText_Output"
DEFAULT_GRADE_BAND = "10-14"
COURSE_TOPIC = "Trigonometry"

ABACUS_API_KEY = os.getenv("ABACUS_API_KEY")
ROUTE_LLM_ENDPOINT = "https://routellm.abacus.ai/v1/chat/completions"
VISION_MODEL = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"
API_TIMEOUT_SEC = 180

# Rendering quality for vector graphics
RENDER_DPI = 300  # Higher = better quality, larger file size

# Minimum size threshold for vector graphics (in points)
MIN_WIDTH = 50
MIN_HEIGHT = 50

if not ABACUS_API_KEY:
    raise EnvironmentError("❌ Missing ABACUS_API_KEY environment variable.")


# =========================
# Abacus call
# =========================

def build_prompt(grade_band: str, course_topic: str, context: str) -> str:
    return f"{utils.STYLE_GUIDE_RULES}\nContext snippet: {context}\nReturn ONLY JSON."


def call_abacus(img_bytes: bytes, prompt: str) -> Dict[str, str]:
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
        return {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}

    if isinstance(js, dict) and "topic" not in js:
        js["topic"] = ""
    return js


# =========================
# Vector graphics extraction
# =========================

def merge_nearby_rects(rects: List[fitz.Rect], margin: float = 10) -> List[fitz.Rect]:
    """
    Merge rectangles that are close to each other.
    This groups vector paths that form a single "image" (like a graph with axes, labels, curves).
    """
    if not rects:
        return []

    # Expand each rect by margin
    expanded = [fitz.Rect(r.x0 - margin, r.y0 - margin, r.x1 + margin, r.y1 + margin) for r in rects]

    merged = []
    used = [False] * len(expanded)

    for i, rect in enumerate(expanded):
        if used[i]:
            continue

        current = fitz.Rect(rect)
        changed = True

        while changed:
            changed = False
            for j, other in enumerate(expanded):
                if not used[j] and current.intersects(other):
                    current |= other  # Union of rectangles
                    used[j] = True
                    changed = True

        merged.append(current)

    return merged


def extract_vector_graphics_from_pdf(pdf_path: Path) -> List[Dict[str, Union[str, int, bytes]]]:
    """
    Extract vector graphics (drawings/paths) from a PDF as rendered PNG images.
    Filters out text boxes and tables.

    Returns a list of dicts:
      - image_bytes: PNG bytes of the rendered region
      - image_name: synthetic name like 'page_3_graphic_2.png'
      - page_number: 1-based page index
      - bbox: bounding box as (x0, y0, x1, y1)
    """
    rows: List[Dict[str, Union[str, int, bytes]]] = []

    doc = fitz.open(pdf_path)

    for page_index in range(len(doc)):
        page = doc[page_index]
        page_num = page_index + 1

        print(f"   📄 Page {page_num}: Analyzing vector graphics...")

        # Get all drawings (vector paths) on the page
        drawings = page.get_drawings()

        if not drawings:
            print(f"      ⚠️  No vector drawings found on page {page_num}")
            continue

        # Collect bounding boxes of all drawings
        rects = []
        for drawing in drawings:
            rect = drawing.get("rect")
            if rect:
                rects.append(fitz.Rect(rect))

        if not rects:
            print(f"      ⚠️  No valid rectangles found on page {page_num}")
            continue

        # Merge nearby rectangles to group related vector paths
        merged_rects = merge_nearby_rects(rects, margin=15)

        print(f"      🔍 Found {len(rects)} drawing paths → merged into {len(merged_rects)} graphic(s)")

        # Filter out tiny graphics (likely decorative lines, borders, etc.)
        significant_rects = [
            r for r in merged_rects
            if (r.width >= MIN_WIDTH and r.height >= MIN_HEIGHT)
        ]

        # Get all text blocks on the page for filtering
        text_blocks = page.get_text("blocks")  # Returns list of (x0, y0, x1, y1, "text", block_no, block_type)

        # Filter out rectangles that are mostly text (text boxes, tables)
        graphics_only = []
        for rect in significant_rects:
            # Calculate how much of this rectangle overlaps with text
            text_area = 0
            rect_area = rect.width * rect.height

            for block in text_blocks:
                # block format: (x0, y0, x1, y1, "text", block_no, block_type)
                text_rect = fitz.Rect(block[:4])

                # Calculate intersection
                intersection = rect & text_rect  # Intersection of rectangles
                if not intersection.is_empty:
                    text_area += intersection.width * intersection.height

            # If more than 30% of the rectangle is text, skip it
            text_ratio = text_area / rect_area if rect_area > 0 else 0

            if text_ratio < 0.3:  # Less than 30% text = likely a graphic
                graphics_only.append(rect)
            else:
                print(f"      ⏭️  Skipping text-heavy region (text ratio: {text_ratio:.1%})")

        print(f"      ✅ {len(graphics_only)} graphic(s) after filtering text boxes/tables")

        for graphic_idx, rect in enumerate(graphics_only, start=1):
            # Render this region at high DPI
            mat = fitz.Matrix(RENDER_DPI / 72, RENDER_DPI / 72)  # 72 DPI is PDF default
            pix = page.get_pixmap(matrix=mat, clip=rect)

            # Convert pixmap to PNG bytes
            png_bytes = pix.tobytes("png")

            if not png_bytes:
                continue

            image_name = f"{pdf_path.stem}_page{page_num}_graphic{graphic_idx}.png"

            rows.append({
                "image_bytes": png_bytes,
                "image_name": image_name,
                "page_number": page_num,
                "bbox": (rect.x0, rect.y0, rect.x1, rect.y1),
            })

    doc.close()
    return rows

# =========================
# Main
# =========================

def main():
    print(f"🔍 Checking input path: {INPUT_PATH}")

    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"❌ INPUT_PATH does not exist: {INPUT_PATH}")

    if INPUT_PATH.is_dir():
        pdf_files = sorted(INPUT_PATH.glob("*.pdf"))
        output_dir = INPUT_PATH
        print(f"📂 Folder mode. Found {len(pdf_files)} .pdf file(s)")
    elif INPUT_PATH.is_file():
        if INPUT_PATH.suffix.lower() != ".pdf":
            raise ValueError(f"❌ INPUT_PATH is a file but not a .pdf: {INPUT_PATH}")
        pdf_files = [INPUT_PATH]
        output_dir = INPUT_PATH.parent
        print(f"📄 Single-file mode. Processing: {INPUT_PATH.name}")
    else:
        raise ValueError(f"❌ INPUT_PATH is neither file nor folder: {INPUT_PATH}")

    if not pdf_files:
        print("⚠️ No .pdf files found to process.")
        return

    # === Excel setup ===
    wb = Workbook()
    ws = wb.active
    ws.title = "Alt Text AI Output"

    headers = [
        "Image Thumbnail",  # A
        "Image Filename",  # B
        "Picture ID",  # C
        "Image Size",  # D
        "Image Topic",  # E
        "Image Context",  # F
        "Grade Level",  # G
        "Course Topic",  # H
        "File Name/Source",  # I
        "Section/Module",  # J
        "Page",  # K
        "Doc Author",  # L
        "Current Alt Text",  # M
        "Current Alt Text Character Count",  # N
        "Current Alt Text Notes",  # O
        "A.I. Short Alt Text",  # P
        "A.I. Med Alt Text",  # Q
        "A.I. Long Alt Text",  # R
        "A.I. Notes",  # S
        "Suggested Alt Text",  # T
        "Suggested Character Count",  # U
        "Suggested Alt Text Notes",  # V
        "Reviewer Approved",  # W
        "Reviewer notes",  # X
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True)

    row = 2
    image_refs: List[Tuple[BytesIO, ExcelImage]] = []

    for pdf_path in pdf_files:
        print(f"\n📄 Processing PDF: {pdf_path.name}")
        pdf_rows = extract_vector_graphics_from_pdf(pdf_path)
        print(f"   → Extracted {len(pdf_rows)} vector graphic(s) total\n")

        for r in pdf_rows:
            img_bytes = r["image_bytes"]
            page_number = r["page_number"]
            image_name = r["image_name"]
            bbox = r["bbox"]

            context = f"Vector graphic from page {page_number} of '{pdf_path.name}'. Bounding box: {bbox}"

            prompt = build_prompt(DEFAULT_GRADE_BAND, COURSE_TOPIC, context)

            print(f"   🤖 Calling AI for: {image_name}")
            try:
                ai_json = call_abacus(img_bytes, prompt)
            except Exception as e:
                print(f"      ⚠️ Abacus API error: {e}")
                ai_json = {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}

            short_alt = utils.clean_text(ai_json.get("short", ""), 250)
            standard_alt = utils.clean_text(ai_json.get("standard", ""), 600)
            long_alt = utils.clean_text(ai_json.get("long", ""), 2000)
            notes = utils.clean_text(ai_json.get("notes", ""), 1000)
            topic = utils.clean_text(ai_json.get("topic", ""), 100)

            # Thumbnail
            width = height = 0
            if img_bytes:
                img_stream = BytesIO(img_bytes)
                excel_img = ExcelImage(img_stream)
                image_refs.append((img_stream, excel_img))

                try:
                    pil_img = PILImage.open(BytesIO(img_bytes))
                    width, height = pil_img.size
                except Exception:
                    width, height = (0, 0)

                ws.add_image(excel_img, f"A{row}")
                ws.row_dimensions[row].height = int(height * utils.ROW_HEIGHT_SCALE)
            else:
                ws.cell(row=row, column=1, value="[No Thumbnail]")

            ws.cell(row=row, column=2, value=image_name)
            ws.cell(row=row, column=3, value="")
            ws.cell(row=row, column=4, value=f"{width} x {height}" if width and height else "")
            ws.cell(row=row, column=5, value=topic)
            ws.cell(row=row, column=6, value=context)
            ws.cell(row=row, column=7, value=DEFAULT_GRADE_BAND)
            ws.cell(row=row, column=8, value=COURSE_TOPIC)
            ws.cell(row=row, column=9, value=pdf_path.name)
            ws.cell(row=row, column=11, value=page_number)
            ws.cell(row=row, column=13, value="")
            ws.cell(row=row, column=14, value=f"=LEN(M{row})")
            ws.cell(row=row, column=16, value=short_alt)
            ws.cell(row=row, column=17, value=standard_alt)
            ws.cell(row=row, column=18, value=long_alt)
            ws.cell(row=row, column=19, value=notes)
            ws.cell(row=row, column=21, value=f"=LEN(T{row})")

            row += 1

    # === Column sizing ===
    ws.column_dimensions['A'].width = utils._pixels_to_column_width(600)
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 40
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 30
    ws.column_dimensions['J'].width = 25
    ws.column_dimensions['K'].width = 8
    ws.column_dimensions['L'].width = 20
    ws.column_dimensions['M'].width = 50
    ws.column_dimensions['N'].width = 16
    ws.column_dimensions['O'].width = 40
    ws.column_dimensions['P'].width = 50
    ws.column_dimensions['Q'].width = 50
    ws.column_dimensions['R'].width = 50
    ws.column_dimensions['S'].width = 40
    ws.column_dimensions['T'].width = 50
    ws.column_dimensions['U'].width = 18
    ws.column_dimensions['V'].width = 40
    ws.column_dimensions['W'].width = 18
    ws.column_dimensions['X'].width = 30

    text_columns = [
        "B", "C", "D", "E", "F", "G", "H", "I",
        "J", "K", "L", "M", "O", "P", "Q", "R",
        "S", "T", "V", "W", "X",
    ]

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