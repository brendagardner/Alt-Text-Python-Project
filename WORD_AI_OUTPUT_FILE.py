# THE PURPOSE OF THIS SCRIPT IS TO Build Alt Text Pipeline: DOCX → Images → Abacus (multi-threaded) → Rich Excel ===
# Improvement of Word_AI_Output_File

# TO FIX:
# 1. .EMF FILES

# NOTES BEFORE RUNNING:
# 1. SET API KEY CONFIGURATIONS - BuildAPIVisionKey
# 2. REVIEW FOLDER, OUTPUT FILE NAME AND VISION MODEL

# Brenda Gardner — Nov 4 2025
# Updated - Added LocationIndex in Column D.  Captured duplicate images.


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

from docx import Document
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from lxml import etree

# =========================
# CONFIG — EDIT THESE
# =========================
FOLDER_PATH = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1010 Information/1010 Lecture Notes Chapters 1-5")
OUTPUT_EXCEL_PREFIX = "1010_Mod6-7_Create_AltText"
DEFAULT_GRADE_BAND = "10-12"
COURSE_TOPIC = "Intermediate Algebra"

LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
TARGET_IMAGE_WIDTH_PX = 800
ROW_HEIGHT_SCALE = 0.85

ABACUS_API_KEY = os.getenv("ABACUS_API_KEY")
ROUTE_LLM_ENDPOINT = "https://routellm.abacus.ai/v1/chat/completions"
VISION_MODEL = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"
MAX_THREADS = 3
API_TIMEOUT_SEC = 180

if not ABACUS_API_KEY:
    raise EnvironmentError("❌ Missing ABACUS_API_KEY environment variable.")

STYLE_GUIDE_RULES = """
You are an accessibility expert for STEM education. For each image, in addition to your JSON alt text, also suggest a topic for the image that best matches its mathematical or visual content and context (e.g., 'number line', 'parabola', 'coordinate system', 'table', etc.). Use your best judgment; the following are just examples, and you may add more based on your reasoning:
number line, points, definition, coordinate system, slope, linear graph, linear systems, linear inequalities, exponential, growth model, decay model, logarithmic, absolute value, parabola, polynomial, rational function, domain and range, general function, histogram, scatterplot, miscellaneous (non-math)
Return valid JSON with keys: short, standard, long, notes, topic
"""

def clean_text(s: str, max_len: int) -> str:
    return re.sub(r"\s+", " ", s).strip()[:max_len] if s else ""

def safe_to_str(value):
    if value is None:
        return ""
    if isinstance(value, (list, dict)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)

def guess_page_number(paragraph_index: Union[int, None]) -> Union[int, None]:
    if paragraph_index is None or paragraph_index < 0:
        return None
    return (paragraph_index // 20) + 1

def read_paragraph_texts(doc: Document) -> List[str]:
    return [p.text for p in doc.paragraphs]

def extract_section_from_heading(paragraphs: List, p_idx: int) -> str:
    if p_idx is None or p_idx < 0 or p_idx >= len(paragraphs):
        return ""
    for i in range(p_idx, -1, -1):
        para = paragraphs[i]
        style_name = para.style.name if para.style else ""
        if style_name == "Heading 1":
            text = para.text.strip()
            m = re.match(r"^([A-Za-z]*\s*\d+(\.\d+)*)(\s|$)", text)
            if m:
                return m.group(1)
            return text
    return ""

def build_context_snippet(paragraph_texts: List[str], para_idx: int, window: int = 1) -> str:
    left = max(0, para_idx - window)
    right = min(len(paragraph_texts), para_idx + window + 1)
    return clean_text(" ".join(paragraph_texts[left:right]), 500)

def parse_picture_ids(docx_path: Path) -> Dict[str, str]:
    result = {}
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            matches = re.findall(r'<wp:docPr[^>]*id="(\d+)"[^>]*name="([^"]*)"[^>]*>', xml)
            for pic_id, name in matches:
                result[name] = pic_id
    except Exception:
        pass
    return result

def parse_alt_text(docx_path: Path) -> Dict[str, str]:
    alt_map = {}
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            for alt, embed in re.findall(r'descr="([^"]+)".+?r:embed="([^"]+)"', xml):
                alt_map[embed] = alt
    except:
        pass
    return alt_map

def parse_relationships(docx_path):
    rels_map = {}
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        rels_path = "word/_rels/document.xml.rels"
        if rels_path not in docx_zip.namelist():
            return rels_map
        with docx_zip.open(rels_path) as f:
            xml_content = f.read()
    root = etree.fromstring(xml_content)
    ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
    for rel in root.findall("r:Relationship", namespaces=ns):
        r_id = rel.get("Id")
        target = rel.get("Target")
        if target and target.startswith("media/"):
            rels_map[r_id] = os.path.basename(target)
    return rels_map

def extract_docpr_tags_with_rid(docx_path):
    ns = {
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    }
    rels_map = parse_relationships(docx_path)
    docpr_map = {}
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        with docx_zip.open("word/document.xml") as f:
            xml_content = f.read()
    root = etree.fromstring(xml_content)
    for drawing in root.findall(".//w:drawing", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
        docpr = drawing.find(".//wp:docPr", namespaces=ns)
        blip = drawing.find(".//a:blip", namespaces=ns)
        if docpr is not None and blip is not None:
            r_id = blip.get(f"{{{ns['r']}}}embed")
            image_name = rels_map.get(r_id, "")
            pic_id = docpr.get("id", "")
            title = docpr.get("title", "")
            docpr_map[r_id] = {"picture_id": pic_id, "title": title, "image_file": image_name}
    return docpr_map

def parse_rels(docx_path: Path) -> Dict[str, str]:
    rels = {}
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/_rels/document.xml.rels").decode("utf-8", errors="ignore")
            for rid, target in re.findall(r'Id="([^"]+)".+?Target="([^"]+)"', xml):
                if "image" in target.lower():
                    rels[rid] = target
    except:
        pass
    return rels

def locate_all_image_instances(doc: Document) -> Dict[str, List[int]]:
    rid_to_paras: Dict[str, List[int]] = {}
    for i, p in enumerate(doc.paragraphs):
        for run in p.runs:
            blips = run._element.xpath(
                ".//*[local-name()='blip' and namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main']"
            )
            for blip in blips:
                rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if rid:
                    rid_to_paras.setdefault(rid, []).append(i)
    def inspect_cell(cell, row_idx_hint):
        for p in cell.paragraphs:
            for run in p.runs:
                blips = run._element.xpath(
                    ".//*[local-name()='blip' and namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main']"
                )
                for blip in blips:
                    rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    if rid:
                        rid_to_paras.setdefault(rid, []).append(row_idx_hint)
        for t in cell.tables:
            inspect_table(t, row_idx_hint)
    def inspect_table(table, row_idx_hint):
        for row in table.rows:
            for cell in row.cells:
                inspect_cell(cell, row_idx_hint)
    for tbl in doc.tables:
        parent_idx = -1
        for idx, p in enumerate(doc.paragraphs):
            if tbl in p._element.getparent().getchildren():
                parent_idx = idx
                break
        if parent_idx == -1:
            parent_idx = len(doc.paragraphs) - 1
        inspect_table(tbl, parent_idx)
    return rid_to_paras

def load_image_bytes(docx_path: Path, target: str):
    try:
        with zipfile.ZipFile(docx_path) as z:
            return z.read(f"word/{target}") if f"word/{target}" in z.namelist() else None
    except Exception:
        return None

def convert_vector_to_png_bytes(input_bytes: bytes, suffix: str, temp_dir: Path) -> Union[bytes, None]:
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
    except:
        pass
    return None

def resize_png(png_bytes: bytes) -> Tuple[bytes, Tuple[int, int]]:
    pil = PILImage.open(BytesIO(png_bytes))
    aspect = pil.height / pil.width
    new_w = TARGET_IMAGE_WIDTH_PX
    new_h = int(new_w * aspect)
    out = BytesIO()
    pil.resize((new_w, new_h), PILImage.Resampling.LANCZOS).save(out, format="PNG", dpi=(300, 300))
    out.seek(0)
    return out.getvalue(), (new_w, new_h)

def build_prompt(grade_band: str, course_topic: str, context: str) -> str:
    return f"{STYLE_GUIDE_RULES}\nContext snippet: {context}\nReturn ONLY JSON."

def call_abacus(img_bytes: bytes, prompt: str) -> Dict[str, str]:
    import requests
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    payload = {
        "model": VISION_MODEL,
        "messages": [
            {"role": "system", "content": "You are an accessibility expert returning ONLY JSON alt text."},
            {"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
            ]}
        ]
    }
    headers = {"Authorization": f"Bearer {ABACUS_API_KEY}", "Content-Type": "application/json"}
    try:
        r = requests.post(ROUTE_LLM_ENDPOINT, headers=headers, json=payload, timeout=API_TIMEOUT_SEC)
        r.raise_for_status()
        content = r.json().get("choices", [{}])[0].get("message", {}).get("content", "")
        # Accept topic as a suggested key
        js = json.loads(content.strip("`").replace("json", "").strip())
        # If no topic key, fallback to ""
        if isinstance(js, dict) and "topic" not in js:
            js["topic"] = ""
        return js
    except Exception as e:
        return {"short": "", "standard": "", "long": "", "notes": "API error", "topic": ""}

def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Images"

    headers = [
        "Thumbnail", "Image Filename", "Picture ID", "Location Index", "Image Size",
        "Image Topic", "Image Context", "Grade Level", "Course Title",
        "File Name/Source", "Section/Module", "~Page", "Doc Author",
        "Current Alt Text", "Current Alt Text #",
        "AI Short", "AI Medium", "AI Long", "AI Notes",
        "Brenda Suggested", "Suggested Count", "Brenda Notes"
    ]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h).font = Font(bold=True)

    row = 2
    openpyxl_image_refs = []
    for docx in FOLDER_PATH.glob("*.docx"):
        print(f"\n📘 Processing: {docx.name}")
        doc = Document(str(docx))
        paras = doc.paragraphs

        alt_map = parse_alt_text(docx)
        rels = parse_rels(docx)
        rid_para_map = locate_all_image_instances(doc)
        pic_id_map = parse_picture_ids(docx)
        docpr_map = extract_docpr_tags_with_rid(docx)

        base_name = docx.name
        title = docx.stem
        author = doc.core_properties.author or ""

        imgid_first_idx = {}

        for rid, target in rels.items():
            if target.lower().endswith(".wmf"):
                continue

            img_bytes = load_image_bytes(docx, target)
            if not img_bytes:
                continue

            suffix = Path(target).suffix.lower()
            if suffix in (".emf", ".svg"):
                converted = convert_vector_to_png_bytes(img_bytes, suffix, FOLDER_PATH)
                if converted:
                    img_bytes = converted
                else:
                    continue

            pic_id = docpr_map.get(rid, {}).get("picture_id", "")
            para_indices = rid_para_map.get(rid, [-1])
            if pic_id not in imgid_first_idx:
                imgid_first_idx[pic_id] = row

            for idx_num, p_idx in enumerate(para_indices):
                img, (w, h) = resize_png(img_bytes)
                img_stream = BytesIO(img)
                excel_img = ExcelImage(img_stream)
                openpyxl_image_refs.append((img_stream, excel_img))

                context = build_context_snippet(read_paragraph_texts(doc), p_idx)
                section = extract_section_from_heading(paras, p_idx)

                prompt = build_prompt(DEFAULT_GRADE_BAND, COURSE_TOPIC, context)
                current_notes = ""

                is_duplicate = idx_num > 0 or row > imgid_first_idx[pic_id]
                alt_result = {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}
                if is_duplicate:
                    current_notes = "duplicate image. copy suggested alt text"
                else:
                    try:
                        alt_result = call_abacus(img, prompt)
                    except Exception as e:
                        print(f"⚠️ API call failed for {target}: {e}")

                ws.add_image(excel_img, f"A{row}")
                ws.row_dimensions[row].height = int(h * ROW_HEIGHT_SCALE)

                ws.cell(row=row, column=2, value=os.path.basename(target))
                ws.cell(row=row, column=3, value=pic_id)
                ws.cell(row=row, column=4, value=p_idx)
                ws.cell(row=row, column=5, value=f"{w}x{h}")
                ws.cell(row=row, column=6, value=safe_to_str(alt_result.get("topic", "")))
                ws.cell(row=row, column=7, value=context)
                ws.cell(row=row, column=8, value=DEFAULT_GRADE_BAND)
                ws.cell(row=row, column=9, value=COURSE_TOPIC)
                ws.cell(row=row, column=10, value=base_name)
                ws.cell(row=row, column=11, value=section)
                ws.cell(row=row, column=12, value=guess_page_number(p_idx))
                ws.cell(row=row, column=13, value=author)

                existing_alt = alt_map.get(rid, "")
                existing_alt = existing_alt.strip() if existing_alt else ""

                ws.cell(row=row, column=14, value=safe_to_str(existing_alt))
                ws.cell(row=row, column=15, value=f"=LEN({get_column_letter(14)}{row})")

                ws.cell(row=row, column=16, value=safe_to_str(alt_result.get("short", "")))
                ws.cell(row=row, column=17, value=safe_to_str(alt_result.get("standard", "")))
                ws.cell(row=row, column=18, value=safe_to_str(alt_result.get("long", "")))
                ws.cell(row=row, column=19, value=current_notes or alt_result.get("notes", ""))

                ws.cell(row=row, column=21, value=f"=LEN({get_column_letter(20)}{row})")

                row += 1

    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_length = 0
        for cell in col_cells:
            try:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
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
