# helpers/alttext_utils.py

from __future__ import annotations

import base64
import json
import os
import re
import subprocess
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple, Union, Optional

from docx import Document
from lxml import etree
from PIL import Image as PILImage

from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment

LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
TARGET_IMAGE_WIDTH_PX = 800
ROW_HEIGHT_SCALE = 0.85

STYLE_GUIDE_RULES = STYLE_GUIDE_RULES = """
You are an accessibility expert for STEM education. Generate alt text according to these rules:
- Be accurate and concise. Do not reveal answers not explicitly shown.
- Use math screen-reader conventions:
• Negatives: say “negative three,” not “-3” unless the negative is in parentheses.
• Fractions: say “three-fourths”, not “3/4.”
• Intervals/inequalities: say “x is greater than or equal to three,” avoid symbols.
• Exponents/Roots: “x squared,” “x cubed,” “x to the fourth power.”
• Roots: “square root of,” “cube root of”, “f of x,” “g of x.”
• Coordinate plane: Coordinate plane grid with both x-axis and y-axis from negative 10 to 10
• Coordinates: "(-1,5)"
• Axes: “x-axis,” “y-axis,” include units if shown.
• Lines and slopes: say “increasing,” “decreasing,” “constant,” “positive slope,” “negative slope.”
• Fractions in expressions: “quantity a plus b over c,” when needed for clarity.
• Functions: type if clear (linear, quadratic), key features if shown (intercepts, vertex, max/min, asymptotes).
• Order of description: most critical to context, then left to right, bottom-up
• Graphs and charts - include: what type (line graph, bar chart, scatter plot), axes names/units if visible, describe comparison or trend, not exact numbers unless the image highlights them.
- Provide three variants per image:
  1) Short (<= 120 characters)
  2) Standard (<= 250 characters)
  3) Long (no limit; only if image is complex)
- Adjust language by grade level:
  - 4–6: simple words; avoid jargon.
  - 7–9: basic terms like slope, intercept, proportional.
  - 10–14: precise terms like parabola, asymptote, vertex.
- If a variant might reveal an answer, include a note: [risk: might reveal answer].
- Do not include colors unless colors are mentioned in context.  
- Do not use "an image of" and similar.
- No need to include information also found in the image caption or image context (written just before or after the image). Do not guess about context not evident.  Note this in "AI notes" 
- If there is a clear question in the context, ensure alt text can answer that question without noting answer. 

Return ONLY a valid JSON object with these exact keys: short, standard, long, notes, topic
- short, standard, and long are alt-text variants.
- topic is a one- or two-word high-level description of the image's
  mathematical or conceptual topic (for example: "Linear Graph",
  "Graph Theory", "Definitions", "Non-Math", "Coordinate System",
  "Right Triangle", "Parabola", "Translations", "Linear Inequality",
  "Step Function", "Logarithmic").
Do not include any markdown formatting or code blocks, just the raw JSON.
"""

def clean_pid(pid) -> str:
    return str(pid).strip()


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
    except Exception:
        pass
    return alt_map


def parse_relationships(docx_path: Path) -> Dict[str, str]:
    """
    Return only image/media relationships: rid -> target (e.g. 'media/image1.png'),
    and explicitly skip .wmf files.
    """
    rels_map: Dict[str, str] = {}

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
        target = (rel.get("Target") or "").strip()
        lower_target = target.lower()

        # Only keep image/media-ish things
        if "media/" not in lower_target and "image" not in lower_target:
            continue

        # ❌ Skip .wmf entirely
        if lower_target.endswith(".wmf"):
            continue

        # Keep full target path, e.g. 'media/image1.emf' or 'media/image1.png'
        rels_map[r_id] = target

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
    for drawing in root.findall(
        ".//w:drawing",
        namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    ):
        docpr = drawing.find(".//wp:docPr", namespaces=ns)
        blip = drawing.find(".//a:blip", namespaces=ns)
        if docpr is not None and blip is not None:
            r_id = blip.get(f"{{{ns['r']}}}embed")
            image_name = rels_map.get(r_id, "")
            pic_id = docpr.get("id", "")
            title = docpr.get("title", "")
            docpr_map[r_id] = {
                "picture_id": pic_id,
                "title": title,
                "image_file": image_name,
            }
    return docpr_map


def parse_rels(docx_path: Path) -> Dict[str, str]:
    rels = {}
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/_rels/document.xml.rels").decode("utf-8", errors="ignore")
            for rid, target in re.findall(r'Id="([^"]+)".+?Target="([^"]+)"', xml):
                if "image" in target.lower():
                    rels[rid] = target
    except Exception:
        pass
    return rels


def locate_all_image_instances(doc: Document) -> Dict[str, List[int]]:
    """
    Map relationship id (rId) -> list of paragraph index hints, including:

    • Images in normal paragraphs
    • Images inside tables (in any cell / nested table), even if they are
      not directly attached to a Run.

    The integer index is still based on doc.paragraphs, as in your original.
    """
    rid_to_paras: Dict[str, List[int]] = {}

    def _add_rids_from_element(elem, idx_hint: int) -> None:
        """Find all rIds for images under an arbitrary XML element."""
        # DrawingML images (a:blip)
        blips = elem.xpath(
            ".//*[local-name()='blip' and "
            "namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main']"
        )
        for blip in blips:
            rid = blip.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
            )
            if rid:
                rid_to_paras.setdefault(rid, []).append(idx_hint)

        # Legacy VML images (v:imagedata) — just in case
        imagedatas = elem.xpath(
            ".//*[local-name()='imagedata' and "
            "namespace-uri()='urn:schemas-microsoft-com:vml']"
        )
        for img in imagedatas:
            rid = img.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            )
            if rid:
                rid_to_paras.setdefault(rid, []).append(idx_hint)

    # --- 1. All "flat" paragraphs in the document ---
    for i, p in enumerate(doc.paragraphs):
        # Instead of iterating runs, search the entire paragraph XML subtree.
        _add_rids_from_element(p._element, i)

    # --- 2. All tables in the body, mapped back to a nearby paragraph index ---
    def inspect_table(table, row_idx_hint: int) -> None:
        # Scan the WHOLE table subtree (cells, nested tables, etc.)
        tbl_elm = table._element
        _add_rids_from_element(tbl_elm, row_idx_hint)

    for tbl in doc.tables:
        tbl_elm = tbl._element
        parent_idx = -1

        # Find the closest preceding paragraph that shares the same parent
        for idx, p in enumerate(doc.paragraphs):
            parent = p._element.getparent()
            if parent is None:
                continue

            for child in parent.iterchildren():
                if child is tbl_elm:
                    parent_idx = idx
                    break
            if parent_idx != -1:
                break

        # Fallback: treat as "last paragraph" if we can't find a good anchor
        if parent_idx == -1:
            parent_idx = len(doc.paragraphs) - 1 if doc.paragraphs else 0

        inspect_table(tbl, parent_idx)

    return rid_to_paras


def load_image_bytes(docx_path: Path, target: str):
    try:
        with zipfile.ZipFile(docx_path) as z:
            candidates = []

            # 1) Use the target as-is under word/
            candidates.append(f"word/{target}")

            # 2) If target is just "image1.png", also try word/media/image1.png
            name_only = Path(target).name
            candidates.append(f"word/media/{name_only}")

            for cand in candidates:
                if cand in z.namelist():
                    return z.read(cand)

            return None
    except Exception:
        return None


def convert_vector_to_png_bytes(
    input_bytes: bytes,
    suffix: str,
    temp_dir: Path,
) -> Optional[bytes]:
    temp_in = temp_dir / f"vector_input{suffix}"
    png_path = temp_dir / (temp_in.stem + ".png")

    try:
        temp_in.write_bytes(input_bytes)

        subprocess.run(
            [
                LIBREOFFICE_PATH,
                "--headless",
                "--convert-to",
                "png",
                "--outdir",
                str(temp_dir),
                str(temp_in),
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        if png_path.exists():
            return png_path.read_bytes()
        return None

    except Exception:
        return None

    finally:
        # Best effort cleanup
        for p in (temp_in, png_path):
            try:
                p.unlink(missing_ok=True)
            except Exception:
                pass


def resize_png(png_bytes: bytes) -> Tuple[bytes, Tuple[int, int]]:
    pil = PILImage.open(BytesIO(png_bytes))
    aspect = pil.height / pil.width
    new_w = TARGET_IMAGE_WIDTH_PX
    new_h = int(new_w * aspect)
    out = BytesIO()
    pil.resize((new_w, new_h), PILImage.Resampling.LANCZOS).save(
        out, format="PNG", dpi=(300, 300)
    )
    out.seek(0)
    return out.getvalue(), (new_w, new_h)

def create_thumbnail_from_bytes(
    img_bytes: bytes,
    max_width_px: int = 600,
    allow_upscale: bool = False,
) -> Tuple[bytes, int, int, int, int]:
    """
    Given raw image bytes (any format Pillow can open), return:
      - thumbnail PNG bytes
      - thumbnail width (px)
      - thumbnail height (px)
      - original width (px)
      - original height (px)

    If allow_upscale=False, we never enlarge images beyond their original width.
    """
    if not img_bytes:
        return b"", 0, 0, 0, 0

    pil = PILImage.open(BytesIO(img_bytes))
    orig_w, orig_h = pil.size

    if orig_w <= 0 or orig_h <= 0:
        return b"", 0, 0, orig_w, orig_h

    # Decide whether to resize
    if (not allow_upscale) and orig_w <= max_width_px:
        thumb = pil
        thumb_w, thumb_h = orig_w, orig_h
    else:
        aspect = orig_h / orig_w if orig_w else 1
        thumb_w = max_width_px
        thumb_h = int(thumb_w * aspect) if aspect > 0 else max_width_px
        thumb = pil.resize((thumb_w, thumb_h), PILImage.Resampling.LANCZOS)

    out = BytesIO()
    thumb.save(out, format="PNG", dpi=(300, 300))
    out.seek(0)
    return out.getvalue(), thumb_w, thumb_h, orig_w, orig_h

def build_prompt(grade_band: str, course_topic: str, context: str) -> str:
    return f"{STYLE_GUIDE_RULES}\nContext snippet: {context}\nReturn ONLY JSON."

def build_topic_classifier_prompt(context: str = "") -> str:
    """
    Build a prompt for classifying an image into a small set of math
    topic labels. Optionally include surrounding text as context.
    """
    ctx = clean_text(context, 400)
    if ctx:
        return f"{TOPIC_CLASSIFIER_RULES}\nContext snippet: {ctx}\nReturn ONLY JSON."
    else:
        return f"{TOPIC_CLASSIFIER_RULES}\nReturn ONLY JSON."


def normalize_topic_label(raw: str) -> str:
    """
    Map a free-form topic string from the model into one of the
    fixed TOPIC_LABELS. Fall back to 'Non-Math' if uncertain.
    """
    if not raw:
        return "Non-Math"

    raw_lower = raw.lower().strip()

    # Direct exact/substring matches to canonical labels
    mapping = {
        "definition": "Definitions",
        "non-math": "Non-Math",
        "coordinate": "Coordinate System",
        "axis": "Coordinate System",
        "axes": "Coordinate System",
        "linear graph": "Linear Graph",
        "line graph": "Linear Graph",
        "line": "Linear Graph",
        "exponential": "Exponential Graph",
        "growth": "Exponential Graph",
        "decay": "Exponential Graph",
        "linear system": "Linear Systems",
        "system of linear": "Linear Systems",
        "midpoint": "Midpoint Formula",
        "pythag": "Pythagorean Theorem",
        "right triangle": "Right Triangle",
        "triangle": "Right Triangle",
        "chart": "Chart Application",
        "table": "Chart Application",
        "bar chart": "Chart Application",
        "number line": "Number Line",
        "parabola": "Parabola",
        "quadratic": "Parabola",
        "translation": "Translations",
        "shift": "Translations",
        "inequality": "Linear Inequality",
        "step function": "Step Function",
        "piecewise": "Step Function",
        "logarithm": "Logarithmic",
        "log ": "Logarithmic",
        "graph theory": "Graph Theory",
        "network": "Graph Theory",
        "euler": "Graph Theory",
    }

    for key, label in mapping.items():
        if key in raw_lower:
            return label

    # If the model already returned an exact label name (case-insensitive)
    for label in TOPIC_LABELS:
        if label.lower() in raw_lower:
            return label

    return "Non-Math"



def _pixels_to_column_width(pixels: int) -> float:
    """
    Approximate conversion from pixels → Excel column width units.
    This is not exact, but works well enough for thumbnails.
    """
    if pixels <= 0:
        return 8.43  # default Excel width
    return (pixels - 5) / 7.0


def _pixels_to_row_height(pixels: int) -> float:
    """
    Approximate conversion from pixels → Excel row height (points).
    1 point ≈ 1.33 pixels → 1 px ≈ 0.75 pt
    """
    if pixels <= 0:
        return 15.0  # default Excel height
    return pixels * 0.75


def insert_png_bytes_in_cell(
    ws: Worksheet,
    row: int,
    col: int,
    *,
    png_bytes: bytes,
    width_px: int,
    height_px: int,
    adjust_row_col: bool = True,
    max_col_width_px: int = 100,
) -> Optional[ExcelImage]:
    """
    Insert a PNG (already resized) into a specific cell so that it moves with
    the row when you filter/sort.

    Column width will be increased as needed, but never beyond max_col_width_px.
    """
    if not png_bytes:
        return None

    buf = BytesIO(png_bytes)
    img = ExcelImage(buf)

    # Use the actual image size for the drawing itself
    img.width = width_px
    img.height = height_px

    col_letter = get_column_letter(col)
    cell_addr = f"{col_letter}{row}"

    # Anchor to the cell so it moves with filters/sorts
    img.anchor = cell_addr
    ws.add_image(img)

    if adjust_row_col:
        # ---- Column width (cap at max_col_width_px) ----
        desired_col_px = min(width_px, max_col_width_px)
        new_col_width = _pixels_to_column_width(desired_col_px)

        col_dim = ws.column_dimensions[col_letter]
        if not col_dim.width or col_dim.width < new_col_width:
            col_dim.width = new_col_width

        # ---- Row height (no cap; let it grow with the thumbnail) ----
        row_dim = ws.row_dimensions[row]
        new_row_height = _pixels_to_row_height(height_px)
        if not row_dim.height or row_dim.height < new_row_height:
            row_dim.height = new_row_height

    return img

def wrap_and_autofit_columns(
    ws: Worksheet,
    text_columns: List[str],
    max_col_width_px: int = 240,
    vertical: str = "top",
) -> None:
    """
    Wrap text in the given columns and approximate autofit,
    capping max width at max_col_width_px (in pixels).
    """
    # Wrap text
    for col_letter in text_columns:
        for cell in ws[col_letter]:
            cell.alignment = Alignment(wrap_text=True, vertical=vertical)

    # Autofit with max width cap
    max_excel_width = _pixels_to_column_width(max_col_width_px)

    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        if col_letter not in text_columns:
            continue

        max_len = 0
        for cell in col_cells:
            value = cell.value
            if value is not None:
                length = len(str(value))
                if length > max_len:
                    max_len = length

        if max_len == 0:
            continue

        desired_width = min(max_len + 2, max_excel_width)
        ws.column_dimensions[col_letter].width = desired_width
