
#Created Nov 24 to capture/audit the current Alt Text from DOCX files in preferred Excel format
#columns match output excel file from Alt Text to Word Report

"""
Single DOCX → Excel (A–V headers). No API calls.
Fills in:
A: Thumbnail, B: Image Filename, C: Picture ID (what is eventually matched), D: File Name/Source
E: Image Topic (blank)
F: Section/Module
G = Page Number
H = In Word? (auto "Yes")
I = In Excel? (blank)
J = Current Alt Text
K = Alt Text Count
L = Notes (blank)
"""


from pathlib import Path
from io import BytesIO
import zipfile, re, datetime, subprocess
from typing import Dict, List, Tuple, Union
from docx import Document
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from lxml import etree

# ✅ Hardcoded input and output paths
INPUT_DOCX_PATH = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Shared Alt Text Files/1010 Completed Lecture Notes/Math 1010 Lecture Notes Module 3_withAltText.docx")
OUTPUT_FOLDER = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1010 Information/1010 Test Files")

LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
TARGET_IMAGE_WIDTH_PX = 800
ROW_HEIGHT_SCALE = 0.85

def guess_page_number(paragraph_index: Union[int, None]) -> Union[int, None]:
    return (paragraph_index // 20) + 1 if paragraph_index and paragraph_index >= 0 else None

def extract_section_from_heading(paragraphs: List, p_idx: int) -> str:
    if p_idx is None or p_idx < 0 or p_idx >= len(paragraphs):
        return ""
    for i in range(p_idx, -1, -1):
        para = paragraphs[i]
        if para.style and para.style.name == "Heading 1":
            text = para.text.strip()
            m = re.match(r"^([A-Za-z]*\s*\d+(\.\d+)*)(\s|$)", text)
            return m.group(1) if m else text
    return ""

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

def parse_relationships(docx_path: Path) -> Dict[str, str]:
    rels_map = {}
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        rels_path = "word/_rels/document.xml.rels"
        if rels_path not in docx_zip.namelist():
            return rels_map
        with docx_zip.open(rels_path) as f:
            root = etree.fromstring(f.read())
            ns = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
            for rel in root.findall("r:Relationship", namespaces=ns):
                r_id = rel.get("Id")
                target = rel.get("Target")
                if target and target.startswith("media/"):
                    rels_map[r_id] = target
    return rels_map

def extract_docpr_tags_with_rid(docx_path: Path) -> Dict[str, Dict[str, str]]:
    ns = {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    rels_map = parse_relationships(docx_path)
    docpr_map = {}
    with zipfile.ZipFile(docx_path, "r") as docx_zip:
        root = etree.fromstring(docx_zip.read("word/document.xml"))
        for drawing in root.findall(".//w:drawing", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
            docpr = drawing.find(".//wp:docPr", namespaces=ns)
            blip = drawing.find(".//a:blip", namespaces=ns)
            if docpr is not None and blip is not None:
                r_id = blip.get(f"{{{ns['r']}}}embed")
                docpr_map[r_id] = {"picture_id": docpr.get("id", ""), "title": docpr.get("title", ""), "image_file": rels_map.get(r_id, "")}
    return docpr_map

def locate_all_image_instances(doc: Document) -> Dict[str, List[int]]:
    rid_to_paras: Dict[str, List[int]] = {}
    for i, p in enumerate(doc.paragraphs):
        for run in p.runs:
            for blip in run._element.xpath(".//*[local-name()='blip' and namespace-uri()='http://schemas.openxmlformats.org/drawingml/2006/main']"):
                rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if rid:
                    rid_to_paras.setdefault(rid, []).append(i)
    return rid_to_paras

def load_image_bytes(docx_path: Path, target: str) -> Union[bytes, None]:
    try:
        with zipfile.ZipFile(docx_path) as z:
            return z.read(f"word/{target}") if f"word/{target}" in z.namelist() else None
    except:
        return None

def convert_vector_to_png_bytes(input_bytes: bytes, suffix: str, temp_dir: Path) -> Union[bytes, None]:
    temp_in = temp_dir / f"vector_input{suffix}"
    temp_in.write_bytes(input_bytes)
    try:
        subprocess.run([LIBREOFFICE_PATH, "--headless", "--convert-to", "png", "--outdir", str(temp_dir), str(temp_in)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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

def process_single_docx(docx_path: Path, output_folder: Path):
    doc = Document(str(docx_path))
    paragraphs = doc.paragraphs
    alt_map = parse_alt_text(docx_path)
    rels = parse_relationships(docx_path)
    rid_para_map = locate_all_image_instances(doc)
    docpr_map = extract_docpr_tags_with_rid(docx_path)
    base_name = docx_path.name

    wb = Workbook()
    ws = wb.active
    ws.title = "Alt Text Audit"

    headers = ["Thumbnail", "Image Filename", "Picture ID", "File Name/Source", "Image Topic", "Section/Module", "Page Number", "In Word?", "In Excel?", "Current Alt Text", "Alt Text Count", "Notes"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h).font = Font(bold=True)

    row = 2
    openpyxl_image_refs = []
    for rid, target in rels.items():
        img_bytes = load_image_bytes(docx_path, target)
        if not img_bytes:
            continue
        suffix = Path(target).suffix.lower()
        if suffix in (".emf", ".svg", ".wmf"):
            converted = convert_vector_to_png_bytes(img_bytes, suffix, docx_path.parent)
            img_bytes = converted if converted else None
        thumb_bytes = None
        if img_bytes:
            try:
                pil = PILImage.open(BytesIO(img_bytes))
                if pil.format != "PNG":
                    buf = BytesIO(); pil.save(buf, format="PNG"); img_bytes = buf.getvalue()
                thumb_bytes, (w, h) = resize_png(img_bytes)
            except:
                thumb_bytes = None
        para_indices = rid_para_map.get(rid, [-1])
        for p_idx in para_indices:
            if thumb_bytes:
                img_stream = BytesIO(thumb_bytes)
                excel_img = ExcelImage(img_stream)
                openpyxl_image_refs.append((img_stream, excel_img))
                ws.add_image(excel_img, f"A{row}"); ws.row_dimensions[row].height = int(h * ROW_HEIGHT_SCALE)
            else:
                ws.cell(row=row, column=1, value="[No Thumbnail]")

            ws.cell(row=row, column=2, value=Path(target).name)
            ws.cell(row=row, column=3, value=docpr_map.get(rid, {}).get("picture_id", ""))
            ws.cell(row=row, column=4, value=base_name)
            ws.cell(row=row, column=5, value="")
            ws.cell(row=row, column=6, value=extract_section_from_heading(paragraphs, p_idx))
            ws.cell(row=row, column=7, value=guess_page_number(p_idx))
            ws.cell(row=row, column=8, value="Yes")
            ws.cell(row=row, column=9, value="")
            ws.cell(row=row, column=10, value=existing_alt)
            ws.cell(row=row, column=11, value=f"=LEN({get_column_letter(11)}{row})")
            ws.cell(row=row, column=12, value="")
            existing_alt = alt_map.get(rid, "").strip()

            row += 1

    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max((len(str(cell.value)) for cell in col_cells if cell.value), default=0)
        for cell in col_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)
    ws.column_dimensions["A"].width = 30

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_xlsx = output_folder / f"{docx_path.stem}_Audit_{timestamp}.xlsx"
    wb.save(output_xlsx)
    print(f"✅ Excel saved to: {output_xlsx}")

# Run
process_single_docx(INPUT_DOCX_PATH, OUTPUT_FOLDER)
