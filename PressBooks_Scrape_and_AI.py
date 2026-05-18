# Scrapes images from a Pressbooks page and generates alt text variants.
# Writes to an Excel file
#
# Created by Brenda Gardner, May 17, 2026
#
# TO CUSTOMIZE:
# Step 1. Capture URLs and Subjects.
# a. Go to Website TOC and copy paste URLS for use in PRESSBOOKS_URLS.
# use in conjunction with https://slccbruins-my.sharepoint.com/:x:/g/personal/bgardn61_slcc_edu/IQD6AhzRzSxNRohDxUOKgeNYAQrWIctBAHw6ag8_bIYE-U8?e=sQFPWW
# Step 2. Set ABACUS_API_KEY as an environment variable, or paste directly below.
# Step 3. Update OUTPUT_FOLDER, OUTPUT_PREFIX, COURSE_GRADE, TOPIC to specify where the Excel file is saved
#
# End Step. After processing: Use macro to put the images in cell to allow filtering
"""
PressBooks_Scrape_and_AI.py
============================
Script 1 of 2 in the Pressbooks Accessibility Workflow.

WHAT THIS SCRIPT DOES
----------------------
1. Scrapes every page in PRESSBOOKS_URLS for <img> tags
   - will skip decorative UI elements (logos, banners, icons, headers/footers).
2. Downloads each unique image and generates alt-text variants using
   the Abacus AI vision API (short, medium, long + notes + topic).
3. Writes a single Excel review file in the Master 30-column layout
   (A–AD) — including full-size thumbnails in Column A.
4. Leaves Column W (New/Suggested Alt Text) BLANK for content experts to fill in by hand.

WORKFLOW
--------
  Step 1 → Run this script → send Excel to content expert(s)
  Step 2 → Expert fills Column W with approved alt text
  Step 3 → Run PressBooks_Push_AltText.py to push Column W back to Pressbooks

REQUIREMENTS
------------
  pip install requests beautifulsoup4 openpyxl Pillow

CREDENTIALS
-----------
Set ABACUS_API_KEY as an environment variable, or paste directly below.
"""

# ============================================================
# CONFIG — EDIT THIS SECTION BEFORE RUNNING
# ============================================================

# --- Pressbooks pages to scrape ---
# Add as many (url, section_label) tuples as you need.
# section_label becomes Column I (Section/Module/Figure #).
PRESSBOOKS_URLS = [
    ("https://slcc.pressbooks.pub/collegebiology1/back-matter/appendix/", "Appendix"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/1-2-levels-of-biological-organization/",
     "1.2 Levels of Biological Organization"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/1-3-properties-of-life/", "1.3 Properties of Life"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/22-1-the-energy-transformations-that-sustain-life/",
     "22.1 The Energy Transformations that Sustain Life"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/22-2-the-laws-of-thermodynamics/",
     "21.2 The Laws of Thermodynamics"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/25-1-capturing-light-energy/",
     "24.1 Overview of Photosynthesis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/4-1-the-composition-of-organisms/",
     "4.1 The Composition of Organisms"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/4-2-the-structure-of-atoms/", "4.2 The Structure of Atoms"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/4-3-chemical-bonds/", "4.3 Covalent Bonds"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/4-4-non-covalent-bonds/", "4.4 Non-covalent bonds"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/5-1-the-polar-nature-of-water/", "5.1 Water is Polar"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/5-2-waters-interactions-with-other-molecules/",
     "5.2 Water's Interactions with Other Molecules"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/5-3-properties-of-water/", "5.3 Other Properties of Water"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/5-4-ph-acids-and-bases/", "5.4 pH, Acids, and Bases"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/6-1-hydrocarbons/", "6.1 Hydrocarbons"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/6-2-functional-groups/", "6.2 Functional Groups"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/6-3-synthesis-of-biological-macromolecules/",
     "6.3 Synthesis of Biological Macromolecules"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/7-1-protein-function/", "7.1 Protein Function"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/7-2-amino-acids/", "7.2 Amino Acids"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/7-3-protein-structure/", "7.3 Protein Structure"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/7-4-protein-folding-regulation-and-denaturation/",
     "7.4 Protein Folding, Regulation, and Denaturation"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/8-1-metabolic-pathways/", "8.1 Metabolic Pathways"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/8-2-enzyme-function/", "8.2 Enzyme Function"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/8-3-molecular-regulation-of-enzymes/", "8.3 Enzyme Activity"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/8-4-feedback-inhibition-of-metabolic-pathways/",
     "8.4 Feedback Inhibition of Metabolic Pathways"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/active-transport/", "10.3 Active Transport"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/adaptive-evolution/", "3.3 Adaptive Evolution"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/atp-synthase/", "22.2 Chemiosmosis and ATP Synthase"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/bulk-transport/", "10.4 Bulk Transport"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/cancer-and-the-cell-cycle/",
     "17.4 Cancer and the Cell Cycle"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/cancer/", "17.3 Cancer"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/causes-of-mutation/", "15.2 Causes of Mutation"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/cell-cycle-checkpoints/", "17.1 Cell Cycle Checkpoints"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/cell-division-and-genomic-dna/",
     "16.1 Cell Division and Genomic DNA"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/central-dogma/", "13.1 The Central Dogma"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/chapter-1-what-is-life/",
     "1.1 Major Groups of Living Organisms"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/chapter-1/", "2.1 Scientific Inquiry"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/dna-replication/", "12.4 DNA Replication"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/dna-structure/", "12.2 DNA Structure"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/electron-transport-chain/", "22.3 Electron Transport Chains"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/energy/", "21.3 Energy"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/etcs-in-respiration-and-photosynthesis/",
     "22.4 Electron Transport Chains in Respiration and Photosynthesis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/eukaryotic-transcription/", "13.3 Eukaryotic Transcription"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/evidence-for-evolution/", "3.4 Evidence for Evolution"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/evolution/", "3.1 Darwin and Natural Selection"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/foundations-of-modern-genetics/",
     "19.1 Foundations of Modern Genetics"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/glycolysis/", "23.2 Glycolysis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/incomplete-dominance-and-codominance/",
     "20.1 Incomplete Dominance, Codominance, and Multiple Alleles"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/independent-assortment/", "19.4 Independent Assortment"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/life-cycles-of-sexually-reproducing-organisms/",
     "18.2 Life Cycles of Sexually Reproducing Organisms"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/ligands-and-receptors/", "11.1 Ligands and Receptors"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/light-energy/", "24.2 Light Energy"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/lipids/", "9.1 Lipids"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/meiosis/", "18.3 Meiosis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/metabolism/", "21.1 Metabolism"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/mitosis/", "16.3 Mitosis and Cytokinesis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/mrna-processing/", "13.4 mRNA Processing"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/nondisjunction/", "18.4 Nondisjunction"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/nucleic-acids/", "12.1 Functions of Nucleic Acids"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/organization-of-dna-in-the-cell/",
     "12.3 Organization of DNA in the cell"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/osmosis/", "10.2 Osmosis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/overview-of-cellular-respiration/",
     "23.1 Overview of Cellular Respiration"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/passive-transport/", "10.1 Passive Transport"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/pedigree-analysis/", "20.3 Pedigree Analysis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/polygenic-inheritance-and-epistasis/",
     "20.4 Polygenic Inheritance and Epistasis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/prokaryotic-transcription/",
     "13.2 Prokaryotic Transcription"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/punnett-squares-and-laws-of-probability/",
     "19.3 Punnett Squares and Laws of Probability"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/pyruvate-oxidation-and-the-citric-acid-cycle/",
     "23.3 Pyruvate Oxidation and the Citric Acid Cycle"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/regulator-molecules-of-the-cell-cycle/",
     "17.2 Regulator Molecules of the Cell Cycle"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/response-to-the-signal/",
     "11.4 Response to the Signal and Termination"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/segregation/", "19.2 Segregation"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/sex-linked-traits/", "20.2 Sex-linked traits"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/sexual-reproduction/", "18.1 Sexual Reproduction"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/signal-transduction/", "11.3 Signal Transduction"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-atp-cycle/", "21.4 The ATP cycle"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-calvin-cycle/", "24.4 The Calvin Cycle"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-cell-cycle/", "16.2 The Cell Cycle"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-cytoskeleton/", "16.4 The Cytoskeleton"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-endomembrane-system/", "9.3 The Endomembrane System"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-genetic-code/", "14.1 The Genetic Code"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-light-dependent-reactions/",
     "24.3 The Light-Dependent Reactions"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-mechanism-of-protein-synthesis/",
     "14.3 The Mechanism of Protein Synthesis"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-plasma-membrane/", "9.2 The Plasma Membrane"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/the-protein-synthesis-machinery/",
     "14.2 The Protein Synthesis Machinery"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/types-of-cell-signaling/", "11.2 Types of Cell Signaling"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/types-of-mutation/", "15.3 Types of Mutation"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/variation-adaptation/", "3.2 Variation and Adaptation"),
    ("https://slcc.pressbooks.pub/collegebiology1/chapter/variation/", "15.1 Variation and Evolution"),
    ("https://slcc.pressbooks.pub/collegebiology1/front-matter/introduction/", "Introduction"),
]

# --- Output ---
OUTPUT_FOLDER   = "/Users/brendagardner/Desktop"    # ← folder where Excel is saved
OUTPUT_PREFIX   = "CollegeBiology_AltText_Review"       # ← prefix for the output filename

# --- Course metadata (written into every row) ---
DEFAULT_GRADE_BAND = "12-14"
COURSE_TOPIC       = "College Biology I"

# --- Abacus AI ---
import os
ABACUS_API_KEY     = os.getenv("ABACUS_API_KEY", "")  # set env var or paste key here
ROUTE_LLM_ENDPOINT = "https://routellm.abacus.ai/v1/chat/completions"
VISION_MODEL       = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"
API_TIMEOUT_SEC    = 180

# ============================================================
# IMPORTS
# ============================================================

import base64
import datetime
import json
import re
import sys
import warnings
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.parse import urljoin, urlparse

try:
    import requests
    from bs4 import BeautifulSoup
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as ExcelImage
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    from PIL import Image as PILImage
except ImportError as e:
    sys.exit(
        f"Missing dependency: {e}\n"
        "Run:  pip install requests beautifulsoup4 openpyxl Pillow"
    )

try:
    from urllib3.exceptions import NotOpenSSLWarning
    warnings.simplefilter("ignore", NotOpenSSLWarning)
except ImportError:
    pass

# ============================================================
# CONSTANTS
# ============================================================

MASTER_HEADERS = [
    "Image Thumbnail",                                       # A  1 (embedded PNG)
    "Image Filename",                                        # B  2
    "Picture ID",                                            # C  3  ← blank (website source)
    "Image URL",                                             # D  4 (hyperlink)
    "Image Size",                                            # E  5
    "Image Topic",                                           # F  6 (AI-generated, 1–2 words)
    "Image Context",                                         # G  7 (surrounding paragraph text)
    "Page #",                                                # H  8 ← blank (website source)
    "Section/Module/Figure #",                               # I  9
    "Filename (Hyperlinked) or Webpage Name (URL)",          # J  10 (hyperlink to source page)
    "Grade Level",                                           # K  11
    "Course Name",                                           # L  12
    "Original Source Author",                                # M  13 ← blank
    "Copyright",                                             # N  14 ← blank
    "Date of Audit",                                         # O  15 (auto-filled)
    "Original Alt Text",                                     # P  16 (shown on Pressbooks today)
    "Original Alt Text Character Count",                     # Q  17 = LEN(P)
    "Original Alt Text Notes",                               # R  18 ← blank for reviewer
    "A.I. Short Alt Text",                                   # S  19
    "A.I. Med Alt Text",                                     # T  20
    "A.I. Long Alt Text",                                    # U  21
    "A.I. Notes",                                            # V  22
    "New/Suggested Alt Text",                                # W  23 ← CONTENT EXPERT FILLS THIS IN
    "New/Suggested Alt Text Count",                          # X  24 =LEN(W)
    "New/Suggested Alt Text Notes",                          # Y  25 ← blank for reviewer
    "Status - Alt Text Completed?\n(Original/Suggested/N)", # Z  26  ← Filled in in Part 2
    "Reviewer Notes",                                        # AA 27 ← blank for reviewer
    "Date of Completion",                                    # AB 30 ← blank
]

COLUMN_WIDTHS = {
    "A": 80,   # thumbnail — extra wide
    "B": 30, "C": 18, "D": 40, "E": 15, "F": 25, "G": 40,
    "H": 8,  "I": 25, "J": 40, "K": 12, "L": 20, "M": 20,
    "N": 20, "O": 16, "P": 50, "Q": 16, "R": 40, "S": 50,
    "T": 50, "U": 50, "V": 40, "W": 50, "X": 18, "Y": 40,
    "Z": 20, "AA": 30, "AB": 12, "AC": 12, "AD": 18,
}

ROW_HEIGHT_SCALE = 0.75   # pixels → Excel row-height points
MAX_THUMB_WIDTH  = 600     # pixels — max thumbnail width

STYLE_GUIDE = """
You are an accessibility expert and are a content expert in the course name. Generate alt text for images according to these rules:
- Be accurate and concise. Do not reveal answers not explicitly shown.
- Use screen-reader conventions:
  • Graphs and charts: include type, axis names/units if visible, describe trend, not exact numbers.
- Provide three variants per image:
  1) Short  (≤ 120 characters)
  2) Standard (≤ 250 characters)
  3) Long   (no limit; only for complex images)
- Adjust language by grade level:
- Do NOT include colors unless mentioned in context.
- Do NOT use "an image of" or similar phrases.
- Do NOT repeat information already in a surrounding caption.
- topic is a 1–2-word label (e.g. "Linear Graph", "Normal Curve", "Histogram").

Return ONLY a valid JSON object with exactly these keys:
  short, standard, long, notes, topic
No markdown, no code fences — raw JSON only.
"""
# --- Scraping behaviour ---
# Image src URL fragments that indicate decorative / non-content images.
# Any image whose absolute URL contains one of these strings is skipped.
SKIP_URL_PATTERNS = [
    "logo", "banner", "footer", "header",
    "favicon", "sprite", "icon", "avatar",
    "pressbooks-logo", "pb-logo",
]

# HTML parent tags that wrap decorative images (skip anything inside them).
SKIP_PARENT_TAGS = ["header", "footer", "nav"]

# HTML class/id values that indicate a decorative wrapper.
SKIP_CLASSES = [
    "site-header", "header", "footer", "site-footer",
    "footer-logo", "masthead", "page-header", "navbar",
]
SKIP_IDS = ["header", "footer", "colophon", "masthead", "navbar"]
# ============================================================
# HELPERS
# ============================================================

def _pixels_to_col_width(px: int) -> float:
    """Approximate Excel column-width units from pixel width."""
    return max(8.43, (px - 5) / 7.0)


def _clean(s: str, max_len: int = 0) -> str:
    s = re.sub(r"\s+", " ", (s or "")).strip()
    return s[:max_len] if max_len else s


def _make_session() -> requests.Session:
    sess = requests.Session()
    sess.headers.update({"User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    )})
    return sess


def _is_decorative(img_tag, full_url: str) -> bool:
    """Return True if this <img> should be skipped as decorative / UI chrome."""
    url_lower = full_url.lower()
    if any(pat in url_lower for pat in SKIP_URL_PATTERNS):
        return True
    # Check parent HTML elements
    for parent in img_tag.parents:
        tag_name = getattr(parent, "name", None)
        if tag_name in SKIP_PARENT_TAGS:
            return True
        if tag_name and tag_name not in ("#document", "[document]"):
            classes = parent.get("class") or []
            if any(c in SKIP_CLASSES for c in classes):
                return True
            if parent.get("id") in SKIP_IDS:
                return True
    return False


def _extract_context(img_tag, max_chars: int = 400) -> str:
    """Pull surrounding text: prefer <figcaption>, then nearby <p>/<li>."""
    def _text(node) -> str:
        t = node.get_text(" ", strip=True)
        return re.sub(r"\s+", " ", t).strip()[:max_chars]

    # 1) figcaption
    fig = img_tag.find_parent("figure")
    if fig:
        cap = fig.find("figcaption")
        if cap and _text(cap):
            return _text(cap)

    # 2) previous sibling blocks
    for sib in img_tag.previous_siblings:
        name = getattr(sib, "name", None)
        if name in ("p", "li", "div", "h1", "h2", "h3", "h4", "h5", "h6"):
            t = _text(sib)
            if t:
                return t

    # 3) Walk up and look at siblings of ancestors
    for ancestor in img_tag.parents:
        for sib in ancestor.previous_siblings:
            name = getattr(sib, "name", None)
            if name in ("p", "li", "div"):
                t = _text(sib)
                if t:
                    return t
        if getattr(ancestor, "name", None) in ("article", "main", "body"):
            break

    # 4) Next sibling fallback
    for sib in img_tag.next_siblings:
        name = getattr(sib, "name", None)
        if name in ("p", "li", "div"):
            t = _text(sib)
            if t:
                return t

    return ""


def _make_thumbnail(img_bytes: bytes) -> Tuple[bytes, int, int, int, int]:
    """
    Resize image to at most MAX_THUMB_WIDTH wide.
    Returns (thumb_bytes, thumb_w, thumb_h, orig_w, orig_h).
    """
    pil = PILImage.open(BytesIO(img_bytes)).convert("RGBA")
    orig_w, orig_h = pil.size

    if orig_w > MAX_THUMB_WIDTH:
        ratio   = MAX_THUMB_WIDTH / orig_w
        thumb_w = MAX_THUMB_WIDTH
        thumb_h = int(orig_h * ratio)
        pil     = pil.resize((thumb_w, thumb_h), PILImage.Resampling.LANCZOS)
    else:
        thumb_w, thumb_h = orig_w, orig_h

    out = BytesIO()
    pil.save(out, format="PNG")
    return out.getvalue(), thumb_w, thumb_h, orig_w, orig_h


def _insert_thumbnail(ws, row: int, png_bytes: bytes,
                      thumb_w: int, thumb_h: int) -> None:
    """Embed a PNG into Excel Column A, adjusting row height."""
    img = ExcelImage(BytesIO(png_bytes))
    img.width  = thumb_w
    img.height = thumb_h
    img.anchor = f"A{row}"
    ws.add_image(img)
    desired_h = max(thumb_h * ROW_HEIGHT_SCALE, 20)
    if not ws.row_dimensions[row].height or ws.row_dimensions[row].height < desired_h:
        ws.row_dimensions[row].height = desired_h

# ============================================================
# AI CALL
# ============================================================

def call_abacus_ai(img_bytes: bytes, context: str,
                   grade_band: str, course_topic: str) -> Dict[str, str]:
    """Call Abacus AI vision endpoint; return dict with short/standard/long/notes/topic."""
    if not ABACUS_API_KEY:
        print("  ⚠️  ABACUS_API_KEY not set — skipping AI call.")
        return {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}

    prompt = (
        f"{STYLE_GUIDE}\n"
        f"Grade level: {grade_band}\n"
        f"Course topic: {course_topic}\n"
        f"Context snippet: {_clean(context, 400)}\n"
        "Return ONLY JSON."
    )
    b64 = base64.b64encode(img_bytes).decode("utf-8")
    payload = {
        "model": VISION_MODEL,
        "messages": [
            {"role": "system",
             "content": "You are an accessibility expert. Return ONLY valid JSON."},
            {"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url",
                 "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    }
    headers = {
        "Authorization": f"Bearer {ABACUS_API_KEY}",
        "Content-Type": "application/json",
    }
    try:
        resp = requests.post(ROUTE_LLM_ENDPOINT, headers=headers,
                             json=payload, timeout=API_TIMEOUT_SEC)
        resp.raise_for_status()
        raw = (resp.json().get("choices", [{}])[0]
                          .get("message", {})
                          .get("content", ""))
        # Strip markdown fences if the model wrapped the JSON
        raw = re.sub(r"^```(?:json)?\s*", "", raw.strip())
        raw = re.sub(r"\s*```$", "", raw)
        js  = json.loads(raw)
        js.setdefault("topic", "")
        return js
    except Exception as e:
        print(f"  ⚠️  AI error: {e}")
        return {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}

# ============================================================
# SCRAPER
# ============================================================

def scrape_pressbooks_page(session: requests.Session,
                           url: str, section: str) -> List[Dict]:
    """
    Fetch one Pressbooks page and return a list of image-info dicts,
    one per non-decorative content image found.
    """
    print(f"\n🔍 Scraping: {url}")
    try:
        resp = session.get(url, timeout=20)
        resp.raise_for_status()
    except Exception as e:
        print(f"  ❌ Could not fetch page: {e}")
        return []

    soup = BeautifulSoup(resp.content, "html.parser")
    images = []
    seen_urls: set = set()

    for img in soup.find_all("img"):
        src = img.get("src", "").strip()
        if not src:
            continue

        full_url = urljoin(url, src)

        # Skip duplicates on the same page
        if full_url in seen_urls:
            continue

        # Skip decorative / UI chrome
        if _is_decorative(img, full_url):
            continue

        seen_urls.add(full_url)

        alt_text = (img.get("alt") or "").strip()
        context  = _extract_context(img)
        filename = Path(urlparse(full_url).path).name or "image.png"

        images.append({
            "image_url":  full_url,
            "filename":   filename,
            "alt_text":   alt_text,   # current alt on Pressbooks
            "context":    context,
            "section":    section,
            "webpage":    url,
        })

    print(f"  → {len(images)} content image(s) found.")
    return images

# ============================================================
# EXCEL BUILDER
# ============================================================

def _setup_worksheet(wb: Workbook) -> object:
    ws = wb.active
    ws.title = "Pressbooks Alt Text Review"

    for col_idx, header in enumerate(MASTER_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    # Column A gets a fixed wide width; others come from COLUMN_WIDTHS
    ws.column_dimensions["A"].width = _pixels_to_col_width(MAX_THUMB_WIDTH)
    for letter, px in COLUMN_WIDTHS.items():
        if letter != "A":
            ws.column_dimensions[letter].width = px

    return ws


def _write_row(ws, row: int, data: Dict, today: str) -> None:
    """Write one image's data into the worksheet at the given row."""
    def _cell(col: int, value, hyperlink: str = None, formula: str = None):
        c = ws.cell(row=row, column=col, value=value)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if hyperlink:
            c.hyperlink = hyperlink
            c.style = "Hyperlink"
        if formula:
            c.value = formula
        return c

    # B  Image Filename
    _cell(2,  data.get("filename", ""))
    # C  Picture ID — blank for website
    # D  Image URL
    _cell(4,  data.get("image_url", ""), hyperlink=data.get("image_url", ""))
    # E  Image Size
    _cell(5,  data.get("size_str", ""))
    # F  Image Topic (AI)
    _cell(6,  data.get("topic", ""))
    # G  Image Context
    _cell(7,  data.get("context", ""))
    # H  Page # — blank
    # I  Section
    _cell(9,  data.get("section", ""))
    # J  Webpage URL
    _cell(10, data.get("webpage", ""), hyperlink=data.get("webpage", ""))
    # K  Grade Level
    _cell(11, data.get("grade_band", DEFAULT_GRADE_BAND))
    # L  Course Name
    _cell(12, data.get("course_topic", COURSE_TOPIC))
    # M  Author — blank
    # N  Copyright — blank
    # O  Date of Audit
    _cell(15, today)
    # P  Original Alt Text (current on Pressbooks)
    _cell(16, data.get("alt_text", ""))
    # Q  =LEN(P)
    _cell(17, formula=f"=LEN(P{row})")
    # R  Alt Text Notes — blank for reviewer
    # S  AI Short
    _cell(19, _clean(data.get("ai_short",    ""), 250))
    # T  AI Medium
    _cell(20, _clean(data.get("ai_standard", ""), 600))
    # U  AI Long
    _cell(21, _clean(data.get("ai_long",     ""), 5000))
    # V  AI Notes
    _cell(22, _clean(data.get("ai_notes",    ""), 1000))
    # W  New/Suggested Alt Text — LEFT BLANK FOR REVIEWER
    # X  =LEN(W)
    _cell(24, formula=f"=LEN(W{row})")
    # Y  Suggested Notes — blank
    # Z  Status — blank
    # AA Reviewer Notes — blank
    # AB In Word? — blank
    # AC In Excel? — blank
    # AD Date of Completion — blank

# ============================================================
# MAIN
# ============================================================

def main():
    if not PRESSBOOKS_URLS:
        sys.exit("❌  PRESSBOOKS_URLS is empty.  Add at least one URL to scrape.")

    output_dir = Path(OUTPUT_FOLDER).expanduser()
    output_dir.mkdir(parents=True, exist_ok=True)

    session  = _make_session()
    wb       = Workbook()
    ws       = _setup_worksheet(wb)
    today    = datetime.date.today().strftime("%m/%d/%Y")
    excel_row = 2

    # Track seen image URLs across ALL pages so we don't duplicate
    global_seen: set = set()

    for page_url, section_label in PRESSBOOKS_URLS:
        images = scrape_pressbooks_page(session, page_url, section_label)

        for img_info in images:
            img_url = img_info["image_url"]

            if img_url in global_seen:
                print(f"  ⏭  Already processed: {img_info['filename']}")
                continue
            global_seen.add(img_url)

            print(f"  🖼  Downloading: {img_info['filename']}")

            # --- Download the image ---
            try:
                dl = session.get(img_url, timeout=20)
                dl.raise_for_status()
                img_bytes = dl.content
            except Exception as e:
                print(f"  ❌ Download failed: {e}")
                ws.cell(row=excel_row, column=1, value="(download failed)")
                ws.row_dimensions[excel_row].height = 30
                _write_row(ws, excel_row, img_info, today)
                excel_row += 1
                continue

            # --- Thumbnail ---
            try:
                thumb_bytes, tw, th, ow, oh = _make_thumbnail(img_bytes)
                img_info["size_str"] = f"{ow} × {oh} px"
                _insert_thumbnail(ws, excel_row, thumb_bytes, tw, th)
            except Exception as e:
                print(f"  ⚠️  Thumbnail error: {e}")
                ws.cell(row=excel_row, column=1, value="(no thumbnail)")
                ws.row_dimensions[excel_row].height = 30
                thumb_bytes = img_bytes  # still try to send raw to AI

            # --- AI alt text ---
            print(f"  🤖 Calling AI …")
            ai = call_abacus_ai(
                thumb_bytes,
                img_info.get("context", ""),
                DEFAULT_GRADE_BAND,
                COURSE_TOPIC,
            )
            img_info["ai_short"]    = ai.get("short",    "")
            img_info["ai_standard"] = ai.get("standard", "")
            img_info["ai_long"]     = ai.get("long",     "")
            img_info["ai_notes"]    = ai.get("notes",    "")
            img_info["topic"]       = ai.get("topic",    "")

            # --- Write row ---
            _write_row(ws, excel_row, img_info, today)
            excel_row += 1

    # --- Freeze pane + autofilter ---
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:AD{excel_row - 1}"

    # --- Save ---
    timestamp   = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    output_path = output_dir / f"{OUTPUT_PREFIX}_{timestamp}.xlsx"
    wb.save(output_path)

    total = excel_row - 2
    print(f"\n✅  Done — {total} image(s) processed.")
    print(f"   Excel saved to: {output_path}")
    print(
        "\nNEXT STEPS:"
        "\n  1. Open the Excel file."
        "\n  2. Review Columns S–U (AI alt text variants)."
        "\n  3. Enter your approved alt text in Column W (New/Suggested Alt Text)."
        "\n  4. Run PressBooks_Push_AltText.py to push Column W back to Pressbooks."
    )


if __name__ == "__main__":
    main()