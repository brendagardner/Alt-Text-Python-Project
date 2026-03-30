#!/usr/bin/env python3
"""
Convert a Microsoft Word (.docx) file to Pressbooks-ready HTML.
Version 2 – adds free MathType OLE equation → LaTeX conversion.

USAGE
-----
    python3 1035_WordToPressbooksV2.py "My Chapter.docx"


HOW MATHTYPE EQUATIONS ARE CONVERTED (no paid API needed)
----------------------------------------------------------
When MathType equations are inserted with alt text enabled, MathType
stores the equation as MathML inside the image's alt attribute, in the
format:
    {"mathml": "<math ...>...</math>"}

This script reads that MathML, converts it to LaTeX, and replaces the
image tag in the HTML with the correct  [latex]...[/latex]  or  \\(…\\)
notation — entirely free, with no external API.

  ✔  Fractions       → \\frac{}{}
  ✔  Cancelled terms → \\cancel{}  (requires MathJax cancel extension)
  ✔  Text units      → \\text{minutes}, \\text{miles}, etc.
  ✔  Greek letters   → \\alpha, \\pi, etc.
  ✔  Exponents       → ^{}
  ✔  Subscripts      → _{}
  ✔  Square roots    → \\sqrt{}
  ✔  Trig / log      → \\sin, \\cos, \\log, etc.

If an equation image has no MathML in its alt text (older MathType
equations saved before alt-text was added), the script inserts a
highlighted [TODO] marker so you can fill in the LaTeX manually.

Optional: Set MATHPIX_APP_ID and MATHPIX_APP_KEY to use the MathPix API
as a fallback for any images that could not be converted from MathML.
Sign up free at https://mathpix.com.

Pandoc also converts Word's native OMML equations to LaTeX automatically:
    Inline  → \\( e^{i\\pi}+1=0 \\)
    Display → \\[ e^{i\\pi}+1=0 \\]

====================================================================
WHAT THIS SCRIPT DOES WELL (automated, reliable)
====================================================================
✔  Converts MathType equation images → LaTeX via MathML alt text (free)
✔  Falls back to MathPix API if configured (optional, for older equations)
✔  Falls back to [TODO] markers for any unconverted equation images
✔  Runs pandoc to convert OMML equations → \\(...\\) / \\[...\\]
✔  Strips pandoc's wrapper HTML; keeps only body content
✔  Removes the Word-generated Table of Contents block
✔  Removes legacy _Toc anchor IDs left by Word/pandoc
✔  Fixes file:///...#anchor links → plain #anchor links
✔  Flattens all <blockquote> wrappers (keeps content)
✔  Builds the "Section Topics" learning-objectives callout from H2s
✔  H3 headings emitted as plain <h3> subheadings
✔  Wraps ALL-CAPS bold paragraphs as procedure boxes (key-takeaways)
✔  Detects purple ListParagraph items → individual key-takeaways boxes
✔  Detects "Example N" headings AND heading-styled examples (h3/h4/h5) and
   wraps all following content into textbox--examples boxes
✔  Auto-inserts <hr /><strong>Solution</strong> after the first paragraph
   of every example box (soup-level, before the state machine)
✔  Converts grouped display math \\[{eq1}{eq2}\\] → align* environment
✔  Merges consecutive single-equation display paragraphs → align* block
✔  Converts single-letter/digit <em>x</em> → [latex]x[/latex]
✔  Normalizes all <table> elements to width: 100%
✔  Closes open boxes when a new section/box type begins
✔  Slugifies H2 text for id= anchors (needed for Pressbooks TOC links)
✔  Outputs clean, indented HTML ready to paste into Pressbooks

====================================================================
WHAT NEEDS HUMAN REVIEW AFTER RUNNING  (~15–45 min)
====================================================================
✘  [TODO] MARKERS  – Search the HTML for [TODO] and replace each with
   the correct LaTeX.  These mark MathType equations that had no MathML
   in their alt text.  To avoid them: in MathType, make sure "Copy as
   MathML" / accessibility alt text is enabled before saving your Word doc.

✘  COMPLEX FRACTIONS  – Pandoc renders most fractions correctly, but
   deeply nested or unusual fraction layouts may need a manual check.

✘  BOX BOUNDARIES  – Inferred from text cues.  Fix manually if wrong.

✘  IMAGES  – Pandoc extracts images to a media/ folder.  Move it next
   to the final HTML and verify every image displays correctly.

✘  TABLES  – Check header rows (<th>) and merged cells.

✘  UNNESTED LISTS  – Check that list items weren't merged into prose.

✘  INDENTED EQUATIONS  – Add padding manually:
   <p style="padding-left: 40px">[latex]...[/latex]</p>

====================================================================
PRESSBOOKS BOX TEMPLATES  (copy-paste as needed)
====================================================================
<!-- SECTION TOPICS (auto-generated) -->
<div class="textbox textbox--learning-objectives">
  <header class="textbox__header">
    <p class="textbox__title">Section Topics</p>
  </header>
  <div class="textbox__content">
    <div>
      <p style="padding-left: 40px"><a href="#anchor">Heading Text</a></p>
    </div>
  </div>
</div>

<!-- KEY OBJECTIVES / PROCEDURE / KEY-TAKEAWAYS BOX -->
<div class="textbox textbox--key-takeaways">
  <header class="textbox__header">
    <p class="textbox__title">TITLE</p>
  </header>
  <div class="textbox__content">
    <!-- steps / definitions go here -->
  </div>
</div>

<!-- EXAMPLE BOX -->
<div class="textbox textbox--examples">
  <header class="textbox__header">
    <p class="textbox__title">Example N</p>
  </header>
  <div class="textbox__content">
    <!-- problem statement -->
    <hr />
    <strong>Solution</strong>
    <!-- solution steps -->
  </div>
</div>

<!-- TIP / SIDEBAR BOX -->
<div class="textbox textbox--tip">
  <header class="textbox__header">
    <p class="textbox__title">Tip Title</p>
  </header>
  <div class="textbox__content">
    <!-- tip content -->
  </div>
</div>
"""

# ====================================================================
# RUN SETTINGS  –  fill these in before clicking Run in PyCharm
# ====================================================================
INPUT_FILE  = "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/MathType_Test.docx"
OUTPUT_FILE = "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/TestMathtype_PB_HTML_V2.html"

# ====================================================================
# MATHPIX SETTINGS  –  optional fallback for equations without MathML
# ====================================================================
# The MathML-from-alt-text converter runs first and is completely free.
# MathPix is only used as a fallback for older MathType equations that
# were saved without MathML alt text.  Leave both as "" to skip MathPix.
# Sign up free at https://mathpix.com
MATHPIX_APP_ID  = ""   # e.g. "yourname_abc123"
MATHPIX_APP_KEY = ""   # e.g. "1234567890abcdef"

# ====================================================================
# USER SETTINGS  –  adjust these to match your Word document style
# ====================================================================

# Heading level offset ────────────────────────────────────────────────
# Set SHIFT_HEADINGS_DOWN = 1  if your section headings are Heading 1
#                                 in Word  (most full-chapter textbook files)
# Set SHIFT_HEADINGS_DOWN = 0  if your section headings are already
#                                 Heading 2 in Word
SHIFT_HEADINGS_DOWN = 0

# Words that mark a "Solution" break inside an example box.
SOLUTION_WORDS = {"solution", "solutions"}

# Hex color code (without #) used for Key Objective box headers in Word.
KEY_OBJECTIVE_COLOR = "7030A0"  # verified color in Word for Oremus files

# ====================================================================
# IMPORTS
# ====================================================================
import json
import re
import sys
import subprocess
import textwrap
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    from bs4 import BeautifulSoup, NavigableString, Tag
except ImportError:
    sys.exit(
        "ERROR: BeautifulSoup4 is not installed.\n"
        "Run:  pip install beautifulsoup4 --break-system-packages\n"
        "Then re-run this script."
    )

# ====================================================================
# COMPILED PATTERNS
# ====================================================================
_PROC_TITLE_RE  = re.compile(r'^[A-Z][A-Z\s\d,\-:()]+$')
_EXAMPLE_RE     = re.compile(r'^\s*example\s+(\d+)\b', re.I)
_TOC_ID_RE      = re.compile(r'^_Toc\d+$', re.IGNORECASE)
_SINGLE_CHAR_RE = re.compile(r'^[A-Za-z0-9]$')
_SOFT_NL_RE     = re.compile(r'[ \t]*\n[ \t]*')
_DISPLAY_P_RE   = re.compile(r'^\\\[(.*)\\\]$', re.S)   # matches \[...\]


# ====================================================================
# MATHML → LATEX CONVERTER  (no external API needed)
# ====================================================================

# Operator Unicode character → LaTeX
_MO_MAP = {
    '+': '+',  '-': '-',  '×': r'\times',  '÷': r'\div',
    '·': r'\cdot',  '\u00B7': r'\cdot',      # middle dot
    '=': '=',  '≤': r'\leq',  '≥': r'\geq',  '≠': r'\neq',
    '<': '<',  '>': '>',
    '≈': r'\approx',  '≡': r'\equiv',  '~': r'\sim',
    '∞': r'\infty',   '∂': r'\partial', '∇': r'\nabla',
    '∑': r'\sum',     '∏': r'\prod',
    '∫': r'\int',     '∬': r'\iint',    '∭': r'\iiint',
    '→': r'\to',      '←': r'\leftarrow',  '↔': r'\leftrightarrow',
    '⇒': r'\Rightarrow',  '⇐': r'\Leftarrow',  '⇔': r'\Leftrightarrow',
    '∈': r'\in',   '∉': r'\notin',
    '⊂': r'\subset',  '⊃': r'\supset',
    '⊆': r'\subseteq',  '⊇': r'\supseteq',
    '∪': r'\cup',  '∩': r'\cap',  '∅': r'\emptyset',
    '∀': r'\forall',  '∃': r'\exists',
    '|': r'|',  '‖': r'\|',  '±': r'\pm',  '∓': r'\mp',
    '⌊': r'\lfloor',  '⌋': r'\rfloor',
    '⌈': r'\lceil',   '⌉': r'\rceil',
    '%': r'\%',  '.': '.',
    # Invisible/spacing operators — suppress or convert to space
    '\u2061': '',   # function application
    '\u2062': '',   # invisible times
    '\u2063': '',   # invisible separator
    '\u2064': '',   # invisible plus
    '\u00A0': ' ',  # non-breaking space
    '\u2009': r'\,', # thin space
}

# Greek letter → LaTeX (used in <mi> elements)
_GREEK_MAP = {
    'α': r'\alpha',   'β': r'\beta',    'γ': r'\gamma',   'δ': r'\delta',
    'ε': r'\epsilon', 'ζ': r'\zeta',    'η': r'\eta',     'θ': r'\theta',
    'ι': r'\iota',    'κ': r'\kappa',   'λ': r'\lambda',  'μ': r'\mu',
    'ν': r'\nu',      'ξ': r'\xi',      'π': r'\pi',      'ρ': r'\rho',
    'σ': r'\sigma',   'τ': r'\tau',     'υ': r'\upsilon', 'φ': r'\phi',
    'χ': r'\chi',     'ψ': r'\psi',     'ω': r'\omega',
    'Γ': r'\Gamma',   'Δ': r'\Delta',   'Θ': r'\Theta',   'Λ': r'\Lambda',
    'Ξ': r'\Xi',      'Π': r'\Pi',      'Σ': r'\Sigma',   'Υ': r'\Upsilon',
    'Φ': r'\Phi',     'Ψ': r'\Psi',     'Ω': r'\Omega',
}

# Function names that get a backslash in LaTeX (e.g. \sin, \log)
_FUNC_NAMES = {
    'sin', 'cos', 'tan', 'csc', 'sec', 'cot',
    'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh',
    'log', 'ln', 'exp',
    'lim', 'max', 'min', 'sup', 'inf',
    'det', 'dim', 'ker', 'gcd', 'lcm', 'mod', 'deg',
}


def _mml_local_tag(node) -> str:
    """Return local element name, stripping any XML namespace prefix."""
    tag = node.tag
    return tag.split('}', 1)[1] if isinstance(tag, str) and '}' in tag else tag


def _mml_children_to_latex(children: list) -> str:
    """
    Convert a list of MathML child nodes to LaTeX.
    Special rule: a run of 3+ consecutive single-letter <mi> elements
    (each containing one alphabetic character) is treated as a text label
    and wrapped in \\text{...} rather than as individual math variables.
    This handles MathType's habit of encoding units letter-by-letter,
    e.g. <mi>m</mi><mi>i</mi><mi>n</mi><mi>u</mi><mi>t</mi><mi>e</mi><mi>s</mi>
    → \\text{minutes}
    """
    parts = []
    i = 0
    while i < len(children):
        node = children[i]
        tag  = _mml_local_tag(node)
        text = (node.text or '').strip()

        # Detect a run of single-letter alphabetic <mi> nodes (= a word/unit)
        if (tag == 'mi'
                and len(text) == 1
                and text.isalpha()
                and text not in _GREEK_MAP):
            run = [text]
            j = i + 1
            while j < len(children):
                nt = _mml_local_tag(children[j])
                nx = (children[j].text or '').strip()
                if (nt == 'mi'
                        and len(nx) == 1
                        and nx.isalpha()
                        and nx not in _GREEK_MAP):
                    run.append(nx)
                    j += 1
                else:
                    break
            if len(run) >= 3:
                # 3+ letters → text label (unit, word)
                parts.append(r'\text{' + ''.join(run) + '}')
                i = j
            else:
                # 1–2 letters → keep as individual variables
                for k in range(len(run)):
                    parts.append(_mml_node_to_latex(children[i + k]))
                i += len(run)
        else:
            parts.append(_mml_node_to_latex(node))
            i += 1

    return ''.join(parts)


def _mml_node_to_latex(node) -> str:
    """Recursively convert one MathML XML element to a LaTeX string."""
    tag  = _mml_local_tag(node)
    text = (node.text or '').strip()
    ch   = list(node)

    def c(i: int) -> str:
        return _mml_node_to_latex(ch[i]) if i < len(ch) else ''

    def a() -> str:
        return _mml_children_to_latex(ch)

    # ── Root / transparent grouping ──────────────────────────────────
    if tag in ('math', 'mrow', 'mphantom', 'mstyle', 'mpadded'):
        return a()

    # ── Token elements ───────────────────────────────────────────────
    if tag == 'mi':
        if text in _GREEK_MAP:
            return _GREEK_MAP[text]
        if text.lower() in _FUNC_NAMES:
            return '\\' + text.lower()
        if len(text) > 1:
            return r'\mathrm{' + text + '}'
        return text

    if tag == 'mn':
        return text

    if tag == 'mo':
        if text in _MO_MAP:
            return _MO_MAP[text]
        # Stretchy brackets
        bracket_map = {
            '(': r'\left(', ')': r'\right)',
            '[': r'\left[', ']': r'\right]',
            '{': r'\left\{', '}': r'\right\}',
        }
        if text in bracket_map:
            return bracket_map[text]
        return text

    if tag == 'mtext':
        t = text or ''.join((x.text or '') for x in ch)
        return r'\text{' + t.strip() + '}' if t.strip() else ''

    if tag == 'mspace':
        return r'\,'

    # ── Fractions ────────────────────────────────────────────────────
    if tag == 'mfrac':
        num = c(0)
        den = c(1)
        if node.get('linethickness', '1') == '0':
            return r'\binom{' + num + '}{' + den + '}'
        return r'\frac{' + num + '}{' + den + '}'

    # ── Scripts ─────────────────────────────────────────────────────
    if tag == 'msup':
        return c(0) + '^{' + c(1) + '}'

    if tag == 'msub':
        return c(0) + '_{' + c(1) + '}'

    if tag == 'msubsup':
        return c(0) + '_{' + c(1) + '}^{' + c(2) + '}'

    if tag == 'mmultiscripts':
        return c(0)  # simplified: just the base

    # ── Radicals ─────────────────────────────────────────────────────
    if tag == 'msqrt':
        return r'\sqrt{' + a() + '}'

    if tag == 'mroot':
        return r'\sqrt[' + c(1) + ']{' + c(0) + '}'

    # ── Under / Over decorations ─────────────────────────────────────
    if tag == 'munder':
        return c(0) + '_{' + c(1) + '}'   # handles lim, sum, etc.

    if tag == 'mover':
        accent_map = {
            '→': r'\vec', '⃗': r'\vec', '˜': r'\tilde',
            '˙': r'\dot', '¨': r'\ddot', '¯': r'\bar', 'ˆ': r'\hat',
            '`': r'\grave', '´': r'\acute',
        }
        over_text = (ch[1].text or '').strip() if len(ch) > 1 else ''
        if over_text in accent_map:
            return accent_map[over_text] + '{' + c(0) + '}'
        if over_text == '⏞':
            return r'\overbrace{' + c(0) + '}'
        return r'\overset{' + c(1) + '}{' + c(0) + '}'

    if tag == 'munderover':
        return c(0) + '_{' + c(1) + '}^{' + c(2) + '}'

    # ── Enclosures (strikethrough, box, etc.) ────────────────────────
    if tag == 'menclose':
        notation = node.get('notation', '')
        inner = a()
        if 'updiagonalstrike' in notation or 'downdiagonalstrike' in notation:
            return r'\cancel{' + inner + '}'
        if 'box' in notation:
            return r'\boxed{' + inner + '}'
        if 'radical' in notation:
            return r'\sqrt{' + inner + '}'
        if 'horizontalstrike' in notation:
            return r'\sout{' + inner + '}'
        return inner  # unknown notation — show content as-is

    # ── Fenced expressions  (parentheses, brackets, etc.) ────────────
    if tag == 'mfenced':
        od  = node.get('open', '(')
        cd  = node.get('close', ')')
        sep = node.get('separators', ',')
        inner = sep.join(_mml_node_to_latex(x) for x in ch)
        open_map  = {'(': r'\left(', '[': r'\left[',  '{': r'\left\{',
                     '|': r'\left|',  '': ''}
        close_map = {')': r'\right)', ']': r'\right]', '}': r'\right\}',
                     '|': r'\right|', '': ''}
        op = open_map.get(od,  r'\left'  + od)
        cl = close_map.get(cd, r'\right' + cd)
        return op + inner + cl

    # ── Tables / matrices ────────────────────────────────────────────
    if tag == 'mtable':
        rows = [x for x in ch if _mml_local_tag(x) == 'mtr']
        row_strs = []
        for row in rows:
            cells = [x for x in row if _mml_local_tag(x) == 'mtd']
            row_strs.append(
                ' & '.join(_mml_children_to_latex(list(cell)) for cell in cells)
            )
        return r'\begin{matrix}' + r' \\ '.join(row_strs) + r'\end{matrix}'

    if tag in ('mtr', 'mtd'):
        return a()

    # ── Fallback ─────────────────────────────────────────────────────
    return text + a()


def mathml_to_latex(mathml_str: str):
    """
    Convert a MathML XML string to a LaTeX string.
    Returns None if the conversion fails (bad XML, empty result, etc.).
    Uses only Python's built-in xml.etree.ElementTree — no extra installs.
    """
    try:
        ET.register_namespace('', 'http://www.w3.org/1998/Math/MathML')
        root   = ET.fromstring(mathml_str)
        result = _mml_node_to_latex(root).strip()
        return result if result else None
    except ET.ParseError:
        return None
    except Exception:
        return None


# ====================================================================
# STEP 0 – PRE-SCAN DOCX FOR PURPLE LIST PARAGRAPH HEADERS
# ====================================================================

def extract_purple_list_headers(docx_path: Path) -> set:
    """
    Open the Word XML directly (no python-docx required) and return the
    set of paragraph texts that use both the 'ListParagraph' style AND
    the KEY_OBJECTIVE_COLOR font color.
    """
    headers = set()
    try:
        with zipfile.ZipFile(str(docx_path)) as z:
            xml = z.read("word/document.xml").decode("utf-8")
    except Exception as e:
        print(f"  Warning: could not scan docx for Key Objective headers: {e}")
        return headers

    color = KEY_OBJECTIVE_COLOR.upper()
    paras = re.findall(r'<w:p[ >].*?</w:p>', xml, re.DOTALL)
    for p in paras:
        if 'ListParagraph' in p and color in p.upper():
            text = ' '.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', p)).strip()
            if text:
                headers.add(text)
    return headers


# ====================================================================
# STEP 0b – DETECT MATHTYPE OLE EQUATIONS IN DOCX
# ====================================================================

def detect_mathtype_equations(docx_path: Path) -> int:
    """Count MathType OLE objects (ProgID="Equation.3") in the docx."""
    try:
        with zipfile.ZipFile(str(docx_path)) as z:
            xml = z.read("word/document.xml").decode("utf-8")
        return (xml.count('ProgID="Equation.3"')
                + xml.count("ProgID='Equation.3'"))
    except Exception as e:
        print(f"  Warning: could not scan docx for MathType equations: {e}")
        return 0


# ====================================================================
# STEP 1 – RUN PANDOC
# ====================================================================

def run_pandoc(docx_path: Path) -> str:
    """Call pandoc and return the HTML string it produces."""
    cmd = [
        "pandoc",
        str(docx_path),
        "--from", "docx",
        "--to", "html",
        "--mathjax",
        "--wrap=none",
        "--strip-comments",
        "--extract-media", str(docx_path.parent / "media"),
    ]
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, check=True
        )
    except FileNotFoundError:
        sys.exit(
            "ERROR: pandoc is not found on your PATH.\n"
            "Install pandoc from https://pandoc.org/installing.html\n"
            "then re-run this script."
        )
    except subprocess.CalledProcessError as e:
        sys.exit(
            f"ERROR: pandoc failed (exit code {e.returncode}).\n"
            f"Stderr: {e.stderr.strip()}"
        )
    return result.stdout


# ====================================================================
# STEP 1b – CONVERT MATHTYPE IMAGES → LATEX VIA MathML ALT TEXT
# ====================================================================

def convert_mathml_alt_to_latex(soup: BeautifulSoup) -> tuple:
    """
    Find every <img> tag whose alt attribute contains MathML in the format
    MathType uses:
        {"mathml": "<math xmlns=...>...</math>"}
    or (less commonly) raw MathML:
        <math xmlns=...>...</math>

    Converts the MathML to LaTeX and replaces the <img> tag with:
      • Display math  (img is the only content of its <p>):
            <p style="text-align: center">[latex]...[/latex]</p>
      • Inline math (img is inside a paragraph with other content):
            \\( ... \\)

    Returns (converted_count, no_mathml_count) where no_mathml_count is
    the number of media images that had no MathML and need other handling.
    """
    converted    = 0
    no_mathml    = 0

    for img in list(soup.find_all('img')):
        src = img.get('src', '')
        # Only process images that came from the docx (pandoc's media folder)
        if not ('media/' in src):
            continue

        alt = img.get('alt', '').strip()
        if not alt:
            no_mathml += 1
            continue

        # ── Parse MathML from alt text ───────────────────────────────
        mathml = None

        # Format A: JSON wrapper {"mathml":"<math...>"}  (MathType default)
        if alt.startswith('{') and 'mathml' in alt:
            try:
                data   = json.loads(alt)
                mathml = data.get('mathml', '') or None
            except (json.JSONDecodeError, ValueError):
                pass

        # Format B: raw MathML string
        elif alt.lstrip().startswith('<math'):
            mathml = alt

        if not mathml:
            # This image has alt text but no MathML — might be a diagram/photo
            # with a regular text description.  Leave it alone.
            no_mathml += 1
            continue

        # ── Convert MathML → LaTeX ───────────────────────────────────
        latex = mathml_to_latex(mathml)
        if not latex:
            _replace_img_with_todo(soup, img, Path(src).name,
                                   reason='MathML could not be parsed')
            continue

        # ── Insert LaTeX in place of the <img> ──────────────────────
        parent     = img.parent
        is_display = (
            isinstance(parent, Tag)
            and parent.name == 'p'
            and all(
                (isinstance(c, NavigableString) and not c.strip()) or c is img
                for c in parent.children
            )
        )

        if is_display:
            latex_text = f'[latex]{latex}[/latex]'
            new_p = soup.new_tag('p', attrs={'style': 'text-align: center'})
            new_p.string = latex_text
            if parent.parent:
                parent.replace_with(new_p)
            else:
                img.replace_with(NavigableString(latex_text))
        else:
            img.replace_with(NavigableString(f'\\({latex}\\)'))

        converted += 1

    return converted, no_mathml


# ====================================================================
# STEP 1c – OPTIONAL MATHPIX FALLBACK FOR REMAINING EQUATION IMAGES
# ====================================================================

def convert_image_to_png(img_path: Path):
    """
    Convert a WMF or EMF image to PNG using ImageMagick or Inkscape.
    Returns the PNG path on success, or None if conversion failed.
    Install ImageMagick on macOS:  brew install imagemagick
    """
    if img_path.suffix.lower() in ('.png', '.jpg', '.jpeg'):
        return img_path
    png_path = img_path.with_suffix('.png')

    for cmd in [
        ['convert', str(img_path), str(png_path)],
        ['inkscape', f'--export-png={png_path}', str(img_path)],
        ['inkscape', '--export-type=png',
         f'--export-filename={png_path}', str(img_path)],
    ]:
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=20)
            if r.returncode == 0 and png_path.exists():
                return png_path
        except (FileNotFoundError, subprocess.TimeoutExpired):
            continue
    return None


def convert_remaining_images_mathpix(
        soup: BeautifulSoup, media_dir: Path) -> tuple:
    """
    Use the MathPix API to convert any remaining <img> tags from the
    media folder that still have equation-like content (i.e. the MathML
    step did not handle them).  Only called if MATHPIX_APP_ID is set.
    Returns (converted_count, failed_count).
    """
    if not MATHPIX_APP_ID or not MATHPIX_APP_KEY:
        return 0, 0
    try:
        import requests
        import base64
    except ImportError:
        print("  Warning: 'requests' not installed — MathPix skipped.\n"
              "  Run: pip install requests --break-system-packages")
        return 0, 0

    converted = 0
    failed    = 0
    for img in list(soup.find_all('img')):
        src = img.get('src', '')
        if 'media/' not in src:
            continue
        img_path = (media_dir.parent / src).resolve()
        if not img_path.exists():
            continue
        png_path = convert_image_to_png(img_path)
        if png_path is None:
            _replace_img_with_todo(soup, img, img_path.name,
                                   reason='no image converter (install ImageMagick)')
            failed += 1
            continue
        try:
            with open(str(png_path), 'rb') as fh:
                img_b64 = base64.b64encode(fh.read()).decode('utf-8')
            mime = ('image/jpeg' if png_path.suffix.lower() in ('.jpg', '.jpeg')
                    else 'image/png')
            resp = requests.post(
                'https://api.mathpix.com/v3/text',
                headers={'app_id': MATHPIX_APP_ID, 'app_key': MATHPIX_APP_KEY,
                         'Content-type': 'application/json'},
                json={'src': f'data:{mime};base64,{img_b64}',
                      'formats': ['latex_simplified'],
                      'math_inline_delimiters':  ['\\(', '\\)'],
                      'math_display_delimiters': ['\\[', '\\]']},
                timeout=30,
            )
            if resp.status_code != 200:
                raise ValueError(f'API error {resp.status_code}')
            latex = (resp.json() or {}).get('latex_simplified', '').strip()
            if not latex:
                raise ValueError('empty LaTeX returned')
        except Exception as exc:
            _replace_img_with_todo(soup, img, img_path.name, reason=str(exc))
            failed += 1
            continue

        parent     = img.parent
        is_display = (
            isinstance(parent, Tag)
            and parent.name == 'p'
            and all(
                (isinstance(c, NavigableString) and not c.strip()) or c is img
                for c in parent.children
            )
        )
        if is_display:
            latex_text = f'[latex]{latex}[/latex]'
            new_p = soup.new_tag('p', attrs={'style': 'text-align: center'})
            new_p.string = latex_text
            if parent.parent:
                parent.replace_with(new_p)
            else:
                img.replace_with(NavigableString(latex_text))
        else:
            img.replace_with(NavigableString(f'\\({latex}\\)'))
        converted += 1
    return converted, failed


def _replace_img_with_todo(
        soup: BeautifulSoup, img: Tag, filename: str, reason: str = '') -> None:
    """Replace an <img> tag with a visible highlighted TODO span."""
    note = f' ({reason})' if reason else ''
    span_html = (
        f'<span style="background:#ffe066; color:#333; padding:2px 6px; '
        f'font-family:monospace; font-size:0.9em;">'
        f'[TODO: insert LaTeX for {filename}{note}]</span>'
    )
    img.replace_with(BeautifulSoup(span_html, 'html.parser'))


# ====================================================================
# STEP 2 – NORMALIZE MATH SPANS
# ====================================================================

def normalize_math(soup: BeautifulSoup) -> None:
    """
    Replace pandoc's math spans with bare MathJax delimiters.
    Inline  → plain text \\(...\\)
    Display → <p class="align-center">\\[...\\]</p>
    """
    for span in soup.find_all("span", class_="math inline"):
        span.replace_with(NavigableString(span.get_text()))

    for span in soup.find_all("span", class_="math display"):
        if span.parent is None:
            continue
        math_text = span.get_text()
        parent = span.parent
        new_p = soup.new_tag("p", attrs={"class": "align-center"})
        new_p.string = math_text
        if parent and parent.name == "p" and parent.parent is not None:
            parent.replace_with(new_p)
        else:
            span.replace_with(new_p)


# ====================================================================
# STEP 2b – CONVERT GROUPED DISPLAY MATH → align* ENVIRONMENT
# ====================================================================

def _split_math_groups(s: str) -> list:
    """Split '{eq1}{eq2}{eq3}' → ['eq1', 'eq2', 'eq3'] with brace tracking."""
    groups, depth, current = [], 0, []
    for char in s:
        if char == '{':
            if depth == 0:
                current = []
            else:
                current.append(char)
            depth += 1
        elif char == '}':
            depth -= 1
            if depth == 0:
                groups.append(''.join(current))
                current = []
            else:
                current.append(char)
        else:
            if depth > 0:
                current.append(char)
    return groups


def convert_grouped_display_math(soup: BeautifulSoup) -> None:
    for p in list(soup.find_all("p", class_="align-center")):
        text = p.get_text().strip()
        if not (text.startswith(r'\[{') and text.endswith(r'}\]')):
            continue
        inner = text[2:-2]
        groups = _split_math_groups(inner)
        if len(groups) < 2:
            continue
        aligned_lines = []
        for eq in groups:
            eq = eq.strip()
            if '=' in eq:
                eq = eq.replace('=', ' &=', 1)
            aligned_lines.append(eq)
        latex_body = ' \\\\\n'.join(aligned_lines)
        latex = f'[latex]\\begin{{align*}}\n{latex_body}\n\\end{{align*}}[/latex]'
        new_p = soup.new_tag("p", attrs={"style": "text-align: center"})
        new_p.string = latex
        p.replace_with(new_p)


# ====================================================================
# STEP 2c – SHIFT HEADING LEVELS
# ====================================================================

def shift_headings(soup: BeautifulSoup, n: int) -> None:
    if n <= 0:
        return
    for level in range(6, 0, -1):
        new_level = min(level + n, 6)
        if level == new_level:
            continue
        for tag in soup.find_all(f'h{level}'):
            tag.name = f'h{new_level}'


# ====================================================================
# STEP 2d – REMOVE WORD TABLE OF CONTENTS
# ====================================================================

def remove_word_toc(soup: BeautifulSoup) -> None:
    body = soup.find("body") or soup
    toc_heading = None
    for tag in body.find_all(["h1", "h2", "h3"], recursive=False):
        if tag.get_text().strip().lower() in ("contents", "table of contents"):
            toc_heading = tag
            break
    if toc_heading is None:
        return
    to_remove = [toc_heading]
    for sibling in toc_heading.next_siblings:
        if isinstance(sibling, Tag) and sibling.name in ("h1", "h2", "h3", "h4"):
            break
        to_remove.append(sibling)
    for el in to_remove:
        el.decompose()
    print(f"  Removed Word TOC block ({len(to_remove)} element(s))")


# ====================================================================
# STEP 2e – HTML CLEANUP PASSES
# ====================================================================

def remove_scripts_styles(soup: BeautifulSoup) -> None:
    for tag in soup.find_all(["script", "style"]):
        tag.decompose()


def fix_file_anchor_links(soup: BeautifulSoup) -> None:
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith("file:///") and "#" in href:
            a["href"] = "#" + href.split("#", 1)[1]


def remove_legacy_toc_anchors(soup: BeautifulSoup) -> None:
    for a in list(soup.find_all("a")):
        if _TOC_ID_RE.match(a.get("id", "")) or _TOC_ID_RE.match(a.get("name", "")):
            a.decompose()
    for t in soup.find_all(True, id=True):
        if _TOC_ID_RE.match(t["id"]):
            del t["id"]


def remove_br_between_display_math(soup: BeautifulSoup) -> None:
    def _is_display(t):
        return (isinstance(t, Tag) and t.name == "span"
                and "math" in (t.get("class") or [])
                and "display" in (t.get("class") or []))
    for br in list(soup.find_all("br")):
        if _is_display(br.find_previous_sibling()) and _is_display(br.find_next_sibling()):
            br.decompose()


def flatten_blockquotes(soup: BeautifulSoup) -> None:
    for bq in list(soup.find_all("blockquote")):
        bq.unwrap()


def normalize_tables(soup: BeautifulSoup) -> None:
    for table in soup.find_all("table"):
        style = table.get("style", "").strip()
        if not style:
            table["style"] = "width: 100%;"
        elif "width" not in style.lower():
            table["style"] = style.rstrip(";") + "; width: 100%;"


def collapse_soft_newlines(soup: BeautifulSoup) -> None:
    for node in soup.find_all(string=True):
        if not isinstance(node, NavigableString):
            continue
        parent = node.parent
        if not isinstance(parent, Tag) or parent.name in ("script", "style", "pre", "code"):
            continue
        s = str(node)
        if "\n" not in s:
            continue
        s2 = re.sub(r'[ \t]{2,}', ' ', _SOFT_NL_RE.sub(' ', s))
        if s2 != s:
            node.replace_with(s2)


# ====================================================================
# STEP 2f – MERGE CONSECUTIVE DISPLAY MATH PARAGRAPHS → align*
# ====================================================================

def merge_consecutive_display_math_paragraphs(soup: BeautifulSoup) -> int:
    def _is_display_para(p):
        if not isinstance(p, Tag) or p.name != "p":
            return False
        if p.get("class") != ["align-center"]:
            return False
        return bool(_DISPLAY_P_RE.match(p.get_text().strip()))

    def _inner(p):
        m = _DISPLAY_P_RE.match(p.get_text().strip())
        return m.group(1).strip().rstrip(".,;:") if m else ""

    def _align_line(eq):
        eq = eq.strip()
        if '=' in eq:
            left, _, right = eq.partition('=')
            left, right = left.strip(), right.strip()
            return f"{left} &= {right}" if left else f"&= {right}"
        return eq

    def _next_tag(el):
        sib = el.find_next_sibling()
        while sib and isinstance(sib, NavigableString) and not sib.strip():
            sib = sib.find_next_sibling()
        return sib

    merged    = 0
    all_paras = list(soup.find_all("p"))
    i         = 0
    while i < len(all_paras):
        p = all_paras[i]
        if p.parent is None or not _is_display_para(p):
            i += 1
            continue
        group = [p]
        nxt   = _next_tag(p)
        while nxt and isinstance(nxt, Tag) and _is_display_para(nxt):
            group.append(nxt)
            nxt = _next_tag(nxt)
        if len(group) >= 2:
            lines      = [_align_line(_inner(gp)) for gp in group]
            align_body = ' \\\\\n'.join(lines)
            latex      = f'[latex]\\begin{{align*}}\n{align_body}\n\\end{{align*}}[/latex]'
            new_p      = soup.new_tag("p", attrs={"style": "text-align: center"})
            new_p.string = latex
            group[0].replace_with(new_p)
            for gp in group[1:]:
                gp.decompose()
            merged += 1
        i += len(group)
    return merged


# ====================================================================
# STEP 2g – SINGLE-LETTER ITALIC → LATEX
# ====================================================================

def convert_em_to_latex(soup: BeautifulSoup) -> int:
    converted = 0
    for em in list(soup.find_all("em")):
        if em.find(True) is not None:
            continue
        txt = em.get_text(strip=True)
        if _SINGLE_CHAR_RE.match(txt):
            em.replace_with(NavigableString(f"[latex]{txt}[/latex]"))
            converted += 1
    return converted


# ====================================================================
# STEP 2h – KEY OBJECTIVE BOXES
# ====================================================================

def convert_key_objective_boxes(soup: BeautifulSoup, purple_headers: set) -> int:
    if not purple_headers:
        return 0

    def _sig_children(tag):
        return [c for c in tag.children
                if not (isinstance(c, NavigableString) and not c.strip())]

    to_convert = []
    for li in soup.find_all("li"):
        children = _sig_children(li)
        if len(children) < 2:
            continue
        first = children[0]
        if not (isinstance(first, Tag) and first.name == "p"):
            continue
        p_children = _sig_children(first)
        if len(p_children) != 1:
            continue
        strong = p_children[0]
        if not (isinstance(strong, Tag) and strong.name == "strong"):
            continue
        title = strong.get_text(strip=True)
        if title not in purple_headers:
            continue
        second = children[1]
        if not (isinstance(second, Tag) and second.name == "ul"):
            continue
        to_convert.append((li, title, second))

    converted = 0
    for li, title, content_ul in to_convert:
        if li.parent is None:
            continue
        box = soup.new_tag("div")
        box["class"] = ["textbox", "textbox--key-takeaways"]
        header_tag = soup.new_tag("header")
        header_tag["class"] = ["textbox__header"]
        title_p = soup.new_tag("p")
        title_p["class"] = ["textbox__title"]
        title_p.string = title
        header_tag.append(title_p)
        content_div = soup.new_tag("div")
        content_div["class"] = ["textbox__content"]
        for item in content_ul.find_all("li", recursive=False):
            for p in list(item.find_all("p", recursive=False)):
                content_div.append(p.extract())
        box.append(header_tag)
        box.append(content_div)
        parent_ul = li.parent
        parent_ul.insert_before(box)
        li.decompose()
        converted += 1

    for ul in soup.find_all("ul"):
        if not any(isinstance(c, Tag) for c in ul.children):
            ul.decompose()
    return converted


# ====================================================================
# STEP 2i – CONVERT EXAMPLE HEADINGS/PARAGRAPHS → TEXTBOX--EXAMPLES
# ====================================================================

_EXAMPLE_LABEL_RE = re.compile(r'(?i)^\s*example\s*([0-9]+)?\s*[:.\-]?\s*(.*)$')

def _norm_ws(s: str) -> str:
    return re.sub(r'\s+', ' ', (s or '')).strip()

def _is_heading_tag(t) -> bool:
    return isinstance(t, Tag) and t.name in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6')

def _is_example_starter(t) -> bool:
    if not isinstance(t, Tag):
        return False
    if t.name not in ('p', 'h2', 'h3', 'h4', 'h5', 'h6'):
        return False
    txt = _norm_ws(t.get_text(' ', strip=True))
    return bool(txt) and txt.lower().startswith('example')

def _parse_example_label(txt: str):
    m = _EXAMPLE_LABEL_RE.match(_norm_ws(txt))
    if not m:
        return None, ''
    num  = int(m.group(1)) if m.group(1) else None
    rest = (m.group(2) or '').strip()
    return num, rest

def _make_example_box_tag(soup: BeautifulSoup, num: int) -> Tag:
    box = soup.new_tag('div')
    box['class'] = ['textbox', 'textbox--examples']
    header = soup.new_tag('header')
    header['class'] = ['textbox__header']
    title = soup.new_tag('p')
    title['class'] = ['textbox__title']
    title.string = f'Example {num}'
    header.append(title)
    content = soup.new_tag('div')
    content['class'] = ['textbox__content']
    box.append(header)
    box.append(content)
    return box


def convert_examples_to_textboxes(soup: BeautifulSoup) -> int:
    wrapped   = 0
    container = soup.body if soup.body else soup

    max_explicit = 0
    for t in container.find_all(['p', 'h2', 'h3', 'h4', 'h5', 'h6']):
        if _is_example_starter(t):
            n, _ = _parse_example_label(t.get_text(' ', strip=True))
            if n is not None:
                max_explicit = max(max_explicit, n)
    next_auto = max_explicit + 1

    starters = list(container.find_all(['p', 'h2', 'h3', 'h4', 'h5', 'h6']))
    for starter in starters:
        if not _is_example_starter(starter):
            continue
        if starter.parent is None:
            continue

        n, rest = _parse_example_label(starter.get_text(' ', strip=True))
        if n is None:
            n = next_auto
            next_auto += 1

        box         = _make_example_box_tag(soup, n)
        content_div = box.find('div', class_='textbox__content')

        if rest:
            p0 = soup.new_tag('p')
            p0.string = rest
            content_div.append(p0)

        starter.insert_before(box)
        starter.extract()

        sib = box.next_sibling
        while sib is not None:
            if isinstance(sib, NavigableString) and not sib.strip():
                nxt = sib.next_sibling
                sib.extract()
                sib = nxt
                continue
            if isinstance(sib, Tag) and (_is_heading_tag(sib) or _is_example_starter(sib)):
                break
            nxt = sib.next_sibling
            content_div.append(sib.extract())
            sib = nxt

        wrapped += 1
    return wrapped


def insert_solution_breaks_in_examples(soup: BeautifulSoup) -> int:
    modified = 0
    for box in soup.find_all('div', class_=lambda c: c and 'textbox--examples' in c):
        content = box.find('div', class_='textbox__content')
        if not content:
            continue
        paragraphs = [p for p in content.find_all('p', recursive=False)
                      if p.get_text(strip=True)]
        if len(paragraphs) < 2:
            continue
        first_p = paragraphs[0]
        nxt = first_p.find_next_sibling()
        if isinstance(nxt, Tag) and nxt.name == 'hr':
            continue
        hr     = soup.new_tag('hr')
        strong = soup.new_tag('strong')
        strong.string = 'Solution'
        first_p.insert_after(hr)
        hr.insert_after(strong)
        modified += 1
    return modified


# ====================================================================
# STEP 3 – SLUGIFY  (for heading anchors)
# ====================================================================

def slugify(text: str) -> str:
    s = text.lower()
    s = re.sub(r'[^a-z0-9\s-]', '', s)
    s = re.sub(r'\s+', '-', s.strip())
    return s


# ====================================================================
# STEP 4 – CLASSIFY PARAGRAPHS
# ====================================================================

def inner_text(tag: Tag) -> str:
    return tag.get_text(separator=" ").strip()


def is_bold_allcaps(tag: Tag) -> bool:
    strongs = tag.find_all("strong")
    if not strongs:
        return False
    bold_text = " ".join(s.get_text() for s in strongs).strip()
    non_bold  = tag.get_text().replace(bold_text, "").strip()
    if non_bold and non_bold not in ("", ".", ":"):
        return False
    return bool(_PROC_TITLE_RE.match(bold_text))


def is_solution_line(tag: Tag) -> bool:
    return inner_text(tag).lower().rstrip(".") in SOLUTION_WORDS


# ====================================================================
# STEP 5 – BUILD PRESSBOOKS HTML
# ====================================================================

def ind(level: int) -> str:
    return "  " * level

def open_box(box_class: str, title: str, level: int = 0) -> list:
    i = ind(level)
    return [
        f'{i}<div class="textbox {box_class}">',
        f'{i}  <header class="textbox__header">',
        f'{i}    <p class="textbox__title">{title}</p>',
        f'{i}  </header>',
        f'{i}  <div class="textbox__content">',
    ]

def close_box(level: int = 0) -> list:
    i = ind(level)
    return [f'{i}  </div>', f'{i}</div>']


def build_pressbooks_html(soup: BeautifulSoup) -> str:
    lines = []
    body  = soup.find("body") or soup

    h2_headings = [inner_text(tag) for tag in body.find_all("h2") if inner_text(tag)]
    if h2_headings:
        lines += [
            '<div class="textbox textbox--learning-objectives">',
            '  <header class="textbox__header">',
            '    <p class="textbox__title">Section Topics</p>',
            '  </header>',
            '  <div class="textbox__content">',
            '    <div>',
        ]
        for heading in h2_headings:
            anchor = slugify(heading)
            lines.append(
                f'      <p style="padding-left: 40px">'
                f'<a href="#{anchor}">{heading}</a></p>'
            )
        lines += ['    </div>', '  </div>', '</div>', '']

    box_open = False

    def close_current_box() -> list:
        nonlocal box_open
        out = []
        if box_open:
            out += close_box()
            box_open = False
        return out

    for node in body.children:
        if isinstance(node, NavigableString):
            text = str(node).strip()
            if text:
                lines.append(f'<p>{text}</p>')
            continue
        if not isinstance(node, Tag):
            continue

        tag_name = node.name

        if tag_name == "h1":
            lines += close_current_box()
            lines.append(f'<h1>{inner_text(node)}</h1>')
            lines.append('')
            continue

        if tag_name == "h2":
            lines += close_current_box()
            text   = inner_text(node)
            anchor = slugify(text)
            lines.append(f'<h2 id="{anchor}"><a id="{anchor}"></a>{text}</h2>')
            lines.append('')
            continue

        if tag_name == "h3":
            lines += close_current_box()
            lines.append(f'<h3>{inner_text(node)}</h3>')
            continue

        if tag_name in ("h4", "h5", "h6"):
            lines.append(f'<{tag_name}>{inner_text(node)}</{tag_name}>')
            continue

        if tag_name == "p":
            text = inner_text(node)
            if not text and not node.find("img") and not node.find("span"):
                continue
            if (node.get("class") == ["align-center"]
                    or node.get("style") == "text-align: center"):
                lines.append(str(node))
                continue
            if is_bold_allcaps(node):
                lines += close_current_box()
                lines += open_box("textbox--key-takeaways", inner_text(node))
                box_open = True
                continue
            lines.append(str(node))
            continue

        if tag_name in ("ul", "ol"):
            lines.append(str(node))
            continue

        if tag_name == "table":
            lines.append(str(node))
            lines.append('')
            continue

        if tag_name == "hr":
            lines.append('<hr />')
            continue

        if tag_name in ("div", "figure", "figcaption", "pre"):
            lines += close_current_box()
            lines.append(str(node))
            continue

        raw = str(node).strip()
        if raw:
            lines.append(raw)

    lines += close_current_box()
    return "\n".join(lines)


# ====================================================================
# STEP 6 – WRAP IN PRESSBOOKS PAGE SHELL
# ====================================================================

PRESSBOOKS_HEADER = textwrap.dedent("""\
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      <title>Pressbooks Chapter</title>
      <!--
        Copy the content between <body> and </body> into the Pressbooks
        "Text" (HTML source) editor for a chapter or back matter page.

        Math notation:
          Inline  : \\( ... \\)
          Display : \\[ ... \\]  or  [latex]...[/latex]
        Both render with MathJax when enabled in Pressbooks.

        Note: \\cancel{} requires the MathJax cancel extension.
        In Pressbooks, enable it under: Appearance → Theme Options → MathJax.

        Grouped/aligned equations use [latex]\\begin{align*}...\\end{align*}[/latex].

        Images are referenced relative to a  media/  folder.
        Upload images via Pressbooks Media Library and update the src= paths,
        or use UploadImagesToPress.py to automate this.
      -->
    </head>
    <body>
    """)

PRESSBOOKS_FOOTER = textwrap.dedent("""\
    </body>
    </html>
    """)


# ====================================================================
# MAIN
# ====================================================================

def main():
    # ── Resolve input path ───────────────────────────────────────────
    if INPUT_FILE:
        docx_path = Path(INPUT_FILE)
    elif len(sys.argv) >= 2:
        docx_path = Path(sys.argv[1])
    else:
        sys.exit(
            "No input file specified.\n\n"
            "To run from PyCharm:\n"
            "  1. Open this script.\n"
            "  2. Find the line  INPUT_FILE = \"\"\n"
            "  3. Paste your .docx path between the quotes.\n"
            "  4. Click the green Run button.\n\n"
            "To run from the terminal:\n"
            "  python 1035_WordToPressbooksV2.py \"Chapter 3.docx\""
        )

    if not docx_path.exists():
        sys.exit(f"ERROR: File not found:\n  {docx_path}")
    if docx_path.suffix.lower() != ".docx":
        sys.exit(f"ERROR: Expected a .docx file, got: {docx_path.suffix}")

    # ── Resolve output path ──────────────────────────────────────────
    if OUTPUT_FILE:
        out_path = Path(OUTPUT_FILE)
        out_path.parent.mkdir(parents=True, exist_ok=True)
    else:
        out_path = docx_path.with_name(docx_path.stem + "_V2.html")

    # ── 0. Pre-scan docx for purple Key Objective headers ────────────
    print(f"Scanning for purple Key Objective headers in: {docx_path.name} …")
    purple_headers = extract_purple_list_headers(docx_path)
    if purple_headers:
        print(f"  Found {len(purple_headers)} Key Objective header(s).")
    else:
        print("  None found (or KEY_OBJECTIVE_COLOR not matched).")

    # ── 0b. Detect MathType OLE equations ───────────────────────────
    print("Checking for MathType OLE equations …")
    mathtype_count = detect_mathtype_equations(docx_path)
    if mathtype_count > 0:
        print(f"  Found {mathtype_count} MathType equation(s).")
    else:
        print("  No MathType OLE equations detected.")

    # ── 1. Pandoc → raw HTML ─────────────────────────────────────────
    print(f"Running pandoc on: {docx_path.name} …")
    raw_html = run_pandoc(docx_path)

    # ── 2. Parse ─────────────────────────────────────────────────────
    soup = BeautifulSoup(raw_html, "html.parser")

    # ── 3a. Shift heading levels ─────────────────────────────────────
    if SHIFT_HEADINGS_DOWN:
        print(f"Shifting headings down by {SHIFT_HEADINGS_DOWN} level(s) …")
    shift_headings(soup, SHIFT_HEADINGS_DOWN)

    # ── 3b. Remove the Word TOC block ───────────────────────────────
    print("Removing Word TOC block (if present) …")
    remove_word_toc(soup)

    # ── 3c. HTML cleanup ─────────────────────────────────────────────
    remove_scripts_styles(soup)
    fix_file_anchor_links(soup)
    remove_legacy_toc_anchors(soup)
    remove_br_between_display_math(soup)
    flatten_blockquotes(soup)
    normalize_tables(soup)
    collapse_soft_newlines(soup)

    # ── 3d. Normalize math spans → bare MathJax delimiters ──────────
    normalize_math(soup)

    # ── 3e. Convert grouped display math → align* ───────────────────
    convert_grouped_display_math(soup)

    # ── 3f. Merge consecutive single-equation display paragraphs ────
    merged = merge_consecutive_display_math_paragraphs(soup)
    if merged:
        print(f"  Merged {merged} consecutive equation group(s) → align* block(s).")

    # ── 3g. Single-letter italic → [latex] ──────────────────────────
    em_converted = convert_em_to_latex(soup)
    if em_converted:
        print(f"  Converted {em_converted} single-letter <em> → [latex].")

    # ── 3h. Convert Key Objective boxes ─────────────────────────────
    ko_count = convert_key_objective_boxes(soup, purple_headers)
    if ko_count:
        print(f"  Converted {ko_count} Key Objective box(es).")

    # ── 3i. Convert Example headings/paragraphs → textbox--examples ─
    print("Converting Example blocks …")
    ex_count = convert_examples_to_textboxes(soup)
    if ex_count:
        print(f"  Wrapped {ex_count} Example box(es).")
    else:
        print("  No 'Example N' headings or paragraphs found.")

    # ── 3j. Insert Solution breaks inside Example boxes ──────────────
    sol_inserted = insert_solution_breaks_in_examples(soup)
    if sol_inserted:
        print(f"  Added Solution break to {sol_inserted} example box(es).")

    # ── 3k. Convert MathType image equations → LaTeX (free, via MathML)
    mt_converted  = 0
    mt_todo       = 0
    mp_converted  = 0
    mp_failed     = 0

    if mathtype_count > 0:
        print("Converting MathType equation images → LaTeX via MathML alt text …")
        mt_converted, no_mathml = convert_mathml_alt_to_latex(soup)
        if mt_converted:
            print(f"  ✔  Converted {mt_converted} equation(s) via MathML.")

        # ── MathPix fallback for any remaining images ────────────────
        remaining_imgs = [
            img for img in soup.find_all('img')
            if 'media/' in img.get('src', '')
        ]
        if remaining_imgs and MATHPIX_APP_ID and MATHPIX_APP_KEY:
            print(f"  Trying MathPix for {len(remaining_imgs)} remaining image(s) …")
            mp_converted, mp_failed = convert_remaining_images_mathpix(
                soup, docx_path.parent / "media"
            )
            if mp_converted:
                print(f"  ✔  MathPix converted {mp_converted} additional equation(s).")

        # ── TODO placeholders for anything still not converted ───────
        still_remaining = [
            img for img in soup.find_all('img')
            if 'media/' in img.get('src', '')
        ]
        for img in still_remaining:
            _replace_img_with_todo(
                soup, img, Path(img.get('src', 'unknown')).name,
                reason='no MathML alt text — add alt text in MathType before saving'
            )
            mt_todo += 1
        if mt_todo:
            print(
                f"  ⚠  {mt_todo} equation image(s) had no MathML alt text.\n"
                f"     [TODO] markers inserted.  To fix: in MathType, ensure\n"
                f"     accessibility/alt text is enabled before saving your .docx."
            )

    # ── 4. Build Pressbooks HTML ─────────────────────────────────────
    pb_body = build_pressbooks_html(soup)

    # ── 5. Wrap in shell ─────────────────────────────────────────────
    final_html = PRESSBOOKS_HEADER + pb_body + "\n" + PRESSBOOKS_FOOTER

    # ── 6. Write output ──────────────────────────────────────────────
    out_path.write_text(final_html, encoding="utf-8")
    print(f"\nSaved: {out_path}")

    # ── 7. Quick stats ───────────────────────────────────────────────
    h2_count    = pb_body.count('<h2 ')
    proc_count  = pb_body.count('textbox--key-takeaways')
    sol_count   = pb_body.count('<strong>Solution</strong>')
    align_count = pb_body.count('begin{align*}')
    math_disp   = pb_body.count(r'\[')
    math_inl    = pb_body.count(r'\(')
    latex_inl   = pb_body.count('[latex]')
    todo_count  = pb_body.count('[TODO:')

    print(f"\nQuick stats:")
    print(f"  H2 sections          : {h2_count}")
    print(f"  Example boxes        : {ex_count}  ({sol_inserted} with Solution break)")
    print(f"  Key-takeaways boxes  : {proc_count}  ({ko_count} Key Objective box(es))")
    print(f"  Solution breaks      : {sol_count}")
    print(f"  align* blocks        : {align_count}")
    print(f"  Display math \\[      : {math_disp}")
    print(f"  Inline math \\(       : {math_inl}")
    print(f"  [latex] shortcodes   : {latex_inl}")
    if mathtype_count > 0:
        total_converted = mt_converted + mp_converted
        print(f"  MathType equations   : {mathtype_count} found  |  "
              f"{total_converted} converted  |  {mt_todo} need manual entry")
    if todo_count > 0:
        print(f"\n  ⚠  {todo_count} [TODO] marker(s) need manual LaTeX.")

    print(
        "\nNext steps:"
        "\n  1. Open the HTML in a browser and compare with the Word doc."
        "\n  2. Search for [TODO] and fill in any missing LaTeX equations."
        "\n  3. To enable \\cancel{} in Pressbooks: Appearance → Theme Options"
        "\n     → MathJax → add  cancel  to the extension list."
        "\n  4. Fix any box boundaries that look wrong."
        "\n  5. Verify fractions, exponents, and align* blocks."
        "\n  6. Run UploadImagesToPress.py to upload images and rewrite src= paths."
        "\n  7. Paste body content into Pressbooks chapter HTML editor."
    )


if __name__ == "__main__":
    main()
