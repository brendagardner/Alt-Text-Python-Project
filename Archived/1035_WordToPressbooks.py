#!/usr/bin/env python3
"""
Convert a Microsoft Word (.docx) file to Pressbooks-ready HTML.

USAGE
-----
    python3 WordToPressbooks.py "My Chapter.docx"


MATH NOTATION
-------------
Pandoc converts Word's native OMML equations to LaTeX automatically:
    Inline  → \( e^{i\pi}+1=0 \)
    Display → \[ e^{i\pi}+1=0 \]
Grouped/aligned equations (multiple equations in one Word math object)
are converted to a LaTeX align* environment wrapped in [latex]...[/latex].
Consecutive single-equation display paragraphs are also merged into
a single align* block automatically.

====================================================================
WHAT THIS SCRIPT DOES WELL (automated, reliable)
====================================================================
✔  Runs pandoc to convert OMML equations → \(...\) / \[...\]
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
✔  Converts grouped display math \[{eq1}{eq2}\] → align* environment
✔  Merges consecutive single-equation display paragraphs → align* block
✔  Converts single-letter/digit <em>x</em> → [latex]x[/latex]
✔  Normalizes all <table> elements to width: 100%
✔  Closes open boxes when a new section/box type begins
✔  Slugifies H2 text for id= anchors (needed for Pressbooks TOC links)
✔  Outputs clean, indented HTML ready to paste into Pressbooks

====================================================================
WHAT NEEDS HUMAN REVIEW AFTER RUNNING  (~15–45 min)
====================================================================
✘  COMPLEX FRACTIONS  – Pandoc renders most fractions correctly, but
   deeply nested or unusual fraction layouts may need a manual check.
   Look for \frac{}{} and verify the numerator/denominator are right.

✘  BOX BOUNDARIES  – The script infers box openings/closings from text
   cues (ALL-CAPS headings, "Example N" paragraphs or headings, purple
   list items).  If your Word document uses a different convention you
   will need to fix box boundaries manually.

✘  IMAGES  – Pandoc extracts images to a media/ folder.  The script
   references them with relative paths.  Move the media/ folder next
   to the final HTML file and verify every image displays correctly.

✘  TABLES  – Pandoc converts tables to <table> HTML, which Pressbooks
   accepts.  Check that header rows (<th>) are correct and that any
   merged cells look right.

✘  UNNESTED LISTS  – Pandoc renders Word lists as <ul>/<ol>.  If items
   got merged into a single paragraph in Word they may appear as prose.

✘  INDENTED EQUATIONS  – Equations that should be indented need to be
   added manually using <p style="padding-left: 40px">[latex]...[/latex]</p>.

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
OUTPUT_FILE = "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/TestMathtype_PB_HTML.html"

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
KEY_OBJECTIVE_COLOR = "7030A0" #verified color in Word for Oremus files

# ====================================================================
# IMPORTS
# ====================================================================
import re
import sys
import subprocess
import textwrap
import zipfile
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
# STEP 0 – PRE-SCAN DOCX FOR PURPLE LIST PARAGRAPH HEADERS
# ====================================================================

def extract_purple_list_headers(docx_path: Path) -> set:
    """
    Open the Word XML directly (no python-docx required) and return the
    set of paragraph texts that use both the 'ListParagraph' style AND
    the KEY_OBJECTIVE_COLOR font color.

    These become the titles of individual key-takeaways textboxes.
    Returns an empty set if the file cannot be read or has no matches.
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
# STEP 2 – NORMALIZE MATH SPANS
# ====================================================================

def normalize_math(soup: BeautifulSoup) -> None:
    """
    Replace pandoc's math spans with bare MathJax delimiters.
    Inline  → plain text \(...\)
    Display → <p class="align-center">\[...\]</p>
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
    """
    Convert pandoc grouped display math \[{eq1}{eq2}...\] → align* block.
    Single-equation display paragraphs are left for the consecutive-merge step.
    """
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
    """Rename heading tags down by n levels (e.g. h1→h2 when n=1)."""
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
    """
    Strip the Word-generated TOC block (Contents / Table of Contents heading
    plus all following siblings up to the next real heading).
    """
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
    """Strip <script> and <style> tags (pandoc --mathjax adds CDN scripts)."""
    for tag in soup.find_all(["script", "style"]):
        tag.decompose()


def fix_file_anchor_links(soup: BeautifulSoup) -> None:
    """Convert file:///...#anchor links to just #anchor."""
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith("file:///") and "#" in href:
            a["href"] = "#" + href.split("#", 1)[1]


def remove_legacy_toc_anchors(soup: BeautifulSoup) -> None:
    """Remove Word's auto-generated _Toc anchor tags and IDs."""
    for a in list(soup.find_all("a")):
        if _TOC_ID_RE.match(a.get("id", "")) or _TOC_ID_RE.match(a.get("name", "")):
            a.decompose()
    for t in soup.find_all(True, id=True):
        if _TOC_ID_RE.match(t["id"]):
            del t["id"]


def remove_br_between_display_math(soup: BeautifulSoup) -> None:
    """Remove <br> tags sandwiched between two display math spans."""
    def _is_display(t):
        return (
            isinstance(t, Tag) and t.name == "span"
            and "math" in (t.get("class") or [])
            and "display" in (t.get("class") or [])
        )
    for br in list(soup.find_all("br")):
        if _is_display(br.find_previous_sibling()) and _is_display(br.find_next_sibling()):
            br.decompose()


def flatten_blockquotes(soup: BeautifulSoup) -> None:
    """Unwrap all <blockquote> tags, promoting their contents in-place."""
    for bq in list(soup.find_all("blockquote")):
        bq.unwrap()


def normalize_tables(soup: BeautifulSoup) -> None:
    """Ensure every <table> has width: 100% so it fills the Pressbooks column."""
    for table in soup.find_all("table"):
        style = table.get("style", "").strip()
        if not style:
            table["style"] = "width: 100%;"
        elif "width" not in style.lower():
            table["style"] = style.rstrip(";") + "; width: 100%;"


def collapse_soft_newlines(soup: BeautifulSoup) -> None:
    """Collapse newline characters inside text nodes into spaces."""
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
    """
    After normalize_math() runs, single display equations live in:
        <p class="align-center">\\[...\\]</p>

    This function merges runs of 2+ consecutive such paragraphs into
    a single aligned environment:
        <p style="text-align: center">
          [latex]\\begin{align*}
          lhs &= rhs \\\\
          &= rhs2
          \\end{align*}[/latex]
        </p>

    Returns number of groups merged.
    """
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

    merged = 0
    all_paras = list(soup.find_all("p"))
    i = 0
    while i < len(all_paras):
        p = all_paras[i]
        if p.parent is None or not _is_display_para(p):
            i += 1
            continue
        group = [p]
        nxt = _next_tag(p)
        while nxt and isinstance(nxt, Tag) and _is_display_para(nxt):
            group.append(nxt)
            nxt = _next_tag(nxt)
        if len(group) >= 2:
            lines = [_align_line(_inner(gp)) for gp in group]
            align_body = ' \\\\\n'.join(lines)
            latex = f'[latex]\\begin{{align*}}\n{align_body}\n\\end{{align*}}[/latex]'
            new_p = soup.new_tag("p", attrs={"style": "text-align: center"})
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
    """
    Convert <em>x</em> to [latex]x[/latex] when the content is exactly
    one letter or digit (a math variable in running text).
    Skips <em> that contain nested tags (e.g. <em>P<sub>2</sub></em>).
    Returns number of conversions made.
    """
    converted = 0
    for em in list(soup.find_all("em")):
        if em.find(True) is not None:
            continue          # has nested tags — skip
        txt = em.get_text(strip=True)
        if _SINGLE_CHAR_RE.match(txt):
            em.replace_with(NavigableString(f"[latex]{txt}[/latex]"))
            converted += 1
    return converted


# ====================================================================
# STEP 2h – KEY OBJECTIVE BOXES
# (purple ListParagraph → individual textbox--key-takeaways boxes)
# ====================================================================

def convert_key_objective_boxes(soup: BeautifulSoup, purple_headers: set) -> int:
    """
    Convert purple ListParagraph items to individual key-takeaways textboxes.

    In pandoc output, each purple header + its sub-bullet definitions appears as:
        <ul>
          <li>
            <p><strong>Vertex</strong></p>
            <ul>
              <li><p>A vertex is a dot in the graph...</p></li>
            </ul>
          </li>
          ...
        </ul>

    Each matching <li> is converted to:
        <div class="textbox textbox--key-takeaways">
          <header class="textbox__header">
            <p class="textbox__title">Vertex</p>
          </header>
          <div class="textbox__content">
            <p>A vertex is a dot in the graph...</p>
          </div>
        </div>

    The textbox is inserted before the parent <ul> in the document,
    and the processed <li> is removed.  Empty <ul> containers are
    cleaned up after all conversions.

    Returns number of boxes created.
    """
    if not purple_headers:
        return 0

    def _sig_children(tag):
        """Children that are not blank whitespace NavigableStrings."""
        return [c for c in tag.children
                if not (isinstance(c, NavigableString) and not c.strip())]

    # Collect candidates first to avoid modifying the tree while iterating
    to_convert = []
    for li in soup.find_all("li"):
        children = _sig_children(li)
        if len(children) < 2:
            continue
        # First child: <p> with only a single <strong>
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
        # Second child: nested <ul>
        second = children[1]
        if not (isinstance(second, Tag) and second.name == "ul"):
            continue
        to_convert.append((li, title, second))

    converted = 0
    for li, title, content_ul in to_convert:
        if li.parent is None:
            continue   # already detached by a previous iteration

        # Build the textbox div
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

        # Move content paragraphs from nested <li> items into the textbox
        for item in content_ul.find_all("li", recursive=False):
            for p in list(item.find_all("p", recursive=False)):
                content_div.append(p.extract())

        box.append(header_tag)
        box.append(content_div)

        # Insert the textbox before the PARENT <ul> (same body level),
        # then remove the processed <li> from that parent <ul>.
        parent_ul = li.parent
        parent_ul.insert_before(box)
        li.decompose()
        converted += 1

    # Remove any <ul> elements that are now completely empty
    for ul in soup.find_all("ul"):
        if not any(isinstance(c, Tag) for c in ul.children):
            ul.decompose()

    return converted


# ====================================================================
# STEP 2i – CONVERT EXAMPLE HEADINGS/PARAGRAPHS → TEXTBOX--EXAMPLES
# (ported & adapted from HTMLtoPressbooks.py)
# ====================================================================

_EXAMPLE_LABEL_RE = re.compile(r'(?i)^\s*example\s*([0-9]+)?\s*[:.\-]?\s*(.*)$')

def _norm_ws(s: str) -> str:
    return re.sub(r'\s+', ' ', (s or '')).strip()

def _is_heading_tag(t) -> bool:
    return isinstance(t, Tag) and t.name in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6')

def _is_example_starter(t) -> bool:
    """True if t is a <p> or heading whose text starts with 'Example'."""
    if not isinstance(t, Tag):
        return False
    if t.name not in ('p', 'h2', 'h3', 'h4', 'h5', 'h6'):
        return False
    txt = _norm_ws(t.get_text(' ', strip=True))
    return bool(txt) and txt.lower().startswith('example')

def _parse_example_label(txt: str):
    """Return (number_or_None, remainder_text)."""
    m = _EXAMPLE_LABEL_RE.match(_norm_ws(txt))
    if not m:
        return None, ''
    num = int(m.group(1)) if m.group(1) else None
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
    """
    Find every paragraph or heading that starts with "Example N" and wrap
    it plus all following siblings (until the next heading or next Example)
    into a  <div class="textbox textbox--examples">  structure.

    Works on BOTH <p> paragraphs AND headings (h3/h4/h5 etc.), so it
    catches examples whether they were plain paragraphs or heading styles
    in Word.

    Returns number of example boxes created.
    """
    wrapped = 0
    container = soup.body if soup.body else soup

    # First pass: find highest explicit number so auto-numbering starts
    # above any explicit example number already in the document.
    max_explicit = 0
    for t in container.find_all(['p', 'h2', 'h3', 'h4', 'h5', 'h6']):
        if _is_example_starter(t):
            n, _ = _parse_example_label(t.get_text(' ', strip=True))
            if n is not None:
                max_explicit = max(max_explicit, n)
    next_auto = max_explicit + 1

    # Second pass: wrap each starter in document order.
    # Re-query each time because inserting boxes changes the tree.
    starters = list(container.find_all(['p', 'h2', 'h3', 'h4', 'h5', 'h6']))
    for starter in starters:
        if not _is_example_starter(starter):
            continue
        if starter.parent is None:
            continue   # already moved into a box by a previous iteration

        n, rest = _parse_example_label(starter.get_text(' ', strip=True))
        if n is None:
            n = next_auto
            next_auto += 1

        box = _make_example_box_tag(soup, n)
        content_div = box.find('div', class_='textbox__content')

        # If the starter line had text AFTER "Example N:", keep it
        if rest:
            p0 = soup.new_tag('p')
            p0.string = rest
            content_div.append(p0)

        # Insert box where the starter was, then remove the starter
        starter.insert_before(box)
        starter.extract()

        # Pull subsequent siblings into the box until we hit a heading
        # or another Example starter
        sib = box.next_sibling
        while sib is not None:
            if isinstance(sib, NavigableString) and not sib.strip():
                nxt = sib.next_sibling
                sib.extract()
                sib = nxt
                continue
            if isinstance(sib, Tag) and (
                    _is_heading_tag(sib) or _is_example_starter(sib)):
                break
            nxt = sib.next_sibling
            content_div.append(sib.extract())
            sib = nxt

        wrapped += 1

    return wrapped


def insert_solution_breaks_in_examples(soup: BeautifulSoup) -> int:
    """
    Inside every .textbox--examples box, insert:
        <hr />
        <strong>Solution</strong>
    immediately after the first <p> in .textbox__content,
    but only when there is at least a second paragraph to follow.

    Skips boxes that already have an <hr> after the first paragraph
    (idempotent).

    Returns number of boxes modified.
    """
    modified = 0
    for box in soup.find_all('div', class_=lambda c: c and 'textbox--examples' in c):
        content = box.find('div', class_='textbox__content')
        if not content:
            continue
        paragraphs = [
            p for p in content.find_all('p', recursive=False)
            if p.get_text(strip=True)
        ]
        if len(paragraphs) < 2:
            continue
        first_p = paragraphs[0]
        # Avoid double-inserting
        nxt = first_p.find_next_sibling()
        if isinstance(nxt, Tag) and nxt.name == 'hr':
            continue
        hr = soup.new_tag('hr')
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
    """'Convert Fractions (Review)' → 'convert-fractions-review'"""
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
    """True if the paragraph is fully bold text matching the ALL-CAPS pattern."""
    strongs = tag.find_all("strong")
    if not strongs:
        return False
    bold_text = " ".join(s.get_text() for s in strongs).strip()
    non_bold = tag.get_text().replace(bold_text, "").strip()
    if non_bold and non_bold not in ("", ".", ":"):
        return False
    return bool(_PROC_TITLE_RE.match(bold_text))


def is_example_heading(tag: Tag) -> tuple:
    """Return (True, number_string) if paragraph looks like 'Example 3'."""
    m = _EXAMPLE_RE.match(inner_text(tag))
    if m:
        return True, m.group(1)
    return False, None


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
    return [
        f'{i}  </div>',
        f'{i}</div>',
    ]


def build_pressbooks_html(soup: BeautifulSoup) -> str:
    """
    Walk the body elements and emit Pressbooks-structured HTML.

    By the time this runs, the soup has already been transformed:
      • Key Objective textboxes  → <div class="textbox textbox--key-takeaways">
      • Example textboxes        → <div class="textbox textbox--examples">
    Both are emitted as-is when the state machine encounters a <div>.

    The state machine handles:
      H1 / H2          – section boundaries; close any open procedure box
      H3 / H4 / H5 / H6 – plain subheadings (H3 no longer auto-opens boxes)
      ALL-CAPS bold <p> – opens a textbox--key-takeaways procedure box
      Regular <p>       – emitted inside whatever box is open (or at body level)
      <div>             – closes any open procedure box, then emits the div as-is
                          (this handles both pre-built K.O. and Example boxes)
    """
    lines = []
    body = soup.find("body") or soup

    # --- Section Topics callout (H2 only) ---
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

    # --- State machine (procedure boxes only; examples handled at soup level) ---
    box_open = False   # True only while inside an ALL-CAPS procedure box

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

        # ── H1 ──────────────────────────────────────────────────────
        if tag_name == "h1":
            lines += close_current_box()
            lines.append(f'<h1>{inner_text(node)}</h1>')
            lines.append('')
            continue

        # ── H2 ──────────────────────────────────────────────────────
        if tag_name == "h2":
            lines += close_current_box()
            text = inner_text(node)
            anchor = slugify(text)
            lines.append(f'<h2 id="{anchor}"><a id="{anchor}"></a>{text}</h2>')
            lines.append('')
            continue

        # ── H3 – plain subheading (no auto box conversion) ──────────
        if tag_name == "h3":
            lines += close_current_box()
            lines.append(f'<h3>{inner_text(node)}</h3>')
            continue

        # ── H4 / H5 / H6 ────────────────────────────────────────────
        if tag_name in ("h4", "h5", "h6"):
            lines.append(f'<{tag_name}>{inner_text(node)}</{tag_name}>')
            continue

        # ── Paragraphs ───────────────────────────────────────────────
        if tag_name == "p":
            text = inner_text(node)
            if not text and not node.find("img"):
                continue

            # Display math (align-center or centered style)
            if (node.get("class") == ["align-center"]
                    or node.get("style") == "text-align: center"):
                lines.append(str(node))
                continue

            # ALL-CAPS bold → procedure box
            # (Key Objective boxes are handled at soup level by
            #  convert_key_objective_boxes(); this handles ALLCAPS procedures)
            if is_bold_allcaps(node):
                lines += close_current_box()
                lines += open_box("textbox--key-takeaways", inner_text(node))
                box_open = True
                continue

            # Regular paragraph
            lines.append(str(node))
            continue

        # ── Lists ────────────────────────────────────────────────────
        if tag_name in ("ul", "ol"):
            lines.append(str(node))
            continue

        # ── Tables ───────────────────────────────────────────────────
        if tag_name == "table":
            lines.append(str(node))
            lines.append('')
            continue

        # ── Horizontal rules ─────────────────────────────────────────
        if tag_name == "hr":
            lines.append('<hr />')
            continue

        # ── Divs and other block elements ────────────────────────────
        # Close any open procedure box first so pre-built textboxes
        # (Key Objectives, Examples) don't nest inside a procedure box.
        if tag_name in ("div", "figure", "figcaption", "pre"):
            lines += close_current_box()
            lines.append(str(node))
            continue

        # ── Fallback ─────────────────────────────────────────────────
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
            "  python WordToPressbooks.py \"Chapter 3.docx\""
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
        out_path = docx_path.with_name(docx_path.stem + "_V1.html")

    # ── 0. Pre-scan docx for purple Key Objective headers ────────────
    print(f"Scanning for purple Key Objective headers in: {docx_path.name} …")
    purple_headers = extract_purple_list_headers(docx_path)
    if purple_headers:
        print(f"  Found {len(purple_headers)} Key Objective header(s).")
    else:
        print("  None found (or KEY_OBJECTIVE_COLOR not matched).")

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
    remove_br_between_display_math(soup)   # must run before normalize_math
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

    # ── 4. Build Pressbooks HTML ─────────────────────────────────────
    pb_body = build_pressbooks_html(soup)

    # ── 5. Wrap in shell ─────────────────────────────────────────────
    final_html = PRESSBOOKS_HEADER + pb_body + "\n" + PRESSBOOKS_FOOTER

    # ── 6. Write output ──────────────────────────────────────────────
    out_path.write_text(final_html, encoding="utf-8")
    print(f"\nSaved: {out_path}")

    # ── 7. Quick stats ───────────────────────────────────────────────
    h2_count    = pb_body.count('<h2 ')
    # ex_count and sol_inserted are already set from soup-level transforms above
    proc_count  = pb_body.count('textbox--key-takeaways')
    sol_count   = pb_body.count('<strong>Solution</strong>')
    align_count = pb_body.count('begin{align*}')
    math_disp   = pb_body.count(r'\[')
    math_inl    = pb_body.count(r'\(')
    latex_inl   = pb_body.count('[latex]')

    print(f"\nQuick stats:")
    print(f"  H2 sections          : {h2_count}")
    print(f"  Example boxes        : {ex_count}  (soup-level, {sol_inserted} with Solution break)")
    print(f"  Key-takeaways boxes  : {proc_count}  "
          f"(includes {ko_count} Key Objective box(es))")
    print(f"  Solution breaks      : {sol_count}")
    print(f"  align* blocks        : {align_count}")
    print(f"  Display math \\[      : {math_disp}")
    print(f"  Inline math \\(       : {math_inl}")
    print(f"  [latex] shortcodes   : {latex_inl}")
    print(
        "\nNext steps:"
        "\n  1. Open the HTML file in a browser and compare with the Word doc."
        "\n  2. Fix any box boundaries that look wrong."
        "\n  3. Verify all fractions and exponents rendered correctly."
        "\n  4. Check align* blocks — confirm &= alignment looks right."
        "\n  5. Run UploadImagesToPress.py to upload images and rewrite src= paths."
        "\n  6. Paste body content into Pressbooks chapter HTML editor."
    )


if __name__ == "__main__":
    main()
