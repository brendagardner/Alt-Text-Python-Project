#!/usr/bin/env python3
"""
HTML -> Pressbooks-ready HTML (math + structure)

What it does:
1) Accepts INPUT_PATH as either a single HTML file or a folder containing HTML files
2) Removes Pandoc/MathJax scripts & styles and strips <head> (keeps body content)
3) Removes any Table of Contents block it can detect
4) Converts MathJax-rendered blobs (<mjx-container> / <mjx-assistive-mml><math>...) into Pressbooks [latex]...[/latex]
5) Builds a "Section Topics" callout from headings (h2/h3/h4)
6) Wraps "Example..." paragraphs into <div class="textbox textbox--examples"> blocks
7) Adds a separation Solutions barrier between "Example..." and "Solution..." paragraphs
8) Align = sign in consecutive [latex]...[/latex] blocks

Dependencies:
- pandoc installed and available on PATH
- pip install beautifulsoup4
"""

from __future__ import annotations

import re
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple
from bs4 import BeautifulSoup, Tag, NavigableString


# =========================
# USER SETTINGS (EDIT THESE)
# =========================

INPUT_PATH = "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/103x Files/Canvas_Textbook_Algebra_Review/Readiness Check Finance Spr24.html5/Readiness Check Finance Spr24.html"  # folder OR single .html file. Example: "/Users/brenda/.../Orem_Pandoc_Folder" or ".../page.html"
OUTPUT_SUFFIX = "_PB"  # output becomes: originalname_PB.html for Pressbook version
OVERWRITE = True       # if False, skips if output exists

# Heading levels to include in "Section Topics"
SECTION_TOPICS_LEVELS = ("h2", "h3")  #includes the two levels of headings

# Example wrapping behavior
WRAP_EXAMPLES = True

# If pandoc is not on PATH, set full path here (optional)
PANDOC_EXE = "pandoc"

# =========================
# HELPERS
# =========================

_slug_cleanup_re = re.compile(r"[^a-z0-9\-]+")
_soft_newline_re = re.compile(r"[ \t]*\n[ \t]*")
_inline_wrap = re.compile(r"^\\\((.*)\\\)\s*$", re.S)
_display_wrap = re.compile(r"^\\\[(.*)\\\]\s*$", re.S)
_heading_num_prefix = re.compile(r"^\s*\d+(\.\d+)*\s+")
_heading_num_suffix = re.compile(r"\s+\d+(\.\d+)*\s*$")
_toc_id_re = re.compile(r"^_Toc\d+$", re.IGNORECASE)
_BLOCK_TAGS = (
    "p|div|h1|h2|h3|h4|h5|h6|ul|ol|li|table|thead|tbody|tr|td|th|"
    "blockquote|section|nav|header|footer|figure|figcaption|pre"
)
_LATEX_BLOCK_RE = re.compile(r"\[latex\](.*?)\[/latex\]", re.S)
_latex_only_re = re.compile(r"^\s*\[latex\].*?\[/latex\]\s*$", re.S)

def slugify(text: str) -> str:
    """Make a stable anchor id from heading text."""
    s = text.strip().lower()
    s = s.replace("&", "and")
    s = re.sub(r"\s+", "-", s)
    s = _slug_cleanup_re.sub("", s)
    s = re.sub(r"\-+", "-", s).strip("-")
    return s or "section"

def is_heading(tag: Tag) -> bool:
    return isinstance(tag, Tag) and tag.name in ("h1", "h2", "h3", "h4", "h5", "h6")

def normalize_whitespace(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def starts_with_example(text: str) -> bool:
    t = normalize_whitespace(text)
    return t.lower().startswith("example")

def tag_text(tag: Tag) -> str:
    return normalize_whitespace(tag.get_text(" ", strip=True))

def safe_write_text(path: Path, text: str):
    path.write_text(text, encoding="utf-8")

def collapse_soft_newlines_in_text(soup: BeautifulSoup) -> int:
    """
    Collapses newline characters that occur INSIDE text nodes into spaces.
    Does NOT change structure (paragraphs remain paragraphs).
    Returns count of text nodes changed.
    """
    changed = 0

    for node in soup.find_all(string=True):
        if not isinstance(node, NavigableString):
            continue

        parent = node.parent
        if not isinstance(parent, Tag):
            continue

        # Skip script/style and preserve pre/code behavior
        if parent.name in ("script", "style", "pre", "code"):
            continue

        s = str(node)
        if "\n" not in s:
            continue

        # Turn internal newlines into spaces, then collapse multiple spaces
        s2 = _soft_newline_re.sub(" ", s)
        s2 = re.sub(r"[ \t]{2,}", " ", s2)

        if s2 != s:
            node.replace_with(s2)
            changed += 1

    return changed

def remove_existing_section_topics(soup: BeautifulSoup) -> int:
    """
    Removes any existing 'Section Topics' textbox/callout or heading+list block.
    Returns number of blocks removed.
    """
    removed = 0

    # 1) Remove Pressbooks-style textbox that already has Section Topics title
    for title_p in list(soup.find_all("p", class_="textbox__title")):
        if tag_text(title_p).lower() == "section topics":
            wrapper = title_p.find_parent("div", class_=lambda c: c and "textbox" in c)
            if wrapper:
                wrapper.decompose()
                removed += 1

    # 2) Remove a plain heading "Section Topics" followed by a list (ul/ol)
    for h in list(soup.find_all(re.compile(r"^h[1-6]$"))):
        if tag_text(h).lower() == "section topics":
            nxt = h.find_next_sibling()
            h.decompose()
            removed += 1
            if isinstance(nxt, Tag) and nxt.name in ("ul", "ol", "div", "section"):
                nxt.decompose()
                removed += 1

    return removed

def clean_heading_text_for_topics(text: str) -> str:
    t = normalize_whitespace(text)
    t = _heading_num_prefix.sub("", t)  # removes "1.2 Rates" -> "Rates"
    t = _heading_num_suffix.sub("", t)  # removes "Rates 1.2" -> "Rates"
    return t.strip()

def ensure_heading_anchor_ids(soup: BeautifulSoup, levels=("h2", "h3")) -> int:
    """
    For each heading in `levels`, ensure:
      <hX id="foo"><a id="foo"></a>Heading Text</hX>

    Returns number of anchors added.
    """
    added = 0
    seen_ids = set()

    # Track all existing ids to avoid collisions
    for t in soup.find_all(True, id=True):
        seen_ids.add(t["id"])

    for level in levels:
        for h in soup.find_all(level):
            if not isinstance(h, Tag):
                continue

            text = tag_text(h)
            if not text:
                continue

            hid = h.get("id")

            # Create id if missing
            if not hid:
                base = slugify(text)
                hid = base
                i = 2
                while hid in seen_ids:
                    hid = f"{base}-{i}"
                    i += 1
                h["id"] = hid
                seen_ids.add(hid)

            # Check if anchor already exists
            has_anchor = any(
                isinstance(child, Tag)
                and child.name == "a"
                and child.get("id") == hid
                for child in h.contents
            )

            if not has_anchor:
                a = soup.new_tag("a")
                a["id"] = hid

                # Remove leading whitespace text nodes
                while h.contents and isinstance(h.contents[0], NavigableString) and not h.contents[0].strip():
                    h.contents[0].extract()

                # Insert anchor inline, before the first real content
                if h.contents:
                    h.contents[0].insert_before(a)
                else:
                    h.append(a)

                added += 1

    return added

def remove_legacy_toc_anchors(soup: BeautifulSoup) -> int:
    """
    Removes legacy Word/Pandoc TOC anchors like:
      <a id="_Toc123456"></a>
      <a name="_Toc123456"></a>
    and strips id attributes like id="_Toc123456" from headings.

    Returns number of removals.
    """
    removed = 0

    # 1) Remove standalone <a> anchors with _Toc ids or names
    for a in list(soup.find_all("a")):
        aid = a.get("id") or ""
        aname = a.get("name") or ""
        if _toc_id_re.match(aid) or _toc_id_re.match(aname):
            a.decompose()
            removed += 1

    # 2) Strip _Toc ids from other tags (e.g., headings)
    for t in soup.find_all(True, id=True):
        if _toc_id_re.match(t["id"]):
            del t["id"]
            removed += 1

    return removed

def normalize_tables_full_width(soup: BeautifulSoup) -> int:
    """
    Ensures all <table> elements have width: 100%.
    - Adds style if missing
    - Appends width if style exists but no width is set
    Returns number of tables modified.
    """
    modified = 0

    for table in soup.find_all("table"):
        style = table.get("style", "").strip()

        # Case 1: no style attribute at all
        if not style:
            table["style"] = "width: 100%;"
            modified += 1
            continue

        # Case 2: style exists but no width
        if "width" not in style.lower():
            # ensure proper semicolon spacing
            if not style.endswith(";"):
                style += ";"
            table["style"] = f"{style} width: 100%;"
            modified += 1

    return modified


def unwrap_blockquotes(soup: BeautifulSoup) -> int:
    """
    Unwrap ALL remaining <blockquote> tags (keep contents).
    Assumes TOC blockquote has already been removed earlier.
    """
    changed = 0
    for bq in list(soup.find_all("blockquote")):
        bq.unwrap()
        changed += 1
    return changed

# =========================
# MATH CONVERSION helpers
# =========================

_single_letter_or_digit_re = re.compile(r"^[A-Za-z0-9]$")

def convert_simple_em_to_latex(soup: BeautifulSoup) -> int:
    """
    Converts <em>x</em> to [latex]x[/latex] ONLY when the text content (trimmed) is exactly ONE alphanumeric character:
      - <em> contains NO child tags (so <em>P<sub>2</sub></em> is untouched)
    """
    changed = 0

    for em in list(soup.find_all("em")):
        if not isinstance(em, Tag):
            continue

        # If <em> contains any nested tags (sub/sup/etc.), skip it
        if em.find(True) is not None:
            continue

        # Get trimmed text inside <em>
        txt = em.get_text(strip=True)

        # Only a single letter or digit
        if not _single_letter_or_digit_re.match(txt):
            continue

        em.replace_with(NavigableString(f"[latex]{txt}[/latex]"))
        changed += 1

    return changed

def remove_br_between_display_math(soup: BeautifulSoup) -> int:
    """
    Removes <br> tags that appear between consecutive <span class="math display"> blocks.
    Returns count removed.
    """
    removed = 0
    for br in list(soup.find_all("br")):
        prev = br.find_previous_sibling()
        nxt  = br.find_next_sibling()

        def is_display_math_span(t):
            return (
                isinstance(t, Tag)
                and t.name == "span"
                and "math" in (t.get("class") or [])
                and "display" in (t.get("class") or [])
            )

        if is_display_math_span(prev) and is_display_math_span(nxt):
            br.decompose()
            removed += 1

    return removed

def mathml_to_latex(mathml: str) -> str:
    """
    Convert a <math>...</math> MathML string to LaTeX using pandoc.
    We ask pandoc for LaTeX output directly (NOT markdown) to avoid
    the '$' / '\$' parsing trap.
    Returns TeX content with wrappers removed, or "" on failure.
    """
    try:
        with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as f:
            f.write(mathml.encode("utf-8"))
            tmp = f.name

        # Key change: output LaTeX, not Markdown
        tex = subprocess.check_output(
            [PANDOC_EXE, "-f", "html", "-t", "latex", tmp],
            text=True
        ).strip()

        # pandoc typically wraps inline math as \( ... \) and display as \[ ... \]
        m = _inline_wrap.match(tex)
        if m:
            return m.group(1).strip()

        m = _display_wrap.match(tex)
        if m:
            return m.group(1).strip()

        # Fallback: if pandoc returns raw TeX without wrappers
        return tex.strip()

    except Exception:
        return ""

def replace_bullet_with_cdot(soup: BeautifulSoup) -> int:
    """
    Replaces \\bullet with \\cdot inside [latex]...[/latex] blocks only.
    Returns number of replacements made.
    """
    replaced = 0

    for node in soup.find_all(string=True):
        if not node or "[latex]" not in node:
            continue

        text = str(node)

        def repl(m):
            nonlocal replaced
            content = m.group(1)
            if r"\bullet" in content:
                count = content.count(r"\bullet")
                replaced += count
                content = content.replace(r"\bullet", r"\cdot")
            return f"[latex]{content}[/latex]"

        new_text = _LATEX_BLOCK_RE.sub(repl, text)

        if new_text != text:
            node.replace_with(new_text)

    return replaced


def center_latex_paragraphs(soup: BeautifulSoup) -> int:
    """
    Adds class="align-center" to any <p> that contains ONLY a [latex]...[/latex] shortcode.
    Works anywhere (including inside textbox__content).
    Returns number of paragraphs centered.
    """
    centered = 0

    for p in soup.find_all("p"):
        if not isinstance(p, Tag):
            continue

        # Build a normalized view of what's inside the paragraph, ignoring whitespace nodes
        parts = []
        for c in p.contents:
            if isinstance(c, NavigableString):
                if c.strip():
                    parts.append(str(c))
            elif isinstance(c, Tag):
                # If there are tags inside (like <a> or <span>), it's not "latex-only"
                parts.append("<TAG>")

        if not parts:
            continue

        # Must be exactly one text part and it must be latex-only
        if len(parts) == 1 and parts[0] != "<TAG>" and _latex_only_re.match(parts[0]):
            # Add align-center without clobbering existing classes
            classes = p.get("class") or []
            if "align-center" not in classes:
                p["class"] = classes + ["align-center"]
                centered += 1

    return centered

def shrink_long_latex(soup: BeautifulSoup, max_chars=90) -> int:
    """
    Shrinks LaTeX equations that are likely to overflow the page width
    by wrapping them in {\small ...}.

    Heuristics:
    - character length > max_chars
    - OR contains long sequences without breaks (fractions, sums, etc.)

    Returns number of equations modified.
    """
    modified = 0

    for node in soup.find_all(string=True):
        if not node or "[latex]" not in node:
            continue

        text = str(node)

        def repl(m):
            nonlocal modified
            content = m.group(1).strip()

            # Skip if already resized
            if content.startswith(r"{\small") or content.startswith(r"{\footnotesize"):
                return m.group(0)

            # Heuristic 1: character length
            if len(content) > max_chars:
                modified += 1
                return f"[latex]{{\\small {content}}}[/latex]"

            # Heuristic 2: long unbroken structures
            if any(tok in content for tok in (r"\frac", r"\sum", r"\prod", r"\left", r"\right")) and len(content) > max_chars * 0.7:
                modified += 1
                return f"[latex]{{\\small {content}}}[/latex]"

            return m.group(0)

        new_text = _LATEX_BLOCK_RE.sub(repl, text)

        if new_text != text:
            node.replace_with(new_text)

    return modified

# =========================
# Align equal sign in consecutive equations
# =========================
#
# There are TWO alignment functions in this section, handling two different
# stages of the pipeline:
#
#   merge_consecutive_equations()
#       Works on Pandoc-generated HTML that still has
#       <span class="math display"> tags (i.e. input that was NOT yet
#       processed through MathJax in a browser).
#
#   merge_consecutive_latex_equations()   ← the post-conversion version
#       Works on HTML that has ALREADY been processed by
#       convert_mathjax_to_pressbooks_latex(), where every equation is now
#       a [latex]...[/latex] shortcode inside a <p> tag.
#       This is the function that fires for the typical LibreOffice/HTML5
#       workflow, because the MathJax spans are gone by the time alignment
#       would run.
#
# Both are called in process_html_file(); the first handles Pandoc input,
# the second handles MathJax-rendered input.
# =========================

def merge_consecutive_equations(soup: BeautifulSoup) -> int:
    """
    Merges consecutive paragraphs containing ONLY math display spans into a single
    align* environment, aligning on the = sign.

    Only merges paragraphs that:
    1. Contain a <span class="math display">
    2. Have no other significant text content
    3. Are directly consecutive (no other elements in between)

    Example input:
        <p><span class="math inline">\(P(...) = 1 - ...\)</span>.</p>
        <p><span class="math display">\[= 1 - 0.98\]</span></p>
        <p><span class="math display">\[= 0.1492\]</span></p>

    Output (only the last two are merged):
        <p><span class="math inline">\(P(...) = 1 - ...\)</span>.</p>
        <p style="text-align: center;">[latex]\begin{align*}
        &= 1 - 0.98 \\
        &= 0.1492
        \end{align*}[/latex]</p>

    Returns number of equation groups merged.
    """
    merged = 0

    # Work on a copy of all paragraphs to avoid modification during iteration
    all_paragraphs = list(soup.find_all("p"))

    i = 0
    while i < len(all_paragraphs):
        p = all_paragraphs[i]

        # Check if this paragraph has been removed already
        if p.parent is None:
            i += 1
            continue

        # Check if this paragraph contains ONLY a math display span
        if not _is_display_math_only_paragraph(p):
            i += 1
            continue

        # Start collecting consecutive display math paragraphs
        math_group = [p]
        next_p = p.find_next_sibling()

        # Skip whitespace text nodes
        while next_p and isinstance(next_p, NavigableString) and not next_p.strip():
            next_p = next_p.find_next_sibling()

        while next_p and isinstance(next_p, Tag) and next_p.name == "p" and _is_display_math_only_paragraph(next_p):
            math_group.append(next_p)
            next_p = next_p.find_next_sibling()

            # Skip whitespace text nodes
            while next_p and isinstance(next_p, NavigableString) and not next_p.strip():
                next_p = next_p.find_next_sibling()

        # Only merge if we have 2+ consecutive display math paragraphs
        if len(math_group) >= 2:
            _merge_math_group_to_align(soup, math_group)
            merged += 1
            # Skip past the paragraphs we just merged
            i += len(math_group)
        else:
            i += 1

    return merged


def _is_display_math_only_paragraph(p: Tag) -> bool:
    """
    Check if paragraph contains ONLY a math display span (and whitespace).
    Returns False if:
    - Not a <p> tag
    - Contains math inline (not display)
    - Contains other text content besides the math span
    """
    if not isinstance(p, Tag) or p.name != "p":
        return False

    # Must have a math display span
    math_span = p.find("span", class_=lambda c: c and "math" in c and "display" in c)
    if not math_span:
        return False

    # Check that the paragraph contains ONLY the math span (plus whitespace/punctuation)
    # Get all text content
    all_text = p.get_text(strip=True)
    math_text = math_span.get_text(strip=True)

    # Remove the math text from all text
    remaining = all_text.replace(math_text, "").strip()

    # Allow only whitespace or minimal punctuation (period, comma)
    remaining = remaining.strip(".,;: \t\n")

    # If there's any other text, this is not a display-math-only paragraph
    return len(remaining) == 0


def _extract_latex_from_math_span(p: Tag) -> str:
    """
    Extract LaTeX content from a paragraph containing math display.
    Handles \[...\] display formats.
    Returns the inner LaTeX without wrappers, or empty string if not found.
    """
    math_span = p.find("span", class_=lambda c: c and "math" in c and "display" in c)
    if not math_span:
        return ""

    # Get the text content
    text = math_span.get_text(strip=True)

    # Remove \[ ... \] wrappers (display)
    if text.startswith(r"\[") and text.endswith(r"\]"):
        text = text[2:-2].strip()

    # Remove any trailing punctuation (like periods or commas)
    text = text.rstrip(".,;:")

    return text


def _merge_math_group_to_align(soup: BeautifulSoup, math_group: List[Tag]) -> None:
    """
    Merges a group of consecutive math display paragraphs into a single align* block.
    Replaces the first paragraph with the merged result and removes the rest.
    """
    lines = []

    for p in math_group:
        latex = _extract_latex_from_math_span(p)
        if not latex:
            continue

        # Check if this line contains an equals sign
        if "=" in latex:
            # Split on first = to align properly
            parts = latex.split("=", 1)
            if len(parts) == 2:
                left = parts[0].strip()
                right = parts[1].strip()

                # If left side is empty, this is a continuation line
                if not left:
                    lines.append(f"&= {right}")
                else:
                    lines.append(f"{left} &= {right}")
            else:
                # Shouldn't happen, but fallback
                lines.append(latex)
        else:
            # No equals sign, just add the line as-is
            lines.append(latex)

    if not lines:
        return

    # Build the align* environment
    align_content = r"\begin{align*}" + "\n"
    align_content += " \\\\\n".join(lines)
    align_content += "\n" + r"\end{align*}"

    # Create new paragraph with centered align* block
    new_p = soup.new_tag("p")
    new_p["style"] = "text-align: center;"
    new_p.string = f"[latex]{align_content}[/latex]"

    # Replace first paragraph with merged result
    math_group[0].replace_with(new_p)

    # Remove the rest
    for p in math_group[1:]:
        p.decompose()

# =========================
# POST-CONVERSION ALIGNMENT
# Operates on [latex]...[/latex] blocks after convert_mathjax_to_pressbooks_latex()
# has already replaced all <span class="math display"> tags.
# =========================

_LATEX_ONLY_PARA_RE = re.compile(r'^\s*\[latex\](.*?)\[/latex\]\s*$', re.S)


def _is_latex_only_paragraph(p: Tag) -> bool:
    """
    Returns True if <p> contains ONLY a [latex]...[/latex] shortcode
    (possibly with surrounding whitespace) and no child HTML tags.
    """
    if not isinstance(p, Tag) or p.name != "p":
        return False
    # No child tags allowed – only plain text
    if p.find(True) is not None:
        return False
    text = p.get_text()
    return bool(_LATEX_ONLY_PARA_RE.match(text))


def _extract_latex_content(p: Tag) -> str:
    """Pull the inner LaTeX string from a [latex]...[/latex] paragraph."""
    m = _LATEX_ONLY_PARA_RE.match(p.get_text())
    return m.group(1).strip() if m else ""


def merge_consecutive_latex_equations(soup: BeautifulSoup) -> int:
    """
    Post-conversion version of merge_consecutive_equations.

    After convert_mathjax_to_pressbooks_latex() has run, each display
    equation lives in a <p> containing ONLY a [latex]...[/latex] shortcode.
    This function finds runs of 2+ such consecutive <p> tags and merges
    them into a single align* environment, aligning on the = sign.

    Example input (after MathJax conversion):
        <p class="align-center">[latex]9{,}000 = 1000(1.0025)^{12t}[/latex]</p>
        <p class="align-center">[latex]9 = (1.0025)^{12t}[/latex]</p>
        <p class="align-center">[latex]\\log(9) = 12t\\cdot\\log(1.0025)[/latex]</p>

    Example output:
        <p class="align-center">[latex]\\begin{align*}
        9{,}000 &= 1000(1.0025)^{12t} \\\\
        9 &= (1.0025)^{12t} \\\\
        \\log(9) &= 12t\\cdot\\log(1.0025)
        \\end{align*}[/latex]</p>

    Returns number of equation groups merged.
    """
    merged = 0
    all_paragraphs = list(soup.find_all("p"))

    i = 0
    while i < len(all_paragraphs):
        p = all_paragraphs[i]

        if p.parent is None:
            i += 1
            continue

        if not _is_latex_only_paragraph(p):
            i += 1
            continue

        # Collect a run of consecutive [latex]-only paragraphs
        group = [p]
        next_el = p.find_next_sibling()

        # Skip bare whitespace text nodes between siblings
        while (next_el is not None
               and isinstance(next_el, NavigableString)
               and not next_el.strip()):
            next_el = next_el.find_next_sibling()

        while (next_el is not None
               and isinstance(next_el, Tag)
               and next_el.name == "p"
               and _is_latex_only_paragraph(next_el)):
            group.append(next_el)
            next_el = next_el.find_next_sibling()
            while (next_el is not None
                   and isinstance(next_el, NavigableString)
                   and not next_el.strip()):
                next_el = next_el.find_next_sibling()

        if len(group) >= 2:
            # Build the align* lines
            lines = []
            for gp in group:
                content = _extract_latex_content(gp)
                if not content:
                    continue
                if "=" in content:
                    left, _, right = content.partition("=")
                    left = left.strip()
                    right = right.strip()
                    if left:
                        lines.append(f"{left} &= {right}")
                    else:
                        lines.append(f"&= {right}")
                else:
                    lines.append(content)

            if len(lines) >= 2:
                align_body = " \\\\\n".join(lines)
                align_latex = (
                    r"\begin{align*}" + "\n"
                    + align_body + "\n"
                    + r"\end{align*}"
                )

                new_p = soup.new_tag("p")
                new_p["class"] = ["align-center"]
                new_p.string = f"[latex]{align_latex}[/latex]"

                group[0].replace_with(new_p)
                for gp in group[1:]:
                    gp.decompose()
                merged += 1

            i += len(group)
        else:
            i += 1

    return merged


# =========================
# TRANSFORM STEPS
# =========================

def strip_head_keep_body(html: str) -> BeautifulSoup:
    soup = BeautifulSoup(html, "html.parser")
    body = soup.body
    if body:
        # Create a new soup that contains only body content
        new = BeautifulSoup(str(body), "html.parser")
        return new
    return soup

def remove_scripts_styles(soup: BeautifulSoup) -> None:
    for t in soup.find_all(["script", "style"]):
        t.decompose()

def remove_leftover_mjx(soup: BeautifulSoup) -> None:
    # Remove any remaining mjx-* tags after conversion
    for t in soup.find_all(re.compile(r"^mjx-")):
        t.decompose()
    # Also remove <mjx-container> if any survived
    for t in soup.find_all("mjx-container"):
        t.decompose()

def fix_file_anchor_links(soup: BeautifulSoup) -> None:
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.startswith("file:///") and "#" in href:
            a["href"] = "#" + href.split("#", 1)[1]

def convert_mathjax_to_pressbooks_latex(soup: BeautifulSoup) -> Tuple[int, int]:
    """
    Replaces MathJax HTML output that contains <mjx-assistive-mml><math>...</math>
    with Pressbooks [latex]...[/latex]. Returns (converted_count, failed_count).
    """
    converted = 0
    failed = 0

    # Typical structure: <span class="math inline"> ... <mjx-assistive-mml><math>...</math></mjx-assistive-mml> ...</span>
    math_spans = soup.find_all("span", class_=lambda c: c and "math" in c)
    for span in math_spans:
        assist = span.find("mjx-assistive-mml")
        if not assist:
            continue
        math = assist.find("math")
        if not math:
            continue

        tex = mathml_to_latex(str(math))
        if not tex:
            failed += 1
            continue

        # Determine display vs inline:
        # pandoc often produces inline for small expressions; we rely on class hint.
        is_display = (
                             span.get("class") and "display" in span.get("class")
                     ) or (
                             assist.get("display") == "block"
                     )

        replace_math_with_latex(soup, span, tex, is_display)
        converted += 1

    return converted, failed

def replace_math_with_latex(
    soup: BeautifulSoup,
    span: Tag,
    tex: str,
    is_display: bool
):
    """
    Replaces a MathJax span with Pressbooks LaTeX.
    Display math is centered on its own line.
    """
    latex_code = f"[latex]{tex}[/latex]"

    parent = span.parent

    # CASE 1: Display math → always centered, own line
    if is_display:
        p = soup.new_tag("p")
        p["class"] = ["align-center"]
        p.string = latex_code
        span.replace_with(p)
        return

    # CASE 2: Inline math → check if it's the only content in <p>
    if parent and parent.name == "p":
        # Remove whitespace to test if span is alone
        siblings = [
            s for s in parent.contents
            if not (isinstance(s, NavigableString) and not s.strip())
        ]
        if len(siblings) == 1:
            # Treat as standalone → center it
            p = soup.new_tag("p")
            p["class"] = ["align-center"]
            p.string = latex_code
            parent.replace_with(p)
            return

    # CASE 3: True inline math → keep inline
    span.replace_with(NavigableString(latex_code))


def remove_table_of_contents(soup: BeautifulSoup) -> int:
    """
    Removes a TOC block. Best-effort:
    - headings containing 'Table of Contents'
    - <nav> with class/id containing toc
    - div/section with class/id containing toc
    Returns number of removed blocks.
    """
    removed = 0

    # 1) Remove explicit headings labeled "Table of Contents" + following list block if adjacent
    for h in list(soup.find_all(re.compile(r"^h[1-6]$"))):
        if "table of contents" in tag_text(h).lower():
            # remove heading
            nxt = h.find_next_sibling()
            h.decompose()
            removed += 1
            # remove the next sibling if it's a TOC-like list/nav/div
            if isinstance(nxt, Tag) and nxt.name in ("ul", "ol", "nav", "div", "section"):
                nxt.decompose()
                removed += 1

    # 2) Remove any nav/div/section with toc-ish id/class
    def has_toc_id_or_class(t: Tag) -> bool:
        if not isinstance(t, Tag):
            return False
        tid = (t.get("id") or "").lower()
        cls = " ".join(t.get("class") or []).lower()
        return ("toc" in tid) or ("table-of-contents" in tid) or ("toc" in cls) or ("table-of-contents" in cls)

    for t in list(soup.find_all(["nav", "div", "section"], has_toc_id_or_class)):
        t.decompose()
        removed += 1

    # 3) Remove a "TOC" list near the top if it’s mostly file:/// links
    # (common in pandoc exports)
    # 3) Remove a "TOC" list near the top if it’s mostly file:/// links
    for lst in list(soup.find_all(["ul", "ol"])):
        links = lst.find_all("a", href=True)
        if links and sum(1 for a in links if a["href"].startswith("file:///")) >= max(2, len(links) // 2):
            lst.decompose()
            removed += 1

    # ✅ NEW: remove blockquote-style TOC
    removed += remove_blockquote_toc(soup)

    return removed

def remove_blockquote_toc(soup: BeautifulSoup) -> int:
    """
    Removes a <blockquote> that is acting like a Table of Contents, e.g.
    <blockquote>
      <p><a href="#...">Table of Contents <span>1</span></a></p>
      <p><a href="#...">Some heading <span>3</span></a></p>
      ...
    </blockquote>
    Returns number of blockquotes removed.
    """
    removed = 0

    for bq in list(soup.find_all("blockquote")):
        links = bq.find_all("a", href=True)
        if len(links) < 2:
            continue

        # Must be mostly internal anchor links
        internal = [a for a in links if a["href"].startswith("#")]
        if len(internal) < 2:
            continue

        # First link text contains "Table of Contents"
        first_txt = tag_text(links[0]).lower()
        if "table of contents" not in first_txt:
            continue

        # Strong signal: lots of spans (page numbers) inside the blockquote
        spans = bq.find_all("span")
        if len(spans) >= 2:
            bq.decompose()
            removed += 1
            continue

        # Even without spans, this is still almost certainly TOC
        bq.decompose()
        removed += 1

    return removed

def serialize_html_readable(soup: BeautifulSoup) -> str:
    """
    Make HTML readable by putting newlines around block-level tags,
    WITHOUT forcing newlines between inline tags like <a> inside headings.
    """
    html = soup.decode(formatter="minimal")

    # Newline BEFORE opening block tags
    html = re.sub(rf"\s*<({_BLOCK_TAGS})\b", r"\n<\1", html, flags=re.IGNORECASE)

    # Newline AFTER closing block tags
    html = re.sub(rf"</({_BLOCK_TAGS})>\s*", r"</\1>\n", html, flags=re.IGNORECASE)

    # Clean up excessive blank lines
    html = re.sub(r"\n{3,}", "\n\n", html).strip() + "\n"

    return html

def ensure_heading_ids(soup: BeautifulSoup) -> int:
    """
    Ensures headings in SECTION_TOPICS_LEVELS have ids.
    If missing, generates stable id from text; de-dupes if needed.
    Returns count of ids created.
    """
    created = 0
    seen = set()

    # collect existing ids
    for t in soup.find_all(True, id=True):
        seen.add(t["id"])

    for level in SECTION_TOPICS_LEVELS:
        for h in soup.find_all(level):
            if not isinstance(h, Tag):
                continue
            if h.get("id"):
                continue
            base = slugify(tag_text(h))
            new_id = base
            i = 2
            while new_id in seen:
                new_id = f"{base}-{i}"
                i += 1
            h["id"] = new_id
            seen.add(new_id)
            created += 1

    return created

def build_section_topics_callout(soup: BeautifulSoup) -> Optional[Tag]:
    """
    Builds Section Topics as indented <p> links (not bullets), like:
      <div class="textbox textbox--learning-objectives">
        <header ...>Section Topics</header>
        <div class="textbox__content">
          <div>
            <p style="padding-left: 40px"><a href="#h2-id">H2 Title</a></p>
            <p style="padding-left: 80px"><a href="#h3-id">H3 Title</a></p>
            ...
          </div>
        </div>
      </div>
    Excludes headings that start with "Example".
    """
    items: list[tuple[str, str, int]] = []  # (id, text, indent_px)

    def add_heading(h: Tag, indent_px: int):
        hid = h.get("id")
        text = clean_heading_text_for_topics(tag_text(h))
        if not hid or not text:
            return
        if text.lower().startswith("example"):
            return
        items.append((hid, text, indent_px))

    for h2 in soup.find_all("h2"):
        add_heading(h2, 40)

        # Include following h3s until next h2
        sib = h2.find_next_sibling()
        while sib and not (isinstance(sib, Tag) and sib.name == "h2"):
            if isinstance(sib, Tag) and sib.name == "h3":
                add_heading(sib, 80)
            sib = sib.find_next_sibling()

    if not items:
        return None

    outer = soup.new_tag("div")
    outer["class"] = ["textbox", "textbox--learning-objectives"]

    header = soup.new_tag("header")
    header["class"] = ["textbox__header"]
    title_p = soup.new_tag("p")
    title_p["class"] = ["textbox__title"]
    title_p.string = "Section Topics"
    header.append(title_p)

    content = soup.new_tag("div")
    content["class"] = ["textbox__content"]

    inner = soup.new_tag("div")  # matches your sample (a plain div container)

    for hid, text, indent in items:
        p = soup.new_tag("p")
        p["style"] = f"padding-left: {indent}px"
        a = soup.new_tag("a", href=f"#{hid}")
        a.string = text
        p.append(a)
        inner.append(p)

    content.append(inner)
    outer.append(header)
    outer.append(content)
    return outer

def insert_section_topics_near_top(soup: BeautifulSoup) -> bool:
    """
    Inserts Section Topics callout after the first H1 if present,
    otherwise at the beginning of body content.
    Returns True if inserted.
    """
    callout = build_section_topics_callout(soup)
    if not callout:
        return False

    # find first meaningful container: inside <body> if present
    body = soup.body if hasattr(soup, "body") else None
    container = body if body else soup

    h1 = container.find("h1")
    if isinstance(h1, Tag):
        h1.insert_after(callout)
        return True

    # otherwise insert at top
    if container.contents:
        container.insert(0, callout)
    else:
        container.append(callout)
    return True

_example_num_re = re.compile(r"(?i)^\s*example\s*([0-9]+)?\s*[:.\-]?\s*(.*)$")

def parse_example_lead(text: str) -> tuple[Optional[int], str]:
    """
    Returns (number_if_present, remaining_text_after_example_label).
    Examples:
      'Example 3: Solve...' -> (3, 'Solve...')
      'Example - Find...'   -> (None, 'Find...')
      'Example'             -> (None, '')
    """
    m = _example_num_re.match(normalize_whitespace(text))
    if not m:
        return None, ""
    num = int(m.group(1)) if m.group(1) else None
    rest = (m.group(2) or "").strip()
    return num, rest

def is_example_starter(tag: Tag) -> bool:
    """True if tag is a <p> or heading that starts with 'Example'."""
    if not isinstance(tag, Tag):
        return False
    if tag.name not in ("p", "h2", "h3", "h4", "h5", "h6"):
        return False
    txt = tag_text(tag)
    return bool(txt) and txt.lower().startswith("example")

def make_example_textbox(soup: BeautifulSoup, ex_num: int) -> Tag:
    box = soup.new_tag("div")
    box["class"] = ["textbox", "textbox--examples"]

    header = soup.new_tag("header")
    header["class"] = ["textbox__header"]

    title_p = soup.new_tag("p")
    title_p["class"] = ["textbox__title"]
    title_p.string = f"Example {ex_num}"

    header.append(title_p)

    content = soup.new_tag("div")
    content["class"] = ["textbox__content"]

    box.append(header)
    box.append(content)
    return box



_example_re = re.compile(r"(?i)^\s*example\s*([0-9]+)?\s*[:.\-]?\s*(.*)$")

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()

def _is_heading(t: Tag) -> bool:
    return isinstance(t, Tag) and t.name in ("h1","h2","h3","h4","h5","h6")

def _is_example_starter(t: Tag) -> bool:
    if not isinstance(t, Tag):
        return False
    if t.name not in ("p","h2","h3","h4","h5","h6"):
        return False
    txt = _norm(t.get_text(" ", strip=True))
    return txt.lower().startswith("example")

def _parse_example_label(txt: str):
    """
    Returns (num_or_None, remainder_text_after_label)
    """
    m = _example_re.match(_norm(txt))
    if not m:
        return None, ""
    num = int(m.group(1)) if m.group(1) else None
    rest = (m.group(2) or "").strip()
    return num, rest

def _make_example_box(soup: BeautifulSoup, num: int) -> Tag:
    box = soup.new_tag("div")
    box["class"] = ["textbox", "textbox--examples"]

    header = soup.new_tag("header")
    header["class"] = ["textbox__header"]

    title = soup.new_tag("p")
    title["class"] = ["textbox__title"]
    title.string = f"Example {num}"

    header.append(title)

    content = soup.new_tag("div")
    content["class"] = ["textbox__content"]

    box.append(header)
    box.append(content)
    return box

def convert_examples_to_textboxes(soup: BeautifulSoup) -> int:
    """
    Converts any Example starter that is a heading like:
      <h4 id="example-4">Example 4</h4>
    (or paragraph starting with Example)
    into the Pressbooks textbox structure.

    Captures subsequent siblings into textbox__content until:
      - next heading (h1-h6), OR
      - next Example starter
    """
    wrapped = 0
    container = soup.body if soup.body else soup

    # Establish numbering: keep explicit numbers; auto-number the rest in reading order
    # First pass: find max explicit
    max_explicit = 0
    for t in container.find_all(["p","h2","h3","h4","h5","h6"]):
        if _is_example_starter(t):
            n, _ = _parse_example_label(t.get_text(" ", strip=True))
            if n is not None:
                max_explicit = max(max_explicit, n)
    next_auto = max_explicit + 1

    # Walk all starters in document order (safer than sibling-only scans)
    starters = list(container.find_all(["p","h2","h3","h4","h5","h6"]))
    for starter in starters:
        if not _is_example_starter(starter):
            continue
        # It might have been moved already by a previous wrap
        if starter.parent is None:
            continue

        n, rest = _parse_example_label(starter.get_text(" ", strip=True))
        if n is None:
            n = next_auto
            next_auto += 1

        box = _make_example_box(soup, n)
        content_div = box.find("div", class_="textbox__content")

        # If starter had extra text beyond "Example N", keep it as first paragraph
        if rest:
            p0 = soup.new_tag("p")
            p0.string = rest
            content_div.append(p0)

        # Insert box where starter was, then remove starter
        starter.insert_before(box)
        starter.extract()

        # Move following siblings into content until stop condition
        sib = box.next_sibling
        while sib is not None:
            # skip pure whitespace nodes
            if isinstance(sib, NavigableString) and not sib.strip():
                nxt = sib.next_sibling
                sib.extract()
                sib = nxt
                continue

            if isinstance(sib, Tag) and (_is_heading(sib) or _is_example_starter(sib)):
                break

            nxt = sib.next_sibling
            content_div.append(sib.extract())
            sib = nxt

        wrapped += 1

    return wrapped

def insert_solution_breaks_in_examples(soup: BeautifulSoup) -> int:
    """
    In each .textbox--examples block, insert:
        <hr />
        <strong>Solution</strong>
    immediately AFTER the first <p> inside .textbox__content,
    but only if there is a second paragraph to follow.

    Returns number of example boxes modified.
    """
    modified = 0

    for box in soup.find_all("div", class_=lambda c: c and "textbox--examples" in c):
        content = box.find("div", class_="textbox__content")
        if not content:
            continue

        # Only look at direct paragraph children (ignore whitespace)
        paragraphs = [
            p for p in content.find_all("p", recursive=False)
            if p.get_text(strip=True)
        ]

        # Must have at least two paragraphs
        if len(paragraphs) < 2:
            continue

        first_p = paragraphs[0]

        # Avoid double-inserting if already present
        next_tag = first_p.find_next_sibling()
        if isinstance(next_tag, Tag) and next_tag.name == "hr":
            continue

        hr = soup.new_tag("hr")
        strong = soup.new_tag("strong")
        strong.string = "Solution"

        first_p.insert_after(hr)
        hr.insert_after(strong)

        modified += 1

    return modified

def remove_repeating_topics_after_section_topics(soup: BeautifulSoup) -> int:
    """
    After we insert our Section Topics textbox, remove any repeating topics/TOC block
    that appears immediately after it (before real prose begins).
    Returns number of blocks removed.
    """
    removed = 0
    container = soup.body if soup.body else soup

    # Find our Section Topics textbox (the one we inserted)
    section_box = None
    for title in container.find_all("p", class_="textbox__title"):
        if tag_text(title).lower() == "section topics":
            section_box = title.find_parent("div", class_=lambda c: c and "textbox" in c)
            if section_box:
                break

    if not section_box:
        return 0

    def looks_like_topics_block(tag: Tag) -> bool:
        if not isinstance(tag, Tag):
            return False
        if tag.name in ("ul", "ol"):
            links = tag.find_all("a", href=True)
            # "topics list" = multiple internal anchor links
            return len(links) >= 3 and sum(1 for a in links if a["href"].startswith("#")) >= 3
        if tag.name in ("div", "section", "nav"):
            # If it contains a UL/OL that looks like topics, treat as topics block
            lst = tag.find(["ul", "ol"])
            return bool(lst and looks_like_topics_block(lst))
        return False

    def is_real_content(tag: Tag) -> bool:
        if not isinstance(tag, Tag):
            return False
        # First prose paragraph is "real content"
        if tag.name == "p" and tag_text(tag):
            return True
        # Images or tables often indicate real content begins
        if tag.name in ("img", "table", "figure"):
            return True
        # If we already hit an example textbox, that’s content
        cls = " ".join(tag.get("class") or [])
        if "textbox--examples" in cls:
            return True
        # A major section heading can be considered content start too
        if tag.name in ("h2", "h3") and tag_text(tag):
            return True
        return False

    # Walk siblings right after the Section Topics box, removing “topics-like” blocks
    sib = section_box.find_next_sibling()
    while sib is not None:
        # skip blank whitespace nodes
        if isinstance(sib, NavigableString) and not sib.strip():
            nxt = sib.next_sibling
            sib.extract()
            sib = nxt
            continue

        if isinstance(sib, Tag) and is_real_content(sib):
            break

        if isinstance(sib, Tag) and looks_like_topics_block(sib):
            nxt = sib.next_sibling
            sib.decompose()
            removed += 1
            sib = nxt
            continue

        # If it’s just an empty wrapper, remove it
        if isinstance(sib, Tag) and not tag_text(sib) and not sib.find(["p","img","table","ul","ol","h2","h3","div","section","nav"]):
            nxt = sib.next_sibling
            sib.decompose()
            removed += 1
            sib = nxt
            continue

        # Otherwise, stop — we don’t want to delete too aggressively
        break

    return removed


# =========================
# FILE IO / PIPELINE
# =========================

@dataclass
class Report:
    file_in: Path
    file_out: Path
    toc_removed: int = 0
    heading_ids_created: int = 0
    section_topics_inserted: bool = False
    math_converted: int = 0
    math_failed: int = 0
    examples_wrapped: int = 0
    soft_newlines_collapsed: int = 0
    equations_merged: int = 0

def process_html_file(path: Path) -> Report:
    out_path = path.with_name(f"{path.stem}{OUTPUT_SUFFIX}{path.suffix}")
    if out_path.exists() and not OVERWRITE:
        return Report(file_in=path, file_out=out_path)

    raw = path.read_text(encoding="utf-8", errors="ignore")
    soup = strip_head_keep_body(raw)

    remove_scripts_styles(soup)
    fix_file_anchor_links(soup)

    toc_removed = remove_table_of_contents(soup)
    section_topics_removed = remove_existing_section_topics(soup)

    # Ensure ids for headings we want in Section Topics
    heading_ids_created = ensure_heading_ids(soup)  # assigns ids to h2/h3
    anchors_added = ensure_heading_anchor_ids(soup, levels=("h2", "h3"))

    # Insert Section Topics callout near top
    section_topics_inserted = insert_section_topics_near_top(soup)
    repeating_topics_removed = remove_repeating_topics_after_section_topics(soup)

    # Remove legacy Word/Pandoc anchors AFTER ours exist
    legacy_toc_removed = remove_legacy_toc_anchors(soup)

    # Convert math
    math_converted, math_failed = convert_mathjax_to_pressbooks_latex(soup)

    remove_leftover_mjx(soup)

    # Convert <em>n</em> → [latex]n[/latex] (only when it's a single letter/digit)
    em_latex_converted = convert_simple_em_to_latex(soup)

    # Alignment pass 1: for Pandoc-sourced HTML that still has math display spans
    equations_merged = merge_consecutive_equations(soup)
    # Alignment pass 2: for MathJax-rendered HTML already converted to [latex] blocks
    equations_merged += merge_consecutive_latex_equations(soup)

    # Wrap examples, replace \bullets, center latex paragraphs, normalize tables
    examples_wrapped = convert_examples_to_textboxes(soup) if WRAP_EXAMPLES else 0
    bq_unwrapped = unwrap_blockquotes(soup)
    cdot_replaced = replace_bullet_with_cdot(soup)
    latex_centered = center_latex_paragraphs(soup)
    latex_shrunk = shrink_long_latex(soup, max_chars=90)
    tables_normalized = normalize_tables_full_width(soup)
    soft_newlines_collapsed = collapse_soft_newlines_in_text(soup)
    br_removed = remove_br_between_display_math(soup)
    solution_markers_added = insert_solution_breaks_in_examples(soup)

    safe_write_text(out_path, serialize_html_readable(soup))
    return Report(
        file_in=path,
        file_out=out_path,
        toc_removed=toc_removed,
        heading_ids_created=heading_ids_created,
        section_topics_inserted=section_topics_inserted,
        math_converted=math_converted,
        math_failed=math_failed,
        examples_wrapped=examples_wrapped,
        soft_newlines_collapsed=soft_newlines_collapsed,
        equations_merged=equations_merged,
    )

def collect_html_files(input_path: Path) -> List[Path]:
    if input_path.is_file():
        return [input_path]
    if input_path.is_dir():
        return sorted([p for p in input_path.glob("*.html") if p.is_file()])
    return []

def main():
    input_path = Path(INPUT_PATH).expanduser()
    if not INPUT_PATH or not input_path.exists():
        print("❌ INPUT_PATH is blank or does not exist. Edit INPUT_PATH at the top of the script.")
        return

    files = collect_html_files(input_path)
    if not files:
        print(f"❌ No HTML files found at: {input_path}")
        return

    print(f"✅ Found {len(files)} HTML file(s). Processing...\n")

    reports: List[Report] = []
    for f in files:
        r = process_html_file(f)
        reports.append(r)

        print(f"— {r.file_in.name}")
        print(f"   out: {r.file_out.name}")
        print(f"   TOC blocks removed: {r.toc_removed}")
        print(f"   heading ids created: {r.heading_ids_created}")
        print(f"   Section Topics inserted: {r.section_topics_inserted}")
        print(f"   math converted / failed: {r.math_converted} / {r.math_failed}")
        print(f"   example blocks wrapped: {r.examples_wrapped}\n")
        print(f"   equation groups merged: {r.equations_merged}")
        print(f"   soft newlines collapsed: {r.soft_newlines_collapsed}")

    total_failed = sum(r.math_failed for r in reports)
    if total_failed:
        print(f"⚠️ NOTE: {total_failed} math item(s) failed conversion. "
              f"Those will remain as original content; spot-check those pages.")
    else:
        print("🎉 All detected MathML converted successfully.")

if __name__ == "__main__":
    main()
