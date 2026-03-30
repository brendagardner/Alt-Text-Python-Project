#!/usr/bin/env python3
"""
Created by Brenda Gardner via Claude CoWork on Feb 24, 2026.
PDF → Pressbooks-ready HTML for OER math content

What it does:
  1. Extracts text from a PDF using pdfplumber
  2. Classifies lines into structural blocks (title, H2 section, ALL-CAPS
     procedure box, Example, Solution, prose, math, list items)
  3. Groups blocks and renders Pressbooks HTML with:
       - Section Topics callout  (textbox--learning-objectives)  ← from PDF's TOC
       - Procedure boxes         (textbox--key-takeaways)        ← from ALL-CAPS headings
       - Example boxes           (textbox--examples)             ← from "Example N" labels
       - Tip/sidebar boxes       (textbox--tip)
       - Step explanations as inline text; equations centered below
  4. Math notation used:
       Display / centered:  \\[ ... \\]
       Inline in prose:     \\( ... \\)

═══════════════════════════════════════════════════════
  WHAT THIS SCRIPT DOES WELL (reliable automation):
═══════════════════════════════════════════════════════
  ✔  Detects section headings from the PDF's Table of Contents / Learning
     Outcomes list at the top — produces correct Section Topics anchors.
  ✔  Detects ALL-CAPS procedure box titles → key-takeaways textboxes.
  ✔  Detects "Example N" / "Solution" labels → example textboxes with hr break.
  ✔  Wraps simple inline math: percentages (6%), dollar amounts ($14.90),
     and plain fractions like 6/100 → \\(\\frac{6}{100}\\).
  ✔  Detects log(x) / ln(x) patterns and wraps them as inline LaTeX.
  ✔  Detects common exponent patterns (a^2, x^n) and wraps them.
  ✔  Merges prose lines into paragraphs; strips dot-leader TOC lines.
  ✔  Skips page numbers and boilerplate header/footer lines.

═══════════════════════════════════════════════════════
  WHAT NEEDS HUMAN REVIEW AFTER RUNNING (script limitations):
═══════════════════════════════════════════════════════
  ✘  FRACTIONS WITH STACKED NUMERATOR/DENOMINATOR
       pdfplumber extracts "36/100" on two lines as "36" then "100".
       The script cannot reliably know which lone number is a numerator
       and which is a denominator.  These appear as <<<MATH: ...>>> placeholders.
       Replace each one with the correct \\[\\dfrac{a}{b}\\] or \\(\\dfrac{a}{b}\\).

  ✘  NTH ROOT SYMBOLS (√, ∛, ⁸√)
       Radical symbols either disappear entirely or are extracted as garbage
       characters.  The script marks lines with radical-like content as
       <<<MATH: ...>>> placeholders.

  ✘  SUPERSCRIPTS AND SUBSCRIPTS
       PDF superscripts (exponents, indices) are usually extracted on a
       SEPARATE line from their base.  For example, "5³" may appear as
       "5" on one line and "3" on the next.  Simple cases (like "a^2") are
       auto-wrapped; anything complex becomes a placeholder.

  ✘  MULTI-COLUMN PAGE LAYOUTS
       When the PDF has a two-column layout (e.g., solution steps on the
       left, a sidebar box on the right), pdfplumber reads them interleaved.
       The resulting output may mix content from both columns.  You will
       need to separate and re-order manually.

  ✘  SPECIAL MATH SYMBOLS
       Symbols like → ⇒ · ≈ ≤ ≥ ∞ ≠ are sometimes extracted correctly and
       sometimes appear as garbled characters.  Review and correct as needed.

  ✘  VIDEO / IMAGE PLACEHOLDERS
       Videos and images appear as black boxes in PDF text extraction.
       They are silently dropped.  Add video embeds manually.

  ✘  BOLD / ITALIC SEMANTIC MEANING
       The script cannot detect which words are bold or italic in the PDF.
       Emphasis that carries meaning (e.g., "this is NOT multiplication")
       must be added manually.

  ✘  NUMBERED SUB-PROBLEMS WITH MATCHING SUB-SOLUTIONS
       When an example has sub-parts (1. … 2. … 3. …) and solutions are
       on the same page but widely separated, the script may not correctly
       pair each problem with its solution.

  ───────────────────────────────────────────────────────
  WORKFLOW RECOMMENDATION:
    1. Run the script → get a rough HTML draft.
    2. Search for  <<<MATH:  and replace each placeholder with proper LaTeX.
    3. Review all other math (auto-wrapped) for accuracy.
    4. Fix any content that was garbled by multi-column extraction.
    5. Add video embeds and restore bold/italic as needed.
  Typical manual cleanup time: 20–60 minutes depending on math density.
  ───────────────────────────────────────────────────────

Dependencies:
  pip install pdfplumber --break-system-packages

Usage:
  Edit INPUT_PATH (and optionally OUTPUT_PATH) below, then run:
      python3 PDFtoPressbooks.py
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

import pdfplumber

# ═══════════════════════════════════════════════════════
#  USER SETTINGS  (edit these)
# ═══════════════════════════════════════════════════════

INPUT_PATH  = "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/103x Files/Canvas_Textbook_Algebra_Review/Readiness Check Finance Spr24.pdf"   # PDF to convert
OUTPUT_PATH = ""   # blank → auto: same folder, _PB.html  ← must be "" not " "
OVERWRITE   = True

# Heading levels to include in Section Topics callout
SECTION_TOPICS_LEVELS = ("h2",)

# ═══════════════════════════════════════════════════════
#  PATTERNS
# ═══════════════════════════════════════════════════════

# ALL-CAPS lines → procedure / rule box titles
_ALLCAPS_RE   = re.compile(r'^[A-Z][A-Z][A-Z\s\-\/\(\)]{3,}$')

# (Intentionally removed: _SEC_HEAD_RE heuristic replaced by TOC-based detection)

# Matches "Example", "Example 1", "Example 1A", "Example 1A:", etc.
_EXAMPLE_RE   = re.compile(
    r'^\s*[Ee]xample\s*(\d+[A-Za-z]?)?\s*[:.\-]?\s*(.*)', re.S
)

# Matches "Solution", "Solution:", "Solutions:"
_SOLUTION_RE  = re.compile(r'^\s*[Ss]olutions?\s*:?\s*$')

# TOC filler dots
_TOC_DOTS_RE  = re.compile(r'\.{4,}')

# List items: numbered (1. or 1)) or bulleted (•, –, -)
_LIST_RE      = re.compile(r'^\s*(\d+[\.\)]|[•·\–\-])\s+\S')

# Simple inline math patterns worth auto-wrapping
_PCT_RE       = re.compile(r'(?<!\w)([\d]+(?:\.\d+)?)\s*%(?!\w)')
_DOLLAR_RE    = re.compile(r'\$([\d,]+(?:\.\d+)?)')
_FRAC_TEXT_RE = re.compile(r'\b(\d+)\s*/\s*(\d+)\b')          # e.g.  6/100

# Exponent patterns:  a^2  or  x^n  (letter/digit caret letter/digit)
_EXP_RE       = re.compile(r'([A-Za-z0-9])\^([A-Za-z0-9]+)')

# log / ln  patterns:  log(x)  ln(x)  log(100)
_LOG_RE       = re.compile(r'\b(log|ln)\s*\(([^)]+)\)')

# Radical / root indicator — lines containing √ or the word "root"
# These are flagged as needing a placeholder (can't auto-convert)
_RADICAL_RE   = re.compile(r'[√∛∜]|\bsqrt\b|\bnth root\b', re.I)

# ═══════════════════════════════════════════════════════
#  DATA MODEL
# ═══════════════════════════════════════════════════════

@dataclass
class Block:
    kind: str   # 'h1' | 'h2' | 'allcaps' | 'example' | 'solution' |
                # 'text' | 'math' | 'list_item' | 'blank' | 'tip_open' | 'tip_close'
    text: str


# ═══════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════

def slugify(text: str) -> str:
    s = text.lower().strip()
    s = re.sub(r'[^a-z0-9\s\-]', '', s)
    s = re.sub(r'\s+', '-', s).strip('-')
    return s or 'section'


def is_math_line(text: str) -> bool:
    """
    Heuristic: does this line look like a standalone mathematical expression
    rather than prose?  We check for a high ratio of non-alphabetic chars
    and the presence of math operators.
    """
    s = text.strip()
    if not s or len(s) > 100:
        return False
    # Count math-character types
    alpha   = sum(1 for c in s if c.isalpha())
    numeric = sum(1 for c in s if c.isdigit())
    math_op = sum(1 for c in s if c in '+-*/=()[]{}^_\\,.…')
    total   = len(s)

    if alpha / total > 0.55:
        return False          # mostly letters → prose

    # Needs at least one digit or operator to be math
    if (numeric + math_op) / total > 0.45:
        # Also require at least one digit or known math keyword
        if numeric > 0 or re.search(r'[=\+\-\*\/\^\\]', s):
            return True
    return False


def wrap_inline_math(text: str) -> str:
    r"""
    Auto-wrap recognizable inline math patterns with \( ... \).

    Handles automatically:
      - Percentages:       6%        →  \(6\%\)
      - Dollar amounts:    $14.90    →  \(\$14.90\)
      - Simple fractions:  6/100     →  \(\frac{6}{100}\)
      - log/ln:            log(100)  →  \(\log(100)\)
      - Simple exponents:  x^2       →  \(x^{2}\)

    Everything else needs manual LaTeX markup.
    """
    # $14.90  →  \(\$14.90\)
    text = _DOLLAR_RE.sub(lambda m: r'\(\$' + m.group(1) + r'\)', text)
    # 6%  →  \(6\%\)
    text = _PCT_RE.sub(lambda m: r'\(' + m.group(1) + r'\%\)', text)
    # 6/100  →  \(\frac{6}{100}\)
    text = _FRAC_TEXT_RE.sub(
        lambda m: r'\(\frac{' + m.group(1) + r'}{' + m.group(2) + r'}\)', text
    )
    # log(100)  →  \(\log(100)\)    ln(x)  →  \(\ln(x)\)
    text = _LOG_RE.sub(
        lambda m: r'\(' + '\\' + m.group(1) + r'(' + m.group(2) + r')\)', text
    )
    # x^2  →  \(x^{2}\)
    text = _EXP_RE.sub(
        lambda m: r'\(' + m.group(1) + r'^{' + m.group(2) + r'}\)', text
    )
    return text


def format_display(text: str) -> str:
    r"""Wrap a math expression as Pressbooks display (centered) math."""
    s = text.strip()
    return rf'\[{s}\]'


def math_placeholder(text: str) -> str:
    """Use when pdfplumber output is too garbled to auto-convert."""
    return f'<<<MATH: {text.strip()} >>>'


# ═══════════════════════════════════════════════════════
#  EXTRACTION
# ═══════════════════════════════════════════════════════

def extract_lines(pdf_path: Path) -> List[str]:
    """Return all text lines from every page, in reading order."""
    lines: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            raw = page.extract_text(layout=False)
            if raw:
                for ln in raw.splitlines():
                    lines.append(ln.rstrip())
    return lines


def extract_toc_items(raw_lines: List[str]) -> List[str]:
    """
    Parse the Learning Outcomes / Table of Contents block near the top of
    the PDF and return the cleaned topic strings.  These become the ONLY
    H2 section headings (and Section Topics anchors) in the output.

    Handles two formats found in OER review PDFs:

      Format 1 – dot-leader lines (most common):
          Table of Contents
          Convert between percents, decimals and fractions ......... 1
          Evaluate Exponents ................................................. 5

      Format 2 – bullet-point list (e.g. Trigonometry, Set Theory):
          By the end of this section, you will be able to:
          • Evaluate whole numbers in exponential form
          • Evaluate square roots

    The function stops collecting as soon as it hits a line that is neither
    a dot-leader item, a bullet item, a blank line, nor a pure page number —
    i.e. when it reaches the introductory paragraph of the body text.
    """
    items: List[str] = []
    seen_by_end = False   # saw "By the end of this section…"
    in_toc      = False   # actively collecting TOC items

    for line in raw_lines:
        s = line.strip()

        # ── Trigger lines ──────────────────────────────────────────────
        if re.match(r'^by the end of this section', s.lower()):
            seen_by_end = True
            continue

        if s.lower() == 'table of contents':
            in_toc = True
            continue

        # Bullet items immediately after "By the end…" → bullet-format TOC
        if seen_by_end and not in_toc and re.match(r'^[•·]', s):
            in_toc = True
            # fall through so this first bullet is also collected below

        if not in_toc:
            continue

        # ── Collecting TOC items ───────────────────────────────────────

        # Skip blanks and lone page numbers
        if not s or re.match(r'^\d{1,3}$', s):
            continue

        _boilerplate = {
            'learning outcomes', 'table of contents',
            'by the end of this section, you will be able to:',
            'by the end of this section you will be able to',
        }

        # Dot-leader format:  "Topic text ......... 1"
        if _TOC_DOTS_RE.search(s):
            topic = _TOC_DOTS_RE.split(s)[0].strip()
            if topic and topic.lower() not in _boilerplate:
                items.append(topic)
            continue

        # Bullet format:  "• Topic text"
        if re.match(r'^[•·]\s*', s):
            topic = re.sub(r'^[•·]\s*', '', s).strip()
            if topic and topic.lower() not in _boilerplate:
                items.append(topic)
            continue

        # Skip known boilerplate that can appear as plain lines
        if s.lower() in _boilerplate:
            continue

        # Anything else (real prose) signals the end of the TOC block
        break

    return items


# ═══════════════════════════════════════════════════════
#  CLASSIFICATION
# ═══════════════════════════════════════════════════════

def classify(line: str, toc_set: Optional[set] = None) -> Block:
    """
    Map one raw text line to a Block with a semantic kind.

    toc_set: a set of lowercased TOC item strings.  When provided, any line
             whose text matches a TOC item (case-insensitive, allowing a
             parenthetical suffix like '(Review)') is classified as 'h2'.
             Pass this so that section headings are protected from being
             merged into prose paragraphs.
    """
    s = line.strip()

    if not s:
        return Block('blank', '')

    # Skip TOC dot-leaders
    if _TOC_DOTS_RE.search(s):
        return Block('blank', '')

    # Skip obvious page-number-only lines
    if re.match(r'^\d{1,3}$', s):
        return Block('blank', '')

    # Detect section headings from the TOC list BEFORE any other check,
    # so they won't be merged into prose paragraphs later.
    if toc_set:
        # Strip common suffixes that appear on body headings but not in the TOC:
        #   (1) a trailing parenthetical like "(Review)" or "(continued)"
        #   (2) a trailing standalone number, e.g. "Heading text 200" where
        #       200 is a stray value from an adjacent column in the PDF layout
        s_clean = re.sub(r'\s*\([^)]*\)\s*$', '', s).strip()   # strip parens
        s_clean = re.sub(r'\s+\d+\s*$', '', s_clean).strip()   # strip trailing number
        if s.lower() in toc_set:
            return Block('h2', s)
        if s_clean.lower() in toc_set:
            return Block('h2', s_clean)

    # Example starter
    m = _EXAMPLE_RE.match(s)
    if m:
        return Block('example', s)

    # Solution label
    if _SOLUTION_RE.match(s):
        return Block('solution', s)

    # ALL-CAPS line → procedure/rule box heading
    if _ALLCAPS_RE.match(s):
        return Block('allcaps', s)

    # Numbered or bulleted list item
    if _LIST_RE.match(s):
        return Block('list_item', s)

    # Math-heavy line
    if is_math_line(s):
        return Block('math', s)

    # Default: prose
    return Block('text', s)


def merge_prose_lines(blocks: List[Block]) -> List[Block]:
    """
    Collapse consecutive 'text' blocks (separated only by blanks)
    into single paragraph blocks.  A blank line acts as a paragraph
    boundary; a 'math', 'allcaps', etc. block also terminates prose.
    """
    out: List[Block] = []
    buf: List[str] = []

    def flush():
        if buf:
            out.append(Block('text', ' '.join(buf)))
            buf.clear()

    for b in blocks:
        if b.kind == 'text':
            buf.append(b.text)
        elif b.kind == 'blank':
            flush()
        else:
            flush()
            out.append(b)

    flush()
    return out


# ═══════════════════════════════════════════════════════
#  HTML RENDERING
# ═══════════════════════════════════════════════════════

def _p(text: str, cls: str = '', style: str = '') -> str:
    attrs = ''
    if cls:
        attrs += f' class="{cls}"'
    if style:
        attrs += f' style="{style}"'
    return f'<p{attrs}>{text}</p>'


def _open_box(modifier: str, title: str) -> str:
    return (
        f'<div class="textbox textbox--{modifier}">\n'
        f'<header class="textbox__header">\n'
        f'<p class="textbox__title">{title}</p>\n'
        f'</header>\n'
        f'<div class="textbox__content">\n'
    )


def _close_box() -> str:
    return '</div>\n</div>'


def render_math_block(raw: str) -> str:
    r"""
    Decide how to render a 'math' line:
      - Contains radical symbol  → <<<MATH: ...>>> placeholder (can't auto-convert)
      - Short, clean expression  → \[ ... \]
      - Longer / garbled         → <<<MATH: ...>>> placeholder
    """
    s = raw.strip()

    # Lines with radical characters cannot be auto-converted
    if _RADICAL_RE.search(s):
        return _p(math_placeholder(s))

    # Very short expressions are likely clean enough
    if len(s) <= 40 and re.search(r'\d', s):
        # Clean up common pdfplumber artifacts
        s = s.replace('⋅', r'\cdot ')
        s = s.replace('…', r'\ldots ')
        # Auto-wrap log/ln and exponents
        s = _LOG_RE.sub(
            lambda m: '\\' + m.group(1) + r'(' + m.group(2) + r')', s
        )
        s = _EXP_RE.sub(
            lambda m: m.group(1) + r'^{' + m.group(2) + r'}', s
        )
        return format_display(s)

    # Longer or suspiciously encoded → placeholder
    return _p(math_placeholder(s))


def render_content_line(line: str, kind: str) -> str:
    """Turn one classified line into an HTML snippet."""
    if kind == 'math':
        raw = line.strip()
        # Radical lines → always placeholder
        if _RADICAL_RE.search(raw):
            return _p(math_placeholder(raw))
        if len(raw) <= 60 and re.search(r'[\d=\+\-]', raw):
            raw = raw.replace('⋅', r'\cdot ').replace('…', r'\ldots ')
            raw = _LOG_RE.sub(
                lambda m: '\\' + m.group(1) + r'(' + m.group(2) + r')', raw
            )
            raw = _EXP_RE.sub(
                lambda m: m.group(1) + r'^{' + m.group(2) + r'}', raw
            )
            return _p(format_display(raw), 'align-center')
        return _p(math_placeholder(raw))
    else:
        return _p(wrap_inline_math(line.strip()))


def render_step_pairs(pairs: List[Tuple[str, str]]) -> str:
    r"""
    Render a list of (explanation, equation) pairs as the standard format:
        <p>Explanation text</p>
        <p class="align-center">\[ equation \]</p>
    Either part may be empty string to omit that element.
    """
    parts: List[str] = []
    for explanation, equation in pairs:
        if explanation.strip():
            parts.append(_p(wrap_inline_math(explanation.strip())))
        if equation.strip():
            eq = equation.strip().replace('⋅', r'\cdot ').replace('…', r'\ldots ')
            parts.append(_p(format_display(eq), 'align-center'))
    return '\n'.join(parts)


# ═══════════════════════════════════════════════════════
#  SOLUTION STEP PARSER
# ═══════════════════════════════════════════════════════
# Inside a Solution section, lines often alternate:
#   prose explanation line
#   math result line
# We pair them up so explanation text comes inline above its equation.

def pair_solution_lines(lines: List[Block]) -> str:
    """
    Given the blocks inside a Solution, produce the step-by-step HTML:
      <p>explanation</p>
      <p class="align-center">\\[equation\\]</p>
    """
    parts: List[str] = []
    pending_text: Optional[str] = None

    for b in lines:
        if b.kind == 'blank':
            if pending_text:
                parts.append(_p(wrap_inline_math(pending_text)))
                pending_text = None
            continue

        if b.kind == 'math':
            # Output any pending explanation first, then the centered equation
            if pending_text:
                parts.append(_p(wrap_inline_math(pending_text)))
                pending_text = None
            eq = b.text.strip()
            eq = eq.replace('⋅', r'\cdot ').replace('…', r'\ldots ')
            if len(eq) <= 80:
                parts.append(_p(format_display(eq), 'align-center'))
            else:
                parts.append(_p(math_placeholder(eq)))

        elif b.kind == 'text':
            if pending_text:
                # Two prose lines in a row → join as one paragraph
                pending_text += ' ' + b.text.strip()
            else:
                pending_text = b.text.strip()

        elif b.kind == 'list_item':
            if pending_text:
                parts.append(_p(wrap_inline_math(pending_text)))
                pending_text = None
            item = re.sub(r'^\s*[\d]+[\.\)]\s*', '', b.text)
            item = re.sub(r'^\s*[•·\–\-]\s*', '', item)
            parts.append(_p(wrap_inline_math(item)))

    if pending_text:
        parts.append(_p(wrap_inline_math(pending_text)))

    return '\n'.join(parts)


# ═══════════════════════════════════════════════════════
#  MAIN PIPELINE
# ═══════════════════════════════════════════════════════

def build_html(blocks: List[Block], toc_items: Optional[List[str]] = None) -> str:
    """
    Walk the classified + merged blocks and produce Pressbooks HTML.

    toc_items: cleaned topic strings from extract_toc_items().  When
               provided, these become the sole H2 headings and Section
               Topics anchors.  If None (or empty), H2s are skipped
               and no Section Topics callout is emitted.
    """

    # ── Pre-pass: locate H1 and build H2 anchor list ──────────────────
    h1_text = ''
    h2_items: List[Tuple[str, str]] = []   # (anchor_id, display_text)

    skip_keywords = {
        'learning outcomes', 'table of contents',
        'by the end of this section you will be able to',
        'by the end of this section, you will be able to:',
    }

    # Find title: the first non-blank text block that isn't boilerplate
    for b in blocks:
        if b.kind == 'text' and b.text.strip():
            if b.text.strip().lower() not in skip_keywords:
                h1_text = b.text.strip()
            break

    # H2 anchors come ONLY from the TOC extracted at the top of the PDF
    if toc_items:
        for topic in toc_items:
            anchor = slugify(topic)
            if not any(a == anchor for a, _ in h2_items):
                h2_items.append((anchor, topic))

    # ── Build Section Topics callout ───────────────────────────────────
    section_topics_html = ''
    if h2_items:
        links = '\n'.join(
            f'<p style="padding-left: 40px">'
            f'<a href="#{aid}">{txt}</a></p>'
            for aid, txt in h2_items
        )
        section_topics_html = (
            '<div class="textbox textbox--learning-objectives">\n'
            '<header class="textbox__header">\n'
            '<p class="textbox__title">Section Topics</p>\n'
            '</header>\n'
            '<div class="textbox__content">\n'
            '<div>\n'
            f'{links}\n'
            '</div>\n'
            '</div>\n'
            '</div>'
        )

    # ── Main rendering pass ────────────────────────────────────────────
    parts: List[str] = []

    # H1
    if h1_text:
        parts.append(f'<h1>{h1_text}</h1>')

    # Section Topics
    if section_topics_html:
        parts.append(section_topics_html)

    # State machine
    i = 0
    found_h1 = False

    # Buffers
    in_procedure    = False
    proc_title      = ''
    proc_items      : List[str] = []
    proc_text_lines : List[str] = []   # prose lines inside a procedure box

    in_example      = False
    ex_label        = ''
    ex_problem      : List[Block] = []
    ex_solution     : List[Block] = []
    in_solution     = False

    pending_text    : List[str] = []

    # ── flush helpers ─────────────────────────────────────────────────

    def flush_text():
        nonlocal pending_text
        if pending_text:
            parts.append(_p(wrap_inline_math(' '.join(pending_text))))
            pending_text = []

    def flush_procedure():
        nonlocal in_procedure, proc_title, proc_items, proc_text_lines
        if not in_procedure:
            return
        title_display = proc_title.title()
        inner = ''
        if proc_items:
            li_html = '\n'.join(f'<li>{item}</li>' for item in proc_items)
            inner = f'<ol>\n{li_html}\n</ol>\n'
        elif proc_text_lines:
            inner = '\n'.join(_p(ln) for ln in proc_text_lines) + '\n'
        parts.append(_open_box('key-takeaways', title_display) + inner + _close_box())
        in_procedure = False
        proc_title = ''
        proc_items = []
        proc_text_lines = []

    def flush_example():
        nonlocal in_example, ex_label, ex_problem, ex_solution, in_solution

        if not in_example:
            return

        # Render problem lines
        prob_parts: List[str] = []
        # Check if the problem is a numbered list
        list_items = [b for b in ex_problem if b.kind == 'list_item']
        prose_items = [b for b in ex_problem if b.kind not in ('list_item', 'blank')]

        if list_items and len(list_items) == len([b for b in ex_problem if b.kind != 'blank']):
            # Pure list problem
            strip_lead = re.compile(r'^\s*[\d]+[\.)\s]*')
            li_html = '\n'.join(
                f'<li>{wrap_inline_math(strip_lead.sub("", b.text))}</li>'
                for b in list_items
            )
            prob_parts.append(f'<ol>\n{li_html}\n</ol>')
        else:
            for b in ex_problem:
                if b.kind == 'blank':
                    continue
                prob_parts.append(render_content_line(b.text, b.kind))

        # Render solution lines using the step-pairing logic
        sol_html = pair_solution_lines(ex_solution) if ex_solution else ''

        # Assemble the textbox
        box_html = (
            _open_box('examples', ex_label) +
            '\n'.join(prob_parts) + '\n' +
            '<hr />\n'
            '<strong>Solution</strong>\n' +
            sol_html + '\n' +
            _close_box()
        )
        parts.append(box_html)

        in_example = False
        ex_label = ''
        ex_problem = []
        ex_solution = []
        in_solution = False

    # ── Walk blocks ───────────────────────────────────────────────────

    known_skip = {h1_text.strip().lower()} | skip_keywords

    for b in blocks:
        s = b.text.strip()

        # ── Skip the title line (already emitted as H1)
        if b.kind == 'text' and s.lower() == h1_text.strip().lower() and not found_h1:
            found_h1 = True
            continue

        # ── Skip boilerplate
        if s.lower() in known_skip:
            continue

        # ── Section heading (H2) ──────────────────────────────────────
        # Match against TOC items (case-insensitive to allow minor
        # capitalisation differences between TOC and body text)
        h2_match = next(
            ((aid, txt) for aid, txt in h2_items
             if s.lower() == txt.lower()),
            None
        )
        if h2_match:
            flush_text()
            flush_procedure()
            flush_example()
            anchor, display = h2_match
            parts.append(f'<h2 id="{anchor}"><a id="{anchor}"></a>{display}</h2>')
            continue

        # ── ALL-CAPS → procedure box ──────────────────────────────────
        if b.kind == 'allcaps':
            flush_text()
            flush_example()
            flush_procedure()
            in_procedure = True
            proc_title = s
            continue

        # Content inside a procedure box
        if in_procedure:
            if b.kind == 'list_item':
                item = re.sub(r'^\s*[\d]+[\.\)]\s*', '', s)
                item = re.sub(r'^\s*[•·\–\-]\s*', '', item)
                proc_items.append(wrap_inline_math(item))
                continue
            elif b.kind == 'text':
                proc_text_lines.append(wrap_inline_math(s))
                continue
            elif b.kind == 'blank':
                continue
            else:
                # Non-list, non-text block closes the procedure box
                flush_procedure()
                # fall through to handle below

        # ── Example ───────────────────────────────────────────────────
        if b.kind == 'example':
            flush_text()
            flush_procedure()
            flush_example()
            m = _EXAMPLE_RE.match(s)
            num  = (m.group(1) or '').strip() if m else ''
            rest = (m.group(2) or '').strip() if m else ''
            ex_label   = f'Example {num}'.strip() if num else 'Example'
            in_example = True
            in_solution = False
            if rest:
                ex_problem.append(Block('text', rest))
            continue

        # ── Solution marker ───────────────────────────────────────────
        if b.kind == 'solution' and in_example:
            in_solution = True
            continue

        # ── Content inside an example ─────────────────────────────────
        if in_example:
            if in_solution:
                ex_solution.append(b)
            else:
                ex_problem.append(b)
            continue

        # ── Regular prose / math ──────────────────────────────────────
        if b.kind == 'text':
            pending_text.append(s)
        elif b.kind == 'math':
            flush_text()
            eq = s.replace('⋅', r'\cdot ').replace('…', r'\ldots ')
            if len(eq) <= 80:
                parts.append(_p(format_display(eq), 'align-center'))
            else:
                parts.append(_p(math_placeholder(eq)))
        elif b.kind == 'list_item':
            flush_text()
            item = re.sub(r'^\s*[\d]+[\.\)]\s*', '', s)
            item = re.sub(r'^\s*[•·\–\-]\s*', '', item)
            parts.append(_p(wrap_inline_math(item)))
        elif b.kind == 'blank':
            flush_text()

    # ── Final flushes ─────────────────────────────────────────────────
    flush_text()
    flush_procedure()
    flush_example()

    return '\n\n'.join(parts) + '\n'


# ═══════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════

def main():
    input_path = Path(INPUT_PATH).expanduser()
    if not input_path.exists():
        print(f"❌  INPUT_PATH not found: {input_path}")
        return

    out_path = (
        Path(OUTPUT_PATH).expanduser()
        if OUTPUT_PATH
        else input_path.with_name(f"{input_path.stem}_PB.html")
    )

    if out_path.exists() and not OVERWRITE:
        print(f"⚠️   Output exists and OVERWRITE=False — skipping.")
        return

    print(f"📄  Reading:  {input_path}")
    raw_lines  = extract_lines(input_path)
    toc_items  = extract_toc_items(raw_lines)
    # Build a lowercased set for O(1) heading lookup during classification
    toc_set    = {item.lower() for item in toc_items}
    blocks     = [classify(ln, toc_set) for ln in raw_lines]
    blocks     = merge_prose_lines(blocks)

    if toc_items:
        print(f"📑  TOC items found ({len(toc_items)}):")
        for item in toc_items:
            print(f"    • {item}")
    else:
        print("⚠️   No TOC / Learning Outcomes list detected — no Section Topics callout will be generated.")

    html = build_html(blocks, toc_items)
    out_path.write_text(html, encoding='utf-8')

    placeholder_count = html.count('<<<MATH:')
    print(f"✅  Saved  →  {out_path}")
    print(f"\n📋  Summary:")
    print(f"    Lines extracted  : {len(raw_lines)}")
    print(f"    Blocks classified: {len(blocks)}")
    print(f"    Math placeholders: {placeholder_count}")
    if placeholder_count:
        print(
            f"\n⚠️   {placeholder_count} complex math expression(s) need manual LaTeX markup.\n"
            f"    Search for  <<<MATH:  in the output and replace with proper \\[...\\]."
        )
    else:
        print("\n🎉  No complex math placeholders — review equations for accuracy.")


if __name__ == "__main__":
    main()
