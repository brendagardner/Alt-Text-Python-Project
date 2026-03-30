# This script scrapes images + alt text from URLs and then calls the Abacus
# vision API to generate candidate alt text variants.
# Updated by Brenda Gardner, December 1 2025.
# To improve - vision model using context - review last lines of STYLE_GUIDE_RULES in alttext_utils.py

# Step 1. Capture URLs and Subjects.
# a. Go to Website TOC and copy paste URLS.
# use in conjunction with https://slccbruins-my.sharepoint.com/:x:/g/personal/bgardn61_slcc_edu/IQD6AhzRzSxNRohDxUOKgeNYAQrWIctBAHw6ag8_bIYE-U8?e=sQFPWW

# Step 2. Configure Environment for Abacus API key.  (BuildAPIVisionKey)
#      Note:  LLama4-Maverick is only Input: $0.14/M, Output: $0.59/M, 1 CR ~ 7 images.  $10/month = 20,000 CR ~ 150,000 images (Graph Theory images not accurate)
#      If Llama doesn't work well, choose "gpt-4o-2024-11-20".  Input: $2.50/M, Output: $10.00/M.  3 CR ~ 1 images (Graph Theory images still not accurate)

# Step 3. Run this script.
# Step 4. VBA script to put images in cell and filter, as desired.

# Master 30-column layout (A–AD) — standardized across all scripts:
# A:  Image Thumbnail
# B:  Image Filename
# C:  Picture ID          — blank for website sources
# D:  Image URL
# E:  Image Size (original width x height)
# F:  Image Topic         (AI topic 1-2 words)
# G:  Image Context       (nearby text / figcaption)
# H:  Page #              — blank for website sources
# I:  Section/Module/Figure #
# J:  Filename (Hyperlinked) or Webpage Name (URL)
# K:  Grade Level         (from config)
# L:  Course Name         (from config)
# M:  Original Source Author
# N:  Copyright
# O:  Date of Audit       — auto-filled
# P:  Original Alt Text
# Q:  Original Alt Text Character Count  — =LEN(P)
# R:  Original Alt Text Notes            — blank
# S:  A.I. Short Alt Text
# T:  A.I. Med Alt Text
# U:  A.I. Long Alt Text
# V:  A.I. Notes
# W:  New/Suggested Alt Text             — blank
# X:  New/Suggested Alt Text Count       — =LEN(W)
# Y:  New/Suggested Alt Text Notes       — blank
# Z:  Status - Alt Text Completed?       — blank
# AA: Reviewer Notes                     — blank
# AB: In Word?                           — blank for website sources
# AC: In Excel?                          — blank
# AD: Date of Completion                 — blank

import os
import json
import base64
import warnings
from io import BytesIO
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional

from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from urllib3.exceptions import NotOpenSSLWarning

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from helpers import alttext_utils as utils

# =========================
# CONFIG — EDIT THESE
# =========================

# These will be written into every row (can be left "" if you want to fill later).
DEFAULT_GRADE_BAND = "6-8"           # e.g. "10–12"
COURSE_TOPIC = "Elementary Algebra"
BOOK_SOURCE_NAME = "Lumen Find LCD"

# Folder where the Excel output file will be written
OUTPUT_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/"
    "Sabbatical Files - Personal (Unshared)/0990 Information"
)

# List of (URL, Section/Topic) pairs
URLS_AND_TOPICS = [
    (
        "https://courses.lumenlearning.com/slcc-elementaryalgebra/chapter/finding-least-common-denominators/",
        "Finding LCD",
    ),
    # Add more (url, section/topic) pairs as needed
]

# ---------- Abacus AI configuration ----------

ABACUS_API_KEY = os.getenv("ABACUS_API_KEY")
ROUTE_LLM_ENDPOINT = "https://routellm.abacus.ai/v1/chat/completions"
VISION_MODEL = "meta-llama/Llama-4-Maverick-17B-128E-Instruct-FP8"
API_TIMEOUT_SEC = 180

if not ABACUS_API_KEY:
    raise EnvironmentError("❌ Missing ABACUS_API_KEY environment variable.")


# =========================
# Abacus helpers
# =========================

def build_prompt(grade_band: str, course_topic: str, context: str) -> str:
    """
    Build the prompt for Abacus using the shared STYLE_GUIDE_RULES from utils.
    """
    context_snippet = context or ""
    return f"{utils.STYLE_GUIDE_RULES}\nContext snippet: {context_snippet}\nReturn ONLY JSON."


def call_abacus(img_bytes: bytes, prompt: str) -> Dict[str, str]:
    """
    Call the Abacus vision endpoint and return the parsed JSON dict.
    Expected keys: short, standard, long, notes, topic (topic optional).
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
    r = requests.post(
        ROUTE_LLM_ENDPOINT,
        headers=headers,
        json=payload,
        timeout=API_TIMEOUT_SEC,
    )
    r.raise_for_status()
    content = r.json().get("choices", [{}])[0].get("message", {}).get("content", "")

    try:
        js = json.loads(content.strip("`").replace("json", "").strip())
    except Exception:
        # Fallback on parse error
        return {"short": "", "standard": "", "long": "", "notes": "", "topic": ""}

    if isinstance(js, dict) and "topic" not in js:
        js["topic"] = ""
    return js


# =========================
# Scraper + Excel builder
# =========================

class CombinedImageExtractor:
    def __init__(self, proxies: Optional[Dict[str, str]] = None):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/91.0.4472.124 Safari/537.36"
            )
        })
        if proxies:
            self.session.proxies.update(proxies)

    # ---------- Scraping ----------

    def extract_images_from_url(self, url: str, topic: str = "") -> List[Dict[str, str]]:
        """Extract all images and alt text from a webpage and its iframes."""
        try:
            if not url:
                return []

            print(f"Processing main page: {url}")
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")

            images_data: List[Dict[str, str]] = []

            # Extract images from main page
            main_images = self._extract_images_from_soup(
                soup, url, topic, source_type="main page"
            )
            images_data.extend(main_images)

            # Find and process all iframes
            iframes = soup.find_all("iframe")
            print(f"Found {len(iframes)} iframes to process")

            for i, iframe in enumerate(iframes):
                iframe_src = iframe.get("src")
                if not iframe_src:
                    continue

                iframe_url = urljoin(url, iframe_src)

                # Skip iframes that likely contain video or ads
                skip_domains = [
                    "youtube.com", "vimeo.com", "facebook.com",
                    "twitter.com", "ads.", "analytics.", "google.com/maps",
                ]
                if any(domain in iframe_url.lower() for domain in skip_domains):
                    print(
                        f"Skipping iframe {i + 1}: {iframe_url} "
                        f"(likely ads/social/video)"
                    )
                    continue

                print(f"Processing iframe {i + 1}: {iframe_url}")
                iframe_images = self._extract_images_from_iframe(
                    iframe_url, topic, original_webpage=url
                )
                images_data.extend(iframe_images)

            print(f"Total found: {len(images_data)} images for {url}")
            return images_data

        except Exception as e:
            print(f"Error extracting from {url}: {str(e)}")
            return []

    def _extract_image_context(self, img, max_chars: int = 500) -> str:
        """
        Try to pull a meaningful text context for an <img> element.

        Priority:
          1) <figcaption> (if inside a <figure>)
          2) Closest previous sibling <p>/<li>/<div> with real text
          3) Walk up ancestors, checking their previous siblings
          4) Fallback: closest following <p>/<li>/<div> with text
        """
        def clean_text(node):
            if not node:
                return ""
            txt = node.get_text(" ", strip=True)
            txt = " ".join(txt.split())
            return txt[:max_chars]

        # 1) If inside a <figure>, prefer <figcaption>
        figure = img.find_parent("figure")
        if figure:
            cap = figure.find("figcaption")
            if cap:
                cap_text = clean_text(cap)
                if cap_text:
                    return cap_text

        # helper to scan previous siblings
        def prev_sibling_text(node):
            for sib in getattr(node, "previous_siblings", []):
                if getattr(sib, "name", None) in ("p", "li", "div"):
                    txt = clean_text(sib)
                    if txt:
                        return txt
            return ""

        # 2) Previous siblings of img
        ctx = prev_sibling_text(img)
        if ctx:
            return ctx

        # 3) Walk up ancestors and check their previous siblings
        parent = img.parent
        while parent is not None and getattr(parent, "name", None) not in ("body", "html"):
            ctx = prev_sibling_text(parent)
            if ctx:
                return ctx
            parent = parent.parent

        # 4) Fallback: next siblings
        for sib in getattr(img, "next_siblings", []):
            if getattr(sib, "name", None) in ("p", "li", "div"):
                txt = clean_text(sib)
                if txt:
                    return txt

        return ""

    def _extract_images_from_soup(
        self,
        soup: BeautifulSoup,
        source_url: str,
        topic: str,
        source_type: str = "page",
        original_webpage: Optional[str] = None,
    ) -> List[Dict[str, str]]:
        """Helper method to extract images from a BeautifulSoup object."""
        images_data: List[Dict[str, str]] = []
        img_tags = soup.find_all("img")

        for img in img_tags:
            src = img.get("src")
            if not src:
                continue

            # 1) Skip images inside <header> or <footer>
            if img.find_parent(["header", "footer"]):
                continue

            # 2) Skip by parent class/id (common footer/header patterns)
            parent = img.parent
            if parent:
                skip_classes = [
                    "site-header", "header", "footer", "site-footer",
                    "footer-logo", "masthead", "page-header",
                ]
                skip_ids = ["header", "footer", "colophon", "masthead"]

                if any(cls in skip_classes for cls in parent.get("class", [])):
                    continue
                if parent.get("id") in skip_ids:
                    continue

            # 3) Skip by URL filename patterns (logos, banners, icons)
            full_url = urljoin(source_url, src)
            full_url_lower = full_url.lower()

            skip_patterns = [
                "logo", "banner", "footer", "header",
                "favicon", "sprite", "icon",
            ]
            if any(p in full_url_lower for p in skip_patterns):
                continue

            # SAFE — process the content image
            alt_text = img.get("alt", "") or ""
            title_text = img.get("title", "") or ""
            author = img.get("data-author") or img.get("data-credit") or ""
            copyright_text = img.get("data-copyright") or img.get("copyright") or ""

            image_context = self._extract_image_context(img)

            images_data.append({
                "image_location": full_url,
                "alt_text": alt_text,
                "webpage": original_webpage if original_webpage else source_url,
                "section": topic,
                "image_title": title_text,
                "author": author,
                "copyright": copyright_text,
                "source_type": source_type,
                "image_context": image_context,
            })

        print(f"Found {len(images_data)} images in {source_type}")
        return images_data

    def _extract_images_from_iframe(
        self,
        iframe_url: str,
        topic: str,
        original_webpage: str,
    ) -> List[Dict[str, str]]:
        """Extract images from an iframe."""
        try:
            response = self.session.get(iframe_url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "html.parser")

            return self._extract_images_from_soup(
                soup,
                iframe_url,
                topic,
                source_type=f"iframe ({iframe_url})",
                original_webpage=original_webpage,
            )
        except Exception as e:
            print(f"Error processing iframe {iframe_url}: {str(e)}")
            return []

    # ---------- Excel builder ----------

    def create_excel_with_images(
        self,
        all_images_data: List[Dict[str, str]],
        output_filename: Path,
        file_source: str = "",
        grade_band: str = "",
        course_topic: str = "",
    ) -> None:
        if not all_images_data:
            print("No images to export.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Website Alt Text AI Output"

        # Headers — Master 30-column layout A–AD
        headers = [
            "Image Thumbnail",                               # A  (1)
            "Image Filename",                                # B  (2)
            "Picture ID",                                    # C  (3) — blank for website sources
            "Image URL",                                     # D  (4)
            "Image Size",                                    # E  (5)
            "Image Topic",                                   # F  (6)
            "Image Context",                                 # G  (7)
            "Page #",                                        # H  (8) — blank for website sources
            "Section/Module/Figure #",                       # I  (9)
            "Filename (Hyperlinked) or Webpage Name (URL)",  # J  (10)
            "Grade Level",                                   # K  (11)
            "Course Name",                                   # L  (12)
            "Original Source Author",                        # M  (13)
            "Copyright",                                     # N  (14)
            "Date of Audit",                                 # O  (15) — auto-filled
            "Original Alt Text",                             # P  (16)
            "Original Alt Text Character Count",             # Q  (17) — =LEN(P)
            "Original Alt Text Notes",                       # R  (18) — blank
            "A.I. Short Alt Text",                           # S  (19)
            "A.I. Med Alt Text",                             # T  (20)
            "A.I. Long Alt Text",                            # U  (21)
            "A.I. Notes",                                    # V  (22)
            "New/Suggested Alt Text",                        # W  (23) — blank
            "New/Suggested Alt Text Count",                  # X  (24) — =LEN(W)
            "New/Suggested Alt Text Notes",                  # Y  (25) — blank
            "Status - Alt Text Completed?\n(Original/Suggested/N)",  # Z  (26) — blank
            "Reviewer Notes",                                # AA (27) — blank
            "In Word?",                                      # AB (28) — blank for website sources
            "In Excel?",                                     # AC (29) — blank
            "Date of Completion",                            # AD (30) — blank
        ]

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)

        # Column widths — Master 30-column layout
        ws.column_dimensions["A"].width = utils._pixels_to_column_width(600)  # Image Thumbnail
        ws.column_dimensions["B"].width = 30   # Image Filename
        ws.column_dimensions["C"].width = 18   # Picture ID
        ws.column_dimensions["D"].width = 40   # Image URL
        ws.column_dimensions["E"].width = 15   # Image Size
        ws.column_dimensions["F"].width = 25   # Image Topic
        ws.column_dimensions["G"].width = 40   # Image Context
        ws.column_dimensions["H"].width = 8    # Page #
        ws.column_dimensions["I"].width = 25   # Section/Module/Figure #
        ws.column_dimensions["J"].width = 40   # Filename (Hyperlinked) or Webpage Name
        ws.column_dimensions["K"].width = 12   # Grade Level
        ws.column_dimensions["L"].width = 20   # Course Name
        ws.column_dimensions["M"].width = 20   # Original Source Author
        ws.column_dimensions["N"].width = 20   # Copyright
        ws.column_dimensions["O"].width = 16   # Date of Audit
        ws.column_dimensions["P"].width = 50   # Original Alt Text
        ws.column_dimensions["Q"].width = 16   # Original Alt Text Character Count
        ws.column_dimensions["R"].width = 40   # Original Alt Text Notes
        ws.column_dimensions["S"].width = 50   # A.I. Short Alt Text
        ws.column_dimensions["T"].width = 50   # A.I. Med Alt Text
        ws.column_dimensions["U"].width = 50   # A.I. Long Alt Text
        ws.column_dimensions["V"].width = 40   # A.I. Notes
        ws.column_dimensions["W"].width = 50   # New/Suggested Alt Text
        ws.column_dimensions["X"].width = 18   # New/Suggested Alt Text Count
        ws.column_dimensions["Y"].width = 40   # New/Suggested Alt Text Notes
        ws.column_dimensions["Z"].width = 20   # Status - Alt Text Completed?
        ws.column_dimensions["AA"].width = 30  # Reviewer Notes
        ws.column_dimensions["AB"].width = 12  # In Word?
        ws.column_dimensions["AC"].width = 12  # In Excel?
        ws.column_dimensions["AD"].width = 18  # Date of Completion

        current_row = 2

        for img_info in all_images_data:
            image_url = img_info.get("image_location", "") or ""
            alt_text = img_info.get("alt_text", "") or ""
            webpage = img_info.get("webpage", "") or ""
            section = img_info.get("section", "") or ""
            image_title = img_info.get("image_title", "") or ""
            author = img_info.get("author", "") or ""
            copyright_text = img_info.get("copyright", "") or ""
            image_context = img_info.get("image_context", "") or ""

            # B: image filename from URL path
            parsed = urlparse(image_url)
            image_filename = Path(parsed.path).name if parsed.path else ""

            # K/L: grade level & course topic
            grade_value = grade_band or ""
            course_value = course_topic or ""

            # P: Original Alt Text
            current_alt = alt_text

            # Column B (2): Image Filename
            ws.cell(row=current_row, column=2, value=image_filename)

            # Column C (3): Picture ID — blank for website sources

            # Column D (4): Image URL (hyperlink)
            ws.cell(row=current_row, column=4, value=image_url)
            if image_url:
                ws.cell(row=current_row, column=4).hyperlink = image_url
                ws.cell(row=current_row, column=4).style = "Hyperlink"

            # Column G (7): Image Context
            ws.cell(row=current_row, column=7, value=image_context)

            # Column H (8): Page # — blank for website sources

            # Column I (9): Section/Module/Figure #
            ws.cell(row=current_row, column=9, value=section)

            # Column J (10): Webpage/Source (hyperlink)
            ws.cell(row=current_row, column=10, value=webpage)
            if webpage:
                ws.cell(row=current_row, column=10).hyperlink = webpage
                ws.cell(row=current_row, column=10).style = "Hyperlink"

            # Column K (11): Grade Level
            ws.cell(row=current_row, column=11, value=grade_value)

            # Column L (12): Course Name
            ws.cell(row=current_row, column=12, value=course_value)

            # Column M (13): Original Source Author
            ws.cell(row=current_row, column=13, value=author)

            # Column N (14): Copyright
            ws.cell(row=current_row, column=14, value=copyright_text)

            # Column O (15): Date of Audit — auto-filled with today's date
            ws.cell(row=current_row, column=15, value=datetime.now().strftime("%m/%d/%Y"))

            # Column P (16): Original Alt Text
            ws.cell(row=current_row, column=16, value=current_alt)

            # Column Q (17): Original Alt Text Character Count — formula =LEN(P)
            ws.cell(row=current_row, column=17, value=f"=LEN(P{current_row})")

            # Column R (18): Original Alt Text Notes — blank

            # --- Download image once: use for thumbnail + API ---
            try:
                if image_url:
                    print(f"Downloading image for row {current_row}: {image_url}")
                    resp = self.session.get(image_url, timeout=10)
                    resp.raise_for_status()
                    img_bytes_raw = resp.content
                else:
                    img_bytes_raw = b""
            except Exception as e:
                print(f"Failed to download image at row {current_row}: {e}")
                img_bytes_raw = b""

            # --- Call Abacus API if we have image bytes ---
            ai_short = ai_standard = ai_long = ai_notes = ""
            image_topic_ai = ""  # <-- NEW

            if img_bytes_raw:
                prompt = build_prompt(
                    grade_band=grade_value or grade_band,
                    course_topic=course_value or course_topic,
                    context=image_context or section or webpage,
                )
                try:
                    ai_json = call_abacus(img_bytes_raw, prompt)
                except Exception as e:
                    print(f"⚠️ Abacus API error for row {current_row}: {e}")
                    ai_json = {
                        "short": "",
                        "standard": "",
                        "long": "",
                        "notes": "",
                        "topic": "",
                    }

                ai_short = utils.clean_text(ai_json.get("short", ""), 250)
                ai_standard = utils.clean_text(ai_json.get("standard", ""), 600)
                ai_long = utils.clean_text(ai_json.get("long", ""), 2000)
                ai_notes = utils.clean_text(ai_json.get("notes", ""), 1000)
                image_topic_ai = utils.clean_text(ai_json.get("topic", ""), 80)

            else:
                image_topic_ai = ""

            # Column F (6): Image Topic — prefer AI topic (1–2 words), fallback to HTML title
            topic_value = image_topic_ai or image_title
            ws.cell(row=current_row, column=6, value=topic_value)

            # Columns S–V (19–22): A.I. alt text variants
            ws.cell(row=current_row, column=19, value=ai_short)        # S: A.I. Short Alt Text
            ws.cell(row=current_row, column=20, value=ai_standard)     # T: A.I. Med Alt Text
            ws.cell(row=current_row, column=21, value=ai_long)         # U: A.I. Long Alt Text
            ws.cell(row=current_row, column=22, value=ai_notes)        # V: A.I. Notes

            # Column W (23): New/Suggested Alt Text — blank (reviewer fills in)

            # Column X (24): New/Suggested Alt Text Count — formula =LEN(W)
            ws.cell(row=current_row, column=24, value=f"=LEN(W{current_row})")

            # Columns Y–AA (25–27): Notes / Status / Reviewer Notes — blank
            # Column AB (28): In Word? — blank for website sources
            # Columns AC–AD (29–30): In Excel? / Date of Completion — blank


            # --- A: Thumbnail + E: Image Size ---
            try:
                if img_bytes_raw:
                    thumb_bytes, thumb_w, thumb_h, orig_w, orig_h = utils.create_thumbnail_from_bytes(
                        img_bytes_raw,
                        max_width_px=600,
                        allow_upscale=False,
                    )

                    if orig_w and orig_h:
                        img_size_str = f"{orig_w} x {orig_h} px"
                        ws.cell(row=current_row, column=5, value=img_size_str)

                    if thumb_bytes:
                        utils.insert_png_bytes_in_cell(
                            ws,
                            row=current_row,
                            col=1,  # A
                            png_bytes=thumb_bytes,
                            width_px=thumb_w,
                            height_px=thumb_h,
                            adjust_row_col=True,
                            max_col_width_px=600,
                        )
                    else:
                        ws.row_dimensions[current_row].height = 20
                else:
                    ws.row_dimensions[current_row].height = 20
            except Exception as e:
                print(f"Failed to insert thumbnail at row {current_row}: {e}")
                ws.row_dimensions[current_row].height = 20

            current_row += 1


        # Wrap text and autofit B–AD (text columns), cap at ~240px
        text_columns = [
            "B", "C", "D", "E", "F", "G",
            "H", "I", "J", "K", "L", "M",
            "N", "O", "P", "Q", "R", "S",
            "T", "U", "V", "W", "X", "Y",
            "Z", "AA", "AB", "AC", "AD",
        ]
        utils.wrap_and_autofit_columns(
            ws,
            text_columns=text_columns,
            max_col_width_px=240,
        )

        # Optional: crude row height adjust so wrapped text is visible
        for row_idx in range(2, ws.max_row + 1):
            max_lines = 1
            for col_letter in text_columns:
                cell = ws[f"{col_letter}{row_idx}"]
                if cell.value:
                    text = str(cell.value)
                    est_lines = max(1, len(text) // 60 + 1)
                    if est_lines > max_lines:
                        max_lines = est_lines
            # 18–20 pts per line is usually comfortable
            ws.row_dimensions[row_idx].height = max_lines * 18

        wb.save(output_filename)
        print(f"Excel file with embedded images & AI alt text saved as: {output_filename}")


def main():
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    extractor = CombinedImageExtractor()
    all_images_data: List[Dict[str, str]] = []

    for url, topic in URLS_AND_TOPICS:
        if not url:
            continue
        images_data = extractor.extract_images_from_url(url, topic)
        all_images_data.extend(images_data)

    if not all_images_data:
        print("No images found to process.")
        return

    date_str = datetime.now().strftime("%Y%m%d")
    output_filename = OUTPUT_FOLDER / f"{BOOK_SOURCE_NAME}_website_AI_{date_str}.xlsx"

    print(f"\nCreating Excel file with {len(all_images_data)} total images...")
    extractor.create_excel_with_images(
        all_images_data,
        output_filename,
        file_source=BOOK_SOURCE_NAME,
        grade_band=DEFAULT_GRADE_BAND,
        course_topic=COURSE_TOPIC,
    )
    print("Done!")


if __name__ == "__main__":
    main()
