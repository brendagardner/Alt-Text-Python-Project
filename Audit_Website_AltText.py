#This Script will take URLs and pull the current images and alt text from the webpage and add them to an Excel file..
#created by Brenda Gardner, September 2025.
# Updated Dec 1 2025 to use alttext_utils helpers and reorganize columns and add text wrapping, skip logo images, scrape website for context.

# Step 1. Capture URLs and Subjects.
# a. Go to Website TOC and copy paste URLS.
# use in conjunction with https://slccbruins-my.sharepoint.com/:x:/g/personal/bgardn61_slcc_edu/IQD6AhzRzSxNRohDxUOKgeNYAQrWIctBAHw6ag8_bIYE-U8?e=sQFPWW

# End Step. After processing: Use macro to put the images in cell to allow filtering


import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from io import BytesIO
from pathlib import Path
from datetime import datetime
import warnings
from urllib3.exceptions import NotOpenSSLWarning
warnings.simplefilter("ignore", NotOpenSSLWarning)

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from helpers import alttext_utils as utils

# =========================
# CONFIG — EDIT THESE
# =========================
# These will be written into every row (can be left "" if you want to fill later).
DEFAULT_GRADE_BAND = "12-24"       # e.g. "10–12"
COURSE_TOPIC = "Graph Theory"      # e.g. "College Algebra"
BOOK_SOURCE_NAME = "Math in Society"

# Folder where the Excel output file will be written
OUTPUT_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1050 Information"
)

# List of (URL, Section/Topic) pairs
URLS_AND_TOPICS = [
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.01:_Introduction" , "6.1: Introduction"),
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.02:_Graphs" , "6.2: Graphs"),
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.03:_Shortest_Path" , "6.3: Shortest Path"),
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.04:_Euler_Circuits_and_the_Chinese_Postman_Problem" , "6.4: Euler Circuits and the Chinese Postman Problem"),
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.05:_Eulerization_and_the_Chinese_Postman_Problem" , "6.5: Eulerization and the Chinese Postman Problem"),
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.06:_Hamiltonian_Circuits_and_the_Traveling_Salesman_Problem" , "6.6: Hamiltonian Circuits and the Traveling Salesman Problem"),
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.07:_Spanning_Trees" , "6.7: Spanning Trees"),
("https://math.libretexts.org/Bookshelves/Applied_Mathematics/Math_in_Society_(Lippman)/06:_Graph_Theory/6.07:_Spanning_Trees" , "6.8: Exercise"),
]


class CombinedImageExtractor:
    def __init__(self, proxies=None):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': (
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                'AppleWebKit/537.36 (KHTML, like Gecko) '
                'Chrome/91.0.4472.124 Safari/537.36'
            )
        })
        if proxies:
            self.session.proxies.update(proxies)

    def extract_images_from_url(self, url, topic=""):
        """Extract all images and alt text from a webpage and its iframes."""
        try:
            if not url:
                return []

            print(f"Processing main page: {url}")
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')

            images_data = []

            # Extract images from main page
            main_images = self._extract_images_from_soup(
                soup, url, topic, source_type="main page"
            )
            images_data.extend(main_images)

            # Find and process all iframes
            iframes = soup.find_all('iframe')
            print(f"Found {len(iframes)} iframes to process")

            for i, iframe in enumerate(iframes):
                iframe_src = iframe.get('src')
                if not iframe_src:
                    continue

                # Convert relative URLs to absolute
                iframe_url = urljoin(url, iframe_src)

                # Skip certain iframe types that won't have useful images
                skip_domains = [
                    'youtube.com', 'vimeo.com', 'facebook.com',
                    'twitter.com', 'ads.', 'analytics.', 'google.com/maps'
                ]
                if any(domain in iframe_url.lower() for domain in skip_domains):
                    print(f"Skipping iframe {i + 1}: {iframe_url} "
                          f"(likely ads/social/video)")
                    continue

                print(f"Processing iframe {i + 1}: {iframe_url}")
                iframe_images = self._extract_images_from_iframe(
                    iframe_url, topic, original_webpage=url
                )
                images_data.extend(iframe_images)

            print(f"Total found: {len(images_data)} images")
            return images_data

        except Exception as e:
            print(f"Error extracting from {url}: {str(e)}")
            return []

    def _extract_images_from_soup(
            self,
            soup,
            source_url,
            topic,
            source_type="page",
            original_webpage=None
    ):
        """Helper method to extract images from a BeautifulSoup object."""
        images_data = []
        img_tags = soup.find_all('img')

        for img in img_tags:
            src = img.get('src')
            if not src:
                continue

            # --------------------------------------------------------
            # 1️⃣ Skip images inside <header> or <footer>
            # --------------------------------------------------------
            if img.find_parent(['header', 'footer']):
                continue

            # --------------------------------------------------------
            # 2️⃣ Skip by parent class/id (common footer/header patterns)
            # --------------------------------------------------------
            parent = img.parent
            if parent:
                skip_classes = [
                    "site-header", "header", "footer", "site-footer",
                    "footer-logo", "masthead", "page-header"
                ]
                skip_ids = ["header", "footer", "colophon", "masthead"]

                # skip if parent contains a class we are avoiding
                if any(cls in skip_classes for cls in parent.get("class", [])):
                    continue

                # skip if parent has an ID we are avoiding
                if parent.get("id") in skip_ids:
                    continue

            # --------------------------------------------------------
            # 3️⃣ Skip by URL filename patterns (logos, banners, icons)
            # --------------------------------------------------------
            # Convert relative URLs to absolute NOW so we can inspect filename
            full_url = urljoin(source_url, src)
            full_url_lower = full_url.lower()

            skip_patterns = [
                "logo", "banner", "footer", "header",
                "favicon", "sprite", "icon"
            ]

            if any(p in full_url_lower for p in skip_patterns):
                continue

            # --------------------------------------------------------
            # SAFE — process the content image
            # --------------------------------------------------------

            alt_text = img.get('alt', '') or ""
            title_text = img.get('title', '') or ""
            author = img.get('data-author') or img.get('data-credit') or ''
            copyright_text = img.get('data-copyright') or img.get('copyright') or ''

            # grab context around the image
            image_context = self._extract_image_context(img)

            images_data.append({
                'image_location': full_url,
                'alt_text': alt_text,
                'webpage': original_webpage if original_webpage else source_url,
                'section': topic,
                'image_title': title_text,
                'author': author,
                'copyright': copyright_text,
                'source_type': source_type,
                'image_context': image_context,  # 🔹 NEW FIELD
            })

        print(f"Found {len(images_data)} images in {source_type}")
        return images_data

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

        # Helper to scan previous siblings of a node
        def prev_sibling_text(node):
            for sib in getattr(node, "previous_siblings", []):
                if getattr(sib, "name", None) in ("p", "li", "div"):
                    txt = clean_text(sib)
                    if txt:
                        return txt
            return ""

        # 2) Previous siblings of the <img> itself
        ctx = prev_sibling_text(img)
        if ctx:
            return ctx

        # 3) Walk up ancestors and check THEIR previous siblings
        parent = img.parent
        while parent is not None and getattr(parent, "name", None) not in ("body", "html"):
            ctx = prev_sibling_text(parent)
            if ctx:
                return ctx
            parent = parent.parent

        # 4) Fallback: next siblings after the <img>
        for sib in getattr(img, "next_siblings", []):
            if getattr(sib, "name", None) in ("p", "li", "div"):
                txt = clean_text(sib)
                if txt:
                    return txt

        return ""

    def _extract_images_from_iframe(self, iframe_url, topic, original_webpage):
        """Extract images from an iframe."""
        try:
            response = self.session.get(iframe_url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')

            return self._extract_images_from_soup(
                soup, iframe_url, topic, source_type=f"iframe ({iframe_url})",
                original_webpage=original_webpage
            )

        except Exception as e:
            print(f"Error processing iframe {iframe_url}: {str(e)}")
            return []

    def _safe_filename(self, name: str) -> str:
        return "".join(
            c if c.isalnum() or c in (' ', '-', '_') else '_'
            for c in name
        ).strip().replace(' ', '_')

    def _extract_book_name_from_url(self, url: str) -> str:
        """Extract first path segment as the book name, e.g., slcc-elementaryalgebra."""
        try:
            parsed = urlparse(url)
            first_segment = (parsed.path or '/').strip('/').split('/')[0] or 'book'
            return self._safe_filename(first_segment)
        except Exception:
            return "book"

    # ------------------------------------------------------------------
    # NEW: Excel builder using alttext_utils + new column layout
    # ------------------------------------------------------------------
    def create_excel_with_images(
        self,
        all_images_data,
        output_filename: str,
        file_source: str = "",
        grade_band: str = "",
        course_topic: str = ""
    ):
        """
        Create a single Excel file with embedded thumbnails and the
        NEW column layout:

        A: Image Thumbnail
        B: Image Filename
        C: Image Link
        D: File Name/Source (Hard coded in configurations)
        E: Image Size
        F: Image Topic (Blank for now)
        G: Image Context
        H: Grade Level (Hard coded)
        I: Course Topic (Hard coded)
        J: Webpage/Source
        K: Section
        L: Copyright
        M: Author
        N: Current Alt Text
        O: Current Alt Text Character Count
        """
        if not all_images_data:
            print("No images to export.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Images"

        # ---- Headers ----
        headers = [
            "Image Thumbnail",                    # A
            "Image Filename",                     # B
            "Image Link",                         # C
            "Image Hosting",                   # D
            "Image Size",                         # E
            "Image Topic",                        # F
            "Image Context",                      # G
            "Grade Level",                        # H
            "Course Topic",                       # I
            "Webpage/Source",                     # J
            "Section",                            # K
            "Copyright",                          # L
            "Author",                             # M
            "Current Alt Text",                   # N
            "Current Alt Text Character Count",   # O
        ]

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)

        # Set some base column widths (thumbnail column will be adjusted by helper)
        ws.column_dimensions['A'].width = utils._pixels_to_column_width(600)
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 40
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 25
        ws.column_dimensions['G'].width = 25
        ws.column_dimensions['H'].width = 15
        ws.column_dimensions['I'].width = 25
        ws.column_dimensions['J'].width = 40
        ws.column_dimensions['K'].width = 25
        ws.column_dimensions['L'].width = 20
        ws.column_dimensions['M'].width = 20
        ws.column_dimensions['N'].width = 40
        ws.column_dimensions['O'].width = 15

        # ---- Data rows ----
        current_row = 2

        for img_info in all_images_data:
            image_url = img_info.get("image_location", "") or ""
            alt_text = img_info.get("alt_text", "") or ""
            webpage = img_info.get("webpage", "") or ""
            section = img_info.get("section", "") or ""
            image_title = img_info.get("image_title", "") or ""
            author = img_info.get("author", "") or ""
            copyright_text = img_info.get("copyright", "") or ""
            image_context = img_info.get("image_context", "") or ""  # 🔹 NEW

            # B: image filename from URL path
            parsed = urlparse(image_url)
            image_filename = Path(parsed.path).name if parsed.path else ""

            # H/I: grade level & course topic pulled from config
            grade_value = grade_band or ""
            course_value = course_topic or ""

            # N/O: alt text and character count
            current_alt = alt_text
            alt_char_count = len(current_alt) if current_alt else 0

            # J: webpage/source (hyperlink)
            ws.cell(row=current_row, column=10, value=webpage)
            if webpage:
                ws.cell(row=current_row, column=10).hyperlink = webpage
                ws.cell(row=current_row, column=10).style = "Hyperlink"

            # C: image link (hyperlink)
            ws.cell(row=current_row, column=3, value=image_url)
            if image_url:
                ws.cell(row=current_row, column=3).hyperlink = image_url
                ws.cell(row=current_row, column=3).style = "Hyperlink"

            # B: filename
            ws.cell(row=current_row, column=2, value=image_filename)

            # 🔹 D: File Name/Source as an Excel formula based on C (image link)
            formula_str = (
                f'=IF(LEFT(C{current_row},10)="https://s3","Amazon AWS",'
                f'IF(LEFT(C{current_row},19)="https://textimgs.s3","AMS MIT","Other"))'
            )
            ws.cell(row=current_row, column=4, value=formula_str)

            # F: Image Topic (using HTML title as a placeholder)
            ws.cell(row=current_row, column=6, value=image_title)

            # 🔹 G: Image Context (paragraph before the image)
            ws.cell(row=current_row, column=7, value=image_context)

            # H: Grade Level
            ws.cell(row=current_row, column=8, value=grade_value)

            # I: Course Topic
            ws.cell(row=current_row, column=9, value=course_value)

            # K: Section (the chapter/section label you passed in)
            ws.cell(row=current_row, column=11, value=section)

            # L: Copyright
            ws.cell(row=current_row, column=12, value=copyright_text)

            # M: Author
            ws.cell(row=current_row, column=13, value=author)

            # N: Current Alt Text
            ws.cell(row=current_row, column=14, value=current_alt)

            # O: Current Alt Text Character Count
            ws.cell(row=current_row, column=15, value=alt_char_count)

            # --- A: Thumbnail + E: Image Size ---
            try:
                if image_url:
                    print(f"Downloading image for row {current_row}: {image_url}")
                    resp = self.session.get(image_url, timeout=10)
                    resp.raise_for_status()

                    # Use shared thumbnail helper
                    thumb_bytes, thumb_w, thumb_h, orig_w, orig_h = utils.create_thumbnail_from_bytes(
                        resp.content,
                        max_width_px=600,
                        allow_upscale=False,
                    )

                    # E: Image Size (original dimensions)
                    if orig_w and orig_h:
                        img_size_str = f"{orig_w} x {orig_h} px"
                        ws.cell(row=current_row, column=5, value=img_size_str)

                    # Insert thumbnail into Column A
                    if thumb_bytes:
                        utils.insert_png_bytes_in_cell(
                            ws,
                            row=current_row,
                            col=1,  # Column A
                            png_bytes=thumb_bytes,
                            width_px=thumb_w,
                            height_px=thumb_h,
                            adjust_row_col=True,
                            max_col_width_px=600,
                        )
                    else:
                        ws.row_dimensions[current_row].height = 20
                else:
                    # no URL, no thumbnail
                    ws.row_dimensions[current_row].height = 20

            except Exception as e:
                print(f"Failed to insert thumbnail at row {current_row}: {e}")
                ws.row_dimensions[current_row].height = 20

            current_row += 1


        # Wrap text and autofit B–O (all text columns), cap at ~240px
        text_columns = [
            "B", "C", "D", "E", "F", "G",
            "H", "I", "J", "K", "L", "M",
            "N", "O",
        ]
        utils.wrap_and_autofit_columns(
            ws,
            text_columns=text_columns,
            max_col_width_px=240,
        )

        # AUTO-ADJUST ROW HEIGHTS
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            max_lines = 1
            for cell in row:
                if cell.value is not None:
                    text = str(cell.value)
                    # Rough estimate: every ~60 characters creates a new wrapped line
                    est_lines = max(1, len(text) // 60 + 1)
                    if est_lines > max_lines:
                        max_lines = est_lines

            # Convert estimated lines into row height
            ws.row_dimensions[cell.row].height = max_lines * 30

        wb.save(output_filename)
        print(f"Excel file with embedded images saved as: {output_filename}")


def main():
    extractor = CombinedImageExtractor()

    all_images_data = []

    for url, topic in URLS_AND_TOPICS:
        if not url:
            continue
        images_data = extractor.extract_images_from_url(url, topic)
        all_images_data.extend(images_data)

    if not all_images_data:
        print("No images found to process.")
        return

    date_str = datetime.now().strftime("%Y%m%d")
    output_filename = OUTPUT_FOLDER / f"{BOOK_SOURCE_NAME}_website_{date_str}.xlsx"

    print(f"\nCreating Excel file with {len(all_images_data)} total images...")
    extractor.create_excel_with_images(
        all_images_data,
        output_filename,
        file_source=BOOK_SOURCE_NAME,
        grade_band=DEFAULT_GRADE_BAND,
        course_topic=COURSE_TOPIC
    )

    print(f"Done! Excel saved to:\n{output_filename}")


if __name__ == "__main__":
    main()
