# This script takes an existing Excel file whose columns A–O are already filled (same layout as your website alt-text/export sheet), and then:
# - Uses Column C (Image Link) to re-download images
# - Calls Abacus AI to generate alt text if Column N (Current Alt Text) is empty - Change on Row 155
# - Fills Columns P–S:
#     P: AI Short Alt Text
#     Q: AI Standard Alt Text
#     R: AI Long Alt Text
#     S: AI Notes
# The output is a *copy* of the original .xlsx with "_AI" appended.
# Can reuse since it SKIPS rows with AI alt text already filled.

import os
import json
import base64
import warnings
from pathlib import Path
from typing import Dict

import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from urllib3.exceptions import NotOpenSSLWarning

from helpers import alttext_utils as utils

# Suppress LibreSSL / NotOpenSSL warnings on macOS
warnings.simplefilter("ignore", NotOpenSSLWarning)

# =========================
# CONFIG — EDIT THESE
# =========================

# Excel file that already has columns A–O filled
INPUT_EXCEL_PATH = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/103x Information/Test_Excel_AI_Euler_AI.xlsx"
)

# Optional: sheet name; if "", use active sheet
SHEET_NAME = "v2"  # e.g. "Website Alt Text AI Output"

# Grade band & course topic are only used for the AI prompt context
DEFAULT_GRADE_BAND = "6-8"
COURSE_TOPIC = "Elementary Algebra"

# Abacus AI config
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
    import requests  # keep local if you prefer

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
# Main processing
# =========================

def main():
    if not INPUT_EXCEL_PATH.exists():
        raise FileNotFoundError(f"❌ Input Excel not found: {INPUT_EXCEL_PATH}")

    print(f"📂 Loading workbook: {INPUT_EXCEL_PATH}")
    wb = load_workbook(INPUT_EXCEL_PATH)
    if SHEET_NAME and SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
    else:
        ws = wb.active

    print(f"Using sheet: {ws.title}")

    # Make sure Column A is wide enough for thumbnails (~600 px)
    desired_thumb_width = utils._pixels_to_column_width(600)
    col_a_dim = ws.column_dimensions["A"]
    if not col_a_dim.width or col_a_dim.width < desired_thumb_width:
        col_a_dim.width = desired_thumb_width

    # Ensure headers P–S exist (row 1), but don't overwrite if already present
    header_map = {
        16: "AI Short Alt Text",    # P
        17: "AI Standard Alt Text", # Q
        18: "AI Long Alt Text",     # R
        19: "AI Notes",             # S
    }
    for col_idx, header in header_map.items():
        cell = ws.cell(row=1, column=col_idx)
        if not cell.value:
            cell.value = header

    # Loop over rows 2..max_row
    for row_idx in range(2, ws.max_row + 1):
        image_link_cell = ws.cell(row=row_idx, column=3)   # C: Image Link
        
        #Update below if changing column to review and skip AI
        current_alt_cell = ws.cell(row=row_idx, column=20)  # N: Current Alt Text Change column = 14, # T: Suggested Alt Text column = 20
        if current_alt_cell.value not in (None, ""):
            continue

        image_url = (image_link_cell.value or "").strip()

        # Skip rows with no image link
        if not image_url:
            continue

        # Pull some context from existing columns:
        # G: Image Context, K: Section, N: Current Alt Text, J: Webpage/Source
        image_context = (ws.cell(row=row_idx, column=7).value or "").strip()   # G
        section = (ws.cell(row=row_idx, column=11).value or "").strip()        # K
        current_alt = (ws.cell(row=row_idx, column=14).value or "").strip()    # N
        webpage = (ws.cell(row=row_idx, column=10).value or "").strip()        # J

        context_parts = [image_context, section, current_alt, webpage]
        context = " | ".join([c for c in context_parts if c])

        grade_band = DEFAULT_GRADE_BAND
        course_topic = COURSE_TOPIC

        print(f"Row {row_idx}: downloading image {image_url!r}")
        try:
            resp = requests.get(image_url, timeout=10)
            resp.raise_for_status()
            img_bytes_raw = resp.content
        except Exception as e:
            print(f"  ⚠️ Failed to download image at row {row_idx}: {e}")
            continue

        # --- Call Abacus ---
        prompt = build_prompt(
            grade_band=grade_band,
            course_topic=course_topic,
            context=context,
        )

        try:
            ai_json = call_abacus(img_bytes_raw, prompt)
        except Exception as e:
            print(f"  ⚠️ Abacus API error at row {row_idx}: {e}")
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

        # --- P–S: write AI text ---
        ws.cell(row=row_idx, column=16, value=ai_short)    # P
        ws.cell(row=row_idx, column=17, value=ai_standard) # Q
        ws.cell(row=row_idx, column=18, value=ai_long)     # R
        ws.cell(row=row_idx, column=19, value=ai_notes)    # S

        # Make the AI cells wrap & top-align
        for col_idx in range(16, 20):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = Alignment(
                wrap_text=True,
                vertical="top"
            )

        # --- A & E: thumbnail + image size (NEW) ---
        try:
            thumb_bytes, thumb_w, thumb_h, orig_w, orig_h = utils.create_thumbnail_from_bytes(
                img_bytes_raw,
                max_width_px=600,
                allow_upscale=False,
            )

            # E: Image Size (if that column exists / is part of your layout)
            if orig_w and orig_h:
                size_str = f"{orig_w} x {orig_h} px"
                ws.cell(row=row_idx, column=5, value=size_str)  # E

            if thumb_bytes:
                utils.insert_png_bytes_in_cell(
                    ws,
                    row=row_idx,
                    col=1,  # A
                    png_bytes=thumb_bytes,
                    width_px=thumb_w,
                    height_px=thumb_h,
                    adjust_row_col=True,
                    max_col_width_px=600,
                )
        except Exception as e:
            print(f"  ⚠️ Failed to insert thumbnail at row {row_idx}: {e}")
            # leave row height as-is

    # Save as copy with "_AI" appended
    out_path = INPUT_EXCEL_PATH.with_name(
        INPUT_EXCEL_PATH.stem + "_AI" + INPUT_EXCEL_PATH.suffix
    )
    wb.save(out_path)
    print(f"\n✅ Completed — Excel with AI alt text + thumbnails saved to:\n{out_path}")


if __name__ == "__main__":
    main()
