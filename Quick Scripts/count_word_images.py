#!/usr/bin/env python3
import zipfile
from pathlib import Path
from collections import Counter

# --- CONFIGURATION ---
# 1. Replace the path below with your actual folder path.
# TIP: You can drag your folder from Finder into this window to see the path.
TARGET_FOLDER = Path("/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1060 Information/OER Trigonometry Textbook 2023")

# 2. Set to True if you want to search inside subfolders
RECURSIVE = True
# ---------------------

TARGET_EXTS = {".png", ".jpeg", ".jpg", ".emf"}


def count_images_in_docx(docx_path):
    counts = Counter()
    try:
        with zipfile.ZipFile(docx_path, "r") as z:
            for name in z.namelist():
                if name.lower().startswith("word/media/"):
                    ext = Path(name).suffix.lower()
                    if ext in TARGET_EXTS:
                        # Group .jpg and .jpeg together
                        if ext == ".jpg": ext = ".jpeg"
                        counts[ext] += 1
    except zipfile.BadZipFile:
        pass  # Skip corrupted files
    return counts


def main():
    if not TARGET_FOLDER.exists():
        print(f"Error: The folder '{TARGET_FOLDER}' does not exist.")
        return

    pattern = "**/*.docx" if RECURSIVE else "*.docx"
    docx_files = list(TARGET_FOLDER.glob(pattern))

    total_counts = Counter()
    for docx in docx_files:
        total_counts.update(count_images_in_docx(docx))

    print(f"--- Results for: {TARGET_FOLDER} ---")
    print(f"Word documents scanned: {len(docx_files)}")
    print(f"Total .png files:  {total_counts['.png']}")
    print(f"Total .jpeg files: {total_counts['.jpeg']}")
    print(f"Total .emf files:  {total_counts['.emf']}")


if __name__ == "__main__":
    main()