# -*- coding: utf-8 -*-
"""
Invoice File Organizer — Step 3: Dual Structure
Author: [Your Name or GitHub Handle]
Date: 2026

Description:
    Step 3 of the invoice processing pipeline.
    Takes the renamed PDFs produced by Step 2 and organises them into two
    parallel directory structures for different access patterns:

    Structure A — browsing by supplier:
        organized/by_supplier/{location_folder}/{supplier}/{date}/{invoice}.pdf

    Structure B — browsing by date:
        organized/by_date/{location_folder}/{date}/{supplier}_{invoice}.pdf

    The file is moved into Structure A (primary archive) and copied into
    Structure B (secondary index). The source date folder is deleted if it
    is empty after processing.

    Location folder names are resolved by scanning the existing by_supplier
    tree for a folder whose name contains the 4-digit location code, allowing
    descriptive names like "0042 - Central Warehouse" without hardcoding them.

Usage:
    Ensure a .env file exists (see .env.example), then run:
        python step3_organize.py
"""

import os
import re
import shutil
from datetime import datetime, timedelta

from dotenv import load_dotenv

load_dotenv()


# =============================================================================
# 1. CONFIGURATION  (all paths from .env)
# =============================================================================

BASE_DIR          = os.getenv("INVOICE_BASE_FOLDER",      os.path.join(os.getcwd(), "invoices"))
BY_SUPPLIER_ROOT  = os.getenv("ORGANIZED_BY_SUPPLIER",    os.path.join(BASE_DIR, "organized", "by_supplier"))
BY_DATE_ROOT      = os.getenv("ORGANIZED_BY_DATE",        os.path.join(BASE_DIR, "organized", "by_date"))

# Date folder — defaults to yesterday
yesterday    = datetime.now() - timedelta(days=1)
DATE_FOLDER  = yesterday.strftime("%d-%m-%Y")
SOURCE_FOLDER = os.path.join(BASE_DIR, DATE_FOLDER)

print(f"📁 Source folder : {SOURCE_FOLDER}")
print(f"📂 By-supplier   : {BY_SUPPLIER_ROOT}")
print(f"📂 By-date       : {BY_DATE_ROOT}")


# =============================================================================
# 2. HELPERS
# =============================================================================

def build_location_map(root: str) -> dict[str, str]:
    """
    Scans the by_supplier root for existing location folders and returns a
    dict mapping 4-digit location code → full folder name.

    Example:
        {"0042": "0042 - Central Warehouse", "0017": "0017 - North Branch"}

    Only processes subdirectories whose name contains exactly one 4-digit
    sequence so that other folders (e.g. archive years) are ignored.
    """
    location_map: dict[str, str] = {}

    if not os.path.exists(root):
        os.makedirs(root, exist_ok=True)
        return location_map

    for folder_name in os.listdir(root):
        if os.path.isdir(os.path.join(root, folder_name)):
            match = re.search(r"\b(\d{4})\b", folder_name)
            if match:
                location_map[match.group(1)] = folder_name

    return location_map


def safe_destination(directory: str, filename: str) -> str:
    """
    Returns a unique file path inside directory by appending _(N) suffixes
    if a file with the same name already exists.
    """
    dest = os.path.join(directory, filename)
    if not os.path.exists(dest):
        return dest

    stem, ext = os.path.splitext(filename)
    counter   = 1
    while os.path.exists(dest):
        dest = os.path.join(directory, f"{stem}_({counter}){ext}")
        counter += 1
    return dest


# =============================================================================
# 3. MAIN ORGANISER
# =============================================================================

def organize_and_cleanup() -> None:
    """
    Iterates over all renamed PDFs in SOURCE_FOLDER.

    Expected filename format (produced by step2_rename.py):
        DD-MM-YYYY_LOCATIONCODE_SUPPLIER_INVOICENUMBER.pdf

    For each valid file:
      1. Parses the four fields from the filename.
      2. Resolves the location folder name via the existing by_supplier tree.
         Files whose location code has no matching folder are skipped (logged).
      3. Copies the file to Structure B (by date).
      4. Moves the file to Structure A (by supplier).

    After processing, removes the source folder if it is empty.
    """
    print("🚀 Starting dual-structure file organisation...\n")

    location_map = build_location_map(BY_SUPPLIER_ROOT)
    moved_count  = 0

    if not os.path.exists(SOURCE_FOLDER):
        print(f"⚠️  Source folder not found: {SOURCE_FOLDER}")
        return

    for filename in os.listdir(SOURCE_FOLDER):
        if not filename.lower().endswith(".pdf"):
            continue

        stem  = filename[:-4]
        parts = stem.split("_")

        # Filename must have at least 4 underscore-delimited parts
        if len(parts) < 4:
            print(f"⏩ Skipping '{filename}' — does not match expected format.")
            continue

        date_str      = parts[0]
        location_code = parts[1]
        supplier      = parts[2]
        invoice_num   = "_".join(parts[3:])   # invoice number may contain underscores

        # Resolve the human-readable location folder name
        location_folder = location_map.get(location_code)
        if not location_folder:
            print(
                f"⏩ Skipping '{filename}' — "
                f"location code '{location_code}' has no matching folder in by_supplier/."
            )
            continue

        source_path = os.path.join(SOURCE_FOLDER, filename)

        # ── Structure A: by_supplier / location / supplier / date / invoice.pdf ──
        dir_A  = os.path.join(BY_SUPPLIER_ROOT, location_folder, supplier, date_str)
        os.makedirs(dir_A, exist_ok=True)
        dest_A = safe_destination(dir_A, f"{invoice_num}.pdf")

        # ── Structure B: by_date / location / date / supplier_invoice.pdf ──
        dir_B  = os.path.join(BY_DATE_ROOT, location_folder, date_str)
        os.makedirs(dir_B, exist_ok=True)
        dest_B = safe_destination(dir_B, f"{supplier}_{invoice_num}.pdf")

        try:
            shutil.copy2(source_path, dest_B)   # copy to Structure B first
            shutil.move(source_path,  dest_A)   # then move to Structure A
            print(f"✅ Organised: {filename}")
            moved_count += 1
        except Exception as exc:
            print(f"❌ Error processing '{filename}': {exc}")

    # ── Cleanup: remove source folder if empty ────────────────────────────────
    print("\n🧹 Checking source folder for cleanup...")
    try:
        remaining = os.listdir(SOURCE_FOLDER)
        if not remaining:
            os.rmdir(SOURCE_FOLDER)
            print(f"🗑️  Source folder was empty and has been removed: {SOURCE_FOLDER}")
        else:
            print(
                f"⚠️  Source folder NOT removed — "
                f"{len(remaining)} file(s) remain (check for unmatched location codes)."
            )
    except Exception as exc:
        print(f"❌ Cleanup error: {exc}")

    print(f"\n🏁 ORGANISATION COMPLETE — {moved_count} file(s) placed into both structures.")


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    organize_and_cleanup()
