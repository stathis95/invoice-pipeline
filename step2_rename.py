# -*- coding: utf-8 -*-
"""
Invoice PDF Processor — Step 2: Extract & Rename
Author: [Your Name or GitHub Handle]
Date: 2026

Description:
    Step 2 of the invoice processing pipeline.
    Scans a folder of downloaded PDF invoices, extracts structured fields
    (supplier VAT, invoice number, date, location code) using text extraction
    and an OCR fallback, then renames each file to a standardised format:

        DD-MM-YYYY_LOCATIONCODE_SUPPLIERNAME_INVOICENUMBER.pdf

    Supplier identification uses a lookup Excel table keyed by VAT number.
    Location attribution uses a keyword-scoring system loaded from a second
    Excel table (postal codes carry higher weight than generic keywords).

Usage:
    Ensure a .env file exists (see .env.example), then run:
        python step2_rename.py
"""

import os
import re
import gc
import shutil
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import pdfplumber
import fitz
import easyocr
from dotenv import load_dotenv

load_dotenv()


# =============================================================================
# 1. CONFIGURATION  (all paths from .env)
# =============================================================================

BASE_DIR       = os.getenv("INVOICE_BASE_FOLDER",  os.path.join(os.getcwd(), "invoices"))
SUPPLIERS_FILE = os.getenv("SUPPLIERS_FILE",        os.path.join(BASE_DIR, "suppliers.xlsx"))
LOCATIONS_FILE = os.getenv("LOCATIONS_FILE",        os.path.join(BASE_DIR, "locations.xlsx"))
TRASH_FOLDER   = os.getenv("TRASH_FOLDER",          os.path.join(BASE_DIR, "UNPROCESSABLE"))
RULES_FILE     = os.path.join(BASE_DIR, "invoice_rules.txt")

# Supplier Excel column names — update to match your spreadsheet headers
SUPPLIER_VAT_COL  = os.getenv("SUPPLIER_VAT_COL",  "VAT_Number")
SUPPLIER_NAME_COL = os.getenv("SUPPLIER_NAME_COL",  "Supplier_Name")

# Location Excel column names
LOC_CODE_COL = os.getenv("LOC_CODE_COL", "LOCATION_CODE")
LOC_ADDR_COL = os.getenv("LOC_ADDR_COL", "ADDRESS")
LOC_KEYS_COL = os.getenv("LOC_KEYS_COL", "KEYWORDS")

# Date folder — defaults to yesterday
yesterday   = datetime.now() - timedelta(days=1)
DATE_FOLDER = yesterday.strftime("%d-%m-%Y")
PDF_FOLDER  = os.path.join(BASE_DIR, DATE_FOLDER)

print(f"📁 Processing folder: {PDF_FOLDER}")


# =============================================================================
# 2. REGEX RULES  (loaded from external file for easy customisation)
# =============================================================================

def load_invoice_rules() -> list[str]:
    """
    Reads invoice-number regex patterns from invoice_rules.txt.
    Each non-blank, non-comment line is treated as one pattern.
    Capture group 1 must contain the invoice number.

    Falls back to a generic pattern if the file is absent.
    """
    rules: list[str] = []

    if os.path.exists(RULES_FILE):
        with open(RULES_FILE, "r", encoding="utf-8") as fh:
            for line in fh:
                clean = line.strip()
                if clean and not clean.startswith("#"):
                    rules.append(clean)
        print(f"✅ Loaded {len(rules)} invoice-number rule(s).")
    else:
        print("⚠️  invoice_rules.txt not found — using built-in fallback rule.")
        rules = [r"INVOICE\s*NO\.?\s*:?\s*([A-Z0-9/_-]+)"]

    return rules


INVOICE_PATTERNS = load_invoice_rules()


# =============================================================================
# 3. LOOKUP LOADERS
# =============================================================================

def load_suppliers() -> dict[str, str]:
    """
    Reads the supplier lookup table from an Excel file.
    Returns a dict mapping VAT number (zero-padded to 9 digits) → supplier name.
    """
    print("📊 Loading supplier list...")
    df = pd.read_excel(SUPPLIERS_FILE)
    df[SUPPLIER_VAT_COL] = (
        df[SUPPLIER_VAT_COL]
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .str.zfill(9)
    )
    df = df.drop_duplicates(subset=[SUPPLIER_VAT_COL], keep="first")
    return dict(zip(df[SUPPLIER_VAT_COL], df[SUPPLIER_NAME_COL]))


def load_locations() -> dict[str, list[dict]]:
    """
    Reads the location/cost-centre lookup table from an Excel file.

    Each location entry contains a list of weighted keywords:
      - 5-digit postal codes found in the address field → weight 10
      - Custom keywords from the KEYWORDS column         → weight 10 (postal) / 3 (text)

    Returns a dict mapping location_code → list of {'word': str, 'weight': int}.
    """
    print("🏢 Loading location list...")

    try:
        df = pd.read_excel(LOCATIONS_FILE, header=None)
    except FileNotFoundError:
        print("⚠️  locations.xlsx not found — location attribution disabled.")
        return {}

    code_col = addr_col = keys_col = None
    start_row = 0

    # Detect header row dynamically
    for i, row in df.iterrows():
        row_upper = [str(c).strip().upper() for c in row.values]
        if LOC_CODE_COL.upper() in row_upper:
            code_col  = row_upper.index(LOC_CODE_COL.upper())
            start_row = i + 1
        if LOC_ADDR_COL.upper() in row_upper:
            addr_col = row_upper.index(LOC_ADDR_COL.upper())
        for j, cell in enumerate(row_upper):
            if LOC_KEYS_COL.upper() in cell:
                keys_col = j
        if code_col is not None and keys_col is not None:
            break

    if code_col is None:
        print(f"⚠️  Column '{LOC_CODE_COL}' not found in locations.xlsx.")
        return {}

    locations: dict[str, list[dict]] = {}

    for i in range(start_row, len(df)):
        row  = df.iloc[i]
        code = str(row[code_col]).strip().replace(".0", "")
        if code in ("nan", "") or not code.isdigit():
            continue

        keywords: list[dict] = []

        # Postal code extracted from address → high confidence
        if addr_col is not None:
            address  = str(row[addr_col])
            tk_match = re.search(r"\b(\d{3})\s*(\d{2})\b", address)
            if tk_match:
                postal = tk_match.group(1) + tk_match.group(2)
                keywords.append({"word": postal, "weight": 10})

        # Custom keywords from the KEYWORDS column
        if keys_col is not None:
            kw_cell = str(row[keys_col])
            if kw_cell != "nan":
                for kw in (k.strip() for k in kw_cell.split(",") if k.strip()):
                    is_postal = kw.replace(" ", "").isdigit() and len(kw.replace(" ", "")) == 5
                    keywords.append({"word": kw, "weight": 10 if is_postal else 3})

        if keywords:
            locations[code] = keywords

    print(f"✅ Loaded {len(locations)} location(s).")
    return locations


# =============================================================================
# 4. EXTRACTION HELPERS
# =============================================================================

def remove_accents(text: str) -> str:
    """Uppercases a string and strips diacritical marks for fuzzy matching."""
    if not isinstance(text, str):
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", text.upper())
        if unicodedata.category(c) != "Mn"
    )


def identify_location(raw_text: str, locations: dict[str, list[dict]]) -> str:
    """
    Scores each known location against the invoice text using keyword matching.
    Returns the location code with the highest score, or '0000' if none match.

    Postal codes (weight=10) dominate plain text keywords (weight=3),
    which greatly reduces false positives for locations in the same city.
    """
    clean      = remove_accents(re.sub(r"\s+", " ", raw_text))
    clean_bare = re.sub(r"[^\w\s]", "", clean)

    best_code  = "0000"
    best_score = 0

    for code, keywords in locations.items():
        score = 0
        for item in keywords:
            kw      = remove_accents(str(item["word"]).strip())
            kw_bare = re.sub(r"[^\w\s]", "", kw)
            if kw and (kw in clean or kw_bare in clean_bare):
                score += item["weight"]
        if score > best_score:
            best_score = score
            best_code  = code

    if best_score > 0:
        print(f"🎯 Location match: {best_code} (confidence score: {best_score})")
    return best_code


def clean_filename(name: str, max_len: int = 40) -> str:
    """Strips filesystem-unsafe characters and truncates to max_len."""
    safe = re.sub(r'[\\/*?:"<>|]', "", str(name).replace("\xa0", " ")).strip()
    return safe[:max_len].strip()


def extract_invoice_number(raw_text: str) -> str | None:
    """
    Tries each pattern in INVOICE_PATTERNS in order.
    Returns the first captured group on a match, or None.
    """
    for pattern in INVOICE_PATTERNS:
        try:
            match = re.search(pattern, raw_text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        except re.error as exc:
            print(f"⚠️  Invalid regex pattern '{pattern}': {exc}")
    return None


def extract_date(raw_text: str) -> str | None:
    """
    Searches the first 30 lines of the invoice text for a date.
    Context-aware patterns (e.g. 'DATE:', 'ISSUED:') are tried first,
    followed by a bare date scan.
    Returns 'DD-MM-YYYY' or None.
    """
    top = "\n".join(raw_text.splitlines()[:30])
    base = r"\b(\d{1,2})\s*[/.\s-]\s*(\d{1,2})\s*[/.\s-]\s*(20\d{2})\b"

    context_patterns = [
        r"DATE\s*:?\s*" + base,
        r"ISSUED\s*:?\s*" + base,
        r"INVOICE\s+DATE\s*:?\s*" + base,
        r"ΗΜΕΡΟΜΗΝΙΑ\s*:?\s*" + base,   # Greek: "date"
        r"ΗΜ/ΝΙΑ\s*:?\s*" + base,        # Greek shorthand
        r"ΕΚΔΟΣΗΣ\s*:?\s*" + base,        # Greek: "of issue"
    ]

    for cp in context_patterns:
        match = re.search(cp, top, re.IGNORECASE)
        if match:
            day, month, year = match.groups()
            return f"{day.zfill(2)}-{month.zfill(2)}-{year}"

    for day, month, year in re.findall(base, top):
        if 1 <= int(day) <= 31 and 1 <= int(month) <= 12:
            return f"{day.zfill(2)}-{month.zfill(2)}-{year}"

    return None


# =============================================================================
# 5. MAIN PROCESSOR
# =============================================================================

def process_pdfs() -> None:
    """
    Iterates over all PDFs in PDF_FOLDER:

    1. Extracts text with pdfplumber (fast, text-based PDFs).
    2. Falls back to PyMuPDF if pdfplumber yields nothing.
    3. Falls back to EasyOCR (English, CPU) if no VAT found via text extraction.
    4. Renames the file to: DATE_LOCATIONCODE_SUPPLIER_INVOICENUMBER.pdf
    5. Moves unresolvable files to TRASH_FOLDER for manual review.
    """
    os.makedirs(TRASH_FOLDER, exist_ok=True)

    known_suppliers = load_suppliers()
    known_locations = load_locations()

    print(
        f"✅ Ready — {len(known_suppliers)} supplier(s), "
        f"{len(known_locations)} location(s).\n" + "=" * 50
    )

    if not os.path.exists(PDF_FOLDER):
        print(f"⚠️  Folder not found: {PDF_FOLDER}")
        return

    ocr_reader = None  # Initialised lazily — EasyOCR load is slow

    for filename in os.listdir(PDF_FOLDER):
        if not filename.lower().endswith(".pdf"):
            continue

        pdf_path   = os.path.join(PDF_FOLDER, filename)
        trash_path = os.path.join(TRASH_FOLDER, filename)
        raw_text   = ""
        supplier   = None

        # ── Text extraction: pdfplumber ───────────────────────────────────
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    extracted = page.extract_text()
                    if extracted:
                        raw_text += extracted + "\n"
        except Exception:
            pass

        # ── Text extraction: PyMuPDF fallback ────────────────────────────
        if not raw_text.strip():
            try:
                doc = fitz.open(pdf_path)
                for page in doc:
                    extracted = page.get_text("text")
                    if extracted:
                        raw_text += extracted + "\n"
                doc.close()
            except Exception:
                pass

        raw_text = raw_text.replace("\xa0", " ").replace("\u200b", "")

        # ── VAT lookup in extracted text ──────────────────────────────────
        for afm in re.findall(r"\b(?:EL[- \t]*)?(\d{9})\b", raw_text, re.IGNORECASE):
            if afm in known_suppliers:
                supplier = known_suppliers[afm]
                break

        # ── OCR fallback (only if VAT not found via text) ────────────────
        if not supplier:
            try:
                if ocr_reader is None:
                    print("\n🤖 [OCR] Loading reader (English + digits only)...")
                    ocr_reader = easyocr.Reader(["en"], gpu=False)

                print(f"📸 [OCR] Scanning first page of: {filename}")
                doc = fitz.open(pdf_path)
                pix      = doc[0].get_pixmap(dpi=100)
                img_data = pix.tobytes("png")
                doc.close()

                ocr_text = " ".join(ocr_reader.readtext(img_data, detail=0)) + "\n"

                del img_data, pix
                gc.collect()

                ocr_text = ocr_text.replace("\xa0", " ").replace("\u200b", "")
                for afm in re.findall(r"\b(?:EL[- \t]*)?(\d{9})\b", ocr_text, re.IGNORECASE):
                    if afm in known_suppliers:
                        supplier = known_suppliers[afm]
                        print(f"🎯 [OCR] VAT matched: {afm}")
                        break

            except Exception as exc:
                print(f"❌ [OCR] Error: {exc}")

        # ── Guard: empty and unidentified ─────────────────────────────────
        if not raw_text.strip() and not supplier:
            shutil.move(pdf_path, trash_path)
            print(f"🗑️  [Empty/Corrupt] Moved to trash: {filename}")
            continue

        # ── Rename ────────────────────────────────────────────────────────
        if supplier:
            date_str      = extract_date(raw_text) or "0000-00-00"
            invoice_num   = extract_invoice_number(raw_text)
            location_code = identify_location(raw_text, known_locations)

            safe_supplier = clean_filename(supplier)
            safe_invoice  = clean_filename(invoice_num) if invoice_num else "UNKNOWN"

            base_name    = f"{date_str}_{location_code}_{safe_supplier}_{safe_invoice}"
            new_filename = f"{base_name}.pdf"
            new_path     = os.path.join(PDF_FOLDER, new_filename)

            counter = 1
            while os.path.exists(new_path) and new_path != pdf_path:
                new_filename = f"{base_name}_{counter}.pdf"
                new_path     = os.path.join(PDF_FOLDER, new_filename)
                counter += 1

            try:
                if pdf_path != new_path:
                    os.rename(pdf_path, new_path)
                    print(f"✅ Renamed: {filename}  →  {new_filename}")
                else:
                    print(f"⏩ Already correctly named: {filename}")
            except Exception as exc:
                print(f"❌ Rename failed for {filename}: {exc}")
        else:
            print(f"⚠️  [Unknown supplier] {filename} — could not match VAT via text or OCR.")

    print("\n🏁 EXTRACTION & RENAME PHASE COMPLETE!")


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    process_pdfs()
