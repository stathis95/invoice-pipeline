# 🧾 Invoice Processing Pipeline

An automated 3-step pipeline for downloading, extracting, and organising supplier invoice PDFs — built for accounts-payable workflows that receive invoices via email across multiple supplier portals.

---

## Pipeline Overview

```
┌─────────────────────────────────────────────────────────┐
│  STEP 1 — Download                 step1_download.py    │
│                                                         │
│  📧 IMAP mailbox scan                                   │
│     ├── Extract invoice links from email body           │
│     └── Save direct PDF attachments (trusted senders)   │
│                                                         │
│  🌐 Browser automation (DrissionPage + CDP)             │
│     ├── Navigate to each link                           │
│     ├── Detect & wait out Cloudflare CAPTCHAs           │
│     ├── Download native PDFs or print-to-PDF via CDP    │
│     └── Retry queue for transiently failed links        │
│                                                         │
│  Output: invoices/DD-MM-YYYY/TEMP_*.pdf                 │
└────────────────────────┬────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────┐
│  STEP 2 — Extract & Rename         step2_rename.py      │
│                                                         │
│  📄 PDF text extraction (pdfplumber → PyMuPDF → OCR)   │
│     ├── Match 9-digit VAT number → supplier name        │
│     ├── Extract invoice number via configurable regex   │
│     ├── Extract issue date (context-aware patterns)     │
│     └── Attribute location via keyword scoring          │
│                                                         │
│  Output: invoices/DD-MM-YYYY/                           │
│          DD-MM-YYYY_LOCCODE_SUPPLIER_INVNO.pdf          │
└────────────────────────┬────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────┐
│  STEP 3 — Organise                 step3_organize.py    │
│                                                         │
│  📂 Dual parallel folder structure                      │
│     ├── by_supplier / location / supplier / date /      │
│     │     invoice.pdf          ← primary archive        │
│     └── by_date / location / date /                     │
│             supplier_invoice.pdf  ← date-based index    │
│                                                         │
│  Source date folder removed if empty after processing.  │
└─────────────────────────────────────────────────────────┘
```

---

## Features

- **Multi-strategy PDF acquisition** — direct attachments, native portal downloads, and CDP print-to-PDF for HTML-only pages
- **Layered text extraction** — pdfplumber for text-based PDFs, PyMuPDF as fallback, EasyOCR for scanned documents
- **VAT-based supplier identification** — 9-digit VAT lookup against a configurable Excel table
- **Configurable invoice-number extraction** — regex rules loaded from an external `.txt` file; add new patterns without touching code
- **Location attribution via keyword scoring** — postal codes (high weight) and text keywords (low weight) to minimise false positives in same-city locations
- **Standardised filename format** — `DD-MM-YYYY_LOCCODE_SUPPLIER_INVNO.pdf` for predictable downstream processing
- **Dual folder organisation** — simultaneous archive by supplier and index by date
- **Optional concurrency lock** — file-based daily token on a shared network drive prevents duplicate runs across multiple machines
- **Unprocessable file triage** — corrupt or unidentifiable PDFs moved to a dedicated review folder

---

## Tech Stack

| Library | Purpose |
|---|---|
| [imap-tools](https://github.com/ikvk/imap_tools) | IMAP email scanning and attachment handling |
| [DrissionPage](https://github.com/g1879/DrissionPage) | Browser automation via Chromium |
| [pdfplumber](https://github.com/jsvine/pdfplumber) | Primary PDF text extraction |
| [PyMuPDF](https://pymupdf.readthedocs.io/) | Fallback text extraction + OCR pre-processing |
| [EasyOCR](https://github.com/JaidedAI/EasyOCR) | OCR for scanned / image-based invoices |
| [pandas](https://pandas.pydata.org/) | Excel lookup tables (suppliers, locations) |
| [python-dotenv](https://github.com/theskumar/python-dotenv) | Environment-based configuration |
| [OpenCV](https://opencv.org/) *(optional)* | Computer vision for CAPTCHA element detection |

---

## Project Structure

```
invoice-pipeline/
│
├── step0_trigger.py        # Orchestrator + optional concurrency lock
├── step1_download.py       # IMAP scanner + browser downloader
├── step2_rename.py         # PDF extractor + renamer
├── step3_organize.py       # Dual folder organiser
│
├── invoice_rules.txt       # Configurable regex patterns for invoice numbers
│
├── .env.example            # Environment variable template (commit this)
├── .env                    # Your local values            (never commit)
├── .gitignore
├── requirements.txt
└── README.md
```

---

## Setup

### 1. Clone & install

```bash
git clone https://github.com/yourusername/invoice-pipeline.git
cd invoice-pipeline
pip install -r requirements.txt
```

### 2. Configure environment

```bash
cp .env.example .env
# Open .env and fill in your values
```

### 3. Prepare lookup files

Place these in your `INVOICE_BASE_FOLDER` (or the paths you set in `.env`):

**`suppliers.xlsx`**

| VAT_Number  | Supplier_Name       |
|-------------|---------------------|
| 123456789   | Acme Corp           |
| 987654321   | Global Supplies Ltd |

**`locations.xlsx`**

| LOCATION_CODE | ADDRESS                        | KEYWORDS              |
|---------------|--------------------------------|-----------------------|
| 0001          | 12 Main St, 10431 Athens       | Main Street, downtown |
| 0002          | 7 Port Ave, 18536 Piraeus      | Port, harbour         |

**`invoice_rules.txt`**

```
# One regex per line — first match wins
# Capture group 1 must contain the invoice number
INVOICE\s*NO\.?\s*:?\s*([A-Z0-9/_-]+)
INV[#\s-]*([0-9]{4,})
ΑΡ\.\s*ΕΝΤΥΠΟΥ\s*:\s*([Α-ΩA-Z0-9_.-]+)
```

### 4. Gmail setup (if using Gmail)

1. Enable IMAP: *Gmail Settings → See all settings → Forwarding and POP/IMAP → Enable IMAP*
2. Generate an App Password: *Google Account → Security → 2-Step Verification → App passwords*
3. Create a Gmail label (e.g. `accounting`) and set up a filter to route invoice emails there

---

## Usage

```bash
# Run the full pipeline (Steps 1 → 2 → 3)
python step0_trigger.py

# Run individual steps
python step1_download.py
python step2_rename.py
python step3_organize.py
```

---

## Configuration Reference

All configuration is via the `.env` file. See [`.env.example`](.env.example) for the full list with descriptions.

| Variable | Used by | Description |
|---|---|---|
| `IMAP_USERNAME` | Step 1 | Email address |
| `IMAP_PASSWORD` | Step 1 | App password |
| `IMAP_SERVER` | Step 1 | IMAP hostname (default: `imap.gmail.com`) |
| `IMAP_LABEL` | Step 1 | Mailbox label/folder to scan |
| `INVOICE_BASE_FOLDER` | All | Root folder for all invoice data |
| `SUPPLIERS_FILE` | Step 2 | Path to `suppliers.xlsx` |
| `LOCATIONS_FILE` | Step 2 | Path to `locations.xlsx` |
| `TRASH_FOLDER` | Step 2 | Destination for unprocessable files |
| `SUPPLIER_VAT_COL` | Step 2 | VAT column name in suppliers.xlsx |
| `SUPPLIER_NAME_COL` | Step 2 | Name column name in suppliers.xlsx |
| `LOC_CODE_COL` | Step 2 | Code column name in locations.xlsx |
| `LOC_ADDR_COL` | Step 2 | Address column name in locations.xlsx |
| `LOC_KEYS_COL` | Step 2 | Keywords column name in locations.xlsx |
| `ORGANIZED_BY_SUPPLIER` | Step 3 | Root of the by-supplier archive |
| `ORGANIZED_BY_DATE` | Step 3 | Root of the by-date index |
| `LOCK_DIR` | Step 0 | Shared folder for the concurrency lock (optional) |

---

## Notes

### Adding a new supplier portal

1. Add the portal's invoice-link URL pattern to `TARGET_PATTERNS` in `step1_download.py`
2. If the portal has non-standard download behaviour, add a handler function following the `_handle_provider_a` / `_handle_provider_b` pattern and register it in `run_download()`

### Adding a new invoice-number format

Add a new regex line to `invoice_rules.txt` — no code changes required. Patterns are tried in order; the first match wins.

### Computer vision module

For portals with particularly persistent anti-bot protection, the pipeline was extended with an optional computer vision component (OpenCV multi-scale template matching + PyAutoGUI hardware mouse simulation) to handle Cloudflare Turnstile CAPTCHAs that do not resolve via the standard scroll-and-wait loop. This module is not included in the public repository.

---

## License

MIT
