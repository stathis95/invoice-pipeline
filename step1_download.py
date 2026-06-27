# -*- coding: utf-8 -*-
"""
Automated Invoice Downloader
Author: [Your Name or GitHub Handle]
Date: 2026

Description:
    Step 1 of the invoice processing pipeline.
    Connects to an IMAP mailbox, scans a designated label for invoice emails,
    extracts download links or direct PDF attachments, and saves everything to
    a date-stamped local folder.

    Supported download strategies:
      - Direct PDF attachments from trusted senders
      - Browser-rendered PDFs via Chrome DevTools Protocol (CDP)
      - Cloudflare CAPTCHA detection with human-like wait loop
      - Retry queue for transiently failed links

Usage:
    Ensure a .env file exists (see .env.example), then run:
        python step1_download.py
"""

import os
import re
import time
import base64
import random
import logging
import datetime

from dotenv import load_dotenv
from imap_tools import MailBox, AND
from DrissionPage import ChromiumPage, ChromiumOptions

# Load environment variables from .env file
load_dotenv()

# Suppress noisy PDF-parser warnings from imap_tools internals
logging.getLogger("pdfminer").setLevel(logging.ERROR)


# =============================================================================
# 1. CONFIGURATION  (all sensitive values come from .env)
# =============================================================================

USERNAME    = os.getenv("IMAP_USERNAME")
PASSWORD    = os.getenv("IMAP_PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER", "imap.gmail.com")
LABEL_NAME  = os.getenv("IMAP_LABEL",  "accounting")
BASE_FOLDER = os.getenv("INVOICE_BASE_FOLDER", os.path.join(os.getcwd(), "invoices"))

if not USERNAME or not PASSWORD:
    raise EnvironmentError(
        "IMAP_USERNAME and IMAP_PASSWORD must be set in your .env file.\n"
        "See .env.example for the required variables."
    )

# Trusted sender domains — emails from these are always considered valid invoice sources
TRUSTED_DOMAINS = [
    "provider-a.com",
    "provider-b.gr",
    "erp-system.com",
    "e-invoice-hub.gr",
]


# =============================================================================
# 2. TRUSTED SUPPLIER EMAILS  (loaded from an external flat file)
# =============================================================================

EMAILS_FILE = os.path.join(BASE_FOLDER, "trusted_emails.txt")


def load_trusted_emails() -> list[str]:
    """
    Reads extra trusted supplier email addresses from a plain-text file
    (one address per line; lines starting with '#' are treated as comments).

    Falls back to a small hardcoded list if the file is absent.
    """
    emails: list[str] = []

    if os.path.exists(EMAILS_FILE):
        with open(EMAILS_FILE, "r", encoding="utf-8") as fh:
            for line in fh:
                clean = line.strip().lower()
                if clean and not clean.startswith("#"):
                    emails.append(clean)
        print(f"✅ Loaded {len(emails)} trusted supplier email(s) from file.")
    else:
        print("⚠️  trusted_emails.txt not found — using built-in fallback list.")
        emails = [
            "noreply@supplier1.com",
            "no-reply@supplier2.gr",
            "invoices@supplier3.com",
        ]

    return emails


ADDITIONAL_SUPPLIER_EMAILS = load_trusted_emails()


# =============================================================================
# 3. DATE & FOLDER SETUP
# =============================================================================

TARGET_DATE  = datetime.date.today() - datetime.timedelta(days=1)
DAILY_FOLDER = os.path.join(BASE_FOLDER, TARGET_DATE.strftime("%d-%m-%Y"))
os.makedirs(DAILY_FOLDER, exist_ok=True)

print(f"📁 Output folder: {DAILY_FOLDER}")


# =============================================================================
# 4. INVOICE LINK PATTERNS
#    One regex per provider portal; add new patterns here as needed.
# =============================================================================

TARGET_PATTERNS = [
    r"(https://(?:www\.)?provider-a\.gr/download/[a-zA-Z0-9\-]+)",
    r"(https://einvoice\.provider-b\.com/p/[a-zA-Z0-9]+/[a-zA-Z0-9]+/[a-zA-Z0-9]+[^\s<>\"'\]]*)",
    r"(https://documents[a-zA-Z0-9-]*\.erp-system\.com/(?:fd|FileDocument)/[a-zA-Z0-9\-:/]+)",
    r"(https://(?:www\.)?e-invoice-hub\.gr/invoices/[a-zA-Z0-9]+)",
    r"(https://[a-zA-Z0-9.-]+\.email-sender\.net/ls/click\?upn=[a-zA-Z0-9\-_.~%+]+)",
]


# =============================================================================
# 5. HELPERS
# =============================================================================

def human_delay(min_sec: float = 15, max_sec: float = 25) -> None:
    """Waits a randomised amount of time to mimic human browsing pace."""
    delay = random.uniform(min_sec, max_sec)
    print(f"☕ Pausing for {delay:.1f}s...")
    time.sleep(delay)


def short_pause(min_sec: float = 3, max_sec: float = 6) -> None:
    """Short randomised pause between UI interactions."""
    time.sleep(random.uniform(min_sec, max_sec))


# =============================================================================
# 6. EMAIL COLLECTOR
# =============================================================================

def collect_and_download_attachments() -> list[str]:
    """
    Connects to the IMAP mailbox and scans all emails from TARGET_DATE.

    For each message it:
      1. Searches the body for known invoice-link patterns.
      2. Falls back to downloading PDF attachments directly if the sender
         is trusted and no link was found.

    Returns a list of URLs to be processed by the browser downloader.
    """
    invoice_links: list[str] = []
    print(f"🔓 Connecting to {IMAP_SERVER} — checking mail for {TARGET_DATE.strftime('%d/%m/%Y')}.")

    try:
        with MailBox(IMAP_SERVER).login(USERNAME, PASSWORD) as mailbox:
            mailbox.folder.set(LABEL_NAME)

            for msg in mailbox.fetch(AND(date=TARGET_DATE), reverse=True):
                body = msg.html or msg.text or ""
                has_found_link = False

                # --- Link extraction ---
                for pattern in TARGET_PATTERNS:
                    for link in re.findall(pattern, body):
                        clean_link = link.strip("'\"<>")
                        if clean_link not in invoice_links:
                            invoice_links.append(clean_link)
                            has_found_link = True

                # --- Trust check for attachment fallback ---
                sender_headers = f"{msg.from_} {msg.reply_to}".lower()
                if "from" in msg.headers:
                    sender_headers += " " + " ".join(msg.headers["from"]).lower()

                sender_domain = (
                    msg.from_.split("@")[-1].lower() if "@" in msg.from_ else ""
                )

                is_trusted = any(d in sender_domain for d in TRUSTED_DOMAINS) or \
                             any(e in sender_headers for e in ADDITIONAL_SUPPLIER_EMAILS)

                # --- Direct attachment download (only if no link was found) ---
                if not has_found_link and is_trusted:
                    for att in msg.attachments:
                        fname = att.filename.lower()
                        is_pdf = fname.endswith(".pdf")
                        is_junk = re.search(r"(terms|gdpr|policy|flyer|promo)", fname)

                        if is_pdf and not is_junk:
                            save_path = os.path.join(
                                DAILY_FOLDER,
                                f"TEMP_DirectAttach_{int(time.time())}_{att.filename}",
                            )
                            with open(save_path, "wb") as fh:
                                fh.write(att.payload)
                            print(f"📎 [Direct attachment] Saved: {att.filename}")

    except Exception as exc:
        print(f"❌ Email error: {exc}")

    return invoice_links


# =============================================================================
# 7. BROWSER DOWNLOADER
# =============================================================================

def _cdp_print(page_or_tab, save_path: str) -> None:
    """Prints the current page to PDF using the Chrome DevTools Protocol."""
    result = page_or_tab.run_cdp(
        "Page.printToPDF",
        printBackground=True,
        preferCSSPageSize=True,
        paperWidth=8.27,
        paperHeight=11.69,
    )
    with open(save_path, "wb") as fh:
        fh.write(base64.b64decode(result["data"]))


def _handle_provider_a(page: ChromiumPage, save_path: str, failed_links: list, url: str) -> None:
    """
    Provider A — dynamic invoice viewer with optional Cloudflare CAPTCHA.

    Strategy:
      1. Poll for the 'Save' button (invoice loaded) OR a Cloudflare iframe.
      2. If CAPTCHA is detected, enter a human-scroll loop for up to 2 minutes.
      3. Once the page is clear, trigger CDP print-to-PDF.
    """
    print("🔍 Provider A: Checking for invoice page or CAPTCHA...")
    check_load = False
    cf_iframe  = None

    for _ in range(20):
        check_load = (
            page.ele("text:Save",        timeout=0.5) or
            page.ele("text:Αποθήκευση", timeout=0.5)
        )
        if check_load:
            print("✅ Save button found.")
            break
        cf_iframe = page.ele('xpath://iframe[contains(@src, "cloudflare")]', timeout=0.5)
        if cf_iframe:
            print("🛡️  Cloudflare CAPTCHA iframe detected.")
            break

    if cf_iframe and not check_load:
        print("🤖 Waiting for CAPTCHA resolution (timeout: 120s)...")
        max_wait = 120
        start_t  = time.time()
        passed   = False

        while time.time() - start_t < max_wait:
            try:
                page.scroll.down(random.randint(10, 50))
                short_pause(1, 3)
                page.scroll.up(random.randint(10, 50))
                check_load = (
                    page.ele("text:Save",        timeout=1) or
                    page.ele("text:Αποθήκευση", timeout=1)
                )
                if check_load:
                    elapsed = int(time.time() - start_t)
                    print(f"✅ Passed CAPTCHA in {elapsed}s.")
                    passed = True
                    break
            except Exception:
                pass

        if not passed:
            print(f"⚠️  CAPTCHA timeout ({max_wait}s). Adding to retry queue.")

    elif not check_load and not cf_iframe:
        print("ℹ️  No invoice page or CAPTCHA detected on initial poll.")

    if check_load:
        short_pause(1, 2)
        pdf_radio = page.ele("text:PDF", timeout=3)
        if pdf_radio and pdf_radio.next():
            pdf_radio.next().click(by_js=True)
            short_pause(3, 5)
        else:
            print("⚠️  PDF radio button not found. Proceeding directly to CDP print.")

        _cdp_print(page, save_path)
        print(f"✅ Saved via CDP print: {os.path.basename(save_path)}")
    else:
        print("⚠️  Invoice page not ready. Adding to retry queue.")
        failed_links.append(url)


def _handle_provider_b(page: ChromiumPage, save_path: str, failed_links: list, url: str) -> None:
    """
    Provider B — ERP portal with a native download button.

    Strategy:
      1. Look for a download button by text or attribute.
      2. If found and a native download starts  → wait for completion.
      3. If a new HTML tab opens instead        → CDP print that tab.
      4. If no button at all                    → CDP print the current page.
    """
    print("🏢 Provider B: Looking for download button...")
    short_pause(2, 4)

    dl_btn = (
        page.ele("text:Κατεβάστε",    timeout=2) or
        page.ele("text:Λήψη",         timeout=2) or
        page.ele("text:Download",     timeout=2) or
        page.ele("@title:Λήψη",       timeout=2) or
        page.ele("@title:Download",   timeout=2) or
        page.ele("#download",         timeout=2)
    )

    if dl_btn:
        print("🖱️  Download button found. Clicking...")
        dl_btn.click(by_js=True)
        mission = page.wait.download_begin(timeout=8)

        if mission:
            print("✅ Native PDF download started.")
            mission.wait()
            return

        print("📑 New HTML tab opened. Printing via CDP...")
        new_tab_id = page.wait.new_tab(timeout=5)
        if new_tab_id:
            pdf_tab = page.get_tab(new_tab_id)
            pdf_tab.wait.load_start()
            short_pause(1, 3)
            _cdp_print(pdf_tab, save_path)
            pdf_tab.close()
            print(f"✅ Saved via CDP print: {os.path.basename(save_path)}")
        else:
            print("⚠️  Download did not start. Adding to retry queue.")
            failed_links.append(url)
    else:
        print("ℹ️  No download button found — falling back to CDP print.")
        _cdp_print(page, save_path)
        print(f"✅ Saved via CDP print: {os.path.basename(save_path)}")


def run_download() -> None:
    """
    Main entry point for Step 1.

    Orchestrates email collection followed by browser-based downloading.
    Implements a two-pass retry queue: any URL that fails on the first pass
    is retried once before the process terminates.
    """
    links_to_download = collect_and_download_attachments()

    if not links_to_download:
        print("\n⚠️  No invoice links found. Exiting download phase.")
        return

    print(f"\n🎯 Found {len(links_to_download)} link(s). Starting browser...\n" + "=" * 50)

    options = ChromiumOptions()
    options.set_argument("--start-maximized")
    options.set_pref("plugins.always_open_pdf_externally", True)

    page = ChromiumPage(options)
    page.set.download_path(DAILY_FOLDER)

    links_to_process = links_to_download.copy()
    max_attempts    = 2
    current_attempt = 1

    while current_attempt <= max_attempts and links_to_process:
        if current_attempt > 1:
            print(f"\n🔄 Retry pass {current_attempt} — {len(links_to_process)} link(s) remaining...")

        failed_links: list[str] = []

        for index, url in enumerate(links_to_process, start=1):
            if index > 1:
                print("\n⏳ Preparing for next download...")
                human_delay(10, 20)

            print(f"\n📥 [{index}/{len(links_to_process)}] {url}")

            domain_name = (
                url.split("//")[1].split("/")[0]
                   .replace("www.", "")
                   .split(".")[0]
                   .capitalize()
            )
            save_path = os.path.join(
                DAILY_FOLDER, f"TEMP_{domain_name}_{index}_{int(time.time())}.pdf"
            )

            try:
                page.get(url)
                short_pause(1, 3)

                if "provider-a.gr" in url:
                    _handle_provider_a(page, save_path, failed_links, url)

                elif "provider-b.gr" in url:
                    _handle_provider_b(page, save_path, failed_links, url)

                else:
                    # Generic fallback: CDP print for any unrecognised portal
                    short_pause(3, 5)
                    _cdp_print(page, save_path)
                    print(f"✅ Saved via CDP print: {os.path.basename(save_path)}")

            except Exception as exc:
                print(f"⚠️  Unexpected error: {exc}. Adding to retry queue.")
                failed_links.append(url)
                continue

        links_to_process = failed_links
        current_attempt += 1

    print("\n🏁 DOWNLOAD PHASE COMPLETE!")


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    run_download()
