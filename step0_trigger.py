# -*- coding: utf-8 -*-
"""
Invoice Pipeline Orchestrator — Step 0: Trigger & Concurrency Control
Author: [Your Name or GitHub Handle]
Date: 2026

Description:
    Entry point for the full 3-step invoice processing pipeline.

    Also implements a simple file-based concurrency lock: if the pipeline is
    deployed on multiple machines sharing a network drive, only the first
    machine to run on a given day will proceed; subsequent runs exit silently.
    Disable this behaviour by leaving LOCK_DIR unset in your .env.

Usage:
    python step0_trigger.py
"""

import os
import sys
from datetime import datetime
from dotenv import load_dotenv

# Required to avoid duplicate-library crashes on some Windows setups
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

load_dotenv()

import step1_download
import step2_rename
import step3_organize


# =============================================================================
# CONCURRENCY LOCK  (optional — only active when LOCK_DIR is set in .env)
# =============================================================================

LOCK_DIR  = os.getenv("LOCK_DIR", "")          # e.g. a shared network folder
LOCK_FILE = os.path.join(LOCK_DIR, "daily_lock.txt") if LOCK_DIR else ""


def check_and_acquire_lock() -> bool:
    """
    Tries to acquire today's run token.

    Returns True  if this machine should proceed (token written successfully).
    Returns False if another machine already ran today (token already present).
    Skips the check entirely if LOCK_DIR is not configured.
    """
    if not LOCK_DIR:
        return True     # Lock disabled — always proceed

    today = datetime.now().strftime("%d-%m-%Y")
    os.makedirs(LOCK_DIR, exist_ok=True)

    if os.path.exists(LOCK_FILE):
        with open(LOCK_FILE, "r", encoding="utf-8") as fh:
            saved_date = fh.read().strip()
        if saved_date == today:
            print(f"\n🛑 Pipeline already ran today ({today}) on another machine. Exiting.")
            return False

    with open(LOCK_FILE, "w", encoding="utf-8") as fh:
        fh.write(today)

    print(f"\n✅ Lock acquired for {today}. This machine will run the pipeline.\n")
    return True


# =============================================================================
# MAIN
# =============================================================================

def run_all() -> None:
    """Runs all three pipeline steps in sequence."""
    print("🚀 Invoice Pipeline — Starting...\n")

    if not check_and_acquire_lock():
        return

    print("=" * 50)
    print("▶  STEP 1 — Download")
    print("=" * 50)
    step1_download.run_download()

    print("\n" + "=" * 50)
    print("▶  STEP 2 — Extract & Rename")
    print("=" * 50)
    step2_rename.process_pdfs()

    print("\n" + "=" * 50)
    print("▶  STEP 3 — Organise")
    print("=" * 50)
    step3_organize.organize_and_cleanup()

    print("\n🎉 All steps completed successfully!")


if __name__ == "__main__":
    try:
        run_all()
    except Exception as exc:
        print(f"\n❌ Fatal error: {exc}")
    finally:
        input("\nPress ENTER to exit...")
