"""
FSG CCBOM Automation Tool
=========================

Automates the upload of Costed Carbonized Bill of Material (CCBOM) data
from Excel files to the Formula Student Germany (FSG) website.

Developed by ELBFLORACE e.V. — Dresden, Germany.
Open-sourced for all Formula Student teams.

Usage:
    python bom_automation.py

Configuration:
    See .env.example for all available options.
"""

import os
import sys
import glob
import time
import pandas as pd
import openpyxl
from datetime import datetime
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

load_dotenv()

# FSG Website URLs — TEAM_ID must be explicitly set in .env
TEAM_ID = os.getenv("TEAM_ID", "").strip()
if not TEAM_ID:
    print("ERROR: TEAM_ID not set. Copy .env.example to .env and set TEAM_ID.")
    sys.exit(1)
BASE_URL = "https://www.formulastudent.de"
LOGIN_URL = f"{BASE_URL}/login"
BOM_URL = f"{BASE_URL}/teams/fse/details/bom/tid/{TEAM_ID}"

# Credentials & Behaviour
FSG_USERNAME = os.getenv("FSG_USERNAME")
FSG_PASSWORD = os.getenv("FSG_PASSWORD")
# Safer default for new users: TEST_MODE enabled by default.
TEST_MODE = os.getenv("TEST_MODE", "true").lower() == "true"
# Dry-run option: if true, log actions but do not perform uploads.
DRY_RUN = os.getenv("DRY_RUN", "false").lower() == "true"
TEST_LIMIT = int(os.getenv("TEST_LIMIT", "3"))
DEFAULT_SYSTEM = os.getenv("DEFAULT_SYSTEM", "").strip().upper()
LOG_FILE = os.getenv("LOG_FILE", "bom_log.txt")
BOMS_DIR = os.getenv("BOMS_DIR", "BOMs")

# Row‑colour hex codes used for skip detection (openpyxl format, no leading #)
SKIP_COLORS = {
    "FF00FF00",   # Pure Green  — already uploaded
    "0000FF00",   # Green variant
    "FFFF0000",   # Pure Red    — do not upload
    "00FF0000",   # Red variant
}

# ──────────────────────────────────────────────────────────────────────────────
# FSG System Codes → Full Dropdown Labels
# ──────────────────────────────────────────────────────────────────────────────

SYSTEM_MAP = {
    "AT": "AT - Autonomous System",
    "BR": "BR - Brake System",
    "DT": "DT - Drivetrain",
    "ET": "ET - Engine and Tractive System",
    "FR": "FR - Chassis and Body",
    "LV": "LV - Grounded Low Voltage System",
    "MS": "MS - Miscellaneous Fit and Finish",
    "ST": "ST - Steering System",
    "SU": "SU - Suspension System",
    "WT": "WT - Wheels, Wheel Bearings and Tires",
}

# ──────────────────────────────────────────────────────────────────────────────
# Smart Assembly Remapping
#
# If your Excel uses a different name for an assembly than the FSG dropdown,
# add the mapping here.  Keys must be lowercase.
# ──────────────────────────────────────────────────────────────────────────────

ASSEMBLY_REMAP = {
    # Brake System
    "brake caliper":        "Calipers",
    "brake calipers":       "Calipers",
    "caliper":              "Calipers",
    "reservoire":           "Brake Master Cylinder",
    "reservoir":            "Brake Master Cylinder",
    "resovoir":             "Brake Master Cylinder",
    "fitting screw":        "Fasteners",
    "fastener":             "Fasteners",
    "screws":               "Fasteners",
    "bolts":                "Fasteners",
    "brake disc":           "Brake Discs",
    "brake disk":           "Brake Discs",
    "brake pad":            "Brake Pads",
    "brake line":           "Brake Lines",
    "master cylinder":      "Brake Master Cylinder",
    # Suspension
    "damper":               "Dampers",
    "spring":               "Springs",
    "pushrod":              "Pushrods",
    "rocker":               "Rockers",
    "a-arm":                "A-Arms",
    # Drivetrain
    "chain":                "Chain",
    "sprocket":             "Sprockets",
    "differential":         "Differential",
    "half shaft":           "Half Shafts",
    "halfshaft":            "Half Shafts",
    # Steering
    "steering rack":        "Steering Rack",
    "tie rod":              "Tie Rods",
    "steering wheel":       "Steering Wheel",
    # Wheels & Tires
    "tire":                 "Tires",
    "tyre":                 "Tires",
    "wheel bearing":        "Wheel Bearings",
    "rim":                  "Wheels",
    "wheel":                "Wheels",
}


# ──────────────────────────────────────────────────────────────────────────────
# Logging
# ──────────────────────────────────────────────────────────────────────────────

def log(message: str, status: str = "INFO") -> None:
    """Log to both console and log file with timestamp."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{status}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ──────────────────────────────────────────────────────────────────────────────
# Excel Helpers
# ──────────────────────────────────────────────────────────────────────────────

def get_cell_color(sheet, row: int, col: int = 1) -> str | None:
    """Return the fill colour of a cell as an uppercase hex string, or None."""
    try:
        fill = sheet.cell(row=row, column=col).fill
        if fill.patternType is None:
            return None
        idx = fill.start_color.index
        return str(idx).upper() if idx and idx != "00000000" else None
    except Exception:
        return None


def should_skip_color(sheet, row: int) -> str | None:
    """Check if a row's colour indicates it should be skipped.
    Returns a human-readable reason or None."""
    color = get_cell_color(sheet, row)
    if color is None:
        return None
    if color in SKIP_COLORS:
        if "FF00" in color or "00FF00" in color.replace("FF", "", 1):
            return "green (already uploaded)"
        if "FF0000" in color:
            return "red (do not upload)"
        return f"skipped colour ({color})"
    return None


# ──────────────────────────────────────────────────────────────────────────────
# Fuzzy Dropdown Matching
# ──────────────────────────────────────────────────────────────────────────────

def fuzzy_select(page, selector: str, target: str) -> bool:
    """Select a dropdown option using smart remapping + fuzzy matching.

    Matching priority:
      0. Check ASSEMBLY_REMAP dictionary
      1. Exact match
      2. Case-insensitive match
      3. Substring / contains match
    """
    try:
        # 0 — Remap
        resolved = ASSEMBLY_REMAP.get(target.lower().strip(), target)

        options = page.eval_on_selector(
            selector,
            "el => Array.from(el.options).map(o => o.text)",
        )

        # 1 — Exact
        if resolved in options:
            page.locator(selector).select_option(label=resolved)
            return True

        # 2 — Case-insensitive
        resolved_lower = resolved.lower().strip()
        for opt in options:
            if opt.lower().strip() == resolved_lower:
                page.locator(selector).select_option(label=opt)
                return True

        # 3 — Contains
        for opt in options:
            ol = opt.lower().strip()
            if resolved_lower in ol or ol in resolved_lower:
                page.locator(selector).select_option(label=opt)
                return True

        return False
    except Exception as e:
        log(f"Fuzzy-match error for '{target}': {e}", "WARN")
        return False


# ──────────────────────────────────────────────────────────────────────────────
# File Selection
# ──────────────────────────────────────────────────────────────────────────────

def discover_excel_files() -> list[str]:
    """Look for .xlsx files inside the BOMS_DIR folder."""
    search_dir = os.path.join(os.getcwd(), BOMS_DIR)
    if not os.path.isdir(search_dir):
        os.makedirs(search_dir, exist_ok=True)
        log(f"Created '{BOMS_DIR}/' directory. Place your Excel files there.", "WARN")
        return []
    files = glob.glob(os.path.join(search_dir, "*.xlsx"))
    return sorted(files)


def select_file() -> str:
    """Interactive file picker for Excel BOMs."""
    files = discover_excel_files()
    if not files:
        print(f"\n  No .xlsx files found in '{BOMS_DIR}/'.")
        print("  Place your BOM Excel files there and try again.\n")
        sys.exit(1)

    print(f"\nExcel files in '{BOMS_DIR}/':")
    for i, f in enumerate(files):
        print(f"  {i + 1}. {os.path.basename(f)}")

    while True:
        try:
            choice = int(input("\nSelect a file (number): ")) - 1
            if 0 <= choice < len(files):
                return files[choice]
        except ValueError:
            pass
        print("Invalid — please enter a valid number.")


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def main() -> None:
    log("=" * 60)
    log("FSG CCBOM Automation — Starting")
    if TEST_MODE:
        log(f"TEST MODE: Only the first {TEST_LIMIT} parts will be processed.", "WARN")
    log(f"Config: TEAM_ID={TEAM_ID} TEST_MODE={TEST_MODE} DRY_RUN={DRY_RUN} BOMS_DIR={BOMS_DIR}")

    # ── 1. File selection ────────────────────────────────────────────────
    filepath = select_file()
    filename = os.path.basename(filepath)
    log(f"Selected file: {filename}")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    sheet = wb.active
    df = pd.read_excel(filepath)

    # Normalise column names (strip whitespace, lowercase)
    df.columns = [c.strip().lower() for c in df.columns]

    required_cols = {"system", "assembly", "part"}
    missing = required_cols - set(df.columns)
    if missing:
        log(f"Missing required columns: {missing}. "
            f"Available: {list(df.columns)}", "ERROR")
        sys.exit(1)

    # ── 2. System selection ──────────────────────────────────────────────
    unique_systems = [
        str(s).strip().upper()
        for s in df["system"].dropna().unique()
        if str(s).strip().upper() not in ("NAN", "BEISPIEL", "")
    ]

    if DEFAULT_SYSTEM and DEFAULT_SYSTEM in unique_systems:
        run_system = DEFAULT_SYSTEM
        log(f"System filter (from .env): {run_system} — "
            f"{SYSTEM_MAP.get(run_system, run_system)}")
    else:
        print("\nSystems found in Excel:")
        for s in unique_systems:
            label = SYSTEM_MAP.get(s, s)
            print(f"  • {s:4s}  {label}")
        run_system = input(
            "\nEnter system code to filter (e.g. 'BR') "
            "or 'ALL' for everything: "
        ).strip().upper()

    # ── 3. Filter rows ───────────────────────────────────────────────────
    filtered = []
    skipped_green = 0
    skipped_red = 0
    skipped_example = 0
    skipped_empty = 0

    for idx, row in df.iterrows():
        excel_row = idx + 2  # pandas 0-indexed + header row

        system = str(row.get("system", "")).strip().upper()
        assembly = str(row.get("assembly", "")).strip()
        part = str(row.get("part", "")).strip()
        quantity = row.get("part_quantity", "")
        makebuy = str(row.get("make o. buy", "")).strip().lower()
        comments = str(row.get("part_comments", "")).strip()

        # System filter
        if run_system != "ALL" and system != run_system:
            continue

        # Empty rows
        if (pd.isna(row.get("system")) and pd.isna(row.get("part"))) or \
           not system or system == "NAN" or not part or part == "NAN":
            skipped_empty += 1
            continue

        # Example rows
        if "BEISPIEL" in system or "BEISPIEL" in part.upper() or \
           "EXAMPLE" in system or "EXAMPLE" in part.upper():
            skipped_example += 1
            continue

        # Colour check (green = uploaded, red = do not upload)
        skip_reason = should_skip_color(sheet, excel_row)
        if skip_reason:
            if "green" in skip_reason:
                skipped_green += 1
            elif "red" in skip_reason:
                skipped_red += 1
            log(f"Row {excel_row}: Skipped — {skip_reason}", "SKIP")
            continue

        # Clean make/buy to single char
        mb = makebuy[0] if makebuy and makebuy[0] in ("m", "b") else "m"

        # Clean comments
        if comments in ("nan", "NaN", ""):
            comments = ""

        # Clean quantity
        qty_str = str(quantity).strip()
        if qty_str in ("nan", "NaN", ""):
            qty_str = ""

        filtered.append({
            "row": excel_row,
            "system": system,
            "assembly": assembly,
            "part": part,
            "makebuy": mb,
            "quantity": qty_str,
            "comments": comments,
        })

    # Summary
    log(f"Filtering complete: {len(filtered)} parts to upload "
        f"({skipped_green} green / {skipped_red} red / "
        f"{skipped_example} example / {skipped_empty} empty skipped)")

    if TEST_MODE and len(filtered) > TEST_LIMIT:
        log(f"Test Mode: limiting {len(filtered)} → {TEST_LIMIT} parts")
        filtered = filtered[:TEST_LIMIT]

    if not filtered:
        log("Nothing to upload — exiting.")
        sys.exit(0)

    # If credentials are not provided, require explicit confirmation to proceed in manual login mode.
    if not (FSG_USERNAME and FSG_PASSWORD):
        print("\nWARNING: FSG_USERNAME and/or FSG_PASSWORD not set.")
        print("The script will open a browser and you will need to log in manually.")
        manual_confirm = input("Type 'YES' to continue in manual login mode, or anything else to abort: ").strip()
        if manual_confirm != "YES":
            log("Aborted by user: credentials missing and manual login declined.", "ERROR")
            sys.exit(1)

    # Final pre-upload confirmation to prevent accidental uploads.
    print(f"\nReady to upload to TEAM_ID={TEAM_ID}. Parts to upload: {len(filtered)}")
    print(f"TEST_MODE={TEST_MODE} DRY_RUN={DRY_RUN}")
    confirm = input("Type 'YES' to proceed with uploading (or anything else to abort): ").strip()
    if confirm != "YES":
        log("Aborted by user before upload.", "WARN")
        sys.exit(0)

    # ── 4. Browser automation ────────────────────────────────────────────
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Login
        if FSG_USERNAME and FSG_PASSWORD:
            log(f"Logging in as '{FSG_USERNAME}'...")
            page.goto(LOGIN_URL)
            page.fill("#tx-felogin-input-username", FSG_USERNAME)
            page.fill("#tx-felogin-input-password", FSG_PASSWORD)
            page.click('input[name="submit"]')
            page.wait_for_load_state("networkidle")

        page.goto(BOM_URL)
        input(
            "\n  ┌──────────────────────────────────────────────────────┐\n"
            "  │  Verify you are logged in and on the BOM page.     │\n"
            "  │  Press ENTER to begin uploading.                   │\n"
            "  └──────────────────────────────────────────────────────┘\n"
        )

        # ── 5. Deduplication ─────────────────────────────────────────────
        log("Fetching existing parts for deduplication...")
        existing: set[str] = set()
        try:
            data = page.evaluate(
                """() => {
                    try {
                        return $('#bom-table').DataTable().data().toArray();
                    } catch(e) {
                        return [];
                    }
                }"""
            )
            for r in data:
                if isinstance(r, dict):
                    key = (
                        f"{str(r.get('system','')).strip()}_"
                        f"{str(r.get('assembly','')).strip()}_"
                        f"{str(r.get('part','')).strip()}"
                    ).lower()
                    existing.add(key)
            log(f"Found {len(existing)} existing parts on the website.")
        except Exception as e:
            log(f"Could not read existing parts: {e}", "WARN")

        # ── 6. Upload loop ───────────────────────────────────────────────
        success = 0
        failed = 0
        skipped_dup = 0
        start_time = time.time()

        for item in filtered:
            sys_code = item["system"]
            asm_raw = item["assembly"]
            part_name = item["part"]
            row_num = item["row"]

            # Deduplication check
            dup_key = f"{sys_code}_{asm_raw}_{part_name}".lower()
            if dup_key in existing:
                log(f"Row {row_num}: Duplicate — '{part_name}' already exists",
                    "SKIP")
                skipped_dup += 1
                continue

            try:
                # Open "New" form
                page.get_by_text("New", exact=True).click()
                page.wait_for_selector(".DTE_Action_Create")

                # System
                sys_label = SYSTEM_MAP.get(sys_code, sys_code)
                if not fuzzy_select(page, "#DTE_Field_system", sys_label):
                    raise RuntimeError(
                        f"System '{sys_label}' not found in dropdown"
                    )
                page.locator("#DTE_Field_system").dispatch_event("change")
                page.wait_for_timeout(1000)

                # Assembly
                if not fuzzy_select(page, "#DTE_Field_assembly", asm_raw):
                    raise RuntimeError(
                        f"Assembly '{asm_raw}' not found in dropdown"
                    )
                page.locator("#DTE_Field_assembly").dispatch_event("change")

                # Part name
                page.locator("#DTE_Field_part").fill(part_name)

                # Make / Buy
                if item["makebuy"] == "m":
                    page.locator("#DTE_Field_makebuy_0").check()
                else:
                    page.locator("#DTE_Field_makebuy_1").check()

                # Optional fields
                if item["comments"]:
                    page.locator("#DTE_Field_comments").fill(item["comments"])
                if item["quantity"]:
                    page.locator("#DTE_Field_quantity").fill(item["quantity"])

                # Submit (or simulate when DRY_RUN enabled)
                if DRY_RUN:
                    log(f"Row {row_num}: [DRY RUN] Would create '{part_name}'", "DRY")
                    existing.add(dup_key)
                    success += 1
                else:
                    page.get_by_text("Create", exact=True).click()
                    page.wait_for_selector(
                        ".DTE_Action_Create", state="hidden", timeout=10000
                    )

                    log(f"Row {row_num}: ✓ '{part_name}'", "OK")
                    existing.add(dup_key)  # prevent re-upload in same run
                    success += 1

            except Exception as e:
                log(f"Row {row_num}: ✗ '{part_name}' — {e}", "ERROR")
                failed += 1
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                except Exception:
                    pass

        # ── 7. Summary ───────────────────────────────────────────────────
        elapsed = round(time.time() - start_time, 1)
        log("─" * 60)
        log(f"Done in {elapsed}s — "
            f"{success} uploaded / {skipped_dup} duplicates / {failed} failed")
        log("─" * 60)

        input("\nPress ENTER to close the browser...")
        browser.close()


if __name__ == "__main__":
    main()
