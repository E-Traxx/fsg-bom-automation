"""
FSG CCBOM Automation Tool — e-traxx variant (logic module)
==========================================================

Variant of bom_automation.py adapted to the e-traxx master BOM schema
(BOMs/BOM_Final.xlsx), which differs from the EFRxx template:

  * Header row is row 2 (row 1 is a category banner).
  * "System" contains full names ("Chassis and Body"), not codes.
  * "Make/Buy" uses the words "Make" / "Buy".
  * Part column is "Part Name"; comments column is "Comment".
  * A boolean column "if Make CCBOM Eintrag erstellt?" marks already-
    uploaded parts — treated equivalently to a green row.
  * An "Eingebaut?" boolean gates whether a part is actually on the car;
    set ETRAXX_REQUIRE_INSTALLED=true in .env to only upload installed parts.

All environment configuration lives in `src/env.py`; this module contains
the constants, helpers, and main automation loop.
"""

import os
import sys
import glob
import time
import pandas as pd
import openpyxl
from datetime import datetime
from playwright.sync_api import sync_playwright

from src.env import (
    TEAM_ID, LOGIN_URL, BOM_URL,
    FSG_USERNAME, FSG_PASSWORD,
    TEST_MODE, DRY_RUN, DRY_RUN_HOLD_MS, TEST_LIMIT,
    DEFAULT_SYSTEM, LOG_FILE, BOMS_DIR,
    DEFAULT_FILE, REQUIRE_INSTALLED,
)

# ──────────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────────

SKIP_COLORS = {
    "FF00FF00", "0000FF00",   # Green — already uploaded
    "FFFF0000", "00FF0000",   # Red   — do not upload
}

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

# Full name (as written in BOM_Final.xlsx) → system code
SYSTEM_NAME_TO_CODE = {
    "autonomous system":                    "AT",
    "brake system":                         "BR",
    "drivetrain":                           "DT",
    "engine and tractive system":           "ET",
    "chassis and body":                     "FR",
    "grounded low voltage system":          "LV",
    "miscellaneous fit and finish":         "MS",
    "steering system":                      "ST",
    "suspension system":                    "SU",
    "wheels, wheel bearings and tires":     "WT",
    "wheels wheel bearings and tires":      "WT",
}

ASSEMBLY_REMAP = {
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
    "damper":               "Dampers",
    "spring":               "Springs",
    "pushrod":              "Pushrods",
    "rocker":               "Rockers",
    "a-arm":                "A-Arms",
    "chain":                "Chain",
    "sprocket":             "Sprockets",
    "differential":         "Differential",
    "half shaft":           "Half Shafts",
    "halfshaft":            "Half Shafts",
    "steering rack":        "Steering Rack",
    "tie rod":              "Tie Rods",
    "steering wheel":       "Steering Wheel",
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
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] [{status}] {message}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ──────────────────────────────────────────────────────────────────────────────
# Excel helpers
# ──────────────────────────────────────────────────────────────────────────────

def get_cell_color(sheet, row: int, col: int = 1) -> str | None:
    try:
        fill = sheet.cell(row=row, column=col).fill
        if fill.patternType is None:
            return None
        idx = fill.start_color.index
        return str(idx).upper() if idx and idx != "00000000" else None
    except Exception:
        return None


def should_skip_color(sheet, row: int) -> str | None:
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


def as_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    s = str(v).strip().lower()
    return s in ("true", "1", "yes", "x", "ja")


# ──────────────────────────────────────────────────────────────────────────────
# Fuzzy dropdown matching
# ──────────────────────────────────────────────────────────────────────────────

def snapshot_options(page, selector: str) -> list[str]:
    try:
        return page.eval_on_selector(
            selector,
            "el => Array.from(el.options).map(o => o.text)",
        )
    except Exception:
        return []


def wait_for_options(page, selector: str, expected: str | None = None,
                     timeout_ms: int = 5000, poll_ms: int = 100,
                     previous: list[str] | None = None) -> list[str]:
    # poll the <select> until the AJAX-driven option list settles.
    # `previous` is the stale list from the prior row — we ignore it until
    # the dropdown has actually repopulated, otherwise a stale "Brackets"
    # can satisfy the match and then vanish before fuzzy_select runs.
    resolved = ASSEMBLY_REMAP.get(expected.lower().strip(), expected) if expected else None
    resolved_lower = resolved.lower().strip() if resolved else None
    deadline = time.time() + timeout_ms / 1000
    options: list[str] = []
    last_snapshot: list[str] | None = None
    stable_hits = 0
    while time.time() < deadline:
        options = snapshot_options(page, selector)
        # stale list from previous row — keep waiting for the refresh
        if previous is not None and options == previous:
            last_snapshot = None
            stable_hits = 0
            page.wait_for_timeout(poll_ms)
            continue
        real = [o for o in options if o and o.strip()]
        populated = real and (
            len(real) > 1 or real[0].strip().lower() not in ("", "select", "-")
        )
        if populated:
            match_ok = resolved_lower is None or any(
                ol == resolved_lower or resolved_lower in ol or ol in resolved_lower
                for ol in (o.lower().strip() for o in real)
            )
            if match_ok:
                # require two consecutive identical snapshots — confirms AJAX settled
                if options == last_snapshot:
                    stable_hits += 1
                    if stable_hits >= 1:
                        return options
                else:
                    last_snapshot = options
                    stable_hits = 0
        page.wait_for_timeout(poll_ms)
    return options


def fuzzy_select(page, selector: str, target: str) -> bool:
    resolved = ASSEMBLY_REMAP.get(target.lower().strip(), target)
    resolved_lower = resolved.lower().strip()

    def pick(options: list[str]) -> str | None:
        if resolved in options:
            return resolved
        for opt in options:
            if opt.lower().strip() == resolved_lower:
                return opt
        for opt in options:
            ol = opt.lower().strip()
            if resolved_lower in ol or ol in resolved_lower:
                return opt
        return None

    # retry once: the options list can change underfoot when AJAX repopulates
    # between our read and the select_option call.
    for attempt in range(2):
        try:
            options = snapshot_options(page, selector)
            choice = pick(options)
            if choice is None:
                if attempt == 0:
                    page.wait_for_timeout(200)
                    continue
                return False
            page.locator(selector).select_option(label=choice)
            return True
        except Exception as e:
            if attempt == 0:
                page.wait_for_timeout(200)
                continue
            log(f"Fuzzy-match error for '{target}': {e}", "WARN")
            return False
    return False


# ──────────────────────────────────────────────────────────────────────────────
# File selection
# ──────────────────────────────────────────────────────────────────────────────

def discover_excel_files() -> list[str]:
    search_dir = os.path.join(os.getcwd(), BOMS_DIR)
    if not os.path.isdir(search_dir):
        os.makedirs(search_dir, exist_ok=True)
        log(f"Created '{BOMS_DIR}/' directory. Place your Excel files there.", "WARN")
        return []
    return sorted(glob.glob(os.path.join(search_dir, "*.xlsx")))


def select_file() -> str:
    # Prefer the configured default file if present.
    default_path = os.path.join(os.getcwd(), BOMS_DIR, DEFAULT_FILE)
    if os.path.isfile(default_path):
        log(f"Using default e-traxx file: {DEFAULT_FILE}")
        return default_path

    files = discover_excel_files()
    if not files:
        print(f"\n  No .xlsx files found in '{BOMS_DIR}/'.\n")
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
    log("FSG CCBOM Automation (e-traxx) — Starting")
    if TEST_MODE:
        log(f"TEST MODE: Only the first {TEST_LIMIT} parts will be processed.", "WARN")
    log(f"Config: TEAM_ID={TEAM_ID} TEST_MODE={TEST_MODE} DRY_RUN={DRY_RUN} "
        f"BOMS_DIR={BOMS_DIR} REQUIRE_INSTALLED={REQUIRE_INSTALLED}")

    # ── 1. File ──────────────────────────────────────────────────────────
    filepath = select_file()
    filename = os.path.basename(filepath)
    log(f"Selected file: {filename}")

    wb = openpyxl.load_workbook(filepath, data_only=True)
    if "BOM" in wb.sheetnames:
        sheet = wb["BOM"]
    else:
        sheet = wb.active
        log(f"Sheet 'BOM' not found — using active sheet '{sheet.title}'.", "WARN")

    # Header row is row 2 in the e-traxx template (row 1 is a banner).
    df = pd.read_excel(filepath, sheet_name=sheet.title, header=1)
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Map the e-traxx column names to the internal keys we use below.
    COLMAP = {
        "system":       "system",
        "assembly":     "assembly",
        "sub-assembly": "subassembly",
        "part name":    "part",
        "make/buy":     "makebuy",
        "comment":      "comments",
        "quantity":     "quantity",
        "eingebaut?":                       "installed",
        "if make ccbom eintrag erstellt?":  "uploaded",
    }
    missing = [c for c in ("system", "assembly", "part name") if c not in df.columns]
    if missing:
        log(f"Missing required columns: {missing}. Available: {list(df.columns)}", "ERROR")
        sys.exit(1)

    # Header is on row 2 → first data row is Excel row 3.
    HEADER_EXCEL_ROW = 2

    def g(row, key, default=""):
        col = next((c for c, k in COLMAP.items() if k == key), None)
        if col is None or col not in row.index:
            return default
        return row[col]

    # ── 2. System selection ──────────────────────────────────────────────
    unique_codes: list[str] = []
    seen = set()
    for s in df["system"].dropna().unique():
        code = SYSTEM_NAME_TO_CODE.get(str(s).strip().lower())
        if code and code not in seen:
            seen.add(code)
            unique_codes.append(code)

    if DEFAULT_SYSTEM and DEFAULT_SYSTEM in unique_codes:
        run_system = DEFAULT_SYSTEM
        log(f"System filter (from .env): {run_system} — "
            f"{SYSTEM_MAP.get(run_system, run_system)}")
    else:
        print("\nSystems found in Excel:")
        for s in unique_codes:
            print(f"  • {s:4s}  {SYSTEM_MAP.get(s, s)}")
        run_system = input(
            "\nEnter system code to filter (e.g. 'BR') or 'ALL' for everything: "
        ).strip().upper()

    # ── 3. Filter rows ───────────────────────────────────────────────────
    filtered = []
    skipped_green = 0
    skipped_red = 0
    skipped_example = 0
    skipped_empty = 0
    skipped_uploaded = 0
    skipped_not_installed = 0

    for idx, row in df.iterrows():
        excel_row = idx + HEADER_EXCEL_ROW + 1

        sys_name = str(g(row, "system", "")).strip()
        sys_code = SYSTEM_NAME_TO_CODE.get(sys_name.lower(), "")
        assembly = str(g(row, "assembly", "")).strip()
        subassembly = str(g(row, "subassembly", "")).strip()
        if subassembly.lower() in ("nan", ""):
            subassembly = ""
        part = str(g(row, "part", "")).strip()
        quantity = g(row, "quantity", "")
        makebuy = str(g(row, "makebuy", "")).strip().lower()
        comments = str(g(row, "comments", "")).strip()
        installed = as_bool(g(row, "installed", False))
        uploaded = as_bool(g(row, "uploaded", False))

        if run_system != "ALL" and sys_code != run_system:
            continue

        if not sys_code or not part or part.lower() in ("nan", ""):
            skipped_empty += 1
            continue

        if "BEISPIEL" in part.upper() or "EXAMPLE" in part.upper():
            skipped_example += 1
            continue

        if uploaded:
            skipped_uploaded += 1
            log(f"Row {excel_row}: Skipped — already uploaded (CCBOM Eintrag erstellt)", "SKIP")
            continue

        if REQUIRE_INSTALLED and not installed:
            skipped_not_installed += 1
            continue

        skip_reason = should_skip_color(sheet, excel_row)
        if skip_reason:
            if "green" in skip_reason:
                skipped_green += 1
            elif "red" in skip_reason:
                skipped_red += 1
            log(f"Row {excel_row}: Skipped — {skip_reason}", "SKIP")
            continue

        mb = makebuy[0] if makebuy and makebuy[0] in ("m", "b") else "m"

        if comments.lower() in ("nan", ""):
            comments = ""

        qty_str = str(quantity).strip()
        if qty_str.lower() in ("nan", ""):
            qty_str = ""
        elif qty_str.endswith(".0"):
            qty_str = qty_str[:-2]

        filtered.append({
            "row": excel_row,
            "system": sys_code,
            "assembly": assembly,
            "subassembly": subassembly,
            "part": part,
            "makebuy": mb,
            "quantity": qty_str,
            "comments": comments,
        })

    log(f"Filtering complete: {len(filtered)} parts to upload "
        f"({skipped_green} green / {skipped_red} red / "
        f"{skipped_uploaded} already uploaded / "
        f"{skipped_not_installed} not installed / "
        f"{skipped_example} example / {skipped_empty} empty skipped)")

    if TEST_MODE and len(filtered) > TEST_LIMIT:
        log(f"Test Mode: limiting {len(filtered)} → {TEST_LIMIT} parts")
        filtered = filtered[:TEST_LIMIT]

    if not filtered:
        log("Nothing to upload — exiting.")
        sys.exit(0)

    if not (FSG_USERNAME and FSG_PASSWORD):
        print("\nWARNING: FSG_USERNAME and/or FSG_PASSWORD not set.")
        print("The script will open a browser and you will need to log in manually.")
        manual_confirm = input("Type 'YES' to continue in manual login mode, or anything else to abort: ").strip()
        if manual_confirm != "YES":
            log("Aborted by user: credentials missing and manual login declined.", "ERROR")
            sys.exit(1)

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

        log("Fetching existing parts for deduplication...")
        existing: set[str] = set()
        try:
            data = page.evaluate(
                """() => {
                    try { return $('#bom-table').DataTable().data().toArray(); }
                    catch(e) { return []; }
                }"""
            )
            for r in data:
                if isinstance(r, dict):
                    key = (
                        f"{str(r.get('system','')).strip()}_"
                        f"{str(r.get('assembly','')).strip()}_"
                        f"{str(r.get('subassembly','')).strip()}_"
                        f"{str(r.get('part','')).strip()}"
                    ).lower()
                    existing.add(key)
            log(f"Found {len(existing)} existing parts on the website.")
        except Exception as e:
            log(f"Could not read existing parts: {e}", "WARN")

        success = 0
        failed = 0
        skipped_dup = 0
        start_time = time.time()

        for item in filtered:
            sys_code = item["system"]
            asm_raw = item["assembly"]
            sub_raw = item["subassembly"]
            part_name = item["part"]
            row_num = item["row"]

            dup_key = f"{sys_code}_{asm_raw}_{sub_raw}_{part_name}".lower()
            if dup_key in existing:
                log(f"Row {row_num}: Duplicate — '{part_name}' already exists", "SKIP")
                skipped_dup += 1
                continue

            try:
                page.get_by_text("New", exact=True).click()
                page.wait_for_selector(".DTE_Action_Create")

                sys_label = SYSTEM_MAP.get(sys_code, sys_code)
                asm_before = snapshot_options(page, "#DTE_Field_assembly")
                if not fuzzy_select(page, "#DTE_Field_system", sys_label):
                    raise RuntimeError(f"System '{sys_label}' not found in dropdown")
                page.locator("#DTE_Field_system").dispatch_event("change")
                wait_for_options(page, "#DTE_Field_assembly", asm_raw,
                                 previous=asm_before)

                if not fuzzy_select(page, "#DTE_Field_assembly", asm_raw):
                    raise RuntimeError(f"Assembly '{asm_raw}' not found in dropdown")
                sub_before = snapshot_options(page, "#DTE_Field_subassembly")
                page.locator("#DTE_Field_assembly").dispatch_event("change")

                if sub_raw and page.locator("#DTE_Field_subassembly").count():
                    wait_for_options(page, "#DTE_Field_subassembly", sub_raw,
                                     previous=sub_before)
                    if not fuzzy_select(page, "#DTE_Field_subassembly", sub_raw):
                        log(f"Row {row_num}: Sub-assembly '{sub_raw}' not in dropdown — leaving blank", "WARN")
                    else:
                        page.locator("#DTE_Field_subassembly").dispatch_event("change")

                page.locator("#DTE_Field_part").fill(part_name)

                if item["makebuy"] == "m":
                    page.locator("#DTE_Field_makebuy_0").check()
                else:
                    page.locator("#DTE_Field_makebuy_1").check()

                if item["comments"]:
                    page.locator("#DTE_Field_comments").fill(item["comments"])
                if item["quantity"]:
                    page.locator("#DTE_Field_quantity").fill(item["quantity"])

                if DRY_RUN:
                    log(f"Row {row_num}: [DRY RUN] Would create '{part_name}'", "DRY")
                    existing.add(dup_key)
                    success += 1
                    # hold the filled form so the user can eyeball it before moving on
                    page.wait_for_timeout(DRY_RUN_HOLD_MS)
                    # close the modal so the next row starts from a clean form
                    try:
                        cancel = page.get_by_text("Cancel", exact=True)
                        if cancel.count():
                            cancel.first.click()
                        else:
                            page.keyboard.press("Escape")
                        page.wait_for_selector(".DTE_Action_Create", state="hidden", timeout=3000)
                    except Exception:
                        pass
                else:
                    page.get_by_text("Create", exact=True).click()
                    page.wait_for_selector(
                        ".DTE_Action_Create", state="hidden", timeout=10000
                    )
                    log(f"Row {row_num}: ✓ '{part_name}'", "OK")
                    existing.add(dup_key)
                    success += 1

            except Exception as e:
                log(f"Row {row_num}: ✗ '{part_name}' — {e}", "ERROR")
                failed += 1
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(500)
                except Exception:
                    pass

        elapsed = round(time.time() - start_time, 1)
        log("─" * 60)
        log(f"Done in {elapsed}s — {success} uploaded / {skipped_dup} duplicates / {failed} failed")
        log("─" * 60)

        input("\nPress ENTER to close the browser...")
        browser.close()
