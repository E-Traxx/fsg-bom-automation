"""Environment variables for the FSG CCBOM automation tools.

All values are loaded from `.env` via python-dotenv. See `.env.example` for
the full list of supported options.
"""

import os
import sys

from dotenv import load_dotenv

load_dotenv()

TEAM_ID = os.getenv("TEAM_ID", "").strip()
if not TEAM_ID:
    print("ERROR: TEAM_ID not set. Copy .env.example to .env and set TEAM_ID.")
    sys.exit(1)

BASE_URL = "https://www.formulastudent.de"
LOGIN_URL = f"{BASE_URL}/login"
BOM_URL = f"{BASE_URL}/teams/fse/details/bom/tid/{TEAM_ID}"

FSG_USERNAME = os.getenv("FSG_USERNAME")
FSG_PASSWORD = os.getenv("FSG_PASSWORD")

TEST_MODE = os.getenv("TEST_MODE", "true").lower() == "true"
DRY_RUN = os.getenv("DRY_RUN", "false").lower() == "true"
DRY_RUN_HOLD_MS = int(os.getenv("DRY_RUN_HOLD_MS", "1500"))
TEST_LIMIT = int(os.getenv("TEST_LIMIT", "3"))

DEFAULT_SYSTEM = os.getenv("DEFAULT_SYSTEM", "").strip().upper()
LOG_FILE = os.getenv("LOG_FILE", "bom_log.txt")
BOMS_DIR = os.getenv("BOMS_DIR", "BOMs")

DEFAULT_FILE = os.getenv("ETRAXX_FILE", "BOM_Final.xlsx")
REQUIRE_INSTALLED = os.getenv("ETRAXX_REQUIRE_INSTALLED", "false").lower() == "true"
