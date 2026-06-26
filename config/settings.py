# =============================================================
# StationDeck - Configuration Settings
# =============================================================
# This file is the single source of truth for all settings.
# Change values here and they update across the entire system.
# =============================================================
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# =============================================================
# BASE DIRECTORY DETECTION
# =============================================================
# When running as a PyInstaller .exe, __file__ points inside
# the _internal bundle folder, NOT the install folder.
# We must detect this and use sys.executable's parent instead.
#
# Development:   BASE_DIR = stationdeck/          (project root)
# Frozen .exe:   BASE_DIR = dist/StationDeck/     (exe folder)

if getattr(sys, 'frozen', False):
    # Running as compiled .exe — BASE_DIR is the (read-only) install folder.
    BASE_DIR = Path(sys.executable).parent
    # Writable data MUST NOT live under Program Files (non-admin users cannot
    # write there) and must survive app updates. Use %LOCALAPPDATA%\StationDeck.
    _appdata = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    DATA_DIR = (Path(_appdata) / "StationDeck") if _appdata else BASE_DIR
else:
    # Running in development — settings.py is at config/settings.py
    # so parent.parent = project root. Data stays in the project root.
    BASE_DIR = Path(__file__).parent.parent
    DATA_DIR = BASE_DIR

# Load secret keys from .env file (read-only — lives next to the install/dev root)
load_dotenv(BASE_DIR / ".env")

# =============================================================
# PROJECT PATHS
# =============================================================
# WRITABLE locations live under DATA_DIR; read-only bundled assets under BASE_DIR.
DATA_INPUT_DIR      = DATA_DIR / "data" / "input"
DATA_PROCESSED_DIR  = DATA_DIR / "data" / "processed"
DATA_ARCHIVE_DIR    = DATA_DIR / "data" / "archive"
REPORTS_PDF_DIR     = DATA_DIR / "reports" / "pdf"
REPORTS_DOCX_DIR    = DATA_DIR / "reports" / "docx"
REPORTS_HISTORY_DIR = DATA_DIR / "reports" / "history"
LOGS_DIR            = DATA_DIR / "logs"
TEMPLATES_DIR       = BASE_DIR / "templates"   # read-only bundled assets

# Shorthand aliases used by processor.py and main.py
REPORTS_DIR         = DATA_DIR / "reports"
ARCHIVE_DIR         = REPORTS_DIR / "archive"

# =============================================================
# STATION INFORMATION
# =============================================================
STATION_NAME        = os.getenv("STATION_NAME", "Total Fuel Station")
STATION_LOCATION    = os.getenv("STATION_LOCATION", "Uganda")

# =============================================================
# OPENAI API SETTINGS
# =============================================================
OPENAI_API_KEY      = os.getenv("OPENAI_API_KEY")
OPENAI_MODEL        = "gpt-4o"   # full model — far better at handwritten digits than -mini
AI_MAX_TOKENS       = 2000
AI_TEMPERATURE      = 0.3

# =============================================================
# REPORT SETTINGS
# =============================================================
REPORT_CURRENCY     = "UGX"
REPORT_AUTHOR       = "StationDeck AI System"
FUEL_PRODUCTS = {
    "PMS": "Petrol (PMS)",
    "AGO": "Diesel (AGO)",
    "LPG": "LPG Gas",
}

# =============================================================
# LOGGING SETTINGS
# =============================================================
LOG_FILE            = LOGS_DIR / "app.log"
LOG_LEVEL           = "INFO"

# =============================================================
# EMAIL / SMTP SETTINGS
# =============================================================
SMTP_HOST           = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT           = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER           = os.getenv("SMTP_USER", "")
SMTP_PASS           = os.getenv("SMTP_PASS", "")

# Comma-separated in .env -> Python list here
# Example in .env: REPORT_RECIPIENTS=manager@station.com,owner@station.com
_recipients_raw     = os.getenv("REPORT_RECIPIENTS", "")
REPORT_RECIPIENTS   = [r.strip() for r in _recipients_raw.split(",") if r.strip()]