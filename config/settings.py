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
    # Running as compiled .exe — use the folder containing the .exe
    BASE_DIR = Path(sys.executable).parent
else:
    # Running in development — settings.py is at config/settings.py
    # so parent.parent = project root
    BASE_DIR = Path(__file__).parent.parent

# Load secret keys from .env file (must come after BASE_DIR is set)
load_dotenv(BASE_DIR / ".env")

# =============================================================
# PROJECT PATHS
# =============================================================
DATA_INPUT_DIR      = BASE_DIR / "data" / "input"
DATA_PROCESSED_DIR  = BASE_DIR / "data" / "processed"
DATA_ARCHIVE_DIR    = BASE_DIR / "data" / "archive"
REPORTS_PDF_DIR     = BASE_DIR / "reports" / "pdf"
REPORTS_DOCX_DIR    = BASE_DIR / "reports" / "docx"
REPORTS_HISTORY_DIR = BASE_DIR / "reports" / "history"
LOGS_DIR            = BASE_DIR / "logs"
TEMPLATES_DIR       = BASE_DIR / "templates"

# Shorthand aliases used by processor.py and main.py
REPORTS_DIR         = BASE_DIR / "reports"
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