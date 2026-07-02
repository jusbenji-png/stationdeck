# ============================================================
# config/station_loader.py
# StationDeck — Station Configuration Loader
# ============================================================

import os
import sys
import yaml
from pathlib import Path

# --- Path detection: works both frozen (.exe) and in development ---
if getattr(sys, 'frozen', False):
    PROJECT_ROOT = Path(sys.executable).parent          # read-only install dir
    _appdata = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    DATA_ROOT = (Path(_appdata) / "StationDeck") if _appdata else PROJECT_ROOT
else:
    PROJECT_ROOT = Path(__file__).resolve().parent.parent
    DATA_ROOT = PROJECT_ROOT

# Station YAML is a read-only bundled asset → install dir.
STATIONS_DIR = PROJECT_ROOT / "config" / "stations"
# Input data and report output are WRITABLE → DATA_ROOT (%LOCALAPPDATA% when frozen).
DATA_INPUT_DIR = DATA_ROOT / "data" / "input"
REPORTS_DIR = DATA_ROOT / "reports"


# Uploaded files have been saved under different names across versions
# (e.g. "Stock mvt-Template-2025-2026-V1-7-25.xlsx" vs "Stock_mvt.xlsx"),
# and station YAMLs written at different times expect different ones. If the
# configured filename isn't on disk, fall back to the newest file in
# data/input whose name matches the file kind — so imports keep feeding
# reports and exports regardless of which name they were saved under.
_INPUT_FILE_PATTERNS = {
    "cashflow": ("cash",),
    "stock":    ("stock", "mvt", "movement"),
    "shop":     ("shop",),
    "manager":  ("manager",),
}


def _resolve_input_file(key: str, filename: str):
    exact = DATA_INPUT_DIR / filename
    if exact.exists():
        return exact
    patterns = _INPUT_FILE_PATTERNS.get(key)
    if patterns and DATA_INPUT_DIR.exists():
        candidates = [
            f for f in DATA_INPUT_DIR.glob("*.xls*")
            if any(p in f.name.lower() for p in patterns)
        ]
        if candidates:
            return max(candidates, key=lambda f: f.stat().st_mtime)
    return exact


def load_station_config(station_id: str) -> dict:
    # Check writable user dir first (stations registered at runtime),
    # then fall back to bundled install dir (te_rwizi and pre-built configs).
    user_yaml    = DATA_ROOT / "stations" / f"{station_id}.yaml"
    bundled_yaml = STATIONS_DIR / f"{station_id}.yaml"
    yaml_path    = user_yaml if user_yaml.exists() else bundled_yaml

    if not yaml_path.exists():
        raise FileNotFoundError(
            f"No config found for station '{station_id}'.\n"
            f"Expected: {yaml_path}\n"
            f"Available configs: {_list_available_stations()}"
        )

    with open(yaml_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    required_sections = ["station", "files", "reports", "email", "branding", "settings"]
    for section in required_sections:
        if section not in raw:
            raise KeyError(
                f"Missing required section '{section}' in {yaml_path.name}"
            )

    resolved_files = {}
    for key, filename in raw["files"].items():
        resolved_files[key] = _resolve_input_file(key, filename)

    station_output = raw["reports"]["output_dir"]
    resolved_reports = {
        "pdf":     REPORTS_DIR / station_output / "pdf",
        "docx":    REPORTS_DIR / station_output / "docx",
        "xlsx":    REPORTS_DIR / station_output / "xlsx",
        "archive": REPORTS_DIR / station_output / "archive",
    }

    config = {
        "station_id":   raw["station"]["id"],
        "station_name": raw["station"]["name"],
        "location":     raw["station"]["location"],
        "operator":     raw["station"]["operator"],
        "currency":     raw["station"]["currency"],
        "timezone":     raw["station"]["timezone"],
        "files":        resolved_files,
        "report_dirs":  resolved_reports,
        "recipients":   raw["email"]["recipients"],
        "branding":     raw["branding"],
        "placeholder_mode": raw["settings"]["placeholder_mode"],
        "archive_enabled":  raw["settings"]["archive_enabled"],
        "email_enabled":    raw["settings"]["email_enabled"],
    }

    return config


def _list_available_stations() -> list:
    if not STATIONS_DIR.exists():
        return []
    return [f.stem for f in STATIONS_DIR.glob("*.yaml")]


if __name__ == "__main__":
    from pprint import pprint
    try:
        config = load_station_config("te_rwizi")
        print("Config loaded successfully!")
        pprint(config)
    except Exception as e:
        print(f"Error: {e}")