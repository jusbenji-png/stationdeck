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


def load_station_config(station_id: str) -> dict:
    yaml_path = STATIONS_DIR / f"{station_id}.yaml"

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
        resolved_files[key] = DATA_INPUT_DIR / filename

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