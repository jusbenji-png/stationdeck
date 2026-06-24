"""
src/updater.py — StationDeck Auto-Update Checker
-------------------------------------------------

WHY THIS EXISTS:
  Once StationDeck is installed at real stations, Benjamin needs to push
  updates (bug fixes, new features, payment integration, etc.) without
  visiting each station manually. This module checks whether a newer
  version exists every time the app starts, then tells the dashboard
  to show an update banner if one is available.

HOW IT WORKS:
  1. On startup, launcher.py calls start_update_check() in a background thread
  2. That thread calls GET /version on the Railway auth server
  3. The auth server returns the latest version number + download URL
  4. We compare that to the installed version in config/version.txt
  5. The result is stored in a module-level variable (_update_status)
  6. The Flask route /check_update reads that variable and returns JSON
  7. dashboard.html calls /check_update on page load and shows a banner if needed

WHAT THE STATION SEES:
  A dismissable yellow banner at the top of the dashboard:
  "StationDeck v1.1.0 is available. [Download Update]"
  Clicking opens their browser to the GitHub Releases download page.
  They download the new installer, run it, done.

DESIGN DECISIONS:
  - Background thread: update check never slows down app startup
  - Fail silently: if the auth server is unreachable, no banner, no error
  - One check per startup: we don't poll repeatedly (not needed)
  - No automatic download: station clicks and runs installer themselves
    (safer on Windows — avoids file lock and permission issues)
"""

import threading
import requests
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# ── Auth server URL (same server used for login verification) ─────────────────
AUTH_SERVER = "https://web-production-46077.up.railway.app"

# ── Module-level result store ─────────────────────────────────────────────────
# This dictionary is written by the background thread and read by Flask routes.
# It starts as None (check not yet complete) and is filled within a few seconds.
#
# Possible states:
#   {"status": "pending"}                          — check still running
#   {"status": "up_to_date", "version": "1.0.0"}  — already on latest
#   {"status": "update_available",                 — newer version exists
#    "installed_version": "1.0.0",
#    "latest_version":    "1.1.0",
#    "download_url":      "https://...",
#    "release_notes":     "What's new..."}
#   {"status": "error"}                            — server unreachable or error

_update_status: dict = {"status": "pending"}
_check_done = threading.Event()


# ── Version file reader ───────────────────────────────────────────────────────

def get_installed_version() -> str:
    """
    Read the installed version from config/version.txt.

    This file lives at stationdeck/config/version.txt and contains
    a single line like: 1.0.0

    We use sys.frozen-aware path detection so this works both when
    running as a .py file and when bundled as a .exe by PyInstaller.
    """
    import sys

    if getattr(sys, 'frozen', False):
        # Running as .exe — executable is in the install folder
        base = Path(sys.executable).parent
    else:
        # Running as .py — go up from src/ to project root
        base = Path(__file__).resolve().parent.parent

    version_file = base / "config" / "version.txt"

    try:
        version = version_file.read_text(encoding="utf-8").strip()
        if version:
            return version
    except Exception as e:
        logger.warning(f"Could not read version.txt: {e}")

    # Fallback — should never happen in a properly built installer
    return "1.0.0"


# ── Version comparison ────────────────────────────────────────────────────────

def _version_tuple(v: str) -> tuple:
    """
    Convert a version string like "1.2.3" into a comparable tuple (1, 2, 3).

    This lets us compare versions correctly:
      "1.10.0" > "1.9.0"  ← string comparison would get this wrong
      (1,10,0) > (1,9,0)  ← tuple comparison gets it right
    """
    try:
        return tuple(int(x) for x in v.strip().split("."))
    except Exception:
        return (0, 0, 0)


def is_newer(latest: str, installed: str) -> bool:
    """Return True if latest version is newer than installed version."""
    return _version_tuple(latest) > _version_tuple(installed)


# ── Background check thread ───────────────────────────────────────────────────

def _do_check():
    """
    Called in a background thread by start_update_check().

    Fetches /version from the auth server, compares to installed version,
    writes result into _update_status. Never raises — all errors are caught
    so a network problem never crashes the app.
    """
    global _update_status

    installed = get_installed_version()
    logger.info(f"Update check: installed version is {installed}")

    try:
        response = requests.get(
            f"{AUTH_SERVER}/version",
            timeout=8,          # don't hang startup for more than 8 seconds
            headers={"User-Agent": f"StationDeck/{installed}"},
        )

        if response.status_code != 200:
            logger.warning(f"Update check: server returned {response.status_code}")
            _update_status = {"status": "error"}
            _check_done.set()
            return

        data = response.json()
        latest        = data.get("version", "")
        download_url  = data.get("download_url", "")
        release_notes = data.get("release_notes", "")

        if not latest:
            logger.warning("Update check: server returned empty version")
            _update_status = {"status": "error"}
            _check_done.set()
            return

        if is_newer(latest, installed):
            logger.info(f"Update available: {installed} → {latest}")
            _update_status = {
                "status":            "update_available",
                "installed_version": installed,
                "latest_version":    latest,
                "download_url":      download_url,
                "release_notes":     release_notes,
            }
        else:
            logger.info(f"Update check: already up to date ({installed})")
            _update_status = {
                "status":  "up_to_date",
                "version": installed,
            }

    except requests.exceptions.ConnectionError:
        logger.info("Update check: auth server unreachable (offline) — skipping")
        _update_status = {"status": "error"}

    except requests.exceptions.Timeout:
        logger.warning("Update check: timed out after 8s — skipping")
        _update_status = {"status": "error"}

    except Exception as e:
        logger.warning(f"Update check failed: {e}")
        _update_status = {"status": "error"}

    finally:
        _check_done.set()


# ── Public API ────────────────────────────────────────────────────────────────

def start_update_check():
    """
    Launch the update check in a background daemon thread.

    Call this once from launcher.py after Flask starts.
    It returns immediately — the check runs in the background.
    The result appears in _update_status within a few seconds.
    """
    thread = threading.Thread(target=_do_check, daemon=True, name="UpdateChecker")
    thread.start()
    logger.info("Update check started in background.")


def get_update_status() -> dict:
    """
    Return the current update status dictionary.

    Called by the Flask /check_update route.
    If the background check hasn't finished yet, returns {"status": "pending"}.
    """
    return _update_status.copy()