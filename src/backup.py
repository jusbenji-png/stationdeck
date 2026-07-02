# =============================================================
# src/backup.py — StationDeck Automatic Data Backup
# =============================================================
#
# All station data lives only on the station PC. Staff were told to
# copy the data folder to a USB drive monthly, which nobody does
# reliably. This runs automatically on startup instead:
#
#   - At most once every 7 days (checked against existing backups)
#   - Zips the database + imported Excel files into
#     %LOCALAPPDATA%\StationDeck\backups\stationdeck_backup_YYYY-MM-DD.zip
#   - Keeps the newest 8 backups (~2 months of weekly history)
#
# A backup on the same disk does not protect against theft or disk
# death — the Settings page still tells stations to copy backups to
# a USB drive or cloud folder — but it does protect against the far
# more common cases: bad imports, accidental Clear All Data, and
# app-upgrade problems.
# =============================================================

import logging
import re
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

logger = logging.getLogger(__name__)

BACKUP_EVERY_DAYS = 7
KEEP_BACKUPS      = 8
_NAME_RE = re.compile(r"stationdeck_backup_(\d{4}-\d{2}-\d{2})\.zip$")


def _backup_dir(data_root: Path) -> Path:
    return data_root / "backups"


def last_backup_date(data_root: Path):
    """Date of the newest backup, or None."""
    bdir = _backup_dir(data_root)
    if not bdir.exists():
        return None
    dates = []
    for f in bdir.glob("stationdeck_backup_*.zip"):
        m = _NAME_RE.search(f.name)
        if m:
            try:
                dates.append(datetime.strptime(m.group(1), "%Y-%m-%d").date())
            except ValueError:
                pass
    return max(dates) if dates else None


def run_backup_if_due(data_root: Path) -> bool:
    """Create a weekly backup if the newest one is older than the interval.

    Returns True if a backup was created. Never raises — a backup failure
    must never break app startup.
    """
    try:
        last = last_backup_date(data_root)
        today = datetime.now().date()
        if last and (today - last) < timedelta(days=BACKUP_EVERY_DAYS):
            return False

        bdir = _backup_dir(data_root)
        bdir.mkdir(parents=True, exist_ok=True)
        out = bdir / f"stationdeck_backup_{today:%Y-%m-%d}.zip"

        db_file   = data_root / "data" / "stationdeck.db"
        input_dir = data_root / "data" / "input"
        stations  = data_root / "stations"

        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            if db_file.exists():
                z.write(db_file, "data/stationdeck.db")
            if input_dir.exists():
                for f in input_dir.glob("*.xls*"):
                    z.write(f, f"data/input/{f.name}")
            if stations.exists():
                for f in stations.glob("*.yaml"):
                    z.write(f, f"stations/{f.name}")
            id_file = data_root / "station_id.txt"
            if id_file.exists():
                z.write(id_file, "station_id.txt")

        # Prune old backups beyond the retention window
        backups = sorted(bdir.glob("stationdeck_backup_*.zip"),
                         key=lambda f: f.name, reverse=True)
        for old in backups[KEEP_BACKUPS:]:
            try:
                old.unlink()
            except Exception:
                pass

        logger.info(f"Automatic backup created: {out.name} "
                    f"({out.stat().st_size:,} bytes)")
        return True

    except Exception as e:
        logger.warning(f"Automatic backup failed (non-fatal): {e}")
        return False
