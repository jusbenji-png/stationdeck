# ============================================================
# main.py
# StationDeck — Production Pipeline Entry Point
#
# Runs the full 6-step monthly report pipeline for a station.
#
# Usage:
#   python main.py --station te_rwizi --month 4 --year 2026
# ============================================================

import sys
import io

# Fix Windows cp1252 encoding — MUST be first
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import argparse
import logging
import shutil
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path

from config.station_loader import load_station_config
from src.processor         import process_monthly_report
from src.ai_engine         import generate_report
from src.exporter          import ExportEngine
from src.emailer           import send_monthly_report

# ============================================================
# LOGGING SETUP
# ============================================================

LOG_DIR = Path("logs")
LOG_DIR.mkdir(exist_ok=True)

log_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

file_handler = RotatingFileHandler(
    LOG_DIR / "app.log",
    maxBytes=5 * 1024 * 1024,  # 5 MB
    backupCount=3,
    encoding="utf-8"
)
file_handler.setFormatter(log_formatter)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_formatter)

logger = logging.getLogger("stationdeck")
logger.setLevel(logging.INFO)
logger.addHandler(file_handler)
logger.addHandler(console_handler)


# ============================================================
# HELPERS
# ============================================================

def build_stamped_stem(period_label: str) -> str:
    """
    Build a filename stem with today's date stamp.
    e.g. "April 2026 20260526"
    """
    today = datetime.now().strftime("%Y%m%d")
    return f"{period_label} {today}"


def archive_reports(pdf: str, docx: str, xlsx: str,
                    archive_dir: Path, period_label: str) -> None:
    """
    Move all 3 exported files into the station's monthly archive folder.
    e.g. reports/te_rwizi/archive/2026-04/
    """
    month_names = {
        "January": "01", "February": "02", "March": "03",    "April": "04",
        "May": "05",      "June": "06",     "July": "07",     "August": "08",
        "September": "09","October": "10",  "November": "11", "December": "12"
    }

    parts = period_label.split()          # ["April", "2026"]
    month_num = month_names.get(parts[0], "00")
    year_num  = parts[1] if len(parts) > 1 else "0000"

    dest = archive_dir / f"{year_num}-{month_num}"
    dest.mkdir(parents=True, exist_ok=True)

    for src_path in [pdf, docx, xlsx]:
        if src_path and Path(src_path).exists():
            shutil.move(src_path, dest / Path(src_path).name)
            logger.info(f"Archived: {Path(src_path).name} → {dest}")


# ============================================================
# PIPELINE
# ============================================================

def run_pipeline(station_id: str, month: int, year: int) -> None:

    start_time = datetime.now()

    month_names = {
        1: "January", 2: "February",  3: "March",    4: "April",
        5: "May",     6: "June",       7: "July",     8: "August",
        9: "September", 10: "October", 11: "November", 12: "December"
    }
    period_label = f"{month_names[month]} {year}"

    logger.info("=" * 55)
    logger.info(f"  StationDeck Pipeline Started")
    logger.info(f"  Station : {station_id}")
    logger.info(f"  Period  : {period_label}")
    logger.info(f"  Started : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 55)

    # ----------------------------------------------------------
    # STEP 1 — Load and validate station config
    # ----------------------------------------------------------
    logger.info("\n[1/6] Loading station config...")
    try:
        station_config = load_station_config(station_id)
        logger.info(f"  ✅ Station: {station_config['station_name']} "
                    f"({station_config['location']})")
    except Exception as e:
        logger.error(f"  ❌ Failed to load station config: {e}")
        sys.exit(1)

    # Validate input files exist
    missing = [
        name for name, path in station_config["files"].items()
        if not path.exists()
    ]
    if missing:
        logger.error(f"  ❌ Missing input files: {missing}")
        sys.exit(1)

    # Create output directories
    for dir_path in station_config["report_dirs"].values():
        dir_path.mkdir(parents=True, exist_ok=True)

    # ----------------------------------------------------------
    # STEP 2 — Process data
    # ----------------------------------------------------------
    logger.info("\n[2/6] Processing data...")
    try:
        metrics = process_monthly_report(month, year, station_config)
        logger.info(f"  ✅ Metrics built — "
                    f"{metrics.get('total_days', 0)} days | "
                    f"Fuel Rev: UGX {metrics.get('total_fuel_revenue', 0):,.0f}")
    except Exception as e:
        logger.error(f"  ❌ Processing failed: {e}")
        sys.exit(1)

    # ----------------------------------------------------------
    # STEP 3 — Generate AI narrative
    # ----------------------------------------------------------
    logger.info("\n[3/6] Generating AI narrative...")
    try:
        ai_result = generate_report(metrics, period_label)
        sections  = ai_result["sections"]
        mode      = ai_result.get("mode", "unknown")
        logger.info(f"  ✅ Report generated (mode: {mode})")
    except Exception as e:
        logger.error(f"  ❌ AI generation failed: {e}")
        sys.exit(1)

    # ----------------------------------------------------------
    # STEP 4 — Export PDF + DOCX + XLSX
    # ----------------------------------------------------------
    logger.info("\n[4/6] Exporting reports...")
    try:
        stamped_stem = build_stamped_stem(period_label)
        engine = ExportEngine(station_config=station_config)

        pdf_path  = engine.generate_pdf(metrics, sections, stamped_stem)
        docx_path = engine.generate_docx(metrics, sections, stamped_stem)
        xlsx_path = engine.generate_xlsx(metrics, metrics["daily_df"], stamped_stem)

        logger.info(f"  ✅ PDF  → {Path(pdf_path).name}")
        logger.info(f"  ✅ DOCX → {Path(docx_path).name}")
        logger.info(f"  ✅ XLSX → {Path(xlsx_path).name}")
    except Exception as e:
        logger.error(f"  ❌ Export failed: {e}")
        sys.exit(1)

    # ----------------------------------------------------------
    # STEP 5 — Archive reports
    # ----------------------------------------------------------
    logger.info("\n[5/6] Archiving reports...")
    if station_config.get("archive_enabled", True):
        try:
            archive_reports(
                pdf_path, docx_path, xlsx_path,
                station_config["report_dirs"]["archive"],
                period_label
            )
            logger.info("  ✅ All files archived")
        except Exception as e:
            logger.warning(f"  ⚠️  Archive failed (non-fatal): {e}")
    else:
        logger.info("  ⏭️  Archive disabled for this station")

    # ----------------------------------------------------------
    # STEP 6 — Email delivery
    # ----------------------------------------------------------
    logger.info("\n[6/6] Sending email...")
    if station_config.get("email_enabled", True):
        try:
            # Resolve archived paths (files moved in step 5)
            archive_dir = station_config["report_dirs"]["archive"]
            month_num   = str(month).zfill(2)
            archived    = archive_dir / f"{year}-{month_num}"

            pdf_final  = archived / Path(pdf_path).name
            docx_final = archived / Path(docx_path).name
            xlsx_final = archived / Path(xlsx_path).name

            sent = send_monthly_report(
                pdf_path   = str(pdf_final),
                docx_path  = str(docx_final),
                xlsx_path  = str(xlsx_final),
                period_label = period_label,
                metrics    = metrics,
                recipients = station_config["recipients"]
            )
            status = "Sent" if sent else "Failed"
            logger.info(f"  Email: {status}")
        except Exception as e:
            logger.warning(f"  ⚠️  Email failed (non-fatal): {e}")
    else:
        logger.info("  ⏭️  Email disabled for this station")

    # ----------------------------------------------------------
    # DONE
    # ----------------------------------------------------------
    elapsed = (datetime.now() - start_time).total_seconds()
    logger.info(f"\n{'='*55}")
    logger.info(f"  ✅ Pipeline complete: {period_label}")
    logger.info(f"  Station  : {station_config['station_name']}")
    logger.info(f"  Duration : {elapsed:.1f} seconds")
    logger.info(f"{'='*55}\n")


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="StationDeck — Monthly Report Pipeline"
    )
    parser.add_argument(
        "--station", required=True,
        help="Station config ID (e.g. te_rwizi)"
    )
    parser.add_argument(
        "--month", type=int, required=True,
        help="Report month as integer (1–12)"
    )
    parser.add_argument(
        "--year", type=int, required=True,
        help="Report year (e.g. 2026)"
    )

    args = parser.parse_args()
    run_pipeline(args.station, args.month, args.year)