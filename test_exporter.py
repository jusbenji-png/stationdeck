# =============================================================
# StationDeck - Full Pipeline Test
# =============================================================
# Tests the complete export pipeline with all data sources:
#   processor -> ai_engine -> exporter
#
# Generates PDF, DOCX, and XLSX reports for the target month.
#
# Run from your project root with venv active:
#   python test_exporter.py
# =============================================================

import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

from src.processor import process_monthly_report
from src.ai_engine  import generate_report
from src.exporter   import ExportEngine
from src.reader     import read_daily_cashflow
from config.settings import DATA_INPUT_DIR

import pandas as pd

# =============================================================
# CONFIGURATION — change these to test different months
# =============================================================

TARGET_MONTH  = 4       # April
TARGET_YEAR   = 2026
PERIOD_LABEL  = "April 2026"
CASHFLOW_FILE = DATA_INPUT_DIR / "Daily_cash_flow.xlsx"


# =============================================================
# STEP 1 — Process all data sources into one metrics package
# =============================================================

def step_process():
    print("\n" + "=" * 55)
    print("  STEP 1 — Processing All Data Sources")
    print("=" * 55)

    metrics = process_monthly_report(month=TARGET_MONTH, year=TARGET_YEAR)

    print(f"\n  Summary:")
    print(f"    Days covered     : {metrics.get('total_days')}")
    print(f"    Fuel revenue     : UGX {metrics.get('total_fuel_revenue', 0):,.0f}")
    print(f"    Total revenue    : UGX {metrics.get('total_revenue', 0):,.0f}")
    print(f"    Delta status     : {metrics.get('delta_status')}")
    print(f"    Net profit (PNL) : UGX {metrics.get('pnl', {}).get('net_profit', 0):,.0f}")
    print(f"    Shop turnover    : UGX {metrics.get('shop_sales_detail', {}).get('total_turnover', 0):,.0f}")
    print(f"    Debtors          : UGX {metrics.get('debtors', {}).get('total_outstanding', 0):,.0f}")
    print(f"    Claims           : UGX {metrics.get('claims', {}).get('total_claims', 0):,.0f}")

    return metrics


# =============================================================
# STEP 2 — Generate AI report sections
# =============================================================

def step_ai(metrics):
    print("\n" + "=" * 55)
    print("  STEP 2 — Generating AI Report Sections")
    print("=" * 55)

    result   = generate_report(metrics, PERIOD_LABEL)
    sections = result["sections"]
    mode     = result["mode"]

    print(f"\n  Mode: {mode.upper()}")
    for section_name, content in sections.items():
        word_count = len(content.split()) if content else 0
        print(f"  ✅ {section_name}: {word_count} words")

    return sections


# =============================================================
# STEP 3 — Load daily data for XLSX export
# =============================================================

def step_load_daily():
    """Loads the daily cashflow DataFrame for the XLSX daily sheet."""
    df = read_daily_cashflow(CASHFLOW_FILE)
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    monthly = df[
        (df["date"].dt.month == TARGET_MONTH) &
        (df["date"].dt.year  == TARGET_YEAR)
    ].copy()
    return monthly


# =============================================================
# STEP 4 — Export to PDF, DOCX, and XLSX
# =============================================================

def step_export(metrics, sections, daily_df):
    print("\n" + "=" * 55)
    print("  STEP 3 — Exporting Reports")
    print("=" * 55)

    exporter = ExportEngine()

    print("\n  Generating PDF...")
    pdf_path = exporter.generate_pdf(metrics, sections, PERIOD_LABEL)

    print("\n  Generating Word document...")
    docx_path = exporter.generate_docx(metrics, sections, PERIOD_LABEL)

    print("\n  Generating Excel workbook...")
    xlsx_path = exporter.generate_xlsx(metrics, daily_df, PERIOD_LABEL)

    return pdf_path, docx_path, xlsx_path


# =============================================================
# MAIN
# =============================================================

def main():
    print("\n" + "=" * 55)
    print("  StationDeck — Full Pipeline Test")
    print(f"  Period : {PERIOD_LABEL}")
    print(f"  Run at : {datetime.now().strftime('%d %B %Y, %H:%M:%S')}")
    print("=" * 55)

    metrics                        = step_process()
    sections                       = step_ai(metrics)
    daily_df                       = step_load_daily()
    pdf_path, docx_path, xlsx_path = step_export(metrics, sections, daily_df)

    print("\n" + "=" * 55)
    print("  ✅ EXPORT COMPLETE")
    print("=" * 55)
    print(f"\n  📄 PDF  : {pdf_path}")
    print(f"  📝 DOCX : {docx_path}")
    print(f"  📊 XLSX : {xlsx_path}")
    print("\n  Open all three files to verify the report.")
    print("=" * 55 + "\n")


if __name__ == "__main__":
    main()