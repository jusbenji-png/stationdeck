# =============================================================
# StationDeck - Export Engine Test
# =============================================================
# Tests the full export pipeline:
#   reader → processor → ai_engine → exporter
#
# Generates a real PDF and DOCX report for April 2026.
#
# Run from your project root with venv active:
#   python test_exporter.py
# =============================================================

import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent))

from src.reader    import read_daily_cashflow
from src.processor import process_monthly_report
from src.ai_engine import generate_report
from src.exporter  import ExportEngine
from config.settings import DATA_INPUT_DIR

# =============================================================
# CONFIGURATION — change these to test different months
# =============================================================

TARGET_MONTH  = 4
TARGET_YEAR   = 2026
PERIOD_LABEL  = "April 2026"
EXCEL_FILE    = DATA_INPUT_DIR / "Daily_cash_flow.xlsx"


# =============================================================
# STEP 1 — Read Excel data
# =============================================================

def step_read():
    print("\n" + "=" * 55)
    print("  STEP 1 — Reading Excel Data")
    print("=" * 55)

    if not EXCEL_FILE.exists():
        print(f"  ❌ File not found: {EXCEL_FILE}")
        sys.exit(1)

    df = read_daily_cashflow(EXCEL_FILE)

    print(f"  Total records loaded:   {len(df)}")

    monthly = df[
        (df["date"].dt.month == TARGET_MONTH) &
        (df["date"].dt.year  == TARGET_YEAR)
    ].copy()

    print(f"  Records for {PERIOD_LABEL}: {len(monthly)}")

    if monthly.empty:
        print(f"  ❌ No data found for {PERIOD_LABEL}. Check your Excel file.")
        sys.exit(1)

    return monthly


# =============================================================
# STEP 2 — Process data into metrics
# =============================================================

def step_process(monthly_df):
    print("\n" + "=" * 55)
    print("  STEP 2 — Processing Metrics")
    print("=" * 55)

    metrics = process_monthly_report(monthly_df, TARGET_YEAR, TARGET_MONTH)

    print(f"  Days covered:        {metrics.get('total_days')}")
    print(f"  Total fuel revenue:  UGX {metrics.get('total_fuel_revenue'):,.0f}")
    print(f"  Total sales:         UGX {metrics.get('total_sales'):,.0f}")
    print(f"  Delta status:        {metrics.get('delta_status')}")
    print(f"  Anomaly days:        {metrics.get('anomaly_days_count')}")

    return metrics


# =============================================================
# STEP 3 — Generate AI report sections
# =============================================================

def step_ai(metrics):
    print("\n" + "=" * 55)
    print("  STEP 3 — Generating AI Report Sections")
    print("=" * 55)

    ai_metrics = {
        "station_name":           "Total Energies Rwizi",
        "report_period":          PERIOD_LABEL,
        "total_days":             metrics.get("total_days"),
        "pms_volume":             metrics.get("pms_volume_total"),
        "ago_volume":             metrics.get("ago_volume_total"),
        "pms_revenue":            metrics.get("pms_revenue_total"),
        "ago_revenue":            metrics.get("ago_revenue_total"),
        "total_fuel_revenue":     metrics.get("total_fuel_revenue"),
        "avg_daily_fuel_revenue": metrics.get("avg_daily_fuel_revenue"),
        "lubes_revenue":          metrics.get("lubes_revenue_total"),
        "lpg_revenue":            metrics.get("lpg_revenue_total"),
        "shop_sales":             metrics.get("shop_sales_total"),
        "total_nonfuel_revenue":  metrics.get("total_nonfuel_revenue"),
        "total_sales":            metrics.get("total_sales"),
        "cash_collected":         metrics.get("cash_collected"),
        "cashless_collected":     metrics.get("cashless_collected"),
        "cash_percentage":        metrics.get("cash_percentage"),
        "cashless_percentage":    metrics.get("cashless_percentage"),
        "plus_card_total":        metrics.get("plus_card_total"),
        "visa_total":             metrics.get("visa_total"),
        "credit_sales_total":     metrics.get("credit_sales_total"),
        "total_expenses":         metrics.get("total_expenses"),
        "total_cash_banked":      metrics.get("total_cash_banked"),
        "total_cash_expected":    metrics.get("total_cash_expected"),
        "total_delta":            metrics.get("total_delta"),
        "delta_status":           metrics.get("delta_status"),
        "anomaly_days_count":     metrics.get("anomaly_days_count"),
        "total_revenue":          metrics.get("total_revenue"),
    }

    sections = generate_report(ai_metrics, PERIOD_LABEL)

    for section_name, content in sections.items():
        word_count = len(content.split())
        print(f"  ✅ {section_name}: {word_count} words")

    return sections


# =============================================================
# STEP 4 — Export to PDF and DOCX
# =============================================================

def step_export(metrics, sections):
    print("\n" + "=" * 55)
    print("  STEP 4 — Exporting Reports")
    print("=" * 55)

    exporter = ExportEngine()

    print("\n  Generating PDF...")
    pdf_path = exporter.generate_pdf(metrics, sections, PERIOD_LABEL)

    print("\n  Generating Word document...")
    docx_path = exporter.generate_docx(metrics, sections, PERIOD_LABEL)

    return pdf_path, docx_path


# =============================================================
# MAIN
# =============================================================

def main():
    print("\n" + "=" * 55)
    print("  StationDeck — Export Pipeline Test")
    print(f"  Period: {PERIOD_LABEL}")
    print(f"  Run at: {datetime.now().strftime('%d %B %Y, %H:%M:%S')}")
    print("=" * 55)

    monthly_df           = step_read()
    metrics              = step_process(monthly_df)
    sections             = step_ai(metrics)
    pdf_path, docx_path  = step_export(metrics, sections)

    print("\n" + "=" * 55)
    print("  ✅ EXPORT COMPLETE")
    print("=" * 55)
    print(f"\n  📄 PDF:  {pdf_path}")
    print(f"  📝 DOCX: {docx_path}")
    print("\n  Open both files to verify the report looks correct.")
    print("=" * 55 + "\n")


if __name__ == "__main__":
    main()