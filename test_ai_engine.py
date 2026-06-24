# test_ai_engine.py
# ─────────────────────────────────────────────────────────────────────────────
# StationDeck — Phase 3 Test Script
#
# PURPOSE:
#   Tests the full Phase 3 chain:
#   reader.py → processor.py → ai_engine.py
#
# USAGE:
#   python test_ai_engine.py
# ─────────────────────────────────────────────────────────────────────────────

import sys
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

sys.path.insert(0, str(Path(__file__).parent))

from src.reader import read_daily_cashflow
from src.processor import process_monthly_report
from src.ai_engine import generate_report


def build_ai_metrics(raw):
    """
    Translate processor.py key names into the key names ai_engine.py expects.
    This keeps processor.py and ai_engine.py independently maintainable.
    """
    return {
        "days_covered":           raw.get("total_days", 0),
        "pms_volume_litres":      raw.get("pms_volume_total", 0),
        "ago_volume_litres":      raw.get("ago_volume_total", 0),
        "pms_revenue":            raw.get("pms_revenue_total", 0),
        "ago_revenue":            raw.get("ago_revenue_total", 0),
        "total_fuel_revenue":     raw.get("total_fuel_revenue", 0),
        "avg_daily_revenue":      raw.get("avg_daily_fuel_revenue", 0),
        "lubricants_revenue":     raw.get("lubes_revenue_total", 0),
        "lpg_revenue":            raw.get("lpg_revenue_total", 0),
        "shop_sales_revenue":     raw.get("shop_sales_total", 0),
        "total_non_fuel_revenue": raw.get("total_nonfuel_revenue", 0),
        "total_sales":            raw.get("total_sales", 0),
        "total_revenue":          raw.get("total_revenue", 0),
        "cash_collected":         raw.get("cash_collected", 0),
        "cashless_collected":     raw.get("cashless_collected", 0),
        "cash_percentage":        raw.get("cash_percentage", 0),
        "cashless_percentage":    raw.get("cashless_percentage", 0),
        "plus_card_total":        raw.get("plus_card_total", 0),
        "visa_payments":          raw.get("visa_total", 0),
        "credit_sales":           raw.get("credit_sales_total", 0),
        "total_expenses":         raw.get("total_expenses", 0),
        "total_banked":           raw.get("total_cash_banked", 0),
        "expected_to_bank":       raw.get("total_cash_expected", 0),
        "total_delta":            raw.get("total_delta", 0),
        "delta_status":           raw.get("delta_status", "UNKNOWN"),
        "anomaly_days":           raw.get("anomaly_days_count", 0),
    }


def run_test():

    print("=" * 65)
    print("  StationDeck — Phase 3 AI Engine Test")
    print("=" * 65)

    # ── Step 1: Read Excel ────────────────────────────────────────────────
    data_path = Path("data/input/Daily_cash_flow.xlsx")

    if not data_path.exists():
        print(f"\n  ERROR: File not found at {data_path}")
        sys.exit(1)

    print(f"\n  [1/3] Reading Excel file...")
    df = read_daily_cashflow(data_path)
    print(f"        Loaded {len(df)} records.")

    # ── Step 2: Process April 2026 ────────────────────────────────────────
    print(f"\n  [2/3] Processing metrics for April 2026...")
    raw_metrics = process_monthly_report(df, year=2026, month=4)

    if not raw_metrics:
        print("\n  ERROR: Processor returned no metrics.")
        sys.exit(1)

    metrics = build_ai_metrics(raw_metrics)

    print(f"        Days covered:    {metrics['days_covered']}")
    print(f"        Total revenue:   UGX {int(metrics['total_revenue']):,}")
    print(f"        Delta status:    {metrics['delta_status']}")

    # ── Step 3: Generate report ───────────────────────────────────────────
    print(f"\n  [3/3] Generating AI report...")
    result = generate_report(metrics, period_label="April 2026")

    print(f"        Mode:            {result['mode'].upper()}")
    print(f"        Period:          {result['period']}")
    print(f"        Sections found:  {list(result['sections'].keys())}")

    # ── Print report ──────────────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  GENERATED REPORT — APRIL 2026")
    print("=" * 65)
    print()
    print(result["report_text"])

    # ── Section word counts ───────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  SECTION VERIFICATION")
    print("=" * 65)
    for name, text in result["sections"].items():
        count = len(text.split())
        status = "OK" if count > 20 else "WARNING — may be empty"
        print(f"  {name:<38} {count:>4} words   [{status}]")

    print()
    print("=" * 65)
    print("  Phase 3 test complete.")
    print("=" * 65)


if __name__ == "__main__":
    run_test()