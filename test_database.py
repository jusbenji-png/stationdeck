"""
test_database.py — Phase 9A verification script

Run from project root with:
  python test_database.py

Tests:
  1. DB initialisation (creates stationdeck.db)
  2. Single record save
  3. Excel bulk import (all historical data)
  4. Query by month (April 2026)
  5. Query by week
  6. Query by financial year
  7. Duplicate prevention
  8. Record count and date range
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from pathlib import Path
from datetime import date

# Ensure src/ is importable
sys.path.insert(0, str(Path(__file__).parent))

from src.database import (
    init_db,
    save_daily_record,
    import_from_excel,
    get_records_by_month,
    get_records_by_week,
    get_records_by_date_range,
    get_records_by_financial_year,
    get_record_count,
    get_date_range_stored,
    DB_PATH,
)

STATION = "te_rwizi"
CASHFLOW_PATH = Path("data/input/Daily_cash_flow.xlsx")

PASS = "[PASS]"
FAIL = "[FAIL]"

def separator(title):
    print(f"\n{'='*55}")
    print(f"  {title}")
    print(f"{'='*55}")

# ── TEST 1: Init DB ──────────────────────────────────────
separator("TEST 1: Initialise Database")
result = init_db()
print(f"  init_db() returned: {result}")
print(f"  DB file exists: {DB_PATH.exists()}")
print(f"  DB path: {DB_PATH}")
print(f"  {PASS if result and DB_PATH.exists() else FAIL}")

# ── TEST 2: Save a single record ─────────────────────────
separator("TEST 2: Save Single Record")
test_record = {
    "date": date(2020, 1, 1),
    "pms_volume": 1000.0,
    "ago_volume": 500.0,
    "pms_price": 4800.0,
    "ago_price": 4600.0,
    "pms_revenue": 4800000.0,
    "ago_revenue": 2300000.0,
    "lubes_litres": 0.0,
    "lubes_revenue": 0.0,
    "lpg_kgs": 0.0,
    "lpg_revenue": 0.0,
    "tba_credits": 0.0,
    "plus_card_payment": 0.0,
    "shop_sales": 50000.0,
    "plus_card_payment_total": 0.0,
    "tyre_sales": 0.0,
    "cashless_total": 7150000.0,
    "plus_card_pms": 0.0,
    "plus_card_ago": 0.0,
    "other_payments": 0.0,
    "momo_pay": 0.0,
    "airtel_pay": 0.0,
    "visa": 0.0,
    "credit_sales": 0.0,
    "total_cash": 7150000.0,
    "expense_umeme": 0.0,
    "expense_water": 0.0,
    "expense_security": 0.0,
    "expense_stationery": 0.0,
    "expense_generator": 0.0,
    "expense_meals": 0.0,
    "expense_transport": 0.0,
    "expense_salaries": 0.0,
    "expense_sanitary": 0.0,
    "expense_airtime": 0.0,
    "expense_misc": 0.0,
    "expense_shop_packaging": 0.0,
    "total_expenses": 0.0,
    "stock_tba": 0.0,
    "stock_lpg_acc": 0.0,
    "stock_shop_purchase": 0.0,
    "cash_to_bank": 7150000.0,
    "actual_cash_banked": 7150000.0,
    "delta": 0.0,
    "source_sheet": "TEST",
}
ok = save_daily_record(test_record, STATION)
print(f"  save_daily_record() returned: {ok}")
print(f"  {PASS if ok else FAIL}")

# ── TEST 3: Bulk import from Excel ───────────────────────
separator("TEST 3: Bulk Import from Excel")
if not CASHFLOW_PATH.exists():
    print(f"  {FAIL} — File not found: {CASHFLOW_PATH}")
    print("  Make sure Daily_cash_flow.xlsx is in data/input/")
else:
    count = import_from_excel(CASHFLOW_PATH, STATION)
    print(f"  Records imported: {count}")
    print(f"  {PASS if count > 0 else FAIL}")

# ── TEST 4: Query by month ───────────────────────────────
separator("TEST 4: Query by Month (April 2026)")
df_april = get_records_by_month(4, 2026, STATION)
print(f"  Rows returned: {len(df_april)}")
if not df_april.empty:
    print(f"  Date range: {df_april['date'].min().date()} → {df_april['date'].max().date()}")
    print(f"  Columns: {list(df_april.columns[:6])} ...")
    pms_total = df_april["pms_volume"].sum()
    print(f"  Total PMS volume (April 2026): {pms_total:,.2f} L")
print(f"  {PASS if len(df_april) > 0 else FAIL}")

# ── TEST 5: Query by week ────────────────────────────────
separator("TEST 5: Query by Week (week of 2026-04-01)")
df_week = get_records_by_week("2026-04-01", STATION)
print(f"  Rows returned: {len(df_week)}")
if not df_week.empty:
    print(f"  Date range: {df_week['date'].min().date()} → {df_week['date'].max().date()}")
print(f"  {PASS if len(df_week) > 0 else FAIL}")

# ── TEST 6: Query by financial year ─────────────────────
separator("TEST 6: Query by Financial Year (FY 2025/2026)")
df_fy = get_records_by_financial_year(2025, STATION)
print(f"  Rows returned: {len(df_fy)}")
if not df_fy.empty:
    print(f"  Date range: {df_fy['date'].min().date()} → {df_fy['date'].max().date()}")
print(f"  {PASS if len(df_fy) > 0 else FAIL}")

# ── TEST 7: Duplicate prevention ─────────────────────────
separator("TEST 7: Duplicate Prevention")
count_before = get_record_count(STATION)
# Save the same test record again — should replace, not duplicate
save_daily_record(test_record, STATION)
count_after = get_record_count(STATION)
print(f"  Record count before re-save: {count_before}")
print(f"  Record count after re-save:  {count_after}")
print(f"  {PASS if count_before == count_after else FAIL} (counts should be equal)")

# ── TEST 8: Record count and date range ──────────────────
separator("TEST 8: Record Count and Date Coverage")
total = get_record_count(STATION)
coverage = get_date_range_stored(STATION)
print(f"  Total records in DB: {total}")
print(f"  Earliest date: {coverage['earliest']}")
print(f"  Latest date:   {coverage['latest']}")
print(f"  {PASS if total > 0 else FAIL}")

# ── SUMMARY ──────────────────────────────────────────────
separator("TEST COMPLETE")
print("  Paste your full output above into the chat.")
print("  Claude will review before we proceed to Phase 9B.")
print()