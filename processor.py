"""
processor.py
------------
StationDeck Data Processor — the central assembly point of the pipeline.

Reads from ALL data sources and assembles one complete metrics
dictionary for a given month/year. This single dict is then passed
to the AI engine and exporter.

Data sources:
  1. Daily_cash_flow.xlsx          → fuel volumes, cash, payments, reconciliation
  2. Stock_mvt-Template-...xlsx    → fuel stock loss/gain, product sales, expenses
  3. SHOP_MONTHLY_SALES_REPORT...  → shop category breakdown
  4. End_of_May_manager_s_report   → PNL, debtors, depositors, claims, financial position

Usage:
    from src.processor import process_monthly_report
    metrics = process_monthly_report(month=4, year=2026)
"""

import pandas as pd
from pathlib import Path
from datetime import datetime

# ── Existing cashflow reader ──────────────────────────────────
from src.reader import read_daily_cashflow

# ── New readers added in Phase 5 prep ────────────────────────
from src.reader_stock   import get_monthly_stock_metrics
from src.reader_shop    import get_monthly_shop_metrics
from src.reader_manager import get_all_manager_metrics

# ── Settings ─────────────────────────────────────────────────
from config.settings import DATA_INPUT_DIR


# =============================================================
# FILE PATH CONFIGURATION
# =============================================================

CASHFLOW_FILE   = DATA_INPUT_DIR / "Daily_cash_flow.xlsx"
STOCK_FILE      = DATA_INPUT_DIR / "Stock mvt-Template-2025-2026-V1-7-25.xlsx"
SHOP_FILE       = DATA_INPUT_DIR / "SHOP MONTHLY SALES REPORT FOR APRIL 2026 T.E. Rwizi .xlsx"
MANAGER_FILE    = DATA_INPUT_DIR / "End of May manager's report.xlsx"


# =============================================================
# EXISTING CASHFLOW METRIC CALCULATORS (unchanged from Phase 4)
# =============================================================

def filter_by_month(df, year, month):
    """Filters DataFrame to the specified year and month."""
    return df[
        (df["date"].dt.month == month) &
        (df["date"].dt.year  == year)
    ].copy()


def calculate_fuel_metrics(df):
    """Calculates PMS and AGO volume and revenue totals."""
    pms_vol = df["pms_volume"].sum()
    ago_vol = df["ago_volume"].sum()
    pms_rev = df["pms_revenue"].sum()
    ago_rev = df["ago_revenue"].sum()
    total_fuel_rev = pms_rev + ago_rev

    total_days = len(df)
    avg_daily = total_fuel_rev / total_days if total_days > 0 else 0

    return {
        "pms_volume_total":      round(pms_vol, 2),
        "ago_volume_total":      round(ago_vol, 2),
        "pms_revenue_total":     round(pms_rev, 2),
        "ago_revenue_total":     round(ago_rev, 2),
        "total_fuel_revenue":    round(total_fuel_rev, 2),
        "avg_daily_fuel_revenue": round(avg_daily, 2),
        "total_days":            total_days,
    }


def calculate_nonfuel_metrics(df):
    """Calculates non-fuel product revenue totals."""
    lubes  = df["lubes_revenue"].sum()       if "lubes_revenue"  in df.columns else 0
    lpg    = df["lpg_revenue"].sum()         if "lpg_revenue"    in df.columns else 0
    shop   = df["shop_sales"].sum()          if "shop_sales"     in df.columns else 0
    other  = df["tba_credits"].sum()         if "tba_credits"    in df.columns else 0

    total_nonfuel = lubes + lpg + shop + other

    return {
        "lubes_revenue_total":   round(lubes,  2),
        "lpg_revenue_total":     round(lpg,    2),
        "shop_sales_total":      round(shop,   2),
        "other_income_total":    round(other,  2),
        "total_nonfuel_revenue": round(total_nonfuel, 2),
    }


def calculate_cashless_metrics(df):
    """Calculates cashless payment breakdown and percentages."""
    plus_card    = df["plus_card_payment"].sum()   if "plus_card_payment"  in df.columns else 0
    visa         = df["visa"].sum()                if "visa"               in df.columns else 0
    credit_sales = df["credit_sales"].sum()        if "credit_sales"       in df.columns else 0
    cash         = df["total_cash"].sum()          if "total_cash"         in df.columns else 0

    cashless   = plus_card + visa + credit_sales
    total_sales = cash + cashless

    cash_pct     = (cash     / total_sales * 100) if total_sales > 0 else 0
    cashless_pct = (cashless / total_sales * 100) if total_sales > 0 else 0

    return {
        "plus_card_total":      round(plus_card,    2),
        "visa_total":           round(visa,          2),
        "credit_sales_total":   round(credit_sales,  2),
        "cash_collected":       round(cash,          2),
        "cashless_collected":   round(cashless,      2),
        "cash_percentage":      round(cash_pct,      2),
        "cashless_percentage":  round(cashless_pct,  2),
        "total_sales":          round(total_sales,   2),
    }


def calculate_expense_metrics(df):
    """Calculates total expenses from the cashflow sheet."""
    expenses = df["total_expenses"].sum() if "total_expenses" in df.columns else 0
    return {
        "total_expenses": round(expenses, 2),
    }


def calculate_cash_reconciliation(df):
    """Calculates cash reconciliation — banked vs expected, delta and anomalies."""
    banked   = df["actual_cash_banked"].sum()   if "actual_cash_banked"  in df.columns else 0
    expected = df["cash_to_bank"].sum()          if "cash_to_bank"         in df.columns else 0
    delta    = banked - expected

    anomaly_mask  = (df["actual_cash_banked"] - df["cash_to_bank"]).abs() > 0
    anomaly_days  = int(anomaly_mask.sum())

    delta_status = "SURPLUS" if delta >= 0 else "DEFICIT"

    return {
        "total_cash_banked":   round(banked,   2),
        "total_cash_expected": round(expected, 2),
        "total_delta":         round(delta,    2),
        "delta_status":        delta_status,
        "anomaly_days_count":  anomaly_days,
    }


def calculate_summary_metrics(fuel, nonfuel, cashless):
    """Combines fuel and non-fuel into total revenue."""
    total_revenue = fuel["total_fuel_revenue"] + nonfuel["total_nonfuel_revenue"]
    return {
        "total_revenue": round(total_revenue, 2),
    }


# =============================================================
# FORMAT HELPERS
# =============================================================

def format_ugx(value):
    """Formats a number as UGX currency string."""
    try:
        return f"UGX {int(value):,}"
    except (TypeError, ValueError):
        return str(value)


def format_litres(value):
    """Formats a number as litres string."""
    try:
        return f"{float(value):,.2f} L"
    except (TypeError, ValueError):
        return str(value)


# =============================================================
# MAIN FUNCTION — PROCESS MONTHLY REPORT
# =============================================================

def process_monthly_report(month: int, year: int) -> dict:
    """
    Master function — loads ALL data sources and returns one
    complete metrics dictionary for the given month and year.
    """

    metrics = {}

    # ── SECTION 1: Daily Cashflow (core data) ─────────────────
    print("  [1/4] Reading daily cashflow data...")
    try:
        df_all = read_daily_cashflow(CASHFLOW_FILE)
        df_all["date"] = pd.to_datetime(df_all["date"], errors="coerce")
        df = filter_by_month(df_all, year, month)

        if df.empty:
            print(f"  [WARNING] No cashflow data found for {month}/{year}")
        else:
            fuel_m    = calculate_fuel_metrics(df)
            nonfuel_m = calculate_nonfuel_metrics(df)
            cashless_m= calculate_cashless_metrics(df)
            expense_m = calculate_expense_metrics(df)
            recon_m   = calculate_cash_reconciliation(df)
            summary_m = calculate_summary_metrics(fuel_m, nonfuel_m, cashless_m)

            metrics.update(fuel_m)
            metrics.update(nonfuel_m)
            metrics.update(cashless_m)
            metrics.update(expense_m)
            metrics.update(recon_m)
            metrics.update(summary_m)

            # ✅ FIX: store the filtered daily DataFrame so exporter can use it
            metrics["daily_df"] = df

            print(f"      Days: {fuel_m['total_days']} | "
                  f"Fuel Rev: UGX {fuel_m['total_fuel_revenue']:,.0f} | "
                  f"Delta: {recon_m['delta_status']}")

    except Exception as e:
        print(f"  [WARNING] Cashflow read failed: {e}")

    # ── SECTION 2: Stock Template (fuel stock + expenses) ──────
    print("  [2/4] Reading stock movement data...")
    try:
        if STOCK_FILE.exists():
            stock_m = get_monthly_stock_metrics(STOCK_FILE, month=month, year=year)
            metrics["fuel_stock"]    = stock_m.get("fuel_stock", {})
            metrics["product_sales"] = stock_m.get("product_sales", {})
            if stock_m.get("expenses", {}).get("total_expenses", 0) > 0:
                metrics["expense_detail"] = stock_m.get("expenses", {})
                metrics["total_expenses"] = stock_m["expenses"]["total_expenses"]
            print(f"      PMS loss/gain: {metrics['fuel_stock'].get('pms', {}).get('loss_gain_ltrs', 'N/A')} L | "
                  f"AGO: {metrics['fuel_stock'].get('ago', {}).get('loss_gain_ltrs', 'N/A')} L")
        else:
            print(f"  [WARNING] Stock file not found: {STOCK_FILE.name}")
            metrics["fuel_stock"]    = {}
            metrics["product_sales"] = {}
            metrics["expense_detail"] = {}

    except Exception as e:
        print(f"  [WARNING] Stock data read failed: {e}")
        metrics["fuel_stock"]    = {}
        metrics["product_sales"] = {}
        metrics["expense_detail"] = {}

    # ── SECTION 3: Shop Monthly Sales ──────────────────────────
    print("  [3/4] Reading shop sales data...")
    try:
        if SHOP_FILE.exists():
            shop_m = get_monthly_shop_metrics(SHOP_FILE, month=month, year=year)
            metrics["shop_sales_detail"] = shop_m
            if shop_m:
                print(f"      Shop total: UGX {shop_m.get('global_total', 0):,.0f} | "
                      f"Top category: {shop_m.get('top_3_categories', [{}])[0].get('category', 'N/A')}")
        else:
            print(f"  [WARNING] Shop file not found: {SHOP_FILE.name}")
            metrics["shop_sales_detail"] = {}

    except Exception as e:
        print(f"  [WARNING] Shop data read failed: {e}")
        metrics["shop_sales_detail"] = {}

    # ── SECTION 4: Manager's Report (PNL, debtors, etc.) ───────
    print("  [4/4] Reading manager report data...")
    try:
        if MANAGER_FILE.exists():
            mgr_m = get_all_manager_metrics(MANAGER_FILE)
            metrics["pnl"]                = mgr_m.get("pnl", {})
            metrics["debtors"]            = mgr_m.get("debtors", {})
            metrics["depositors"]         = mgr_m.get("depositors", {})
            metrics["financial_position"] = mgr_m.get("financial_position", {})
            metrics["claims"]             = mgr_m.get("claims", {})
            metrics["money_meters"]       = mgr_m.get("money_meters", {})

            pnl = mgr_m.get("pnl", {})
            if pnl:
                print(f"      Gross income: UGX {pnl.get('gross_income', 0):,.0f} | "
                      f"Net profit: UGX {pnl.get('net_profit', 0):,.0f} | "
                      f"Reserve: UGX {pnl.get('reserve_balance', 0):,.0f}")
        else:
            print(f"  [WARNING] Manager report not found: {MANAGER_FILE.name}")
            metrics["pnl"]                = {}
            metrics["debtors"]            = {}
            metrics["depositors"]         = {}
            metrics["financial_position"] = {}
            metrics["claims"]             = {}
            metrics["money_meters"]       = {}

    except Exception as e:
        print(f"  [WARNING] Manager report read failed: {e}")
        metrics["pnl"]                = {}
        metrics["debtors"]            = {}
        metrics["depositors"]         = {}
        metrics["financial_position"] = {}
        metrics["claims"]             = {}
        metrics["money_meters"]       = {}

    # ── Attach period metadata ──────────────────────────────────
    import calendar
    metrics["report_month"]  = month
    metrics["report_year"]   = year
    metrics["period_label"]  = f"{calendar.month_name[month]} {year}"
    metrics["generated_at"]  = datetime.now().isoformat()

    return metrics


# =============================================================
# QUICK TEST  (run directly to verify)
# =============================================================

if __name__ == "__main__":
    import json

    print("=" * 60)
    print("TEST: process_monthly_report() — April 2026")
    print("=" * 60)

    metrics = process_monthly_report(month=4, year=2026)

    print()
    print("── CASHFLOW METRICS ──────────────────────────────────")
    cashflow_keys = [
        "total_days", "pms_volume_total", "ago_volume_total",
        "pms_revenue_total", "ago_revenue_total", "total_fuel_revenue",
        "total_expenses", "total_cash_banked", "total_delta", "delta_status",
        "anomaly_days_count", "total_sales", "total_revenue",
    ]
    for k in cashflow_keys:
        v = metrics.get(k, "MISSING")
        if isinstance(v, float):
            print(f"  {k:<30}: {v:,.2f}")
        else:
            print(f"  {k:<30}: {v}")

    print()
    print("── FUEL STOCK ────────────────────────────────────────")
    for prod, data in metrics.get("fuel_stock", {}).items():
        print(f"  {prod.upper()}: loss/gain = {data.get('loss_gain_ltrs')} L "
              f"({data.get('loss_gain_pct')}%) | "
              f"value = UGX {data.get('loss_gain_value_ugx', 0):,.0f}")

    print()
    print("── PRODUCT SALES ─────────────────────────────────────")
    for k, v in metrics.get("product_sales", {}).items():
        print(f"  {k:<30}: UGX {v:,.0f}")

    print()
    print("── SHOP SUMMARY ──────────────────────────────────────")
    shop = metrics.get("shop_sales_detail", {})
    if shop:
        print(f"  Total turnover   : UGX {shop.get('total_turnover', 0):,.0f}")
        print(f"  Trading days     : {shop.get('trading_days')}")
        print(f"  Avg daily sales  : UGX {shop.get('avg_daily_sales', 0):,.0f}")
        print(f"  Best day         : {shop.get('best_day')}")
        print(f"  Top 3 categories : {shop.get('top_3_categories')}")

    print()
    print("── PNL ───────────────────────────────────────────────")
    pnl = metrics.get("pnl", {})
    if pnl:
        print(f"  Gross income     : UGX {pnl.get('gross_income', 0):,.0f}")
        print(f"  Total expenses   : UGX {pnl.get('total_expenses', 0):,.0f}")
        print(f"  Net profit       : UGX {pnl.get('net_profit', 0):,.0f}")
        print(f"  Price change fx  : UGX {pnl.get('price_change_effect', 0):,.0f}")
        print(f"  Reserve balance  : UGX {pnl.get('reserve_balance', 0):,.0f}")
        print(f"  Expense lines    : {len(pnl.get('expense_lines', []))}")

    print()
    print("── DEBTORS ───────────────────────────────────────────")
    deb = metrics.get("debtors", {})
    if deb:
        print(f"  Customers        : {len(deb.get('customers', []))}")
        print(f"  Total outstanding: UGX {deb.get('total_outstanding', 0):,.0f}")
        print(f"  Overdue count    : {deb.get('overdue_count')}")

    print()
    print("── DEPOSITORS ────────────────────────────────────────")
    dep = metrics.get("depositors", {})
    if dep:
        print(f"  Total customers  : {len(dep.get('customers', []))}")
        print(f"  Active depositors: {dep.get('active_depositors')}")
        print(f"  Total balance    : UGX {dep.get('total_balance', 0):,.0f}")

    print()
    print("── FINANCIAL POSITION ────────────────────────────────")
    fp = metrics.get("financial_position", {})
    if fp:
        print(f"  Total assets     : UGX {fp.get('total_assets', 0):,.0f}")
        print(f"  Total liabilities: UGX {fp.get('total_liabilities', 0):,.0f}")
        print(f"  Net position     : UGX {fp.get('net_position', 0):,.0f}")

    print()
    print("── CLAIMS ────────────────────────────────────────────")
    cl = metrics.get("claims", {})
    if cl:
        print(f"  Claims           : {len(cl.get('claims', []))}")
        print(f"  Total value      : UGX {cl.get('total_claims', 0):,.0f}")

    print()
    print("=" * 60)
    print("ALL SECTIONS LOADED ✅")
    print("=" * 60)