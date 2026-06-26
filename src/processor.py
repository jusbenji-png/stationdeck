# ============================================================
# src/processor.py
# StationDeck — Monthly + Annual Data Processor
#
# Reads all 4 Excel sources for a given station/month/year,
# merges them into one complete metrics dictionary, and
# returns it ready for the AI engine and exporters.
#
# EXPENSE RULE (locked):
#   All expense figures — monthly and annual — come exclusively
#   from the Daily Expenses sheet of the Stock Movement file.
#   The cashflow file's total_expenses (col 37) is NEVER used
#   in any report output. This ensures both reports always agree.
#
# SHOP TURNOVER RULE (locked):
#   shop_sales_total always comes from product_sales["shop_ugx"],
#   which is the General Sales Summary sheet col 9.
#   This matches the Shop Sales sheet daily totals exactly.
#   The cashflow file's shop_sales column (col 13) is ignored.
#
# Usage:
#   from src.processor import process_monthly_report
#   metrics = process_monthly_report(4, 2026, station_config)
#
#   from src.processor import process_annual_report
#   metrics = process_annual_report(2025, station_config)
# ============================================================

import pandas as pd
from pathlib import Path
from datetime import datetime

from src.reader         import read_daily_cashflow
from src.reader_stock   import get_monthly_stock_metrics, get_annual_stock_metrics
from src.reader_shop    import get_monthly_shop_metrics
from src.reader_manager import get_all_manager_metrics


MONTH_NAMES = {
    1: "January",  2: "February", 3: "March",     4: "April",
    5: "May",      6: "June",     7: "July",       8: "August",
    9: "September",10: "October", 11: "November", 12: "December",
}


# ============================================================
# MONTHLY REPORT PROCESSOR
# ============================================================

def process_monthly_report(month: int, year: int, station_config: dict) -> dict:
    """
    Master processing function for a single calendar month.

    Reads all 4 data sources, merges into one complete metrics dict.

    Args:
        month:          Integer month (1-12)
        year:           Integer year (e.g. 2026)
        station_config: Dict from load_station_config()

    Returns:
        Complete metrics dict ready for ai_engine and exporter.
    """
    files         = station_config["files"]
    CASHFLOW_FILE = files["cashflow"]
    STOCK_FILE    = files["stock"]
    SHOP_FILE     = files["shop"]
    MANAGER_FILE  = files["manager"]

    period_label = f"{MONTH_NAMES[month]} {year}"

    print(f"\n{'='*55}")
    print(f"  Processing: {period_label}")
    print(f"  Station:    {station_config['station_name']}")
    print(f"{'='*55}")

    metrics = {}

    # --------------------------------------------------------
    # SOURCE 1 — Daily Cash Flow
    # --------------------------------------------------------
    print("\n[1/4] Reading Daily Cash Flow...")
    try:
        df = read_daily_cashflow(CASHFLOW_FILE)
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

        mask     = (df["date"].dt.month == month) & (df["date"].dt.year == year)
        df_month = df[mask].copy()

        if df_month.empty:
            print(f"  ⚠️  No cashflow data found for {period_label}")
            df_month = pd.DataFrame()

        total_days = len(df_month)
        print(f"  ✅ {total_days} days of cashflow data found")

        # Fuel volumes and revenue
        pms_volume_total   = df_month["pms_volume"].sum()
        ago_volume_total   = df_month["ago_volume"].sum()
        pms_revenue_total  = df_month["pms_revenue"].sum()
        ago_revenue_total  = df_month["ago_revenue"].sum()
        total_fuel_revenue = pms_revenue_total + ago_revenue_total

        # Payment collection
        # cashless_total (col 16) = grand total station sales
        # Cashless payments = cashless_total - total_cash
        plus_card_total    = df_month["plus_card_payment_total"].sum()
        visa_total         = df_month["visa"].sum()
        credit_sales_total = df_month["credit_sales"].sum()
        cash_collected     = df_month["total_cash"].sum()
        total_sales        = df_month["cashless_total"].sum()
        cashless_collected = total_sales - cash_collected

        cash_percentage     = (cash_collected / total_sales * 100)     if total_sales > 0 else 0
        cashless_percentage = (cashless_collected / total_sales * 100) if total_sales > 0 else 0

        # Cash reconciliation
        total_cash_banked  = df_month["cash_to_bank"].sum()
        actual_cash_banked = df_month["actual_cash_banked"].sum()
        total_delta        = df_month["delta"].sum()
        delta_status       = (
            "SURPLUS"  if total_delta > 0 else
            "DEFICIT"  if total_delta < 0 else
            "BALANCED"
        )
        anomaly_days_count     = int((df_month["delta"] != 0).sum())
        avg_daily_fuel_revenue = total_fuel_revenue / total_days if total_days > 0 else 0

        # NOTE: We do NOT use cashflow total_expenses or shop_sales here.
        # Those are overridden in SOURCE 2 with stock file values.
        # We set placeholder zeros so keys exist in the dict.
        metrics.update({
            "total_days":               total_days,
            "pms_volume_total":         pms_volume_total,
            "ago_volume_total":         ago_volume_total,
            "pms_revenue_total":        pms_revenue_total,
            "ago_revenue_total":        ago_revenue_total,
            "total_fuel_revenue":       total_fuel_revenue,
            "avg_daily_fuel_revenue":   avg_daily_fuel_revenue,
            # Non-fuel totals — placeholders, overridden by stock file in SOURCE 2
            "lubes_revenue_total":      0,
            "lpg_revenue_total":        0,
            "shop_sales_total":         0,   # always overridden by stock file
            "tba_revenue_total":        0,
            "lpg_accessories_total":    0,
            "car_wash_total":           0,
            "other_income_total":       0,
            "total_nonfuel_revenue":    0,
            "total_revenue":            total_fuel_revenue,
            # Payment collection
            "plus_card_total":          plus_card_total,
            "visa_total":               visa_total,
            "credit_sales_total":       credit_sales_total,
            "cash_collected":           cash_collected,
            "cashless_collected":       cashless_collected,
            "cash_percentage":          round(cash_percentage, 1),
            "cashless_percentage":      round(cashless_percentage, 1),
            "total_sales":              total_sales,
            # Expenses — placeholder zero, always overridden by stock file in SOURCE 2
            "total_expenses":           0,
            "expense_detail":           {},
            # Reconciliation
            "total_cash_banked":        total_cash_banked,
            "actual_cash_banked":       actual_cash_banked,
            "total_cash_expected":      total_cash_banked,
            "total_delta":              total_delta,
            "delta_status":             delta_status,
            "anomaly_days_count":       anomaly_days_count,
            # Daily DataFrame — used by exporter for Daily Records sheet
            "daily_df":                 df_month,
        })

    except Exception as e:
        print(f"  ❌ Cashflow read failed: {e}")
        metrics["daily_df"] = pd.DataFrame()

    # --------------------------------------------------------
    # SOURCE 2 — Stock Movement (reader_stock.py)
    # --------------------------------------------------------
    print("\n[2/4] Reading Stock Movement...")
    try:
        stock = get_monthly_stock_metrics(STOCK_FILE, month, year)
        stock_available = stock.get("data_available", False)

        # ── Fuel stock movement ──────────────────────────────
        metrics["fuel_stock"]           = stock.get("fuel_stock", {})
        metrics["stock_data_available"] = stock_available

        # ── Product sales (authoritative for all non-fuel products) ──
        ps = stock.get("product_sales", {})
        metrics["product_sales"] = ps

        if stock_available and ps:
            # Lubricants, LPG, accessories, TBA, car wash — from General Sales Summary
            metrics["lubes_revenue_total"]   = ps.get("lubricants_ugx", 0)
            metrics["lpg_revenue_total"]     = ps.get("lpg_ugx", 0)
            metrics["tba_revenue_total"]     = ps.get("tba_ugx", 0)
            metrics["lpg_accessories_total"] = ps.get("lpg_accessories_ugx", 0)
            metrics["car_wash_total"]        = ps.get("car_wash_ugx", 0)

            # SHOP TURNOVER — authoritative source: General Sales Summary col 9
            # (verified to match Shop Sales sheet daily totals exactly)
            metrics["shop_sales_total"] = ps.get("shop_ugx", 0)

            metrics["total_nonfuel_revenue"] = (
                metrics["lubes_revenue_total"]   +
                metrics["lpg_revenue_total"]     +
                metrics["shop_sales_total"]      +
                metrics["tba_revenue_total"]     +
                metrics["lpg_accessories_total"] +
                metrics["car_wash_total"]
            )
            metrics["total_revenue"] = (
                metrics.get("total_fuel_revenue", 0) +
                metrics["total_nonfuel_revenue"]
            )

        # ── EXPENSES — stock file is the ONLY source ─────────
        # The Daily Expenses sheet total (col 31) is authoritative.
        # We never use cashflow col 37 in any report output.
        exp = stock.get("expenses_detail", {})
        if exp and exp.get("total_expenses", 0) > 0:
            metrics["expenses_detail_stock"] = exp
            metrics["total_expenses"]        = exp.get("total_expenses", 0)
        else:
            metrics["expenses_detail_stock"] = {}
            metrics["total_expenses"]        = 0
            print("  [INFO] No expense data from stock file for this period")

        # ── Daily non-fuel product breakdown (for XLSX Daily Records) ──
        # general_sales_daily is a list of dicts, one per day,
        # with lubricants, lpg, accessories, tba, car_wash, shop, solar
        metrics["general_sales_daily"] = stock.get("general_sales_daily", [])

        print(f"  ✅ Stock metrics loaded (data_available={stock_available})")

        if stock_available:
            pms_s = stock.get("fuel_stock", {}).get("pms", {})
            ago_s = stock.get("fuel_stock", {}).get("ago", {})
            print(f"     PMS: {pms_s.get('total_sales_ltrs', 0):,.2f} L  "
                  f"deliveries: {pms_s.get('delivery_count', 0)}")
            print(f"     AGO: {ago_s.get('total_sales_ltrs', 0):,.2f} L  "
                  f"deliveries: {ago_s.get('delivery_count', 0)}")
            print(f"     Shop turnover: UGX {metrics['shop_sales_total']:,.0f}")
            print(f"     Total expenses: UGX {metrics['total_expenses']:,.0f}")

    except Exception as e:
        print(f"  ⚠️  Stock read failed (non-fatal): {e}")
        metrics["fuel_stock"]            = {}
        metrics["product_sales"]         = {}
        metrics["general_sales_daily"]   = []
        metrics["stock_data_available"]  = False
        metrics["expenses_detail_stock"] = {}

    # --------------------------------------------------------
    # SOURCE 3 — Shop Sales (reader_shop.py)
    # --------------------------------------------------------
    print("\n[3/4] Reading Shop Sales...")
    try:
        shop = get_monthly_shop_metrics(SHOP_FILE, month, year)
        metrics["shop_sales_detail"] = shop.get("shop_sales_detail", {})
        print(f"  ✅ Shop metrics loaded")
    except Exception as e:
        print(f"  ⚠️  Shop read failed (non-fatal): {e}")
        metrics["shop_sales_detail"] = {}

    # --------------------------------------------------------
    # SOURCE 4 — Manager's Report (reader_manager.py)
    # --------------------------------------------------------
    print("\n[4/4] Reading Manager's Report...")
    try:
        mgr = get_all_manager_metrics(MANAGER_FILE)
        metrics["pnl"]                = mgr.get("pnl", {})
        metrics["debtors"]            = mgr.get("debtors", {})
        metrics["depositors"]         = mgr.get("depositors", {})
        metrics["financial_position"] = mgr.get("financial_position", {})
        metrics["claims"]             = mgr.get("claims", {})
        metrics["money_meters"]       = mgr.get("money_meters", {})
        print(f"  ✅ Manager metrics loaded")
    except Exception as e:
        print(f"  ⚠️  Manager report read failed (non-fatal): {e}")
        metrics["pnl"]                = {}
        metrics["debtors"]            = {}
        metrics["depositors"]         = {}
        metrics["financial_position"] = {}
        metrics["claims"]             = {}
        metrics["money_meters"]       = {}

    # --------------------------------------------------------
    # METADATA
    # --------------------------------------------------------
    metrics.update({
        "report_type":   "monthly",
        "report_month":  month,
        "report_year":   year,
        "period_label":  period_label,
        "station_id":    station_config["station_id"],
        "station_name":  station_config["station_name"],
        "location":      station_config.get("location", ""),
        "currency":      station_config.get("currency", "UGX"),
        "generated_at":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })

    print(f"\n{'='*55}")
    print(f"  ✅ Processing complete: {period_label}")
    print(f"  Total revenue:    UGX {metrics.get('total_revenue', 0):,.0f}")
    print(f"  Shop turnover:    UGX {metrics.get('shop_sales_total', 0):,.0f}")
    print(f"  Total expenses:   UGX {metrics.get('total_expenses', 0):,.0f}")
    print(f"{'='*55}\n")

    return metrics


# ============================================================
# ANNUAL REPORT PROCESSOR (Uganda FY: July–June)
# ============================================================

def process_annual_report(fy_start_year: int, station_config: dict) -> dict:
    """
    Processes a full Uganda financial year (July 1 – June 30).

    EXPENSE RULE: monthly_breakdown expenses come from the stock file
    Daily Expenses sheet — same source as monthly reports. This ensures
    the annual report's monthly expense column matches each monthly report.

    Args:
        fy_start_year:  The year the FY starts. E.g. 2025 = FY2025/26
        station_config: Dict from load_station_config()

    Returns:
        Complete annual metrics dict ready for ai_engine and exporter.
    """
    files         = station_config["files"]
    CASHFLOW_FILE = files["cashflow"]
    STOCK_FILE    = files["stock"]

    fy_label     = f"FY {fy_start_year}/{str(fy_start_year + 1)[-2:]}"
    period_label = f"Annual Report {fy_label}"

    print(f"\n{'='*55}")
    print(f"  Processing: {period_label}")
    print(f"  Station:    {station_config['station_name']}")
    print(f"{'='*55}")

    metrics = {}

    # ── Build the 12 months in the FY ───────────────────────
    fy_months = []
    for m in range(7, 13):
        fy_months.append((m, fy_start_year))
    for m in range(1, 7):
        fy_months.append((m, fy_start_year + 1))

    # ── Step 1: Aggregate cashflow across all 12 months ─────
    # Cashflow gives us: volumes, revenues, payment splits,
    # reconciliation figures. NOT expenses (stock file only).
    print("\n[1/2] Aggregating Daily Cash Flow across FY...")

    annual_cashflow = {
        "pms_volume_total":      0,
        "ago_volume_total":      0,
        "pms_revenue_total":     0,
        "ago_revenue_total":     0,
        "total_fuel_revenue":    0,
        "total_sales":           0,
        "cash_collected":        0,
        "cashless_collected":    0,
        "plus_card_total":       0,
        "visa_total":            0,
        "credit_sales_total":    0,
        "total_cash_banked":     0,
        "total_delta":           0,
        "total_days":            0,
    }

    # monthly_cashflow_breakdown carries cashflow-side figures per month
    monthly_cashflow_breakdown = []

    try:
        df_all = read_daily_cashflow(CASHFLOW_FILE)
        df_all["date"] = pd.to_datetime(df_all["date"], errors="coerce")

        months_with_data = 0

        for (month, year) in fy_months:
            mask     = (df_all["date"].dt.month == month) & (df_all["date"].dt.year == year)
            df_month = df_all[mask].copy()

            if df_month.empty:
                monthly_cashflow_breakdown.append({
                    "month": month, "year": year,
                    "label": pd.Timestamp(year=year, month=month, day=1).strftime("%b %Y"),
                    "data_available": False,
                })
                continue

            months_with_data += 1
            total_sales_m = df_month["cashless_total"].sum()
            cash_m        = df_month["total_cash"].sum()
            cashless_m    = total_sales_m - cash_m
            fuel_rev_m    = df_month["pms_revenue"].sum() + df_month["ago_revenue"].sum()

            monthly_cashflow_breakdown.append({
                "month":          month,
                "year":           year,
                "label":          pd.Timestamp(year=year, month=month, day=1).strftime("%b %Y"),
                "data_available": True,
                "days":           len(df_month),
                "pms_volume":     df_month["pms_volume"].sum(),
                "ago_volume":     df_month["ago_volume"].sum(),
                "fuel_revenue":   fuel_rev_m,
                "total_sales":    total_sales_m,
                "cash":           cash_m,
                "cashless":       cashless_m,
                # NOTE: expenses column is NOT set here from cashflow.
                # It will be filled in from the stock file below.
                "expenses":       0,
                "delta":          df_month["delta"].sum(),
            })

            annual_cashflow["pms_volume_total"]   += df_month["pms_volume"].sum()
            annual_cashflow["ago_volume_total"]   += df_month["ago_volume"].sum()
            annual_cashflow["pms_revenue_total"]  += df_month["pms_revenue"].sum()
            annual_cashflow["ago_revenue_total"]  += df_month["ago_revenue"].sum()
            annual_cashflow["total_fuel_revenue"] += fuel_rev_m
            annual_cashflow["total_sales"]        += total_sales_m
            annual_cashflow["cash_collected"]     += cash_m
            annual_cashflow["cashless_collected"] += cashless_m
            annual_cashflow["plus_card_total"]    += df_month["plus_card_payment_total"].sum()
            annual_cashflow["visa_total"]         += df_month["visa"].sum()
            annual_cashflow["credit_sales_total"] += df_month["credit_sales"].sum()
            annual_cashflow["total_cash_banked"]  += df_month["cash_to_bank"].sum()
            annual_cashflow["total_delta"]        += df_month["delta"].sum()
            annual_cashflow["total_days"]         += len(df_month)

        ts = annual_cashflow["total_sales"]
        annual_cashflow["cash_percentage"]     = round(annual_cashflow["cash_collected"] / ts * 100, 1) if ts > 0 else 0
        annual_cashflow["cashless_percentage"] = round(annual_cashflow["cashless_collected"] / ts * 100, 1) if ts > 0 else 0

        td = annual_cashflow["total_delta"]
        annual_cashflow["delta_status"] = "SURPLUS" if td > 0 else ("DEFICIT" if td < 0 else "BALANCED")

        avg_monthly_fuel = (
            annual_cashflow["total_fuel_revenue"] / months_with_data
            if months_with_data > 0 else 0
        )
        annual_cashflow["avg_monthly_fuel_revenue"] = avg_monthly_fuel
        annual_cashflow["months_with_data"]         = months_with_data

        metrics.update(annual_cashflow)
        print(f"  ✅ Cashflow aggregated — {months_with_data}/12 months with data")

    except Exception as e:
        print(f"  ❌ Annual cashflow aggregation failed: {e}")

    # ── Step 2: Aggregate stock data across FY ──────────────
    # Stock file provides: fuel stock, product sales, expenses.
    # Expenses here are the AUTHORITATIVE annual total.
    print("\n[2/2] Aggregating Stock Movement across FY...")
    try:
        stock_annual = get_annual_stock_metrics(STOCK_FILE, fy_start_year)

        metrics["fuel_stock"]             = stock_annual.get("fuel_stock", {})
        metrics["stock_data_available"]   = stock_annual.get("data_available", False)
        metrics["stock_months_available"] = stock_annual.get("months_available", 0)

        # Product sales (non-fuel) — stock file is more accurate
        ps = stock_annual.get("product_sales", {})
        metrics["product_sales"] = ps
        if ps and ps.get("total_ugx", 0) > 0:
            metrics["lubes_revenue_total"]   = ps.get("lubricants_ugx", 0)
            metrics["lpg_revenue_total"]     = ps.get("lpg_ugx", 0)
            metrics["shop_sales_total"]      = ps.get("shop_ugx", 0)
            metrics["tba_revenue_total"]     = ps.get("tba_ugx", 0)
            metrics["lpg_accessories_total"] = ps.get("lpg_accessories_ugx", 0)
            metrics["car_wash_total"]        = ps.get("car_wash_ugx", 0)
            metrics["total_nonfuel_revenue"] = (
                metrics["lubes_revenue_total"]   +
                metrics["lpg_revenue_total"]     +
                metrics["shop_sales_total"]      +
                metrics["tba_revenue_total"]     +
                metrics["lpg_accessories_total"] +
                metrics["car_wash_total"]
            )
            metrics["total_revenue"] = (
                metrics.get("total_fuel_revenue", 0) +
                metrics["total_nonfuel_revenue"]
            )
        else:
            metrics["lubes_revenue_total"]   = 0
            metrics["lpg_revenue_total"]     = 0
            metrics["shop_sales_total"]      = 0
            metrics["total_nonfuel_revenue"] = 0
            metrics["total_revenue"]         = metrics.get("total_fuel_revenue", 0)

        # EXPENSES — stock file only, authoritative for annual report
        exp = stock_annual.get("expenses_detail", {})
        metrics["expenses_detail_stock"] = exp
        metrics["total_expenses"]        = exp.get("total_expenses", 0)
        metrics["total_expenses_stock"]  = exp.get("total_expenses", 0)

        # ── Merge stock expenses into the monthly cashflow breakdown ──
        # This is the fix for "Annual expenses don't match Monthly report":
        # The monthly breakdown table in the annual report now shows
        # expenses from the stock file — exactly the same source as
        # each individual monthly report.
        stock_monthly_breakdown = stock_annual.get("monthly_breakdown", [])
        stock_exp_by_label = {
            row["label"]: row.get("total_expenses", 0)
            for row in stock_monthly_breakdown
            if row.get("data_available", False)
        }
        for row in monthly_cashflow_breakdown:
            label = row["label"]
            if label in stock_exp_by_label:
                row["expenses"] = stock_exp_by_label[label]
            # If stock file has no data for this month, expenses stays 0

        print(f"  ✅ Stock aggregated — {stock_annual.get('months_available', 0)}/12 months")
        print(f"     Annual expenses (stock):   UGX {metrics['total_expenses']:,.0f}")
        print(f"     Annual shop total (stock): UGX {metrics.get('shop_sales_total', 0):,.0f}")

    except Exception as e:
        print(f"  ⚠️  Annual stock aggregation failed (non-fatal): {e}")
        metrics["fuel_stock"]            = {}
        metrics["product_sales"]         = {}
        metrics["stock_data_available"]  = False
        metrics["expenses_detail_stock"] = {}
        metrics["total_expenses"]        = 0

    # ── Build final monthly breakdown for exporter ───────────
    # This is what appears in the annual report's monthly table.
    # fuel_revenue, total_sales, cash, delta from cashflow.
    # expenses from stock file (filled in above).
    metrics["monthly_breakdown"] = monthly_cashflow_breakdown

    # ── Metadata ─────────────────────────────────────────────
    metrics.update({
        "report_type":    "annual",
        "fy_start_year":  fy_start_year,
        "fy_label":       fy_label,
        "period_label":   f"{fy_label} (July {fy_start_year} – June {fy_start_year + 1})",
        "station_id":     station_config["station_id"],
        "station_name":   station_config["station_name"],
        "location":       station_config.get("location", ""),
        "currency":       station_config.get("currency", "UGX"),
        "generated_at":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "daily_df":       pd.DataFrame(),   # not used for annual
        "general_sales_daily": [],          # not used for annual
    })

    print(f"\n{'='*55}")
    print(f"  ✅ Annual processing complete: {fy_label}")
    print(f"  Total fuel revenue:  UGX {metrics.get('total_fuel_revenue', 0):,.0f}")
    print(f"  Total expenses:      UGX {metrics.get('total_expenses', 0):,.0f}")
    print(f"{'='*55}\n")

    return metrics


# ============================================================
# LAST-12-MONTHS REPORT PROCESSOR
# Period: the 12 calendar months ending on the latest month
# that has data in the cashflow file.
# ============================================================

def process_last12_months_report(station_config: dict) -> dict:
    """
    Processes the 12 most recent calendar months for which the cashflow
    file contains at least one record.

    Unlike process_annual_report(), this is not tied to July–June.
    The period label is derived automatically (e.g. "Jul 2025 – Jun 2026").

    Returns the same metrics dict shape as process_annual_report().
    """
    files         = station_config["files"]
    CASHFLOW_FILE = files["cashflow"]
    STOCK_FILE    = files["stock"]

    # ── Detect the 12 months to cover ──────────────────────────
    try:
        df_all = read_daily_cashflow(CASHFLOW_FILE)
        df_all["date"] = pd.to_datetime(df_all["date"], errors="coerce")
        df_all = df_all.dropna(subset=["date"])
        if df_all.empty:
            raise ValueError("No data in cashflow file.")

        latest_date  = df_all["date"].max()
        end_month    = int(latest_date.month)
        end_year     = int(latest_date.year)

        # Build 12 months ending on (end_month, end_year)
        target_months = []
        for i in range(11, -1, -1):
            m = end_month - i
            y = end_year
            while m < 1:
                m += 12
                y -= 1
            target_months.append((m, y))

        start_m, start_y = target_months[0]
        period_label = (
            f"{pd.Timestamp(year=start_y, month=start_m, day=1).strftime('%b %Y')}"
            f" – "
            f"{pd.Timestamp(year=end_year, month=end_month, day=1).strftime('%b %Y')}"
        )
        fy_label = f"Last 12 Months ({period_label})"

    except Exception as e:
        raise RuntimeError(f"Could not determine date range from cashflow file: {e}")

    print(f"\n{'='*55}")
    print(f"  Processing: Last 12 Months")
    print(f"  Period:     {period_label}")
    print(f"  Station:    {station_config['station_name']}")
    print(f"{'='*55}")

    metrics = {}

    annual_cashflow = {
        "pms_volume_total":      0,
        "ago_volume_total":      0,
        "pms_revenue_total":     0,
        "ago_revenue_total":     0,
        "total_fuel_revenue":    0,
        "total_sales":           0,
        "cash_collected":        0,
        "cashless_collected":    0,
        "plus_card_total":       0,
        "visa_total":            0,
        "credit_sales_total":    0,
        "total_cash_banked":     0,
        "total_delta":           0,
        "total_days":            0,
    }

    monthly_cashflow_breakdown = []
    months_with_data = 0

    print("\n[1/2] Aggregating Daily Cash Flow across last 12 months...")
    try:
        for (month, year) in target_months:
            mask     = (df_all["date"].dt.month == month) & (df_all["date"].dt.year == year)
            df_month = df_all[mask].copy()

            if df_month.empty:
                monthly_cashflow_breakdown.append({
                    "month": month, "year": year,
                    "label": pd.Timestamp(year=year, month=month, day=1).strftime("%b %Y"),
                    "data_available": False,
                })
                continue

            months_with_data += 1
            total_sales_m = df_month["cashless_total"].sum()
            cash_m        = df_month["total_cash"].sum()
            cashless_m    = total_sales_m - cash_m
            fuel_rev_m    = df_month["pms_revenue"].sum() + df_month["ago_revenue"].sum()

            monthly_cashflow_breakdown.append({
                "month":          month,
                "year":           year,
                "label":          pd.Timestamp(year=year, month=month, day=1).strftime("%b %Y"),
                "data_available": True,
                "days":           len(df_month),
                "pms_volume":     df_month["pms_volume"].sum(),
                "ago_volume":     df_month["ago_volume"].sum(),
                "fuel_revenue":   fuel_rev_m,
                "total_sales":    total_sales_m,
                "cash":           cash_m,
                "cashless":       cashless_m,
                "expenses":       0,
                "delta":          df_month["delta"].sum(),
            })

            annual_cashflow["pms_volume_total"]   += df_month["pms_volume"].sum()
            annual_cashflow["ago_volume_total"]   += df_month["ago_volume"].sum()
            annual_cashflow["pms_revenue_total"]  += df_month["pms_revenue"].sum()
            annual_cashflow["ago_revenue_total"]  += df_month["ago_revenue"].sum()
            annual_cashflow["total_fuel_revenue"] += fuel_rev_m
            annual_cashflow["total_sales"]        += total_sales_m
            annual_cashflow["cash_collected"]     += cash_m
            annual_cashflow["cashless_collected"] += cashless_m
            annual_cashflow["plus_card_total"]    += df_month["plus_card_payment_total"].sum()
            annual_cashflow["visa_total"]         += df_month["visa"].sum()
            annual_cashflow["credit_sales_total"] += df_month["credit_sales"].sum()
            annual_cashflow["total_cash_banked"]  += df_month["cash_to_bank"].sum()
            annual_cashflow["total_delta"]        += df_month["delta"].sum()
            annual_cashflow["total_days"]         += len(df_month)

        ts = annual_cashflow["total_sales"]
        annual_cashflow["cash_percentage"]     = round(annual_cashflow["cash_collected"] / ts * 100, 1) if ts > 0 else 0
        annual_cashflow["cashless_percentage"] = round(annual_cashflow["cashless_collected"] / ts * 100, 1) if ts > 0 else 0

        td = annual_cashflow["total_delta"]
        annual_cashflow["delta_status"]              = "SURPLUS" if td > 0 else ("DEFICIT" if td < 0 else "BALANCED")
        annual_cashflow["avg_monthly_fuel_revenue"]  = (
            annual_cashflow["total_fuel_revenue"] / months_with_data if months_with_data > 0 else 0
        )
        annual_cashflow["months_with_data"] = months_with_data

        metrics.update(annual_cashflow)
        print(f"  Cashflow aggregated — {months_with_data}/12 months with data")

    except Exception as e:
        print(f"  Annual cashflow aggregation failed: {e}")

    # ── Step 2: Stock data — cover all 12 months ────────────────
    # get_annual_stock_metrics expects a fy_start_year (July–June).
    # For last-12-months we call it for each relevant FY year that
    # overlaps with our window, then filter by the target months.
    print("\n[2/2] Aggregating Stock Movement across last 12 months...")
    try:
        # Collect unique (month, year) pairs that have cashflow data
        stock_exp_by_label = {}
        stock_data_available = False
        lubes_total = lpg_total = shop_total = tba_total = lpg_acc_total = car_wash_total = 0

        fy_years_needed = set()
        for (month, year) in target_months:
            # A month belongs to FY starting in (year-1) if month <= 6
            fy_years_needed.add(year - 1 if month <= 6 else year)

        for fy_yr in fy_years_needed:
            try:
                stock_annual = get_annual_stock_metrics(STOCK_FILE, fy_yr)
                for row in stock_annual.get("monthly_breakdown", []):
                    if row.get("data_available"):
                        stock_exp_by_label[row["label"]] = row.get("total_expenses", 0)
                        stock_data_available = True

                ps = stock_annual.get("product_sales", {})
                if ps.get("total_ugx", 0) > 0:
                    lubes_total   += ps.get("lubricants_ugx", 0)
                    lpg_total     += ps.get("lpg_ugx", 0)
                    shop_total    += ps.get("shop_ugx", 0)
                    tba_total     += ps.get("tba_ugx", 0)
                    lpg_acc_total += ps.get("lpg_accessories_ugx", 0)
                    car_wash_total+= ps.get("car_wash_ugx", 0)
            except Exception:
                pass

        # Merge expenses into monthly breakdown
        for row in monthly_cashflow_breakdown:
            if row["label"] in stock_exp_by_label:
                row["expenses"] = stock_exp_by_label[row["label"]]

        total_expenses = sum(stock_exp_by_label.get(r["label"], 0) for r in monthly_cashflow_breakdown)

        metrics["stock_data_available"]  = stock_data_available
        metrics["expenses_detail_stock"] = {}
        metrics["total_expenses"]        = total_expenses
        metrics["total_expenses_stock"]  = total_expenses
        metrics["lubes_revenue_total"]   = lubes_total
        metrics["lpg_revenue_total"]     = lpg_total
        metrics["shop_sales_total"]      = shop_total
        metrics["tba_revenue_total"]     = tba_total
        metrics["lpg_accessories_total"] = lpg_acc_total
        metrics["car_wash_total"]        = car_wash_total
        metrics["total_nonfuel_revenue"] = (
            lubes_total + lpg_total + shop_total + tba_total + lpg_acc_total + car_wash_total
        )
        metrics["total_revenue"] = (
            metrics.get("total_fuel_revenue", 0) + metrics["total_nonfuel_revenue"]
        )
        metrics["product_sales"] = {
            "lubricants_ugx":      lubes_total,
            "lpg_ugx":             lpg_total,
            "shop_ugx":            shop_total,
            "tba_ugx":             tba_total,
            "lpg_accessories_ugx": lpg_acc_total,
            "car_wash_ugx":        car_wash_total,
            "total_ugx":           metrics["total_nonfuel_revenue"],
        }
        print(f"  Stock aggregated across {len(fy_years_needed)} FY period(s)")

    except Exception as e:
        print(f"  Stock aggregation failed (non-fatal): {e}")
        metrics["stock_data_available"]  = False
        metrics["expenses_detail_stock"] = {}
        metrics["total_expenses"]        = 0
        metrics["total_nonfuel_revenue"] = 0
        metrics["total_revenue"]         = metrics.get("total_fuel_revenue", 0)

    metrics["monthly_breakdown"] = monthly_cashflow_breakdown

    metrics.update({
        "report_type":   "annual",
        "fy_start_year": None,
        "fy_label":      fy_label,
        "period_label":  period_label,
        "station_id":    station_config["station_id"],
        "station_name":  station_config["station_name"],
        "location":      station_config.get("location", ""),
        "currency":      station_config.get("currency", "UGX"),
        "generated_at":  datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "daily_df":      pd.DataFrame(),
        "general_sales_daily": [],
    })

    print(f"\n{'='*55}")
    print(f"  Last-12-months processing complete")
    print(f"  Period:             {period_label}")
    print(f"  Total fuel revenue: UGX {metrics.get('total_fuel_revenue', 0):,.0f}")
    print(f"  Total expenses:     UGX {metrics.get('total_expenses', 0):,.0f}")
    print(f"{'='*55}\n")

    return metrics


# ============================================================
# BUILT-IN TEST
# Run with: python -m src.processor
# ============================================================
if __name__ == "__main__":
    from config.station_loader import load_station_config

    config = load_station_config("te_rwizi")

    print("\n" + "="*55)
    print("TEST: Monthly — April 2026")
    print("="*55)
    m = process_monthly_report(4, 2026, config)
    print(f"  Shop turnover:      UGX {m.get('shop_sales_total', 0):,.0f}  (expected: 17,426,000)")
    print(f"  Total expenses:     UGX {m.get('total_expenses', 0):,.0f}  (expected: 9,182,900)")
    print(f"  Daily product rows: {len(m.get('general_sales_daily', []))}  (expected: 30)")

    print("\n" + "="*55)
    print("TEST: Annual — FY 2025/26")
    print("="*55)
    a = process_annual_report(2025, config)
    print(f"  Annual expenses:    UGX {a.get('total_expenses', 0):,.0f}")
    print()
    print("  Monthly breakdown (expenses should match each monthly report):")
    for row in a.get("monthly_breakdown", []):
        mark = "OK" if row["data_available"] else "--"
        print(f"    [{mark}] {row['label']:12}  "
              f"fuel: UGX {row.get('fuel_revenue', 0):>15,.0f}  "
              f"exp: UGX {row.get('expenses', 0):>12,.0f}")