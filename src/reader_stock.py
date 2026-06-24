"""
reader_stock.py
---------------
Reads fuel stock movement and product sales data from the
Stock Movement Template Excel file.

Verified column positions against real April/May 2026 data.

Functions:
  read_fuel_stock_movement()    — daily PMS/AGO movement with deliveries
  read_general_sales_summary()  — daily all-product revenue breakdown
  read_shop_sales()             — daily shop sales (Day/Night/Total)
  read_daily_expenses_stock()   — daily expense breakdown
  get_monthly_stock_metrics()   — aggregated monthly dict for the report processor
  get_annual_stock_metrics()    — aggregated annual dict (Uganda FY July–June)
"""

import pandas as pd
from pathlib import Path
from datetime import datetime


# ─────────────────────────────────────────────────────────────
# INTERNAL HELPERS
# ─────────────────────────────────────────────────────────────

def _validate_path(filepath) -> Path:
    p = Path(filepath)
    if not p.exists():
        raise FileNotFoundError(f"Stock template not found: {p}")
    return p


def _is_date_row(val) -> bool:
    """Return True if val is a datetime-like object (valid data row)."""
    return isinstance(val, datetime) or (
        hasattr(val, 'year') and hasattr(val, 'month') and hasattr(val, 'day')
    )


def _safe_float(val) -> float:
    """Convert a cell value to float safely; return 0.0 on failure."""
    try:
        return float(val) if pd.notna(val) else 0.0
    except (ValueError, TypeError):
        return 0.0


# ─────────────────────────────────────────────────────────────
# COLUMN MAP — verified against real data 2026-04 and 2026-05
# ─────────────────────────────────────────────────────────────
#
# PMS sheet: 'Fuel Mvt-2025-2026-PMS ....'
# AGO sheet: 'Fuel Mvt-2025-2026-AGO'
#
# col 0  — Date
# col 1  — Opening Dip (Ltrs)
# col 2  — Purchases (Ltrs)             ← 0 on non-delivery days
# col 5  — Station Sales (Ltrs)         ← USE THIS
# col 8  — Closing DIP (Ltrs)
# col 9  — Expected Balance (Ltrs)
# col 10 — Loss/Gain (Ltrs)
# col 11 — Cost Price (UGX/Ltr)
# col 12 — Selling Price (UGX/Ltr)
# col 13 — Purchase Value (UGX)
# col 14 — Turnover (UGX)
# col 15 — Stock Value (UGX)
# col 16 — Loss/Gain Value (UGX)

FUEL_COLS = {
    "date":               0,
    "opening_dip_ltrs":   1,
    "purchases_ltrs":     2,
    "station_sales_ltrs": 5,
    "closing_dip_ltrs":   8,
    "expected_balance_ltrs": 9,
    "loss_gain_ltrs":    10,
    "cost_price_ugx":    11,
    "selling_price_ugx": 12,
    "purchase_value_ugx":13,
    "turnover_ugx":      14,
    "stock_value_ugx":   15,
    "loss_gain_value_ugx":16,
}

#
# General Sales Summary sheet: 'General Sales Sumary'
# col 0  — Date
# col 1  — PMS (UGX)
# col 2  — AGO (UGX)
# col 3  — Lubricants (UGX)
# col 4  — TBA (UGX)
# col 5  — LPG Accessories (UGX)
# col 6  — LPG (UGX)
# col 7  — Car Wash (UGX)
# col 8  — (blank separator)
# col 9  — Shop (UGX)               ← matches Shop Sales sheet total
# col 10 — Solar Lanterns (UGX)
# col 11 — General Total (UGX)

GENERAL_SALES_COLS = {
    "date":               0,
    "pms_ugx":            1,
    "ago_ugx":            2,
    "lubricants_ugx":     3,
    "tba_ugx":            4,
    "lpg_accessories_ugx":5,
    "lpg_ugx":            6,
    "car_wash_ugx":       7,
    "shop_ugx":           9,
    "solar_ugx":         10,
    "total_ugx":         11,
}

#
# Shop Sales sheet: 'Shop Sales'
# col 0  — Date
# col 1  — Day Shift
# col 2  — Night Shift
# col 3  — Total                     ← USE THIS for shop_turnover

SHOP_SALES_COLS = {
    "date":        0,
    "day_shift":   1,
    "night_shift": 2,
    "total":       3,
}

#
# Daily Expenses sheet: 'Daily Expenses'
# col 0  — Date
# col 1  — Meals
# col 2  — Generator
# col 3  — Electricity
# col 4  — Water
# col 5  — Salaries
# col 6  — Stationary
# col 7  — Security
# col 8  — Sanitation
# col 9  — Airtime/Data
# col 10 — Transport
# col 11 — NSSF
# col 12 — Sundries
# col 13 — Maintenance
# col 14 — VAT/Tax
# col 31 — TOTAL  ← verified position

EXPENSE_COLS = {
    "date":         0,
    "meals":        1,
    "generator":    2,
    "electricity":  3,
    "water":        4,
    "salaries":     5,
    "stationary":   6,
    "security":     7,
    "sanitation":   8,
    "airtime":      9,
    "transport":   10,
    "nssf":        11,
    "sundries":    12,
    "maintenance": 13,
    "vat_tax":     14,
}
EXPENSE_TOTAL_COL = 31


# ─────────────────────────────────────────────────────────────
# FUNCTION 1 — FUEL STOCK MOVEMENT (PMS + AGO)
# ─────────────────────────────────────────────────────────────

def read_fuel_stock_movement(filepath) -> pd.DataFrame:
    """
    Reads daily fuel stock movement for both PMS and AGO.

    Returns a DataFrame with columns:
      date, product, opening_dip_ltrs, purchases_ltrs, station_sales_ltrs,
      closing_dip_ltrs, expected_balance_ltrs, loss_gain_ltrs,
      cost_price_ugx, selling_price_ugx, purchase_value_ugx,
      turnover_ugx, stock_value_ugx, loss_gain_value_ugx,
      is_delivery_day (bool)
    """
    path = _validate_path(filepath)

    sheet_map = {
        "PMS": "Fuel Mvt-2025-2026-PMS ....",
        "AGO": "Fuel Mvt-2025-2026-AGO",
    }

    records = []

    for product, sheet_name in sheet_map.items():
        try:
            df_raw = pd.read_excel(
                path, sheet_name=sheet_name,
                header=None, engine="openpyxl",
            )
        except Exception as e:
            print(f"  [WARNING] Could not read sheet '{sheet_name}': {e}")
            continue

        for _, row in df_raw.iterrows():
            date_val = row.iloc[FUEL_COLS["date"]] if len(row) > 0 else None
            if not _is_date_row(date_val):
                continue

            record = {
                "date":    pd.Timestamp(date_val),
                "product": product,
            }

            for col_name, idx in FUEL_COLS.items():
                if col_name == "date":
                    continue
                val = row.iloc[idx] if idx < len(row) else None
                record[col_name] = _safe_float(val)

            record["is_delivery_day"] = record["purchases_ltrs"] > 0
            records.append(record)

    if not records:
        print("  [WARNING] No fuel stock movement data found.")
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values(["product", "date"]).reset_index(drop=True)

    numeric_cols = [
        c for c in df.columns
        if c not in ("date", "product", "is_delivery_day")
    ]
    df[numeric_cols] = (
        df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    )

    return df


# ─────────────────────────────────────────────────────────────
# FUNCTION 2 — GENERAL SALES SUMMARY (ALL PRODUCTS, DAILY)
# ─────────────────────────────────────────────────────────────

def read_general_sales_summary(filepath) -> pd.DataFrame:
    """
    Reads the daily general sales summary covering all product lines.

    Returns a DataFrame with columns:
      date, pms_ugx, ago_ugx, lubricants_ugx, tba_ugx,
      lpg_accessories_ugx, lpg_ugx, car_wash_ugx,
      shop_ugx, solar_ugx, total_ugx

    NOTE: shop_ugx here matches the Shop Sales sheet total exactly.
    This is the authoritative source for shop_turnover.
    """
    path = _validate_path(filepath)

    try:
        df_raw = pd.read_excel(
            path, sheet_name="General Sales Sumary",
            header=None, engine="openpyxl",
        )
    except Exception as e:
        raise RuntimeError(f"Could not read General Sales Sumary sheet: {e}")

    records = []
    for _, row in df_raw.iterrows():
        date_val = row.iloc[GENERAL_SALES_COLS["date"]] if len(row) > 0 else None
        if not _is_date_row(date_val):
            continue

        record = {"date": pd.Timestamp(date_val)}
        for col_name, idx in GENERAL_SALES_COLS.items():
            if col_name == "date":
                continue
            val = row.iloc[idx] if idx < len(row) else None
            record[col_name] = _safe_float(val)

        records.append(record)

    if not records:
        print("  [WARNING] No general sales summary data found.")
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)

    numeric_cols = [c for c in df.columns if c != "date"]
    df[numeric_cols] = (
        df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    )

    return df


# ─────────────────────────────────────────────────────────────
# FUNCTION 3 — SHOP SALES (DAY/NIGHT/TOTAL, DAILY)
# ─────────────────────────────────────────────────────────────

def read_shop_sales(filepath) -> pd.DataFrame:
    """
    Reads daily shop sales from the Shop Sales sheet.

    Columns returned: date, day_shift, night_shift, total

    The 'total' column is the authoritative shop turnover figure.
    It matches the shop_ugx column in General Sales Summary exactly,
    so we use General Sales Summary's shop_ugx as the single source
    in monthly metrics. This function is available for detailed
    shift-level analysis if needed in future.
    """
    path = _validate_path(filepath)

    try:
        df_raw = pd.read_excel(
            path, sheet_name="Shop Sales",
            header=None, engine="openpyxl",
        )
    except Exception as e:
        raise RuntimeError(f"Could not read Shop Sales sheet: {e}")

    records = []
    for _, row in df_raw.iterrows():
        date_val = row.iloc[SHOP_SALES_COLS["date"]] if len(row) > 0 else None
        if not _is_date_row(date_val):
            continue

        # Skip subtotal rows — they have a date-like cell from merged cells
        # Real data rows always have a numeric total in col 3
        total_val = row.iloc[SHOP_SALES_COLS["total"]] if len(row) > 3 else None
        total = _safe_float(total_val)

        # Skip rows that look like subtotals (very large values relative to one day)
        # A single day's shop sales should be under 5,000,000 UGX
        if total > 5_000_000:
            continue

        record = {
            "date":        pd.Timestamp(date_val),
            "day_shift":   _safe_float(row.iloc[SHOP_SALES_COLS["day_shift"]]  if len(row) > 1 else None),
            "night_shift": _safe_float(row.iloc[SHOP_SALES_COLS["night_shift"]] if len(row) > 2 else None),
            "total":       total,
        }
        records.append(record)

    if not records:
        print("  [WARNING] No shop sales data found.")
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)

    numeric_cols = [c for c in df.columns if c != "date"]
    df[numeric_cols] = (
        df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    )

    return df


# ─────────────────────────────────────────────────────────────
# FUNCTION 4 — DAILY EXPENSES FROM STOCK TEMPLATE
# ─────────────────────────────────────────────────────────────

def read_daily_expenses_stock(filepath) -> pd.DataFrame:
    """
    Reads the daily expenses breakdown from the stock template.

    Returns a DataFrame with columns:
      date, meals, generator, electricity, water, salaries,
      stationary, security, sanitation, airtime, transport,
      nssf, sundries, maintenance, vat_tax, total_expenses

    This is the SINGLE authoritative source for all expense figures —
    both monthly and annual reports must use this, never cashflow col37.
    """
    path = _validate_path(filepath)

    try:
        df_raw = pd.read_excel(
            path, sheet_name="Daily Expenses",
            header=None, engine="openpyxl",
        )
    except Exception as e:
        raise RuntimeError(f"Could not read Daily Expenses sheet: {e}")

    records = []
    for _, row in df_raw.iterrows():
        date_val = row.iloc[EXPENSE_COLS["date"]] if len(row) > 0 else None
        if not _is_date_row(date_val):
            continue

        record = {"date": pd.Timestamp(date_val)}
        for col_name, idx in EXPENSE_COLS.items():
            if col_name == "date":
                continue
            val = row.iloc[idx] if idx < len(row) else None
            record[col_name] = _safe_float(val)

        total_val = (
            row.iloc[EXPENSE_TOTAL_COL]
            if EXPENSE_TOTAL_COL < len(row) else None
        )
        record["total_expenses"] = _safe_float(total_val)

        records.append(record)

    if not records:
        print("  [WARNING] No daily expense data found in stock template.")
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)

    numeric_cols = [c for c in df.columns if c != "date"]
    df[numeric_cols] = (
        df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    )

    return df


# ─────────────────────────────────────────────────────────────
# INTERNAL — FUEL METRICS FOR ONE PRODUCT/PERIOD
# ─────────────────────────────────────────────────────────────

def _fuel_metrics_for_period(df_fuel: pd.DataFrame, product: str) -> dict:
    """
    Given a filtered (already month/year masked) fuel DataFrame,
    compute all metrics for one product (PMS or AGO).
    Returns an empty dict if no data is available.
    """
    p = df_fuel[df_fuel["product"] == product].copy()
    if p.empty:
        return {}

    total_sales      = p["station_sales_ltrs"].sum()
    total_purchases  = p["purchases_ltrs"].sum()
    total_loss_ltrs  = p["loss_gain_ltrs"].sum()
    total_loss_ugx   = p["loss_gain_value_ugx"].sum()
    total_turnover   = p["turnover_ugx"].sum()
    total_purchase_value = p["purchase_value_ugx"].sum()

    opening_dip         = float(p.iloc[0]["opening_dip_ltrs"])
    closing_dip         = float(p.iloc[-1]["closing_dip_ltrs"])
    closing_stock_value = float(p.iloc[-1]["stock_value_ugx"])

    prices = p[p["cost_price_ugx"] > 0]
    avg_cost_price    = prices["cost_price_ugx"].mean()    if not prices.empty else 0
    avg_selling_price = prices["selling_price_ugx"].mean() if not prices.empty else 0

    loss_pct = (total_loss_ltrs / total_sales * 100) if total_sales != 0 else 0

    deliveries = p[p["is_delivery_day"]].copy()
    delivery_list = []
    for _, row in deliveries.iterrows():
        delivery_list.append({
            "date":        row["date"].strftime("%d %b %Y"),
            "litres":      round(row["purchases_ltrs"], 0),
            "cost_price":  round(row["cost_price_ugx"], 0),
            "total_value": round(row["purchase_value_ugx"], 0),
        })

    return {
        "opening_dip_ltrs":        round(opening_dip, 2),
        "closing_dip_ltrs":        round(closing_dip, 2),
        "total_purchases_ltrs":    round(total_purchases, 2),
        "total_sales_ltrs":        round(total_sales, 2),
        "loss_gain_ltrs":          round(total_loss_ltrs, 2),
        "loss_gain_pct":           round(loss_pct, 4),
        "loss_gain_value_ugx":     round(total_loss_ugx, 0),
        "turnover_ugx":            round(total_turnover, 0),
        "purchase_value_ugx":      round(total_purchase_value, 0),
        "closing_stock_value_ugx": round(closing_stock_value, 0),
        "avg_cost_price_ugx":      round(avg_cost_price, 0),
        "avg_selling_price_ugx":   round(avg_selling_price, 0),
        "delivery_count":          len(delivery_list),
        "deliveries":              delivery_list,
    }


# ─────────────────────────────────────────────────────────────
# FUNCTION 5 — MONTHLY AGGREGATION
# ─────────────────────────────────────────────────────────────

def get_monthly_stock_metrics(
    filepath,
    month: int,
    year: int,
) -> dict:
    """
    Aggregates all stock template data for a given month/year into
    a single metrics dictionary ready for the report processor.

    Key additions vs old version:
    - product_sales["shop_ugx"] now comes from General Sales Summary
      col 9, which exactly matches the Shop Sales sheet daily totals.
      This is the authoritative shop_turnover figure.
    - general_sales_daily: list of daily dicts for all non-fuel products,
      used by the exporter to add product columns to Daily Records sheet.
    - expenses come exclusively from Daily Expenses sheet (col 31 total).

    Returns:
    {
      "fuel_stock": {
          "pms": { opening_dip, closing_dip, purchases, sales, loss/gain,
                   turnover, stock_value, avg prices, deliveries list },
          "ago": { same structure }
      },
      "product_sales": {
          "pms_ugx", "ago_ugx", "lubricants_ugx", "tba_ugx",
          "lpg_ugx", "lpg_accessories_ugx", "car_wash_ugx",
          "shop_ugx",          ← authoritative shop turnover
          "solar_ugx", "total_ugx"
      },
      "general_sales_daily": [
          {
            "date": "01 Apr 2026",
            "lubricants_ugx": ..., "lpg_ugx": ..., "lpg_accessories_ugx": ...,
            "tba_ugx": ..., "car_wash_ugx": ..., "shop_ugx": ..., "solar_ugx": ...
          }, ...
      ],
      "expenses_detail": {
          "meals", "generator", "electricity", "water", "salaries",
          "stationary", "security", "sanitation", "airtime", "transport",
          "nssf", "sundries", "maintenance", "vat_tax",
          "total_expenses"     ← from Daily Expenses col 31 (authoritative)
      },
      "data_available": True/False
    }
    """
    path = _validate_path(filepath)
    result = {
        "fuel_stock":          {"pms": {}, "ago": {}},
        "product_sales":       {},
        "general_sales_daily": [],
        "expenses_detail":     {},
        "data_available":      False,
    }

    # ── Fuel stock movement ──────────────────────────────────
    try:
        df_fuel = read_fuel_stock_movement(path)
        if not df_fuel.empty:
            mask = (
                (df_fuel["date"].dt.month == month) &
                (df_fuel["date"].dt.year  == year)
            )
            df_m = df_fuel[mask]

            if not df_m.empty and df_m["station_sales_ltrs"].sum() > 0:
                result["fuel_stock"]["pms"] = _fuel_metrics_for_period(df_m, "PMS")
                result["fuel_stock"]["ago"] = _fuel_metrics_for_period(df_m, "AGO")
                result["data_available"] = True
            else:
                print(f"  [INFO] Fuel stock data not yet entered for {month}/{year}")

    except Exception as e:
        print(f"  [WARNING] Fuel stock aggregation error: {e}")

    # ── General product sales + daily breakdown ──────────────
    try:
        df_sales = read_general_sales_summary(path)
        if not df_sales.empty:
            mask = (
                (df_sales["date"].dt.month == month) &
                (df_sales["date"].dt.year  == year)
            )
            df_m = df_sales[mask]

            if not df_m.empty and df_m["total_ugx"].sum() > 0:

                # Aggregated monthly product totals
                result["product_sales"] = {
                    "pms_ugx":             round(df_m["pms_ugx"].sum(), 0),
                    "ago_ugx":             round(df_m["ago_ugx"].sum(), 0),
                    "lubricants_ugx":      round(df_m["lubricants_ugx"].sum(), 0),
                    "tba_ugx":             round(df_m["tba_ugx"].sum(), 0),
                    "lpg_ugx":             round(df_m["lpg_ugx"].sum(), 0),
                    "lpg_accessories_ugx": round(df_m["lpg_accessories_ugx"].sum(), 0),
                    "car_wash_ugx":        round(df_m["car_wash_ugx"].sum(), 0),
                    # shop_ugx from General Sales Summary = authoritative shop turnover
                    # Verified: matches Shop Sales sheet total exactly for April 2026
                    "shop_ugx":            round(df_m["shop_ugx"].sum(), 0),
                    "solar_ugx":           round(df_m["solar_ugx"].sum(), 0),
                    "total_ugx":           round(df_m["total_ugx"].sum(), 0),
                }
                result["data_available"] = True

                # Daily breakdown for Daily Records XLSX sheet
                # Each row gives non-fuel product revenue per day
                daily_rows = []
                for _, row in df_m.iterrows():
                    daily_rows.append({
                        "date":                row["date"].strftime("%d %b %Y"),
                        "lubricants_ugx":      round(_safe_float(row.get("lubricants_ugx", 0)), 0),
                        "lpg_ugx":             round(_safe_float(row.get("lpg_ugx", 0)), 0),
                        "lpg_accessories_ugx": round(_safe_float(row.get("lpg_accessories_ugx", 0)), 0),
                        "tba_ugx":             round(_safe_float(row.get("tba_ugx", 0)), 0),
                        "car_wash_ugx":        round(_safe_float(row.get("car_wash_ugx", 0)), 0),
                        "shop_ugx":            round(_safe_float(row.get("shop_ugx", 0)), 0),
                        "solar_ugx":           round(_safe_float(row.get("solar_ugx", 0)), 0),
                    })
                result["general_sales_daily"] = daily_rows

    except Exception as e:
        print(f"  [WARNING] General sales aggregation error: {e}")

    # ── Daily expenses (AUTHORITATIVE — always use this, never cashflow) ──
    try:
        df_exp = read_daily_expenses_stock(path)
        if not df_exp.empty:
            mask = (
                (df_exp["date"].dt.month == month) &
                (df_exp["date"].dt.year  == year)
            )
            df_m = df_exp[mask]

            if not df_m.empty:
                exp_cols = [
                    "meals", "generator", "electricity", "water", "salaries",
                    "stationary", "security", "sanitation", "airtime",
                    "transport", "nssf", "sundries", "maintenance", "vat_tax",
                ]
                expenses = {
                    col: round(df_m[col].sum(), 0)
                    for col in exp_cols
                    if col in df_m.columns
                }
                # total_expenses is the sum of col 31 from Daily Expenses sheet.
                # This is the ONLY figure that should appear in any report.
                expenses["total_expenses"] = round(
                    df_m["total_expenses"].sum(), 0
                )
                result["expenses_detail"] = expenses

    except Exception as e:
        print(f"  [WARNING] Expenses aggregation error: {e}")

    return result


# ─────────────────────────────────────────────────────────────
# FUNCTION 6 — ANNUAL AGGREGATION (Uganda FY: July–June)
# ─────────────────────────────────────────────────────────────

def get_annual_stock_metrics(
    filepath,
    fy_start_year: int,
) -> dict:
    """
    Aggregates stock data for a full Uganda financial year.

    Uganda FY runs July 1 – June 30.
    fy_start_year=2025 means July 2025 – June 2026.

    Returns the same structure as get_monthly_stock_metrics()
    but aggregated across all 12 months, plus a monthly_breakdown
    list for charting/tables.

    Also returns "months_available" — count of months that had data.
    """
    path = _validate_path(filepath)

    # Build the 12 months in the financial year
    fy_months = []
    for m in range(7, 13):
        fy_months.append((m, fy_start_year))
    for m in range(1, 7):
        fy_months.append((m, fy_start_year + 1))

    pms_totals = {
        "total_purchases_ltrs": 0, "total_sales_ltrs": 0,
        "loss_gain_ltrs": 0, "loss_gain_value_ugx": 0,
        "turnover_ugx": 0, "purchase_value_ugx": 0,
        "delivery_count": 0,
    }
    ago_totals = dict(pms_totals)

    product_sales_total = {
        "pms_ugx": 0, "ago_ugx": 0, "lubricants_ugx": 0,
        "tba_ugx": 0, "lpg_ugx": 0, "lpg_accessories_ugx": 0,
        "car_wash_ugx": 0, "shop_ugx": 0, "solar_ugx": 0, "total_ugx": 0,
    }

    exp_keys = [
        "meals", "generator", "electricity", "water", "salaries",
        "stationary", "security", "sanitation", "airtime", "transport",
        "nssf", "sundries", "maintenance", "vat_tax", "total_expenses",
    ]
    expenses_total = {k: 0 for k in exp_keys}

    monthly_breakdown = []
    months_available  = 0

    for (month, year) in fy_months:
        m_metrics = get_monthly_stock_metrics(path, month, year)

        if not m_metrics.get("data_available", False):
            monthly_breakdown.append({
                "month": month, "year": year,
                "label": pd.Timestamp(year=year, month=month, day=1).strftime("%b %Y"),
                "data_available": False,
            })
            continue

        months_available += 1

        # Accumulate fuel
        for product, totals in [("pms", pms_totals), ("ago", ago_totals)]:
            fs = m_metrics["fuel_stock"].get(product, {})
            for key in totals:
                totals[key] += fs.get(key, 0)

        # Accumulate product sales
        ps = m_metrics.get("product_sales", {})
        for key in product_sales_total:
            product_sales_total[key] += ps.get(key, 0)

        # Accumulate expenses (stock file only — authoritative)
        exp = m_metrics.get("expenses_detail", {})
        for key in expenses_total:
            expenses_total[key] += exp.get(key, 0)

        monthly_breakdown.append({
            "month":           month,
            "year":            year,
            "label":           pd.Timestamp(year=year, month=month, day=1).strftime("%b %Y"),
            "data_available":  True,
            "pms_sales_ltrs":  m_metrics["fuel_stock"].get("pms", {}).get("total_sales_ltrs", 0),
            "ago_sales_ltrs":  m_metrics["fuel_stock"].get("ago", {}).get("total_sales_ltrs", 0),
            "pms_turnover":    m_metrics["fuel_stock"].get("pms", {}).get("turnover_ugx", 0),
            "ago_turnover":    m_metrics["fuel_stock"].get("ago", {}).get("turnover_ugx", 0),
            "total_sales_ugx": ps.get("total_ugx", 0),
            # total_expenses here is from stock file Daily Expenses sheet
            "total_expenses":  exp.get("total_expenses", 0),
        })

    return {
        "fy_label":          f"FY {fy_start_year}/{str(fy_start_year+1)[-2:]}",
        "fy_start_year":     fy_start_year,
        "months_available":  months_available,
        "fuel_stock": {
            "pms": pms_totals,
            "ago": ago_totals,
        },
        "product_sales":     product_sales_total,
        "expenses_detail":   expenses_total,
        "monthly_breakdown": monthly_breakdown,
        "data_available":    months_available > 0,
    }


# ─────────────────────────────────────────────────────────────
# QUICK TEST  (run this file directly to verify)
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import json

    filepath = r"data\input\Stock mvt-Template-2025-2026-V1-7-25.xlsx"

    print("=" * 60)
    print("TEST: get_monthly_stock_metrics() — April 2026")
    print("=" * 60)
    m = get_monthly_stock_metrics(filepath, month=4, year=2026)
    print(f"  data_available:    {m['data_available']}")
    print(f"  shop_ugx:          {m['product_sales'].get('shop_ugx', 0):,.0f}  (expected: 17,426,000)")
    print(f"  total_expenses:    {m['expenses_detail'].get('total_expenses', 0):,.0f}  (expected: 9,182,900)")
    print(f"  daily rows count:  {len(m['general_sales_daily'])}  (expected: 30)")
    if m['general_sales_daily']:
        print(f"  first daily row:   {m['general_sales_daily'][0]}")

    print()
    print("=" * 60)
    print("TEST: get_annual_stock_metrics() — FY 2025/26")
    print("=" * 60)
    a = get_annual_stock_metrics(filepath, fy_start_year=2025)
    print(f"  FY label:          {a['fy_label']}")
    print(f"  Months available:  {a['months_available']}")
    print(f"  PMS total sales:   {a['fuel_stock']['pms']['total_sales_ltrs']:,.2f} L")
    print(f"  AGO total sales:   {a['fuel_stock']['ago']['total_sales_ltrs']:,.2f} L")
    print(f"  Shop total:        {a['product_sales']['shop_ugx']:,.0f}")
    print(f"  Annual expenses:   {a['expenses_detail']['total_expenses']:,.0f}")
    print()
    print("  Monthly breakdown (expenses from stock file):")
    for row in a["monthly_breakdown"]:
        avail = "OK" if row["data_available"] else "--"
        exp   = row.get("total_expenses", 0)
        print(f"  [{avail}] {row['label']:12}  expenses: {exp:>14,.0f}")