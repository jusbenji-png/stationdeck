"""
reader_stock.py
---------------
Reads fuel stock movement and product sales data from the
Stock Movement Template Excel file.

Provides three functions:
  - read_fuel_stock_movement()  : daily PMS/AGO dips, loss/gain per litre
  - read_general_sales_summary(): daily breakdown of all product revenue
  - read_daily_expenses_stock() : daily expenses from the stock template

All functions return a pandas DataFrame. Call them with the path to
your Stock_mvt-Template file.
"""

import pandas as pd
from pathlib import Path
from datetime import datetime


# ─────────────────────────────────────────────────────────────
# INTERNAL HELPERS
# ─────────────────────────────────────────────────────────────

def _validate_path(filepath: str | Path) -> Path:
    p = Path(filepath)
    if not p.exists():
        raise FileNotFoundError(f"Stock template not found: {p}")
    return p


def _is_valid_date(val) -> bool:
    return isinstance(val, datetime)


# ─────────────────────────────────────────────────────────────
# FUNCTION 1 — FUEL STOCK MOVEMENT (PMS + AGO)
# ─────────────────────────────────────────────────────────────

def read_fuel_stock_movement(filepath: str | Path) -> pd.DataFrame:
    """
    Reads daily fuel stock movement for both PMS and AGO.

    Columns returned:
      date, product, opening_dip_ltrs, purchases_ltrs, net_sales_ltrs,
      closing_dip_ltrs, expected_balance_ltrs, loss_gain_ltrs,
      cost_price_ugx, selling_price_ugx, purchases_value_ugx,
      turnover_ugx, stock_value_ugx, loss_gain_value_ugx

    Source sheets: 'Fuel Mvt-2025-2026-PMS ....' and 'Fuel Mvt-2025-2026-AGO'
    """
    path = _validate_path(filepath)

    sheet_map = {
        "PMS": "Fuel Mvt-2025-2026-PMS ....",
        "AGO": "Fuel Mvt-2025-2026-AGO",
    }

    # Column positions in the sheet (0-indexed), based on header analysis:
    # Date(0), OpeningDip(1), Purchases(2), ReceiptUPDF(3), ReturnToTank(4),
    # StationSales(5), UPDFConsumption(6), NetSales(7), ClosingDIP(8),
    # ExpectedBalance(9), LossGain(10), CostPrice(11), SellingPrice(12),
    # PurchasesValue(13), Turnover(14), StockValue(15), LossGainValue(16)

    COL = {
        "date": 0,
        "opening_dip_ltrs": 1,
        "purchases_ltrs": 2,
        "net_sales_ltrs": 7,
        "closing_dip_ltrs": 8,
        "expected_balance_ltrs": 9,
        "loss_gain_ltrs": 10,
        "cost_price_ugx": 11,
        "selling_price_ugx": 12,
        "purchases_value_ugx": 13,
        "turnover_ugx": 14,
        "stock_value_ugx": 15,
        "loss_gain_value_ugx": 16,
    }

    records = []

    for product, sheet_name in sheet_map.items():
        try:
            df_raw = pd.read_excel(
                path,
                sheet_name=sheet_name,
                header=None,
                engine="openpyxl",
            )
        except Exception as e:
            print(f"  [WARNING] Could not read sheet '{sheet_name}': {e}")
            continue

        for _, row in df_raw.iterrows():
            date_val = row.iloc[COL["date"]]
            if not _is_valid_date(date_val):
                continue

            record = {"date": pd.Timestamp(date_val), "product": product}
            for col_name, idx in COL.items():
                if col_name == "date":
                    continue
                val = row.iloc[idx] if idx < len(row) else None
                record[col_name] = val if pd.notna(val) else 0.0

            records.append(record)

    if not records:
        print("  [WARNING] No fuel stock movement data found.")
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values(["product", "date"]).reset_index(drop=True)

    # Ensure numeric columns are float
    numeric_cols = [c for c in df.columns if c not in ("date", "product")]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)

    return df


# ─────────────────────────────────────────────────────────────
# FUNCTION 2 — GENERAL SALES SUMMARY (ALL PRODUCTS)
# ─────────────────────────────────────────────────────────────

def read_general_sales_summary(filepath: str | Path) -> pd.DataFrame:
    """
    Reads the daily general sales summary covering all product lines.

    Columns returned:
      date, pms_ugx, ago_ugx, lubricants_ugx, tba_ugx,
      lpg_accessories_ugx, lpg_ugx, car_wash_ugx,
      shop_ugx, solar_ugx, total_ugx

    Source sheet: 'General Sales Sumary'
    """
    path = _validate_path(filepath)

    # Row 0 is the header. Columns (0-indexed):
    # 0=DATE, 1=PMS-Ugx, 2=AGO-Ugx, 3=Lubricants, 4=TBA,
    # 5=LPG Accessories, 6=LPG, 7=Car Wash, 8=None, 9=Shop,
    # 10=Solar Lanterns, 11=General Total

    try:
        df_raw = pd.read_excel(
            path,
            sheet_name="General Sales Sumary",
            header=0,
            engine="openpyxl",
        )
    except Exception as e:
        raise RuntimeError(f"Could not read General Sales Sumary sheet: {e}")

    # The header row has the column names — rename them clearly
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    rename_map = {
        df_raw.columns[0]: "date",
        df_raw.columns[1]: "pms_ugx",
        df_raw.columns[2]: "ago_ugx",
        df_raw.columns[3]: "lubricants_ugx",
        df_raw.columns[4]: "tba_ugx",
        df_raw.columns[5]: "lpg_accessories_ugx",
        df_raw.columns[6]: "lpg_ugx",
        df_raw.columns[7]: "car_wash_ugx",
        df_raw.columns[9]: "shop_ugx",
        df_raw.columns[10]: "solar_ugx",
        df_raw.columns[11]: "total_ugx",
    }
    df_raw = df_raw.rename(columns=rename_map)

    # Keep only the columns we renamed
    keep_cols = list(rename_map.values())
    df = df_raw[[c for c in keep_cols if c in df_raw.columns]].copy()

    # Filter to valid date rows only
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df[df["date"].notna()].copy()

    # Fill numeric NaN with 0
    numeric_cols = [c for c in df.columns if c != "date"]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)

    df = df.sort_values("date").reset_index(drop=True)
    return df


# ─────────────────────────────────────────────────────────────
# FUNCTION 3 — DAILY EXPENSES FROM STOCK TEMPLATE
# ─────────────────────────────────────────────────────────────

def read_daily_expenses_stock(filepath: str | Path) -> pd.DataFrame:
    """
    Reads the daily expenses breakdown from the stock template.

    Columns returned:
      date, meals, generator, electricity, water, salaries,
      stationary, security, sanitation, airtime, transport,
      nssf, sundries, maintenance, vat_tax, total_expenses

    Source sheet: 'Daily Expenses'
    """
    path = _validate_path(filepath)

    try:
        df_raw = pd.read_excel(
            path,
            sheet_name="Daily Expenses",
            header=0,
            engine="openpyxl",
        )
    except Exception as e:
        raise RuntimeError(f"Could not read Daily Expenses sheet: {e}")

    # Rename columns by position — header row defines order:
    # DATE, MEALS, GENERATOR, ELECTRICITY, WATER, SALARIES, STATIONARY,
    # SECURITY, SANITATION, AIRTIME/DATA, TRANSPORT, NSSF, SANDRIES,
    # MANTAINANCE, VAT/TAX, PHOTOCOPY, TAX COMPLIANCE FEES, ..., TOTAL

    col_rename = {
        0: "date",
        1: "meals",
        2: "generator",
        3: "electricity",
        4: "water",
        5: "salaries",
        6: "stationary",
        7: "security",
        8: "sanitation",
        9: "airtime",
        10: "transport",
        11: "nssf",
        12: "sundries",
        13: "maintenance",
        14: "vat_tax",
    }

    # Find the TOTAL column — it's the last non-empty named column
    # Based on our analysis it's at index 31
    total_col_idx = 31

    records = []
    for _, row in df_raw.iterrows():
        date_val = row.iloc[0] if len(row) > 0 else None
        if not _is_valid_date(date_val):
            continue

        record = {"date": pd.Timestamp(date_val)}
        for idx, col_name in col_rename.items():
            if col_name == "date":
                continue
            val = row.iloc[idx] if idx < len(row) else None
            record[col_name] = float(val) if pd.notna(val) else 0.0

        # Total expenses
        total_val = row.iloc[total_col_idx] if total_col_idx < len(row) else None
        record["total_expenses"] = float(total_val) if pd.notna(total_val) else 0.0

        records.append(record)

    if not records:
        print("  [WARNING] No daily expense data found in stock template.")
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)

    numeric_cols = [c for c in df.columns if c != "date"]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)

    return df


# ─────────────────────────────────────────────────────────────
# MONTHLY AGGREGATION HELPER
# ─────────────────────────────────────────────────────────────

def get_monthly_stock_metrics(
    filepath: str | Path,
    month: int,
    year: int,
) -> dict:
    """
    Aggregates all stock template data for a given month/year into
    a single dictionary ready for use in the report processor.

    Returns a dict with keys:
      fuel_stock   : { pms: {...}, ago: {...} }
      product_sales: { lubricants, tba, lpg, lpg_accessories,
                       car_wash, shop, solar, total }
      expenses     : { meals, generator, electricity, water,
                       salaries, stationary, security, sanitation,
                       airtime, transport, nssf, sundries,
                       maintenance, vat_tax, total }
    """
    path = _validate_path(filepath)
    result = {}

    # ── Fuel stock movement ──────────────────────────────────
    try:
        df_fuel = read_fuel_stock_movement(path)
        if not df_fuel.empty:
            mask = (df_fuel["date"].dt.month == month) & (df_fuel["date"].dt.year == year)
            df_m = df_fuel[mask]

            fuel_stock = {}
            for product in ["PMS", "AGO"]:
                p = df_m[df_m["product"] == product]
                if p.empty:
                    fuel_stock[product.lower()] = {}
                    continue

                opening = p.iloc[0]["opening_dip_ltrs"] if not p.empty else 0
                closing = p.iloc[-1]["closing_dip_ltrs"] if not p.empty else 0
                total_purchases = p["purchases_ltrs"].sum()
                total_sales = p["net_sales_ltrs"].sum()
                total_loss_gain_ltrs = p["loss_gain_ltrs"].sum()
                total_loss_gain_value = p["loss_gain_value_ugx"].sum()
                total_turnover = p["turnover_ugx"].sum()
                closing_stock_value = p.iloc[-1]["stock_value_ugx"] if not p.empty else 0
                pct_loss_gain = (
                    (total_loss_gain_ltrs / total_sales * 100) if total_sales != 0 else 0
                )

                fuel_stock[product.lower()] = {
                    "opening_dip_ltrs": round(opening, 2),
                    "closing_dip_ltrs": round(closing, 2),
                    "total_purchases_ltrs": round(total_purchases, 2),
                    "total_sales_ltrs": round(total_sales, 2),
                    "loss_gain_ltrs": round(total_loss_gain_ltrs, 2),
                    "loss_gain_pct": round(pct_loss_gain, 4),
                    "loss_gain_value_ugx": round(total_loss_gain_value, 2),
                    "turnover_ugx": round(total_turnover, 2),
                    "closing_stock_value_ugx": round(closing_stock_value, 2),
                }

            result["fuel_stock"] = fuel_stock
        else:
            result["fuel_stock"] = {}

    except Exception as e:
        print(f"  [WARNING] Fuel stock aggregation error: {e}")
        result["fuel_stock"] = {}

    # ── General product sales ────────────────────────────────
    try:
        df_sales = read_general_sales_summary(path)
        if not df_sales.empty:
            mask = (df_sales["date"].dt.month == month) & (df_sales["date"].dt.year == year)
            df_m = df_sales[mask]

            result["product_sales"] = {
                "lubricants_ugx": round(df_m["lubricants_ugx"].sum(), 2),
                "tba_ugx": round(df_m["tba_ugx"].sum(), 2),
                "lpg_ugx": round(df_m["lpg_ugx"].sum(), 2),
                "lpg_accessories_ugx": round(df_m["lpg_accessories_ugx"].sum(), 2),
                "car_wash_ugx": round(df_m["car_wash_ugx"].sum(), 2),
                "shop_ugx": round(df_m["shop_ugx"].sum(), 2),
                "solar_ugx": round(df_m["solar_ugx"].sum(), 2),
                "total_ugx": round(df_m["total_ugx"].sum(), 2),
            }
        else:
            result["product_sales"] = {}

    except Exception as e:
        print(f"  [WARNING] General sales aggregation error: {e}")
        result["product_sales"] = {}

    # ── Daily expenses ───────────────────────────────────────
    try:
        df_exp = read_daily_expenses_stock(path)
        if not df_exp.empty:
            mask = (df_exp["date"].dt.month == month) & (df_exp["date"].dt.year == year)
            df_m = df_exp[mask]

            exp_cols = [
                "meals", "generator", "electricity", "water", "salaries",
                "stationary", "security", "sanitation", "airtime", "transport",
                "nssf", "sundries", "maintenance", "vat_tax",
            ]
            expenses = {col: round(df_m[col].sum(), 2) for col in exp_cols if col in df_m.columns}
            expenses["total_expenses"] = round(df_m["total_expenses"].sum(), 2)
            result["expenses"] = expenses
        else:
            result["expenses"] = {}

    except Exception as e:
        print(f"  [WARNING] Expenses aggregation error: {e}")
        result["expenses"] = {}

    return result


# ─────────────────────────────────────────────────────────────
# QUICK TEST  (run this file directly to verify)
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    filepath = r"data\input\Stock_mvt-Template-2025-2026-V1-7-25.xlsx"

    print("=" * 60)
    print("TEST: read_fuel_stock_movement()")
    print("=" * 60)
    df = read_fuel_stock_movement(filepath)
    print(f"  Rows loaded : {len(df)}")
    print(f"  Date range  : {df['date'].min().date()} → {df['date'].max().date()}")
    print(f"  Products    : {df['product'].unique().tolist()}")
    print(df.head(3).to_string(index=False))

    print()
    print("=" * 60)
    print("TEST: read_general_sales_summary()")
    print("=" * 60)
    df2 = read_general_sales_summary(filepath)
    print(f"  Rows loaded : {len(df2)}")
    print(f"  Date range  : {df2['date'].min().date()} → {df2['date'].max().date()}")
    print(df2.head(3).to_string(index=False))

    print()
    print("=" * 60)
    print("TEST: read_daily_expenses_stock()")
    print("=" * 60)
    df3 = read_daily_expenses_stock(filepath)
    print(f"  Rows loaded : {len(df3)}")
    print(df3.head(3).to_string(index=False))

    print()
    print("=" * 60)
    print("TEST: get_monthly_stock_metrics() — April 2026")
    print("=" * 60)
    metrics = get_monthly_stock_metrics(filepath, month=4, year=2026)
    import json
    print(json.dumps(metrics, indent=2, default=str))
