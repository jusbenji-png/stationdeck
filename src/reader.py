# StationDeck - Excel Reader
import pandas as pd
from pathlib import Path
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

CASHFLOW_COLUMNS = {
    0:  "date",
    1:  "pms_volume",
    2:  "ago_volume",
    3:  "pms_price",
    4:  "ago_price",
    5:  "pms_revenue",
    6:  "ago_revenue",
    7:  "lubes_litres",
    8:  "lubes_revenue",
    9:  "lpg_kgs",
    10: "lpg_revenue",
    11: "tba_credits",
    12: "plus_card_payment",
    13: "shop_sales",
    14: "plus_card_payment_total",
    15: "tyre_sales",
    16: "cashless_total",
    17: "plus_card_pms",
    18: "plus_card_ago",
    19: "other_payments",
    20: "momo_pay",
    21: "airtel_pay",
    22: "visa",
    23: "credit_sales",
    24: "total_cash",
    25: "expense_umeme",
    26: "expense_water",
    27: "expense_security",
    28: "expense_stationery",
    29: "expense_generator",
    30: "expense_meals",
    31: "expense_transport",
    32: "expense_salaries",
    33: "expense_sanitary",
    34: "expense_airtime",
    35: "expense_misc",
    36: "expense_shop_packaging",
    37: "total_expenses",
    38: "stock_tba",
    39: "stock_lpg_acc",
    40: "stock_shop_purchase",
    42: "cash_to_bank",
    43: "actual_cash_banked",
    44: "delta",
}


def convert_excel_date(value):
    if pd.isna(value):
        return None
    try:
        if isinstance(value, (int, float)):
            return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(value))
        if isinstance(value, (datetime, pd.Timestamp)):
            return pd.Timestamp(value)
        return pd.Timestamp(str(value))
    except Exception:
        return None


def is_valid_date_row(value):
    if pd.isna(value):
        return False
    if isinstance(value, (int, float)):
        return 43831 <= int(value) <= 47483
    if isinstance(value, (datetime, pd.Timestamp)):
        return True
    try:
        pd.Timestamp(str(value))
        return True
    except Exception:
        return False


def clean_numeric(value):
    if pd.isna(value):
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def read_cashflow_sheet(sheet_df):
    records = []

    for _, row in sheet_df.iterrows():
        raw_date = row.iloc[0] if len(row) > 0 else None
        if not is_valid_date_row(raw_date):
            continue
        date = convert_excel_date(raw_date)
        if date is None:
            continue

        record = {"date": date.date()}
        for col_index, col_name in CASHFLOW_COLUMNS.items():
            if col_name == "date":
                continue
            if col_index < len(row):
                record[col_name] = clean_numeric(row.iloc[col_index])
            else:
                record[col_name] = 0.0
        records.append(record)

    if not records:
        return None

    df = pd.DataFrame(records)

    financial_cols = ['pms_revenue', 'ago_revenue', 'total_cash']
    df = df[df[financial_cols].sum(axis=1) > 0]
    df = df[pd.to_datetime(df['date']) <= pd.Timestamp('2030-12-31')]

    return df


def read_daily_cashflow(file_path):
    file_path = Path(file_path)
    if not file_path.exists():
        logger.error(f"File not found: {file_path}")
        return None

    logger.info(f"Opening file: {file_path.name}")

    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        logger.info(f"Found {len(sheet_names)} sheets: {sheet_names}")

        all_months = []
        skip_sheets = ['Sheet1', 'Sheet2', 'Sheet3', 'Chart1']

        for sheet_name in sheet_names:
            if sheet_name in skip_sheets:
                logger.info(f"Skipping placeholder sheet: {sheet_name}")
                continue

            logger.info(f"Processing sheet: {sheet_name}")
            df_raw = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                header=None,
                engine='openpyxl'
            )
            df_clean = read_cashflow_sheet(df_raw)

            if df_clean is not None and len(df_clean) > 0:
                df_clean["source_sheet"] = sheet_name
                all_months.append(df_clean)
                logger.info(f"  Extracted {len(df_clean)} daily records")
            else:
                logger.warning(f"  No valid data found in sheet: {sheet_name}")

        if not all_months:
            logger.error("No valid data found in any sheet")
            return None

        df_combined = pd.concat(all_months, ignore_index=True)
        df_combined = df_combined.sort_values("date").reset_index(drop=True)

        logger.info(f"Total records loaded: {len(df_combined)}")
        logger.info(f"Date range: {df_combined['date'].min()} to {df_combined['date'].max()}")

        return df_combined

    except Exception as e:
        logger.error(f"Failed to read file: {e}")
        return None