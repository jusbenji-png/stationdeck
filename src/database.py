"""
database.py — StationDeck Local SQLite Database Layer
Phase 9A

Handles all permanent data storage for StationDeck.
Database file: data/stationdeck.db
"""

import sqlite3
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import date, timedelta
import logging

logger = logging.getLogger(__name__)

# Writable DB location — under %LOCALAPPDATA%\StationDeck when frozen so it is
# writable by non-admin users and survives app updates.
try:
    from config.settings import DATA_DIR
    DB_PATH = DATA_DIR / "data" / "stationdeck.db"
except Exception:
    DB_PATH = Path(__file__).parent.parent / "data" / "stationdeck.db"

DAILY_COLUMNS = [
    "date", "pms_volume", "ago_volume", "pms_price", "ago_price",
    "pms_revenue", "ago_revenue", "lubes_litres", "lubes_revenue",
    "lpg_kgs", "lpg_revenue", "tba_credits", "plus_card_payment",
    "shop_sales", "plus_card_payment_total", "tyre_sales", "cashless_total",
    "plus_card_pms", "plus_card_ago", "other_payments", "momo_pay",
    "airtel_pay", "visa", "credit_sales", "total_cash", "expense_umeme",
    "expense_water", "expense_security", "expense_stationery",
    "expense_generator", "expense_meals", "expense_transport",
    "expense_salaries", "expense_sanitary", "expense_airtime",
    "expense_misc", "expense_shop_packaging", "total_expenses",
    "stock_tba", "stock_lpg_acc", "stock_shop_purchase", "cash_to_bank",
    "actual_cash_banked", "delta", "source_sheet",
]


def _sanitize_value(val):
    """
    Convert any pandas/numpy type to a plain Python type that SQLite accepts.
    - pandas Timestamp  → "YYYY-MM-DD" string
    - numpy float NaN   → None
    - numpy int/float   → plain Python int/float
    - everything else   → unchanged
    """
    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, float) and np.isnan(val):
        return None
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        return float(val)
    return val


def _get_connection() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> bool:
    """Create the database and daily_records table if they don't exist."""
    try:
        conn = _get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS daily_records (
                id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                station_id              TEXT    NOT NULL,
                date                    TEXT    NOT NULL,
                pms_volume              REAL,
                ago_volume              REAL,
                pms_price               REAL,
                ago_price               REAL,
                pms_revenue             REAL,
                ago_revenue             REAL,
                lubes_litres            REAL,
                lubes_revenue           REAL,
                lpg_kgs                 REAL,
                lpg_revenue             REAL,
                tba_credits             REAL,
                plus_card_payment       REAL,
                shop_sales              REAL,
                plus_card_payment_total REAL,
                tyre_sales              REAL,
                cashless_total          REAL,
                plus_card_pms           REAL,
                plus_card_ago           REAL,
                other_payments          REAL,
                momo_pay                REAL,
                airtel_pay              REAL,
                visa                    REAL,
                credit_sales            REAL,
                total_cash              REAL,
                expense_umeme           REAL,
                expense_water           REAL,
                expense_security        REAL,
                expense_stationery      REAL,
                expense_generator       REAL,
                expense_meals           REAL,
                expense_transport       REAL,
                expense_salaries        REAL,
                expense_sanitary        REAL,
                expense_airtime         REAL,
                expense_misc            REAL,
                expense_shop_packaging  REAL,
                total_expenses          REAL,
                stock_tba               REAL,
                stock_lpg_acc           REAL,
                stock_shop_purchase     REAL,
                cash_to_bank            REAL,
                actual_cash_banked      REAL,
                delta                   REAL,
                source_sheet            TEXT,
                created_at              TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(station_id, date)
            )
        """)
        conn.commit()
        conn.close()
        logger.info(f"Database initialised at: {DB_PATH}")
        return True
    except Exception as e:
        logger.error(f"init_db() failed: {e}")
        return False


def save_daily_record(record: dict, station_id: str = "te_rwizi") -> bool:
    """
    Save one day's record to the database.
    Existing records for the same station+date are replaced (safe to re-import).
    """
    try:
        conn = _get_connection()
        cursor = conn.cursor()

        columns = ["station_id", "date"] + DAILY_COLUMNS
        placeholders = ", ".join(["?"] * len(columns))
        col_names = ", ".join(columns)

        # Convert date to string
        date_val = record.get("date")
        if isinstance(date_val, pd.Timestamp):
            date_val = date_val.strftime("%Y-%m-%d")
        elif hasattr(date_val, "strftime"):
            date_val = date_val.strftime("%Y-%m-%d")

        # Build values — sanitize every value, skip "date" (handled above)
        values = [station_id, date_val]
        for col in DAILY_COLUMNS:
            if col == "date":
                values.append(date_val)
            else:
                values.append(_sanitize_value(record.get(col)))

        cursor.execute(
            f"INSERT OR REPLACE INTO daily_records ({col_names}) VALUES ({placeholders})",
            values
        )
        conn.commit()
        conn.close()
        return True

    except Exception as e:
        logger.error(f"save_daily_record() failed for date {record.get('date')}: {e}")
        return False


def import_from_excel(filepath: str | Path, station_id: str = "te_rwizi") -> int:
    """
    Bulk import all records from Daily_cash_flow.xlsx into the database.
    Uses pandas to_sql for fast bulk insertion instead of row-by-row.
    Returns number of records imported (0 on failure).
    """
    try:
        from src.reader import read_daily_cashflow

        logger.info(f"Importing Excel data from: {filepath}")
        df = read_daily_cashflow(str(filepath))
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["date"])

        # Convert date column to plain string — SQLite stores dates as TEXT
        df["date"] = df["date"].dt.strftime("%Y-%m-%d")

        # Add station_id column
        df.insert(0, "station_id", station_id)

        # Keep only columns we care about (station_id + DAILY_COLUMNS)
        keep_cols = ["station_id"] + DAILY_COLUMNS
        # Only keep columns that exist in df
        keep_cols = [c for c in keep_cols if c in df.columns]
        df = df[keep_cols]

        # Replace NaN with None (SQLite NULL)
        df = df.where(pd.notnull(df), None)

        # Use pandas to_sql with INSERT OR REPLACE via raw sqlite3
        conn = _get_connection()

        # Build INSERT OR REPLACE statement
        col_names = ", ".join(keep_cols)
        placeholders = ", ".join(["?"] * len(keep_cols))
        sql = f"INSERT OR REPLACE INTO daily_records ({col_names}) VALUES ({placeholders})"

        imported = 0
        failed = 0
        for row in df.itertuples(index=False, name=None):
            try:
                conn.execute(sql, row)
                imported += 1
            except Exception as row_err:
                logger.error(f"Row insert failed: {row_err}")
                failed += 1

        conn.commit()
        conn.close()

        logger.info(
            f"Import complete — {imported} records saved, {failed} failed | "
            f"Station: {station_id}"
        )
        return imported

    except Exception as e:
        logger.error(f"import_from_excel() failed: {e}")
        return 0


def _query_to_df(query: str, params: tuple) -> pd.DataFrame:
    """Run a SQL query and return results as a DataFrame."""
    try:
        conn = _get_connection()
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        if not df.empty and "date" in df.columns:
            df["date"] = pd.to_datetime(df["date"], errors="coerce")
        return df
    except Exception as e:
        logger.error(f"Database query failed: {e}")
        return pd.DataFrame()


def get_records_by_date_range(
    start_date, end_date, station_id: str = "te_rwizi"
) -> pd.DataFrame:
    if hasattr(start_date, "strftime"):
        start_date = start_date.strftime("%Y-%m-%d")
    if hasattr(end_date, "strftime"):
        end_date = end_date.strftime("%Y-%m-%d")
    query = """
        SELECT * FROM daily_records
        WHERE station_id = ? AND date >= ? AND date <= ?
        ORDER BY date ASC
    """
    return _query_to_df(query, (station_id, start_date, end_date))


def get_records_by_month(
    month: int, year: int, station_id: str = "te_rwizi"
) -> pd.DataFrame:
    start = date(year, month, 1)
    end = date(year + 1, 1, 1) - timedelta(days=1) if month == 12 \
        else date(year, month + 1, 1) - timedelta(days=1)
    return get_records_by_date_range(start, end, station_id)


def get_records_by_week(
    start_date, station_id: str = "te_rwizi"
) -> pd.DataFrame:
    if isinstance(start_date, str):
        start_date = date.fromisoformat(start_date)
    return get_records_by_date_range(start_date, start_date + timedelta(days=6), station_id)


def get_records_by_financial_year(
    fy_year: int, station_id: str = "te_rwizi"
) -> pd.DataFrame:
    return get_records_by_date_range(
        date(fy_year, 7, 1), date(fy_year + 1, 6, 30), station_id
    )


def get_record_count(station_id: str = "te_rwizi") -> int:
    try:
        conn = _get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT COUNT(*) FROM daily_records WHERE station_id = ?", (station_id,)
        )
        count = cursor.fetchone()[0]
        conn.close()
        return count
    except Exception as e:
        logger.error(f"get_record_count() failed: {e}")
        return 0


def delete_all_records(station_id: str = "te_rwizi") -> int:
    """Delete ALL daily records for a station. Returns number of rows removed."""
    try:
        conn = _get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM daily_records WHERE station_id = ?", (station_id,))
        removed = cursor.rowcount
        conn.commit()
        conn.close()
        logger.info(f"delete_all_records(): removed {removed} rows for {station_id}")
        return removed
    except Exception as e:
        logger.error(f"delete_all_records() failed: {e}")
        return 0


def delete_records_by_date(date_str: str, station_id: str = "te_rwizi") -> int:
    """Delete records for a single date (YYYY-MM-DD). Returns rows removed."""
    try:
        conn = _get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM daily_records WHERE station_id = ? AND date = ?",
            (station_id, date_str),
        )
        removed = cursor.rowcount
        conn.commit()
        conn.close()
        logger.info(f"delete_records_by_date({date_str}): removed {removed} rows")
        return removed
    except Exception as e:
        logger.error(f"delete_records_by_date() failed: {e}")
        return 0


def get_date_range_stored(station_id: str = "te_rwizi") -> dict:
    try:
        conn = _get_connection()
        cursor = conn.cursor()
        cursor.execute(
            "SELECT MIN(date) as earliest, MAX(date) as latest "
            "FROM daily_records WHERE station_id = ?",
            (station_id,)
        )
        row = cursor.fetchone()
        conn.close()
        return {"earliest": row["earliest"], "latest": row["latest"]}
    except Exception as e:
        logger.error(f"get_date_range_stored() failed: {e}")
        return {"earliest": None, "latest": None}