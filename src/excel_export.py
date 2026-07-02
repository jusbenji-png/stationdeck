# =============================================================
# src/excel_export.py — StationDeck Filled-Template Excel Export
# =============================================================
#
# WHY THIS EXISTS:
#   Stations import their Daily Cash Flow / Stock Movement Excel files,
#   and also enter days directly in the app (OCR daily entry). Those
#   app-entered days exist only in the database — the station's own
#   Excel files fall behind. This module writes the database back INTO
#   the station's real template files, so the exported workbook contains
#   BOTH previously-imported and app-entered data, in the exact layout
#   TotalEnergies expects, and can be re-imported into StationDeck.
#
# DESIGN RULES (data safety first):
#   1. The master file in data/input/ is NEVER modified. We copy it to
#      data/exports/ and fill the copy.
#   2. Rows that already contain data in the master are left 100%
#      untouched — the master was the import source, so it wins.
#      Only rows that are EMPTY in the master but present in the DB
#      (i.e. app-entered days) are filled.
#   3. Formula cells are preserved wherever possible. On rows we fill,
#      computed cells are written as plain values (openpyxl cannot
#      recalculate), and fullCalcOnLoad is set so Excel refreshes the
#      monthly SUM/subtotal formulas the first time the file is opened.
#
# WHAT EACH EXPORT CAN FILL:
#   Cash Flow      — every column (the DB stores the full row).
#   Stock Movement — Daily Expenses (11 of 14 categories) and the
#                    manual columns of General Sales Sumary (Lubricants,
#                    TBA, LPG). Fuel dips/purchases and the Shop Sales
#                    day/night split are physical readings the app does
#                    not persist, so those stay as they were imported.
# =============================================================

import logging
import shutil
from datetime import datetime, date
from pathlib import Path

import openpyxl

from src.database import get_records_by_date_range

logger = logging.getLogger(__name__)


# ── Cash Flow: DB column → 1-based Excel column on 'CASH FLOW' sheet ─────────
# Mirrors CASHFLOW_COLUMNS in src/reader.py (0-based there, +1 here).
CASHFLOW_INPUT_COLS = {
    "pms_volume":              2,   # B
    "ago_volume":              3,   # C
    "pms_price":               4,   # D
    "ago_price":               5,   # E
    "lubes_litres":            8,   # H
    "lubes_revenue":           9,   # I
    "lpg_kgs":                 10,  # J
    "lpg_revenue":             11,  # K
    "tba_credits":             12,  # L
    "plus_card_payment":       13,  # M
    "shop_sales":              14,  # N
    "plus_card_payment_total": 15,  # O
    "tyre_sales":              16,  # P
    "plus_card_pms":           18,  # R
    "plus_card_ago":           19,  # S
    "other_payments":          20,  # T
    "momo_pay":                21,  # U
    "airtel_pay":              22,  # V
    "visa":                    23,  # W
    "credit_sales":            24,  # X
    "expense_umeme":           26,  # Z
    "expense_water":           27,  # AA
    "expense_security":        28,  # AB
    "expense_stationery":      29,  # AC
    "expense_generator":       30,  # AD
    "expense_meals":           31,  # AE
    "expense_transport":       32,  # AF
    "expense_salaries":        33,  # AG
    "expense_sanitary":        34,  # AH
    "expense_airtime":         35,  # AI
    "expense_misc":            36,  # AJ
    "expense_shop_packaging":  37,  # AK
    "stock_tba":               39,  # AM
    "stock_lpg_acc":           40,  # AN
    "stock_shop_purchase":     41,  # AO
    "actual_cash_banked":      44,  # AR
}

# Computed columns — normally Excel formulas. Written as plain values on
# rows WE fill (the DB already holds the computed results), so the file is
# correct immediately and on re-import into StationDeck.
CASHFLOW_COMPUTED_COLS = {
    "pms_revenue":    6,   # F  (=B*D)
    "ago_revenue":    7,   # G  (=C*E)
    "cashless_total": 17,  # Q
    "total_cash":     25,  # Y
    "total_expenses": 38,  # AL
    "cash_to_bank":   43,  # AQ
    "delta":          45,  # AS
}

# Cells checked to decide whether a master row already has data.
_CASHFLOW_PRESENCE_COLS = (2, 3, 25)   # B pms_vol, C ago_vol, Y total_cash


# ── Stock Movement: 'Daily Expenses' sheet — DB col → 1-based Excel col ─────
# Mirrors EXPENSE_COLS in src/reader_stock.py. NSSF / Maintenance / VAT are
# captured on paper but not stored per-category in the DB → left blank.
STOCK_EXPENSE_COLS = {
    "expense_meals":       2,   # B  Meals
    "expense_generator":   3,   # C  Generator
    "expense_umeme":       4,   # D  Electricity
    "expense_water":       5,   # E  Water
    "expense_salaries":    6,   # F  Salaries
    "expense_stationery":  7,   # G  Stationary
    "expense_security":    8,   # H  Security
    "expense_sanitary":    9,   # I  Sanitation
    "expense_airtime":     10,  # J  Airtime/Data
    "expense_transport":   11,  # K  Transport
    "expense_misc":        13,  # M  Sundries
}

# 'General Sales Sumary' — only its truly manual columns. PMS/AGO (B/C) and
# Shop (J) are cross-sheet formulas and must not be overwritten.
STOCK_GENERAL_SALES_COLS = {
    "lubes_revenue": 4,   # D  Lubricants
    "tba_credits":   5,   # E  TBA
    "lpg_revenue":   7,   # G  LPG
}


# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────

def _all_db_records(station_id: str):
    """Every daily record for the station, as {date_str: row_dict}."""
    df = get_records_by_date_range("2000-01-01", "2100-01-01", station_id)
    records = {}
    if df is None or df.empty:
        return records
    for _, row in df.iterrows():
        d = str(row["date"])[:10]
        records[d] = row.to_dict()
    return records


def _date_key(value):
    """Normalize an Excel cell value to 'YYYY-MM-DD' or None."""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    return None


def _build_date_index(ws_vals):
    """Map date -> row number using the CACHED values of column A.

    Templates chain dates with formulas (=+A2+1), so the formula view of
    the sheet can't be used — this must come from a data_only load.
    """
    index = {}
    for row in range(2, ws_vals.max_row + 1):
        key = _date_key(ws_vals.cell(row=row, column=1).value)
        if key and key not in index:
            index[key] = row
    return index


def _row_is_empty(ws_vals, row, presence_cols):
    """True when none of the presence cells hold a non-zero cached value."""
    for col in presence_cols:
        v = ws_vals.cell(row=row, column=col).value
        if isinstance(v, (int, float)) and v != 0:
            return False
    return True


def _num(record, key):
    v = record.get(key)
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f == int(f) else round(f, 2)


def _bake_row(ws, ws_vals, row, max_col):
    """Replace this row's formula cells with their Excel-cached values.

    openpyxl saves formulas WITHOUT their cached results, so a re-read of
    the saved file (pandas/StationDeck import) would see those cells as
    empty. Baking the cached values into data rows keeps the export
    readable everywhere. Rows without data keep their formulas so managers
    can continue filling the sheet by hand in Excel.
    """
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cached = ws_vals.cell(row=row, column=col).value
            if cached is not None:
                cell.value = cached


def _row_has_any_value(ws_vals, row, max_col):
    """True if any cached cell beyond the date column holds a non-zero number."""
    for col in range(2, max_col + 1):
        v = ws_vals.cell(row=row, column=col).value
        if isinstance(v, (int, float)) and v != 0:
            return True
    return False


def _bake_data_rows(ws, ws_vals, max_col, extra_rows=frozenset()):
    """Bake cached formula results into every dated row that holds data.

    Walks ALL rows (templates can carry duplicate dates — stray partial
    rows next to the real one — and a first-match date index would miss
    the real row). Rows without data keep their formulas so the sheet
    stays hand-fillable; `extra_rows` forces rows we filled from the DB.
    """
    for row in range(2, ws_vals.max_row + 1):
        if _date_key(ws_vals.cell(row=row, column=1).value) is None:
            continue
        if row in extra_rows or _row_has_any_value(ws_vals, row, max_col):
            _bake_row(ws, ws_vals, row, max_col)


def _export_path(exports_dir: Path, label: str) -> Path:
    exports_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d")
    return exports_dir / f"StationDeck_Export_{label}_{stamp}.xlsx"


# ─────────────────────────────────────────────────────────────
# EXPORT 1 — DAILY CASH FLOW
# ─────────────────────────────────────────────────────────────

def export_cashflow(station_config: dict, station_id: str,
                    exports_dir: Path) -> dict:
    """Fill the station's cash-flow template with all DB data.

    Returns {"success", "path", "rows_filled", "message"}.
    """
    src = Path(station_config["files"]["cashflow"])
    if not src.exists():
        return {"success": False, "path": None, "rows_filled": 0,
                "message": "No cash flow file has been imported yet — "
                           "import one first so StationDeck has the template."}

    out = _export_path(exports_dir, "Daily_Cash_Flow")
    shutil.copy2(src, out)

    records = _all_db_records(station_id)
    if not records:
        return {"success": False, "path": None, "rows_filled": 0,
                "message": "The database has no records to export."}

    wb_vals = openpyxl.load_workbook(out, data_only=True)
    wb      = openpyxl.load_workbook(out)

    skip = {"Sheet1", "Sheet2", "Sheet3", "Chart1"}
    rows_filled = 0

    for sheet_name in wb.sheetnames:
        if sheet_name in skip or sheet_name not in wb_vals.sheetnames:
            continue
        ws_vals = wb_vals[sheet_name]
        ws      = wb[sheet_name]
        index   = _build_date_index(ws_vals)
        if not index:
            continue

        max_col = max(CASHFLOW_COMPUTED_COLS.values())
        filled  = set()

        for date_str, row_no in index.items():
            record = records.get(date_str)
            if record is None:
                continue
            # Master rows that already hold data are the import source —
            # never overwrite them.
            if not _row_is_empty(ws_vals, row_no, _CASHFLOW_PRESENCE_COLS):
                continue

            wrote = False
            for db_col, xl_col in CASHFLOW_INPUT_COLS.items():
                v = _num(record, db_col)
                if v is not None:
                    ws.cell(row=row_no, column=xl_col).value = v
                    wrote = True
            if wrote:
                # Replace this row's formulas with the DB's computed values
                # so the row is correct immediately (openpyxl can't recalc).
                for db_col, xl_col in CASHFLOW_COMPUTED_COLS.items():
                    v = _num(record, db_col)
                    if v is not None:
                        ws.cell(row=row_no, column=xl_col).value = v
                filled.add(row_no)
                rows_filled += 1

        # Bake Excel's cached formula results into every data row —
        # openpyxl drops them on save, which would otherwise make the
        # whole file unreadable on re-import. Walks all rows, so stray
        # duplicate-date rows are baked too.
        _bake_data_rows(ws, ws_vals, max_col, extra_rows=filled)

    wb.calculation.fullCalcOnLoad = True   # refresh SUM totals in Excel
    wb.save(out)
    wb.close()
    wb_vals.close()

    logger.info(f"export_cashflow: {rows_filled} app-entered day(s) "
                f"written into {out.name}")
    return {"success": True, "path": out, "rows_filled": rows_filled,
            "message": f"Cash flow exported — {rows_filled} app-entered "
                       f"day(s) added to the imported data."}


# ─────────────────────────────────────────────────────────────
# EXPORT 2 — STOCK MOVEMENT
# ─────────────────────────────────────────────────────────────

def export_stock(station_config: dict, station_id: str,
                 exports_dir: Path) -> dict:
    """Fill the stock-movement template's Daily Expenses and General Sales
    manual columns from the DB. Fuel dips/purchases stay as imported —
    the app never persists physical tank readings."""
    src = Path(station_config["files"]["stock"])
    if not src.exists():
        return {"success": False, "path": None, "rows_filled": 0,
                "message": "No stock movement file has been imported yet — "
                           "import one first so StationDeck has the template."}

    out = _export_path(exports_dir, "Stock_Movement")
    shutil.copy2(src, out)

    records = _all_db_records(station_id)
    if not records:
        return {"success": False, "path": None, "rows_filled": 0,
                "message": "The database has no records to export."}

    wb_vals = openpyxl.load_workbook(out, data_only=True)
    wb      = openpyxl.load_workbook(out)

    rows_filled = 0
    filled_rows = set()   # (sheet_name, row_no) written from the DB

    # ── Daily Expenses ────────────────────────────────────────
    if "Daily Expenses" in wb.sheetnames:
        ws_vals = wb_vals["Daily Expenses"]
        ws      = wb["Daily Expenses"]
        for date_str, row_no in _build_date_index(ws_vals).items():
            record = records.get(date_str)
            if record is None:
                continue
            # Fill ONLY rows that are empty across the ENTIRE row — the
            # sheet has categories beyond what the DB stores (NSSF, VAT,
            # maintenance, cols 15-31), and any of them counts as data.
            if _row_has_any_value(ws_vals, row_no, 32):
                continue
            wrote = False
            for db_col, xl_col in STOCK_EXPENSE_COLS.items():
                v = _num(record, db_col)
                if v:   # expenses: only write non-zero, keep sheet sparse
                    ws.cell(row=row_no, column=xl_col).value = v
                    wrote = True
            if wrote:
                # The TOTAL column (AF) is a per-row formula openpyxl can't
                # recalculate — write the DB total so re-import reads it.
                v = _num(record, "total_expenses")
                if v:
                    ws.cell(row=row_no, column=32).value = v
                filled_rows.add(("Daily Expenses", row_no))
                rows_filled += 1

    # ── General Sales Sumary (manual columns only) ────────────
    if "General Sales Sumary" in wb.sheetnames:
        ws_vals = wb_vals["General Sales Sumary"]
        ws      = wb["General Sales Sumary"]
        presence = tuple(STOCK_GENERAL_SALES_COLS.values())
        for date_str, row_no in _build_date_index(ws_vals).items():
            record = records.get(date_str)
            if record is None:
                continue
            if not _row_is_empty(ws_vals, row_no, presence):
                continue
            wrote = False
            for db_col, xl_col in STOCK_GENERAL_SALES_COLS.items():
                v = _num(record, db_col)
                if v:
                    ws.cell(row=row_no, column=xl_col).value = v
                    wrote = True
            if wrote:
                filled_rows.add(("General Sales Sumary", row_no))

    # ── Bake cached formula results into all data rows ────────
    # StationDeck's readers consume formula columns on these sheets
    # (fuel turnover, cross-sheet revenue, per-row totals). openpyxl
    # drops Excel's cached results on save, so without this the export
    # would re-import as empty. Rows without data keep their formulas.
    # Rows we just filled are also baked so their formula-chained date
    # cells become real dates the readers can recognize.
    _READER_SHEETS = [s for s in wb.sheetnames
                      if s.startswith("Fuel Mvt") or s in
                      ("Fuel Sales Sumary", "General Sales Sumary",
                       "Shop Sales", "Daily Expenses")]
    for sheet_name in _READER_SHEETS:
        ws_vals = wb_vals[sheet_name]
        ws      = wb[sheet_name]
        max_col = min(ws.max_column, 40)
        extra   = {r for (s, r) in filled_rows if s == sheet_name}
        _bake_data_rows(ws, ws_vals, max_col, extra_rows=extra)

    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    wb.close()
    wb_vals.close()

    logger.info(f"export_stock: {rows_filled} expense day(s) "
                f"written into {out.name}")
    return {"success": True, "path": out, "rows_filled": rows_filled,
            "message": f"Stock movement exported — {rows_filled} app-entered "
                       f"expense day(s) added. Fuel dips and the shop "
                       f"day/night split stay as imported (physical readings "
                       f"the app doesn't capture)."}
