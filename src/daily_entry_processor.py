# =============================================================
# src/daily_entry_processor.py
# StationDeck — Daily Entry Processor
#
# Receives the confirmed field dict from the daily_entry/commit
# route, calculates derived values, and writes one row to the
# SQLite database using the exact column names from database.py.
#
# DAILY_COLUMNS reference (from database.py):
#   date, pms_volume, ago_volume, pms_price, ago_price,
#   pms_revenue, ago_revenue, lubes_litres, lubes_revenue,
#   lpg_kgs, lpg_revenue, tba_credits, plus_card_payment,
#   shop_sales, plus_card_payment_total, tyre_sales, cashless_total,
#   plus_card_pms, plus_card_ago, other_payments, momo_pay,
#   airtel_pay, visa, credit_sales, total_cash, expense_umeme,
#   expense_water, expense_security, expense_stationery,
#   expense_generator, expense_meals, expense_transport,
#   expense_salaries, expense_sanitary, expense_airtime,
#   expense_misc, expense_shop_packaging, total_expenses,
#   stock_tba, stock_lpg_acc, stock_shop_purchase, cash_to_bank,
#   actual_cash_banked, delta, source_sheet
# =============================================================

import logging
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)


# =============================================================
# MARGIN FALLBACKS
# =============================================================
# Used ONLY when selling_price was not captured on the Meter Book.
# When the manager fills in the Selling Price box, that value is
# used directly and these constants are not applied.
PMS_MARGIN = 140.0   # UGX per litre above cost price
AGO_MARGIN = 125.0   # UGX per litre above cost price


# =============================================================
# PUBLIC ENTRY POINT
# =============================================================

def process_daily_entry(
    entry_date: str,
    shift: str,
    fields: dict,
    station_id: str,
    mode: str = "ocr",
) -> dict:
    """
    Validate and write a confirmed daily entry to SQLite.

    Args:
        entry_date  : "YYYY-MM-DD"
        shift       : "Day" or "Night"
        fields      : dict of field_id -> value string from the form
        station_id  : e.g. "te_rwizi"
        mode        : "ocr" or "excel"

    Returns:
        {"success": True/False, "message": str, "record": dict or None}
    """

    # ── 1. Parse and validate the date ───────────────────────
    try:
        datetime.strptime(entry_date, "%Y-%m-%d")
    except ValueError:
        return {
            "success": False,
            "message": f"Invalid date format: '{entry_date}'. Expected YYYY-MM-DD.",
            "record":  None,
        }

    # ── 2. Helper: parse a field to float safely ──────────────
    def _n(key, default=0.0):
        raw = fields.get(key, "")
        if raw is None or str(raw).strip() == "":
            return default
        try:
            cleaned = str(raw).replace(",", "").replace(" ", "")
            return float(cleaned)
        except (ValueError, TypeError):
            return default

    # ── 3. Meter Book — PMS volumes and selling price ─────────
    # Net volume per tank = opening_dip - closing_dip
    #                       + return_tank - updf_consumption
    #
    # Selling price priority:
    #   1. Use the actual selling_price field if captured and > 0
    #   2. Fall back to cost_price + PMS_MARGIN
    #
    # We compute a volume-weighted average selling price across
    # all tanks so the revenue figure is accurate even if tanks
    # have different prices on the same shift.

    pms_volume       = 0.0
    pms_sell_total   = 0.0   # sum of (volume × selling_price) per tank

    for t in range(1, 5):
        opening      = _n(f"pms{t}_opening_dip")
        closing      = _n(f"pms{t}_closing_dip")
        ret          = _n(f"pms{t}_return_tank")
        updf_c       = _n(f"pms{t}_updf_consumption")
        cost_p       = _n(f"pms{t}_cost_price")
        selling_p    = _n(f"pms{t}_selling_price")

        # Use actual selling price if captured, else estimate
        if selling_p > 0:
            effective_price = selling_p
        elif cost_p > 0:
            effective_price = cost_p + PMS_MARGIN
        else:
            effective_price = 0.0

        net = opening - closing + ret - updf_c
        if net > 0:
            pms_volume     += net
            pms_sell_total += net * effective_price

    pms_avg_sell = (pms_sell_total / pms_volume) if pms_volume > 0 else 0.0
    pms_revenue  = round(pms_volume * pms_avg_sell)

    # ── 4. Meter Book — AGO volumes and selling price ─────────
    ago_volume       = 0.0
    ago_sell_total   = 0.0

    for t in range(1, 5):
        opening      = _n(f"ago{t}_opening_dip")
        closing      = _n(f"ago{t}_closing_dip")
        ret          = _n(f"ago{t}_return_tank")
        updf_c       = _n(f"ago{t}_updf_consumption")
        cost_p       = _n(f"ago{t}_cost_price")
        selling_p    = _n(f"ago{t}_selling_price")

        if selling_p > 0:
            effective_price = selling_p
        elif cost_p > 0:
            effective_price = cost_p + AGO_MARGIN
        else:
            effective_price = 0.0

        net = opening - closing + ret - updf_c
        if net > 0:
            ago_volume     += net
            ago_sell_total += net * effective_price

    ago_avg_sell = (ago_sell_total / ago_volume) if ago_volume > 0 else 0.0
    ago_revenue  = round(ago_volume * ago_avg_sell)

    # ── 5. Non-fuel product revenue ───────────────────────────
    # Prefer Product Totals sheet (pt_) values; fall back to Cash & Sales (cs_)
    lubes_litres  = _n("cs_lubes_ltrs")
    lubes_revenue = _n("pt_lubricants") or _n("cs_lubes_ugx")
    lpg_kgs       = _n("cs_lpg_kgs")
    lpg_revenue   = _n("pt_lpg")        or _n("cs_lpg_ugx")
    tba_credits   = _n("pt_tba")        or _n("cs_tba_ugx")
    shop_sales    = _n("pt_shop")        or _n("cs_shop_sales")
    tyre_sales    = _n("cs_tyre_ugx")

    lpg_acc_rev   = _n("pt_lpg_acc")
    car_wash_rev  = _n("pt_car_wash")
    solar_rev     = _n("pt_solar")

    # ── 6. Cashless breakdown ─────────────────────────────────
    plus_card_pms           = _n("cs_plus_pms")
    plus_card_ago           = _n("cs_plus_ago")
    plus_card_others        = _n("cs_plus_others")
    plus_card_payment_total = plus_card_pms + plus_card_ago + plus_card_others
    plus_card_payment       = _n("cs_plus_card_credits")

    momo_pay       = _n("cs_momo")
    airtel_pay     = _n("cs_airtel")
    visa           = _n("cs_visa")
    credit_sales   = _n("cs_debtors_credit")
    other_payments = _n("cs_other_payment")

    cashless_total = (
        plus_card_payment_total + momo_pay + airtel_pay +
        visa + credit_sales + other_payments
    )

    # ── 7. Cash calculations ──────────────────────────────────
    fuel_revenue     = pms_revenue + ago_revenue
    non_fuel_revenue = (
        lubes_revenue + lpg_revenue + tba_credits +
        shop_sales + tyre_sales +
        lpg_acc_rev + car_wash_rev + solar_rev
    )
    total_sales_calc   = fuel_revenue + non_fuel_revenue
    total_cash         = total_sales_calc - cashless_total
    actual_cash_banked = _n("cs_cash_banked")

    # ── 8. Expenses ───────────────────────────────────────────
    # Use Daily Expenses sheet (dx_) if entered; fall back to shift expenses (cs_exp_)
    dx_total = (
        _n("dx_meals")       + _n("dx_generator")    +
        _n("dx_electricity") + _n("dx_water")         +
        _n("dx_salaries")    + _n("dx_stationery")    +
        _n("dx_security")    + _n("dx_sanitation")    +
        _n("dx_airtime")     + _n("dx_transport")     +
        _n("dx_nssf")        + _n("dx_sundries")      +
        _n("dx_maintenance") + _n("dx_vat")           +
        _n("dx_photocopy")   + _n("dx_tax_compliance")
    )

    cs_exp_total = (
        _n("cs_exp_umeme")         + _n("cs_exp_water")          +
        _n("cs_exp_security")      + _n("cs_exp_stationery")     +
        _n("cs_exp_generator")     + _n("cs_exp_meals")          +
        _n("cs_exp_transport")     + _n("cs_exp_salaries")       +
        _n("cs_exp_sanitary")      + _n("cs_exp_airtime")        +
        _n("cs_exp_misc")          + _n("cs_exp_shop_packaging")
    )

    total_expenses = dx_total if dx_total > 0 else cs_exp_total

    def _exp(dx_key, cs_key):
        v = _n(dx_key)
        return v if v > 0 else _n(cs_key)

    expense_umeme          = _exp("dx_electricity",  "cs_exp_umeme")
    expense_water          = _exp("dx_water",        "cs_exp_water")
    expense_security       = _exp("dx_security",     "cs_exp_security")
    expense_stationery     = _exp("dx_stationery",   "cs_exp_stationery")
    expense_generator      = _exp("dx_generator",    "cs_exp_generator")
    expense_meals          = _exp("dx_meals",        "cs_exp_meals")
    expense_transport      = _exp("dx_transport",    "cs_exp_transport")
    expense_salaries       = _exp("dx_salaries",     "cs_exp_salaries")
    expense_sanitary       = _exp("dx_sanitation",   "cs_exp_sanitary")
    expense_airtime        = _exp("dx_airtime",      "cs_exp_airtime")
    expense_misc           = _exp("dx_sundries",     "cs_exp_misc")
    expense_shop_packaging = _n("cs_exp_shop_packaging")

    # ── 9. Stock purchases and cash to bank ───────────────────
    stock_tba           = _n("cs_stock_tba")
    stock_lpg_acc       = _n("cs_stock_lpg_acc")
    stock_shop_purchase = _n("cs_stock_shop")
    cash_to_bank        = total_cash - total_expenses

    # ── 10. Delta ─────────────────────────────────────────────
    delta = actual_cash_banked - cash_to_bank

    # ── 11. Build record using exact DAILY_COLUMNS names ─────
    # NOTE: pms_price and ago_price now store the SELLING price
    # (actual if captured, estimated if not), not the cost price.
    # This matches what the report engine uses for revenue display.
    record = {
        "date":                    entry_date,
        "pms_volume":              round(pms_volume, 2),
        "ago_volume":              round(ago_volume, 2),
        "pms_price":               round(pms_avg_sell, 2),
        "ago_price":               round(ago_avg_sell, 2),
        "pms_revenue":             int(pms_revenue),
        "ago_revenue":             int(ago_revenue),
        "lubes_litres":            round(lubes_litres, 2),
        "lubes_revenue":           int(lubes_revenue),
        "lpg_kgs":                 round(lpg_kgs, 2),
        "lpg_revenue":             int(lpg_revenue),
        "tba_credits":             int(tba_credits),
        "plus_card_payment":       int(plus_card_payment),
        "shop_sales":              int(shop_sales),
        "plus_card_payment_total": int(plus_card_payment_total),
        "tyre_sales":              int(tyre_sales),
        "cashless_total":          int(cashless_total),
        "plus_card_pms":           int(plus_card_pms),
        "plus_card_ago":           int(plus_card_ago),
        "other_payments":          int(other_payments),
        "momo_pay":                int(momo_pay),
        "airtel_pay":              int(airtel_pay),
        "visa":                    int(visa),
        "credit_sales":            int(credit_sales),
        "total_cash":              int(total_cash),
        "expense_umeme":           int(expense_umeme),
        "expense_water":           int(expense_water),
        "expense_security":        int(expense_security),
        "expense_stationery":      int(expense_stationery),
        "expense_generator":       int(expense_generator),
        "expense_meals":           int(expense_meals),
        "expense_transport":       int(expense_transport),
        "expense_salaries":        int(expense_salaries),
        "expense_sanitary":        int(expense_sanitary),
        "expense_airtime":         int(expense_airtime),
        "expense_misc":            int(expense_misc),
        "expense_shop_packaging":  int(expense_shop_packaging),
        "total_expenses":          int(total_expenses),
        "stock_tba":               int(stock_tba),
        "stock_lpg_acc":           int(stock_lpg_acc),
        "stock_shop_purchase":     int(stock_shop_purchase),
        "cash_to_bank":            int(cash_to_bank),
        "actual_cash_banked":      int(actual_cash_banked),
        "delta":                   int(delta),
        "source_sheet":            f"daily_entry_{mode}",
    }

    # ── 12. Write to SQLite via save_daily_record() ───────────
    try:
        from src.database import save_daily_record
        success = save_daily_record(record, station_id)
        if not success:
            return {
                "success": False,
                "message": "Database write failed — check the log for details.",
                "record":  record,
            }
    except Exception as e:
        logger.error(f"Daily entry commit failed for {entry_date}: {e}", exc_info=True)
        return {
            "success": False,
            "message": f"Database write failed: {str(e)}",
            "record":  record,
        }

    logger.info(
        f"Daily entry committed: {entry_date} {shift} | "
        f"PMS {pms_volume:.1f}L @ {pms_avg_sell:.0f} UGX/L | "
        f"AGO {ago_volume:.1f}L @ {ago_avg_sell:.0f} UGX/L | "
        f"Fuel rev UGX {fuel_revenue:,.0f} | Delta UGX {delta:,.0f}"
    )

    return {
        "success": True,
        "message": (
            f"Entry for {entry_date} ({shift} Shift) saved. "
            f"PMS: {pms_volume:.1f}L, AGO: {ago_volume:.1f}L, "
            f"Delta: {'SURPLUS' if delta >= 0 else 'DEFICIT'} "
            f"UGX {abs(delta):,.0f}."
        ),
        "record": record,
    }