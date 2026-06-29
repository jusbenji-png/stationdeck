"""
ai_engine.py
------------
StationDeck — AI Report Engine

PURPOSE:
  Receives the full metrics package from processor.py, builds a
  structured prompt, calls the OpenAI API, and returns a set of
  narrative report sections.

  If the OpenAI API is unavailable or the key is a placeholder,
  the engine automatically falls back to placeholder mode so
  the rest of the pipeline can still be tested.

MONTHLY SECTIONS (11):
  1.  executive_summary
  2.  fuel_sales_analysis
  3.  fuel_stock_movement          ← NEW — deliveries, loss/gain, stock position
  4.  product_performance          ← NEW — lubes, LPG, TBA, accessories, car wash
  5.  shop_analysis
  6.  payment_collection_analysis
  7.  expense_pnl_analysis
  8.  cash_reconciliation_analysis
  9.  debtors_depositors_analysis
  10. financial_position_analysis
  11. claims_analysis

ANNUAL SECTIONS (7):
  1.  annual_executive_summary
  2.  annual_fuel_performance
  3.  annual_revenue_breakdown
  4.  annual_payment_trends
  5.  annual_expense_analysis
  6.  annual_reconciliation
  7.  annual_outlook
"""

import os
import logging

logger = logging.getLogger(__name__)

# =============================================================
# AI BACKEND
# =============================================================
# Primary: route narrative generation through the StationDeck server (the
# OpenAI key lives only on the server). Fallback: local OPENAI_API_KEY (dev).
# Final fallback: deterministic placeholder text (never crashes a report).
_AI_SERVER_URL = "https://web-production-46077.up.railway.app/ai"

_MONTHLY_SYSTEM = (
    "You are a Senior Financial Analyst specialising in petroleum retail "
    "operations in East Africa. You write formal, professional operational "
    "reports for fuel station managers and company directors."
)
_ANNUAL_SYSTEM = (
    "You are a Senior Financial Analyst specialising in petroleum retail "
    "operations in East Africa. You write formal annual operational reports "
    "for fuel station directors and regional management."
)


def _server_chat(system, prompt, model, max_tokens, temperature,
                 station_name, machine_id):
    """Call the StationDeck server /ai endpoint. Returns text or None."""
    if not (station_name and machine_id):
        return None
    try:
        import requests
        r = requests.post(_AI_SERVER_URL, json={
            "station_name": station_name, "machine_id": machine_id,
            "system": system, "prompt": prompt,
            "model": model, "max_tokens": max_tokens, "temperature": temperature,
        }, timeout=120)
        data = r.json()
        if data.get("success"):
            return data.get("text")
        logger.error(f"Server AI failed ({r.status_code}): {data.get('message')}")
        return None
    except Exception as e:
        logger.error(f"Server AI request failed: {e}")
        return None


def _local_chat(system, prompt, model, max_tokens, temperature):
    """Fallback: call OpenAI directly with a local key (dev). Returns text or None."""
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key or api_key.startswith("sk-placeholder") or api_key == "your_api_key_here":
        return None
    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": system},
                      {"role": "user", "content": prompt}],
            temperature=temperature, max_tokens=max_tokens,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        logger.error(f"Local OpenAI call failed: {e}")
        return None


# =============================================================
# MAIN PUBLIC FUNCTIONS
# =============================================================

def generate_report(metrics: dict, period_label: str,
                    station_name: str = None, machine_id: str = None) -> dict:
    """
    Generate a full narrative monthly report from processed metrics.

    Attempts OpenAI API. Falls back to placeholder mode automatically.

    Args:
        metrics (dict):     Full metrics package from processor.py
        period_label (str): Human-readable period, e.g. "April 2026"

    Returns:
        dict: { "period", "mode", "report_text", "sections" }
    """
    prompt = _build_monthly_prompt(metrics, period_label)

    # 1) Server (key on server) → 2) local key (dev) → 3) placeholder
    raw_text = _server_chat(_MONTHLY_SYSTEM, prompt, "gpt-4o-mini", 4000, 0.4,
                            station_name, machine_id)
    if raw_text is None:
        raw_text = _local_chat(_MONTHLY_SYSTEM, prompt, "gpt-4o-mini", 4000, 0.4)

    if raw_text is None:
        logger.warning("No AI backend available — monthly report in placeholder mode.")
        return _placeholder_report(period_label, metrics)

    raw_text = raw_text.strip()
    sections = _parse_monthly_sections(raw_text)
    logger.info("Monthly report narrative generated.")
    return {
        "period":      period_label,
        "mode":        "live",
        "report_text": raw_text,
        "sections":    sections,
    }


def generate_annual_report(metrics: dict, period_label: str,
                           station_name: str = None, machine_id: str = None) -> dict:
    """
    Generate a full narrative annual report from processed metrics.

    Args:
        metrics (dict):     Full metrics package from processor.process_annual_report()
        period_label (str): e.g. "FY 2025/26 (July 2025 – June 2026)"

    Returns:
        dict: { "period", "mode", "report_text", "sections" }
    """
    prompt = _build_annual_prompt(metrics, period_label)

    # 1) Server (key on server) → 2) local key (dev) → 3) placeholder
    raw_text = _server_chat(_ANNUAL_SYSTEM, prompt, "gpt-4o-mini", 4000, 0.4,
                            station_name, machine_id)
    if raw_text is None:
        raw_text = _local_chat(_ANNUAL_SYSTEM, prompt, "gpt-4o-mini", 4000, 0.4)

    if raw_text is None:
        logger.warning("No AI backend available — annual report in placeholder mode.")
        return _placeholder_annual_report(period_label, metrics)

    raw_text = raw_text.strip()
    sections = _parse_annual_sections(raw_text)
    logger.info("Annual report narrative generated.")
    return {
        "period":      period_label,
        "mode":        "live",
        "report_text": raw_text,
        "sections":    sections,
    }


# =============================================================
# FORMATTING HELPERS
# =============================================================

def _ugx(val):
    try:    return f"UGX {int(val):,}"
    except: return "N/A"

def _vol(val):
    try:    return f"{float(val):,.2f}"
    except: return "N/A"

def _pct(val):
    try:    return f"{float(val):.1f}%"
    except: return "N/A"

def _n(val):
    try:    return f"{int(val):,}"
    except: return "N/A"


# =============================================================
# MONTHLY PROMPT BUILDER
# =============================================================

def _build_monthly_prompt(metrics: dict, period_label: str) -> str:
    """
    Builds the full structured prompt for a monthly report.
    Incorporates all metric sections from the enhanced processor.
    """

    days         = metrics.get("total_days", 0)
    cur          = metrics.get("currency", "UGX")

    # ── Fuel volumes and revenue ──────────────────────────────
    pms_vol      = _vol(metrics.get("pms_volume_total", 0))
    ago_vol      = _vol(metrics.get("ago_volume_total", 0))
    pms_rev      = _ugx(metrics.get("pms_revenue_total", 0))
    ago_rev      = _ugx(metrics.get("ago_revenue_total", 0))
    fuel_rev     = _ugx(metrics.get("total_fuel_revenue", 0))
    avg_daily    = _ugx(metrics.get("avg_daily_fuel_revenue", 0))

    # ── Fuel stock movement ───────────────────────────────────
    fuel_stock   = metrics.get("fuel_stock", {})
    pms_s        = fuel_stock.get("pms", {})
    ago_s        = fuel_stock.get("ago", {})

    pms_open     = _vol(pms_s.get("opening_dip_ltrs", 0))
    pms_close    = _vol(pms_s.get("closing_dip_ltrs", 0))
    pms_purch    = _vol(pms_s.get("total_purchases_ltrs", 0))
    pms_loss     = _vol(pms_s.get("loss_gain_ltrs", 0))
    pms_loss_v   = _ugx(pms_s.get("loss_gain_value_ugx", 0))
    pms_del_ct   = pms_s.get("delivery_count", 0)
    pms_stk_val  = _ugx(pms_s.get("closing_stock_value_ugx", 0))
    pms_cost     = _ugx(pms_s.get("avg_cost_price_ugx", 0))
    pms_sell     = _ugx(pms_s.get("avg_selling_price_ugx", 0))

    ago_open     = _vol(ago_s.get("opening_dip_ltrs", 0))
    ago_close    = _vol(ago_s.get("closing_dip_ltrs", 0))
    ago_purch    = _vol(ago_s.get("total_purchases_ltrs", 0))
    ago_loss     = _vol(ago_s.get("loss_gain_ltrs", 0))
    ago_loss_v   = _ugx(ago_s.get("loss_gain_value_ugx", 0))
    ago_del_ct   = ago_s.get("delivery_count", 0)
    ago_stk_val  = _ugx(ago_s.get("closing_stock_value_ugx", 0))
    ago_cost     = _ugx(ago_s.get("avg_cost_price_ugx", 0))
    ago_sell     = _ugx(ago_s.get("avg_selling_price_ugx", 0))

    # Format delivery details
    def _delivery_lines(deliveries):
        if not deliveries:
            return "  No deliveries recorded this period."
        lines = []
        for d in deliveries:
            lines.append(
                f"  {d['date']}: {int(d['litres']):,} L @ {_ugx(d['cost_price'])}/L "
                f"(Value: {_ugx(d['total_value'])})"
            )
        return "\n".join(lines)

    pms_deliveries_str = _delivery_lines(pms_s.get("deliveries", []))
    ago_deliveries_str = _delivery_lines(ago_s.get("deliveries", []))

    # ── Product sales ─────────────────────────────────────────
    prod         = metrics.get("product_sales", {})
    lub_ugx      = _ugx(prod.get("lubricants_ugx", metrics.get("lubes_revenue_total", 0)))
    lpg_ugx      = _ugx(prod.get("lpg_ugx", metrics.get("lpg_revenue_total", 0)))
    shop_ugx     = _ugx(prod.get("shop_ugx", metrics.get("shop_sales_total", 0)))
    tba_ugx      = _ugx(prod.get("tba_ugx", metrics.get("tba_revenue_total", 0)))
    lpg_acc_ugx  = _ugx(prod.get("lpg_accessories_ugx", metrics.get("lpg_accessories_total", 0)))
    car_wash_ugx = _ugx(prod.get("car_wash_ugx", metrics.get("car_wash_total", 0)))
    total_prod   = _ugx(prod.get("total_ugx", metrics.get("total_revenue", 0)))

    # ── Expenses ──────────────────────────────────────────────
    total_exp    = _ugx(metrics.get("total_expenses", 0))
    exp_stock    = metrics.get("expenses_detail_stock", {})
    exp_cf       = metrics.get("expense_detail", {})

    # Use stock detail if available, else cashflow detail
    exp_src = exp_stock if exp_stock.get("total_expenses", 0) > 0 else exp_cf

    exp_lines = []
    exp_label_map = {
        "salaries":    "Salaries",
        "electricity": "Electricity (UMEME)",
        "umeme":       "Electricity (UMEME)",
        "generator":   "Generator",
        "meals":       "Meals",
        "water":       "Water",
        "security":    "Security",
        "transport":   "Transport",
        "airtime":     "Airtime/Data",
        "nssf":        "NSSF",
        "maintenance": "Maintenance",
        "stationary":  "Stationery",
        "stationery":  "Stationery",
        "sanitation":  "Sanitation",
        "vat_tax":     "VAT/Tax",
        "misc":        "Miscellaneous",
    }
    seen_labels = set()
    for key, label in exp_label_map.items():
        if label in seen_labels:
            continue
        val = exp_src.get(key, 0)
        if val and float(val) > 0:
            exp_lines.append(f"  {label}: {_ugx(val)}")
            seen_labels.add(label)
    exp_detail_str = "\n".join(exp_lines) if exp_lines else "  Detail not available."

    # ── Shop detail ───────────────────────────────────────────
    shop_detail  = metrics.get("shop_sales_detail", {})
    shop_total   = _ugx(shop_detail.get("total_turnover", 0))
    shop_days    = shop_detail.get("trading_days", 0)
    shop_avg     = _ugx(shop_detail.get("avg_daily_sales", 0))
    top_cats     = shop_detail.get("top_3_categories", [])
    top_cats_str = ", ".join(
        [f"{c['category'].replace('_',' ').title()} ({_ugx(c['total'])})" for c in top_cats]
    ) if top_cats else "N/A"

    # ── Payment collection ────────────────────────────────────
    cash_col     = _ugx(metrics.get("cash_collected", 0))
    cashless_col = _ugx(metrics.get("cashless_collected", 0))
    cash_pct     = _pct(metrics.get("cash_percentage", 0))
    cashless_pct = _pct(metrics.get("cashless_percentage", 0))
    plus_card    = _ugx(metrics.get("plus_card_total", 0))
    visa         = _ugx(metrics.get("visa_total", 0))
    credit       = _ugx(metrics.get("credit_sales_total", 0))
    total_sales  = _ugx(metrics.get("total_sales", 0))

    # ── PNL ───────────────────────────────────────────────────
    pnl          = metrics.get("pnl", {})
    gross_income = _ugx(pnl.get("gross_income", 0))
    net_profit   = _ugx(pnl.get("net_profit", 0))
    price_change = _ugx(pnl.get("price_change_effect", 0))
    reserve      = _ugx(pnl.get("reserve_balance", 0))

    # ── Cash reconciliation ───────────────────────────────────
    banked       = _ugx(metrics.get("total_cash_banked", 0))
    expected     = _ugx(metrics.get("total_cash_expected", 0))
    delta        = _ugx(metrics.get("total_delta", 0))
    delta_status = metrics.get("delta_status", "UNKNOWN")
    anomaly_days = metrics.get("anomaly_days_count", 0)

    # ── Debtors / depositors ──────────────────────────────────
    debtors      = metrics.get("debtors", {})
    deb_total    = _ugx(debtors.get("total_outstanding", 0))
    deb_count    = len(debtors.get("customers", []))
    deb_overdue  = debtors.get("overdue_count", 0)
    depositors   = metrics.get("depositors", {})
    dep_active   = depositors.get("active_depositors", 0)
    dep_total    = _ugx(depositors.get("total_balance", 0))

    # ── Financial position ────────────────────────────────────
    fp           = metrics.get("financial_position", {})
    total_assets = _ugx(fp.get("total_assets", 0))
    total_liab   = _ugx(fp.get("total_liabilities", 0))
    net_position = _ugx(fp.get("net_position", 0))

    # ── Claims ────────────────────────────────────────────────
    claims       = metrics.get("claims", {})
    claims_count = len(claims.get("claims", []))
    claims_total = _ugx(claims.get("total_claims", 0))

    stock_note = (
        "NOTE: Fuel stock movement data is available for this period "
        "and should be incorporated into the fuel analysis section."
        if metrics.get("stock_data_available") else
        "NOTE: Fuel stock movement data was not available for this period. "
        "Base fuel analysis on cashflow volumes only."
    )

    prompt = f"""You are preparing a formal monthly operational report for a TotalEnergies fuel station in Uganda.

REPORT PERIOD: {period_label}
OPERATING DAYS: {days}
{stock_note}

========================================
FUEL OPERATIONS
========================================
PMS (Petrol):
  Volume Sold:      {pms_vol} litres
  Revenue:          {pms_rev}
  Opening Dip:      {pms_open} litres
  Closing Dip:      {pms_close} litres
  Purchases:        {pms_purch} litres ({pms_del_ct} deliveries)
  Loss/Gain:        {pms_loss} litres (Value: {pms_loss_v})
  Avg Cost Price:   {pms_cost}/litre
  Avg Sell Price:   {pms_sell}/litre
  Closing Stock Value: {pms_stk_val}
  PMS Deliveries:
{pms_deliveries_str}

AGO (Diesel):
  Volume Sold:      {ago_vol} litres
  Revenue:          {ago_rev}
  Opening Dip:      {ago_open} litres
  Closing Dip:      {ago_close} litres
  Purchases:        {ago_purch} litres ({ago_del_ct} deliveries)
  Loss/Gain:        {ago_loss} litres (Value: {ago_loss_v})
  Avg Cost Price:   {ago_cost}/litre
  Avg Sell Price:   {ago_sell}/litre
  Closing Stock Value: {ago_stk_val}
  AGO Deliveries:
{ago_deliveries_str}

Combined Fuel Revenue: {fuel_rev}
Average Daily Fuel Revenue: {avg_daily}

========================================
PRODUCT SALES (NON-FUEL)
========================================
Lubricants:        {lub_ugx}
LPG Gas:           {lpg_ugx}
LPG Accessories:   {lpg_acc_ugx}
TBA Credits:       {tba_ugx}
Car Wash:          {car_wash_ugx}
Shop Sales:        {shop_ugx}
  Shop Turnover:   {shop_total} over {shop_days} trading days
  Daily Average:   {shop_avg}
  Top Categories:  {top_cats_str}
Total All Products: {total_prod}

========================================
EXPENSES
========================================
Total Expenses: {total_exp}
Breakdown:
{exp_detail_str}

========================================
PAYMENT COLLECTION
========================================
Cash Collected:    {cash_col} ({cash_pct} of total)
Cashless Total:    {cashless_col} ({cashless_pct} of total)
  Plus Card:       {plus_card}
  Visa:            {visa}
  Credit Sales:    {credit}
Total Sales:       {total_sales}

========================================
PROFIT & LOSS
========================================
Gross Income:      {gross_income}
Price Change Effect:{price_change}
Total Expenses:    {total_exp}
Net Profit:        {net_profit}
Reserve Balance:   {reserve}

========================================
CASH RECONCILIATION
========================================
Cash Collected:    {cash_col}
Total Banked:      {banked}
Expected to Bank:  {expected}
Net Delta:         {delta} ({delta_status})
Days with Anomalies: {anomaly_days} of {days}

========================================
DEBTORS & DEPOSITORS
========================================
Active Debtor Accounts:    {deb_count}
Total Outstanding:         {deb_total}
Accounts Over Limit:       {deb_overdue}
Active Depositor Accounts: {dep_active}
Total Depositor Balance:   {dep_total}

========================================
FINANCIAL POSITION
========================================
Total Assets:      {total_assets}
Total Liabilities: {total_liab}
Net Position:      {net_position}

========================================
CLAIMS
========================================
Outstanding Claims: {claims_count} items worth {claims_total}

========================================
INSTRUCTIONS
========================================
Write a formal operational report with EXACTLY these 11 sections.
Start each section with its exact heading on its own line.
Write 2-3 professional paragraphs per section using specific numbers.
Do not invent figures. Write for a petroleum company director.

SECTION 1: EXECUTIVE SUMMARY
SECTION 2: FUEL SALES AND VOLUME ANALYSIS
SECTION 3: FUEL STOCK MOVEMENT AND DELIVERIES
SECTION 4: NON-FUEL PRODUCT PERFORMANCE
SECTION 5: SHOP PERFORMANCE
SECTION 6: PAYMENT COLLECTION ANALYSIS
SECTION 7: PROFIT AND LOSS ANALYSIS
SECTION 8: CASH RECONCILIATION AND OPERATIONAL INTEGRITY
SECTION 9: DEBTORS AND DEPOSITORS ANALYSIS
SECTION 10: FINANCIAL POSITION
SECTION 11: CLAIMS AND OUTSTANDING OBLIGATIONS
"""
    return prompt


# =============================================================
# ANNUAL PROMPT BUILDER
# =============================================================

def _build_annual_prompt(metrics: dict, period_label: str) -> str:
    """Builds the structured prompt for an annual report."""

    fy_label    = metrics.get("fy_label", period_label)
    months_data = metrics.get("monthly_breakdown", [])
    months_avail = metrics.get("months_with_data", 0)

    # Annual totals
    fuel_rev    = _ugx(metrics.get("total_fuel_revenue", 0))
    total_rev   = _ugx(metrics.get("total_revenue", 0))
    total_sales = _ugx(metrics.get("total_sales", 0))
    pms_vol     = _vol(metrics.get("pms_volume_total", 0))
    ago_vol     = _vol(metrics.get("ago_volume_total", 0))
    total_exp   = _ugx(metrics.get("total_expenses", 0))
    total_delta = _ugx(metrics.get("total_delta", 0))
    delta_stat  = metrics.get("delta_status", "UNKNOWN")
    cash_pct    = _pct(metrics.get("cash_percentage", 0))
    cashless_pct= _pct(metrics.get("cashless_percentage", 0))
    avg_monthly = _ugx(metrics.get("avg_monthly_fuel_revenue", 0))
    total_days  = _n(metrics.get("total_days", 0))

    # Stock annual
    fs          = metrics.get("fuel_stock", {})
    pms_s       = fs.get("pms", {})
    ago_s       = fs.get("ago", {})
    pms_loss    = _vol(pms_s.get("loss_gain_ltrs", 0))
    ago_loss    = _vol(ago_s.get("loss_gain_ltrs", 0))
    pms_purch   = _vol(pms_s.get("total_purchases_ltrs", 0))
    ago_purch   = _vol(ago_s.get("total_purchases_ltrs", 0))
    pms_del_ct  = _n(pms_s.get("delivery_count", 0))
    ago_del_ct  = _n(ago_s.get("delivery_count", 0))

    # Product sales annual
    ps          = metrics.get("product_sales", {})
    lub_total   = _ugx(ps.get("lubricants_ugx", metrics.get("lubes_revenue_total", 0)))
    lpg_total   = _ugx(ps.get("lpg_ugx", metrics.get("lpg_revenue_total", 0)))
    shop_total  = _ugx(ps.get("shop_ugx", metrics.get("shop_sales_total", 0)))

    # Monthly breakdown table
    breakdown_lines = []
    for row in months_data:
        if row.get("data_available"):
            breakdown_lines.append(
                f"  {row['label']:12}  "
                f"Fuel: {_ugx(row.get('fuel_revenue', 0)):>20}  "
                f"Sales: {_ugx(row.get('total_sales', 0)):>20}  "
                f"Delta: {_ugx(row.get('delta', 0)):>15}"
            )
        else:
            breakdown_lines.append(f"  {row['label']:12}  [data not yet available]")
    breakdown_str = "\n".join(breakdown_lines)

    prompt = f"""You are preparing a formal annual operational report for a TotalEnergies fuel station.

FINANCIAL YEAR: {fy_label}
MONTHS WITH DATA: {months_avail} of 12
TOTAL OPERATING DAYS: {total_days}

========================================
ANNUAL FUEL PERFORMANCE
========================================
PMS Volume Sold:    {pms_vol} litres
AGO Volume Sold:    {ago_vol} litres
Total Fuel Revenue: {fuel_rev}
Avg Monthly Fuel:   {avg_monthly}
PMS Purchases:      {pms_purch} litres ({pms_del_ct} deliveries)
AGO Purchases:      {ago_purch} litres ({ago_del_ct} deliveries)
PMS Annual Loss/Gain: {pms_loss} litres
AGO Annual Loss/Gain: {ago_loss} litres

========================================
ANNUAL REVENUE BREAKDOWN
========================================
Fuel Revenue:       {fuel_rev}
Lubricants:         {lub_total}
LPG Gas:            {lpg_total}
Shop Sales:         {shop_total}
Total Revenue:      {total_rev}
Total Sales:        {total_sales}

========================================
ANNUAL PAYMENT TRENDS
========================================
Cash Share:         {cash_pct}
Cashless Share:     {cashless_pct}
Total Expenses:     {total_exp}

========================================
ANNUAL RECONCILIATION
========================================
Net Annual Delta:   {total_delta} ({delta_stat})

========================================
MONTHLY PERFORMANCE BREAKDOWN
========================================
{breakdown_str}

========================================
INSTRUCTIONS
========================================
Write a formal annual operational report with EXACTLY these 7 sections.
Start each section with its exact heading on its own line.
Write 3-4 professional paragraphs per section using specific numbers.
Where months have no data, acknowledge this appropriately.
Write for a petroleum company director reviewing the full financial year.

SECTION 1: ANNUAL EXECUTIVE SUMMARY
SECTION 2: ANNUAL FUEL PERFORMANCE
SECTION 3: ANNUAL REVENUE BREAKDOWN
SECTION 4: ANNUAL PAYMENT TRENDS
SECTION 5: ANNUAL EXPENSE ANALYSIS
SECTION 6: ANNUAL CASH RECONCILIATION
SECTION 7: ANNUAL OUTLOOK AND RECOMMENDATIONS
"""
    return prompt


# =============================================================
# SECTION PARSERS
# =============================================================

def _parse_monthly_sections(report_text: str) -> dict:
    """Parse the 11-section monthly report into a named dict."""

    headings = [
        "SECTION 1: EXECUTIVE SUMMARY",
        "SECTION 2: FUEL SALES AND VOLUME ANALYSIS",
        "SECTION 3: FUEL STOCK MOVEMENT AND DELIVERIES",
        "SECTION 4: NON-FUEL PRODUCT PERFORMANCE",
        "SECTION 5: SHOP PERFORMANCE",
        "SECTION 6: PAYMENT COLLECTION ANALYSIS",
        "SECTION 7: PROFIT AND LOSS ANALYSIS",
        "SECTION 8: CASH RECONCILIATION AND OPERATIONAL INTEGRITY",
        "SECTION 9: DEBTORS AND DEPOSITORS ANALYSIS",
        "SECTION 10: FINANCIAL POSITION",
        "SECTION 11: CLAIMS AND OUTSTANDING OBLIGATIONS",
    ]
    keys = [
        "executive_summary",
        "fuel_sales_analysis",
        "fuel_stock_movement",
        "product_performance",
        "shop_analysis",
        "payment_collection_analysis",
        "expense_pnl_analysis",
        "cash_reconciliation_analysis",
        "debtors_depositors_analysis",
        "financial_position_analysis",
        "claims_analysis",
    ]
    return _parse_sections(report_text, headings, keys)


def _parse_annual_sections(report_text: str) -> dict:
    """Parse the 7-section annual report into a named dict."""

    headings = [
        "SECTION 1: ANNUAL EXECUTIVE SUMMARY",
        "SECTION 2: ANNUAL FUEL PERFORMANCE",
        "SECTION 3: ANNUAL REVENUE BREAKDOWN",
        "SECTION 4: ANNUAL PAYMENT TRENDS",
        "SECTION 5: ANNUAL EXPENSE ANALYSIS",
        "SECTION 6: ANNUAL CASH RECONCILIATION",
        "SECTION 7: ANNUAL OUTLOOK AND RECOMMENDATIONS",
    ]
    keys = [
        "annual_executive_summary",
        "annual_fuel_performance",
        "annual_revenue_breakdown",
        "annual_payment_trends",
        "annual_expense_analysis",
        "annual_reconciliation",
        "annual_outlook",
    ]
    return _parse_sections(report_text, headings, keys)


def _parse_sections(report_text: str, headings: list, keys: list) -> dict:
    """Generic section parser used by both monthly and annual."""

    text_upper = report_text.upper()
    positions  = [text_upper.find(h) for h in headings]

    if all(p == -1 for p in positions):
        logger.warning("Could not parse report sections. Returning full text.")
        return {"full_report": report_text}

    sections = {}
    for i, (key, pos) in enumerate(zip(keys, positions)):
        if pos == -1:
            sections[key] = ""
            continue
        next_positions = [p for p in positions[i + 1:] if p != -1]
        end_pos        = next_positions[0] if next_positions else len(report_text)
        section_text   = report_text[pos:end_pos].strip()
        lines          = section_text.split("\n")
        body_lines     = [l for l in lines[1:] if l.strip()]
        sections[key]  = "\n\n".join(body_lines)

    return sections


# =============================================================
# MONTHLY PLACEHOLDER REPORT
# =============================================================

def _placeholder_report(period_label: str, metrics: dict = None) -> dict:
    """
    Returns a realistic sample report for testing without API credits.
    Uses real metrics if provided.
    """
    m = metrics or {}

    total_rev    = _ugx(m.get("total_revenue", 0))
    fuel_rev     = _ugx(m.get("total_fuel_revenue", 0))
    pms_vol      = _vol(m.get("pms_volume_total", 0))
    ago_vol      = _vol(m.get("ago_volume_total", 0))
    cash_pct     = m.get("cash_percentage", 0)
    cashless_pct = m.get("cashless_percentage", 0)
    delta        = _ugx(m.get("total_delta", 0))
    status       = str(m.get("delta_status", "UNKNOWN")).lower()
    anomalies    = m.get("anomaly_days_count", 0)
    days         = m.get("total_days", 0)
    net_profit   = _ugx(m.get("pnl", {}).get("net_profit", 0))
    reserve      = _ugx(m.get("pnl", {}).get("reserve_balance", 0))
    gross_income = _ugx(m.get("pnl", {}).get("gross_income", 0))
    deb_total    = _ugx(m.get("debtors", {}).get("total_outstanding", 0))
    dep_total    = _ugx(m.get("depositors", {}).get("total_balance", 0))
    net_pos      = _ugx(m.get("financial_position", {}).get("net_position", 0))
    claims_total = _ugx(m.get("claims", {}).get("total_claims", 0))
    ps_top       = m.get("product_sales", {})
    shop_detail  = m.get("shop_sales_detail", {})
    shop_total   = _ugx(
        ps_top.get(
            "shop_ugx",
            m.get("shop_sales_total", shop_detail.get("total_turnover", 0))
        )
    )

    # Fuel stock
    fs         = m.get("fuel_stock", {})
    pms_s      = fs.get("pms", {})
    ago_s      = fs.get("ago", {})
    pms_loss   = pms_s.get("loss_gain_ltrs", 0)
    ago_loss   = ago_s.get("loss_gain_ltrs", 0)
    pms_purch  = _vol(pms_s.get("total_purchases_ltrs", 0))
    ago_purch  = _vol(ago_s.get("total_purchases_ltrs", 0))
    pms_del_ct = pms_s.get("delivery_count", 0)
    ago_del_ct = ago_s.get("delivery_count", 0)
    pms_stk    = _ugx(pms_s.get("closing_stock_value_ugx", 0))
    ago_stk    = _ugx(ago_s.get("closing_stock_value_ugx", 0))

    # Product sales
    ps          = m.get("product_sales", {})
    lub_ugx     = _ugx(ps.get("lubricants_ugx", m.get("lubes_revenue_total", 0)))
    lpg_ugx     = _ugx(ps.get("lpg_ugx", m.get("lpg_revenue_total", 0)))
    tba_ugx     = _ugx(ps.get("tba_ugx", m.get("tba_revenue_total", 0)))
    lpg_acc     = _ugx(ps.get("lpg_accessories_ugx", m.get("lpg_accessories_total", 0)))

    # Top shop categories
    top_cats     = m.get("shop_sales_detail", {}).get("top_3_categories", [])
    top_cats_str = ", ".join(
        [c["category"].replace("_", " ").title() for c in top_cats]
    ) if top_cats else "non-alcoholic beverages, confectionery, and fresh products"

    stock_available = m.get("stock_data_available", False)

    executive_summary = (
        f"The station recorded a total revenue of {total_rev} for the period of {period_label}, "
        f"reflecting consistent operational performance across all product categories. "
        f"Fuel sales remained the primary revenue driver, contributing {fuel_rev} to the overall "
        f"figure, with both PMS and AGO products performing within expected operational ranges. "
        f"The station maintained a balanced payment collection profile, with cashless transactions "
        f"accounting for {cashless_pct:.1f}% of total collections and cash transactions representing "
        f"{cash_pct:.1f}% of the period total. "
        f"Cash reconciliation for the period closed with a {status} position. "
        f"After accounting for all expenses and dealer obligations, the net profit for the period "
        f"was {net_profit}, with a reserve balance of {reserve} remaining after all commitments. "
        f"Overall, {period_label} represents a stable operational period with revenue performance "
        f"in line with expectations for a station of this scale and activity level."
    )

    fuel_sales_analysis = (
        f"Fuel sales for {period_label} were anchored by both PMS (petrol) and AGO (diesel) product lines. "
        f"PMS volumes reached {pms_vol} litres over the {days}-day operating period, while AGO volumes "
        f"totalled {ago_vol} litres, reflecting the station's dual-market position serving both "
        f"private motorists and commercial transport operators. "
        f"The total fuel revenue of {fuel_rev} demonstrates the station's capacity to sustain "
        f"high throughput across both product categories simultaneously. "
        f"Average daily fuel revenue stood at {_ugx(m.get('avg_daily_fuel_revenue', 0))}, "
        f"providing a reliable benchmark for planning and performance comparison in subsequent periods."
    )

    fuel_stock_movement = (
        f"Fuel stock movement for {period_label} recorded {pms_del_ct} PMS deliveries totalling "
        f"{pms_purch} litres and {ago_del_ct} AGO deliveries totalling {ago_purch} litres. "
        f"{'Stock movement data was available for this period and reflects accurate dip-based measurements.' if stock_available else 'Fuel stock movement data was not available for this period.'} "
        f"PMS stock loss/gain for the period was {pms_loss:.2f} litres, while AGO recorded "
        f"{ago_loss:.2f} litres. "
        f"The closing stock values stood at {pms_stk} for PMS and {ago_stk} for AGO, "
        f"representing the station's fuel asset position at period end. "
        f"Any loss figures outside acceptable variance thresholds should be reviewed against "
        f"TotalEnergies Uganda's standard loss/gain guidelines and investigated accordingly."
    )

    product_performance = (
        f"Non-fuel product performance for {period_label} reflected diversified revenue streams "
        f"across lubricants, LPG, and accessories. "
        f"Lubricants generated {lub_ugx}, while LPG gas contributed {lpg_ugx} and "
        f"LPG accessories added {lpg_acc}. "
        f"TBA credit sales for the period amounted to {tba_ugx}. "
        f"These non-fuel products collectively strengthen the station's revenue resilience "
        f"and reduce dependence on fuel margin alone. "
        f"Management is encouraged to monitor stock levels for high-velocity non-fuel products "
        f"to prevent out-of-stock situations that directly impact customer experience and revenue."
    )

    shop_analysis = (
        f"The station shop generated a total turnover of {shop_total} during {period_label}. "
        f"The top-performing categories were {top_cats_str}, which collectively accounted for "
        f"the majority of shop revenue. "
        f"Non-fuel retail continues to be a growing contributor to the station's overall revenue "
        f"diversification, complementing fuel income and improving per-customer revenue metrics. "
        f"Management is encouraged to monitor slow-moving categories and align stock levels with "
        f"actual sales velocity to reduce write-offs and improve working capital efficiency."
    )

    payment_analysis = (
        f"Payment collection for {period_label} reflects a station operating in a progressively "
        f"cashless commercial environment. Cashless transactions accounted for {cashless_pct:.1f}% of "
        f"total collections, while cash payments represented the remaining {cash_pct:.1f}%. "
        f"This distribution is consistent with broader trends in Uganda's petroleum retail sector, "
        f"where fleet card and mobile money adoption continues to grow among commercial customers. "
        f"Plus Card payments represented the dominant cashless channel, underscoring the importance "
        f"of the station's relationship with its fleet card operator. "
        f"Credit sales require continued monitoring to ensure timely settlement and to manage "
        f"the station's receivables exposure within acceptable risk parameters."
    )

    expense_pnl_analysis = (
        f"The Profit and Loss statement for {period_label} recorded a gross income of {gross_income} "
        f"before operational expenses. "
        f"After deducting all operational expenses totalling {_ugx(m.get('total_expenses', 0))}, "
        f"the net profit for the period was {net_profit}. "
        f"Following dealer obligations including shop rent, management fees, and savings commitments, "
        f"the reserve balance available to the dealer stood at {reserve}. "
        f"Salaries, utilities, and transport represented the primary expense categories. "
        f"Management should review any expense lines showing month-on-month increases against "
        f"approved budgets and seek authorisation for any unbudgeted expenditure."
    )

    reconciliation = (
        f"Cash reconciliation for {period_label} closed with a net delta of {delta}, "
        f"resulting in a {status} position for the period. "
        f"This outcome reflects the aggregate variance between cash amounts physically banked "
        f"and the theoretical cash expected based on recorded sales and cashless deductions. "
        f"A total of {anomalies} days out of {days} operating days recorded anomalies in their "
        f"daily cash reconciliation figures, indicating that variance management requires "
        f"continued supervisory attention. "
        f"Station management is advised to review individual anomaly days to determine whether "
        f"variances are attributable to timing differences, rounding, or procedural gaps in "
        f"daily cash handling."
    )

    debtors_analysis = (
        f"The station's debtor ledger as at the control date reflected a total outstanding balance "
        f"of {deb_total} across active customer accounts. "
        f"Credit management requires close attention to ensure that approved credit limits are "
        f"respected and that outstanding balances are settled within agreed payment terms. "
        f"Pre-payment depositor accounts held a combined balance of {dep_total}, "
        f"representing committed fuel and product purchases from institutional and government customers. "
        f"Management should ensure depositor balances are reconciled monthly against consumption "
        f"records and that expiring deposits are flagged for renewal or settlement."
    )

    financial_position_analysis = (
        f"The station's financial position as at the end of {period_label} reflected a net position "
        f"of {net_pos} after deducting all known liabilities. "
        f"Stock assets, cash holdings, trading account balances, and customer receivables formed "
        f"the primary components of the asset base. "
        f"Liabilities primarily comprised uninvoiced product, accrued but unpaid expenses, and "
        f"pre-payment balances owed to depositing customers. "
        f"The overall financial position indicates a well-capitalised station with adequate asset "
        f"coverage for its operational obligations."
    )

    claims_analysis = (
        f"Outstanding promotion and fuel claims as at the control date amounted to {claims_total}. "
        f"These claims include LPG promotion reimbursements, solar product promotions, and "
        f"UN/UPDF fuel consumption claims pending settlement from TotalEnergies Uganda. "
        f"Management should ensure that all claims are supported by required documentation "
        f"and submitted within prescribed claim windows to avoid forfeiture. "
        f"Timely follow-up with TotalEnergies territory representatives is recommended "
        f"to accelerate settlement of long-outstanding claims."
    )

    full_text = "\n\n".join([
        executive_summary, fuel_sales_analysis, fuel_stock_movement,
        product_performance, shop_analysis, payment_analysis,
        expense_pnl_analysis, reconciliation, debtors_analysis,
        financial_position_analysis, claims_analysis,
    ])

    return {
        "period":      period_label,
        "mode":        "placeholder",
        "report_text": full_text,
        "sections": {
            "executive_summary":           executive_summary,
            "fuel_sales_analysis":         fuel_sales_analysis,
            "fuel_stock_movement":         fuel_stock_movement,
            "product_performance":         product_performance,
            "shop_analysis":               shop_analysis,
            "payment_collection_analysis": payment_analysis,
            "expense_pnl_analysis":        expense_pnl_analysis,
            "cash_reconciliation_analysis":reconciliation,
            "debtors_depositors_analysis": debtors_analysis,
            "financial_position_analysis": financial_position_analysis,
            "claims_analysis":             claims_analysis,
        }
    }


# =============================================================
# ANNUAL PLACEHOLDER REPORT
# =============================================================

def _placeholder_annual_report(period_label: str, metrics: dict = None) -> dict:
    """Placeholder annual report for testing without API credits."""
    m = metrics or {}

    fy_label      = m.get("fy_label", period_label)
    fuel_rev      = _ugx(m.get("total_fuel_revenue", 0))
    total_rev     = _ugx(m.get("total_revenue", 0))
    pms_vol       = _vol(m.get("pms_volume_total", 0))
    ago_vol       = _vol(m.get("ago_volume_total", 0))
    total_exp     = _ugx(m.get("total_expenses", 0))
    delta         = _ugx(m.get("total_delta", 0))
    status        = str(m.get("delta_status", "UNKNOWN")).lower()
    avg_monthly   = _ugx(m.get("avg_monthly_fuel_revenue", 0))
    months_avail  = m.get("months_with_data", 0)
    cash_pct      = m.get("cash_percentage", 0)
    cashless_pct  = m.get("cashless_percentage", 0)
    lub_total     = _ugx(m.get("lubes_revenue_total", 0))
    lpg_total     = _ugx(m.get("lpg_revenue_total", 0))
    shop_total    = _ugx(m.get("shop_sales_total", 0))

    fs      = m.get("fuel_stock", {})
    pms_s   = fs.get("pms", {})
    ago_s   = fs.get("ago", {})
    pms_loss= _vol(pms_s.get("loss_gain_ltrs", 0))
    ago_loss= _vol(ago_s.get("loss_gain_ltrs", 0))

    annual_executive_summary = (
        f"This annual operational report covers {fy_label} "
        f"({period_label}) for the station. "
        f"Over the financial year, the station generated total fuel revenue of {fuel_rev} "
        f"and total revenue across all products of {total_rev}. "
        f"PMS sales totalled {pms_vol} litres and AGO sales reached {ago_vol} litres "
        f"across {months_avail} months of complete data. "
        f"Average monthly fuel revenue was {avg_monthly}, demonstrating consistent "
        f"throughput performance across the financial year. "
        f"The station maintained operational continuity across all product categories "
        f"and the annual cash reconciliation closed in a {status} position."
    )

    annual_fuel_performance = (
        f"Annual fuel performance across {fy_label} reflected strong and sustained throughput "
        f"in both PMS and AGO product categories. "
        f"PMS volumes for the year totalled {pms_vol} litres with an annual loss/gain of "
        f"{pms_loss} litres, while AGO volumes reached {ago_vol} litres with a loss/gain of "
        f"{ago_loss} litres. "
        f"Fuel delivery schedules were maintained across the year, with regular procurement "
        f"ensuring adequate stock availability during peak demand periods. "
        f"Loss/gain figures should be reviewed against TotalEnergies Uganda's annual "
        f"acceptable variance thresholds."
    )

    annual_revenue_breakdown = (
        f"Total revenue for {fy_label} was {total_rev}, with fuel contributing the dominant "
        f"share of {fuel_rev}. "
        f"Non-fuel revenue streams included lubricants at {lub_total}, LPG gas at {lpg_total}, "
        f"and shop sales of {shop_total}. "
        f"The diversification of revenue across multiple product categories provides the "
        f"station with resilience against fuel margin compression and price volatility. "
        f"Non-fuel revenue growth should continue to be prioritised as a strategic objective."
    )

    annual_payment_trends = (
        f"Payment collection trends across {fy_label} showed cashless transactions at "
        f"{cashless_pct:.1f}% of total collections, with cash at {cash_pct:.1f}%. "
        f"The progressive shift toward cashless payments reflects growing fleet card and "
        f"mobile money adoption among the station's commercial customer base. "
        f"This trend reduces cash handling risk and improves reconciliation accuracy. "
        f"Credit sales management requires ongoing attention to ensure receivables "
        f"do not exceed acceptable exposure levels."
    )

    annual_expense_analysis = (
        f"Total operational expenses for {fy_label} amounted to {total_exp}. "
        f"Salaries, utilities, and transport collectively represented the largest expense "
        f"categories across the financial year. "
        f"Management should review year-on-year expense trends and benchmark against "
        f"comparable TotalEnergies stations in the region. "
        f"Any expense categories showing significant growth should be investigated "
        f"and approved variances documented for audit purposes."
    )

    annual_reconciliation = (
        f"Annual cash reconciliation for {fy_label} closed with a net delta of {delta}, "
        f"resulting in a {status} position for the year. "
        f"Daily variance management across all operating months requires sustained "
        f"supervisory attention to ensure individual day anomalies are reviewed and explained. "
        f"The annual reconciliation position reflects the aggregate of monthly variances "
        f"and should be reviewed by station management and the territory representative."
    )

    annual_outlook = (
        f"Looking ahead, the station is well-positioned to sustain and grow its operational "
        f"performance in the coming financial year. "
        f"Key priorities should include optimising fuel procurement scheduling to minimise "
        f"stock-out risk, growing non-fuel revenue through active shop and LPG management, "
        f"and continuing to improve cashless payment adoption to reduce cash handling exposure. "
        f"Management is encouraged to use the monthly StationDeck reports as a consistent "
        f"operational review tool to identify trends and address variances proactively. "
        f"With disciplined execution across all product categories and continued investment "
        f"in operational controls, the station is expected to deliver improved performance "
        f"in the next financial year."
    )

    full_text = "\n\n".join([
        annual_executive_summary, annual_fuel_performance,
        annual_revenue_breakdown, annual_payment_trends,
        annual_expense_analysis, annual_reconciliation, annual_outlook,
    ])

    return {
        "period":      period_label,
        "mode":        "placeholder",
        "report_text": full_text,
        "sections": {
            "annual_executive_summary": annual_executive_summary,
            "annual_fuel_performance":  annual_fuel_performance,
            "annual_revenue_breakdown": annual_revenue_breakdown,
            "annual_payment_trends":    annual_payment_trends,
            "annual_expense_analysis":  annual_expense_analysis,
            "annual_reconciliation":    annual_reconciliation,
            "annual_outlook":           annual_outlook,
        }
    }