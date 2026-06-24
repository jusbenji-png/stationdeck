# prompts/report_prompts.py
# ─────────────────────────────────────────────────────────────────────────────
# StationDeck — Report Prompt Builder
#
# PURPOSE:
#   Formats processed station metrics into a structured OpenAI prompt.
#   The AI receives only numeric summaries — never raw Excel data.
#
# USED BY:
#   src/ai_engine.py
# ─────────────────────────────────────────────────────────────────────────────


def build_report_prompt(metrics: dict, period_label: str) -> str:
    """
    Build a structured prompt for OpenAI report generation.

    Args:
        metrics (dict): The processed metrics dictionary from processor.py
        period_label (str): Human-readable period, e.g. "April 2026"

    Returns:
        str: A complete prompt string ready to send to the OpenAI API
    """

    # ── Helper: format large numbers with commas and UGX prefix ──────────────
    def ugx(value):
        try:
            return f"UGX {int(value):,}"
        except (TypeError, ValueError):
            return "UGX 0"

    def litres(value):
        try:
            return f"{float(value):,.2f} L"
        except (TypeError, ValueError):
            return "0.00 L"

    def pct(value):
        try:
            return f"{float(value):.1f}%"
        except (TypeError, ValueError):
            return "0.0%"

    # ── Extract all metrics with safe fallbacks ───────────────────────────────
    days            = metrics.get("days_covered", 0)
    pms_vol         = metrics.get("pms_volume_litres", 0)
    ago_vol         = metrics.get("ago_volume_litres", 0)
    pms_rev         = metrics.get("pms_revenue", 0)
    ago_rev         = metrics.get("ago_revenue", 0)
    fuel_rev        = metrics.get("total_fuel_revenue", 0)
    avg_daily       = metrics.get("avg_daily_revenue", 0)
    lubes_rev       = metrics.get("lubricants_revenue", 0)
    lpg_rev         = metrics.get("lpg_revenue", 0)
    shop_rev        = metrics.get("shop_sales_revenue", 0)
    non_fuel_rev    = metrics.get("total_non_fuel_revenue", 0)
    total_sales     = metrics.get("total_sales", 0)
    total_revenue   = metrics.get("total_revenue", 0)
    cash_col        = metrics.get("cash_collected", 0)
    cashless_col    = metrics.get("cashless_collected", 0)
    cash_pct        = metrics.get("cash_percentage", 0)
    cashless_pct    = metrics.get("cashless_percentage", 0)
    plus_card       = metrics.get("plus_card_total", 0)
    visa_pay        = metrics.get("visa_payments", 0)
    credit_sales    = metrics.get("credit_sales", 0)
    total_expenses  = metrics.get("total_expenses", 0)
    total_banked    = metrics.get("total_banked", 0)
    expected_bank   = metrics.get("expected_to_bank", 0)
    delta_total     = metrics.get("total_delta", 0)
    delta_status    = metrics.get("delta_status", "UNKNOWN")
    anomaly_days    = metrics.get("anomaly_days", 0)

    # ── Compose the prompt ────────────────────────────────────────────────────
    prompt = f"""
You are a Senior Financial Analyst specializing in petroleum retail operations in East Africa.
Your task is to write a formal Monthly Operations Report for a fuel station.

The report covers the period: {period_label}
Total operating days in this report: {days} days

Write in a formal, professional business tone suitable for station owners, operations managers,
and financial stakeholders. Use complete sentences and structured paragraphs.
Do not use bullet points. Do not use markdown formatting.
Do not invent data. Only use the figures provided below.

─────────────────────────────────────────────────────────
OPERATIONAL DATA FOR {period_label.upper()}
─────────────────────────────────────────────────────────

FUEL SALES
  PMS (Petrol) Volume Sold:     {litres(pms_vol)}
  AGO (Diesel) Volume Sold:     {litres(ago_vol)}
  PMS Revenue:                  {ugx(pms_rev)}
  AGO Revenue:                  {ugx(ago_rev)}
  Total Fuel Revenue:           {ugx(fuel_rev)}
  Average Daily Revenue:        {ugx(avg_daily)}

NON-FUEL REVENUE
  Lubricants Revenue:           {ugx(lubes_rev)}
  LPG Gas Revenue:              {ugx(lpg_rev)}
  Shop Sales Revenue:           {ugx(shop_rev)}
  Total Non-Fuel Revenue:       {ugx(non_fuel_rev)}

TOTAL PERFORMANCE
  Total Sales (all products):   {ugx(total_sales)}
  Total Revenue (fuel+non-fuel):{ugx(total_revenue)}

PAYMENT COLLECTION
  Cash Collected:               {ugx(cash_col)} ({pct(cash_pct)} of total)
  Cashless Collected:           {ugx(cashless_col)} ({pct(cashless_pct)} of total)
  Plus Card Payments:           {ugx(plus_card)}
  Visa Payments:                {ugx(visa_pay)}
  Credit Sales:                 {ugx(credit_sales)}

EXPENSES
  Total Operating Expenses:     {ugx(total_expenses)}

CASH RECONCILIATION
  Total Cash Banked:            {ugx(total_banked)}
  Expected Amount to Bank:      {ugx(expected_bank)}
  Net Delta (Variance):         {ugx(delta_total)}
  Reconciliation Status:        {delta_status}
  Days with Anomalies:          {anomaly_days} out of {days} days

─────────────────────────────────────────────────────────
REPORT SECTIONS TO WRITE
─────────────────────────────────────────────────────────

Write the following four sections in order. Use the exact section headings shown.
Each section should be 2 to 4 paragraphs of professional narrative.

SECTION 1: EXECUTIVE SUMMARY
Provide a high-level overview of the station's performance for {period_label}.
Summarise total revenue, fuel volumes, payment collection split, and reconciliation outcome.
Conclude with an overall performance assessment.

SECTION 2: FUEL SALES PERFORMANCE
Analyse PMS and AGO volumes and revenues in detail.
Comment on the balance between petrol and diesel sales.
Discuss average daily revenue and what it indicates about operational throughput.
Include any observations about non-fuel revenue contribution.

SECTION 3: PAYMENT COLLECTION AND CASHLESS ANALYSIS
Analyse the split between cash and cashless payments.
Comment on the significance of Plus Card, Visa, and credit sales figures.
Discuss what the cashless penetration rate indicates for the business.
Note any risks associated with the credit sales figure if it is significant.

SECTION 4: CASH RECONCILIATION AND OPERATIONAL INTEGRITY
Report on the total amount banked versus the expected amount.
Explain the delta and its status (surplus or shortfall).
Comment on the number of anomaly days and what this means for operational controls.
Provide a professional closing statement on the overall integrity of the period's operations.

─────────────────────────────────────────────────────────
IMPORTANT INSTRUCTIONS
─────────────────────────────────────────────────────────
- Write all four sections in full.
- Do not use bullet points or markdown.
- Use formal English throughout.
- Reference specific figures from the data provided above.
- Do not add sections beyond the four specified.
- End the report after Section 4.
"""

    return prompt.strip()