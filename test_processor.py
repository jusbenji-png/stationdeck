# StationDeck - Processor Test
from src.reader import read_daily_cashflow
from src.processor import process_monthly_report, format_ugx, format_litres
from config.settings import DATA_INPUT_DIR

def test():
    print("\n" + "="*60)
    print("  StationDeck - Testing Data Processor")
    print("="*60 + "\n")

    file_path = DATA_INPUT_DIR / "Daily_cash_flow.xlsx"
    df = read_daily_cashflow(file_path)

    if df is None:
        print("ERROR: Could not load data.")
        return

    print("Processing: April 2026\n")
    metrics = process_monthly_report(df, 2026, 4)

    if metrics is None:
        print("ERROR: No data for that month.")
        return

    print("\n" + "="*60)
    print("  APRIL 2026 — STATION METRICS")
    print("="*60)

    print(f"\n REPORTING PERIOD")
    print(f"  Days covered:          {metrics['total_days']}")
    print(f"  Period:                {metrics['reporting_period']}")

    print(f"\n FUEL SALES")
    print(f"  PMS volume:            {format_litres(metrics['pms_volume_total'])}")
    print(f"  AGO volume:            {format_litres(metrics['ago_volume_total'])}")
    print(f"  PMS revenue:           {format_ugx(metrics['pms_revenue_total'])}")
    print(f"  AGO revenue:           {format_ugx(metrics['ago_revenue_total'])}")
    print(f"  Total fuel revenue:    {format_ugx(metrics['total_fuel_revenue'])}")
    print(f"  Avg daily revenue:     {format_ugx(metrics['avg_daily_fuel_revenue'])}")

    print(f"\n NON-FUEL REVENUE")
    print(f"  Lubricants:            {format_ugx(metrics['lubes_revenue_total'])}")
    print(f"  LPG Gas:               {format_ugx(metrics['lpg_revenue_total'])}")
    print(f"  Shop sales:            {format_ugx(metrics['shop_sales_total'])}")
    print(f"  Total non-fuel:        {format_ugx(metrics['total_nonfuel_revenue'])}")

    print(f"\n PAYMENT SPLIT")
    print(f"  Total sales:           {format_ugx(metrics['total_sales'])}")
    print(f"  Cash collected:        {format_ugx(metrics['cash_collected'])}")
    print(f"  Cashless collected:    {format_ugx(metrics['cashless_collected'])}")
    print(f"  Cash percentage:       {metrics['cash_percentage']}%")
    print(f"  Cashless percentage:   {metrics['cashless_percentage']}%")
    print(f"  Plus Card:             {format_ugx(metrics['plus_card_total'])}")
    print(f"  MoMo Pay:              {format_ugx(metrics['momo_total'])}")
    print(f"  Airtel Pay:            {format_ugx(metrics['airtel_total'])}")
    print(f"  Visa:                  {format_ugx(metrics['visa_total'])}")
    print(f"  Credit Sales:          {format_ugx(metrics['credit_sales_total'])}")

    print(f"\n EXPENSES")
    print(f"  Total expenses:        {format_ugx(metrics['total_expenses'])}")
    print(f"  Salaries:              {format_ugx(metrics['expense_salaries'])}")
    print(f"  UMEME (electricity):   {format_ugx(metrics['expense_umeme'])}")
    print(f"  Water:                 {format_ugx(metrics['expense_water'])}")
    print(f"  Security:              {format_ugx(metrics['expense_security'])}")
    print(f"  Transport:             {format_ugx(metrics['expense_transport'])}")
    print(f"  Misc:                  {format_ugx(metrics['expense_misc'])}")

    print(f"\n CASH RECONCILIATION")
    print(f"  Total banked:          {format_ugx(metrics['total_cash_banked'])}")
    print(f"  Expected to bank:      {format_ugx(metrics['total_cash_expected'])}")
    print(f"  Total delta:           {format_ugx(metrics['total_delta'])}")
    print(f"  Avg daily delta:       {format_ugx(metrics['avg_daily_delta'])}")
    print(f"  Delta status:          {metrics['delta_status']}")
    print(f"  Anomaly days:          {metrics['anomaly_days_count']}")

    print(f"\n TOTAL STATION REVENUE")
    print(f"  {format_ugx(metrics['total_revenue'])}")

    print("\n" + "="*60)
    print("  Processor test complete.")
    print("="*60 + "\n")

if __name__ == "__main__":
    test()