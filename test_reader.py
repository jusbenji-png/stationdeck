# StationDeck - Reader Test Script
from src.reader import read_daily_cashflow
from config.settings import DATA_INPUT_DIR

def test():
    print("\n" + "="*60)
    print("  StationDeck - Testing Excel Reader")
    print("="*60 + "\n")

    file_path = DATA_INPUT_DIR / "Daily_cash_flow.xlsx"
    print(f"Reading file: {file_path}\n")

    df = read_daily_cashflow(file_path)

    if df is None:
        print("ERROR: Reader returned no data. Check logs above.")
        return

    print("\n" + "="*60)
    print("  RESULTS")
    print("="*60)
    print(f"\nTotal days loaded:  {len(df)}")
    print(f"Columns extracted:  {len(df.columns)}")
    print(f"Date range:         {df['date'].min()} to {df['date'].max()}")

    print("\n--- First 3 records ---")
    print(df[['date', 'pms_volume', 'ago_volume',
              'pms_revenue', 'ago_revenue',
              'total_cash', 'actual_cash_banked', 'delta']].head(3).to_string(index=False))

    print("\n--- Last 3 records ---")
    print(df[['date', 'pms_volume', 'ago_volume',
              'pms_revenue', 'ago_revenue',
              'total_cash', 'actual_cash_banked', 'delta']].tail(3).to_string(index=False))

    print("\n--- Quick Financial Summary ---")
    print(f"Total PMS volume sold:    {df['pms_volume'].sum():>15,.0f} litres")
    print(f"Total AGO volume sold:    {df['ago_volume'].sum():>15,.0f} litres")
    print(f"Total PMS revenue:        {df['pms_revenue'].sum():>15,.0f} UGX")
    print(f"Total AGO revenue:        {df['ago_revenue'].sum():>15,.0f} UGX")
    print(f"Total cash banked:        {df['actual_cash_banked'].sum():>15,.0f} UGX")
    print(f"Total delta (variance):   {df['delta'].sum():>15,.0f} UGX")

    print("\n" + "="*60)
    print("  Reader test complete.")
    print("="*60 + "\n")

if __name__ == "__main__":
    test()