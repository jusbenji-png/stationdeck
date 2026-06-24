# check_columns.py - Raw Excel inspector
import pandas as pd

df_raw = pd.read_excel(
    "data/input/Daily_cash_flow.xlsx",
    sheet_name="CASH FLOW",
    header=None,
    engine='openpyxl'
)

# Show first 10 rows, all columns, with position numbers
print("RAW EXCEL DATA - First 10 rows:")
print("=" * 80)
for row_idx in range(min(10, len(df_raw))):
    print(f"\n--- Row {row_idx} ---")
    for col_idx in range(len(df_raw.columns)):
        val = df_raw.iloc[row_idx, col_idx]
        if pd.notna(val) and str(val).strip() not in ['', 'nan']:
            print(f"  Col[{col_idx:02d}] = {val}")

print("\n\nSHOWING ROWS 5-15 (likely where data starts):")
print("=" * 80)
for row_idx in range(5, min(15, len(df_raw))):
    print(f"\n--- Row {row_idx} ---")
    for col_idx in range(len(df_raw.columns)):
        val = df_raw.iloc[row_idx, col_idx]
        if pd.notna(val) and str(val).strip() not in ['', 'nan']:
            print(f"  Col[{col_idx:02d}] = {val}")