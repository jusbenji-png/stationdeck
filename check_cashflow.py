import pandas as pd

xl = pd.ExcelFile('data/input/Daily_cash_flow.xlsx')
print('Sheets:', xl.sheet_names)
print('Total sheets:', len(xl.sheet_names))

df = pd.read_excel('data/input/Daily_cash_flow.xlsx', sheet_name=xl.sheet_names[0], header=None, engine='openpyxl')
print('Shape:', df.shape)

df[0] = pd.to_datetime(df[0], errors='coerce')
df = df[df[0].notna()]
print('Valid date rows:', len(df))

mask = (df[0].dt.month == 4) & (df[0].dt.year == 2026)
dm = df[mask]
print('April 2026 rows:', len(dm))

if len(dm) > 0:
    row = dm.iloc[0]
    print('col0 (date):', row.iloc[0])
    print('col1:', row.iloc[1])
    print('col2:', row.iloc[2])
    print('col5:', row.iloc[5])
    print('col16:', row.iloc[16])
    print('col24:', row.iloc[24])
    print('col44:', row.iloc[44])
    print()
    print('April 2026 TOTALS:')
    print('col1 sum:', dm.iloc[:,1].sum())
    print('col5 sum:', dm.iloc[:,5].sum())
    print('col16 sum:', dm.iloc[:,16].sum())
    print('col24 sum:', dm.iloc[:,24].sum())
    print('col44 sum:', dm.iloc[:,44].sum())