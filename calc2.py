
import pandas as pd
import os

# ✅ Set file paths
file_path = 'c:/Users/OmniXXX/Desktop/SToks/Sample.xlsx'
output_path = 'c:/Users/OmniXXX/Desktop/SToks/All_Beta_Alpha_Results.xlsx'

# ✅ Set risk-free rate (e.g., 6% annually → approx 0.06/252 daily)
risk_free_rate = 0.06 / 252  # Daily risk-free rate

# Load Excel file
xls = pd.ExcelFile(file_path)

# ✅ Parse Nifty_50 sheet
df_nifty = xls.parse('Nifty_50')
df_nifty = df_nifty.iloc[:, [0, 1]].copy()
df_nifty.columns = ['Date', 'Nifty']
df_nifty['Date'] = pd.to_datetime(df_nifty['Date'])
df_nifty.sort_values('Date', inplace=True)

# ✅ Calculate daily market return
df_nifty['Market_Return'] = df_nifty['Nifty'].pct_change()

# Prepare results list
results = []

# Loop through all other sheets
for sheet_name in xls.sheet_names:
    if sheet_name == 'Nifty_50':
        continue

    try:
        df_nav = xls.parse(sheet_name)
        df_nav = df_nav.iloc[:, [2, 3]].copy()  # Assuming NAV and Date
        df_nav.columns = ['NAV', 'Date']
        df_nav['Date'] = pd.to_datetime(df_nav['Date'])
        df_nav.sort_values('Date', inplace=True)

        # ✅ Merge with Nifty
        df_merged = pd.merge(df_nav, df_nifty[['Date', 'Market_Return']], on='Date')
        df_merged['Fund_Return'] = df_merged['NAV'].pct_change()

        # ✅ Drop NaNs
        df_returns = df_merged.dropna(subset=['Fund_Return', 'Market_Return'])

        # ✅ Compute Beta
        beta = df_returns['Fund_Return'].cov(df_returns['Market_Return']) / df_returns['Market_Return'].var()

        # ✅ Compute average returns
        Ri = df_returns['Fund_Return'].mean()
        Rm = df_returns['Market_Return'].mean()

        # ✅ Compute Jensen's Alpha
        alpha_daily = Ri - (risk_free_rate + beta * (Rm - risk_free_rate))
        alpha_annual = alpha_daily * 252  # Annualized alpha

        # ✅ Append results
        results.append({
            'Scheme': sheet_name,
            'Beta': round(beta, 4),
            'Avg Fund Return (Daily)': round(Ri, 6),
            'Avg Market Return (Daily)': round(Rm, 6),
            "Jensen's Alpha (Daily)": round(alpha_daily, 6),
            "Jensen's Alpha (Annualized)": round(alpha_annual, 6)
        })

    except Exception as e:
        print(f"[ERROR] Sheet '{sheet_name}': {e}")

# ✅ Save results to Excel
os.makedirs(os.path.dirname(output_path), exist_ok=True)

results_df = pd.DataFrame(results)
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    results_df.to_excel(writer, sheet_name='Alpha Beta Results', index=False)

print(f"[SUCCESS] Results saved to: {output_path}")
