import pandas as pd
import numpy as np
from scipy.stats import linregress
import warnings

warnings.filterwarnings("ignore")

DATA_PATH = r"C:\Users\OmniXXX\Desktop\SToks\Internship Beta Project.xlsx"
OUTPUT_PATH = r"C:\Users\OmniXXX\Desktop\SToks\Fund_Scheme_Level_Analysis_Filtered.xlsx"

def calculate_metrics(fund_returns, benchmark_returns, risk_free_rate=0.076 / 252):
    mask = ~np.isnan(fund_returns) & ~np.isnan(benchmark_returns)
    fund_returns = fund_returns[mask]
    benchmark_returns = benchmark_returns[mask]

    if len(fund_returns) < 2:
        return np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan

    # CAPM Regression using linregress
    slope, intercept, r_value, p_value, std_err = linregress(benchmark_returns, fund_returns)
    beta = slope

    # Jensen's Alpha (annualized, adjusting for risk-free rate)
    mean_fund_return = np.mean(fund_returns)
    mean_benchmark_return = np.mean(benchmark_returns)
    alpha_daily = (mean_fund_return - risk_free_rate) - beta * (mean_benchmark_return - risk_free_rate)
    alpha_annual = alpha_daily * 252

    # Standard Deviation (volatility, daily)
    std_dev_daily = np.std(fund_returns, ddof=1)

    # Sharpe Ratio (annualized)
    sharpe = (mean_fund_return - risk_free_rate) / std_dev_daily if std_dev_daily != 0 else np.nan

    # CAGR
    total_return = np.prod(1 + fund_returns) - 1
    num_days = len(fund_returns)
    cagr = (1 + total_return) ** (252 / num_days) - 1

    # Sortino Ratio
    downside_returns = fund_returns[fund_returns < 0]
    downside_std = np.std(downside_returns, ddof=1)
    sortino = (mean_fund_return - risk_free_rate) / downside_std if downside_std != 0 else np.nan

    # Treynor Ratio
    treynor = (mean_fund_return - risk_free_rate) / beta if beta != 0 else np.nan

    return beta, alpha_annual, std_dev_daily, sharpe, cagr, sortino, treynor

def load_benchmark_returns(xls, sheet_name):
    df = pd.read_excel(xls, sheet_name=sheet_name)
    date_col = df.columns[0]
    price_col = df.columns[1]

    df = df[[date_col, price_col]].dropna()
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col])
    df = df.sort_values(by=date_col)

    returns = df[price_col].pct_change().dropna().values
    return returns

def main():
    xls = pd.ExcelFile(DATA_PATH)
    sheets = xls.sheet_names

    # Load benchmark returns
    benchmark_sheet = 'Nifty 50 Benchmark'
    benchmark_returns = load_benchmark_returns(xls, benchmark_sheet)

    # Load target scheme codes from Sheet1
    filter_sheet = sheets[0]  # Assuming Sheet1 is the first sheet
    df_filter = pd.read_excel(xls, sheet_name=filter_sheet)
    
    if 'Scheme Code' not in df_filter.columns:
        raise ValueError("Sheet1 must contain 'Scheme Code' column.")

    target_scheme_codes = set(df_filter['Scheme Code'].dropna().astype(str))

    results = []

    for sheet in sheets:
        if 'NAV' in sheet:
            df = pd.read_excel(xls, sheet_name=sheet)
            columns = df.columns.tolist()

            date_col = [col for col in columns if 'Date' in col][0]
            scheme_col = [col for col in columns if 'Scheme Name' in col][0]
            nav_col = [col for col in columns if 'Net Asset Value' in col][0]
            scheme_code_col = [col for col in columns if 'Scheme Code' in col][0]

            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df = df.dropna(subset=[date_col, scheme_col, nav_col, scheme_code_col])
            df = df.sort_values(by=date_col)

            amc_name = sheet.replace(' NAV', '').strip()

            for scheme_code in df[scheme_code_col].unique():
                if str(scheme_code) not in target_scheme_codes:
                    continue  # Skip schemes not in Sheet1

                scheme_data = df[df[scheme_code_col] == scheme_code]
                if scheme_data.shape[0] < 2:
                    continue

                fund_series = scheme_data[[date_col, nav_col]].dropna().sort_values(by=date_col)
                fund_returns = fund_series[nav_col].pct_change().dropna().values

                if len(fund_returns) == 0:
                    continue

                min_len = min(len(fund_returns), len(benchmark_returns))
                aligned_fund_returns = fund_returns[-min_len:]
                aligned_benchmark_returns = benchmark_returns[-min_len:]

                beta, alpha, std_dev, sharpe, cagr, sortino, treynor = calculate_metrics(
                    aligned_fund_returns, aligned_benchmark_returns
                )

                scheme_name = scheme_data[scheme_col].iloc[0]

                results.append({
                    'AMC': amc_name,
                    'Scheme Name': scheme_name,
                    'Scheme Code': scheme_code,
                    'Beta': beta,
                    'Jensen Alpha': alpha,
                    'Standard Deviation': std_dev,
                    'Sharpe Ratio': sharpe,
                    'CAGR': cagr,
                    'Sortino Ratio': sortino,
                    'Treynor Ratio': treynor
                })

    df_results = pd.DataFrame(results)

    df_results['Rank Sharpe'] = df_results['Sharpe Ratio'].rank(ascending=False, method='min')
    df_results['Rank Alpha'] = df_results['Jensen Alpha'].rank(ascending=False, method='min')

    top_30 = df_results.sort_values(by=['Sharpe Ratio', 'Jensen Alpha'], ascending=False).head(30)
    bottom_30 = df_results.sort_values(by=['Sharpe Ratio', 'Jensen Alpha'], ascending=True).head(30)

    with pd.ExcelWriter(OUTPUT_PATH) as writer:
        df_results.to_excel(writer, sheet_name='Filtered Schemes', index=False)
        top_30.to_excel(writer, sheet_name='Top 30 Schemes', index=False)
        bottom_30.to_excel(writer, sheet_name='Bottom 30 Schemes', index=False)

    print(f">>> Filtered analysis complete! Output saved to {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
    
# This code calculates various financial metrics for mutual fund schemes based on their NAVs and compares them against a benchmark.
# It filters schemes based on a list of target scheme codes and outputs the results to an Excel file.
# Ensure that the paths for DATA_PATH and OUTPUT_PATH are correctly set before running the script.  