# Mutual Fund Scheme Performance Analysis

This project provides a Python script to perform detailed financial analysis on mutual fund schemes using historical NAV data and a benchmark index. The analysis calculates key performance metrics such as Beta, Jensen's Alpha, Sharpe Ratio, CAGR, Sortino Ratio, and Treynor Ratio for each scheme relative to the benchmark. The results are ranked and saved to an Excel file for further review.

## Prerequisites

- Python 3.x
- pandas
- numpy
- scipy

You can install the required Python packages using pip:

```bash
pip install pandas numpy scipy
```

## Input Data

The script expects an Excel file (`Internship Beta Project.xlsx`) with the following structure:

- A sheet named **Nifty 50 Benchmark** containing date and benchmark price columns.
- A sheet named **Sheet1** (or the first sheet) containing a column `Scheme Code` listing the target scheme codes to analyze.
- Multiple sheets with names containing `NAV` that include mutual fund scheme data with columns for Date, Scheme Name, Net Asset Value, and Scheme Code.

## Usage

Run the script `calc.py`:

```bash
python calc.py
```

The script will read the input Excel file, perform the analysis, and generate an output Excel file.

## Output

The output is saved to `Fund_Scheme_Level_Analysis_Filtered.xlsx` and contains three sheets:

- **Filtered Schemes**: All analyzed schemes with calculated metrics.
- **Top 30 Schemes**: Top 30 schemes ranked by Sharpe Ratio and Jensen's Alpha.
- **Bottom 30 Schemes**: Bottom 30 schemes ranked by Sharpe Ratio and Jensen's Alpha.

## Notes

- The risk-free rate used in calculations is set to an annualized 7.6%, converted to a daily rate.
- The script suppresses warnings for cleaner output.
- Ensure the input Excel file paths in the script match your local setup.

---

Analysis complete! Check the output Excel file for detailed results.
