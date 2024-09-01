import yfinance as yf
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Define the input and output paths
input_path = r"C:\Models\Book1.xlsx"
output_path = input_path  # Overwrite the existing Excel file

# Load the tickers from the Excel workbook
df = pd.read_excel(input_path, sheet_name=0, header=1)  # Skip first row, read from the second row as headers

# Prepare the output DataFrame with the required columns
output_df = pd.DataFrame(columns=[
    'Ticker', 'Company', 'Price', 'Market Cap', 'EV', 'Debt', 
    'TTM Rev', 'TTM Rev Gr', 'Tgt Rev Gr CY', 
    'Tgt Rev Gr NY', 'Gross Profit', 'Gross Mgn', 
    'EBITDA', 'EBITDA Mgn', 'EV/TTM Rev', 'EV/Fwd Rev', 
    'EV/GP', 'EV/GP/Exp Gr', 'Rule of 40'
])

# Iterate over each ticker and fetch data
for ticker in df.iloc[:, 0].dropna():  # Assuming tickers are in the first column
    try:
        stock = yf.Ticker(ticker.strip())  # Strip any whitespace from ticker symbol
        info = stock.info
        
        if not info:
            raise ValueError(f"No data found for ticker {ticker}")

        # Extract required data, check if data exists and default to NaN if not available
        company_name = info.get('longName', 'N/A')
        current_price = info.get('currentPrice', float('nan'))
        market_cap = info.get('marketCap', float('nan'))
        enterprise_value = info.get('enterpriseValue', float('nan'))
        total_debt = info.get('totalDebt', float('nan'))
        ttm_revenue = info.get('totalRevenue', float('nan'))
        ttm_revenue_growth = info.get('revenueGrowth', float('nan'))
        gross_margin = info.get('grossMargins', float('nan'))
        ebitda = info.get('ebitda', float('nan'))
        ebitda_margin = info.get('ebitdaMargins', float('nan'))

        # Initialize with NaN in case the estimates are not available
        analyst_tgt_rev_growth_current_year = float('nan')
        analyst_tgt_rev_growth_next_year = float('nan')

        # Extract the revenue growth estimates from the revenue_estimate DataFrame
        try:
            revenue_estimates = stock.revenue_estimate
            if not revenue_estimates.empty:
                analyst_tgt_rev_growth_current_year = revenue_estimates.loc['0y', 'growth']
                analyst_tgt_rev_growth_next_year = revenue_estimates.loc['+1y', 'growth']
        except (KeyError, AttributeError):
            # Skip if revenue estimates data is missing
            pass

        # Fetch Gross Profit from Income Statement
        try:
            income_statement = stock.financials.T  # Transpose to make it easier to access by date
            gross_profit = income_statement.loc[:, 'Gross Profit'].iloc[0]
        except (KeyError, AttributeError, IndexError):
            gross_profit = float('nan')

        # Calculate additional metrics
        ev_ttm_rev = enterprise_value / ttm_revenue if ttm_revenue else float('nan')
        ev_fwd_rev = enterprise_value / (ttm_revenue * (1 + analyst_tgt_rev_growth_current_year)) if analyst_tgt_rev_growth_current_year else float('nan')
        ev_gp = enterprise_value / gross_profit if gross_profit else float('nan')
        ev_gp_exp_gr = ev_gp / analyst_tgt_rev_growth_current_year if analyst_tgt_rev_growth_current_year else float('nan')
        rule_of_40 = (ttm_revenue_growth + ebitda_margin) if ttm_revenue_growth and ebitda_margin else float('nan')

        # Create a new DataFrame row for the ticker data
        new_row = pd.DataFrame([{
            'Ticker': ticker.strip(),
            'Company': company_name,
            'Price': current_price,
            'Market Cap': market_cap,
            'EV': enterprise_value,
            'Debt': total_debt,
            'TTM Rev': ttm_revenue,
            'TTM Rev Gr': ttm_revenue_growth,
            'Tgt Rev Gr CY': analyst_tgt_rev_growth_current_year,
            'Tgt Rev Gr NY': analyst_tgt_rev_growth_next_year,
            'Gross Profit': gross_profit,
            'Gross Mgn': gross_margin,
            'EBITDA': ebitda,
            'EBITDA Mgn': ebitda_margin,
            'EV/TTM Rev': ev_ttm_rev,
            'EV/Fwd Rev': ev_fwd_rev,
            'EV/GP': ev_gp,
            'EV/GP/Exp Gr': ev_gp_exp_gr,
            'Rule of 40': rule_of_40
        }])

        # Append the new row to the output DataFrame using pd.concat
        output_df = pd.concat([output_df, new_row], ignore_index=True)

    except Exception as e:
        print(f"Error fetching data for {ticker}: {e}")
        continue

# Write the output DataFrame to Excel
with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    output_df.to_excel(writer, index=False, sheet_name='Stock Data')

print("Data extraction completed and saved to Excel.")
