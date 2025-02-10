import pandas as pd
import yfinance as yf
from datetime import datetime
import os
import time
from openpyxl import load_workbook

# Full path to the CSV file
csv_path = 'C:/Users/a_h_J/OneDrive/Documents/_Invest/AHJ_Financial_Spreadsheet/AHJ_Finance_Tickers.csv'

# Read ticker symbols from CSV
tickers_df = pd.read_csv(csv_path, header=None)
tickers = tickers_df[0].tolist()

# Define the columns for the DataFrame
columns = [
    'Ticker', 'Industry', 'Sector', 'Current_Price', 'Market_Cap', 
    'Cash', 'Debt', 'Enterprise_Value', 'Revenue_ttm', 'Net_Income_ttm', 
    'Adjusted_EBITDA_ttm', 'Free_Cash_Flow_ttm', 'Price_To_Earnings', 
    'EPS', 'Gross_Margin', 'Price_To_Book', 'Price_To_Sales', 
    'Earnings_Yield', 'Current_Ratio', 'Quick_Ratio', 'Debt_to_Equity', 
    'Return_On_Assets', 'Operating_Margin', 'Net_Profit_Margin', 
    'Interest_Coverage_Ratio', 'Dividend_Yield', 'Dividend', 'Beta', 'CAPM'
]

# Initialize an empty DataFrame
data = {col: [] for col in columns}

# Fetch financial data for each ticker
for ticker in tickers:
    stock = yf.Ticker(ticker)
    
    while True:
        try:
            info = stock.info
            break
        except yf.exceptions.YFRateLimitError:
            print(f"Rate limit exceeded for {ticker}. Waiting for 60 seconds before retrying...")
            time.sleep(10)
    
    data['Ticker'].append(ticker)
    data['Industry'].append(info.get('industry', 'N/A'))
    data['Sector'].append(info.get('sector', 'N/A'))
    data['Current_Price'].append(round(info.get('currentPrice', 0), 2))
    data['Market_Cap'].append(f"{info.get('marketCap', 0) / 1_000_000:.2f}M")
    data['Cash'].append(f"{info.get('totalCash', 0) / 1_000_000:.2f}M")
    data['Debt'].append(f"{info.get('totalDebt', 0) / 1_000_000:.2f}M")
    data['Enterprise_Value'].append(f"{info.get('enterpriseValue', 0) / 1_000_000:.2f}M")
    data['Revenue_ttm'].append(f"{info.get('totalRevenue', 0) / 1_000_000:.2f}M")
    data['Net_Income_ttm'].append(f"{info.get('netIncome', 0) / 1_000_000:.2f}M")
    data['Adjusted_EBITDA_ttm'].append(f"{info.get('ebitda', 0) / 1_000_000:.2f}M")
    data['Free_Cash_Flow_ttm'].append(f"{info.get('freeCashflow', 0) / 1_000_000:.2f}M")
    
    # Handle non-numeric trailingPE
    pe_ratio = info.get('trailingPE', 'N/A')
    data['Price_To_Earnings'].append(round(float(pe_ratio), 2) if pe_ratio != 'N/A' else 'N/A')

    data['EPS'].append(round(info.get('trailingEps', 0), 2))
    data['Gross_Margin'].append(round(info.get('grossMargins', 0) * 100, 2))
    
    # Additional financial metrics
    data['Price_To_Book'].append(round(info.get('priceToBook', 0), 2))
    data['Price_To_Sales'].append(round(info.get('priceToSalesTrailing12Months', 0), 2))
    
    # Handle non-numeric trailingPE for Earnings_Yield
    pe_ratio = info.get('trailingPE', float('inf'))
    data['Earnings_Yield'].append(round(1 / float(pe_ratio) * 100, 2) if pe_ratio != float('inf') else 'N/A')
    
    data['Current_Ratio'].append(round(info.get('currentRatio', 0), 2))
    data['Quick_Ratio'].append(round(info.get('quickRatio', 0), 2))
    data['Debt_to_Equity'].append(round(info.get('debtToEquity', 0), 2))
    data['Return_On_Assets'].append(round(info.get('returnOnAssets', 0) * 100, 2))
    data['Operating_Margin'].append(round(info.get('operatingMargins', 0) * 100, 2))
    data['Net_Profit_Margin'].append(round(info.get('profitMargins', 0) * 100, 2))
    data['Interest_Coverage_Ratio'].append(round(info.get('ebitda', 0) / info.get('totalDebt', float('inf')), 2))
    data['Dividend_Yield'].append(round(info.get('dividendYield', 0) * 100, 2))
    data['Dividend'].append(round(info.get('dividendRate', 0), 2))
    data['Beta'].append(round(info.get('beta', 0), 2))
    
    # CAPM Calculation: CAPM = Risk-free rate + Beta * Market Risk Premium
    risk_free_rate = 0.04  # Example 4% risk-free rate (can be updated dynamically)
    market_risk_premium = 0.06  # Example 6% market risk premium
    data['CAPM'].append(round(risk_free_rate + info.get('beta', 0) * market_risk_premium, 2))

    # Add a 50-millisecond delay between requests
    time.sleep(0.25)

# Convert dictionary to DataFrame
df = pd.DataFrame(data)

# Display the DataFrame
print("First 5 rows:\n", df.head(5).to_string(), "\n")
print("Last 5 rows:\n", df.tail(5).to_string())

# Generate the filename with the current date and time
now = datetime.now()
filename = f"AHJ_Finance_{now.strftime('%Y%m%d_%H%M')}.xlsx"

# Check if the file exists and delete if it does
if os.path.exists(filename):
    os.remove(filename)

# Save the DataFrame to an Excel file
df.to_excel(filename, index=False, startrow=9, engine='openpyxl')

# Auto-size columns
wb = load_workbook(filename)
ws = wb.active

for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column_letter].width = adjusted_width

wb.save(filename)

print(f"File saved as {filename}")
