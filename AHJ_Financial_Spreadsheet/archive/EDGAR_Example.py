import json
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

# Load ticker symbols from JSON file
def load_ticker_symbols(file_path):
    with open(file_path, "r") as file:
        return json.load(file)

# Fetch recent quarterly filings from EDGAR
def fetch_quarterly_filings(ticker):
    print(f"Fetching filings for: {ticker}")
    base_url = f"https://www.sec.gov/cgi-bin/browse-edgar"
    params = {
        "action": "getcompany",
        "CIK": ticker,
        "type": "10-Q",
        "count": 8,  # Fetch the last 8 quarterly reports
        "output": "atom"
    }
    headers = {"User-Agent": "MyApp/1.0"}
    response = requests.get(base_url, params=params, headers=headers)
    if response.status_code == 200:
        return parse_edgar_response(response.text, ticker)
    else:
        print(f"Error: Failed to fetch filings for {ticker}. Status: {response.status_code}")
        return None

# Parse the EDGAR response to extract filing URLs
def parse_edgar_response(response_text, ticker):
    soup = BeautifulSoup(response_text, "html.parser")
    filings = []
    for entry in soup.find_all("entry"):
        filing_url = entry.find("link", {"rel": "alternate"})["href"]
        filing_date = entry.find("updated").text
        filings.append({"ticker": ticker, "filing_date": filing_date, "url": filing_url})
    return filings

# Extract financial metrics from filings (placeholder function)
def extract_metrics_from_filing(url):
    # This is a placeholder. Parsing requires logic specific to the filing structure.
    # You would parse the HTML/XBRL for metrics like Total Revenue, Net Income, etc.
    print(f"Extracting data from: {url}")
    return {
        "Total Revenue": "N/A",
        "Net Income": "N/A",
        "Earnings per Share": "N/A",
        "Adjusted EPS": "N/A",
        "Operating Income": "N/A",
        "Operating Margin": "N/A",
        "Gross Margin": "N/A",
        "Net Margin": "N/A",
    }

# Main script
def main():
    # Load ticker symbols
    tickers = load_ticker_symbols("tickers.json")
    all_data = []

    # Fetch and process filings for each ticker
    for ticker in tickers:
        filings = fetch_quarterly_filings(ticker)
        if filings:
            for filing in filings:
                metrics = extract_metrics_from_filing(filing["url"])
                metrics["ticker"] = ticker
                metrics["filing_date"] = filing["filing_date"]
                all_data.append(metrics)

    # Save results to a CSV file
    output_file = f"financial_metrics_{datetime.now().strftime('%Y%m%d')}.csv"
    df = pd.DataFrame(all_data)
    df.to_csv(output_file, index=False)
    print(f"Data saved to {output_file}")

if __name__ == "__main__":
    main()
