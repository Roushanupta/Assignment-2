import requests
import pandas as pd
import openpyxl
import schedule
import time

# Constants
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}
EXCEL_FILE = "crypto_data.xlsx"

def fetch_crypto_data():
    """Fetches live cryptocurrency data from CoinGecko API."""
    try:
        response = requests.get(API_URL, params=PARAMS)
        response.raise_for_status()
        data = response.json()

        crypto_list = []
        for coin in data:
            crypto_list.append({
                "Name": coin["name"],
                "Symbol": coin["symbol"].upper(),
                "Price (USD)": coin["current_price"],
                "Market Cap (USD)": coin["market_cap"],
                "24h Volume (USD)": coin["total_volume"],
                "24h Change (%)": coin["price_change_percentage_24h"]
            })

        return pd.DataFrame(crypto_list)

    except requests.exceptions.RequestException as e:
        print("Error fetching data:", e)
        return pd.DataFrame()

def analyze_data(df):
    """Performs analysis on cryptocurrency data."""
    if df.empty:
        return None

    # Top 5 by market cap
    top_5 = df.nlargest(5, "Market Cap (USD)")[["Name", "Market Cap (USD)"]]

    # Average price of top 50
    avg_price = df["Price (USD)"].mean()

    # Highest & lowest 24h percentage change
    highest_change = df.loc[df["24h Change (%)"].idxmax()]
    lowest_change = df.loc[df["24h Change (%)"].idxmin()]

    return {
        "Top 5 by Market Cap": top_5,
        "Average Price": avg_price,
        "Highest 24h Change": highest_change,
        "Lowest 24h Change": lowest_change
    }

def update_excel():
    """Fetches data, analyzes it, and updates an Excel file."""
    df = fetch_crypto_data()
    if df.empty:
        return

    analysis = analyze_data(df)

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Live Data", index=False)

        # Writing analysis to a separate sheet
        if analysis:
            analysis_df = pd.DataFrame({
                "Metric": ["Average Price", "Highest 24h Change", "Lowest 24h Change"],
                "Value": [analysis["Average Price"], 
                          f'{analysis["Highest 24h Change"]["Name"]} ({analysis["Highest 24h Change"]["24h Change (%)"]:.2f}%)',
                          f'{analysis["Lowest 24h Change"]["Name"]} ({analysis["Lowest 24h Change"]["24h Change (%)"]:.2f}%)']
            })
            analysis["Top 5 by Market Cap"].to_excel(writer, sheet_name="Analysis", index=False)
            analysis_df.to_excel(writer, sheet_name="Analysis", startrow=10, index=False)

    print("Excel updated successfully.")

# Schedule the script to run every 5 minutes
schedule.every(5).minutes.do(update_excel)

# Run the first update immediately
update_excel()

print("Live update started. Press Ctrl+C to stop.")
while True:
    schedule.run_pending()
    time.sleep(1)
