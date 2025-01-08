import requests
import pandas as pd
import time
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference, LineChart

url = "https://api.coingecko.com/api/v3/coins/markets"
params = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}

def fetch_and_update_excel():
    while True:
        # Fetch data
        response = requests.get(url, params=params)
        data = response.json()
        
        # Create DataFrame
        df = pd.DataFrame(data)
        df = df[['name', 'symbol', 'current_price', 'market_cap', 'total_volume', 'price_change_percentage_24h']]
        
        # Analysis
        top_5_by_market_cap = df.nlargest(5, 'market_cap')
        average_price = df['current_price'].mean()
        highest_change = df.loc[df['price_change_percentage_24h'].idxmax()]
        lowest_change = df.loc[df['price_change_percentage_24h'].idxmin()]

        # Write data to Excel with charts
        with pd.ExcelWriter("crypto_live_data.xlsx", engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="Top 50 Cryptos", index=False)
            top_5_by_market_cap.to_excel(writer, sheet_name="Top 5 by Market Cap", index=False)
            summary_data = {
                "Metric": ["Average Price", "Highest Change (24h)", "Lowest Change (24h)"],
                "Value": [average_price, highest_change['price_change_percentage_24h'], lowest_change['price_change_percentage_24h']],
                "Cryptocurrency": ["All Cryptos", highest_change['name'], lowest_change['name']]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name="Summary", index=False)
        
        # Load workbook for charting
        wb = load_workbook("crypto_live_data.xlsx")
        ws = wb["Top 50 Cryptos"]

        # Bar Chart for Market Cap
        bar_chart = BarChart()
        bar_chart.title = "Top 10 Market Cap"
        data = Reference(ws, min_col=4, min_row=2, max_row=11)  # Market Cap Data
        labels = Reference(ws, min_col=1, min_row=2, max_row=11)
        bar_chart.add_data(data, titles_from_data=False)
        bar_chart.set_categories(labels)
        ws.add_chart(bar_chart, "H2")

        # Line Chart for Price Change
        line_chart = LineChart()
        line_chart.title = "24h Price Change %"
        data = Reference(ws, min_col=6, min_row=2, max_row=11)  # Price Change Data
        line_chart.add_data(data, titles_from_data=False)
        line_chart.set_categories(labels)
        ws.add_chart(line_chart, "H20")

        # Save the workbook with charts
        wb.save("crypto_live_data.xlsx")

        print("Excel file updated with charts!")
        time.sleep(300)  # Sleep for 2 minutes

# Run the updater
fetch_and_update_excel()
