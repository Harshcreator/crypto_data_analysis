# Cryptocurrency Market Tracker

## Overview
This Python-based project fetches real-time cryptocurrency market data from the CoinGecko API and generates detailed Excel reports with automated charts and analysis. The tool tracks the top 50 cryptocurrencies by market capitalization and provides various insights through organized spreadsheets and visualizations.

## Features
- Real-time data fetching from the CoinGecko API
- Automated Excel report generation
- Dynamic charts and visualizations
- Comprehensive market analysis, including:
  - Top 50 cryptocurrencies tracking
  - Top 5 cryptocurrencies by market cap
  - Price change analysis (24h)
  - Market statistics and summaries

## Data Points Tracked
- Cryptocurrency name and symbol
- Current price (USD)
- Market capitalization
- Total trading volume
- 24-hour price change percentage

## Generated Reports
The tool creates an Excel file (`crypto_live_data.xlsx`) with multiple sheets:

1. **Top 50 Cryptos Sheet**
   - Complete listing of the top 50 cryptocurrencies
   - Includes all tracked metrics
   - Features a bar chart for the top 10 market caps
   - Displays a line chart for 24-hour price changes

2. **Top 5 by Market Cap Sheet**
   - Focused view of the market leaders
   - All relevant metrics for top performers

3. **Summary Sheet**
   - Average price across all tracked cryptocurrencies
   - Highest 24-hour price change with corresponding cryptocurrency
   - Lowest 24-hour price change with corresponding cryptocurrency

## Technical Implementation
The script utilizes several Python libraries:
- `requests` for API communication
- `pandas` for data manipulation
- `openpyxl` for Excel file handling and chart generation

## Auto-Update Feature
The script automatically updates the data every 5 minutes to ensure current market information.

## Usage
1. Ensure Python is installed on your system.
2. Install the required libraries:
   ```bash
   pip install requests pandas openpyxl
   ```
3. Run the script:
   ```bash
   python crypto_market_tracker.py
   ```
4. The Excel file (`crypto_live_data.xlsx`) will be generated in your working directory and update automatically every 5 minutes.

## API Reference
This project uses the CoinGecko API v3:
- Endpoint: `https://api.coingecko.com/api/v3/coins/markets`
- Free tier with rate limiting
- No API key required for basic functionality

## Visualization Details
1. **Market Cap Bar Chart**
   - Displays the top 10 cryptocurrencies by market capitalization
   - Located at position H2 in the Excel worksheet

2. **Price Change Line Chart**
   - Shows 24-hour price changes for top cryptocurrencies
   - Located at position H20 in the Excel worksheet

## Contributing
Feel free to fork this repository and submit pull requests for any improvements.

## Author
Harsh Singh (harshsingh220603@gmail.com)