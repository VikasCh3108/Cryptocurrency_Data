# Cryptocurrency Data Analyzer

This project fetches and analyzes data for the top 50 cryptocurrencies using the CoinGecko API. Data is automatically updated every 5 minutes in both Excel and LibreOffice Calc formats.

## Features

- Fetches real-time data for top 50 cryptocurrencies
- Updates data every 5 minutes
- Saves data to both Excel (.xlsm) and LibreOffice Calc (.ods) formats
- Supports auto-refresh in both spreadsheet applications
- Performs basic analysis including:
  - Top 5 cryptocurrencies by market cap
  - Average price calculation
  - Highest and lowest 24-hour price changes
  - Trading volume analysis

## Setup

1. Make sure you have Python 3.x installed
2. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```

4. Set up your API key:
   - Rename `.env.example` to `.env`
   - Add your CoinGecko API key to the `.env` file:
     ```
     COINGECKO_API_KEY=your_api_key_here
     ```
   Note: The free tier of CoinGecko API doesn't require an API key, but having one increases rate limits.

## Usage

Run the script:
```bash
python crypto_analyzer.py
```

The script will:
- Continuously fetch cryptocurrency data every 5 minutes
- Save the data to both 'crypto_data_auto.xlsm' and 'crypto_data_auto.ods'
- Display analysis results in the terminal
- Automatically update both spreadsheet files

### Auto-Refresh Setup

#### For Excel (.xlsm file)

1. Open `crypto_data_auto.xlsm`
2. Press Alt + F11 to open the VBA editor
3. In the Project Explorer, double-click on 'ThisWorkbook'
4. Copy and paste this code:
   ```vba
   Private Sub Workbook_Open()
       Application.OnTime Now + TimeValue("00:00:05"), "AutoRefresh"
   End Sub
   ```
5. Save and close the VBA editor
6. Enable macros when prompted
7. The sheet will now auto-refresh every 5 minutes

#### For LibreOffice Calc (.ods file)

1. Open `crypto_data_auto.ods`
2. Go to Tools > Macros > Organize Macros > Basic
3. Create a new module named 'AutoRefresh'
4. Copy the macro code from the file
5. Save the macro
6. Go to Tools > Customize > Events
7. Select 'Document' and find 'Document Opens'
8. Click 'Macro' and select the AutoRefresh macro
9. Click OK and save

The sheet will now auto-refresh every 5 minutes.

#### Manual Refresh

If you prefer manual refresh:
- Excel: Press F9
- LibreOffice Calc: Press Ctrl+Shift+F9

To stop the script, press Ctrl+C.

## Data Fields

The following data is collected for each cryptocurrency:
- Cryptocurrency Name
- Symbol
- Current Price (USD)
- Market Capitalization
- 24-hour Trading Volume
- Price Change (24-hour percentage)
