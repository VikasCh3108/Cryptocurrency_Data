import requests
import pandas as pd
from datetime import datetime
import time
import os
from pathlib import Path
from dotenv import load_dotenv
from spreadsheet_handler import SpreadsheetHandler
from report_generator import CryptoReportGenerator

# Load environment variables
load_dotenv()

class CryptoDataFetcher:
    def __init__(self):
        self.base_url = "https://api.coingecko.com/api/v3"
        self.spreadsheet_handler = SpreadsheetHandler()
        self.report_generator = CryptoReportGenerator()
        self.api_key = os.getenv('COINGECKO_API_KEY')

    def fetch_top_50_crypto(self):
        """Fetch top 50 cryptocurrencies data from CoinGecko API"""
        try:
            endpoint = f"{self.base_url}/coins/markets"
            params = {
                'vs_currency': 'usd',
                'order': 'market_cap_desc',
                'per_page': 50,
                'page': 1,
                'sparkline': False,
                'price_change_percentage': '24h'
            }
            
            # Add API key to headers if available
            headers = {}
            if self.api_key:
                headers['X-CG-API-KEY'] = self.api_key
            
            response = requests.get(endpoint, params=params, headers=headers)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            print(f"Error fetching data: {e}")
            return None

    def process_crypto_data(self, data):
        """Process the raw API data into a pandas DataFrame"""
        if not data:
            return None

        df = pd.DataFrame(data)
        df = df[[
            'name',
            'symbol',
            'current_price',
            'market_cap',
            'total_volume',
            'price_change_percentage_24h',
            'high_24h',
            'low_24h',
            'circulating_supply',
            'ath',
            'ath_change_percentage'
        ]]
        
        # Store numeric values for sorting
        df_numeric = df.copy()
        
        df.columns = [
            'Cryptocurrency Name',
            'Symbol',
            'Current Price (USD)',
            'Market Capitalization',
            '24-hour Trading Volume',
            'Price Change (24h %)',
            '24h High',
            '24h Low',
            'Circulating Supply',
            'All Time High',
            'ATH Change %'
        ]
        
        # Keep numeric values for analysis and report generation
        df['Current Price (USD)'] = df_numeric['current_price']
        df['Market Capitalization'] = df_numeric['market_cap']
        df['24-hour Trading Volume'] = df_numeric['total_volume']
        df['24h High'] = df_numeric['high_24h']
        df['24h Low'] = df_numeric['low_24h']
        df['All Time High'] = df_numeric['ath']
        df['Circulating Supply'] = df_numeric['circulating_supply']
        df['Price Change (24h %)'] = df_numeric['price_change_percentage_24h']
        df['ATH Change %'] = df_numeric['ath_change_percentage']
        
        return df, df_numeric

    def analyze_data(self, df, df_numeric):
        """Perform detailed analysis on the cryptocurrency data"""
        if df is None or df.empty:
            return None

        # Reset index to start from 1 for all dataframes
        df.index = df.index + 1
        
        analysis = {
            'Top 5 by Market Cap': df.head()[
                ['Cryptocurrency Name', 'Current Price (USD)', 'Market Capitalization', 'Price Change (24h %)', '24-hour Trading Volume']
            ],
            'Biggest Gainers': df.iloc[df_numeric['price_change_percentage_24h'].nlargest(3).index][
                ['Cryptocurrency Name', 'Current Price (USD)', 'Price Change (24h %)']
            ].reset_index(drop=True),
            'Biggest Losers': df.iloc[df_numeric['price_change_percentage_24h'].nsmallest(3).index][
                ['Cryptocurrency Name', 'Current Price (USD)', 'Price Change (24h %)']
            ].reset_index(drop=True),
            'Most Active': df.iloc[df_numeric['total_volume'].nlargest(3).index][
                ['Cryptocurrency Name', '24-hour Trading Volume', 'Price Change (24h %)']
            ].reset_index(drop=True),
            'Closest to ATH': df.iloc[df_numeric['ath_change_percentage'].nlargest(3).index][
                ['Cryptocurrency Name', 'Current Price (USD)', 'All Time High', 'ATH Change %']
            ].reset_index(drop=True)
        }
        return analysis

    def update_spreadsheets(self, df, analysis):
        """Update both Excel and LibreOffice Calc files with the latest data"""
        if df is None or df.empty:
            return False

        try:
            # Update both spreadsheet formats
            self.spreadsheet_handler.update_data(df, analysis)
            
            # Open files on first update
            if not hasattr(self, '_files_opened'):
                self.spreadsheet_handler.open_files()
                self._files_opened = True
            
            return True
        except Exception as e:
            print(f"Error updating spreadsheets: {e}")
            return False

    def run(self, interval=300):  # 300 seconds = 5 minutes
        """Main function to run the crypto data fetching and analysis continuously"""
        while True:
            print(f"\nFetching data at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            # Fetch and process data
            raw_data = self.fetch_top_50_crypto()
            df, df_numeric = self.process_crypto_data(raw_data)
            
            if df is not None:
                # Perform analysis
                analysis = self.analyze_data(df, df_numeric)
                
                # Update spreadsheets
                self.update_spreadsheets(df, analysis)
                
                # Debug column names
                print("\nDataFrame columns:", df.columns.tolist())
                
                # Generate PDF report
                report_file = self.report_generator.generate_report(df, analysis)
                print(f"\nGenerated report: {report_file}")
                
                # Print analysis results with clear formatting
                print("\n" + "="*100)
                print(f"CRYPTOCURRENCY MARKET ANALYSIS - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                print("="*100)
                
                print("\n🏆 TOP 5 CRYPTOCURRENCIES BY MARKET CAP:")
                print(analysis['Top 5 by Market Cap'].to_string())
                
                # Reset index to start from 1 for other sections
                for section in ['Biggest Gainers', 'Biggest Losers', 'Most Active', 'Closest to ATH']:
                    analysis[section].index = analysis[section].index + 1
                
                print("\n📈 BIGGEST GAINERS (24H):")
                print(analysis['Biggest Gainers'].to_string())
                
                print("\n📉 BIGGEST LOSERS (24H):")
                print(analysis['Biggest Losers'].to_string())
                
                print("\n🔄 MOST ACTIVE BY TRADING VOLUME:")
                print(analysis['Most Active'].to_string())
                
                print("\n⭐ CLOSEST TO ALL-TIME HIGH:")
                print(analysis['Closest to ATH'].to_string())
            
            print(f"\nNext update in {interval} seconds...")
            time.sleep(interval)

if __name__ == "__main__":
    fetcher = CryptoDataFetcher()
    try:
        fetcher.run()
    except KeyboardInterrupt:
        print("\nProgram terminated by user")
