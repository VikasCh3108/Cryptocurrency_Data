import pandas as pd
import time
from datetime import datetime

def read_price_data(file):
    try:
        if file.endswith('.xlsx'):
            df = pd.read_excel(file, sheet_name='Live Data')
        else:
            df = pd.read_excel(file, sheet_name='Live Data', engine='odf')
        return df['Current Price (USD)'].head(3).to_dict()
    except Exception as e:
        print(f"Error reading {file}: {e}")
        return None

def monitor_price_changes():
    files = ['crypto_data.xlsx', 'crypto_data.ods']
    last_prices = {file: None for file in files}
    
    while True:
        for file in files:
            current_prices = read_price_data(file)
            if current_prices:
                if last_prices[file] and current_prices != last_prices[file]:
                    print(f"\n{datetime.now()} - {file} updated!")
                    print("Price changes detected:")
                    for coin, price in current_prices.items():
                        if last_prices[file][coin] != price:
                            print(f"  {coin}: {last_prices[file][coin]} -> {price}")
                last_prices[file] = current_prices
        time.sleep(30)  # Check every 30 seconds

if __name__ == "__main__":
    print("Monitoring price updates... (Press Ctrl+C to stop)")
    monitor_price_changes()
