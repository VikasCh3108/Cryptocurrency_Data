import pandas as pd
import time
from datetime import datetime
import os

def test_file_contents(file_path):
    """Test if file contains all required fields"""
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, sheet_name='Live Data')
        else:
            df = pd.read_excel(file_path, sheet_name='Live Data', engine='odf')
            
        # Check required columns
        required_columns = [
            'Name',
            'Symbol',
            'Current Price (USD)',
            'Market Cap',
            '24-hour Trading Volume',
            'Price Change (24-hour percentage)'
        ]
        
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"❌ Missing columns in {file_path}: {missing_columns}")
            return False
            
        # Check if we have top 50 cryptocurrencies
        if len(df) < 50:
            print(f"❌ {file_path} contains only {len(df)} cryptocurrencies, expected 50")
            return False
            
        print(f"✅ {file_path} contains all required data")
        return True
        
    except Exception as e:
        print(f"❌ Error reading {file_path}: {e}")
        return False

def test_file_updates(file_path, check_interval=60, total_time=300):
    """Test if file is being updated regularly"""
    initial_mtime = os.path.getmtime(file_path)
    start_time = time.time()
    last_check_time = start_time
    
    print(f"\nMonitoring {file_path} for updates...")
    print(f"Initial modification time: {datetime.fromtimestamp(initial_mtime)}")
    
    while time.time() - start_time < total_time:
        current_time = time.time()
        if current_time - last_check_time >= check_interval:
            current_mtime = os.path.getmtime(file_path)
            if current_mtime > initial_mtime:
                print(f"✅ File updated at: {datetime.fromtimestamp(current_mtime)}")
                return True
            last_check_time = current_time
            print(f"Checking... (No update yet)")
        time.sleep(1)
    
    print("❌ No file updates detected within the test period")
    return False

def main():
    files_to_test = ['crypto_data.xlsx', 'crypto_data.ods']
    
    print("=== Testing File Contents ===")
    for file in files_to_test:
        test_file_contents(file)
    
    print("\n=== Testing Update Frequency ===")
    print("Monitoring files for updates (this will take about 5 minutes)...")
    for file in files_to_test:
        test_file_updates(file)

if __name__ == "__main__":
    main()
