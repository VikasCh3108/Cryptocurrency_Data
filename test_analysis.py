import pandas as pd

def test_analysis_sheet(file_path):
    """Test if analysis sheet contains required analysis"""
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, sheet_name='Analysis')
        else:
            df = pd.read_excel(file_path, sheet_name='Analysis', engine='odf')
            
        # Convert to string to make it easier to search
        content = df.to_string().lower()
        
        # Required analysis components
        required_analysis = [
            'market cap',
            'price',
            'change',
            'volume'
        ]
        
        missing_analysis = []
        for analysis in required_analysis:
            if analysis not in content:
                missing_analysis.append(analysis)
        
        if missing_analysis:
            print(f"❌ Missing analysis in {file_path}: {missing_analysis}")
            return False
            
        print(f"✅ {file_path} contains all required analysis")
        return True
        
    except Exception as e:
        print(f"❌ Error reading analysis sheet in {file_path}: {e}")
        return False

def main():
    files_to_test = ['crypto_data.xlsx', 'crypto_data.ods']
    
    print("=== Testing Analysis Sheet ===")
    for file in files_to_test:
        test_analysis_sheet(file)

if __name__ == "__main__":
    main()
