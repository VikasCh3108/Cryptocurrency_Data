import pandas as pd
import os
import subprocess
import time
from pathlib import Path

class CalcUpdater:
    def __init__(self):
        self.calc_file = "crypto_data.ods"
        self.calc_path = os.path.abspath(self.calc_file)
        self.refresh_interval = 300  # 5 minutes
        
    def update_data(self, df, analysis):
        """Update data in LibreOffice Calc file"""
        try:
            # Create directory for the file if it doesn't exist
            os.makedirs(os.path.dirname(self.calc_path), exist_ok=True)
            
            # Write main data
            with pd.ExcelWriter(self.calc_file, engine='odf') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Live Data', index=False)
                
                # Write analysis data
                start_row = 0
                for title, data in analysis.items():
                    # Write title as a separate row
                    pd.DataFrame([title.upper()]).to_excel(
                        writer, 
                        sheet_name='Analysis',
                        startrow=start_row,
                        header=False,
                        index=False
                    )
                    
                    # Write the actual data
                    data.to_excel(
                        writer,
                        sheet_name='Analysis',
                        startrow=start_row + 1,
                        index=False
                    )
                    start_row += len(data) + 3  # Add spacing between analyses
            
            print(f"\nData successfully updated in Calc at {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"File location: {self.calc_path}")
            return True
            
        except Exception as e:
            print(f"Error updating Calc: {e}")
            return False

    def open_calc(self):
        """Open the file in LibreOffice Calc"""
        try:
            subprocess.Popen(['libreoffice', '--calc', self.calc_path])
            print(f"Opened {self.calc_file} in LibreOffice Calc")
            return True
        except Exception as e:
            print(f"Error opening Calc: {e}")
            return False
