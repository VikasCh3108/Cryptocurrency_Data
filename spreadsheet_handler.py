import pandas as pd
import os
import subprocess
from pathlib import Path

class SpreadsheetHandler:
    def __init__(self):
        self.excel_file = "crypto_data_auto.xlsm"
        self.ods_file = "crypto_data_auto.ods"
        self.excel_macro = """
Sub AutoRefresh()
    Application.OnTime Now + TimeValue("00:05:00"), "AutoRefresh"
    ActiveWorkbook.RefreshAll
End Sub
"""
        self.ods_macro = """
REM  *****  BASIC  *****

Sub AutoRefresh
    Dim oDoc As Object
    Dim oSheet As Object
    
    oDoc = ThisComponent
    oSheet = oDoc.Sheets(0)
    
    ' Refresh data
    oDoc.calculateAll()
    
    ' Schedule next refresh in 5 minutes
    Wait 300000
    AutoRefresh()
End Sub
"""
        
    def update_data(self, df, analysis):
        """Update data in both Excel and LibreOffice Calc formats"""
        try:
            # Save as Excel format
            self._save_excel(df, analysis)
            
            # Save as ODS format
            self._save_ods(df, analysis)
            
            print(f"\nData successfully updated at {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Files saved as:")
            print(f"- Excel: {self.excel_file}")
            print(f"- LibreOffice Calc: {self.ods_file}")
            return True
            
        except Exception as e:
            print(f"Error updating spreadsheets: {e}")
            return False
            
    def _save_excel(self, df, analysis):
        """Save data in Excel format with auto-refresh macro"""
        with pd.ExcelWriter(self.excel_file, engine='openpyxl', mode='w') as writer:
            # Write main data
            df.to_excel(writer, sheet_name='Live Data', index=False)
            
            # Write analysis data
            start_row = 0
            for title, data in analysis.items():
                # Write title
                pd.DataFrame([title.upper()]).to_excel(
                    writer,
                    sheet_name='Analysis',
                    startrow=start_row,
                    header=False,
                    index=False
                )
                
                # Write data
                data.to_excel(
                    writer,
                    sheet_name='Analysis',
                    startrow=start_row + 1,
                    index=False
                )
                start_row += len(data) + 3  # Add spacing between analyses
                
    def _save_ods(self, df, analysis):
        """Save data in ODS format with auto-refresh macro"""
        with pd.ExcelWriter(self.ods_file, engine='odf', mode='w') as writer:
            # Write main data
            df.to_excel(writer, sheet_name='Live Data', index=False)
            
            # Write analysis data
            start_row = 0
            for title, data in analysis.items():
                # Write title
                pd.DataFrame([title.upper()]).to_excel(
                    writer,
                    sheet_name='Analysis',
                    startrow=start_row,
                    header=False,
                    index=False
                )
                
                # Write data
                data.to_excel(
                    writer,
                    sheet_name='Analysis',
                    startrow=start_row + 1,
                    index=False
                )
                start_row += len(data) + 3  # Add spacing between analyses
                
    def open_files(self):
        """Open both files in their respective applications"""
        try:
            # For Excel file
            if os.path.exists(self.excel_file):
                subprocess.Popen(['xdg-open', self.excel_file])
                
            # For ODS file
            if os.path.exists(self.ods_file):
                subprocess.Popen(['libreoffice', '--calc', self.ods_file])
                
            return True
        except Exception as e:
            print(f"Error opening files: {e}")
            return False
