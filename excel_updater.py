import pandas as pd
from O365 import Account
import os
from dotenv import load_dotenv

load_dotenv()

class ExcelUpdater:
    def __init__(self):
        self.excel_file = "crypto_data.xlsx"
        self.client_id = os.getenv('MICROSOFT_CLIENT_ID')
        self.client_secret = os.getenv('MICROSOFT_CLIENT_SECRET')
        self.share_link = None
        
    def authenticate(self):
        """Authenticate with Microsoft 365"""
        try:
            if not self.client_id or not self.client_secret:
                print("Please set MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET in .env file")
                return False
                
            credentials = (self.client_id, self.client_secret)
            self.account = Account(credentials)
            if self.account.authenticate():
                print("Successfully authenticated with Microsoft 365")
                return True
            return False
        except Exception as e:
            print(f"Authentication Error: {e}")
            return False
            
    def update_data(self, df, analysis):
        """Update Excel file and upload to OneDrive"""
        try:
            # Create Excel writer
            with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Live Data', index=False)
                
                # Write analysis data
                start_row = 0
                for title, data in analysis.items():
                    # Write title
                    data.to_excel(writer, sheet_name='Analysis', 
                                startrow=start_row, index=False)
                    start_row += len(data) + 3  # Add spacing between analyses
            
            # Upload to OneDrive
            storage = self.account.storage()
            my_drive = storage.get_default_drive()
            
            # Upload file
            folder = my_drive.get_root_folder()
            file = folder.upload_file(self.excel_file)
            
            # Create shareable link
            permission = file.share_with_link(share_type='view')
            self.share_link = permission.share_link
            
            print(f"\nData successfully updated in Excel")
            print(f"View live updates at: {self.share_link}")
            return True
            
        except Exception as e:
            print(f"Error updating Excel: {e}")
            return False

    def get_share_link(self):
        """Get the shareable link"""
        return self.share_link
