import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
from datetime import datetime
import os

class CryptoReportGenerator:
    def __init__(self):
        self.pdf = FPDF()
        self.pdf.set_auto_page_break(auto=True, margin=15)
        self.pdf.add_page()
        
    def _add_header(self):
        """Add report header"""
        self.pdf.set_font('Arial', 'B', 24)
        self.pdf.cell(0, 20, 'Cryptocurrency Market Analysis', ln=True, align='C')
        self.pdf.set_font('Arial', 'I', 12)
        self.pdf.cell(0, 10, f'Generated on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', ln=True, align='C')
        self.pdf.ln(10)
        
    def _add_section_header(self, title):
        """Add section header"""
        self.pdf.set_font('Arial', 'B', 16)
        self.pdf.cell(0, 10, title, ln=True)
        self.pdf.ln(5)
        
    def _add_table(self, df, title):
        """Add table to the report"""
        self.pdf.set_font('Arial', 'B', 12)
        
        # Calculate column widths
        col_widths = []
        for col in df.columns:
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            col_widths.append(min(max_length * 5, 60))  # Limit column width
            
        # Add headers
        for i, col in enumerate(df.columns):
            self.pdf.cell(col_widths[i], 10, str(col), 1)
        self.pdf.ln()
        
        # Add data
        self.pdf.set_font('Arial', '', 10)
        for _, row in df.iterrows():
            for i, value in enumerate(row):
                self.pdf.cell(col_widths[i], 10, str(value), 1)
            self.pdf.ln()
        
        self.pdf.ln(10)
        
    def _add_chart(self, df, title, chart_type='bar'):
        """Add chart to the report"""
        plt.figure(figsize=(10, 6))
        
        if chart_type == 'bar':
            plt.bar(df.index, df.values)
        elif chart_type == 'pie':
            plt.pie(df.values, labels=df.index, autopct='%1.1f%%')
            
        plt.title(title)
        plt.xticks(rotation=45)
        
        # Save chart to temporary file
        chart_file = 'temp_chart.png'
        plt.savefig(chart_file, bbox_inches='tight')
        plt.close()
        
        # Add chart to PDF
        self.pdf.image(chart_file, x=10, w=190)
        self.pdf.ln(140)  # Space for the chart
        
        # Remove temporary file
        os.remove(chart_file)
        
    def generate_report(self, df, analysis):
        """Generate the PDF report"""
        self._add_header()
        
        # Top 10 Cryptocurrencies
        self._add_section_header('Top 10 Cryptocurrencies by Market Cap')
        top_10 = df.head(10)[['Cryptocurrency Name', 'Symbol', 'Current Price (USD)', 'Market Capitalization']]
        self._add_table(top_10, 'Top 10 Cryptocurrencies')
        
        # Market Cap Distribution
        self._add_section_header('Market Cap Distribution (Top 5)')
        market_cap_data = df.head(5).set_index('Cryptocurrency Name')['Market Capitalization']
        self._add_chart(market_cap_data, 'Market Cap Distribution (Top 5)')
        
        # Price Changes
        self._add_section_header('24-hour Price Changes')
        price_changes = df.nlargest(5, 'Price Change (24h %)')[['Cryptocurrency Name', 'Price Change (24h %)']]
        self._add_table(price_changes, 'Largest Price Changes')
        
        # Trading Volume
        self._add_section_header('Trading Volume Analysis')
        volume_data = df.nlargest(5, '24-hour Trading Volume').set_index('Cryptocurrency Name')['24-hour Trading Volume']
        self._add_chart(volume_data, 'Top 5 by Trading Volume')
        
        # Key Statistics
        self._add_section_header('Key Statistics')
        stats_text = [
            f"Total Market Cap: ${df['Market Capitalization'].sum():,.2f}",
            f"Average Price: ${df['Current Price (USD)'].mean():,.2f}",
            f"Average 24h Volume: ${df['24-hour Trading Volume'].mean():,.2f}",
            f"Average Price Change: {df['Price Change (24h %)'].mean():.2f}%"
        ]
        
        self.pdf.set_font('Arial', '', 12)
        for stat in stats_text:
            self.pdf.cell(0, 10, stat, ln=True)
            
        # Save the report
        report_file = f'crypto_analysis_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        self.pdf.output(report_file)
        return report_file
