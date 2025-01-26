import pdfplumber
import pandas as pd
from dataclasses import dataclass, field
from typing import Optional, List
from pathlib import Path
from datetime import datetime
from calendar import month_abbr

@dataclass
class HistoryPoint:
    year: int
    month_id: int
    month_name: str
    balance: Optional[float] = None
    scheduled_payment: Optional[float] = None
    actual_payment: Optional[float] = None
    credit_limit: Optional[float] = None
    amount_past_due: Optional[float] = None

@dataclass
class Tradeline:
    # Summary section
    account_name: str
    account_number: str
    reported_balance: Optional[float]
    account_status: str
    available_credit: Optional[float]
    # Account History section
    history_points: List[HistoryPoint] = field(default_factory=list)
    
    def to_dict(self):
        return {
            'SUMMARY': '',  # Empty value for header
            'Account Name': self.account_name,
            'Account Number': self.account_number,
            'Reported Balance': self.reported_balance,
            'Account Status': self.account_status,
            'Available Credit': self.available_credit,
            '': '',  # Empty row for spacing
            'ACCOUNT HISTORY': ''  # New section header
        }

def extract_tradeline(pdf_path: str) -> Tradeline:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        
        lines = text.split('\n')
        account_name = ""
        account_number = ""
        reported_balance = None
        account_status = ""
        available_credit = None
        history_data = {}
        
        # Get account name from the first line that contains a bank name
        for line in lines:
            if any(x in line.upper() for x in ["BANK", "FCU", "CREDIT UNION", "FINANCIAL"]):
                account_name = line.strip()
                break
        
        # Track if we're in the payment history section
        in_payment_history = False
        current_year = None
        months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", 
                 "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
        
        # First pass to find the most recent year
        most_recent_year = None
        for line in lines:
            if line.strip() and line[0:4].isdigit():
                year = int(line[0:4])
                if most_recent_year is None or year > most_recent_year:
                    most_recent_year = year

        # Extract account details
        for line in lines:
            # Get account number and balance
            if "Account Number" in line and "Reported Balance" in line:
                acc_parts = line.split("Reported Balance")[0].split()
                account_number = ' '.join(acc_parts[-2:])
                try:
                    balance_str = line.split("$")[1].strip()
                    reported_balance = float(balance_str.replace(',', ''))
                except (IndexError, ValueError):
                    print("Error parsing balance")
            
            # Get status and available credit
            if "Account Status" in line and "Available Credit" in line:
                status_parts = line.split("Available Credit")
                account_status = status_parts[0].replace("Account Status", "").strip()
                try:
                    credit_str = status_parts[1].replace("$", "").strip()
                    if credit_str and credit_str != "NONE":
                        available_credit = float(credit_str.replace(',', ''))
                except (IndexError, ValueError):
                    print("Error parsing credit")
            
            # Parse payment history
            if line.strip() and line[0:4].isdigit():
                year = int(line[0:4])
                values = line[4:].split()
                
                # Process each month's value
                for i, value in enumerate(values):
                    if value.strip() and value != "-" and value.replace(',', '').isdigit():
                        month_index = i
                        
                        # Calculate month_id (36 to 1)
                        months_from_recent = ((most_recent_year - year) * 12) + (11 - month_index)
                        month_id = 36 - months_from_recent
                        
                        if month_id > 0:  # Only include last 36 months
                            history_data[(year, month_id)] = {
                                'year': year,
                                'month_id': month_id,
                                'month_name': months[month_index],
                                'balance': float(value.replace(',', ''))
                            }

        # Create HistoryPoint objects for all 36 months
        history_points = []
        for month_id in range(36, 0, -1):  # Changed from 24 to 36
            months_ago = 36 - month_id  # Changed from 24 to 36
            year = most_recent_year - (months_ago // 12)
            month_index = (11 - (months_ago % 12))
            month_name = months[month_index]
            
            # Get data if it exists, otherwise None
            data = next((data for (y, m), data in history_data.items() 
                        if m == month_id), None)
            
            history_points.append(HistoryPoint(
                year=year,
                month_id=month_id,
                month_name=month_name,
                balance=data['balance'] if data else None,
                scheduled_payment=None,
                actual_payment=None,
                credit_limit=None,
                amount_past_due=None
            ))

        return Tradeline(account_name=account_name,
                        account_number=account_number,
                        reported_balance=reported_balance,
                        account_status=account_status,
                        available_credit=available_credit,
                        history_points=history_points)

def save_to_file(tradeline: Tradeline, output_file: str):
    with open(output_file, 'w') as f:
        for key, value in tradeline.to_dict().items():
            f.write(f"{key}: {value}\n")

def save_to_excel(tradeline: Tradeline, output_file: str):
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        pd.DataFrame().to_excel(writer, sheet_name='Sheet1', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Formats
        header_fmt = workbook.add_format({'bold': True, 'font_size': 12})
        money_fmt = workbook.add_format({'num_format': '$#,##0'})
        year_fmt = workbook.add_format({'num_format': '0000'})  # Force 4-digit year
        
        # Write Summary section
        current_row = 0
        summary_data = {
            'SUMMARY': '',
            'Account Name': tradeline.account_name,
            'Account Number': tradeline.account_number,
            'Reported Balance': tradeline.reported_balance,
            'Account Status': tradeline.account_status,
            'Available Credit': tradeline.available_credit,
            '': '',
            'ACCOUNT HISTORY': ''
        }
        
        for key, value in summary_data.items():
            if key in ['SUMMARY', 'ACCOUNT HISTORY']:
                worksheet.write(current_row, 0, key, header_fmt)
            elif key in ['Reported Balance', 'Available Credit'] and value is not None:
                worksheet.write(current_row, 0, key)
                worksheet.write(current_row, 1, value, money_fmt)
            else:
                worksheet.write(current_row, 0, key)
                worksheet.write(current_row, 1, value)
            current_row += 1
        
        # Write Account History section
        current_row += 1
        
        # Headers in column A
        worksheet.write(current_row, 0, 'Year', header_fmt)
        worksheet.write(current_row + 1, 0, 'Month ID', header_fmt)
        worksheet.write(current_row + 2, 0, 'Month', header_fmt)
        worksheet.write(current_row + 3, 0, 'Balance', header_fmt)
        
        # Sort history points by year (descending) and month_id
        sorted_points = sorted(
            tradeline.history_points, 
            key=lambda x: (-x.year, -x.month_id)  # Reverse sort for both
        )
        
        # Write history data
        for col, point in enumerate(sorted_points, start=1):
            worksheet.write(current_row, col, point.year, year_fmt)  # Use year format
            worksheet.write(current_row + 1, col, point.month_id)
            worksheet.write(current_row + 2, col, point.month_name)
            if point.balance is not None:
                worksheet.write(current_row + 3, col, point.balance, money_fmt)
            if point.scheduled_payment is not None:
                worksheet.write(current_row + 4, col, point.scheduled_payment, money_fmt)
            if point.actual_payment is not None:
                worksheet.write(current_row + 5, col, point.actual_payment, money_fmt)
            if point.credit_limit is not None:
                worksheet.write(current_row + 6, col, point.credit_limit, money_fmt)
            if point.amount_past_due is not None:
                worksheet.write(current_row + 7, col, point.amount_past_due, money_fmt)
        
        # Set column widths
        worksheet.set_column('A:A', 25)  # Keep first column wide for labels
        worksheet.set_column('B:B', 15)  # Keep second column wide for summary values
        worksheet.set_column('C:Z', 5)   # Set narrow width for month columns

def setup_directories():
    input_dir = Path('input')
    output_dir = Path('output')
    
    print(f"Creating directories:")
    if not input_dir.exists():
        input_dir.mkdir()
        print(f"Created input directory at: {input_dir.absolute()}")
    else:
        print(f"Input directory exists at: {input_dir.absolute()}")
        
    if not output_dir.exists():
        output_dir.mkdir()
        print(f"Created output directory at: {output_dir.absolute()}")
    else:
        print(f"Output directory exists at: {output_dir.absolute()}")
    
    return input_dir, output_dir

def main():
    input_dir, output_dir = setup_directories()
    
    # Get first PDF from input directory
    pdf_files = list(input_dir.glob('*.pdf'))
    if not pdf_files:
        print(f"\nError: Please put your PDF file in: {input_dir.absolute()}")
        raise FileNotFoundError("No PDF files found in input directory")
    
    pdf_path = pdf_files[0]
    tradeline = extract_tradeline(str(pdf_path))
    
    # Save to output directory
    save_to_file(tradeline, output_dir / 'tradeline_data.txt')
    save_to_excel(tradeline, output_dir / 'tradeline_data.xlsx')

if __name__ == '__main__':
    main() 