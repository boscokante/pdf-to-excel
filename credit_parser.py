import pdfplumber
import pandas as pd
from dataclasses import dataclass, field
from typing import Optional, List
from pathlib import Path
from datetime import datetime
from calendar import month_abbr
import logging
import re

# Configure logging at the start of the file
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('debug.log'),
        logging.StreamHandler()
    ]
)

@dataclass
class MonthlyData:
    year: int
    month_name: str
    value: Optional[float] = None

@dataclass
class AccountHistorySection:
    name: str
    monthly_data: List[MonthlyData] = field(default_factory=list)
    years: List[int] = field(default_factory=list)

@dataclass
class Tradeline:
    # Section 1: Header Information
    date_reported: str
    cra: str
    
    # Section 2: Account Information
    account_name: str
    account_number: str
    reported_balance: Optional[float]
    account_status: str
    available_credit: Optional[float]
    high_credit: Optional[float]
    payment_responsibility: str
    
    # Section 3: Account Details
    credit_limit: Optional[float]
    account_type: str
    terms_frequency: str
    term_duration: Optional[str]
    balance: Optional[float]
    date_opened: str
    amount_past_due: Optional[float]
    date_reported_details: str
    actual_payment_amount: Optional[float]
    date_of_last_payment: str
    date_of_last_activity: Optional[str]
    scheduled_payment_amount: Optional[float]
    
    # Section 4: Historical Information
    months_reviewed: int
    delinquency_first_reported: Optional[str]
    activity_designator: str
    
    # Section 5: Creditor Information
    creditor_classification: str
    deferred_payment_start_date: Optional[str]
    charge_off_amount: Optional[float]
    balloon_payment_date: Optional[str]
    balloon_payment_amount: Optional[float]
    loan_type: str
    
    # Section 6: Account Status
    date_closed: str
    date_of_first_delinquency: str
    
    # Section 7: Comments
    comments: str
    
    # Section 8: Creditor Contact
    contact_name: str
    contact_address: str
    contact_city_state_zip: str
    contact_phone: str

    def to_dict(self):
        return {
            'Date Reported': self.date_reported,
            'CRA': self.cra,
            'Account Name': self.account_name,
            'Account Number': self.account_number,
            'Reported Balance': self.reported_balance,
            'Account Status': self.account_status,
            'Available Credit': self.available_credit,
            'High Credit': self.high_credit,
            'Payment Responsibility': self.payment_responsibility,
            'Credit Limit': self.credit_limit,
            'Account Type': self.account_type,
            'Terms/Frequency': self.terms_frequency,
            'Term Duration': self.term_duration,
            'Balance': self.balance,
            'Date Opened': self.date_opened,
            'Amount Past Due': self.amount_past_due,
            'Date Reported (Details)': self.date_reported_details,
            'Actual Payment Amount': self.actual_payment_amount,
            'Date of Last Payment': self.date_of_last_payment,
            'Date of Last Activity': self.date_of_last_activity,
            'Scheduled Payment Amount': self.scheduled_payment_amount,
            'Months Reviewed': self.months_reviewed,
            'Delinquency First Reported': self.delinquency_first_reported,
            'Activity Designator': self.activity_designator,
            'Creditor Classification': self.creditor_classification,
            'Deferred Payment Start Date': self.deferred_payment_start_date,
            'Charge-off Amount': self.charge_off_amount,
            'Balloon Payment Date': self.balloon_payment_date,
            'Balloon Payment Amount': self.balloon_payment_amount,
            'Loan Type': self.loan_type,
            'Date Closed': self.date_closed,
            'Date of First Delinquency': self.date_of_first_delinquency,
            'Comments': self.comments,
            'Contact Name': self.contact_name,
            'Contact Address': self.contact_address,
            'Contact City/State/ZIP': self.contact_city_state_zip,
            'Contact Phone': self.contact_phone
        }

def extract_tradeline(pdf_path: str) -> Tradeline:
    debug_output = []
    
    with pdfplumber.open(pdf_path) as pdf:
        # Add PDF content debugging
        print("\n=== RAW PDF TEXT ===")
        full_text = ""
        for page_num, page in enumerate(pdf.pages):
            page_text = page.extract_text()
            print(f"\nPage {page_num + 1}:\n{page_text}")
            full_text += page_text + "\n"
        print("=== END PDF TEXT ===\n")
        
        lines = full_text.split('\n')
        
        # Initialize all fields with default values
        data = {
            'date_reported': 'N/A',
            'cra': 'Equifax',
            'account_name': 'N/A',
            'account_number': 'N/A',
            'reported_balance': None,
            'account_status': 'N/A',
            'available_credit': None,
            'high_credit': None,
            'payment_responsibility': 'N/A',
            'credit_limit': None,
            'account_type': 'N/A',
            'terms_frequency': 'N/A',
            'term_duration': None,
            'balance': None,
            'date_opened': 'N/A',
            'amount_past_due': None,
            'date_reported_details': 'N/A',
            'actual_payment_amount': None,
            'date_of_last_payment': 'N/A',
            'date_of_last_activity': None,
            'scheduled_payment_amount': None,
            'months_reviewed': 0,
            'delinquency_first_reported': None,
            'activity_designator': 'N/A',
            'creditor_classification': 'N/A',
            'deferred_payment_start_date': None,
            'charge_off_amount': None,
            'balloon_payment_date': None,
            'balloon_payment_amount': None,
            'loan_type': 'N/A',
            'date_closed': 'N/A',
            'date_of_first_delinquency': 'N/A',
            'comments': 'N/A',
            'contact_name': 'N/A',
            'contact_address': 'N/A',
            'contact_city_state_zip': 'N/A',
            'contact_phone': 'N/A'
        }

        current_section = None
        contact_lines = []
        in_balance_history = False

        for line in lines:
            line = line.replace('\x0c', ' ')  # Keep form feed handling
            debug_output.append(f"RAW LINE: {line}")
            
            # Clean line while preserving field separators
            clean_line = ' '.join(line.strip().split())
            debug_output.append(f"CLEAN LINE: {clean_line}")

            try:
                # 1. Parse Account Number and Balance
                if 'Account Number' in clean_line and 'Reported Balance' in clean_line:
                    # Format: "Account Number xxxxxxxx 5205 Reported Balance $949"
                    account_info = {}
                    parts = re.split(r'\s{2,}', clean_line)  # Split on 2+ spaces
                    
                    for i, part in enumerate(parts):
                        if 'Account Number' in part:
                            account_info['account_number'] = ' '.join(parts[i+1:i+3]).strip()
                        elif 'Reported Balance' in part:
                            account_info['reported_balance'] = parse_currency(parts[i+1])
                        elif 'Account Status' in part:
                            account_info['account_status'] = parts[i+1]
                        elif 'Available Credit' in part:
                            account_info['available_credit'] = parse_currency(parts[i+1])
                    
                    data.update(account_info)
                    logging.info(f"Account Info: {account_info}")

                # 2. Parse Credit Limit with exact pattern matching
                if 'Credit Limit' in clean_line:
                    # Match format: "Credit Limit $500 Account Type REVOLVING"
                    if '$' in clean_line:
                        credit_part = clean_line.split('$', 1)[1]
                        limit_value = credit_part.split(' ', 1)[0].replace(',', '')
                        try:
                            data['credit_limit'] = float(limit_value)
                            logging.info(f"CREDIT LIMIT FOUND: {data['credit_limit']}")
                        except ValueError:
                            logging.error(f"Failed to parse credit limit from: {clean_line}")
                    else:
                        logging.debug(f"Skipping credit limit header line: {clean_line}")

                    # Parse account type if present
                    if 'Account Type' in clean_line:
                        data['account_type'] = clean_line.split('Account Type ', 1)[1].split(' ', 1)[0]

                # 3. Parse Date Opened
                if 'Date Opened' in clean_line:
                    # Format: "Date Opened Jan 05, 2017"
                    date_str = clean_line.split('Date Opened ', 1)[1].strip()
                    data['date_opened'] = datetime.strptime(date_str, '%b %d, %Y').strftime('%Y-%m-%d')

                # 4. Parse Contact Info
                if 'PO Box' in clean_line:
                    data['contact_address'] = clean_line.strip()
                if 'Wilmington, DE' in clean_line:
                    data['contact_city_state_zip'] = clean_line.strip()
                if '(888)' in clean_line:
                    data['contact_phone'] = clean_line.strip()

            except Exception as e:
                logging.error(f"Error parsing line: {clean_line}")
                logging.error(f"Error details: {str(e)}")
                continue

        # Process collected data
        if contact_lines:
            data['contact_name'] = contact_lines[0] if contact_lines else 'N/A'
            data['contact_address'] = contact_lines[1] if len(contact_lines) > 1 else 'N/A'
            if len(contact_lines) > 2:
                data['contact_city_state_zip'] = ' '.join(contact_lines[2:])

        # Convert currency fields
        currency_fields = [
            'reported_balance', 'available_credit', 'high_credit', 'credit_limit',
            'balance', 'amount_past_due', 'actual_payment_amount', 
            'scheduled_payment_amount', 'charge_off_amount', 'balloon_payment_amount'
        ]
        for field in currency_fields:
            if data[field]:
                data[field] = parse_currency(data[field])

        with open('debug.log', 'w') as f:
            f.write('\n'.join(debug_output))

        return Tradeline(**data)

def parse_section(line: str, fields: list) -> dict:
    """Parse a line with multiple fields using known field markers"""
    section_data = {}
    remaining = line
    
    for field in fields:
        if field in remaining:
            parts = remaining.split(field, 1)
            value_part = parts[1].split('  ')[0].strip()
            field_name = field.replace(':', '').lower().replace('/', '_')
            section_data[field_name] = value_part
            remaining = parts[1][len(value_part):]
    
    return section_data

def parse_currency(value: str) -> Optional[float]:
    """Improved currency parsing with error handling"""
    try:
        return float(value.replace('$', '').replace(',', '').strip())
    except (ValueError, TypeError, AttributeError):
        return None

def save_to_file(tradeline: Tradeline, output_file: str):
    with open(output_file, 'w') as f:
        for key, value in tradeline.to_dict().items():
            f.write(f"{key}: {value}\n")

def save_to_excel(tradeline: Tradeline, output_file: str):
    df = pd.DataFrame({
        'Field': list(tradeline.to_dict().keys()),
        'Value': list(tradeline.to_dict().values())
    })

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, header=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Add formats
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
        money_fmt = workbook.add_format({'num_format': '$#,##0'})
        date_fmt = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        
        # Apply formatting
        for row_num, (field, value) in enumerate(zip(df['Field'], df['Value'])):
            worksheet.write(row_num, 0, field, header_fmt)
            
            if 'date' in field.lower():
                worksheet.write(row_num, 1, value, date_fmt)
            elif isinstance(value, (int, float)):
                worksheet.write(row_num, 1, value, money_fmt)
            else:
                worksheet.write(row_num, 1, value)
        
        # Set column widths
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 40)

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
