#script to generate bank statements for RBC accounts
import pandas as pd
import sys
import random
from datetime import datetime, timedelta
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import LETTER, letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from PyPDF2 import PdfReader, PdfWriter
from copy import deepcopy
import traceback

# Configuration
INPUT_FILE = "template/BrightDesk_Consulting_Ledger_Mar2022_to_Aug2025_v11.xlsx"
OUTPUT_DIR = "statements/"
TEMPLATE_PDF = "template/rbc_banktemplate_V1_printed.pdf"

#company details
COMPANY_NAME = "BrightDesk Consulting"
STREET_ADDRESS = "22 WELLINGTON ST E"
CITY_PROVINCE= "TORONTO ON M3C 2Z4"
POSTAL_CODE = "M3C 2Z4"
ACCOUNT_NUMBER = "5213 03XX XXXX 1234"

class TemplateConfig:
    def __init__(self, **kwargs):
        for key, value in kwargs.items():
            setattr(self, key, value)


templates = {
    "rbc": TemplateConfig(
        file= TEMPLATE_PDF,
        address_x_coord= 80,
        company_name_y_coord= 144,
        street_address_y_coord= 156,
        city_province_state_country_y_coord= 168,
        postal_code_x_coord= 80,
        postal_code_y_coord= 180,
        account_number_ending_x_coord= 590,
        account_number_y_coord= 150,
        second_account_number_x_coord= 158,
        second_account_number_y_coord= 319,
        opening_balance_ending_x_coord= 370,
        opening_balance_y_coord= 319,
        closing_balance_x_coord= 370,
        closing_balance_ending_y_coord= 420,
        total_deposits_x_coord= 370,     
        total_deposits_y_coord= 384,
        total_withdrawals_x_coord= 370,
        total_withdrawals_y_coord= 401,
    )
}    


def load_data(file_path):
    """Load transaction data from Excel file."""
    print(f"Loading data from {file_path}...")
    return pd.read_excel(file_path, sheet_name='chequing')

def create_first_page(opening_balance, account_info, beginning_date, closing_date, total_deposits, total_withdrawls, company_name, street_address, city, province, postal_code):
    #create the first page of the bank statement
    template = templates["rbc"]
    month_year = beginning_date.strftime("%B %Y")
    overlay_pdf = f"first_page_{month_year}.pdf"
    c = canvas.Canvas(overlay_pdf, pagesize=letter)
    width, height = letter
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(template.address_x_coord, template.company_name_y_coord, company_name)
    c.drawRightString(template.address_x_coord, template.street_address_y_coord, street_address)
    c.drawRightString(template.address_x_coord, template.city_province_state_country_y_coord, city + "," + province)
    c.drawRightString(template.postal_code_x_coord, template.postal_code_y_coord, postal_code)

    return overlay_pdf

def generate_monthly_statements(data):
    """Generate monthly bank statements based on transaction data.
    Args:
        data: Raw transaction data from Excel
    """
    # Convert Date column to datetime if not already
    data['Date'] = pd.to_datetime(data['Date'])
    
    # Get the range of months that have data
    min_date = data['Date'].min()
    max_date = data['Date'].max()
    
    print(f"Data range: {min_date.strftime('%B %d, %Y')} to {max_date.strftime('%B %d, %Y')}")
    
    # Determine the first statement month
    first_statement_month = min_date.month
    first_statement_year = min_date.year
    
    # Determine the last statement month
    last_statement_month = max_date.month
    last_statement_year = max_date.year
    
    print(f"Generating statements from {first_statement_month}/{first_statement_year} to {last_statement_month}/{last_statement_year}")
    
    # Generate statements for each month
    current_month = first_statement_month
    current_year = first_statement_year
    
    while (current_year < last_statement_year) or (current_year == last_statement_year and current_month <= last_statement_month):
        # Generate statement for this month
        data_for_month = get_transactions(data, current_month, current_year)
        
        if len(data_for_month) > 0:
            print(f"\nGenerating statement for {datetime(current_year, current_month, 1).strftime('%B %Y')}")
            generate_statement(data_for_month, current_month, current_year)
        else:
            print(f"\nNo transactions found for {datetime(current_year, current_month, 1).strftime('%B %Y')}")
        
        # Move to next month
        if current_month == 12:
            current_month = 1
            current_year += 1
        else:
            current_month += 1


def get_transactions(data, month, year):
    #Genearte transactions for the satement period
    monthly_data = data[(data['Date'].dt.month == month) & (data['Date'].dt.year == year)]
    return monthly_data

def get_company_info(file_path):
    df = pd.read_excel(file_path, sheet_name='company_info', header=None)
    df.columns = ['Field', 'Value']
    info = dict(zip(df['Field'], df['Value']))
    return {
        "company_name": info["Company Name"],
        "street_address": info["Street Address"],
        "city_province": info['City, Province'],
        "postal_code": info['Postal Code'],
        "account_number": info['Account Number']
    }

def make_transaction_table(transactions, colWidths):
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    
    styles = getSampleStyleSheet()
    # Create a style for wrapping text in the second column
    wrap_style = styles['Normal']
    wrap_style.fontSize = 10  # Adjust font size as needed
    wrap_style.leading = 12   # Line spacing
    
    # Process transactions to wrap text in second column (Payee)
    processed_transactions = []
    for row in transactions:
        processed_row = []
        for i, cell in enumerate(row):
            if i == 1 and isinstance(cell, str) and cell:  # Second column (Payee)
                # Wrap the text in a Paragraph for text wrapping
                processed_row.append(Paragraph(str(cell), wrap_style))
            else:
                processed_row.append(cell)
        processed_transactions.append(processed_row)
    
    table = Table(processed_transactions, colWidths=colWidths)
    table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0, colors.white),
        ('LINEBELOW', (0, 0), (-1, -1), 0, colors.white),
        ('LINEABOVE', (0, 0), (-1, -1), 0, colors.white),
        ('LINEBEFORE', (0, 0), (-1, -1), 0, colors.white),
        ('LINEAFTER', (0, 0), (-1, -1), 0, colors.white),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black, None, (3, 3)),
        ('VALIGN', (1, 0), (1, 0), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        # Right align columns 3, 4, and 5 (Credit, Debit, Balance)
        ('ALIGN', (2, 0), (4, -1), 'RIGHT'),
        # Set vertical alignment for all cells to TOP
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    return table

def generate_statement(data, month, year):
    #create first page of the statement
    opening_balance = data["Beginning Balance"].iloc[0]
    closing_balance = data["Closing Balance"].iloc[-1]
    total_deposits = data["Debit"].sum()
    total_withdrawls = data["Credit"].sum()
    company_info = get_company_info(INPUT_FILE)
    beginning_date = data['Date'].iloc[0]
    closing_date = data['Date'].iloc[-1]
    company_name = company_info["company_name"] 
    street_address = company_info["street_address"]
    city_province = company_info["city_province"]
    # Split city and province more robustly
    city, province = city_province.rsplit(" ", 1)
    account_info = company_info["account_number"]
    postal_code = company_info["postal_code"]

    create_first_page(opening_balance, account_info, beginning_date, closing_date, total_deposits, total_withdrawls, company_name, street_address, city, province, postal_code)
    line_height = 10
    colWidths = [95-47,  320-95, 60, 90, 122]
    balance = opening_balance

    # Start with header and opening balance row
    header = ["Date", "Payee", "Credit", "Debit", "Balance"]
    transactions = [
        header,
        ["", "Opening Balance", "", "", opening_balance]
    ]

    overlay_pdf = f"transactions_{year}_{month:02d}_.pdf"
    c = canvas.Canvas(overlay_pdf, pagesize=letter)
    row_counter = 2  # header + opening balance
    first_page = True

    for _, row in data.iterrows():
        balance = balance - row['Credit'] + row['Debit']
        transaction = [row['Date'].strftime("%m-%d"), row['Payee'], row['Credit'], row['Debit'], balance]
        transactions.append(transaction)
        row_counter += 1

        if (first_page and row_counter == 12) or (not first_page and row_counter == 22):
            # Only draw if there is at least one data row (header + 1)
            if len(transactions) > 2 and all(len(r) == 5 for r in transactions):
                table = make_transaction_table(transactions, colWidths)
                # Calculate table dimensions before drawing
                available_width = letter[0] - 90  # leave margins
                available_height = letter[1] - 200  # leave space for header/footer
                table.wrap(available_width, available_height)
                y_pos = letter[1] - 738 if first_page else letter[1] - 179
                table.drawOn(c, 45, y_pos)
            first_page = False
            row_counter = 1  # reset for new page, header will be added
            transactions = [header]  # start new page with header

    # Draw any remaining transactions (must have header + at least 1 data row)
    if len(transactions) > 1 and all(len(r) == 5 for r in transactions):
        table = make_transaction_table(transactions, colWidths)
        # Calculate table dimensions before drawing
        available_width = letter[0] - 90  # leave margins
        available_height = letter[1] - 200  # leave space for header/footer
        table.wrap(available_width, available_height)
        y_pos = letter[1] - 738 if first_page else letter[1] - 179
        if not first_page:
            c.showPage()
        table.drawOn(c, 45, y_pos)

    c.save()    
    # Merge overlay with template
    template_pdf = PdfReader(TEMPLATE_PDF)
    overlay_pdf_reader = PdfReader(overlay_pdf)
    writer = PdfWriter()
    for page_num in range(len(overlay_pdf_reader.pages)):
        template_page = template_pdf.pages[0] if page_num == 0 else template_pdf.pages[1]
        overlay_page = overlay_pdf_reader.pages[page_num]
        template_page.merge_page(overlay_page)
        writer.add_page(template_page)

    output_pdf = f"{OUTPUT_DIR}RBC_Statement_{year}_{month:02d}.pdf"
    with open(output_pdf, "wb") as f:
        writer.write(f)

def main():
    """Main function."""
    try:
        data = load_data(INPUT_FILE)
        generate_monthly_statements(data)
    except Exception as e:
        traceback.print_exc()  # <-- This prints the full traceback
        return 1
    return 0

if __name__ == "__main__":
    sys.exit(main())