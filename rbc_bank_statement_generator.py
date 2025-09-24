#script to generate bank statements for RBC accounts
import pandas as pd
import sys
import random
import os
import tempfile
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
TRANSACTIONS_DIR = "transactions/"
TEMPLATE_PDF = "template/rbc_banktemplate_V3_printed.pdf"

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
        opening_balance_y_coord= 367,
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

def create_first_page(canvas, opening_balance, account_info, beginning_date, closing_date, total_deposits, total_withdrawls, company_name, street_address, city, province, postal_code):
    #create the first page of the bank statement
    template = templates["rbc"]
    month_year = beginning_date.strftime("%B %Y")
    overlay_pdf = f"first_page_{month_year}.pdf"
    c = canvas
    width, height = letter
    c.setFont("Helvetica", 10)
    c.drawRightString(594, height - 110, f"From {beginning_date.strftime('%B %d, %Y')} to {closing_date.strftime('%B %d, %Y')}")
    c.setFont("Helvetica", 12)
    c.drawString(template.address_x_coord, height - template.company_name_y_coord, company_name)
    c.drawString(template.address_x_coord, height - template.street_address_y_coord, street_address)
    c.drawString(template.address_x_coord, height - template.city_province_state_country_y_coord, city + province)
    c.drawString(template.postal_code_x_coord, height - template.postal_code_y_coord, postal_code)
    c.setFont("Helvetica", 8)
    c.drawRightString(template.account_number_ending_x_coord, height - template.account_number_y_coord, account_info)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(template.second_account_number_x_coord, height - template.second_account_number_y_coord, account_info)
    c.setFont("Helvetica", 9)
    c.drawString(146, height - template.opening_balance_y_coord - 0.5, beginning_date.strftime("%b %d"))
    c.drawRightString(template.opening_balance_ending_x_coord, height - template.opening_balance_y_coord, f"${opening_balance:,.2f}")
    c.drawRightString(template.total_deposits_x_coord, height - template.total_deposits_y_coord, f"+ {total_deposits:,.2f}")
    c.drawRightString(template.total_withdrawals_x_coord, height - template.total_withdrawals_y_coord, f"-{total_withdrawls:,.2f}")
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(template.closing_balance_x_coord, height - template.closing_balance_ending_y_coord - 2, f"= ${opening_balance - total_withdrawls + total_deposits:,.2f}")
    c.drawString(150, height - template.closing_balance_ending_y_coord, closing_date.strftime("%b %d"))
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

def make_transaction_table(transactions, colWidths, current_page_number, total_pages):
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph
    
    styles = getSampleStyleSheet()
    # Create a style for wrapping text in the second column
    wrap_style = styles['Normal']
    wrap_style.fontSize = 8  # Adjust font size as needed
    wrap_style.fontName = 'Helvetica' 
    wrap_style.leading = 10   # Line spacing
    
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
    # table.setStyle(TableStyle([
    #     ('GRID', (0, 0), (-1, -1), 0, colors.white),
    #     ('LINEBELOW', (0, 0), (-1, -1), 0, colors.white),
    #     ('LINEABOVE', (0, 0), (-1, -1), 0, colors.white),
    #     ('LINEBEFORE', (0, 0), (-1, -1), 0, colors.white),
    #     ('LINEAFTER', (0, 0), (-1, -1), 0, colors.white),
    #     # Add dashed bottom line for each row
    #     ('LINEBELOW', (0, 0), (-1, -1), 1, colors.black, None, (1, 2)),
    #     ('LEFTPADDING', (0, 0), (-1, -1), 6),
    #     ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    #     ('TOPPADDING', (0, 0), (-1, -1), 6),
    #     ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    #     # Right align columns 3, 4, and 5 (Credit, Debit, Balance)
    #     ('ALIGN', (0, 0), (1, -1), 'LEFT'),  # Date column annd Payee column left aligned
    #     ('ALIGN', (2, 0), (4, -1), 'RIGHT'),
    #     # Set vertical alignment for all cells to TOP
    #     ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    #     ('FONTSIZE', (0, 0), (0, -1), 8),
    # ]))
    
    table.setStyle(TableStyle([
        # Remove all grid lines first
        ('GRID', (0, 0), (-1, -1), 0, colors.white),
        
        # Add dotted horizontal lines between rows
        ('LINEBELOW', (0, 0), (-1, -2), 0.5, colors.black, None, (1, 1)),  # Dotted pattern
        
        # Minimal padding
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        
        # Font settings
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        
        # Alignment
        ('ALIGN', (0, 0), (1, -1), 'LEFT'),   # Description columns
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'), # Amount/balance columns
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        # Final row (closing balance) bold
    ]))

    if current_page_number == total_pages:
        table.setStyle(TableStyle([
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ]))
    return table

def calculate_total_pages(data):
    """Calculate total number of pages needed for the statement"""
    total_data_transactions = len(data)
    
    # Account for: opening balance + data transactions + closing balance
    total_rows = 1 + total_data_transactions + 1  # opening + data + closing
    
    # First page can hold 12 rows total
    if total_rows <= 12:
        return 1
    
    # Remaining rows after first page (which holds 12)
    remaining_rows = total_rows - 12
    
    # Each additional page holds 22 rows
    additional_pages = (remaining_rows + 21) // 22  # Ceiling division
    
    return 1 + additional_pages

def draw_page_number(canvas, current_page, total_pages):
    """Draw page number at bottom right of page"""
    canvas.setFont("Helvetica-Bold", 9)
    page_text = f"{current_page} of {total_pages}"
    # Position at bottom right (adjust coordinates as needed)
    canvas.drawRightString(584, 25, page_text)


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

    line_height = 10
    colWidths = [90-47,  320-90, 60, 90, 122]
    balance = opening_balance

    # Month abbreviations mapping
    month_abbr = {
        1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
        7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
    }

    # Reset statement-specific variables
    is_first_page_of_statement = True  # Reset for new statement
    row_counter = 1
    transactions = [
        ["", "Opening Balance", "", "", f"{opening_balance:,.2f}"]
    ]
    last_date = None

    # Create temporary file for overlay PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
        overlay_pdf = temp_file.name

    c = canvas.Canvas(overlay_pdf, pagesize=letter)
    create_first_page(c, opening_balance, account_info, beginning_date, closing_date, total_deposits, total_withdrawls, company_name, street_address, city, province, postal_code)

    total_pages = calculate_total_pages(data)
    current_page_number = 1

    for _, row in data.iterrows():
        balance = balance - row['Credit'] + row['Debit']
        
        # Format date as "Day Mon" (e.g., "15 Feb")
        current_date = row['Date']
        date_str = f"{current_date.day} {month_abbr[current_date.month]}"
        
        # Only show date if it's different from the last date
        if last_date is None or current_date.date() != last_date.date():
            display_date = date_str
            last_date = current_date
        else:
            display_date = ""
        
        credit_str = f"{row['Credit']:,.2f}" if row['Credit'] != 0 else ""
        debit_str = f"{row['Debit']:,.2f}" if row['Debit'] != 0 else ""
        balance_str = f"{balance:,.2f}"
        transaction = [display_date, row['Payee'], credit_str, debit_str, balance_str]
        transactions.append(transaction)
        row_counter += 1

        if (is_first_page_of_statement and row_counter == 12) or (not is_first_page_of_statement and row_counter == 22):
            # Only draw if there is at least one data row
            if len(transactions) > 0 and all(len(r) == 5 for r in transactions):
                table = make_transaction_table(transactions, colWidths, current_page_number, total_pages)
                # Calculate table dimensions before drawing
                available_width = letter[0]   # leave margins
                available_height = letter[1] # leave space for header/footer
                table_width, table_height = table.wrapOn(c,available_width, available_height)
                y_pos = letter[1] - 498 - table_height if is_first_page_of_statement else letter[1] - 190 - table_height
                x_pos = 45 if is_first_page_of_statement else 18
                table.drawOn(c, x_pos, y_pos)
                print('Drew table with {} rows, on y postiion: {}, is first page:{}'.format(len(transactions), y_pos, is_first_page_of_statement))
                print("Height, width of page: {}, {}".format(letter[1], letter[0]))
                if not is_first_page_of_statement:
                    c.setFont("Helvetica", 10)
                    c.drawRightString(563, letter[1] - 90, f"From {beginning_date.strftime('%B %d, %Y')} to {closing_date.strftime('%B %d, %Y')}")
            
            draw_page_number(c, current_page_number, total_pages)
            if not is_first_page_of_statement:
                c.showPage()
                current_page_number += 1
            is_first_page_of_statement = False
            row_counter = 0  # reset for new page
            transactions = []  # start new page with empty transactions
            last_date = None  # reset date tracking for new page

    # Add closing balance row
    transactions.append(["", "<b>Closing Balance</b>", "", "", f"{closing_balance:,.2f}"])
    row_counter += 1

    # Draw any remaining transactions
    if len(transactions) > 0 and all(len(r) == 5 for r in transactions):
        table = make_transaction_table(transactions, colWidths, current_page_number, total_pages)
        # Bold the last row (Closing Balance)
        table.setStyle(TableStyle([
            ( 'FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ]))
        # Calculate table dimensions before drawing
        available_width = letter[0]   # leave margins
        available_height = letter[1]  # leave space for header/footer
        table_width, table_height = table.wrapOn(c, available_width, available_height)
        y_pos = letter[1] - 498 - table_height if is_first_page_of_statement else letter[1] - 190 - table_height
        if not is_first_page_of_statement:
            c.showPage()
            c.setFont("Helvetica", 10)
            c.drawRightString(563, letter[1] - 90, f"From {beginning_date.strftime('%B %d, %Y')} to {closing_date.strftime('%B %d, %Y')}")
            current_page_number += 1
        x_pos = 45 if is_first_page_of_statement else 18
        table.drawOn(c, x_pos, y_pos)
        draw_page_number(c, current_page_number, total_pages)
        print('Remaining: Drew table with {} rows, on y postiion: {}, is first page:{}'.format(len(transactions), y_pos, is_first_page_of_statement))
        print("Height, width of page: {}, {}".format(letter[1], letter[0]))


    c.save()    
    
    try:
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
            
    finally:
        # Clean up temporary file
        try:
            os.unlink(overlay_pdf)
        except OSError:
            pass  # File might already be deleted
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