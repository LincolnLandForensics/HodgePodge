 #!/usr/bin/env python3
# coding: utf-8
'''
read pdf and convert it to pdf
'''
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import re
import os
import sys
from datetime import datetime
import argparse
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import pdfplumber
import pandas as pd

# Colorize section
global color_red, color_yellow, color_green, color_blue, color_purple, color_reset
color_red = ''
color_yellow = ''
color_green = ''
color_blue = ''
color_purple = ''
color_reset = ''

if sys.version_info > (3, 7, 9) and os.name == "nt":
    version_info = os.sys.getwindowsversion()
    major_version = version_info.major
    build_version = version_info.build

    if major_version >= 10 and build_version >= 22000:
        import colorama # pip install colorama
        from colorama import Fore, Back, Style  
        print(f'{Back.BLACK}')
        color_red = Fore.RED
        color_yellow = Fore.YELLOW
        color_green = Fore.GREEN
        color_blue = Fore.BLUE
        color_purple = Fore.MAGENTA
        color_reset = Style.RESET_ALL
        
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>
author = 'LincolnLandForensics'
description = "read pdf and convert it to pdf"
version = '0.1.6'



def main():
    # global row
    # row = 0
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-b', '--blank', help='blank sheet', required=False, action='store_true')
    parser.add_argument('-p', '--pdf', help='parse pdf', required=False, action='store_true')
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-t', '--tables', help='output tables', required=False, action='store_true')

    args = parser.parse_args()

    global tables_out
    # Assuming args.tables is a list or other iterable
    tables_out = False  # Default value

    if args.tables:  # If args.tables is not empty or None
        tables_out = True

    global output_xlsx
    if not args.output: 
        output_xlsx = f"invoices_.xlsx"        
    else:
        output_xlsx = args.output    

    if not args.input: 
        directory_path = "pdfs"        
    else:
        directory_path = args.input

    if args.blank:
        data = []
        write_xlsx(data)
        
    elif args.pdf:
        # Directory containing PDF files
        # directory_path = 'pdfs'
        # Path where the output Excel file will be saved
        output_excel_path = 'output_pdfs.xlsx'


        if not os.path.exists(directory_path):
            print(f"Error: The directory '{directory_path}' does not exist.")
            # return
        else:
            msg_blurb = (f"Reading pdf's in {directory_path}")
            msg_blurb_square(msg_blurb, color_green)             
            
        process_pdfs_in_directory(directory_path, output_excel_path)

        if tables_out:    
            msg_blurb = (f'Tables have been extracted and saved to {output_excel_path}')
            msg_blurb_square(msg_blurb, color_green) 

    else:
        usage()

    # return 0

# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')

def extract_tables_from_pdf(pdf_path):
    """
    Extract tables from a given PDF file.
    
    Parameters:
    pdf_path (str): Path to the PDF file.
    
    Returns:
    list of pd.DataFrame: A list of DataFrames, one for each table found in the PDF.
    """
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df)
    return tables

def extract_invoice_number(tables):
    """
    Extract the invoice number from the tables.
    
    Parameters:
    tables (list of pd.DataFrame): List of DataFrames representing tables.
    
    Returns:
    str: The invoice number if found, otherwise None.
    """
    for table in tables:
        if 'Invoice #' in table.columns:
            invoice_number = table.loc[0, 'Invoice #']
            return invoice_number
    return None

def extract_information(tables):
    """
    Extract specific information from the tables.
    
    Parameters:
    tables (list of pd.DataFrame): List of DataFrames representing tables.
    
    Returns:
    dict: Dictionary containing the extracted values.
    """
    data_tables = {
        "Date": None,
        "Invoice #": None,
        "Bill To:": None,
        "Ship To:": None,
        "Quantity": None,
        "Description": None,
        "Price Each": None,
        "Amount": None,
        "Total": None
    }

    for table in tables:
        for key in data_tables.keys():
            if key in table.columns:
                data_tables[key] = table.loc[0, key]
                
        if all(data_tables.values()):
            break

    return data_tables
    
    
def extract_data(tables):
    """
    Extract the invoice number and date from the tables.
    
    Parameters:
    tables (list of pd.DataFrame): List of DataFrames representing tables.
    
    Returns:
    tuple: The invoice number and date if found, otherwise (None, None).
    """
    invoice_number = None
    date = None
    for table in tables:
        if 'Invoice #' in table.columns:
            invoice_number = table.loc[0, 'Invoice #']
        if 'Date' in table.columns:
            date = table.loc[0, 'Date']
        if invoice_number and date:
            break
    return invoice_number, date
    
    
def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')


def process_pdfs_in_directory(directory_path, output_excel_path):
    """
    Loop through all PDFs in the specified directory, extract tables, and export them to an Excel sheet.
    
    Parameters:
    directory_path (str): Path to the directory containing PDF files.
    output_excel_path (str): Path where the output Excel file will be saved.
    """
    if tables_out:
        writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')
        
    all_tables = []
    data = []
    
    for filename in os.listdir(directory_path):
        if filename.lower().endswith('.pdf'):
            (date, invoice, billTo, shipTo, quantity, description) = ('', '', '', '', '', '')
            (priceEach, subtotal, amount, total, billToState, shipToState) = ('', '', '', '', '', '')
            (customer, state, year, month, day) = ('', '', '', '', '')

            row_data = {}
            pdf_path = os.path.join(directory_path, filename)
            tables = extract_tables_from_pdf(pdf_path)
            all_tables.extend(tables)
            
            # Extract and print the invoice number
            invoice_number = extract_invoice_number(tables)
            if invoice_number:
                invoice = invoice_number

            # if 'Date' in table.columns:
                # date = table.loc[0, 'Date']

            # Extract and print the information
            data_tables = extract_information(tables)
            for key, value in data_tables.items():
                if value:
                    # print(f"key {key} filename {filename} = value = {value}")
                    if key == "Amount":
                        amount = value
                    elif key == "Date":
                        date = value
                    elif key == "Invoice #":
                        invoice = value
                    elif key == "Bill To:":
                        billTo = value
                    elif key == "Ship To:":
                        shipTo = value
                    elif key == "Quantity":
                        quantity = value
                    elif key == "Description":
                        description = value
                    elif key == "Price Each":
                        priceEach = value
                    elif key == "Total":
                        total = value

# billToState
            # Regex pattern to match the state abbreviation between ", " and the zipcode
            pattern1 = r',\s([A-Z]{2})\s\d{5}'

            # Search for the pattern in the billTo string
            match1 = re.search(pattern1, billTo)

            # Extract the state abbreviation if a match is found
            if match1:
                billToState = match1.group(1)
            else:
                billToState = ''

# shipToState
            # Regex pattern to match the state abbreviation between ", " and the zipcode
            pattern1 = r',\s([A-Z]{2})\s\d{5}'

            # Search for the pattern in the shipTo string
            match2 = re.search(pattern1, shipTo)
    
            # Extract the state abbreviation if a match is found
            if match2:
                shipToState = match2.group(1)
            else:
                shipToState = ''

# customer
            # Regex pattern to match the first line of the string
            pattern3 = r'^(.*)$'

            # Search for the pattern in the shipTo string
            match3 = re.search(pattern3, shipTo, re.MULTILINE)

            # Extract the first line if a match is found
            if match3:
                customer = match3.group(1)
            else:
                customer = ''

# total
            # calculate "total" by adding up all the amounts in amount

            # Regex pattern to match dollar amounts
            pattern4 = r'\$([0-9]+\.[0-9]{2})'

            # Find all matches of the pattern in the amount string
            matches = re.findall(pattern4, amount)

            # Convert matched strings to floats and calculate the total
            total = sum(float(match) for match in matches)

# date
            # Assuming the input date is a string
            date2 = date

            # Convert string to datetime object
            input_date = datetime.strptime(date, '%m/%d/%Y')

            # Format datetime object to desired output format
            date = input_date.strftime('%Y-%m-%d')
            year = input_date.strftime('%Y')
            month = input_date.strftime('%m')
            day = input_date.strftime('%d')

            row_data["filename"] = filename

            row_data["date"] = date
            row_data["invoice"] = invoice
            row_data["billTo"] = billTo
            row_data["shipTo"] = shipTo
            row_data["quantity"] = quantity            
            row_data["description"] = description
            row_data["priceEach"] = priceEach
            row_data["subtotal"] = subtotal
            row_data["amount"] = amount
            row_data["total"] = total
            row_data["billToState"] = billToState
            row_data["shipToState"] = shipToState
            row_data["customer"] = customer
            row_data["state"] = state
            row_data["year"] = year
            row_data["month"] = month
            row_data["day"] = day

            data.append(row_data)

    
    if tables_out:
        for idx, table in enumerate(all_tables):
            sheet_name = f'Table{idx+1}'
            table.to_excel(writer, sheet_name=sheet_name, index=False)

        writer.save()
        writer.close()

    write_xlsx(data)
    
    
def write_xlsx(data):

    global workbook
    workbook = Workbook()
    global worksheet
    worksheet = workbook.active

    worksheet.title = 'Invoices'
    header_format = {'bold': True, 'border': True}
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'B2'  # Freeze cells
    worksheet.selection = 'B2'

    headers = [
        "filename", "date", "invoice", "customer", "total", "billTo", "shipTo", "quantity", "description"
        , "priceEach", "amount", "billToState", "shipToState", "year", "month", "day", "subtotal"
    ]

    
    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in range(18):  # range(18) generates numbers from 0 to 25 inclusive orange
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # orange
            cell.fill = fill

    # Excel column width
    worksheet.column_dimensions['A'].width = 27 # 
    worksheet.column_dimensions['B'].width = 9 # 
    worksheet.column_dimensions['C'].width = 11 # 
    worksheet.column_dimensions['D'].width = 30 # 
    # worksheet.column_dimensions['E'].width = 6 # 
    worksheet.column_dimensions['E'].width = 13 # 
    worksheet.column_dimensions['F'].width = 13 #   
    worksheet.column_dimensions['G'].width = 10 # 
    worksheet.column_dimensions['H'].width = 9 # 
    worksheet.column_dimensions['I'].width = 11 # 
    worksheet.column_dimensions['J'].width = 18 # 
    worksheet.column_dimensions['K'].width = 11  # 
    # worksheet.column_dimensions['M'].width = 11  # 
    worksheet.column_dimensions['L'].width = 12  # 
    worksheet.column_dimensions['M'].width = 12  # 
    worksheet.column_dimensions['N'].width = 12  # 
    worksheet.column_dimensions['O'].width = 12  # 
    worksheet.column_dimensions['P'].width = 12  # 



    for row_index, row_data in enumerate(data):
        for col_index, col_name in enumerate(headers):
            cell_data = row_data.get(col_name)
            try:
                if isinstance(cell_data, list):
                    cell_data = str(cell_data)  # Convert lists to strings
                worksheet.cell(row=row_index+2, column=col_index+1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error printing line: {str(e)}{color_reset}")


    msg_blurb = (f'Writing to {output_xlsx}')
    msg_blurb_square(msg_blurb, color_green)

    workbook.save(output_xlsx)

def usage():
    '''
    working examples of syntax
    '''
    file = sys.argv[0].split('\\')[-1]
    print(f'\nDescription: {color_green}{description}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f'\n    {color_yellow}insert your input into intel.xlsx')
    print(f'\nExample:')
    print(f'    {file} -p') 
    print(f'    {file} -p -t')     
    print(f'    {file} -p -I pdfs -O invoices_.xlsx ')     


if __name__ == "__main__":

    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

0.1.5 - working prototype
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
can be adapated to different table column names
-t will optionaly print out the tables in different sheet
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>



