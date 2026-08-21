 #!/usr/bin/env python3
# coding: utf-8
'''
read pdf and convert it to pdf
'''
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import re
import os
import sys
from PyPDF2 import PdfReader, PdfWriter

from datetime import datetime
import argparse
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import pdfplumber
import pandas as pd

import tkinter as tk
from tkinter import ttk, filedialog
import tkinter.scrolledtext as scrolledtext
import threading
import queue

# Colorize section
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
from colorama import Fore, Back, Style
print(Back.BLACK)
color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>
author = 'LincolnLandForensics'
description = "read pdfs and convert their tables to a single .xlsx"
version = '1.0.1'


# <<<<<<<<<<<<<<<<<<<<<<<<<<       Global        >>>>>>>>>>>>>>>>>>>>>>>>>>
global ssort
ssort = 0

def run_gui():
    root = tk.Tk()
    
    scriptName = os.path.basename(sys.argv[0])
    root.title(f"{scriptName} {version}")
    root.geometry("780x620")
    root.configure(bg="#f4f6f7")
    
    # Vista style setup
    style = ttk.Style(root)
    try:
        style.theme_use('vista')
    except Exception:
        pass

    # Header frame for premium look
    header_frame = tk.Frame(root, bg="#2c3e50", height=65)
    header_frame.pack(fill='x', side='top')
    
    lbl_title = tk.Label(header_frame, text=f"{scriptName} v{version}", fg="white", bg="#2c3e50", font=("Segoe UI", 13, "bold"))
    lbl_title.pack(pady=(12, 2))
    
    # Description label
    lbl_desc = tk.Label(root, text=description, font=("Segoe UI", 10, "italic"), bg="#f4f6f7", fg="#566573")
    lbl_desc.pack(pady=6)
    
    # Main container
    main_frame = ttk.Frame(root, padding="15 10 15 15")
    main_frame.pack(fill='both', expand=True)

    # Radiobuttons Frame
    mode_frame = ttk.LabelFrame(main_frame, text=" Select Operation ", padding="12 10 12 10")
    mode_frame.pack(fill='x', pady=5)
    
    mode_var = tk.StringVar(value='Invoice_pdfs')
    
    def on_mode_change():
        if mode_var.get() == 'split':
            ent_input.delete(0, tk.END)
            ent_input.insert(0, r'.\pdfs')
            ent_output.delete(0, tk.END)
            ent_output.insert(0, r'.\pdfs_single_page')
        else:
            ent_input.delete(0, tk.END)
            ent_input.insert(0, r'.\pdfs_single_page')
            ent_output.delete(0, tk.END)
            ent_output.insert(0, r'Invoices_V1.xlsx')
            
    rb_split = ttk.Radiobutton(mode_frame, text="Split PDFs (split)", variable=mode_var, value='split', command=on_mode_change)
    rb_split.pack(side='left', padx=30)
    
    rb_invoice = ttk.Radiobutton(mode_frame, text="Parse Invoices (Invoice_pdfs)", variable=mode_var, value='Invoice_pdfs', command=on_mode_change)
    rb_invoice.pack(side='left', padx=30)

    # Paths Frame
    path_frame = ttk.LabelFrame(main_frame, text=" File / Folder Paths ", padding="12 10 12 10")
    path_frame.pack(fill='x', pady=5)
    
    # Row 1: Input Path
    ttk.Label(path_frame, text="Input Path:").grid(row=0, column=0, sticky='w', pady=6)
    ent_input = ttk.Entry(path_frame, width=60)
    ent_input.grid(row=0, column=1, padx=8, pady=6, sticky='ew')
    ent_input.insert(0, r'.\pdfs_single_page')
    
    def browse_input():
        # First try folder selection
        folder_selected = filedialog.askdirectory(title="Select Input PDF Folder")
        if folder_selected:
            ent_input.delete(0, tk.END)
            ent_input.insert(0, os.path.normpath(folder_selected))
        else:
            # Fallback to single file selection if folder wasn't chosen
            file_selected = filedialog.askopenfilename(title="Select Input PDF File", filetypes=[("PDF files", "*.pdf")])
            if file_selected:
                ent_input.delete(0, tk.END)
                ent_input.insert(0, os.path.normpath(file_selected))
                
    btn_browse_input = ttk.Button(path_frame, text="Browse...", command=browse_input)
    btn_browse_input.grid(row=0, column=2, padx=5, pady=6)

    # Row 2: Output Path
    ttk.Label(path_frame, text="Output Path:").grid(row=1, column=0, sticky='w', pady=6)
    ent_output = ttk.Entry(path_frame, width=60)
    ent_output.grid(row=1, column=1, padx=8, pady=6, sticky='ew')
    ent_output.insert(0, r'Invoices_V1.xlsx')
    
    def browse_output():
        if mode_var.get() == 'split':
            folder_selected = filedialog.askdirectory(title="Select Output Split Folder")
            if folder_selected:
                ent_output.delete(0, tk.END)
                ent_output.insert(0, os.path.normpath(folder_selected))
        else:
            file_selected = filedialog.asksaveasfilename(
                title="Save Output Excel File",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")]
            )
            if file_selected:
                ent_output.delete(0, tk.END)
                ent_output.insert(0, os.path.normpath(file_selected))
                
    btn_browse_output = ttk.Button(path_frame, text="Browse...", command=browse_output)
    btn_browse_output.grid(row=1, column=2, padx=5, pady=6)
    
    path_frame.columnconfigure(1, weight=1)

    # Control Frame (Start button, Progress bar, status label)
    control_frame = ttk.Frame(main_frame, padding="5 5 5 5")
    control_frame.pack(fill='x', pady=5)
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(control_frame, variable=progress_var, maximum=100)
    progress_bar.pack(fill='x', side='top', pady=6)
    
    lbl_status = ttk.Label(control_frame, text="Status: Ready", font=("Segoe UI", 9, "bold"), foreground="#2e4053")
    lbl_status.pack(side='left', padx=5)
    
    # Terminal-like log window
    term_frame = ttk.LabelFrame(main_frame, text=" GUI Terminal Output ", padding="5 5 5 5")
    term_frame.pack(fill='both', expand=True, pady=5)
    
    txt_log = scrolledtext.ScrolledText(term_frame, bg="#1e1e1e", fg="#d4d4d4", insertbackground="white", font=("Consolas", 9), height=11)
    txt_log.pack(fill='both', expand=True)
    txt_log.configure(state='disabled')
    
    def log_to_gui(msg):
        txt_log.configure(state='normal')
        txt_log.insert(tk.END, msg + "\n")
        txt_log.see(tk.END)
        txt_log.configure(state='disabled')
        
    q = queue.Queue()
    
    def worker():
        mode = mode_var.get()
        inp = ent_input.get()
        outp = ent_output.get()
        
        # reset global sort tracker
        global ssort
        ssort = 0
        
        def gui_log_cb(msg, type='info'):
            q.put(('log', msg, type))
            
        def gui_prog_cb(val):
            q.put(('progress', val))
            
        try:
            if mode == 'split':
                gui_log_cb(f"Running mode 'split'...", 'info')
                split_pdfs(inp, outp, progress_cb=gui_prog_cb, log_cb=gui_log_cb)
                q.put(('output_file', outp))
                q.put(('done', f"Split process finished."))
            else:
                gui_log_cb(f"Running mode 'Invoice_pdfs'...", 'info')
                global output_xlsx, sub_folders, output_split_folder
                output_xlsx = outp
                sub_folders = False
                output_split_folder = 'pdfs_single_page'
                
                data = process_invoice_pdfs(inp, outp, progress_cb=gui_prog_cb, log_cb=gui_log_cb)
                write_xlsx(data)
                q.put(('output_file', outp))
                q.put(('done', f"Invoice extraction process finished."))
        except Exception as ex:
            import traceback
            tb = traceback.format_exc()
            print(tb) # output to the original console
            q.put(('error', str(ex)))
            
    def poll_queue():
        while not q.empty():
            item = q.get()
            if item[0] == 'log':
                msg, log_type = item[1], item[2]
                if log_type == 'input_file':
                    log_to_gui(f"Input file: {msg}")
                else:
                    log_to_gui(msg)
            elif item[0] == 'progress':
                progress_var.set(item[1])
                lbl_status.config(text=f"Status: Processing ({item[1]}%)")
            elif item[0] == 'output_file':
                log_to_gui(f"Output file: {item[1]}")
            elif item[0] == 'done':
                lbl_status.config(text="Status: Completed")
                progress_var.set(100)
                log_to_gui(f"Success: {item[1]}")
                btn_start.config(state='normal')
                rb_split.config(state='normal')
                rb_invoice.config(state='normal')
                btn_browse_input.config(state='normal')
                btn_browse_output.config(state='normal')
            elif item[0] == 'error':
                lbl_status.config(text="Status: Error")
                log_to_gui(f"Error: {item[1]}")
                btn_start.config(state='normal')
                rb_split.config(state='normal')
                rb_invoice.config(state='normal')
                btn_browse_input.config(state='normal')
                btn_browse_output.config(state='normal')
        root.after(100, poll_queue)
        
    def start_processing():
        txt_log.configure(state='normal')
        txt_log.delete('1.0', tk.END)
        txt_log.configure(state='disabled')
        
        btn_start.config(state='disabled')
        rb_split.config(state='disabled')
        rb_invoice.config(state='disabled')
        btn_browse_input.config(state='disabled')
        btn_browse_output.config(state='disabled')
        
        progress_var.set(0)
        lbl_status.config(text="Status: Starting...")
        
        t = threading.Thread(target=worker, daemon=True)
        t.start()
        
    btn_start = ttk.Button(control_frame, text="Start", command=start_processing)
    btn_start.pack(side='right', padx=5)
    
    root.after(100, poll_queue)
    root.mainloop()

def main():
    if len(sys.argv) == 1:
        run_gui()
        return
    print(f'{ssort}')
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-b', '--blank', help='blank sheet', required=False, action='store_true')
    parser.add_argument('-N', '--Invoice_pdfs', help='parse invoices', required=False, action='store_true')

    parser.add_argument('-p', '--pdf', help='parse pdf', required=False, action='store_true')
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-t', '--tables', help='output tables', required=False, action='store_true')
    parser.add_argument('-s', '--subfolders', help='parse subfolders', required=False, action='store_true')
    parser.add_argument('-S', '--split', help='split pdfs', required=False, action='store_true')
    parser.add_argument('-d', '--table_dump_1', help='parse pdf type 2', required=False, action='store_true')

    args = parser.parse_args()

    global tables_out
    # Assuming args.tables is a list or other iterable
    tables_out = False  # Default value

    if args.tables:  # If args.tables is not empty or None
        tables_out = True

    global output_xlsx
    if not args.output: 
        output_xlsx = f"Invoices_V1.xlsx"        
    else:
        output_xlsx = args.output    
    
    global input_folder
    if not args.input: 
        input_folder = "pdfs"        
    else:
        input_folder = args.input

    global output_split_folder
    output_split_folder = 'pdfs_single_page'

    global sub_folders
    if args.subfolders:
        sub_folders = True
    else:
        sub_folders = False
    # print(f"PDF's in subfolders are going to be processed")

    if args.blank:
        data = []
        write_xlsx(data)

    elif args.split:
        split_pdfs(input_folder, output_split_folder)

    elif args.Invoice_pdfs:
        output_excel_path = 'Invoices_V1.xlsx'


        if not os.path.exists(input_folder):
            print(f"{color_red}Error: The directory '{input_folder}' does not exist.{color_reset}")
        else:
            msg_blurb = (f"Reading Bank PDF's in {input_folder}")
            msg_blurb_square(msg_blurb, color_green)             
        process_invoice_pdfs(input_folder, output_excel_path)
        data = process_invoice_pdfs(input_folder, output_excel_path)
        # print(f'{color_yellow} {ssort} lines processed{color_reset}')    
        write_xlsx(data)
        
    elif args.pdf:
        # Directory containing PDF files
        # input_folder = 'pdfs'
        # Path where the output Excel file will be saved
        output_excel_path = 'output_pdfs.xlsx'


        if not os.path.exists(input_folder):
            print(f"Error: The directory '{input_folder}' does not exist.")
            # return
        else:
            msg_blurb = (f"Reading pdf's in {input_folder}")
            msg_blurb_square(msg_blurb, color_green)             
        # process_pdfs_in_directory_old(input_folder, output_excel_path) 
        process_pdfs_in_directory(input_folder, output_excel_path)
        # process_pdfs_in_directory(input_folder, output_excel_path, tables_out=True, sub_folders=True)

        if tables_out:    
            msg_blurb = (f'Tables have been extracted and saved to {output_excel_path}')
            msg_blurb_square(msg_blurb, color_green) 


    elif args.table_dump_1:
        output_excel_path = 'output_pdfs.xlsx'

        if not os.path.exists(input_folder):
            print(f"Error: The directory '{input_folder}' does not exist.")
        else:
            msg_blurb = (f"Reading pdf's in {input_folder}")
            msg_blurb_square(msg_blurb, color_green)             
        # process_pdfs_in_directory_old(input_folder, output_excel_path) 
        table_dump1(input_folder, output_excel_path)
        # process_pdfs_in_directory(input_folder, output_excel_path, tables_out=True, sub_folders=True)

        if tables_out:    
            msg_blurb = (f'Tables have been extracted and saved to {output_excel_path}')
            msg_blurb_square(msg_blurb, color_green) 

    else:
        usage()

    # return 0

# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>


def convert_rfc822_date(Order_Date):
    # Example: "Sun, 08 Jun 2025 20:31:18 +0000"
    try:
        dt = datetime.strptime(Order_Date, "%a, %d %b %Y %H:%M:%S %z")
    except ValueError:
        return "", "", "", ""

    # Normalize to YYYY-MM-DD
    date_str = dt.strftime("%Y-%m-%d")
    year  = dt.strftime("%Y")
    month = dt.strftime("%m")
    day   = dt.strftime("%d")

    return date_str, year, month, day

def create_order_date(Order_Date, year, month, day):
    # Month lookup table
    months = {
        "Jan": "01", "Feb": "02", "Mar": "03",
        "Apr": "04", "May": "05", "Jun": "06",
        "Jul": "07", "Aug": "08", "Sep": "09",
        "Oct": "10", "Nov": "11", "Dec": "12"
    }

    # Regex: e.g., "Aug 20th 2025"
    pattern = r"([A-Za-z]{3})\s+(\d{1,2})(?:st|nd|rd|th)\s+(\d{4})"
    m = re.search(pattern, Order_Date)

    if not m:
        return "", "", "", ""

    month_abbrev = m.group(1)
    day_raw      = m.group(2)
    year_raw     = m.group(3)

    # Convert month abbreviation → number
    month = months.get(month_abbrev, "")
    if not month:
        return "", "", "", ""

    # Normalize day
    day = f"{int(day_raw):02d}"

    # Year is already 4 digits
    year = year_raw
    Order_Date = (f'{year}-{month}-{day}')
    return Order_Date, year, month, day



def create_date(Date, year, month, day):
    # Normalize year
    year = str(year)
    if len(year) == 2:
        year = "20" + year
    try:
        # Normalize month and day
        month = f"{int(month):02d}"
        day   = f"{int(day):02d}"
        # Build ISO date
        Date = f"{year}-{month}-{day}"
    except Exception as e:
        err_msg = (f"{color_red}[Error2]: {e}{color_reset}")
            


    return Date, year, month, day


def convert_date(date):
    # print(f'{color_green}convert_date{color_reset}')    # works
    date_str, year, month, day = '', '', '', ''
    try:
        # Check if the date has valid delimiters and length
        if len(date) < 5:
            print(f'{date} is not a valid date: too short.')
            return date_str, year, month, day
        
        if date.count('/') == 2:
            date_obj = datetime.strptime(date, '%m/%d/%y')
        elif date.count('-') == 2:
            date_obj = datetime.strptime(date, '%Y-%m-%d')
        else:
            print(f'{date} is not a valid date format: incorrect delimiters.')
            return date_str, year, month, day
        
        # Format the datetime object into the desired components
        date_str = date_obj.strftime('%Y-%m-%d')
        year = date_obj.strftime('%Y')
        month = date_obj.strftime('%m')
        day = date_obj.strftime('%d')
    
    except ValueError as e:
        print(f'{e} - Date error for input: {date}')
    
    return date_str, year, month, day

def convert_full_date(line):
    # Map month names to numbers ex 'May 31, 2024'
    months = {
        "January": "01", "February": "02", "March": "03",
        "April": "04", "May": "05", "June": "06",
        "July": "07", "August": "08", "September": "09",
        "October": "10", "November": "11", "December": "12"
    }

    # Regex: MonthName Day, Year
    pattern = r"([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})"
    m = re.search(pattern, line)

    if not m:
        return None, None, None

    month_name = m.group(1)
    day = m.group(2)
    year = m.group(3)

    # Convert month name → number
    month = months.get(month_name, None)
    if not month:
        return None, None, None

    # Pad day if needed
    day = day.zfill(2)

    return year, month, day

    
def extract_information(tables):
    # print(f'{color_green}extract_information{color_reset}')
    """
    Extract specific information from the tables.
    
    Parameters:
    tables (list of pd.DataFrame): List of DataFrames representing tables.
    
    Returns:
    dict: Dictionary containing the extracted values.
    """

    data_tables = {
        "Date": "",
        "Invoice #": None,
        "Bill To:": None,
        "Ship To:": None,
        "Quantity": None,
        "Qty": None,
        "Description": None,
        "Price Each": None,
        "Amount": None,
        "Unit Price": None,
        "Total": None,
        "Product": None,
        "Qty": None,
        "Column1": None,
        "Column2": None,
        "Column3": None,
        "Column4": None,
        "Column5": None,
        "Column6": None,
        "Column7": None,
        "Column8": None,
        "Column9": None,
        "Column10": None
    }

    for table in tables:
        for key in data_tables.keys():
            if key in table.columns:
                data_tables[key] = table.loc[0, key]
                
        if all(data_tables.values()):
            break
    # print(f'{color_green}extract_information{color_reset}{data_tables}')
    return data_tables
  
def extract_tables_from_pdf(pdf_path):
    """
    Extract tables from a given PDF file.
    
    Parameters:
    pdf_path (str): Path to the PDF file.
    
    Returns:
    list of pd.DataFrame: A list of DataFrames, one for each table found in the PDF.
    """
    tables = []
    text, total, date = '', '', ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:

            # extract tables
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df)

            # extract text
            text = page.extract_text()
            # print(f"Text on page {i+1}:\n{text}")
            # print(f"Text on page :\n{text}")  # temp


            
    # print(f'{color_green}extract_tables_from_pdf{text}{color_reset}')         # works      
    return tables, text

def extract_invoice_from_text(text, data, filename):
    # print(f'{color_purple}extract_bank_info_from_text{color_reset}')
    global ssort
    
    Toss_List = [{'IL 60101', '\'', '-', '•', '.', '!'}]
    
    DatePaid, Note = '', ''
    TaxID, TP = '', ''
    InventoryTemp, SubtotalTemp, SalesTaxRate, SalesTaxRateTemp = '', '', '', ''
    POTemp, Rep, Ship, QuantityItemTemp = '', '', '', ''
    DateInvoiceTemp, SalesTaxTemp, TotalTemp, BillToTemp, SellerTemp = '', '', '', '', ''
    Invoice, Code, POnumber, SalesTax, Terms = '', '', '', '', ''
    Seller, Illinois_License, Customer = '', '', ''
    TotalStoreCredit, TotalTaxCredited, PreviousStoreCredit = '', '', ''
    Date, Amount, Description, Reference_Number, Bankcard, Account = '', '', '', '', '', ''
    Type, Customer, check, Ledger_Balance = '', '', '', ''
    year, month, day, Period, bank = '', '', '', '', ''
    Begin_Balance, Deposits, Checks_and_deductions, End_Balance, temp, Checks = '', '', '', '', '', '' 
    Withdrawals_Deductions, Service_Fees, ATM, Other_Subtractions = '', '', '', ''
    date_pattern = r"\b\d{2}/\d{2}\b"   # If you want to match just MM/DD
    date_pattern2 = r"\b(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])/\d{2}\b"  # MM/DD/YY
    date_pattern3 = r"\b(0?[1-9]|1[0-2])/(0?[1-9]|[12]\d|3[01])/\d{4}\b"    # 2/23/2024

    invoice_pattern1 = r"^\d{12}\b"
    invoice_pattern2 = r"^\d{7}\b"    
    invoice_pattern3 = r"M^\d{7}\b" 

    date_pattern2 = r"\b(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])/\d{2,4}\b" # mm/dd/yy and mm/dd/yyyy
    
    
    

    match_account1 = re.search(r'Account No\.\:\s*(\d+)\s*Page\:', text)
    match_balance = r'^\s*\d[\d,]*\.\d{2}(?: \d[\d,]*\.\d{2}){3}\s*$'
    # amount_pattern = r"\$?\d[\d,]*\.\d{2}"
    amount_pattern = r"\$\d{1,3}(?:,\d{3})*\.\d{2}" # $4,100.00
    # amount_pattern2 = r"\d{1,3}(?:,\d{3})*\.\d{2}"  # 4,100.00
    amount_pattern2 = r"-?\d{1,3}(?:,\d{3})*\.\d{2}"  # 4,100.00 or   # -4,100.00
    pattern_amount2 = r"^-?\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$"    # 208.81
    
    
    
    Order_Date, Order = '', ''
    Total, Subtotal, Discount, Shipping, SalesTax, SalesTaxRate, GrandTotal = '', '', '', '', '', '', ''
    BillTo, Shipping_Details = '', ''
    Email, Phone, DBA = '', '', ''
    
    # Split the text into lines
    lines = text.split('\n')
    row_data = {}
    # Iterate over each line and process
    tally_line = 0

    for line in lines:
        SKU, Product, Price, Total = '', '', '', ''
        
        Amount, Description, Reference_Number, Bankcard = '', '', '', ''
        MSRP, check, Ledger_Balance = '', '', ''
        temp, month2, Withdrawals, Credits = '', '', '', ''
        First, Middle, Last, Second2last, Third2last, Fourth2last = '', '', '', '', '', ''
        sort, First, Second, Middle, Second2last, Last = '', '', '', '', '', ''
        Date2, year2, month2, day2, Qty = '', '', '', '', ''
        
        if '/ ' in line:    # ex 10/ 12 701.55 Corporate ACH Tax 512666
            line = line.replace('/ ', '/')
        
        matches6 = re.findall(date_pattern, line)
        matches7 = re.findall(date_pattern2, line)  # mm/dd/yy


        tally_line += 1
        ssort += 1
 
        # regex section
        creditPattern = r'^(\d{10,14})\s+(\d+)\s+(\d*)\s+\$\s*(\d+\.\d{2})\s+\$\s*(\d+\.\d{2})$'
        MatchCredit = re.search(creditPattern, line)

        try:
            linetemp = line.split(' ')
            First = linetemp[0]
            Second = linetemp[0]
            Middle = " ".join(linetemp[1:-2])
            # Product = " ".join(linetemp[2:-4])
            Last = linetemp[-1]
            Second2last  = linetemp[-2]
            Third2last  = linetemp[-3]
            Fourth2last  = linetemp[-4]
        except Exception as e:
            print(f"[ERROR] Failed to parse value: {line}")
            print(f"[EXCEPTION] {e}")

        if line == 'Invoice':
            SellerTemp = 'y'
            temp = (f'{temp} _Invoice1').strip()
        elif line == '':
            temp = (f'{temp} _Blank1').strip()
        elif SellerTemp == 'y':
            Seller = line
            SellerTemp = ''
            temp = (f'{temp} _SellerTemp1').strip()
        elif line == 'Date Invoice#' or line == 'Date Invoice #':
            DateInvoiceTemp = 'y'
            temp = (f'{temp} _DateInvoice1').strip()
        elif line == 'Customer Product Quantity' or line == 'Inventory Opening Close' or line == 'Received Quantity':
            InventoryTemp = 'y'
            temp = (f'{temp} __InventoryTemp1').strip()        
        elif InventoryTemp == 'y':
            temp = (f'{temp} __InventoryTemp2').strip()     

        elif line in Toss_List[0]:
            temp = (f'{temp} _toss').strip()

        elif DateInvoiceTemp == 'y' and 'IL 60101' not in line:
            linetemp = line.split(' ')
            m = re.search(date_pattern3, line)  # TP # 00142 4/2/2025 886

            try:
                if m:
                    Date = m.group(0)


                    Order_Date = Date
                    Invoice = linetemp[-1]   # last token = actual invoice number
                    if Date.count('/') == 2:
                        month, day, year = Date.split('/')
                    if "TP #" in line:
                        TP = linetemp[2]
            except Exception as e:
                err_msg = f"[Error1]: {e}"
  
            DateInvoiceTemp = ''
            temp = (f'{temp} DateInvoice2').strip()

        elif 'Bill To' in line or 'Ship To' in line:
            BillToTemp = 'y'
            temp = (f'{temp} _BillToTemp1').strip()
        elif BillToTemp == 'y':
            BillTo = line
            BillToTemp = ''
            temp = (f'{temp} _BillTo1').strip()

        elif 'Tax ID #' in line:
            TaxID = line.replace('Tax ID # ', '')
            temp = (f'{temp} TaxID1').strip()


        elif line == 'P.O. Number Terms Rep Ship Via Project':
            POTemp = 'y'
            temp = (f'{temp} _POTemp1').strip()
        elif line == 'P.O. No. Terms Project':
            # POTemp = 'y'
            temp = (f'{temp} _POTemp2').strip()            

        elif POTemp == 'y':
            linetemp = line.split(' ')
            try:
                POnumber = linetemp[0]
                Terms = linetemp[1]
                Rep = linetemp[2]
                Ship = linetemp[-1]
                
            except:pass    

            m = re.search(date_pattern3, line)

            try:
                if m:
                    Ship = m.group(0)
                m2 = re.search(date_pattern3, Rep)
                if m2:
                    Rep = ''
                    # if Date.count('/') == 2:
                        # month, day, year = Date.split('/')

            except Exception as e:
                err_msg = f"[Error3]: {e}"

            POTemp = ''
            temp = (f'{temp} POTemp').strip()



        elif 'Quantity Description Rate Amount' in line:
            QuantityItemTemp = 'y'
            temp = (f'{temp} _QuantityItemTemp1').strip()
            
        elif 'Quantity Item Code' in line:
            QuantityItemTemp = 'y'
            temp = (f'{temp} _QuantityItemTemp2').strip()
        elif QuantityItemTemp == 'y' and 'Sub' not in First:
            if 'State Tax ' in line:
                SalesTaxRate = Middle.replace('Tax ', '')
                SalesTax = Last
                temp = (f'{temp} SalesTaxRate1').strip()



            elif First == 'Total':
                GrandTotal = Last.replace('$', '')
                temp = (f'{temp} GrandTotal1').strip()                
            # elif re.search(r"[0-9.]", First):
            elif '. Total $' in line:
                linetemp = line.split('. Total $')
                try:
                    GrandTotal = linetemp[1]
                    Note = linetemp[0]
                except:pass
                temp = (f'{temp} GrandTotal2').strip() 



            elif 'All Customers are responsible for their State Sales Tax' in line:
                Note = 'All Customers are responsible for their State Sales Tax'
                temp = (f'{temp} Note1').strip()
            elif 'This Invoice included Illinois State Tobacco Tax' in line:
                Note = 'This Invoice included Illinois State Tobacco Tax'
                temp = (f'{temp} Note2').strip()
                
                
                
                
            elif 'CUSTOMER IS LIABLE FOR THEIR STATE TOBACCO TAXES' in line:
                Note = 'CUSTOMER IS LIABLE FOR THEIR STATE TOBACCO TAXES'
                temp = (f'{temp} Note3').strip()




            elif re.search(r"[0-9]", First):    
                Qty = First
                Description = Middle
                if (re.fullmatch(pattern_amount2, Second2last) and '.' in Second2last):
                    Price = Second2last.replace('$', '')
                Last_clean = Last.replace('T', '')
                if Price:
                    Amount = Last_clean.replace('$', '')

                temp = (f'{temp} QuantityItem1').strip()

            else:
                Description = line
                if (re.fullmatch(pattern_amount2, Second2last) and '.' in Second2last):
                    Price = Second2last.replace('$', '')
                Last_clean = Last.replace('T', '')
                if Price:
                    Amount = Last_clean.replace('$', '')

                temp = (f'{temp} QuantityItem2').strip()

        # SalesTaxRateTemp
        elif ' Tax (' in line:
            SalesTaxRateTemp = 'y'
            try:
                SalesTaxRate = line.split('Tax (')[1].replace(')', '')
            except:pass
            SalesTaxTemp = 'y'
            temp = (f'{temp} SalesTaxRateTemp1').strip()
        elif SalesTaxTemp == 'y':
            SalesTax = line
            SalesTaxTemp = ''
            temp = (f'{temp} SalesTax').strip()



        elif ' Tax (' in line:
            SalesTaxTemp = 'y'
            temp = (f'{temp} _SalesTaxTemp1').strip()
        elif SalesTaxTemp == 'y':
            SalesTax = line
            SalesTaxTemp = ''
            temp = (f'{temp} SalesTax').strip()


        elif 'Sub' in line and 'otal' in line:
            SubtotalTemp = 'y'
            QuantityItemTemp = ''
            temp = (f'{temp} _SubtotalTemp1').strip()
        elif SubtotalTemp == 'y':
            Subtotal = line
            SubtotalTemp = ''
            temp = (f'{temp} Subtotal').strip()





        elif line == 'Total':
            TotalTemp = 'y'
            QuantityItemTemp = ''
            temp = (f'{temp} _GrandTotalTemp1').strip()
        elif TotalTemp == 'y':
            GrandTotal = line

                
            if '$' in GrandTotal:
                GrandTotal = GrandTotal.split('$')[1].strip()
                temp = (f'{temp} GrandTotal').strip()
            elif 'Phone#' in GrandTotal:
                GrandTotal = 'missing$'                
                temp = (f'{temp} GrandTotal_missing').strip()
            TotalTemp = ''
            

        elif TotalTemp == 'y':
            GrandTotal = line
            TotalTemp = ''
            temp = (f'{temp} GrandTotal').strip()
        elif line == 'Phone#':
            GrandTotal, SalesTax, Terms, POnumber, Invoice, Order_Date, Date, SalesTaxRate = '', '', '', '', '', '', '', ''
            Rep = ''
            Subtotal = ''
            temp = (f'{temp} _phone').strip()


            
        elif 'invoice for order #' in line:
            linetemp = line.split('invoice for order #')
            Order = linetemp[1]
            Seller = linetemp[0]
        elif 'INVOICE # ' in line:
            linetemp = line.split('INVOICE # ')
            Order = linetemp[1]           
  
            
        elif 'ORDER SUBTOTAL : ' in line:
            linetemp = line.split('ORDER SUBTOTAL : ')
            Subtotal = linetemp[1].rstrip('00')
        elif 'Subtotal: ' in line:
            linetemp = line.split('Subtotal: ')
            Subtotal = linetemp[1].replace(' USD', '')  




            
        elif 'Discount: ' in line:
            linetemp = line.split('Discount: ')
            Discount = linetemp[1].replace(' USD', '') 
        elif 'Shipping: ' in line:
            linetemp = line.split('Shipping: ')
            Shipping = linetemp[1].replace(' USD', '')
        elif 'State Vapor Tax: ' in line:
            linetemp = line.split('State Vapor Tax: ')
            SalesTax = linetemp[1].replace(' USD', '')
        elif 'Grand total: ' in line:
            linetemp = line.split('Grand total: ')
            GrandTotal = linetemp[1].replace(' USD', '')
        elif 'Midwest Goods Inc. – Invoice Disclaimer' in line:
            Shipping = ''
            Total, Subtotal, Discount, Shipping, SalesTax, GrandTotal = '', '', '', '', '', ''
        elif 'Illinois License: ' in line:
            linetemp = line.split('Illinois_License: ')
            Illinois_License = line
            Illinois_License = Illinois_License.replace('Illinois_License: ', '')
        elif 'Order Date: ' in line:
            linetemp = line.split('Order Date: ')
            Order_Date = linetemp[1]
            if 'Order: #' in linetemp:
                linetemp2 = linetemp[0].replace('Order: #', '')
            Order_Date, year, month, day = create_order_date(Order_Date, year, month, day)
        elif 'ORDER DATE & TIME : ' in line:
            linetemp = line.split('ORDER DATE & TIME : ')
            Order_Date = linetemp[1]
            if Order_Date:
                Order_Date, year, month, day = convert_rfc822_date(Order_Date)




        elif '(MSRP $' in line:
            linetemp = line.split('(MSRP $')
            MSRP = linetemp[1].replace(')', '')
        elif ' DBA ' in line:
            DBA = line

        elif 'Customer Email : ' in line:
            Email = line.split('Customer Email : ')[1]
        elif 'Email: ' in line:
            Email = line.split('Email: ')[1]

        elif 'Customer Name : ' in line:
            Customer = line.split('Customer Name : ')[1]


        elif 'Customer Shipping Address : ' in line:
            Shipping_Details = line.split('Customer Shipping Address : ')[1]



        elif 'Phone: ' in line:
            Phone = line.split('Phone: ')[1]
            Phone = Phone.replace(' Do you Require Lift Gate Service?', '')

        elif 'Phone: ' in line:
            Phone = line.split('Phone: ')[1]
            Phone = Phone.replace(' Do you Require Lift Gate Service?', '')
            
        elif 1==3:    
        # elif line.count("USD") >= 2:
            parts = line.split()

            # Qty and SKU are always the first two fields
            Qty = parts[0]
            SKU = parts[1].rstrip('-')

            # Extract all USD amounts
            usd_amounts = re.findall(r"\$\d[\d,]*\.\d{2}", line)

            # Price = second-to-last USD amount
            # Total = last USD amount
            if len(usd_amounts) >= 2:
                Price = usd_amounts[-2]
                Total = usd_amounts[-1]
            else:
                Price = ""
                Total = ""

            # Product = everything between SKU and the first USD
            first_usd_index = line.index("USD")
            product_section = line[:first_usd_index].strip()

            # Remove Qty and SKU from product section
            product_words = product_section.split()[2:]
            Product = " ".join(product_words)
            Product = Product.replace('{Price}', '').replace('{Total}', '').strip()
            # Product = " ".join(linetemp[2:-4])

        # elif MatchCredit:   # 840395817704 1 1 $ 45.00 $ 45.00
            # SKU     = MatchCredit.group(1)
            # Qty     = MatchCredit.group(2)
            ##  Product = MatchCredit.group(3)
            # Price   = MatchCredit.group(4)
            # Credit  = MatchCredit.group(5)

            # print(SKU, Qty, Product, Price, Credit)

        elif 'Total Store Credit $ ' in line:
            TotalStoreCredit = line.split('Total Store Credit $ ')[1]
        elif 'Total Tax Credited $ ' in line:
            TotalTaxCredited = line.split('Total Tax Credited $ ')[1]
        elif 'Previous Store Credit $ ' in line:
            PreviousStoreCredit = line.split('Previous Store Credit $ ')[1]





        if Date == '':
            if re.search(date_pattern2, First):               
                m = re.search(date_pattern2, First)
                Date = m.group(0)
                print(m.group(0))            
                Date, year, month, day = convert_date(Date)

        # If year is blank, extract last 4 digits from the line
        if year == '':
            Period = Period.strip()
            tail = Period[-4:]                 # last 4 characters
            if tail.isdigit():               # ensure they are digits
                year = tail

        # if not year.strip().startswith('2'):
            # year = ''

        if year and month and day:
            Date, year, month, day = create_date(Date, year, month, day)


 
        print(f'{filename}  {color_blue}{line}{color_reset}')   # worked
        row_data = {}
        row_data["Date"] = Date  
        row_data["DatePaid"] = DatePaid
        row_data["Seller"] = Seller  
        row_data["Invoice"] = Invoice  
        row_data["Qty"] = Qty
        row_data["Code"] = Code
        row_data["Description"] = Description
        row_data["Price"] = Price
        row_data["Total"] = Total
        row_data["POnumber"] = POnumber
        row_data["Subtotal"] = Subtotal
        row_data["SalesTax"] = SalesTax
        row_data["SalesTaxRate"] = SalesTaxRate  
        row_data["BillTo"] = BillTo          
        row_data["Amount"] = Amount
        row_data["TaxID"] = TaxID
        row_data["TP"] = TP
        row_data["Note"] = Note
      
      




        
        row_data["GrandTotal"] = GrandTotal
        row_data["Terms"] = Terms
        row_data["Rep"] = Rep
        row_data["Ship"] = Ship        
        row_data["Customer"] = Customer
        row_data["year"] = year
        row_data["month"] = month
        row_data["day"] = day
        row_data["text"] = line
        row_data["temp"] = temp
        row_data["sort"] = ssort
        row_data["filename"] = filename
        row_data["First"] = First
        row_data["Second"] = Second
        row_data["Second"] = Second        
        row_data["Middle"] = Middle        
        row_data["Fourth2last"] = Fourth2last
        row_data["Third2last"] = Third2last
        row_data["Second2last"] = Second2last
        row_data["Last"] = Last
        row_data["Order"] = Order
        row_data["Shipping"] = Shipping
        row_data["BillTo"] = BillTo
        row_data["Shipping_Details"] = Shipping_Details
        row_data["Customer"] = Customer
        
        data.append(row_data)

    return data

def extract_info_from_text(text, total, date, grand_total, shipping, invoice):
    print(f'{color_green}extract_info_from_text{color_reset}')
    disclaimer = ''

    match_account1 = re.search(r'Account No\.\:\s*(\d+)\s*Page\:', text)

    match_time = re.search(r'Total:\s*\$\s*([\d,]+\.\d{2})', text)
    match_date = re.search(r'Date:\s*(\d{1,2}/\d{1,2}/\d{4})', text)
    match_total = re.search(r'Total\s*\$([\d,]+\.\d{2})', text)
    # match_total2 = re.search(r'Total:\s*\$([0-9]+\.[0-9]{2})', text)
    match_total2 = re.search(r'Total:\s*\$\s*([0-9,]+\.[0-9]{2})', text)

    match_shipping = re.search(r'Shipping charge 1 \$([0-9]+\.[0-9]{2})', text)
    match_invoice = re.search(r'Willow st, suite #B (\d+)', text)

    # todo one of the grand_total has less spaces when the amount is over $999

    if match_account1:
        account = match_account1.group(1)
        print(f'{color_green}&&&&&&&&&&&&&&&&&&&Extracted Account Number: {account}{color_reset}')

    if match_total:
        # Remove commas from the number and return it as a float
        grand_total1 = float(match_total.group(1).replace(',', ''))
        # print(f'text2 = {text}OOOOOOOOOOOOOOOOOOOOOOOOOOOOOO')  # temp
        if grand_total1 is None:
            grand_total1 = ''
        if grand_total == '':
            grand_total = grand_total1
        else:
            grand_total = (f'{grand_total}\t{grand_total1}')
        # print(f'grand_total3 = {grand_total}____________________________') # temp   # todo
    if match_total2:
        grand_total2 = float(match_total2.group(1).replace(',', ''))
        # print(f'grand_total2-1 = {grand_total2}____________________________') # temp   # todo

        if grand_total2 is None:
            grand_total2 = ''
        if grand_total == '':
            grand_total = grand_total2
        else:
            grand_total2 = (f'{grand_total}\t{grand_total2}')    # temp        
        # print(f'grand_total2 = {grand_total2}____________________________') # temp   # todo

    if match_date:
        if date is None or date == '':
            date = match_date.group(1)

            
    return disclaimer, total, date, grand_total, shipping, invoice


def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')

def process_invoice_pdfs(input_folder, output_excel_path, progress_cb=None, log_cb=None):
    # print(f'{color_purple}process_bank_pdfs{color_reset}')  # worked
 
    """
    Loop through all PDFs in the specified directory (and optionally in subdirectories), extract tables, and export them to an Excel sheet.
    
    Parameters:
    input_folder (str): Path to the directory containing PDF files.
    output_excel_path (str): Path where the output Excel file will be saved.
    tables_out (bool): Whether to save the extracted tables into separate sheets in the Excel file.
    sub_folders (bool): Whether to include PDFs from subdirectories.
    """
    # if tables_out:
    #     writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')
    global tally
    
    tally = 0 
    data = []
    # customer = input("CUSTOMER: ")
    customer = ''
    
    # Gather files to process
    pdf_files = []
    if os.path.isfile(input_folder):
        pdf_files.append((os.path.dirname(input_folder), [os.path.basename(input_folder)]))
    elif os.path.exists(input_folder):
        for root, dirs, files in os.walk(input_folder) if sub_folders else [(input_folder, [], os.listdir(input_folder))]:
            pdfs = [f for f in files if f.lower().endswith('.pdf')]
            if pdfs:
                pdf_files.append((root, pdfs))
    else:
        err_msg = f"Error: The input path '{input_folder}' does not exist."
        print(f"{color_red}{err_msg}{color_reset}")
        if log_cb:
            log_cb(err_msg)
        return data

    total_files = sum(len(files) for root, files in pdf_files)
    if total_files == 0:
        msg = "No PDF files found."
        print(msg)
        if log_cb:
            log_cb(msg)
        return data

    processed_count = 0
    for root, files in pdf_files:
        for filename in files:
            pdf_path = os.path.join(root, filename)
            if log_cb:
                log_cb(pdf_path, type='input_file')
            
            bank = ''
            total, amount, description = '', '', ''
            state, year, month, day, folder = '', '', '', '', ''
            grand_total, shipping, subtotal, discount, total_tax  = '', '', '', '', ''
            date1, amt, account, check, SEQ, TC  = '', '', '', '', '', ''
            date2, amt2, account2, check2, SEQ2, TC2 =  '', '', '', '', '', ''
            date3, amt3, account3, check3, SEQ3, TC3  = '', '', '', '', '', ''
            date4, amt4, account4, check4, SEQ4, TC4  = '', '', '', '', '', ''
            
            row_data = {}
            folder = pdf_path
            tables, text = extract_tables_from_pdf(pdf_path)
            data = extract_invoice_from_text(text, data, filename)
            
            tally += 1
            processed_count += 1
            if progress_cb:
                progress_cb(int(processed_count / total_files * 100))

    print(f"{color_green}processed {tally} PDF's{color_reset}")
    return data


def process_pdfs_in_directory(input_folder, output_excel_path):
    """
    Loop through all PDFs in the specified directory (and optionally in subdirectories), extract tables, and export them to an Excel sheet.
    
    Parameters:
    input_folder (str): Path to the directory containing PDF files.
    output_excel_path (str): Path where the output Excel file will be saved.
    tables_out (bool): Whether to save the extracted tables into separate sheets in the Excel file.
    sub_folders (bool): Whether to include PDFs from subdirectories.
    """
    if tables_out:
        writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')
        
    all_tables = []
    data = []
    seller = input("SELLER: ")
    # Walk through directories and subdirectories if sub_folders is True
    for root, dirs, files in os.walk(input_folder) if sub_folders else [(input_folder, [], os.listdir(input_folder))]:
        for filename in files:
            if filename.lower().endswith('.pdf'):
                date, invoice, billTo, shipTo, quantity, description = '', '', '', '', '', ''
                priceEach, subtotal, amount, total, billToState, shipToState = '', '', '', '', '', ''
                customer, state, year, month, day, folder = '', '', '', '', '', ''
                grand_total, shipping = '', ''
                TaxID, TP = '', ''

                row_data = {}
                pdf_path = os.path.join(root, filename)
                folder = pdf_path
                tables, text = extract_tables_from_pdf(pdf_path)
                # print(f'text = {text}') # temp
                disclaimer, total, date, grand_total, shipping, invoice = extract_info_from_text(text, total, date, grand_total, shipping, invoice)
                if total:
                    print(f"Found total value : ${total}")
                # else:
                    # print(f"total not found")                

                all_tables.extend(tables)

                data_tables = extract_information(tables)
                print(f'date2 = {date}') # temp
                for key, value in data_tables.items():
                    if value:
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
                            if quantity is None:
                                quantity == ''
                        elif key == "Qty" and quantity == '' :
                            quantity = value                            
                        elif key == "Qty":
                            quantity = value
                        elif key == "Description":
                            description = value
                            if description is None:
                                description == ''                            
                        elif key == "Product" and description == '':
                            description = value
                        elif key == "Price Each":
                            priceEach = value
                            if priceEach is None:
                                priceEach == ''
                        elif key == "Unit Price" and priceEach == '':
                            priceEach = value

                        elif key == "Total":
                            total = value


                pattern = r',\s([A-Z]{2})\s\d{5}'

                match1 = re.search(pattern, billTo)
                if match1:
                    billToState = match1.group(1)
                else:
                    billToState = ''

                match2 = re.search(pattern, shipTo)
                if match2:
                    shipToState = match2.group(1)
                else:
                    shipToState = ''

                match3 = re.search(r'^(.*)$', shipTo, re.MULTILINE)
                if match3:
                    Customer = match3.group(1)
                else:
                    Customer = ''

# total
                # calculate "total" by adding up all the amounts in amount
                pattern4 = r'\$([0-9,]+\.[0-9]{2})'
                match4 = re.findall(pattern4, amount)

                total = sum(float(value4.replace(',', '')) for value4 in match4)    
                total = round(float(total), 2)
                
                try:
                    if date.count('/') == 2:
                        print(f'date3 = {date}') # temp
                        input_date = datetime.strptime(date, '%m/%d/%Y')
                        date = input_date.strftime('%Y-%m-%d')
                        year = input_date.strftime('%Y')
                        month = input_date.strftime('%m')
                        day = input_date.strftime('%d')
                except Exception as e:
                    print(f"Error converting date: {str(e)}")
                print(f'date4 = {date}') # temp
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
                row_data["folder"] = folder
                row_data["seller"] = seller
                row_data["disclaimer"] = disclaimer
                row_data["grand_total"] = grand_total
                row_data["shipping"] = shipping
                row_data["TaxID"] = TaxID
                row_data["TP"] = TP




                print(f'{filename}    total ={total}')   # test
                data.append(row_data)

    if tables_out:
        for idx, table in enumerate(all_tables):
            sheet_name = f'Table{idx+1}'
            table.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()
        writer.close()

    write_xlsx(data)



def split_pdfs(input_folder, output_split_folder, progress_cb=None, log_cb=None):
    # from PyPDF2 import PdfReader, PdfWriter

    """
    Function to split PDFs into single-page PDFs.
    """
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_split_folder):
        os.makedirs(output_split_folder)
        print(f"Created split PDF folder: {output_split_folder}")
        if log_cb:
            log_cb(f"Created split PDF folder: {output_split_folder}")

    msg_blurb = (f'Splitting PDFs from {input_folder} folder')
    msg_blurb_square(msg_blurb, color_green)
    if log_cb:
        log_cb(msg_blurb)
    
    # Gather files to process
    pdf_files = []
    if os.path.isfile(input_folder):
        pdf_files.append((os.path.dirname(input_folder), [os.path.basename(input_folder)]))
    elif os.path.exists(input_folder):
        for root, dirs, files in os.walk(input_folder):
            pdfs = [f for f in files if f.lower().endswith('.pdf')]
            if pdfs:
                pdf_files.append((root, pdfs))
    else:
        err_msg = f"Error: The input path '{input_folder}' does not exist."
        print(f"{color_red}{err_msg}{color_reset}")
        if log_cb:
            log_cb(err_msg)
        return

    total_files = sum(len(files) for root, files in pdf_files)
    if total_files == 0:
        msg = "No PDF files found."
        print(msg)
        if log_cb:
            log_cb(msg)
        return

    processed_count = 0
    for root, files in pdf_files:
        for filename in files:
            pdf_path = os.path.join(root, filename)
            if log_cb:
                log_cb(pdf_path, type='input_file')
            try:
                with open(pdf_path, 'rb') as pdf_file:
                    reader = PdfReader(pdf_file)
                    
                    for page_num, page in enumerate(reader.pages, start=1):
                        writer = PdfWriter()
                        writer.add_page(page)
                        
                        # Construct output path maintaining subdirectory structure
                        if os.path.isfile(input_folder):
                            relative_path = ""
                        else:
                            relative_path = os.path.relpath(root, input_folder)
                        output_dir = os.path.join(output_split_folder, relative_path)
                        if not os.path.exists(output_dir):
                            os.makedirs(output_dir)
                        
                        output_filename = f"{os.path.splitext(filename)[0]}_page_{page_num:06d}.pdf"
                        output_path = os.path.join(output_dir, output_filename)
                        
                        with open(output_path, 'wb') as output_pdf:
                            writer.write(output_pdf)
                        
                        print(f"Created: {output_path}")
            except Exception as e:
                err_msg = f"Error processing file {pdf_path}: {e}"
                print(err_msg)
                if log_cb:
                    log_cb(err_msg)
            processed_count += 1
            if progress_cb:
                progress_cb(int(processed_count / total_files * 100))

    msg_blurb = (f'Finished writing split PDFs to {output_split_folder} folder')
    msg_blurb_square(msg_blurb, color_green)
    if log_cb:
        log_cb(msg_blurb)

def table_dump1(input_folder, output_excel_path):
    """
    Loop through all PDFs in the specified directory (and optionally in subdirectories), extract tables, and export them to an Excel sheet.
    
    Parameters:
    input_folder (str): Path to the directory containing PDF files.
    output_excel_path (str): Path where the output Excel file will be saved.
    tables_out (bool): Whether to save the extracted tables into separate sheets in the Excel file.
    sub_folders (bool): Whether to include PDFs from subdirectories.
    """
    if tables_out:
        writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')
        
    all_tables = []
    # data = []
    seller = input("SELLER: ")
    # Walk through directories and subdirectories if sub_folders is True
    for root, dirs, files in os.walk(input_folder) if sub_folders else [(input_folder, [], os.listdir(input_folder))]:
        for filename in files:
            if filename.lower().endswith('.pdf'):
                date, invoice, billTo, shipTo, quantity, description = '', '', '', '', '', ''
                priceEach, subtotal, amount, total, billToState, shipToState = '', '', '', '', '', ''
                customer, state, year, month, day, folder = '', '', '', '', '', ''
                
                row_data = {}
                pdf_path = os.path.join(root, filename)
                folder = pdf_path
                # tables = extract_tables_from_pdf(pdf_path)
                tables, text = extract_tables_from_pdf(pdf_path)
                all_tables.extend(tables)


    if tables_out:
        print(f'{color_purple}tables_out section hit{color_reset}')
        
        for idx, table in enumerate(all_tables):
            sheet_name = f'Table{idx+1}'
            table.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()
        writer.close()

    # write_xlsx(data)    


def write_xlsx(data):
    # process_invoice_pdfs(writing test data = {data}')    # temp
    global workbook
    workbook = Workbook()
    global worksheet
    worksheet = workbook.active

    worksheet.title = 'Invoices'
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'D2'  # Freeze cells
    worksheet.selection = 'D2'

    headers = [
        "Date", "Invoice", "GrandTotal", "SalesTax", "SalesTaxRate"
        , "Qty", "Description", "Price", "Amount"
        , "Seller", "BillTo", "TaxID", "Note"
        , "year", "month", "day", "text", "sort", "filename", "temp", "First", "Second"
        , "Middle", "Second2last", "Last", "TP", "POnumber"
        # , "DatePaid", "Subtotal", "Code", "Terms"
        # , "Total", "Discount", "Shipping", "Shipping_Details", "Billing_Details", "TotalStoreCredit", "TotalTaxCredited", "PreviousStoreCredit", "Customer", "Illinois_License"
        # , "Rep", "Ship_Date"
        # ,"Order_Date", "Order", "SKU", "Product", "MSRP", "DBA", "Email", "Phone"
    ]


    # --- Create second worksheet with same headers ---
    worksheet2 = workbook.create_sheet(title="Extra_Lines")

    # Freeze + selection (optional, matches first sheet)
    worksheet2.freeze_panes = 'D2'
    worksheet2.selection = 'D2'

    for col_index, header in enumerate(headers, start=1):
        cell = worksheet2.cell(row=1, column=col_index)
        cell.value = header

        # Columns 1–6 → Orange
        if 1 <= col_index <= 13:
            cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

        # Columns 23-24 → Yellow
        elif 14 <= col_index <= 16:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Columns 15–22 → Green
        elif 17 <= col_index <= 19:
            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    for col_letter, dim in worksheet.column_dimensions.items():
        worksheet2.column_dimensions[col_letter].width = dim.width


    for col_index, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=1, column=col_index)
        cell.value = header

        # Columns 1–9 → Orange
        if 1 <= col_index <= 13:
            cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

        # Columns 23-24 → Yellow
        elif 14 <= col_index <= 16:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Columns 16-20 → Green
        elif 17 <= col_index <= 19:
            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")


    # Excel column width
    worksheet.column_dimensions['A'].width = 12 # 
    worksheet.column_dimensions['B'].width = 11 # 
    worksheet.column_dimensions['C'].width = 12 # 
    worksheet.column_dimensions['D'].width = 11 # 
    worksheet.column_dimensions['E'].width = 13 # 
    worksheet.column_dimensions['F'].width = 14 #   
    worksheet.column_dimensions['G'].width = 20 # 
    worksheet.column_dimensions['H'].width = 8 # 
    worksheet.column_dimensions['I'].width = 8 # 
    worksheet.column_dimensions['J'].width = 20 # 
    worksheet.column_dimensions['K'].width = 20  # 
    worksheet.column_dimensions['L'].width = 10 # 
    worksheet.column_dimensions['M'].width = 12  # 
    worksheet.column_dimensions['N'].width = 5  # 
    worksheet.column_dimensions['O'].width = 6  # 
    worksheet.column_dimensions['P'].width = 4  # 
    worksheet.column_dimensions['Q'].width = 25  # 
    worksheet.column_dimensions['R'].width = 5  # 
    worksheet.column_dimensions['S'].width = 24  # 
    worksheet.column_dimensions['T'].width = 20  # 
    worksheet.column_dimensions['U'].width = 25  # 
    worksheet.column_dimensions['V'].width = 25  #     
    worksheet.column_dimensions['W'].width = 35  #    
    worksheet.column_dimensions['X'].width = 15  #    
    worksheet.column_dimensions['Y'].width = 25  #    
    worksheet.column_dimensions['Z'].width = 25  #    
    worksheet.column_dimensions['AA'].width = 7  #    
    worksheet.column_dimensions['AB'].width = 20  #    
        

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
    print(f'\n    {color_yellow}insert your pdfs into the {input_folder} folder')
    print(f'\nExample:')
    print(f'    {file} -N') 

    print(f'    {file} -N -I pdfs_single_page') 
    print(f'    {file} -S -I pdfs -O pdfs_single_page ')     
    print(f'    {file} -S   # split pdfs into single page into the {output_split_folder} folder')     

if __name__ == "__main__":

    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
1.0.0 - added -S split and -s sub directories
0.1.5 - working prototype
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""



"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
created for 2026-168 BOA bank statements

-S split pdfs into seperate pages before you process them (-p)
can be adapated to different table column names
-t will optionaly print out the tables in different sheet
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
