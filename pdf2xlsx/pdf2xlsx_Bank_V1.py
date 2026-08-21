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

# Colorize section
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
from colorama import Fore, Back, Style
print(Back.BLACK)
color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>
author = 'LincolnLandForensics'
description = "read pdfs and convert them to a single .xlsx"
version = '1.0.3'


# <<<<<<<<<<<<<<<<<<<<<<<<<<       Global        >>>>>>>>>>>>>>>>>>>>>>>>>>
global ssort
ssort = 0

def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog
    import threading
    import sys
    import os

    root = tk.Tk()
    script_name = os.path.basename(sys.argv[0]) if sys.argv[0] else "pdf2xlsx_Chase_V3.py"
    root.title(f"{script_name} {version}")
    root.geometry("620x540")
    root.resizable(True, True)

    # Vista theme setup
    style = ttk.Style()
    try:
        style.theme_use('vista')
    except Exception:
        try:
            style.theme_use('xpnative')
        except Exception:
            pass

    # Custom styles
    style.configure("TLabel", font=("Segoe UI", 10))
    style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"))
    style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"), foreground="#003366")
    style.configure("TButton", font=("Segoe UI", 10))
    style.configure("Start.TButton", font=("Segoe UI", 11, "bold"), foreground="green")

    # Main Frame
    main_frame = ttk.Frame(root, padding="15")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Title & Description
    title_label = ttk.Label(main_frame, text=f"{script_name} {version}", style="Title.TLabel")
    title_label.pack(anchor="w", pady=(0, 2))
    
    desc_label = ttk.Label(main_frame, text=description, font=("Segoe UI", 10, "italic"), foreground="#555555")
    desc_label.pack(anchor="w", pady=(0, 15))

    # Mode Selection Frame
    mode_frame = ttk.LabelFrame(main_frame, text=" Select Mode ", padding="10")
    mode_frame.pack(fill=tk.X, pady=(0, 10))

    mode_var = tk.StringVar(value="parse")

    def set_split_defaults():
        mode_var.set("split")
        input_entry.delete(0, tk.END)
        input_entry.insert(0, r".\PDFs")
        output_entry.delete(0, tk.END)
        output_entry.insert(0, r".\pdfs_single_page")
        output_label.config(text="Output Folder:")
        update_button_visuals()

    def set_parse_defaults():
        mode_var.set("parse")
        input_entry.delete(0, tk.END)
        input_entry.insert(0, r".\PDFs")
        output_entry.delete(0, tk.END)
        output_entry.insert(0, r".\Bank_DRAFT_V1.xlsx")
        output_label.config(text="Output File:")
        update_button_visuals()

    def update_button_visuals():
        if mode_var.get() == "split":
            split_btn.state(['pressed', 'disabled'])
            parse_btn.state(['!pressed', '!disabled'])
        else:
            parse_btn.state(['pressed', 'disabled'])
            split_btn.state(['!pressed', '!disabled'])

    btn_sub_frame = ttk.Frame(mode_frame)
    btn_sub_frame.pack(fill=tk.X)

    split_btn = ttk.Button(btn_sub_frame, text="Split PDF Pages (-S)", command=set_split_defaults)
    split_btn.pack(side=tk.LEFT, padx=(0, 10), expand=True, fill=tk.X)

    parse_btn = ttk.Button(btn_sub_frame, text="Parse Records (-N)", command=set_parse_defaults)
    parse_btn.pack(side=tk.LEFT, padx=(10, 0), expand=True, fill=tk.X)

    # Input / Output Paths Frame
    paths_frame = ttk.LabelFrame(main_frame, text=" Paths ", padding="10")
    paths_frame.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(paths_frame, text="Input Folder:").grid(row=0, column=0, sticky="w", pady=5)
    input_entry = ttk.Entry(paths_frame, width=50)
    input_entry.grid(row=0, column=1, padx=(5, 5), pady=5, sticky="ew")

    def browse_input():
        dir_selected = filedialog.askdirectory(initialdir=".", title="Select Input Directory containing PDFs")
        if dir_selected:
            dir_selected = os.path.relpath(dir_selected) if os.path.isabs(dir_selected) else dir_selected
            input_entry.delete(0, tk.END)
            input_entry.insert(0, dir_selected)

    browse_input_btn = ttk.Button(paths_frame, text="Browse...", command=browse_input)
    browse_input_btn.grid(row=0, column=2, pady=5)

    output_label = ttk.Label(paths_frame, text="Output File:")
    output_label.grid(row=1, column=0, sticky="w", pady=5)
    output_entry = ttk.Entry(paths_frame, width=50)
    output_entry.grid(row=1, column=1, padx=(5, 5), pady=5, sticky="ew")

    def browse_output():
        if mode_var.get() == "split":
            dir_selected = filedialog.askdirectory(initialdir=".", title="Select Output Directory for Split PDFs")
            if dir_selected:
                dir_selected = os.path.relpath(dir_selected) if os.path.isabs(dir_selected) else dir_selected
                output_entry.delete(0, tk.END)
                output_entry.insert(0, dir_selected)
        else:
            file_selected = filedialog.asksaveasfilename(
                initialdir=".",
                title="Save Excel File As",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
            )
            if file_selected:
                file_selected = os.path.relpath(file_selected) if os.path.isabs(file_selected) else file_selected
                output_entry.delete(0, tk.END)
                output_entry.insert(0, file_selected)

    browse_output_btn = ttk.Button(paths_frame, text="Browse...", command=browse_output)
    browse_output_btn.grid(row=1, column=2, pady=5)

    paths_frame.columnconfigure(1, weight=1)

    # Control Frame (Start Button)
    control_frame = ttk.Frame(main_frame)
    control_frame.pack(fill=tk.X, pady=(0, 10))

    # Log Window
    log_frame = ttk.LabelFrame(main_frame, text=" Terminal Window ", padding="5")
    log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    terminal_text = tk.Text(log_frame, height=10, bg="#1E1E1E", fg="#D4D4D4", insertbackground="white", font=("Consolas", 10))
    terminal_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(log_frame, command=terminal_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    terminal_text.config(yscrollcommand=scrollbar.set)
    terminal_text.config(state='disabled')

    def log_message(message):
        def _append():
            terminal_text.config(state='normal')
            terminal_text.insert(tk.END, message + "\n")
            terminal_text.see(tk.END)
            terminal_text.config(state='disabled')
        root.after(0, _append)

    # Progress & Status Frame
    progress_frame = ttk.Frame(main_frame)
    progress_frame.pack(fill=tk.X)

    progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate")
    progress_bar.pack(fill=tk.X, side=tk.TOP, pady=(0, 5))

    status_label = ttk.Label(progress_frame, text="Status: Ready", font=("Segoe UI", 9))
    status_label.pack(side=tk.LEFT)

    # Start button
    start_btn = ttk.Button(control_frame, text="Start Processing", style="Start.TButton")
    start_btn.pack(fill=tk.X)

    def update_progress(percent):
        def _update():
            progress_bar['value'] = percent
            status_label.config(text=f"Status: {percent}% Complete")
        root.after(0, _update)

    def reenable_ui():
        start_btn.config(state='normal')
        split_btn.config(state='normal')
        parse_btn.config(state='normal')
        update_button_visuals()
        browse_input_btn.config(state='normal')
        browse_output_btn.config(state='normal')

    def start_processing():
        input_val = input_entry.get().strip()
        output_val = output_entry.get().strip()

        if not input_val:
            log_message("Error: Input path is required.")
            status_label.config(text="Status: Error - Input path missing")
            return
        if not output_val:
            log_message("Error: Output path is required.")
            status_label.config(text="Status: Error - Output path missing")
            return

        # Disable UI
        start_btn.config(state='disabled')
        split_btn.config(state='disabled')
        parse_btn.config(state='disabled')
        browse_input_btn.config(state='disabled')
        browse_output_btn.config(state='disabled')

        # Reset Progress
        progress_bar['value'] = 0
        status_label.config(text="Status: Running... 0% Complete")

        # Clear Terminal
        terminal_text.config(state='normal')
        terminal_text.delete('1.0', tk.END)
        terminal_text.config(state='disabled')

        log_message(f"Starting processing in {mode_var.get().upper()} mode...")

        def run():
            try:
                if mode_var.get() == "split":
                    split_pdfs(
                        input_val,
                        output_val,
                        progress_callback=update_progress,
                        terminal_callback=log_message
                    )
                    log_message(f"\nDone! Output Folder: {os.path.abspath(output_val)}")
                    status_label.config(text="Status: Completed splitting PDFs")
                else:
                    # Parse mode
                    global sub_folders
                    sub_folders = False

                    global output_xlsx
                    output_xlsx = output_val

                    # Run processing
                    data = process_invoice_pdfs(
                        input_val,
                        output_val,
                        progress_callback=update_progress,
                        terminal_callback=log_message
                    )

                    log_message("Writing parsed data to Excel file...")
                    write_xlsx(data)
                    log_message(f"\nDone! Output File: {os.path.abspath(output_xlsx)}")
                    status_label.config(text="Status: Completed parsing PDFs")
            except Exception as e:
                err_msg = f"\n[ERROR] An error occurred:\n{e}"
                print(err_msg)
                log_message(err_msg)
                status_label.config(text="Status: Failed with errors")
            finally:
                root.after(0, reenable_ui)

        threading.Thread(target=run, daemon=True).start()

    start_btn.config(command=start_processing)

    # Initialize defaults to Parse
    set_parse_defaults()

    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()

def main():
    if len(sys.argv) == 1:
        run_gui()
        return

    print(f'{ssort}')
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-b', '--blank', help='blank sheet', required=False, action='store_true')
    parser.add_argument('-N', '--Inoivce_pdfs', help='parse invoices', required=False, action='store_true')

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
        output_xlsx = f"Bank_statements_V1.xlsx"        
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

    elif args.Inoivce_pdfs:
        output_excel_path = 'Bank_statements_.xlsx'


        if not os.path.exists(input_folder):
            print(f"{color_red}Error: The directory '{input_folder}' does not exist.{color_reset}")
        else:
            msg_blurb = (f"Reading Bank PDF's in {input_folder}")
            msg_blurb_square(msg_blurb, color_green)             
        process_invoice_pdfs(input_folder, output_excel_path)
        data = process_invoice_pdfs(input_folder, output_excel_path)
        print(f'{color_yellow} __________________ {ssort} lines processed{color_reset}')    
        write_xlsx(data)
        
    elif args.pdf:
        # Directory containing PDF files
        # input_folder = 'pdfs'
        # Path where the output Excel file will be saved
        output_excel_path = 'output_pdfs.xlsx'
        print(f"running args.pdf")  # temp

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

def account_resolveold(masked: str, accounts: set) -> str:
    # Only process masked accounts
    masked = masked.replace(' ', '')
    if "xxxx" not in masked.lower() :
        return masked

    last4 = masked[-4:]

    for acc in accounts:
        if acc.endswith(last4):
            return acc

def account_resolve(masked: str, customer: str, accounts: list):
    """
    Resolve masked account numbers like 'xxxx9289' to full account.
    If customer == '', populate it from the account list.
    Returns (resolved_account, resolved_customer)
    """
    masked = masked.replace(' ', '')

    # If not masked, return as-is
    if "xxxx" not in masked.lower():
        return masked, customer

    last4 = masked[-4:]

    for entry in accounts:
        acct = entry["Account"]
        cust = entry["Customer"]

        if acct.endswith(last4):
            # If customer is empty, fill it
            if customer == "":
                customer = cust
            return acct, customer

    # No match found
    return masked, customer
    

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
    
def create_date(Date, year, month, day):
    # Normalize year
    year = str(year)
    if len(year) == 2:
        year = "20" + year

    # Normalize month and day
    month = f"{int(month):02d}"
    day   = f"{int(day):02d}"

    # Build ISO date
    Date = f"{year}-{month}-{day}"

    return Date, year, month, day
    
    
def extract_bank_information(tables):
    print(f'{color_green}extract_bank_information{color_reset}')
    """
    Extract specific information from the tables.
    
    Parameters:
    tables (list of pd.DataFrame): List of DataFrames representing tables.
    
    Returns:
    dict: Dictionary containing the extracted values.
    """
    data_tables = {
        "Date": '',
        "Invoice #": "",
        "Bill To:": "",
        "Ship To:": "",
        "Quantity": "",
        "Qty": "",
        "Description": "",
        "Price Each": "",
        "Amount": "",
        "Unit Price": "",
        "Total": "",
        "Product": "",
        "Qty": "",
        "Column1": "",
        "Column2": "",
        "Column3": "",
        "Column4": "",
        "Column5": "",
        "Column6": "",
        "Column7": "",
        "Column8": "",
        "Column9": "",
        "Column10": ""
    }
    
    for table in tables:
        for key in data_tables.keys():
            # print(f'key = {key}')   # temp
            print(f'table.columns = {table.columns}')   # temp
            if key in table.columns:
                data_tables[key] = table.loc[0, key]
                
        if all(data_tables.values()):
            break

    return data_tables
    
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
        "Invoice #": '',
        "Bill To:": '',
        "Ship To:": '',
        "Quantity": '',
        "Qty": '',
        "Description": '',
        "Price Each": '',
        "Amount": '',
        "Unit Price": '',
        "Total": '',
        "Product": '',
        "Qty": '',
        "Column1": '',
        "Column2": '',
        "Column3": '',
        "Column4": '',
        "Column5": '',
        "Column6": '',
        "Column7": '',
        "Column8": '',
        "Column9": '',
        "Column10": ''
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

def extract_statement_from_text(text, data, filename):
    print(f'{color_purple}extract_bank_info_from_text{color_reset}')
    global ssort


    ACCOUNT_LIST = [
        {"Account": "000000012345678", "Customer": ""},
        {"Account": "000000212345678", "Customer": "Walter E Neuman"}, 
        {"Account": "000000312345678", "Customer": "ABC enterprises"},
    ]
    
    Date, Amount, Description, Reference_Number, Bankcard, Account = '', '', '', '', '', ''
    Type, Customer, check, Date, Ledger_Balance = '', '', '', '', ''
    year, month, day, Period, bank = '', '', '', '', ''
    Begin_Balance, Deposits, Withdrawals, Checks_and_deductions, End_Balance, ATM, temp = '', '', '', '', '', '', ''

    # date_pattern = r"(?<!\d/)\b\d{2}/\d{2}\b(?!/\d{2})"
    date_pattern = r"(?<!\d/)\b\d{2}/\d{2}\b(?!/\d{2,4})"   #mm/dd


    
    date_pattern2 = r"\b(0[1-9]|1[0-2])/(0[1-9]|[12]\d|3[01])/\d{2}\b"
    invoice_pattern1 = r"^\d{12}\b"
    invoice_pattern2 = r"^\d{7}\b"    
    invoice_pattern3 = r"M^\d{7}\b" 
    
    match_account1 = re.search(r'Account No\.\:\s*(\d+)\s*Page\:', text)
    match_balance = r'^\s*\d[\d,]*\.\d{2}(?: \d[\d,]*\.\d{2}){3}\s*$'
    
    # Split the text into lines
    lines = text.split('\n')
    row_data = {}
    # Iterate over each line and process
    tally_line = 0

    for line in lines:
        Date, Amount, Description, day, Reference_Number, Bankcard = '', '', '', '', '', ''
        check, Date, Ledger_Balance = '', '', ''
        Deposit, Withdrawal = '', ''
        # Begin_Balance, Deposits, Checks_and_deductions, End_Balance = '', '', '', ''
        temp, month2 = '', ''
        First, Middle, Last, Second2last, Third2last, Fourth2last = '', '', '', '', '', ''
    
        if '/ ' in line:    # ex 10/ 12 701.55 Corporate ACH Tax 512666
            line = line.replace('/ ', '/')
            
        # if '-' in filename:   # if the file was saved as 2025-01.pdf
            # yrtemp = filename
            # yrtemp = yrtemp.split('-')
            # year = yrtemp[0]
            # month = yrtemp[1]
            # try:
                # month = month.split('_')[0]
            # except: pass    
        # if '/ ' in line:
            # line = line.replace('/ ', '/')
        
        matches6 = re.findall(date_pattern, line)
        matches7 = re.findall(date_pattern2, line)  # mm/dd/yy
        # pattern_amount = r"^\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})$"
        pattern_amount = r"^-?\$?\d{1,3}(?:,\d{3})*|\$?\d+(?:\.\d+)?$"
        pattern_amount2 = r"^-?\$?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$"    # 208.81
        # pattern_account = r"^\d{15}$" # 15 digit account numbers
        pattern_account = r"^\d{15,20}$" # 15-20 digit account numbers
        pattern_through = r"(\d)(through)([A-Za-z])"

        tally_line += 1
        ssort += 1


        try:
            linetemp = line.split(' ')
            First = linetemp[0]
            Middle = " ".join(linetemp[2:-1])
            Last = linetemp[-1]
            Second2last  = linetemp[-2]
            Third2last  = linetemp[-3]
            Fourth2last  = linetemp[-4]
        except Exception as e:
            print(f"{color_red}[ERROR1] Failed to parse value:{color_reset} {line}")
            print(f"[EXCEPTION1] {color_red}{e}{color_reset}")


           
        if (re.fullmatch(pattern_amount2, Last) and '.' in Last) or '$' in Last:
            temp = (f'{temp} $').strip()
            if (re.fullmatch(pattern_amount2, Second2last) and '.' in Second2last) or '$' in Second2last:
                temp = (f'{temp} $').strip()
        
        if 'Web site: Chase.com' in line:
            bank = 'Chase'    


        for entry in ACCOUNT_LIST:
            if normalize(line) == normalize(entry["Account"]):
                Account = entry["Account"]
                Account = normalize(Account)
                Customer = entry["Customer"]
                break


        if 'through' in line and ', 20' in line:    # January 01, 2021throughJanuary 29, 2021
            Period = line.replace('through', ' through ')
            # year = line.split(', 20')[-1]
            year = line.split(', ')[-1]
            temp = (f'{temp} _Period1').strip()

            line = re.sub(pattern_through, r"\1 \2 \3", line)


   
        elif 'Account Number:' in line and Account == '':
            AccountTemp = line.replace('Account Number:', '').strip().replace(' ','')
            Account = AccountTemp            
            temp = (f'{temp} _Account1').strip()

        elif line in ACCOUNT_LIST:
            Account = line
            temp = (f'{temp} Account2').strip()
        elif Account == 'Account Number:':
            temp = (f'{temp} _Account3').strip()
            Account = line 
      
        elif line.startswith('Account Number'):
            m = re.search(r'^Account Number\s+(\d+)$', line)
            if m:
                Account = m.group(1)
                temp = (f'{temp} Account4').strip()
        elif re.search(pattern_account, line) and Account == '':
            Account = line.replace(' ','')
            temp = (f'{temp} _Account5').strip()
            
        elif '*start*' in line:
            Type = line.replace('*start*', '_') 
            temp = (f'_*start*1').strip()            

        elif line == 'ATM & DEBIT CARD SUMMARY'or line == '*start*atm and debit card summary':
            Type = 'ATM & DEBIT CARD SUMMARY' 
            temp = (f'{temp} _ATM & DEBIT CARD SUMMARY1').strip()
        elif line == 'CHECKS PAID' or '*start*checks paid' in line:
            temp = (f'{temp} _CHECKS PAID1').strip()
            Type = 'CHECKS PAID'

        elif 'PLAT BUS CHECKING' in line:
            Type = 'PLAT BUS CHECKING'  # 8/4/2021 PLAT BUS CHECKING (...8878) - chase.com          
            if ' (...' in line:
                linetemp = line.split(' (...')
                try:
                    Account = linetemp[1]
                    if ')' in Account:
                        Account = Account.split(')')[0]
                    Account = (f'xxxx {Account}')
                    # Account = account_resolveold(Account, ACCOUNT_LIST)       
                    Account, Customer = account_resolve(Account, Customer, ACCOUNT_LIST)
                except:pass

        elif "Interest paid in" in line:
            m = re.search(pattern_account, line)
            if m:
                Account = m.group(0)

        elif 'CHECKING SUMMARY' in line:
            Type = 'CHECKING SUMMARY'

        elif line == 'DEPOSITS AND ADDITIONS' or line == '*start*deposits and additions':
            Type = 'DEPOSITS AND ADDITIONS' 
            temp = (f'{temp} _DEPOSITS AND ADDITIONS1').strip()
            
        elif line == 'DAILY ENDING BALANCE' or '*start*daily ending balan' in line:
            Type = 'DAILY ENDING BALANCE'  
            temp = (f'{temp} _DAILY ENDING BALANCE1').strip()    
        elif line == 'ELECTRONIC WITHDRAWALS' or line == '*start*electronic withdrawal':
            Type = 'ELECTRONIC WITHDRAWALS'
            temp = (f'{temp} _ELECTRONIC WITHDRAWALS').strip()
        elif line == 'SERVICE CHARGE SUMMARY' or '*start*service charge summary' in line:
            Type = 'SERVICE CHARGE SUMMARY' 
            temp = (f'{temp} _SERVICE CHARGE SUMMARY1').strip() 
        elif line == 'SERVICE CHARGE DETAIL' or '*start*service charge detail' in line:
            Type = 'SERVICE CHARGE DETAIL'  
            temp = (f'{temp} _SERVICE CHARGE DETAIL1').strip()    
        elif line == 'SAVINGS SUMMARY' or '*start*summary' in line:
            Type = 'SAVINGS SUMMARY' 
            temp = (f'{temp} _SAVINGS SUMMARY1').strip()    
        elif line == 'OTHER WITHDRAWALS' or '*start*other withdrawals' in line:
            Type = 'OTHER WITHDRAWALS'
            temp = (f'{temp} _OTHER WITHDRAWALS1').strip() 
        elif line == 'TRANSACTION DETAIL' or '*start*transaction detail' in line:
            Type = 'TRANSACTION DETAIL'
            temp = (f'{temp} _TRANSACTION DETAIL1').strip()  
        elif line == 'ACCOUNT ACTIVITY' or line == 'ACCOUNT ACTIVITY (CONTINUED)' or line == 'AACCCCOOUUNNTT AACCTTIIVVIITTYY':
            Type = 'ACCOUNT ACTIVITY'
            temp = (f'{temp} _Account Activity1').strip()
        elif 'ACCOUNT ACTIVITY' in line:
            Type = 'ACCOUNT ACTIVITY'
            temp = (f'{temp} Account Activity2').strip()

        elif line == 'FEES CHARGED':
            Type = 'FEES CHARGED'
        elif line == 'FEES'or '*start*fees section' in line:
            Type = 'FEES'
            temp = (f'{temp} _FEES1').strip()
        elif line == 'CASH ADVANCES':
            Type = 'CASH ADVANCES'              
        elif line == 'PURCHASE':
            Type = 'PURCHASE'             
        elif line == 'PAYMENTS AND OTHER CREDITS':
            Type = 'PAYMENTS AND OTHER CREDITS'           
        elif line == 'Chase Private Client Savings':
            Type = 'Chase Private Client Savings'            
            if ' ' in line:
                try:
                    linetemp = line.split(' ')
                    Account = linetemp[4]
                    Begin_Balance = linetemp[-2]
                    End_Balance = linetemp[-1]
                except:pass


            
        elif 'Ending Balance $' in line or 'Ending Balance -$' in line:    # Ending Balance $1,121.36
            End_Balance = line.replace('Ending Balance ', '')
            temp = (f'End_Balance1 {temp}').strip()
            
        elif 'Beginning Balance $' in line or 'Beginning Balance -$' in line:    # Beginning Balance -$894.66
            Begin_Balance = line.replace('Beginning Balance ', '')
            print(f'________________________ Begin_Balance = {Begin_Balance}')  # temp
            temp = (f'{temp} _Begin_Balance1').strip()

        elif 'Deposits and Additions ' in line:
            # if Type == 'CHECKING SUMMARY' or Type == 'SAVINGS SUMMARY':                      'DEPOSITS AND ADDITIONS',
                    # 'ELECTRONIC WITHDRAWALS',
                    # 'TRANSACTION DETAIL',
                    
                    
                Deposits = line.replace('Deposits and Additions ', '')
                temp = (f'{temp} Deposits1').strip()
                if ' ' in Deposits:
                    Deposits = Deposits.split(' ')[-1]
                    temp = (f'{temp} Deposits2').strip()

        elif 'Electronic Withdrawals ' in line:
 
                Withdrawals = line.replace('Electronic Withdrawals ', '')
                temp = (f'{temp} Withdrawals1').strip()
                if ' ' in Withdrawals:
                    Withdrawals = Withdrawals.split(' ')[-1]
                    temp = (f'{temp} Withdrawals2').strip()
                    
                    
        elif 'Checks Paid ' in line and Type == 'CHECKING SUMMARY':
            Checks_and_deductions = line.replace('Checks Paid ', '')
            if ' ' in Checks_and_deductions:
                Checks_and_deductions = Checks_and_deductions.split(' ')[-1]

        elif 'Electronic Withdrawals ' in line and Type == 'CHECKING SUMMARY':
            ATM = line.replace('Electronic Withdrawals ', '')
            if ' ' in ATM:
                ATM = ATM.split(' ')[-1]
                
        elif 'Ending Balance ' in line:
            if Type == 'CHECKING SUMMARY'or Type == 'SAVINGS SUMMARY':
                End_Balance = line.replace('Ending Balance ', '')
                temp = (f'{temp} End_Balance2').strip()
                if ' ' in End_Balance:
                    End_Balance = End_Balance.split(' ')[-1]
                    temp = (f'{temp} End_Balance3').strip()


        elif len(matches6) == 3:    # mm/dd
            if Type == 'Checks and Substitute Checks':
                line = line.replace(' *','')
            if '/ ' in line:
                line = line.replace('/ ', '')
            temp = "3 dates"
            temp = (f'{temp} 3/dates').strip()
            if Type == 'Checks and Substitute Checks':
                line = line.replace(' *','')            
            
            if '*end*' in line:
                Type = (f'_{line}*')
                temp = (f'_{line}').strip() 
            
            elif ' ' in line:
                linetemp = line
                linetemp = linetemp.replace('Effective Date ', '')
                linetemp = linetemp.split(' ')
                if Type == 'Checks and Substitute Checks':
                    Date2, check2, Amount2, Reference_Number2 = '', '', '', ''
                    Date3, check3, Amount3, Reference_Number3 = '', '', '', ''
                    try:
                        Date = linetemp[0]
                        check = linetemp[1]
                        Amount = f'-{linetemp[2]}'
                        Reference_Number = linetemp[3]
                        Date2 = linetemp[4]
                        check2 = linetemp[5]
                        Amount2 = linetemp[6]
                        Reference_Number2 = linetemp[7]
                        Date3 = linetemp[8]
                        check3 = linetemp[9]
                        Amount3 = linetemp[10]
                        Reference_Number3 = linetemp[11]
                    except:
                        pass

                    if '/' in Date:
                        DayTemp = Date.split('/')
                        month = DayTemp[0]
                        day = DayTemp[1]
                    if '/' in Date2:
                        DayTemp2 = Date2.split('/')
                        month2 = DayTemp2[0]
                        day2 = DayTemp2[1] 
                    if '/' in Date3:
                        DayTemp3 = Date3.split('/')
                        month3 = DayTemp3[0]
                        day3 = DayTemp3[1] 


                    row_data = {}
                    row_data["Date"] = Date2   
                    row_data["Amount"] = f'-{Amount2}'
                    row_data["Reference_Number"] = Reference_Number2
                    row_data["Account"] = Account
                    row_data["Type"] = Type
                    row_data["check"] = check2
                    row_data["filename"] = filename
                    row_data["year"] = year
                    row_data["month"] = month
                    row_data["day"] = day2
                    row_data["text"] = line
                    row_data["Period"] = Period
                    row_data["bank"] = bank
                    row_data["sort"] = ssort
                    row_data["temp"] = "3 dates, 2nd row"
                    data.append(row_data)
                    ssort += 1
                    row_data = {}
                    row_data["Date"] = Date3   
                    row_data["Amount"] = f'-{Amount3}'
                    row_data["Reference_Number"] = Reference_Number3
                    row_data["Account"] = Account
                    row_data["Type"] = Type
                    row_data["check"] = check3
                    row_data["filename"] = filename
                    row_data["year"] = year
                    row_data["month"] = month3
                    row_data["day"] = day3
                    row_data["text"] = line
                    row_data["Period"] = Period
                    row_data["bank"] = bank
                    row_data["sort"] = ssort
                    row_data["temp"] = "3 dates, 3rd row"
                    data.append(row_data)
                    ssort += 1




        elif 1 == 2:
        # elif len(matches6) == 2:    # mm/dd
            if '*end*' in line:
                Type = (f'_{line}*')
                temp = (f'_{line}').strip() 
            elif '/ ' in line:
                line = line.replace('/ ', '')
            temp = (f'{temp} 2/dates_').strip()

            if re.fullmatch(pattern_amount2, Last) and re.fullmatch(date_pattern, First):
                Date = First
                Description = Middle
                Amount = Last
                if len(year) == 2:
                    year = "20" + year
                    
                if day == '':
                    Date = Date.replace('/', '-')
                    if '-' in Date:
                        DateTemp = Date.split('-')
                        try:
                            day = DateTemp[0]
                            month = DateTemp[1]
                        except:pass
                if year and month and day:
                    Date = f"{year}-{month}-{day}"
       


            elif ' ' in line:
                if Type in (
                    'DEPOSITS AND ADDITIONS',
                ):
                    # print(f'{color_green}Type = {Type} line = {color_yellow}{line}{color_reset}')   # temp
                    linetemp = line
                    linetemp = linetemp.split(' ')
                    try:
                        Date = linetemp[0]
                        Amount = linetemp[-1]
                        Deposit = Amount
                        Description = " ".join(linetemp[2:-1])  
                        
                        
                        
                        # Description = Description.replace(f'{Date}', '')
                        # Description = Description.replace(f'{Amount}', '').strip()
                    except:pass
                
                if Date == '':
                    Date = line.split(' ')[0]
                    if any(c.isalpha() for c in Date):
                        Date = ""
                elif ' ' in line and Amount == '':
                    m = re.search(pattern_amount2, line)
                    if m:
                        Amount = m.group(0)
                    elif Amount == '':
                        Amount = line.split(' ')[-1]

                if '/' in Date:
                    DayTemp = Date.split('/')
                    month = DayTemp[0]
                    day = DayTemp[1]
                    year = DayTemp[-1]


                

        elif len(matches6) == 1 or len(matches6) == 2:
            if '*end*' in line:
                Type = (f'_{line}*')
                temp = (f'_{line}').strip() 
            elif len(matches6) == 1:            
                temp = (f'{temp} 1/date').strip()
            elif len(matches6) == 2:            
                temp = (f'{temp} 2/dates').strip()              
            
            linetemp = line.split(" ")
            # --- Amount extraction (last token) ---
            
            last_token = linetemp[-1]

            # Amount = None
            if re.fullmatch(pattern_amount, last_token):
                Amount = last_token

            first_token = linetemp[0]
            check = None
            if re.fullmatch(r"^\d+$", first_token):
                check = first_token

            # --- Date Posted ---
            Date = matches6[0]

            if ' ' in line:
                if Type in (
                    'DEPOSITS AND ADDITIONS',
                    'ELECTRONIC WITHDRAWALS',
                    # 'TRANSACTION DETAIL',
                    'PURCHASE',
                    'PAYMENTS AND OTHER CREDITS',
                    'ACCOUNT ACTIVITY',

                ):
                    # print(f'{color_green}Type = {Type} line = {color_yellow}{line}{color_reset}')   # temp
                    linetemp = line
                    linetemp = linetemp.split(' ')
                    try:
                        Date = linetemp[0]
                        Amount = linetemp[-1]
                        Description = " ".join(linetemp[1:-1])                        
                        temp = (f'{temp} batch1').strip()

                    except:pass
                    if Type in (

                        'ACH Deductions',
                        'Debit Card Purchases',
                        'POS Purchases',
                        'Service Charges and Fees'
                    ):
                        Amount = f'-{Amount}'

                elif Type in (
                    'CHECKS PAID',
                ):
                    # print(f'{color_green}Type = {Type} line = {color_yellow}{line}{color_reset}')   # temp
                    linetemp = line
                    linetemp = linetemp.split(' ')
                    try:
                        check = linetemp[0]
                        Date = linetemp[-2]
                        Amount = linetemp[-1]
                        Description = line
                        temp = (f'{temp} batch2').strip()

                    except:pass
                elif Type in (
                    'TRANSACTION DETAIL',
                ):
                    # print(f'{color_green}Type = {Type} line = {color_yellow}{line}{color_reset}')   # temp
                    linetemp = line
                    linetemp = linetemp.split(' ')
                    try:
                        Date = linetemp[0]
                        Amount = linetemp[-2]
                        Description = " ".join(linetemp[1:-2]) 
                        temp = (f'{temp} batch3').strip()                       

                    except:pass                
                    temp = (f'{temp} Transaction1').strip()
                
                
                elif ' ' in line:
                    m = re.search(pattern_amount2, line)
                    if m:
                        Amount = m.group(0)
                    elif Amount == '':
                        Amount = line.split(' ')[-1]
                        


                if any(c.isalpha() for c in Date):
                    Date = ''
                    Amount = ''
                    Type = ''
                    Description = ''    
                if 'This Cycle Deposits' in line:
                    Ledger_Balance = line.split(' ')[-1]
                    Date = line.split(' ')[-2]
                
                if Type == 'Checks and Substitute Checks':
                    try:
                        Date = linetemp[0]
                        check = linetemp[1]
                        Amount = f'-{linetemp[2]}'
                        Reference_Number = linetemp[3]
                        temp = (f'{temp} batch4').strip()
                    except:
                        pass

            if '/' in Date:
                DayTemp = Date.split('/')
                month = DayTemp[0]
                day = DayTemp[1]
                # year = DayTemp[-1]

            if any(c.isalpha() for c in Amount):
                Amount = ""
            elif '$' in Amount:
                Amount = Amount.replace('$','')

        elif 'Bankcard ' in line:
            Bankcard = line       
            Bankcard = Bankcard.replace('Bankcard ', '')
        elif 'Merchant Bnkcd ' in line:
            Bankcard = line       
            Bankcard = Bankcard.replace('Merchant Bnkcd ', '')

        elif Type == "PLAT BUS CHECKING" and '$' in temp:
            Amount = Last           
            if '-' in Amount:
                Withdrawal = Amount.replace('-', '')
            if ' ' in line:
                linetemp = line.split(' ')
                try:
                    Description = " ".join(linetemp[0:-1]) 
                except:pass
        elif '*end*summary' in line or 'Annual Percentage Yield Earned This Period' in line or 'You earned a higher interest rate on your Chase' in line:
            Begin_Balance, Deposits, Withdrawals, Checks_and_deductions, End_Balance, ATM = '', '', '', '', '', ''
            temp = (f'{temp} _EndBalanceClear1').strip()

        elif '*end*' in line:
            Type = (f'_{line}*')
            temp = (f'_{line}').strip()  


        if Account:
            if ' ' in Account:
                Account = Account.replace(' ', '')
        else:
            Account = ''
        
        if 'xxxx' in Account.lower():
            # Account = account_resolveold(Account, ACCOUNT_LIST)
            Account, Customer = account_resolve(Account, Customer, ACCOUNT_LIST)

        # if Amount and Date_Posted:
        if Amount:
            if Type == 'DEPOSITS AND ADDITIONS' or Type == 'Customer Deposits' or Type == 'Customer Deposits':  # Chase



                Deposit = Amount
                temp = (f"{temp} Deposits1").strip()
            elif Type == 'CHECKS PAID' or Type == 'ELECTRONIC WITHDRAWALS' or Type == 'FEES CHARGED' or Type == 'PAYMENTS AND OTHER CREDITS' or Type == 'PLAT BUS CHECKING' or Type == 'PURCHASE':
                Withdrawal = Amount.replace('-', '')
                temp = (f"{temp} Withdrawal1").strip()

            elif Type == 'TRANSACTION DETAIL':
                if '-' in Amount:
                    Withdrawal = Amount.replace("-", "")
                    temp = (f"{temp} Withdrawal2").strip()
                else:        
                    Deposit = Amount
                    temp = (f"{temp} Deposit2").strip()

        # If year is blank, extract last 4 digits from the line
        if year == '':
            Period = Period.strip()
            tail = Period[-4:]                 # last 4 characters
            if tail.isdigit():               # ensure they are digits
                year = tail
        
        if month == '':
            if 'redeem' in Period:
                Period = ''
            elif '20' in Period:
        # if month == '' and '20' in Period and ' ' in Period:
                PeriodTemp = Period.split(' ')[0]
                month = PeriodTemp
                if month.isalpha():
                    month = month_to_number(month)

        if year and month and day:
            Date, year, month, day = create_date(Date, year, month, day)
        if not year.strip().startswith('2'):
            year = ''

        print(f'{filename}  {color_blue}{line}{color_reset}')   # worked
        row_data = {}
        row_data["Date"] = Date   
        row_data["Amount"] = Amount
        row_data["Description"] = Description
        row_data["Reference_Number"] = Reference_Number
        row_data["Bankcard"] = Bankcard
        row_data["Account"] = Account
        row_data["Type"] = Type
        row_data["Customer"] = Customer
        row_data["check"] = check
        row_data["Date"] = Date
        row_data["Ledger_Balance"] = Ledger_Balance
        row_data["filename"] = filename
        row_data["year"] = year
        row_data["month"] = month
        row_data["day"] = day
        row_data["text"] = line
        row_data["Period"] = Period
        row_data["bank"] = bank
        row_data["Begin_Balance"] = Begin_Balance
        row_data["Deposits"] = Deposits
        row_data["Withdrawals"] = Withdrawals
        row_data["Checks_and_deductions"] = Checks_and_deductions
        row_data["ATM"] = ATM        
        row_data["End_Balance"] = End_Balance
        row_data["sort"] = ssort
        row_data["temp"] = temp
        row_data["First"] = First
        row_data["Middle"] = Middle        
        row_data["Last"] = Last
        row_data["Second2last"] = Second2last
        row_data["Third2last"] = Third2last
        row_data["Fourth2last"] = Fourth2last
        row_data["Deposit"] = Deposit
        row_data["Withdrawal"] = Withdrawal
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
        Account = match_account1.group(1)
        print(f'{color_green}&&&&&&&&&&&&&&&&&&&Extracted Account Number: {Account}{color_reset}')

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

def month_to_number(month: str) -> str:
    month_map = {
        "january": "01",
        "february": "02",
        "march": "03",
        "april": "04",
        "may": "05",
        "june": "06",
        "july": "07",
        "august": "08",
        "september": "09",
        "october": "10",
        "november": "11",
        "december": "12",
    }

    m = month.strip().lower()
    return month_map.get(m)
    

def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')

def normalize(acct):
    return acct.replace(" ", "").replace("-", "")
    
def process_invoice_pdfs(input_folder, output_excel_path, progress_callback=None, terminal_callback=None):
    print(f'{color_purple}process_bank_pdfs{color_reset}')  # worked
 
    """
    Loop through all PDFs in the specified directory (and optionally in subdirectories), extract tables, and export them to an Excel sheet.
    
    Parameters:
    input_folder (str): Path to the directory containing PDF files.
    output_excel_path (str): Path where the output Excel file will be saved.
    tables_out (bool): Whether to save the extracted tables into separate sheets in the Excel file.
    sub_folders (bool): Whether to include PDFs from subdirectories.
    """
    # if tables_out:
        # writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')
    global tally
    
    tally = 0 
    all_tables = []
    data = []
    # customer = input("CUSTOMER: ")
    customer = ''
    
    # Gather PDF files
    pdf_files = []
    for root, dirs, files in os.walk(input_folder) if sub_folders else [(input_folder, [], os.listdir(input_folder))]:
        for filename in files:
            if filename.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, filename))

    total_files = len(pdf_files)
    if total_files == 0:
        msg = "No PDF files found to parse."
        print(msg)
        if terminal_callback:
            terminal_callback(msg)
        if progress_callback:
            progress_callback(100)
        return data

    for idx, pdf_path in enumerate(pdf_files):
        filename = os.path.basename(pdf_path)
        root = os.path.dirname(pdf_path)
        msg_file = f"Input file: {filename}"
        print(msg_file)
        if terminal_callback:
            terminal_callback(msg_file)

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
        try:
            tables, text = extract_tables_from_pdf(pdf_path)
            data = extract_statement_from_text(text, data, filename)
            print(f'{color_green}data test 4{color_reset}') # works
        except Exception as e:
            err_msg = f"Error processing file {pdf_path}: {e}"
            print(err_msg)
            if terminal_callback:
                terminal_callback(err_msg)

        tally += 1
        if progress_callback:
            progress_callback(int((idx + 1) / total_files * 100))

    print(f"{color_green}processed {tally} PDF's{color_reset}")
    return data
    # write_xlsx(data)



def process_bank_pdfs_in_directory(input_folder, output_excel_path):
    print(f'{color_green}process_bank_pdfs_in_directory{color_reset}')


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
    tally = 1   
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
                grand_total, shipping, subtotal, discount, total_tax  = '', '', '', '', ''
                date1, amt, Account, check, SEQ, TC  = '', '', '', '', '', ''
                date2, amt2, account2, check2, SEQ2, TC2 =  '', '', '', '', '', ''
                date3, amt3, account3, check3, SEQ3, TC3  = '', '', '', '', '', ''
                date4, amt4, account4, check4, SEQ4, TC4  = '', '', '', '', '', ''
                
                
                row_data = {}
                pdf_path = os.path.join(root, filename)
                folder = pdf_path
                tables, text = extract_tables_from_pdf(pdf_path)
                # print(f'text = {text}') # temp
                disclaimer, total, date, grand_total, shipping, invoice, total_tax, shipToState, amt, Account, check, SEQ, TC, amt2, account2, check2, SEQ2, TC2 = extract_info_from_text(text, total, date, grand_total, shipping, invoice, total_tax, shipToState, amt, Account, check, SEQ, TC, amt2, account2, check2, SEQ2, TC2)

                if " AMT: $" in text:
                    row_data["amt"] = amt
                    row_data["Account"] = Account
                    row_data["check"] = check
                    row_data["SEQ"] = SEQfor 
                    row_data["TC"] = TC
                    data.append(row_data)

                    row_data["amt"] = amt2
                    row_data["account"] = account2
                    row_data["check"] = check2
                    row_data["SEQ"] = SEQ2
                    row_data["TC"] = TC2
                    data.append(row_data)

                    row_data["amt"] = amt3
                    row_data["account"] = account3
                    row_data["check"] = check3
                    row_data["SEQ"] = SEQ3
                    row_data["TC"] = TC3
                    data.append(row_data)

                    row_data["amt"] = amt4
                    row_data["account"] = account4
                    row_data["check"] = check4
                    row_data["SEQ"] = SEQ4
                    row_data["TC"] = TC4
                    data.append(row_data)

           

                all_tables.extend(tables)

                data_tables = extract_information(tables)
                for key, value in data_tables.items():
                    if value:
                        if key == "Amount":
                            amount = value
                        elif key == "Date":
                            date = value
                        elif key == "Date/Time":
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
                        elif key == "Action":
                            description = value
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
                if total is None or total == 0 or total == '0':
                    total == ''
                if description is None or description == '':
                    description = tables

                pattern = r',\s([A-Z]{2})\s\d{5}'

                match1 = re.search(pattern, billTo)
                if match1:
                    billToState = match1.group(1)
                else:
                    billToState = ''

                match2 = re.search(pattern, shipTo)
                if match2:
                    shipToState = match2.group(1)

                match3 = re.search(r'^(.*)$', shipTo, re.MULTILINE)
                if match3:
                    customer = match3.group(1)
                else:
                    customer = ''

                pattern4 = r'\$([0-9,]+\.[0-9]{2})'
                match4 = re.findall(pattern4, amount)

                total = sum(float(value4.replace(',', '')) for value4 in match4)    
                total = round(float(total), 2)
                
                try:
                    if date.count('/') == 2:
                        input_date = datetime.strptime(date, '%m/%d/%Y')
                        date = input_date.strftime('%Y-%m-%d')
                        year = input_date.strftime('%Y')
                        month = input_date.strftime('%m')
                        day = input_date.strftime('%d')
                except Exception as e:
                    print(f"{color_red}Error converting date: {str(e)}{color_reset}")


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
                row_data["subtotal"] = subtotal
                row_data["discount"] = discount
                row_data["total_tax"] = total_tax
                # bank records
                row_data["amt"] = amt
                row_data["Account"] = Account
                row_data["check"] = check
                row_data["SEQ"] = SEQ
                row_data["TC"] = TC


                print(f'{color_blue}{tally}.    {color_green}{date}  {color_yellow}{filename}    {color_green}{grand_total}{color_reset}')   # test
                data.append(row_data)
                tally += 1
    if tables_out:
        for idx, table in enumerate(all_tables):
            sheet_name = f'Table{idx+1}'
            table.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()
        writer.close()

    write_xlsx(data)


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
                    customer = match3.group(1)
                else:
                    customer = ''

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

                print(f'{filename}    total ={total}')   # test
                data.append(row_data)

    if tables_out:
        for idx, table in enumerate(all_tables):
            sheet_name = f'Table{idx+1}'
            table.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()
        writer.close()

    write_xlsx(data)



def split_pdfs(input_folder, output_split_folder, progress_callback=None, terminal_callback=None):
    # from PyPDF2 import PdfReader, PdfWriter

    """
    Function to split PDFs into single-page PDFs.
    """
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_split_folder):
        os.makedirs(output_split_folder)
        msg = f"Created split PDF folder: {output_split_folder}"
        print(msg)
        if terminal_callback:
            terminal_callback(msg)

    msg_blurb = (f'Splitting PDFs from {input_folder} folder')
    msg_blurb_square(msg_blurb, color_green)
    if terminal_callback:
        terminal_callback(msg_blurb)
    
    pdf_files = []
    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            if filename.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, filename))

    total_files = len(pdf_files)
    if total_files == 0:
        msg = "No PDF files found to split."
        print(msg)
        if terminal_callback:
            terminal_callback(msg)
        if progress_callback:
            progress_callback(100)
        return

    for idx, pdf_path in enumerate(pdf_files):
        filename = os.path.basename(pdf_path)
        root = os.path.dirname(pdf_path)
        msg_file = f"Input file: {filename}"
        print(msg_file)
        if terminal_callback:
            terminal_callback(msg_file)

        try:
            with open(pdf_path, 'rb') as pdf_file:
                reader = PdfReader(pdf_file)
                
                for page_num, page in enumerate(reader.pages, start=1):
                    writer = PdfWriter()
                    writer.add_page(page)
                    
                    # Construct output path maintaining subdirectory structure
                    relative_path = os.path.relpath(root, input_folder)
                    output_dir = os.path.join(output_split_folder, relative_path)
                    if not os.path.exists(output_dir):
                        os.makedirs(output_dir)
                    
                    # Pad the page number with 0 if there aren't at least 7 characters (i.e. :07d)
                    output_filename = f"{os.path.splitext(filename)[0]}_page_{page_num:07d}.pdf"
                    output_path = os.path.join(output_dir, output_filename)
                    
                    with open(output_path, 'wb') as output_pdf:
                        writer.write(output_pdf)
                    
                    print(f"Created: {output_path}")
        except Exception as e:
            err_msg = f"Error processing file {pdf_path}: {e}"
            print(err_msg)
            if terminal_callback:
                terminal_callback(err_msg)

        if progress_callback:
            progress_callback(int((idx + 1) / total_files * 100))

    msg_blurb = (f'Finished writing split PDFs to {output_split_folder} folder')
    msg_blurb_square(msg_blurb, color_green)
    if terminal_callback:
        terminal_callback(msg_blurb)

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

    worksheet.title = 'Bank_Statements'
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'D2'  # Freeze cells
    worksheet.selection = 'D2'

    headers = [
        "Date", "Amount", "Description", "Account", "Customer", "Type", "check"
        , "year", "month", "day", "text", "bank", "sort", "temp", "filename", "Period"
        , "Begin_Balance", "End_Balance", "Deposits", "Withdrawals", "Checks_and_deductions"
        , "First", "Middle" , "Fourth2last", "Third2last", "Second2last", "Last"
        , "Deposit", "Withdrawal"
        , "ATM", "Other_Subtractions", "Checks_and_deductions"

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
        if 1 <= col_index <= 7:
            cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

        # Columns 23-24 → Yellow
        elif 8 <= col_index <= 10:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


        # Columns 15–22 → Green
        elif 17 <= col_index <= 21:
            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

    for col_letter, dim in worksheet.column_dimensions.items():
        worksheet2.column_dimensions[col_letter].width = dim.width



    for col_index, header in enumerate(headers, start=1):
        cell = worksheet.cell(row=1, column=col_index)
        cell.value = header

        # Columns 1–9 → Orange
        if 1 <= col_index <= 7:
            cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")

        # Columns 23-24 → Yellow
        elif 8 <= col_index <= 10:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


        # Columns 16-20 → Green
        elif 17 <= col_index <= 21:
            cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")



 
    # for col_index, header in enumerate(headers):
        # cell = worksheet.cell(row=1, column=col_index + 1)
        # cell.value = header
        # if col_index in range(21):  # range(18) generates numbers from 0 to 25 inclusive orange
            # fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # orange
            # cell.fill = fill

    # Excel column width
    worksheet.column_dimensions['A'].width = 12 # 
    worksheet.column_dimensions['B'].width = 9 # 
    worksheet.column_dimensions['C'].width = 20 # 
    worksheet.column_dimensions['D'].width = 22 # 
    worksheet.column_dimensions['E'].width = 27 # 
    worksheet.column_dimensions['F'].width = 23 #   
    worksheet.column_dimensions['G'].width = 10 # 
    worksheet.column_dimensions['H'].width = 5 # 
    worksheet.column_dimensions['I'].width = 7 # 
    worksheet.column_dimensions['J'].width = 4 # 
    worksheet.column_dimensions['K'].width = 40  # 
    worksheet.column_dimensions['L'].width = 10  # 
    worksheet.column_dimensions['M'].width = 6  # 
    worksheet.column_dimensions['N'].width = 15  # 
    worksheet.column_dimensions['O'].width = 38  # 
    worksheet.column_dimensions['P'].width = 14  # 
    worksheet.column_dimensions['Q'].width = 9  # 
    worksheet.column_dimensions['R'].width = 14  # 
    worksheet.column_dimensions['S'].width = 18  # 
    worksheet.column_dimensions['T'].width = 20  # 
    worksheet.column_dimensions['U'].width = 12  # 
    worksheet.column_dimensions['V'].width = 11  #     
    worksheet.column_dimensions['W'].width = 10  #    
    worksheet.column_dimensions['X'].width = 15  #    
    worksheet.column_dimensions['Y'].width = 12  # 
    worksheet.column_dimensions['Z'].width = 11  #     
    worksheet.column_dimensions['AA'].width = 5  #    
    worksheet.column_dimensions['AB'].width = 15  #    
     

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
account is exporting
process the checks

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
updated to Chase
updated for 2026-110
created for 2026-126 PNC bank statements

-S split pdfs into seperate pages before you process them (-p)
can be adapated to different table column names
-t will optionaly print out the tables in different sheet
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
