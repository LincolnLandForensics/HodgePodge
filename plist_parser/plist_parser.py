#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>
import os
import plistlib
import xlsxwriter
from datetime import datetime
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import sys

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = '''
Iterates through .plist files in a specified folder (logs_plist).
Extracts and writes key-value pairs along with metadata (file creation, modification times) to an Excel file (output_plists.xlsx).
'''
version = '1.0.3'
# Metadata
author = 'LincolnLandForensics'
version = '1.0.3'

# <<<<<<<<<<<<<<<<<<<<<<<<<<      GUI Class      >>>>>>>>>>>>>>>>>>>>>>>>>>

class PlistGui:
    def __init__(self, root):
        self.root = root
        script_name = os.path.basename(sys.argv[0])
        self.root.title(f"{script_name} {version}")
        self.root.geometry("700x550")
        
        # Set Vista theme
        self.style = ttk.Style()
        try:
            self.style.theme_use('vista')
        except:
            pass # Fallback to default if vista is not available

        self.setup_gui()
        self.set_defaults()

    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title/Description
        ttk.Label(main_frame, text=f"{description.strip()}", font=("Helvetica", 10, "italic"), wraplength=650).pack(pady=5)

        # Input Folder
        input_frame = ttk.LabelFrame(main_frame, text="Input Folder (containing .plist files)", padding="5")
        input_frame.pack(fill=tk.X, pady=5)
        
        self.input_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).pack(side=tk.RIGHT)

        # Output File
        output_frame = ttk.LabelFrame(main_frame, text="Output Excel File", padding="5")
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_frame, text="Browse", command=self.browse_output).pack(side=tk.RIGHT)

        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)

        # Status/Log Window
        self.status_box = ScrolledText(main_frame, height=15, state='disabled')
        self.status_box.pack(fill=tk.BOTH, expand=True, pady=5)

        # Start Button
        self.btn_run = ttk.Button(main_frame, text="Start", command=self.start_processing)
        self.btn_run.pack(pady=10)

    def set_defaults(self):
        self.input_var.set("logs_plist")
        self.output_var.set("output_plists.xlsx")

    def browse_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_var.set(folder)

    def browse_output(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.output_var.set(file)

    def log(self, message):
        self.status_box.config(state='normal')
        self.status_box.insert(tk.END, message + "\n")
        self.status_box.see(tk.END)
        self.status_box.config(state='disabled')
        self.root.update_idletasks()

    def start_processing(self):
        input_folder = self.input_var.get()
        output_file = self.output_var.get()

        if not os.path.exists(input_folder):
            messagebox.showerror("Error", f"Input folder '{input_folder}' does not exist.")
            return

        self.btn_run.config(state='disabled')
        self.progress_var.set(0)
        self.status_box.config(state='normal')
        self.status_box.delete(1.0, tk.END)
        self.status_box.config(state='disabled')
        
        self.log(f"Starting processing folder: {os.path.abspath(input_folder)}")

        # Start processing in a new thread
        thread = threading.Thread(target=self.run_process, args=(input_folder, output_file))
        thread.daemon = True
        thread.start()

    def run_process(self, input_folder, output_file):
        try:
            process_plists(input_folder, output_file, self.log, self.update_progress)
            self.log(f"Output File: {output_file}")
            self.log("\nDone.")
        except Exception as e:
            self.log(f"Critical Error: {e}")
        finally:
            self.btn_run.config(state='normal')

    def update_progress(self, current, total):
        if total > 0:
            percent = (current / total) * 100
            self.progress_var.set(percent)
        else:
            self.progress_var.set(100)


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu/Logic     >>>>>>>>>>>>>>>>>>>>>>>>>>

def process_plists(logs_folder, output_file, status_callback=None, progress_callback=None):
    """
    Core logic to parse plist files and save to Excel.
    """
    if not os.path.exists(logs_folder):
        msg = f"Directory {logs_folder} does not exist."
        if status_callback: status_callback(msg)
        print(msg)
        return False

    # Create workbook and worksheet
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet('Plists')

    # Headers
    headers = ['Key', 'Value', 'File Name', 'Creation Time', 'Access Time', 'Modified Time']
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Formatting
    worksheet.freeze_panes(1, 1)
    worksheet.set_column(0, 0, 20)
    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 2, 25)
    worksheet.set_column(3, 5, 25)

    row_num = 1
    
    files = [f for f in os.listdir(logs_folder) if f.lower().endswith('.plist')]
    total_files = len(files)

    msg = f'Parsing {total_files} plist files in the {logs_folder} folder...'
    if status_callback: status_callback(msg)
    print(msg)

    for i, file_name in enumerate(files, 1):
        file_path = os.path.join(logs_folder, file_name)
        if status_callback: status_callback(f"Processing: {file_name}")
        try:
            with open(file_path, 'rb') as f:
                plist_data = plistlib.load(f)

            creation_time_utc = datetime.utcfromtimestamp(os.path.getctime(file_path))
            access_time_utc = datetime.utcfromtimestamp(os.path.getatime(file_path))
            modified_time_utc = datetime.utcfromtimestamp(os.path.getmtime(file_path))

            if isinstance(plist_data, dict):
                for key, value in plist_data.items():
                    worksheet.write(row_num, 0, key)
                    worksheet.write(row_num, 1, str(value))
                    worksheet.write(row_num, 2, file_name)
                    worksheet.write(row_num, 3, creation_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                    worksheet.write(row_num, 4, access_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                    worksheet.write(row_num, 5, modified_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                    row_num += 1
            elif isinstance(plist_data, list):
                for index, value in enumerate(plist_data):
                    worksheet.write(row_num, 0, f'Index {index}')
                    worksheet.write(row_num, 1, str(value))
                    worksheet.write(row_num, 2, file_name)
                    worksheet.write(row_num, 3, creation_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                    worksheet.write(row_num, 4, access_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                    worksheet.write(row_num, 5, modified_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                    row_num += 1
            else:
                worksheet.write(row_num, 0, 'Root Element')
                worksheet.write(row_num, 1, str(plist_data))
                worksheet.write(row_num, 2, file_name)
                worksheet.write(row_num, 3, creation_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                worksheet.write(row_num, 4, access_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                worksheet.write(row_num, 5, modified_time_utc.strftime('%Y-%m-%d %H:%M:%S'))
                row_num += 1

        except Exception as e:
            err_msg = f"Error processing file {file_name}: {e}"
            if status_callback: status_callback(err_msg)
            print(err_msg)
        
        if progress_callback:
            progress_callback(i, total_files)

    workbook.close()
    msg = f"Output written to {output_file}"
    if status_callback: status_callback(msg)
    print(msg)
    return True


def main():
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("-i", "--input", help="Input folder containing .plist files", required=False)
    parser.add_argument("-o", "--output", help="Output Excel file", required=False)

    args = parser.parse_args()

    # If no arguments are provided, launch GUI
    if len(sys.argv) == 1:
        root = tk.Tk()
        gui = PlistGui(root)
        root.mainloop()
        return

    # CLI Mode
    logs_folder = args.input if args.input else 'logs_plist'
    output_file = args.output if args.output else 'output_plists.xlsx'

    process_plists(logs_folder, output_file)

if __name__ == "__main__":
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
test plist files named with upper case extensions


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
stick all of your *.plist files into the logs_plist folder, run the script
bobs your uncle

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Copyright (C) 2024 LincolnLandForensics
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License version 2, as published by the
# Free Software Foundation
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details (http://www.gnu.org/licenses/gpl.txt).


# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>