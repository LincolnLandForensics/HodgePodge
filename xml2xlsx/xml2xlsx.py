#!/usr/bin/env python3
# coding: utf-8
'''
read xml files in an xml folder
convert them to xlsx
'''

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook
import argparse
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import sys

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description2 = "convert xml files in a folder to xlsx"
version = '0.2.6'

# <<<<<<<<<<<<<<<<<<<<<<<<<<      GUI Class      >>>>>>>>>>>>>>>>>>>>>>>>>>

class XmlGui:
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
        ttk.Label(main_frame, text=f"{description2}", font=("Helvetica", 10, "italic"), wraplength=650).pack(pady=5)

        # Input Folder
        input_frame = ttk.LabelFrame(main_frame, text="Input Folder (containing .xml files)", padding="5")
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
        self.input_var.set("xml")
        self.output_var.set("output_xml.xlsx")

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
            process_xml_folder(input_folder, output_file, self.log, self.update_progress)
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
def parse_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    data = []
    
    # Extract column headers from the first entry
    headers = [elem.tag for elem in root[0].iter() if elem.tag != root[0].tag]
    
    # Extract data
    for child in root:
        row = [subchild.text for subchild in child.iter() if subchild.tag != child.tag]
        data.append(row)
    
    return headers, data

def process_xml_folder(xml_folder, output_xlsx, status_callback=None, progress_callback=None):
    """
    Core logic to parse XML files in a folder and save to Excel.
    """
    if not os.path.exists(xml_folder) or not os.path.isdir(xml_folder):
        msg = f"The {xml_folder} folder does not exist."
        if status_callback: status_callback(msg)
        print(f"\n\t{msg}")
        return False
    
    # Check if there are XML files in the xml folder
    xml_files = [f for f in os.listdir(xml_folder) if f.lower().endswith('.xml')]
    total_files = len(xml_files)
    if not xml_files:
        msg = f"No XML files found in the {xml_folder} folder."
        if status_callback: status_callback(msg)
        print(f"\n\t{msg}")
        return False

    msg = f'Reading {total_files} XML files out of the {xml_folder} folder'
    if status_callback: status_callback(msg)
    print(f'\n{msg}')

    # Create a workbook
    wb = Workbook()
    ws = wb.active
    
    # Iterate through XML files in the xml folder
    header_written = False
    for i, filename in enumerate(xml_files, 1):
        msg = f"Parsing {filename} ..."
        if status_callback: status_callback(msg)
        print(f"\t{i}. {msg}")
        
        xml_file = os.path.join(xml_folder, filename)
        try:
            headers, data = parse_xml(xml_file)
            
            if not header_written:
                ws.append(headers)
                header_written = True
            
            for row in data:
                ws.append(row)
        except Exception as e:
            err_msg = f"Error processing {filename}: {e}"
            if status_callback: status_callback(err_msg)
            print(f"\t   {err_msg}")
        
        if progress_callback:
            progress_callback(i, total_files)
    
    # Save the workbook
    wb.save(output_xlsx)
    msg = f'Saving to {output_xlsx}'
    if status_callback: status_callback(msg)
    print(f'\n{msg}')
    return True

def main():
    parser = argparse.ArgumentParser(description=description2)
    parser.add_argument("-i", "--input", help="Input folder containing XML files", required=False)
    parser.add_argument("-o", "--output", help="Output Excel file", required=False)

    args = parser.parse_args()

    # If no arguments are provided, launch GUI
    if len(sys.argv) == 1:
        root = tk.Tk()
        gui = XmlGui(root)
        root.mainloop()
        return

    # CLI Mode
    xml_folder = args.input if args.input else 'xml'
    output_xlsx = args.output if args.output else 'output.xlsx'

    process_xml_folder(xml_folder, output_xlsx)

if __name__ == "__main__":
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
0.2.0 - reads xml and convert them to xlsx
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
"""

'''
Copyright (c) 2024 LincolnLandForensics

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
'''

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
