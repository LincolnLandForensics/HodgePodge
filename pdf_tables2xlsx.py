#!/usr/bin/python
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import sys
import datetime
from tabula import read_pdf # pip install tabula-py
import pandas as pd

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "Export tables from PDF's to xlsx"
version = '0.0.2'


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def extract_tables(filename):
    """
    Extract tables from a PDF file using tabula.

    Parameters:
    - filename (str): The path to the PDF file.

    Returns:
    - list: A list of DataFrames, each representing a table.
    """
    # Use tabula to extract tables
    tables = read_pdf(filename, pages='all', multiple_tables=True)

    # Add a new column with the filename to each table
    for i, table in enumerate(tables):
        table['Filename'] = filename

    return tables

def process_pdfs_in_folder(folder_path, output_excel):
    """
    Process all PDF files in a folder, extract tables, and save them to an Excel file.

    Parameters:
    - folder_path (str): The path to the folder containing PDF files (optional).
    - output_excel (str): The path to the output Excel file.
    """
    count = 1
    keyword = 'Post Date'
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        workbook = writer.book  # Get the workbook object
        
        if not os.path.exists(folder_path):
            # os.makedirs(folder_path) # create the path
            print(f"The specified folder path '{folder_path}' does not exist.")
            exit(1)  # Use a non-zero exit status to indicate an error
            
        for filename in os.listdir(folder_path):
            if filename.lower().endswith('.pdf'):
                full_path = os.path.join(folder_path, filename)

                print(f'\n\n------------------ {filename} ------------------')

                # Extract and export tables
                tables = extract_tables(full_path)

                # Save each table to the same Excel file
                for i, table in enumerate(tables):
                    sheet_name = str(count)  # Convert count to string
                    
                    # Check if the keyword is present in any cell of the table
                    if any(keyword in cell for col in table.columns for cell in table[col].astype(str)):
                        print(f'     Post Date in Table {count}')

                    table.to_excel(writer, sheet_name=sheet_name, index=False)
                    count += 1

if __name__ == "__main__":

    # Check if there is a folder path after the filename
    if len(sys.argv) > 1:
        pdf_folder_path = sys.argv[1]
        print(f' reading pdfs from {pdf_folder_path}')
    else:
        pdf_folder_path = 'pdfs'  # Change this to the path of your pdf folder
        print(f' reading pdfs from default: {pdf_folder_path}')

    # Get the current date and time
    current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # Append timestamp to the output file name
    output_excel = f"combined_tables_{current_datetime}.xlsx"
    print(f'Outputing tables to {output_excel}')
    
    process_pdfs_in_folder(pdf_folder_path, output_excel)

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

0.1.0 - read pdf, output to xlsx
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
can specify a different folder

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Copyright (C) 2023 LincolnLandForensics
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
