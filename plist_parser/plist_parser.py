#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>
import os
import plistlib
import xlsxwriter
from datetime import datetime

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

# Constants
logs_folder = 'logs_plist'
output_file = 'output_plists.xlsx'

if not os.path.exists(logs_folder):
    print(f"Directory {logs_folder} does not exist.")
    exit(1)

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

row = 1

print(f'Parsing plist files in the {logs_folder} folder...')

for file_name in os.listdir(logs_folder):
    if file_name.endswith('.plist'):
        file_path = os.path.join(logs_folder, file_name)
        try:
            with open(file_path, 'rb') as f:
                plist_data = plistlib.load(f)

            file_info = os.stat(file_path)
            creation_time_utc = datetime.utcfromtimestamp(os.path.getctime(file_path))
            access_time_utc = datetime.utcfromtimestamp(os.path.getatime(file_path))
            modified_time_utc = datetime.utcfromtimestamp(os.path.getmtime(file_path))

            if isinstance(plist_data, dict):
                for key, value in plist_data.items():
                    worksheet.write(row, 0, key)
                    worksheet.write(row, 1, str(value))
                    worksheet.write(row, 2, file_name)
                    worksheet.write(row, 3, creation_time_utc)
                    worksheet.write(row, 4, access_time_utc)
                    worksheet.write(row, 5, modified_time_utc)
                    row += 1
            elif isinstance(plist_data, list):
                for index, value in enumerate(plist_data):
                    worksheet.write(row, 0, f'Index {index}')
                    worksheet.write(row, 1, str(value))
                    worksheet.write(row, 2, file_name)
                    worksheet.write(row, 3, creation_time_utc)
                    worksheet.write(row, 4, access_time_utc)
                    worksheet.write(row, 5, modified_time_utc)
                    row += 1
            else:
                worksheet.write(row, 0, 'Root Element')
                worksheet.write(row, 1, str(plist_data))
                worksheet.write(row, 2, file_name)
                worksheet.write(row, 3, creation_time_utc)
                worksheet.write(row, 4, access_time_utc)
                worksheet.write(row, 5, modified_time_utc)
                row += 1

        except Exception as e:
            print(f"Error processing file {file_name}: {e}")


workbook.close()
print(f"Output written to {output_file}")

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