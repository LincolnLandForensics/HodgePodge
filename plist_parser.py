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
This script first creates an Excel file named "output.xlsx" and then iterates
through all files in the "logs" folder that have the ".plist" file extension. 
It then opens each of those files, loads their plist data and writes the key-value
 pairs to the Excel file.
Please note that the script assumes that "logs" folder is located in the same 
directory as the script.
'''
version = '1.0.1'


# Create a new Excel file
workbook = xlsxwriter.Workbook('output_plists.xlsx')
worksheet = workbook.add_worksheet()

# Set column headers
worksheet.write(0, 0, 'Key')
worksheet.write(0, 1, 'Value')
worksheet.write(0, 2, 'File Name')
worksheet.write(0, 3, 'creation_time')
worksheet.write(0, 4, 'access_time')
worksheet.write(0, 5, 'modified_time')

# freeze cells at 1,1
worksheet.freeze_panes(1, 1)
# set column widths
worksheet.set_column(0, 0, 20)
worksheet.set_column(1, 1, 30)
worksheet.set_column(2, 2, 25)
worksheet.set_column(3, 3, 25)
worksheet.set_column(4, 4, 25)
worksheet.set_column(5, 5, 25)

worksheet.name = 'Plists'

# Iterate through all plist files in the "logs" folder
row = 1
for file_name in os.listdir('logs'):
    if file_name.endswith('.plist'):
        file_path = os.path.join('logs', file_name)
        with open(file_path, 'rb') as f:
            plist_data = plistlib.load(f)
        
        # get creation, access, modified times of the .plist
        file_info = os.stat(file_path)

        # utc date
        creation_time = os.path.getctime(file_path)
        creation_time_utc = datetime.utcfromtimestamp(creation_time)

        modified_time = os.path.getmtime(file_path)
        modified_time_utc = datetime.utcfromtimestamp(modified_time)
            
        # Write the key-value pairs to the Excel file
        for key, value in plist_data.items():
            worksheet.write(row, 0, key)
            worksheet.write(row, 1, value)
            worksheet.write(row, 2, file_name)
            worksheet.write(row, 3, creation_time_utc)
            # worksheet.write(row, 4, access_time)
            worksheet.write(row, 5, modified_time_utc)            
            row += 1

# Save and close the Excel file
workbook.close()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
test plist files named with upper case extensions


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
stick all of your *.plist files into the logs folder, run the script
bobs your uncle

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