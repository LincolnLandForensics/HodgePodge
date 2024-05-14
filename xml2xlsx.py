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
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description2 = "convert xml files in a folder to xlsx"
version = '0.2.6'

# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>
def parse_xml(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    data = []
    for child in root:
        row = []
        for subchild in child:
            row.append(subchild.text)
        data.append(row)
    return data

def main():
    count = 1
    # Check if the xml folder exists
    output_xlsx = 'output.xlsx'
    xml_folder = 'xml'
    if not os.path.exists(xml_folder) or not os.path.isdir(xml_folder):
        print(f"\n\tThe {xml_folder} folder does not exist.")
        return
    
    # Check if there are XML files in the xml folder
    xml_files = [f for f in os.listdir(xml_folder) if f.endswith('.xml')]
    if not xml_files:
        print(f"\n\tNo XML files found in the {xml_folder} folder.")
        return
    print(f'\nReading XML files out of the {xml_folder} folder')
    # Create a workbook
    wb = Workbook()
    ws = wb.active
    
    # Iterate through XML files in the xml folder
    header_written = False
    for filename in xml_files:
        print(f"\t{count}. Parsing {filename} ...")
        xml_file = os.path.join(xml_folder, filename)
        data = parse_xml(xml_file)
        if not header_written:
            # Write header from the first sheet
            header = [subchild.tag for subchild in ET.parse(xml_file).getroot()[0]]
            ws.append(header)
            header_written = True
        for row in data:
            ws.append(row)

        count +=1
    # Save the workbook
    wb.save(output_xlsx)
    print(f'\nSaving to {output_xlsx}')
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

