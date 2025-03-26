#!/usr/bin/env python3
# coding: utf-8
"""
Read a case sheet and re-organize the order of the headers
just change the order of headers on line 119
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import sys
import argparse
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill

# Color support for Windows 11+
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
if sys.version_info > (3, 7, 9) and os.name == "nt":
    from colorama import Fore, Back, Style
    print(Back.BLACK)
    color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
    color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "Read a case sheet and re-organize the order of the headers"
version = '0.1.2'


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>



def main():
    global row
    row = 0  # defines arguments
    # Row = 1  # defines arguments   # if you want to add headers 
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-r', '--read', help='read xlsx', required=False, action='store_true')

    args = parser.parse_args()


    global input_file
    input_file = args.input if args.input else "input_case.xlsx"

    global outuput_xlsx
    outuput_xlsx = args.output if args.output else "output_test.xlsx"


    if args.read:
        # create_xlsx()

        file_exists = os.path.exists(input_file)
        if file_exists == True:
            msg_blurb = (f'Reading {input_file}')
            msg_blurb_square(msg_blurb, color_green)    
            
            data = read_xlsx(input_file)
            # workbook.save(outuput_xlsx)
            # write_xlsx(data)
            write_xlsx(data, input_file)
            # workbook.close()
            msg_blurb = (f'Writing to {outuput_xlsx}')
            msg_blurb_square(msg_blurb, color_green)            

        else:
            msg_blurb = (f'{input_file} does not exist')
            msg_blurb_square(msg_blurb, color_red)      
            exit()

    else:
        usage()
    
    return 0


def msg_blurb_square(msg, color):
    border = f"+{'-' * (len(msg) + 2)}+"
    print(f"{color}{border}\n| {msg} |\n{border}{color_reset}")

def read_xlsx(file_path):
    """Read data from an XLSX file and return a list of dictionaries."""
    wb = load_workbook(file_path)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    data = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        row_data = dict(zip(headers, row))
        row_data["zero"] = "bobs your uncle" if row_data.get("zero") == "test" else row_data.get("zero")
        row_data["one"] = "avocadoes rule" if row_data.get("one") == "TeslaSucks" else row_data.get("one")
        data.append(row_data)
    
    wb.close()
    return data


def write_xlsx(data, file_path):
    '''
    The write_xlsx() function receives the processed data as a list of 
    dictionaries and writes it to a new Excel file using openpyxl. 
    It defines the column headers, sets column widths, and then iterates 
    through each row of data, writing it into the Excel worksheet.
    '''

    global workbook
    workbook = Workbook()
    global worksheet
    worksheet = workbook.active

    worksheet.title = 'Cases'
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'B2'  # Freeze cells
    worksheet.selection = 'B2'


    headers = ["caseNumber", "exhibit", "caseName", "subjectBusinessName", "caseType"
    , "caseAgent", "forensicExaminer", "reportStatus", "notes", "summary", "tempNotes"
    , "exhibitType", "makeModel", "serial", "OS", "hostname", "userName", "userPwd"
    , "email", "emailPwd", "ip", "phoneNumber", "phoneIMEI", "phone2", "phoneIMEI2"
    , "mobileCarrier", "biosTime", "currentTime", "timezone", "shutdownMethod"
    , "shutdownTime", "seizureAddress", "seizureRoom", "dateSeized", "seizedBy"
    , "seizureStatus", "dateReceived", "receivedBy", "removalDate", "removalStaff"
    , "reasonForRemoval", "inventoryDate", "storageLocation", "status", "imagingTool"
    , "imagingType", "imageMD5", "imageSHA256", "imageSHA1", "verifyHash", "writeBlocker"
    , "imagingStarted", "imagingFinished", "storageType", "storageMakeModel"
    , "storageSerial", "storageSize", "evidenceDataSize", "analysisTool"
    , "analysisTool2", "exportLocation", "exportedEvidence", "qrCode", "operation"
    , "vaultCaseNumber", "vaultTotal", "caseNumberOrig", "Action", "priority", "temp"]

    # Write headers to the first row
    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [0, 2, 3, 4, 5, 6, 7]:  # Indices of columns A, C, D, E, F, G, H
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            cell.fill = fill

    # Excel column width
    worksheet.column_dimensions['A'].width = 15# zero
    worksheet.column_dimensions['B'].width = 7# one
    worksheet.column_dimensions['C'].width = 16# two
    worksheet.column_dimensions['D'].width = 20# three

    for row_index, row_data in enumerate(data):
        # print(f'{color_purple}Processing row: {row_data}{color_reset}')  # Debugging output
        # print("row_index = %s" %(row_index))
        # print(f'{color_red}row = {row_index+1}{color_reset}')        

        for col_index, col_name in enumerate(headers):
            cell_data = row_data.get(col_name)
            try:
                worksheet.cell(row=row_index+2, column=col_index+1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error printing line: {str(e)}{color_reset}")

    workbook.save(outuput_xlsx)
    
def usage():
    print(f"Usage: {sys.argv[0]} -r [-I input_case.xlsx] [-O output.xlsx]")
    print("Example:")
    print(f"   python {sys.argv[0]} -r")
    print(f"   python {sys.argv[0]} -r -I input_case.xlsx -O custom_output.xlsx")

if __name__ == '__main__':
    main()
