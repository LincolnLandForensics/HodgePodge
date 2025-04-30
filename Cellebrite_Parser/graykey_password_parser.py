#!/usr/bin/python
# coding: utf-8

import os
import re
import sys
import argparse
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "convert graykey password file to xlsx"
version = '1.2.3'

headers = [
    "URL", "Username", "Password", "Notes", "Case", "Exhibit", "protocol",
    "fileType", "Encyption", "Complexity", "Hash", "Pwd", "PWDUMPFormat", "Length"
]

color_green = ''
color_red = ''
color_reset = ''

if sys.version_info > (3, 7, 9) and os.name == "nt":
    version_info = os.sys.getwindowsversion()
    major_version = version_info.major
    build_version = version_info.build
    if major_version >= 10 and build_version >= 22000:
        import colorama
        from colorama import Fore, Back, Style
        print(f'{Back.BLACK}')
        color_red = Fore.RED
        color_yellow = Fore.YELLOW
        color_green = Fore.GREEN
        color_blue = Fore.BLUE
        color_purple = Fore.MAGENTA
        color_reset = Style.RESET_ALL

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    global Row
    Row = 1
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-b', '--blank', help='create blank sheet', required=False, action='store_true')
    parser.add_argument('-p', '--passwords', help='passwords module', required=False, action='store_true')
    parser.add_argument('-c', '--convert', help='convert GrayKey passwords to Excel', required=False, action='store_true')

    args = parser.parse_args()

    global input_file
    input_file = args.input if args.input else "sample_passwords.txt"

    global output_xlsx
    global Case
    global Exhibit

    if args.convert:
        data = []
    
        Case = input("Enter Case: ").strip()
        Exhibit = input("Enter Exhibit: ").strip()        
        output_xlsx = args.output if args.output else (f"passwords_{Case}_Ex_{Exhibit}.xlsx")
        
        
        
        read_pwords()
    elif args.blank:
        output_xlsx = 'blank_password_sheet.xlsx'
        data = []
        write_xlsx(data)
        sys.exit(0)
    else:
        usage()

    return 0

# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def complexinator(Password):
    '''
    create a function that evaluates the complexity of a password:

    must be at least 8 characters 
    and 
    have 3 or more of the following 4 character types:
    - Uppercase
    - Lowercase
    - Numeric
    - Special character

    if it meets the complexity requirement, return "complex"
    if blank, return "blank"
    else return "weak"
    '''
    if not Password:
        return "blank"

    length_ok = len(Password) >= 8

    has_upper = any(c.isupper() for c in Password)
    has_lower = any(c.islower() for c in Password)
    has_digit = any(c.isdigit() for c in Password)
    has_special = any(not c.isalnum() for c in Password)

    complexity_criteria = sum([has_upper, has_lower, has_digit, has_special])

    if length_ok and complexity_criteria >= 3:
        return "complex"
    else:
        return "weak"

def message_square(message, color):
    horizontal_line = f"+{'-' * (len(message) + 2)}+"
    empty_line = f"| {' ' * (len(message))} |"
    print(color + horizontal_line)
    print(empty_line)
    print(f"| {message} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')

def read_pwords():
    if not os.path.isfile(input_file):
        print(f"Error: Input file '{input_file}' does not exist.")
        sys.exit(1)
    else:
        message_square(f'reading {input_file}', color_green)

    data = []
    with open(input_file, 'r', encoding='utf-8') as f:
        fileType = input_file
        content = f.read()
        entries = content.split("----------")
        uniq = set()
        output = []
        pattern = re.compile(r'^\d{9}\.\d{6}$')


        for block in entries:
            
            (URL, Username, Password, Notes, Encyption, Complexity) = ('', '', '', block ,'', '')
            (Length, protocol, Hash) = ('', '', '')

            lines = block.strip().splitlines()
            entry = {}
            for line in lines:
                line = line.strip()
                if line.startswith("Account:"):
                    Username = line.split("Account:", 1)[1].strip()
                elif line.startswith("srvr: "):
                    URL = line.split("srvr: ", 1)[1].strip()
                elif line.startswith("ptcl: "):
                    protocol = line.split("ptcl: ", 1)[1].strip() 
                    if protocol == "0":
                       protocol = ''
                elif line.startswith("Service: "):
                    URL = line.split("Service: ", 1)[1].strip()                    
                elif line.strip().startswith("Item value: {"):
                    line = line.replace('Item value: ','')
                    '''
                    extract email, password etc, out of data
                    
                    '''

                elif line.strip().startswith("Item value:"):
                    line = line.replace('Item value: ','')
                    Password = line.strip('')
                    if Password == 'false' or Password == 'true' or Password == 'US' or Password == 'Secret':
                       Password = ''
                    elif Password == '0' or Password == '1' or Password == 'treeup' or Password == 'mobile' or Password == '\"\"':
                       Password = ''                       
                    elif Password == 'ATM,CHK' or Password == 'myPSKkey' or Password == 'PERSONAL' or Password == 'Registered' or Password == 'stayPaired':
                       Password = ''                       
                    elif Password == 'POH' or Password == 'PER' or Password == 'PR' or Password == '10' or Password == '09EA':
                       Password = '' 
                    elif Password == 'ATM+CHK' or Password == 'kcKeepDeviceTrusted' or Password.endswith('.com'):
                       Password = ''                       
                    elif Password == 'dummy_value' or Password == 'myPSKkey' or Password == 'PERSONAL' or Password == 'Registered' or Password == 'stayPaired':
                       Password = ''                       
                    elif Password.startswith('[{') or Password.startswith('|DYN') or Password.startswith('us-east')  or Password.startswith('http'):
                       Password = ''
                    elif "whatsapp.net" in Password:
                        Password = ''
                    elif len(Password) > 33 or pattern.match(Password):
                        Hash = Password
                        Password = ""                    
                    if Password.endswith("=") or Password.endswith("~~"):
                        Hash = Password
                        Password = ""
                        
                    elif Username.startswith('__') or Username == "UUID" or Username == "secretKey" or Username == "acquiredPackages" or Username.startswith('au.'):
                        Username = ''
                        Hash = Password
                        Password = ''

            if URL == "AirPort":
                protocol = "AirPort"
            elif "com.apple.airplay" in URL:
                protocol = "AirPlay"
            
            if "apikey" in Username or "token" in Username.lower() or Username == "_pfo":
                Username = ''
                Hash = Password
                Password = ''
            elif Username.startswith('com.') or "sessionkey" in Username.lower() or Username.lower() == "username":
                Username = ''
                
            Length = len(Password)
            if Length == 0:
               Length = ''
            else:
                Complexity = complexinator(Password)


            if Password not in uniq and len(Password) < 34 and not pattern.match(Password):
                uniq.add(Password)
                output.append(Password)

            if 1==1:
                entry.setdefault("URL", URL)
                entry.setdefault("Username", Username)
                entry.setdefault("Password", Password)
                entry.setdefault("Notes", block.strip())
                entry.setdefault("Case", Case)
                entry.setdefault("Exhibit", Exhibit)
                entry.setdefault("protocol", protocol)
                entry.setdefault("fileType", fileType)
                entry.setdefault("Encyption", Encyption)
                entry.setdefault("Complexity", Complexity)
                entry.setdefault("Hash", Hash)
                entry.setdefault("Length", Length)
                data.append(entry)
    for pwd in sorted(output, key=len):
        print(pwd)

    write_xlsx(data)


def write_xlsx(data):
    message_square(f'Writing {output_xlsx}', color_green)

    global workbook
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Passwords'
    worksheet.freeze_panes = 'B2'
    worksheet.selection = 'B2'

    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if header in ["Username", "Password", "Exhibit", "Case", "Notes"]:
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            cell.fill = fill
        elif header in ["URL", "Length", "Complexity"]:
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.fill = fill

    col_widths = [15, 20, 20, 25, 25, 15, 14, 16, 24, 12, 10, 8, 20]
    for i, width in enumerate(col_widths):
        worksheet.column_dimensions[chr(65+i)].width = width

    for row_index, row_data in enumerate(data):
        for col_index, col_name in enumerate(headers):
            try:
                cell_data = row_data.get(col_name, '')
                worksheet.cell(row=row_index + 2, column=col_index + 1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error printing line: {str(e)}{color_reset}")

    workbook.save(output_xlsx)

def usage():
    file = os.path.basename(sys.argv[0])
    print("\nDescription: " + description)
    print(f"{file} Version: {version} by {author}")
    print("\nExample:")
    print(f"\t{file} -c -I sample_passwords.txt")    
    print(f"\t{file} -c -I sample_passwords.txt -O passwords_sample_.xlsx")
    print(f"\t{file} -b -O blank_sheet.xlsx")

if __name__ == '__main__':
    main()



# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

0.2.2 - working prototype
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
create a seperate sheet for uniq passwords


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""



"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
