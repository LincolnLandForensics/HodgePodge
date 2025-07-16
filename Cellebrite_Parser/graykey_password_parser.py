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
version = '1.4.1'

headers = [
    "URL", "Username", "Password", "Notes", "Case", "Exhibit", "protocol",
    "fileType", "Encryption", "Complexity", "Hash", "Pwd", "PWDUMPFormat", "Length",
    "Email", "IP"
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
        color_green = Fore.GREEN
        color_reset = Style.RESET_ALL

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-b', '--blank', help='create blank sheet', required=False, action='store_true')
    # parser.add_argument('-p', '--passwords', help='passwords module', required=False, action='store_true')
    parser.add_argument('-c', '--convert', help='convert GrayKey passwords to Excel', required=False, action='store_true')

    args = parser.parse_args()

    global input_file, output_xlsx, Case, Exhibit
    input_file = args.input if args.input else "sample_passwords.txt"

    if args.convert:
        Case = input("Enter Case: ").strip()
        Exhibit = input("Enter Exhibit: ").strip()
        output_xlsx = args.output if args.output else (f"passwords_{Case}_Ex_{Exhibit}.xlsx")
        read_pwords()
    elif args.blank:
        output_xlsx = 'blank_password_sheet.xlsx'
        write_xlsx([], [])
        sys.exit(0)
    else:
        usage()

# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def complexinator(password):
    if not password:
        return "blank"

    length_ok = len(password) >= 8
    has_upper = any(c.isupper() for c in password)
    has_lower = any(c.islower() for c in password)
    has_digit = any(c.isdigit() for c in password)
    has_special = any(not c.isalnum() for c in password)
    complexity_criteria = sum([has_upper, has_lower, has_digit, has_special])

    return "complex" if length_ok and complexity_criteria >= 3 else "weak"

def message_square(message, color):
    horizontal_line = f"+{'-' * (len(message) + 2)}+"
    print(color + horizontal_line)
    print(f"| {message} |")
    print(horizontal_line + f'{color_reset}')

def read_pwords():
    if not os.path.isfile(input_file):
        print(f"Error: Input file '{input_file}' does not exist.")
        sys.exit(1)
    else:
        message_square(f'Reading {input_file}', color_green)

    data, uniq = [], set()
    fileType = input_file
    pattern = re.compile(r'^\d{9}\.\d{6}$')
    known_bad_passwords = {
        'false', 'true', 'US', 'Secret', '0', '1', 'treeup', 'mobile', '""',
        'ATM,CHK', 'myPSKkey', 'PERSONAL', 'Registered', 'stayPaired', 'POH',
        'PER', 'PR', '10', '09EA', 'ATM+CHK', 'kcKeepDeviceTrusted', 'dummy_value',
        '2', '4', '[]', '{}', 'YES', 'prod', 'reinstall_value', 'IS_LATEST_KEY_V2',
        'comcast-business', 'VAL_KeychainCanaryPassword', 'TwitterKeychainCanaryPassword'
    }

    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()
        # content = content.replace("IP: N/A", "----------")  # for intel.veraxity.org output
        content = re.sub(r'^(IP: .*)', r'\1----------', content, flags=re.MULTILINE)

        entries = content.split("----------")

        for block in entries:
            entry = {
                "URL": '', "Username": '', "Password": '', "Notes": block.strip(),
                "Case": Case, "Exhibit": Exhibit, "protocol": '', "fileType": fileType,
                "Encryption": '', "Complexity": '', "Hash": '', "Pwd": '',
                "PWDUMPFormat": '', "Length": '', "Email": '', "IP": ''
            }

            for line in block.strip().splitlines():
                line = line.strip()
                if line.startswith("Account:"):
                    entry["Username"] = line.split("Account:", 1)[1].strip()
                elif line.startswith("srvr: "):
                    entry["URL"] = line.split("srvr: ", 1)[1].strip()
                elif line.startswith("ptcl: "):
                    protocol = line.split("ptcl: ", 1)[1].strip()
                    if protocol != "0":
                        entry["protocol"] = protocol
                elif line.startswith("Service: "):
                    entry["URL"] = line.split("Service: ", 1)[1].strip()
                elif line.startswith("Item value:"):
                    pwd = line.replace("Item value:", '').strip()
                    if pwd in known_bad_passwords or \
                       pwd.endswith('.com') or \
                       pwd.startswith('[{') or \
                       pwd.startswith('{"') or \
                       pwd.startswith('|DYN') or \
                       pwd.startswith('us-east') or \
                       pwd.startswith('http') or \
                       pwd.endswith("=") or \
                       pwd.endswith("~~") or \
                       "whatsapp.net" in pwd or \
                       len(pwd) > 33 or pattern.match(pwd):
                        entry["Hash"] = pwd
                    else:
                        entry["Password"] = pwd
                elif line.startswith("Username: "):
                    entry["Username"] = line.split("Username: ", 1)[1].strip().replace('N/A','')
                elif line.startswith("Email: "):
                    entry["Email"] = line.split("Email: ", 1)[1].strip().replace('N/A','')
                elif line.startswith("Password: "):
                    entry["Password"] = line.split("Password: ", 1)[1].strip().replace('N/A','')
                elif line.startswith("Origin: "):
                    entry["URL"] = line.split("Origin: ", 1)[1].strip().replace('N/A','') 
                    entry["fileType"] = "intel.veraxity.org"
                elif line.startswith("IP: "):
                    entry["IP"] = line.split("IP: ", 1)[1].strip().replace('N/A','')

            if entry["URL"] == "AirPort":
                entry["protocol"] = "AirPort"
            elif "com.apple.airplay" in entry["URL"]:
                entry["protocol"] = "AirPlay"
            elif entry["URL"] == "GuidedAccess":
                entry["URL"] = "_phone pin code ***"                
                entry["Username"] = "" 
                
            if any(k in entry["Username"].lower() for k in ["apikey", "token", "sessionkey"]) or \
               entry["Username"].startswith('com.') or entry["Username"] in ["UUID", "secretKey", "acquiredPackages"]:
                entry["Hash"] = entry["Password"]
                entry["Password"] = ''
                entry["Username"] = ''

            if entry["Password"]:
                entry["Length"] = len(entry["Password"])
                entry["Complexity"] = complexinator(entry["Password"])
                if entry["Password"] not in uniq:
                    uniq.add(entry["Password"])

            data.append(entry)

    data = sorted(data, key=lambda x: (x["Length"] if isinstance(x["Length"], int) else 100))
    write_xlsx(data, sorted(uniq, key=len))




def write_xlsx(data, uniq_list):
    message_square(f'Writing {output_xlsx}', color_green)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Passwords'
    worksheet.freeze_panes = 'B2'
    worksheet.selection = 'B2'

    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if header in ["Username", "Password", "Exhibit", "Case", "Notes"]:
            cell.fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        elif header in ["URL", "Length", "Complexity"]:
            cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    col_widths = [20, 20, 20, 35, 7, 6, 10, 20, 8, 12, 4, 17, 5]
    for i, width in enumerate(col_widths):
        worksheet.column_dimensions[chr(65+i)].width = width


    for row_index, row_data in enumerate(data):
        for col_index, col_name in enumerate(headers):
            worksheet.cell(row=row_index + 2, column=col_index + 1).value = row_data.get(col_name, '')

    # Create second sheet with unique passwords
    # Uniq Sheet
    uniq_sheet = workbook.create_sheet(title="uniq")
    uniq_sheet.freeze_panes = 'B2' 
    uniq_sheet['A1'] = 'Unique Passwords (Sorted by Length)'
    uniq_sheet.column_dimensions['A'].width = 40        
    # uniq_sheet.append(["Password"])
    for password in uniq_list:
        uniq_sheet.append([password])

    workbook.save(output_xlsx)

def usage():
    file = os.path.basename(sys.argv[0])
    print("\nDescription: " + description)
    print(f"{file} Version: {version} by {author}")
    print("\nExample:")
    print(f"\t{file} -c -I sample_passwords.txt")
    print(f"\t{file} -c -I sample_passwords.txt -O passwords_sample_.xlsx")
    print(f"\t{file} -b -O blank_sheet.xlsx")
    print(f"\t{file} -v -I input.txt")
    
if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
0.5.0 - created intel.veraxity.org parser
0.4.0 - create a seperate sheet for uniq passwords
0.2.2 - working prototype
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
also create an intel sheet



"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""



"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
