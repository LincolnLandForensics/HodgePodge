import os
import re
import sys
import argparse
from typing import List, Dict
from datetime import datetime
from openpyxl import load_workbook

# <<<<<<<<<<<<<<<<<<<<<<<<<<     Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description2 = "turn intel.xlsx into actionable markdown files"
tech = 'LincolnLandForensics'  # change this to your name if you are using Linux
version = '0.2.1'

headers_intel = [
    "query", "ranking", "fullname", "url", "email", "user", "phone",
    "business", "fulladdress", "city", "state", "country", "note", "AKA",
    "DOB", "SEX", "info", "misc", "firstname", "middlename", "lastname",
    "associates", "case", "sosfilenumber", "owner", "president", "sosagent",
    "managers", "Time", "Latitude", "Longitude", "Coordinate",
    "original_file", "Source", "Source file information", "Plate", "VIS", "VIN",
    "VYR", "VMA", "LIC", "LIY", "DLN", "DLS", "content", "referer", "osurl",
    "titleurl", "pagestatus", "ip", "dnsdomain", "Tag", "Icon", "Type"
]

# Colorize section
global color_red
global color_yellow
global color_green
global color_blue
global color_purple
global color_reset
color_red = ''
color_yellow = ''
color_green = ''
color_blue = ''
color_purple = ''
color_reset = ''

if sys.version_info > (3, 7, 9):
    try:
        import colorama
        from colorama import Fore, Style
        colorama.init()
        color_red = Fore.RED
        color_yellow = Fore.YELLOW
        color_green = Fore.GREEN
        color_blue = Fore.BLUE
        color_purple = Fore.MAGENTA
        color_reset = Style.RESET_ALL
    except ImportError:
        print("colorama module not available. Output will not be colorized.")

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    """
    Main function to parse arguments and initiate file conversion.
    """
    parser = argparse.ArgumentParser(description=description2)
    parser.add_argument('-I', '--input', help='Input folder path', required=False)
    parser.add_argument('-O', '--output', help='Output folder path', required=False)
    parser.add_argument('-c', '--contacts', help='read contacts from xlsx file. Create 1 .md file per sheet', required=False, action='store_true')
    parser.add_argument('-x', '--xlsx', help='read xlsx file, create 1 .md file per row', required=False, action='store_true')
    parser.add_argument('--dry-run', help='Validate the input file without generating output', required=False, action='store_true')

    args = parser.parse_args()

    input_xlsx = args.input or "Intel_.xlsx"
    output_folder = args.output or "output_markdown"

    if args.dry_run:
        validate_input_file(input_xlsx)
        return 0

    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
        except Exception as e:
            msg_blurb_square(f"Error creating output folder: {e}", color_red)
            sys.exit(1)

    if args.xlsx:
        data = read_excel(input_xlsx)
        create_markdown_files(data, output_folder, headers_intel)
    elif args.contacts:
        if input_xlsx.lower().endswith(".xlsx"):
            data = read_excel(input_xlsx)
            create_contacts_markdown_files(data, input_xlsx, output_folder)
        else:
            print(f"{input_xlsx} isn't an xlsx file. Try again.")
    else:
        parser.print_help()
        Usage()

    return 0

# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def clean_data(item):
    """
    Cleans the item by removing any unwanted newlines and other characters 
    that may interfere with Markdown tables.

    Returns:
        str: The cleaned cell without newlines or extra spaces.
    """
    item_cleaned = re.sub(r'[\n\r]+', ' ', str(item))  # Replaces newlines with space
    item_cleaned = item_cleaned.strip()  # Remove leading/trailing spaces
    return item_cleaned

def clean_phone(item):
    """
    Cleans the item by removing any unwanted newlines and other characters 
    that may interfere with Markdown tables.

    Returns:
        str: The cleaned cell without newlines or extra spaces.
    """
    item_cleaned = re.sub(r'[\n\r]+-\(\)', ' ', str(item))  # Replaces newlines with space
    item_cleaned = item_cleaned.strip()  # Remove leading/trailing spaces
    return item_cleaned
    

def create_contacts_markdown_files(data, input_xlsx, output_folder):
    """
    Create a markdown file with a table for all rows out of input_xlsx with the following columns:
    fullname, email, user, phone, AKA, Tag, business, fulladdress, case, original_file.

    Before the table, print the name of {input_xlsx} and the sheet name.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    sheet_name = data[0].get("sheet_name", "Unknown")
    output_file = os.path.join(output_folder, f"{sheet_name}_{current_time}.md")

    with open(output_file, "w", encoding="utf-8") as md_file:
        md_file.write(f"# Contact Summary from {input_xlsx}\n\n")
        md_file.write("| fullname | phone | email | user | AKA | Tag | business | fulladdress | case | original_file |\n")
        md_file.write("|----------|-------|------|-------|-----|-----|---------|--------------|------|---------------|\n")

        phone_regex = re.compile(r"^\+?[1-9]\d{1,14}$")  # E.164 format or similar valid phone patterns

        for row in data:
            fullname = clean_data(row.get("fullname", ""))
            email = clean_data(row.get("email", ""))
            user = clean_data(row.get("user", ""))
            phone = clean_phone(str(row.get("phone", "")))
            if phone and phone_regex.match(phone):
                phone = f"[[{phone}]]"
            aka = clean_data(row.get("AKA", ""))
            Tag = clean_data(row.get("Tag", ""))
            if Tag:
                Tag = f"@{Tag}"
            business = clean_data(row.get("business", ""))
            fulladdress = clean_data(row.get("fulladdress", ""))
            case = clean_data(row.get("case", ""))
            original_file = clean_data(row.get("original_file", ""))

            md_file.write(f"| {fullname} | {phone} | {email} | {user}  | {aka} | {Tag} | {business} | {fulladdress} | {case} | {original_file} |\n")

        md_file.write(f"\nSheet Name: {sheet_name}\n")

    msg_blurb_square(f"{sheet_name}....md file created in {output_folder} folder", color_green)




def create_markdown_files(data: List[Dict[str, str]], output_folder: str, headers):
    """Creates markdown files for each row."""
    # Define the headers to use for creating markdown files
    # headers = ['query', 'phone', 'fullname', 'business', 'Tag']  # Adjust as needed

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    for idx, row in enumerate(data, start=1):
        query = str(row.get("query", "_") or "_")   # .replace(" ", "_")
        # query = re.sub(r'[^\w\-_]', '_', query)  # Sanitize filename
        query = re.sub(r'[^\w\-]', ' ', query)  # Sanitize filename

        # Check for existing files and adjust the name if necessary
        base_filename = os.path.join(output_folder, f"{query}.md")
        filename = base_filename
        counter = 1

        while os.path.exists(filename):
            filename = os.path.join(output_folder, f"{query}_{counter}.md")
            counter += 1

        try:
            with open(filename, "w", encoding="utf-8") as md_file:
                for header in headers:
                    value = row.get(header)
                    if value:  # Only include non-blank fields
                        if header in ['phone', 'fullname', 'business']:
                            formatted_value = f"[[{value}]]"
                        elif header == 'Tag':
                            formatted_value = f"@{value}"
                        else:
                            formatted_value = value
                        md_file.write(f"{header}: {formatted_value}\n\n")
        except Exception as e:
            msg_blurb_square(f"Unexpected error: {e}", color_red)
            continue

    msg_blurb_square(f"Markdown file generation complete in {output_folder} folder", color_green)
    



def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')


def read_excel(file_path: str, validate_only: bool = False) -> List[Dict[str, str]]:
    """Reads an Excel file and returns a list of rows as dictionaries."""
    try:
        workbook = load_workbook(file_path, data_only=True)
        sheet = workbook.active
        sheet_name = sheet.title  # Get the sheet name
        msg_blurb_square(f"Reading {sheet_name} sheet in {file_path} ", color_green)
    except FileNotFoundError:
        raise FileNotFoundError(f"Excel file not found: {file_path}")
    except Exception as e:
        raise Exception(f"Error reading file: {e}")

    headers = [clean_data(cell.value) if cell is not None else "" for cell in sheet[1]]  # Read header row
    
    if not headers:
        raise ValueError("The Excel file contains no headers or is empty.")

    # Map Excel headers dynamically to headers_intel
    header_map = {header: headers_intel[idx] for idx, header in enumerate(headers) if idx < len(headers_intel)}

    if len(header_map) < len(headers_intel):
        missing_headers = set(headers_intel) - set(header_map.values())
        raise ValueError(f"Missing required headers: {', '.join(missing_headers)}")

    rows = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not any(row):  # Skip completely empty rows
            continue

        row_dict = {header_map.get(headers[idx], f"unknown_column_{idx}"): (cell if cell is not None else "") for idx, cell in enumerate(row)}
        row_dict["sheet_name"] = sheet_name  # Add sheet name to each row
        rows.append(row_dict)

    if validate_only:
        return []  # Validation successful; no need to return data

    return rows


def validate_input_file(file_path: str):
    """
    Validates the input Excel file without generating Markdown files.
    """
    try:
        data = read_excel(file_path, validate_only=True)
        msg_blurb_square(f"Validation successful for file: {file_path}", color_green)
    except FileNotFoundError:
        msg_blurb_square(f"Error: File not found: {file_path}", color_red)
    except ValueError as e:
        msg_blurb_square(f"Validation failed: {e}", color_red)
    except Exception as e:
        msg_blurb_square(f"Unexpected error during validation: {e}", color_red)


def Usage():
    file = sys.argv[0].split(os.sep)[-1]

    print(f'\nDescription: {color_green}{description2}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f"    {file} -c -I contacts_sample.xlsx -O c:\\temp")
    print(f"    {file} -x -I <input_folder> -O <output_folder>")
    print(f"    {file} --dry-run")


if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>


"""

0.1.0 - working copy
0.0.1 - created by ChatGPT
"""


# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
if filename doesn't exist, create it else create fullname_{row}.md
copy the original file to original_file when doing a -c

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
