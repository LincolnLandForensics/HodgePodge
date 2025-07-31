#!/usr/bin/env python3
# coding: utf-8
"""
Author: LincolnLandForensics
Version: 1.0.7
Parses .eml and .mbox files in a folder, extracts messages / non-spam contacts, and exports to Excel.
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>> #
import os
import re
import sys
import json
import email
import hashlib
import argparse
import mailbox
from email import policy
from dateutil import parser  # Make sure to install python-dateutil if needed
from datetime import datetime
from openpyxl import Workbook
from bs4 import BeautifulSoup
from email.parser import BytesParser
from email.header import decode_header
from openpyxl.styles import PatternFill
from email.utils import parsedate_to_datetime


# <<<<<<<<<<<<<<<<<<<<<<<<      Color Support        >>>>>>>>>>>>>>>>>>>>>>>> #
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
if sys.version_info > (3, 7, 9) and os.name == "nt":
    from colorama import Fore, Back, Style
    print(Back.BLACK)
    color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
    color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL

# <<<<<<<<<<<<<<<<<<<<<<<<      Main Function        >>>>>>>>>>>>>>>>>>>>>>>> #
def main():
    parser = argparse.ArgumentParser(description="Parse .eml and .mbox files into Excel")
    parser.add_argument('-I', '--input', help='Input folder of .eml files', required=False)
    parser.add_argument('-O', '--output', help='Output Excel file', required=False)
    parser.add_argument('-E', '--eml', help='Enable email parsing', action='store_true')

    args = parser.parse_args()
    input_folder = args.input if args.input else r"C:\Forensics\scripts\python\eml"
    output_xlsx = args.output if args.output else "email.xlsx"

    global data 
    data = []
    
    global contacts_data
    contacts_data = []

    if args.eml:
        if not os.path.exists(input_folder):
            msg_blurb_square(f'{input_folder} does not exist', color_red)
            exit()
        msg_blurb_square(f'Reading .eml and .mbox files from {input_folder} folder', color_green)
        data, contacts_data = read_eml(input_folder)
        # data, contacts_data = read_json(input_folder)

        contacts_data = contacts_deduplicate(contacts_data)
        write_xlsx(data, contacts_data, output_xlsx)
        msg_blurb_square(f'Writing to {output_xlsx}', color_green)

    else:
        usage()
    return 0

# <<<<<<<<<<<<<<<<<<<<<<<<      subroutines        >>>>>>>>>>>>>>>>>>>>>>>> #

def case_number_prompt():
    # Prompt the user to enter the case number
    case_number = input("Please enter the Case Number: ")
    # Assign the entered value to Case
    case_prompt = case_number
    return case_prompt

def contacts_deduplicate(contacts_data):
    # Step 1: Filter out spam entries and invalid emails
    filtered = []
    for contact in contacts_data:
        email = contact.get("email", "").lower()
        tag = contact.get("Tag", "")    # .lower()

        if tag == "Spam":
            continue
        if any(substr in email for substr in ["<", "noreply", "no-reply"]):
            continue

        filtered.append(contact)

    # Step 2: Deduplicate by email
    deduped = {}
    for contact in filtered:
        email = contact.get("email")
        if email and email not in deduped:
            deduped[email] = contact

    # Step 3: Sort by email
    sorted_contacts = sorted(deduped.values(), key=lambda c: c.get("email", "").lower())

    # Optional: Count of unique non-spam, valid emails
    print(f"✅ Count of unique emails in the contacts sheet: {len(sorted_contacts)}")

    return sorted_contacts

def clean_date(date_str):
    try:
        # Try using email parser first (if applicable)
        dt = parsedate_to_datetime(date_str)
        if dt:
            return dt.isoformat().replace('T', ' ')
    except Exception:
        pass  # Fall through to more flexible parsing

    try:
        # Use dateutil parser for natural language formats
        dt = parser.parse(date_str)
        return dt.isoformat().replace('T', ' ')
    except Exception:
        # Fallback for custom timestamp format: YYYY.MM.DD-HH.MM.SS
        match = re.match(r"(\d{4})\.(\d{2})\.(\d{2})-(\d{2})\.(\d{2})\.(\d{2})", date_str)
        if match:
            year, month, day, hour, minute, second = map(int, match.groups())
            dt = datetime(year, month, day, hour, minute, second)
            return dt.isoformat().replace('T', ' ') + 'Z'

    return ''


def count_email_files(folder_path):
    return sum(
        1 for file_name in os.listdir(folder_path)
        if file_name.lower().endswith((".eml", ".mbox"))
    )


def decode_header_str(header_obj):  # MBOX 
    if not header_obj:
        return ""
    decoded_parts = decode_header(header_obj)
    header_str = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            try:
                header_str += part.decode(encoding or 'utf-8', errors='replace')
            except Exception:
                header_str += part.decode('utf-8', errors='replace')
        else:
            header_str += str(part)
    return header_str.strip()


def decode_payload(payload):
    try:
        return payload.decode('utf-8')
    except UnicodeDecodeError:
        try:
            return payload.decode('latin1')
        except Exception:
            return payload.decode(errors='ignore')


def detect_spam(body, email):
    spam_keywords = ["Unsubscribe", "unsubscribe"]
    spam_email_keywords = ["donotreply", "Donotreply", "info@", "no-response", "notifications@", "postmaster@", "support@", "verify@"]

    if any(keyword in body for keyword in spam_keywords) or \
       any(keyword in email for keyword in spam_email_keywords):
        return "Spam"
    return ""


def extract_body(msg):  # mbox
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_dispo = str(part.get('Content-Disposition'))
            if content_type == 'text/plain' and 'attachment' not in content_dispo:
                body += part.get_payload(decode=True).decode(errors='replace')
            elif content_type == 'text/html' and not body:
                html = part.get_payload(decode=True).decode(errors='replace')
                soup = BeautifulSoup(html, 'lxml')
                body = soup.get_text()
    else:
        payload = msg.get_payload(decode=True)
        body = payload.decode(errors='replace') if payload else ''
    return body.strip()


def get_attachments(msg):
    attachments = []
    for part in msg.walk():
        if part.get_content_disposition() == 'attachment':
            filename = part.get_filename()
            if filename:
                attachments.append(filename)
    return "; ".join(attachments)


def get_attachment_filenames(msg):
    filenames = []
    for part in msg.walk():
        content_disposition = part.get("Content-Disposition", "")
        if 'attachment' in content_disposition.lower():
            filename = part.get_filename()
            if filename:
                filenames.append(filename)
    return filenames


def get_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                payload = part.get_payload(decode=True)
                return decode_payload(payload) if payload else ""
    else:
        payload = msg.get_payload(decode=True)
        return decode_payload(payload) if payload else ""
    return ""


def msg_blurb_square(msg, color):
    border = f"+{'-' * (len(msg) + 2)}+"
    print(f"{color}{border}\n| {msg} |\n{border}{color_reset}")


def read_eml(folder_path):
    case = case_number_prompt()
    source = source_prompt()
    
    data = []
    contacts_data = []
    # email_data2 = []

    def extract_contact_details(from_field):
        fullname = email = ''
        if ' <' in from_field:
            parts = from_field.split(' <')
            fullname = parts[0].strip().replace('"', '')
            email = parts[1].replace('>', '').strip()
        else:
            fullname = from_field.strip().replace('"', '')
            email = from_field.replace('>', '').strip()
        return fullname, email

    for file_name in os.listdir(folder_path):
        full_path = os.path.join(folder_path, file_name)
        sha256 = sha256_hash(full_path)

        if file_name.lower().endswith(".eml"):
            body, attachments, sender, recipient, subject, date = '', '', '', '', '', ''
            precedence, listUnsubscribe, xmailer, fullname, email, tag = '', '', '', '', '', ''
            labels = ''
            
            try:
                with open(full_path, 'rb') as f:
                    msg = BytesParser(policy=policy.default).parse(f)
                    sender = sanitize_field(msg.get("From"))
                    recipient = sanitize_field(msg.get("To"))
                    subject = sanitize_field(msg.get("Subject"))
                    labels = sanitize_field(msg.get("X-Gmail-Labels"))
               
                    date = msg.get("Date")
                    date_obj = clean_date(date) if date else ''
                    body = sanitize_field(get_body(msg))
                    attachments = sanitize_field(get_attachments(msg))
                    precedence = sanitize_field(msg.get("Precedence"))
                    listUnsubscribe = sanitize_field(msg.get("List-Unsubscribe"))
                    xmailer = sanitize_field(msg.get("X-Mailer"))

                    fullname, email = extract_contact_details(sender)
                    
                    spam_indicators = {
                        "List-Unsubscribe": listUnsubscribe,
                        "X-Mailer": xmailer,
                        "Precedence": precedence
                    }
                    if any(spam_indicators.values()):
                        tag = 'Spam'                    
                    
                    
                    tag = "CashApp" if "CashApp" in fullname else detect_spam(body, email)
                    tag = "Spam" if "Spam," in labels else tag
                    tag = "Important" if "Important" in labels else tag
                    
                    email_data1 = {
                        "Time": date_obj,
                        "From": sender,
                        "To": recipient,
                        "Subject": subject,
                        "Labels": labels,
                        "Body": body,
                        "Attachments": attachments,
                        "Tag": tag,
                        "original_file": file_name,
                        "Source": source,
                        "case": case,
                        "Precedence": precedence,
                        "List-Unsubscribe": listUnsubscribe,
                        "X-Mailer": xmailer,
                        "fullname": fullname,
                        "email": email,
                        'sha256': sha256,
                    }

                    contacts = {
                        "query": email,
                        "ranking": "3 - contacts",                    
                        "fullname": fullname,
                        "email": email,
                        "note": recipient,
                        "original_file": file_name,
                        "Source": source,
                        "Tag": tag
                    }

                    data.append(email_data1)
                    contacts_data.append(contacts)

            except Exception as e:
                print(f"{color_red}Error parsing {file_name}: {str(e)}{color_reset}")

        elif file_name.lower().endswith(".mbox"):
            body, attachments, sender, recipient, subject, date = '', '', '', '', '', ''
            precedence, listUnsubscribe, xmailer, fullname, email, tag = '', '', '', '', '', ''
            labels = ''
            try:
                mbox = mailbox.mbox(full_path)
                for message in mbox:
                    try:
                        sender = decode_header_str(message['from'] or '')
                        recipient = decode_header_str(message['to'] or '')
                        subject = decode_header_str(message['subject'] or '')
                        subject = sanitize_field(subject)
                        labels = decode_header_str(message['X-Gmail-Labels'] or '')
                        print(f'labels = {labels}') # temp
                        date = message['date']
                        date_obj = clean_date(date) if date else ''
                        body = extract_body(message)
                        body = sanitize_field(body)
                        precedence = message.get('precedence', '')
                        listUnsubscribe = message.get('List-Unsubscribe', '')
                        xmailer = message.get('X-Mailer', '')
                        attachments = get_attachment_filenames(message)
                        attachments = ', '.join(attachments) if attachments else ''
                        fullname, email = extract_contact_details(sender)
                        
                        spam_indicators = {
                            "List-Unsubscribe": listUnsubscribe,
                            "X-Mailer": xmailer,
                            "Precedence": precedence
                        }
                        if any(spam_indicators.values()):
                            tag = 'Spam'               
                        
                        
                        tag = "CashApp" if "CashApp" in fullname else detect_spam(body, email)
                        email_data2 = []
                        email_data2.append({
                            'Time': date_obj,
                            'From': sender,
                            'To': recipient,
                            'Subject': subject,
                            'Labels': labels,
                            'Body': body,
                            'Attachments': attachments,
                            'Tag': tag,
                            'original_file': file_name,
                            'Source': source,
                            'case': case,
                            'Precedence': precedence,
                            'List-Unsubscribe': listUnsubscribe,
                            'X-Mailer': xmailer,
                            'fullname': fullname,
                            'email': email,
                            'sha256': sha256,
                        })

                        contacts = {
                            "query": email,
                            "ranking": "3 - contacts",
                            "fullname": fullname,
                            "email": email,
                            "note": recipient,
                            "original_file": file_name,
                            "Tag": tag,
                            "Source": source
                        }

                        data.extend(email_data2)
                        contacts_data.append(contacts)

                    except Exception as e:
                        print(f"Error parsing message: {e}")

            except Exception as e:
                print(f"Error opening mbox file {file_name}: {e}")

        elif file_name.lower().endswith(".json"):
            print(f'parsing {file_name}')   # temp
            full_path = os.path.join(folder_path, file_name)
            sha256 = sha256_hash(full_path)

            try:
                with open(full_path, "r", encoding="utf-8") as f:
                    messages_json = json.load(f)

                for message in messages_json.get("messages", []):
                    sender_email = message.get("creator", {}).get("email", "")  # 
                    timestamp = message.get("created_date", "")
                    attachment_data = message.get("attached_files", "")
                    if isinstance(attachment_data, list):
                        attachment = ", ".join([a.get("export_name", "") for a in attachment_data if isinstance(a, dict)])
                    else:
                        # Fallback if it's a dict or an empty string
                        attachment = attachment_data.get("export_name", "") if isinstance(attachment_data, dict) else ""

                    
                    # attachment = message.get("attached_files", "")
                    # if attachment == '':
                        # attachment = message.get("attached_files", {}).get("export_name", "")

                    # attachments_list = message.get("attached_files", [])
                    # attachment_names = [item.get("export_name", "") for item in attachments_list if isinstance(item, dict)]
                    # attachment = ", ".join(attachment_names) if attachment_names else ""

                    
                    date_obj = clean_date(timestamp)
                    fullname = message.get("creator", {}).get("name", "")
                    body = message.get("text", "")
                    date_obj = clean_date(message.get("created_date", ""))
                    subject = ""  # Not present in doc
                    recipient = ""  # Not present either

                    tag = "CashApp" if "CashApp" in fullname else detect_spam(body, sender_email)

                    email_data1 = {
                        "Time": date_obj,
                        "From": sender_email,
                        "To": recipient,
                        "Subject": subject,
                        "Body": body,
                        "Attachments": attachment,
                        "Tag": tag,
                        "original_file": file_name,
                        "Source": source,
                        "case": case,
                        "Precedence": "",
                        "List-Unsubscribe": "",
                        "X-Mailer": "",
                        "fullname": fullname,
                        "email": sender_email,
                        "sha256": sha256,
                    }

                    contacts = {
                        "query": sender_email,
                        "ranking": "3 - contacts",
                        "fullname": fullname,
                        "email": sender_email,
                        "note": recipient,
                        "original_file": file_name,
                        "Source": source,
                        "Tag": tag
                    }

                    data.append(email_data1)
                    contacts_data.append(contacts)

            except Exception as e:
                print(f"{color_red}Error parsing {file_name}: {str(e)}{color_reset}")

    eml_count = count_email_files(folder_path)
    print(f"\nFound {eml_count} .eml and .mbox files in the {folder_path} folder.")

    return data, contacts_data


def read_json(folder_path):

    case = case_number_prompt()
    source = source_prompt()

    data = []
    contacts_data = []

    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith(".json"):
            print(f'parsing {file_name}')   # temp
            full_path = os.path.join(folder_path, file_name)
            sha256 = sha256_hash(full_path)

            try:
                with open(full_path, "r", encoding="utf-8") as f:
                    messages_json = json.load(f)

                for message in messages_json.get("messages", []):
                    sender_email = message.get("creator", {}).get("email", "")
                    fullname = message.get("creator", {}).get("name", "")
                    body = message.get("text", "")
                    date_obj = clean_date(message.get("created_date", ""))
                    subject = ""  # Not present in doc
                    recipient = ""  # Not present either

                    tag = "CashApp" if "CashApp" in fullname else detect_spam(body, sender_email)

                    email_data1 = {
                        "Time": date_obj,
                        "From": sender_email,
                        "To": recipient,
                        "Subject": subject,
                        "Body": body,
                        "Attachments": "",
                        "Tag": tag,
                        "original_file": file_name,
                        "Source": source,
                        "case": case,
                        "Precedence": "",
                        "List-Unsubscribe": "",
                        "X-Mailer": "",
                        "fullname": fullname,
                        "email": sender_email,
                        "sha256": sha256,
                    }

                    contacts = {
                        "query": sender_email,
                        "ranking": "3 - contacts",
                        "fullname": fullname,
                        "email": sender_email,
                        "note": recipient,
                        "original_file": file_name,
                        "Source": source,
                        "Tag": tag
                    }

                    data.append(email_data1)
                    contacts_data.append(contacts)

            except Exception as e:
                print(f"{color_red}Error parsing {file_name}: {str(e)}{color_reset}")

    return data, contacts_data
    
def remove_illegal_chars(text):
    if not isinstance(text, str):
        return text
    # Remove control characters and illegal XML characters
    return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', text)


def sanitize_field(text):
    text = (text or '').strip().strip('"')

    return remove_illegal_chars(text.strip()) if isinstance(text, str) else text


def sha256_hash(full_path):
    '''
    Calculate the SHA-256 hash of the file at full_path
    '''
    sha256 = hashlib.sha256()
    try:
        with open(full_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b''):
                sha256.update(chunk)
        sha256 = sha256.hexdigest()

    except Exception as e:
        print(f"⚠️ Error hashing file: {e}")

    return sha256


def source_prompt():
    # Prompt the user to enter the source file, like the Google Return zip file
    source = input("Please enter the original (zip) file name: ")
    return source


def write_xlsx(data, contacts_data, file_path):
    workbook = Workbook()

    # Worksheet 1
    worksheet = workbook.active
    worksheet.title = 'Eml'
    worksheet.freeze_panes = 'B2'

    # Worksheet 2
    worksheet2 = workbook.create_sheet(title='Contacts')
    worksheet2.freeze_panes = 'B2'

    headers = [
        "Time", "From", "To", "Subject", "Body", "Attachments", "Labels", "Tag",
        "original_file", "Source", "case", "Precedence", "List-Unsubscribe", "X-Mailer", "fullname", "email", "sha256"
    ]
    
    headers2 = [
        "query", "ranking", "fullname", "url", "email", "user", "phone", "business", "fulladdress", "city", "state", "country", "note", "AKA", "DOB", "SEX", "info", "misc", "firstname", "middlename", "lastname", "associates", "case", "sosfilenumber", "owner", "president", "sosagent", "managers", "Time", "Latitude", "Longitude", "Coordinate", "original_file", "Source", "Source file information", "Plate", "VIS", "VIN", "VYR", "VMA", "LIC", "LIY", "DLN", "DLS", "content", "referer", "osurl", "titleurl", "pagestatus", "ip", "dnsdomain", "Tag", "Icon", "Type"
    ]

    # Header formatting
    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [0, 1, 2, 3, 4, 6]:
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.fill = fill

    # Column widths
    worksheet.column_dimensions['A'].width = 18
    worksheet.column_dimensions['B'].width = 20
    worksheet.column_dimensions['C'].width = 25
    worksheet.column_dimensions['D'].width = 20
    worksheet.column_dimensions['E'].width = 20
    worksheet.column_dimensions['F'].width = 13    
    worksheet.column_dimensions['G'].width = 10    
    worksheet.column_dimensions['H'].width = 10    
    worksheet.column_dimensions['I'].width = 20    
    worksheet.column_dimensions['J'].width = 7    
        
    worksheet.column_dimensions['N'].width = 18    
    worksheet.column_dimensions['O'].width = 27    
    worksheet.column_dimensions['P'].width = 30    

    worksheet2.column_dimensions['A'].width = 18
    worksheet2.column_dimensions['B'].width = 14
    worksheet2.column_dimensions['C'].width = 25
    worksheet2.column_dimensions['D'].width = 4
    worksheet2.column_dimensions['E'].width = 20
    worksheet2.column_dimensions['F'].width = 5    
    worksheet2.column_dimensions['G'].width = 10    
    worksheet2.column_dimensions['H'].width = 9    
    worksheet2.column_dimensions['I'].width = 12    
    worksheet2.column_dimensions['J'].width = 5    
    worksheet2.column_dimensions['M'].width = 25         
    worksheet2.column_dimensions['N'].width = 4    
    worksheet2.column_dimensions['O'].width = 4    
    worksheet2.column_dimensions['P'].width = 4
    worksheet2.column_dimensions['AZ'].width = 6


    # Header formatting for worksheet 1
    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [0, 1, 2, 3, 4, 6]:
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            cell.fill = fill

    # Header formatting for worksheet 2
    for col_index, header in enumerate(headers2):
        cell = worksheet2.cell(row=1, column=col_index + 1)
        cell.value = header

    # Data rows for worksheet 1
    for row_index, row_data in enumerate(data):
        for col_index, col_name in enumerate(headers):
            cell_data = row_data.get(col_name)
            try:
                worksheet.cell(row=row_index + 2, column=col_index + 1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error writing line in email: {str(e)}{color_reset}")

    for row_index, row_data in enumerate(contacts_data):
        for col_index, col_name in enumerate(headers2):
            try:
                cell_data = row_data.get(col_name)
                worksheet2.cell(row=row_index + 2, column=col_index + 1).value = cell_data
            except Exception as e:
                print(f"Error writing line in Contacts: {str(e)}")

    workbook.save(file_path)


def usage():
    print(f"Usage: {sys.argv[0]} -E [-I input_folder] [-O output.xlsx]")
    print("Examples:")
    print(f"    {sys.argv[0]} -E")
    print(f"    {sys.argv[0]} -E -I C:\\emails -O parsed_emails.xlsx")


# <<<<<<<<<<<<<<<<<<<<<<<<      Run Program        >>>>>>>>>>>>>>>>>>>>>>>> #
if __name__ == '__main__':
    main()
    

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
1.0.6 - Contacts sheet with unique emails that aren't from spam
1.0.2 - -E now does .eml and .mbox files
1.0.0 - .eml and .mbox parser

"""


# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
parse .msg, .emlx, .pst, .ost .vcf, .chat
see if RLEAPP is doing anything different

"""


# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
NDCAC Google Warrant return viewer does convert .json to .eml and this can then parse the .eml

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
    