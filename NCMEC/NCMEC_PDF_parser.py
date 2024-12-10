 #!/usr/bin/env python3
# coding: utf-8
'''
Dump all of your NCMEC .zip or .pdf files into the NCMEC folder.
it will export the emails, ip's, md5's, phone numbers and users into _output folder

Example:
    python NCMEC_PDFs_parser.py
'''

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import re
import zipfile
# import py7zr  # For handling .7z files  # pip install py7zr
from datetime import datetime
from PyPDF2 import PdfReader

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

pdf_folder = "NCMEC"    # Specify the folder containing .PDF and/or .zip files
target_folder = os.path.join(pdf_folder, '_output')

author = 'LincolnLandForensics'
description = "read .zip and .pdfs and extract out NCMEC intel"
version = '0.1.2'


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def cleanup_email(email_list):
    """
    Removes emails that end with specific terms in the whitelist.
    
    Args:
        email_list (list): A list of email addresses to clean up.
    
    Returns:
        list: A cleaned list of email addresses.
    """
    # Define the whitelist of unwanted endings
    whitelist = ['.gov', 'ncmec.org', '.Upon']
    # whitelist = ['.gov', 'USlawenforcement@google.com', 'ncmec.org', '.Upon', 'lawenforcement@discordapp.com']
  
    
    # Filter emails that do not match any of the whitelist terms and do not contain 'lawenforcement'
    cleaned_emails = [
        email for email in email_list
        if not any(email.lower().endswith(term.lower()) for term in whitelist)  # Exclude unwanted endings
        and 'lawenforcement' not in email.lower()  # Exclude emails containing 'lawenforcement'
    ]
    
    return cleaned_emails
    
    
def process_pdfs_and_unzip(pdf_folder, target_folder):
    # Ensure the folder exists
    if not os.path.exists(pdf_folder):
        print(f"The folder '{pdf_folder}' does not exist. Exiting.")
        exit(1)
    # print(f"The folder '{pdf_folder}' exists. Proceeding with PDF processing.")

    # First unzip all .zip files
    unzip_all_in_folder(pdf_folder, target_folder)

    print(f"Processing PDFs in {pdf_folder} folder....")


    
    today = datetime.now().strftime("%Y-%m-%d") # Get today's date in YY-MM-DD format

    output_folder = "Output"    # Output folder
    os.makedirs(output_folder, exist_ok=True)
    # Initialize lists to store extracted data
    word_list = []
    sentence_list = []
    email_list = []
    md5_list = []
    ip_list = []
    phone_list = []
    user_list = []
    
    # Regular expression patterns
    email_pattern = re.compile(r'[\w.-]+@[\w.-]+\.[a-zA-Z]{2,6}')
    md5_pattern = re.compile(r'\b[a-fA-F0-9]{32}\b')
    ip_pattern = re.compile(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b')
    ipv6_pattern = re.compile(r'([0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}')  # Example pattern for full IPv6 addresses
    phone_pattern = re.compile(r'\b(?:\+?\d{1,3})?[ .-]?(?:\(\d{1,4}\)|\d{1,4})[ .-]?\d{1,4}[ .-]?\d{1,4}[ .-]?\d{1,9}\b')

    # Helper function to process a single PDF
    def process_pdf(file_path):
        nonlocal word_list, sentence_list, email_list, md5_list, ip_list, phone_list
        reader = PdfReader(file_path)
        for page in reader.pages:
            text = page.extract_text()

            # Split text into words and sentences
            words = re.split(r'["<>\\s]', text)
            word_list.extend(filter(None, words))  # Filter out empty strings
            
            sentences = text.split('\n')
            
            for sentence in sentences:
                sentence = sentence.strip()
                if sentence.startswith('Screen/User Name: '):
                    user = sentence.split(':')[1].strip()  # Extract user name after ':'
                    user_list.append(user)  # Append user to user_list
                elif sentence.startswith('CyberTipline Report ') and 'was submitted by a member' not in sentence:
                    user = sentence
                    user = user.replace('CyberTipline Report', '').strip()  # Removing the prefix and any extra spaces
                    user = user.split(' ')[0].strip()  # Extract user name after ':'
                    user_list.append(user)  # Append user to user_list
                elif sentence.startswith('IP Address: ') and ' (Login)' in sentence:
                    ip = sentence
                    ip = ip.replace('IP Address: ', '').strip()
                    ip = ip.split(' ')[0].strip()
                    ip_list.append(ip)  # Append ip
                elif sentence.startswith('Phone: +') and ' (Verified' in sentence:
                    phone = sentence
                    phone = phone.replace('Phone: +', '')
                    phone = phone.split(' ')[0]
                    phone_list.append(phone)  # Append phone

                elif sentence.startswith('Mobile Phone: +') and 'Verified' in sentence:
                    phone = sentence
                    phone = phone.replace('Mobile Phone: +', '')
                    phone = phone.split(' ')[0]
                    phone_list.append(phone)  # Append phone
                if sentence:  # Only append non-empty sentences
                    sentence_list.append(sentence)            
            
            
            
            
            sentence_list.extend(filter(None, sentences))

            # Extract data using regex
            email_list.extend(email_pattern.findall(text))
            md5_list.extend(md5_pattern.findall(text))

            ip_list.extend(ip_pattern.findall(text))    # ipv4

            # ip_list.extend(ipv6_pattern.findall(text))  # ipv6



# ipv6
            # Define the regex pattern for matching IPv6 addresses (or any other regex pattern)

            # Example text containing IPv6 addresses
            # text = "The IPv6 addresses are 2001:db8:85a3::8a2e:370:7334 and fe80::1ff:fe23:4567:890a."

            # Initialize ip_list
            # ip_list = []

            # Find all matches (this could return tuples, depending on your pattern)
            # matches = ipv6_pattern.findall(text)

            # Flatten the list if needed (extract strings from tuples)
            # flat_list = [item[0] for item in matches] if matches and isinstance(matches[0], tuple) else matches

            # Now apply unique_sorted to get a sorted list with unique values
            # unique_sorted = lambda x: sorted(set(x))
            # sorted_unique_ips = unique_sorted(flat_list)

            # ip_list.extend(sorted_unique_ips)  # ipv6


            # phone_list.extend(phone_pattern.findall(text))

    # Unzip all .zip files and process PDFs
    for root, _, files in os.walk(pdf_folder):
        for file in files:
            file_path = os.path.join(root, file)

            ## If the file is a .zip, extract it
            # if file.endswith('.zip'):
                # with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    # zip_ref.extractall(root)

            ## elif the file is a .pdf, process it
            if file.endswith('.pdf'):
                process_pdf(file_path)
                print(f'\t{file_path}')   # temp
    # Function to sort and remove duplicates
    unique_sorted = lambda x: sorted(set(x))

    email_list = cleanup_email(email_list)
    word_list = split_word_list(word_list)

    # Prepare data for export
    export_data = {
        # f"words_{today}.txt": unique_sorted(word_list),
        # f"sentences_{today}.txt": unique_sorted(sentence_list),
        f"emails_{today}.txt": unique_sorted(email_list),
        f"md5_hashes_{today}.txt": unique_sorted(md5_list),
        f"ips_{today}.txt": unique_sorted(ip_list),
        f"phones_{today}.txt": unique_sorted(phone_list),
        f"users_{today}.txt": unique_sorted(user_list),        
    }

    # Write data to text files
    for filename, data in export_data.items():
        # print(f'filename = {filename} to output_folder= {output_folder} target_folder= {target_folder}') # test
        
        # with open(os.path.join(output_folder, filename), "w") as file:
        with open(os.path.join(target_folder, filename), "w") as file:
               
        
            file.write("\n".join(data))

    print(f"See text files in {target_folder} folder.")


def split_word_list(word_list):
    """
    Splits each word in the word_list by spaces, parentheses, and new lines.
    Cleans each word by stripping leading/trailing whitespace and removing '.' or ',' from the end.
    
    Args:
        word_list (list): A list of strings (words) to split.
    
    Returns:
        list: A new list with all the split and cleaned components.
    """
    # Define the regex pattern to split by spaces, parentheses, and new lines
    split_pattern = r'[ \n()]+'  # Matches spaces, new lines, and parentheses
    
    cleaned_words = []  # To store the cleaned and split words
    for word in word_list:
        # Split each word using the regex pattern
        split_components = filter(None, re.split(split_pattern, word))
        # Clean each split component
        for component in split_components:
            # Strip leading/trailing whitespace
            component = component.strip()
            # Remove '.' or ',' from the end of the word
            component = component.rstrip('.,:')
            # Add cleaned component to the list
            cleaned_words.append(component)
    
    return cleaned_words


def unzip_all_in_folder(pdf_folder, target_folder):
    """
    Unzips all .zip and .7z files in the pdf_folder (including subfolders) into the specified target_folder.
    If a folder with the same name already exists in the target folder, appends the current date string.
    """
    # Check if the folder exists
    if not os.path.exists(pdf_folder):
        print(f"The folder '{pdf_folder}' does not exist. Exiting.")
        exit(1)

    # Ensure the target folder exists
    if not os.path.exists(target_folder):
        print(f"The folder '{target_folder}' does not exist. Exiting.")
        os.makedirs(target_folder, exist_ok=True)
        exit(1)
    # print(f"The folder '{target_folder}' exists. Proceeding.")


    zip_count = 0
    date_string = datetime.now().strftime("%Y-%m-%d")  # Format date string

    # Walk through the directory and process .zip and .7z files
    for root, _, files in os.walk(pdf_folder):
        for file in files:
            file_path = os.path.join(root, file)

            # Determine the output folder name based on the archive file name
            folder_name = os.path.splitext(file)[0]
            # output_folder = os.path.join(target_folder, folder_name)
            output_folder = os.path.join(pdf_folder, folder_name)   # test
             
            # If the folder already exists, append the date string
            if os.path.exists(output_folder):
                output_folder = f"{output_folder}_{date_string}"

            # Handle .zip files
            if file.endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    os.makedirs(output_folder, exist_ok=True)
                    zip_ref.extractall(output_folder)
                print(f"Unzipped: {file_path} to {output_folder}")
                zip_count += 1

            # Handle .7z files
            # elif file.endswith('.7z'):
                # with py7zr.SevenZipFile(file_path, mode='r') as seven_zip_ref:
                    # os.makedirs(output_folder, exist_ok=True)
                    # seven_zip_ref.extractall(path=output_folder)
                # print(f"Unzipped: {file_path} to {output_folder}")
                # zip_count += 1

    print(f"{zip_count} files unzipped in {pdf_folder} folder")
    return zip_count


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

0.1.0 - working prototype
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
split ipv6 in half and add them to the ip list

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
This assumes NCMEC .zip files have no password
IP and MD5 don't have to be NCMEC specific
"""


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Main   >>>>>>>>>>>>>>>>>>>>>>>>>>

process_pdfs_and_unzip(pdf_folder, target_folder)

