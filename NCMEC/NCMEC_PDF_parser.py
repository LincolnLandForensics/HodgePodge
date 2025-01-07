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
import sys
import zipfile
import argparse
from datetime import datetime
try:
    import pdfplumber  # pip install pdfplumber
except Exception as e:
    print(f"{str(e)}")
    print(f'pip install pdfplumber')
    sys.exit()

sys.stdout.reconfigure(encoding='utf-8')  # Python 3.7+

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "read .zip and .pdfs and extract out NCMEC intel"
version = '0.1.8'

# <<<<<<<<<<<<<<<<<<<<<<<<<<    Main    >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    """
    Main function to parse arguments and initiate file conversion.
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='Input folder path', required=False)
    parser.add_argument('-O', '--output', help='Output folder path', required=False)
    parser.add_argument('-p', '--parse', help='parse PDFs', required=False, action='store_true')

    args = parser.parse_args()

    global input_folder
    global output_folder

    input_folder = "NCMEC"    # Specify the folder containing .PDF and/or .zip files
    output_folder = os.path.join(input_folder, '_output')

    # Set input and output folders based on arguments, if provided
    if args.input:
        input_folder = args.input
    if args.output:
        output_folder = args.output
        output_folder = output_folder.lstrip("\\")  # doesn't work with something like \temp
    # else:
        # print(f'output folder is: {output_folder}')  # temp        

    # Ensure the input folder exists
    if not os.path.exists(input_folder):
        print(f"Input folder doesn't exist: {input_folder}")
        return 1
        
    # Ensure the output folder exists, or create it if it doesn’t
    if not os.path.exists(output_folder):
        os.makedirs(output_folder, exist_ok=True)
        print(f"Output folder doesn't exist: {output_folder}")
        print(f'Output folder created: {output_folder}')
        
    if args.parse:
        process_pdfs_and_unzip(input_folder, output_folder)

    else:
        parser.print_help()
        Usage()

    return 0


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>


def check_pdf_readability(file_path):
    try:
        with pdfplumber.open(file_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text and text.strip():  # If any page has text, return
                    print(f"{file_path} is readable.")
                    return
        print(f"{file_path} needs to be OCR'd to read it.")
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        
def cleanup_email(email_list):
    whitelist = ['.gov', 'ncmec.org', '.Upon']
    return [
        email for email in email_list
        if not any(email.lower().endswith(term.lower()) for term in whitelist)
        and 'lawenforcement' not in email.lower()
    ]


def process_pdfs_and_unzip(input_folder, output_folder):
    print(f'\nParsing files in {input_folder} folder\n')
    unzip_all_in_folder(input_folder, output_folder)

    today = datetime.now().strftime("%Y-%m-%d")
    # word_list, sentence_list, email_list, md5_list, ip_list, phone_list, user_list = ([] for _ in range(7))
    word_list, sentence_list, email_list, file_list, md5_list, ip_list, phone_list, user_list = ([] for _ in range(8))

    email_pattern = re.compile(r'[\w.-]+@[\w.-]+\.[a-zA-Z]{2,6}')
    md5_pattern = re.compile(r'\b[a-fA-F0-9]{32}\b')
    ip_pattern = re.compile(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b')
    phone_pattern = re.compile(r'\b(?:\+?\d{1,3})?[ .-]?(?:\(\d{1,4}\)|\d{1,4})[ .-]?\d{1,4}[ .-]?\d{1,4}[ .-]?\d{1,9}\b')

    global pdfs_parsed
    pdfs_parsed = 0
        
    def process_pdf(file_path):

        nonlocal word_list, sentence_list, email_list, file_list, md5_list, ip_list, phone_list, user_list

        try:
            text = ''
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    words = re.split(r'["<>\\s]', text)
                    word_list.extend(filter(None, words))

                    sentences = text.split('\n')
                    for sentence in sentences:
                        sentence = sentence.strip()
                        if sentence.startswith('Screen/User Name: '):
                            user_list.append(sentence.split(':', 1)[1].strip())
                        elif sentence.startswith('ESP User ID: '):
                            user_list.append(sentence.replace('ESP User ID: ', '').strip())
                        elif sentence.startswith('Filename:'):
                            file_list.append(sentence.replace('Filename:', '').strip())
                        elif sentence.startswith('IP Address: ') and ' (Login)' in sentence:
                            ip_list.append(sentence.replace('IP Address: ', '').split(' ')[0].strip())
                        elif sentence.startswith('Phone: +') and ' (Verified' in sentence:
                            phone_list.append(sentence.replace('Phone: +', '').split(' ')[0])
                        elif sentence.startswith('Mobile Phone: +') and 'Verified' in sentence:
                            phone_list.append(sentence.replace('Mobile Phone: +', '').split(' ')[0])
                        if sentence:
                            sentence_list.append(sentence)

                    email_list.extend(email_pattern.findall(text))
                    md5_list.extend(md5_pattern.findall(text))
                    ip_list.extend(ip_pattern.findall(text))
            # pdfs_parsed += 1
        except Exception as e:
            print(f"Error processing {file_path}: {str(e)}")

        # test for pdf's that need to be OCR'd first
        if text and text.strip():  # If any page has text, return
            blah = 'blah'
        else:
            print(f"{file_path} needs to be OCR'd to read it.")  


    for root, _, files in os.walk(input_folder):
        for file in files:
            file_path = os.path.join(root, file)
            if file.endswith('.pdf'):
                process_pdf(file_path)

    unique_sorted = lambda x: sorted(set(x))
    email_list = cleanup_email(email_list)
    word_list = split_word_list(word_list)

    export_data = {
        # f"sentences_{today}.txt": unique_sorted(sentence_list),
        f"emails_{today}.txt": unique_sorted(email_list),
        f"files_{today}.txt": unique_sorted(file_list),        
        f"md5_hashes_{today}.txt": unique_sorted(md5_list),
        f"ips_{today}.txt": unique_sorted(ip_list),
        f"phones_{today}.txt": unique_sorted(phone_list),
        f"users_{today}.txt": unique_sorted(user_list),
    }

    for filename, data in export_data.items():
        with open(os.path.join(output_folder, filename), "w", encoding="utf-8", errors="replace") as file:
            if "md5_hashes_" in filename:  # Check if this is the MD5 file
                file.write("MD5\n")       # Prepend "MD5" as the first line

            file.write("\n".join(data))

    if pdfs_parsed != 0:
        # print(f"\n{pdfs_parsed} PDfs parsed")
        print(f"\nPDfs parsed")

    print(f"\n	{len(unique_sorted(email_list))} emails")
    print(f"\t{len(unique_sorted(file_list))} files")    
    print(f"\t{len(unique_sorted(ip_list))} IPs")
    print(f"\t{len(unique_sorted(md5_list))} MD5 hashes")
    print(f"\t{len(unique_sorted(phone_list))} phone numbers")
    print(f"\t{len(unique_sorted(user_list))} users")
    print(f"\nSee output in {output_folder} folder.")


def split_word_list(word_list):
    split_pattern = r'[ \n()]+'
    cleaned_words = []
    for word in word_list:
        split_components = filter(None, re.split(split_pattern, word))
        for component in split_components:
            component = component.strip().rstrip('.,:')
            cleaned_words.append(component)
    return cleaned_words


def unzip_all_in_folder(input_folder, output_folder):
    zip_count = 0
    date_string = datetime.now().strftime("%Y-%m-%d")

    for root, _, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith('.zip'):
                file_path = os.path.join(root, file)
                folder_name = os.path.splitext(file)[0]
                target_folder = os.path.join(output_folder, folder_name)
                if os.path.exists(target_folder):
                    target_folder = f"{target_folder}_{date_string}"

            if file.lower().endswith('.zip'):
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    os.makedirs(target_folder, exist_ok=True)
                    try:
                        zip_ref.extractall(target_folder)
                    except Exception as e:
                        print(f"{str(e)}")
                    
                zip_count += 1

    if zip_count != 0:
        print(f"\n{zip_count} files unzipped")
    
    return zip_count


def Usage():
    """
    Prints usage information for the script.
    """
    print("\nDescription: " + description)
    print(sys.argv[0] + " Version: %s by %s" % (version, author))
    print(f'\nExample:')
    print("\t" + sys.argv[0] + " -p")
    print("\t" + sys.argv[0] + " -p -I C:\\Forensics\\scripts\\python\\NCMEC -O C:\\Forensics\\scripts\\python\\NCMEC\\_output")

if __name__ == '__main__':
    main()
    

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
0.2.0 - added an additional output file with filenames?
0.1.9 - automatically create outputful folder inside the input folder, if it doesn't exist and -O isn't specified. 
0.1.8 - add "MD5" as the first line of the hash file (X-Ways requires it.) 
0.1.7 - fixed -O option to work in specified folder
0.1.3 - handle udf8 characters (like emoji), removed cybertip number from users.txt
0.1.2 - stopped exporting sentences and words, added menu, added 'ESP User ID: ' to users.txt
0.1.0 - working prototype
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
detect if pdf needs to be OCR'd. 
split ipv6 in half and add them to the ip list
delete the unzipped folders at the end of the script 

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
This assumes NCMEC .zip files have not password
IP and MD5 don't have to be NCMEC specific
.7z extraction can be easily added if needed

"""
print(f"\nBob's your uncle\n")


