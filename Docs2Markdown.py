#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>
import os
import sys
import fitz  # pip install PyMuPDF
import json
import docx
import email
import time
from email.policy import default
import argparse # for menu system


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>
author = 'LincolnLandForensics'
description = "Convert the content of .txt, .pdf, .docx, and .eml to Markdown, for use in Obsidian"
version = '0.1.3'

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>


def main():
    """
    Main function to parse arguments and initiate file conversion.
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='Input folder path', required=False)
    parser.add_argument('-O', '--output', help='Output folder path', required=False)
    parser.add_argument('-c', '--convert', help='Convert files to markdown', required=False, action='store_true')

    args = parser.parse_args()

    global input_folder
    global output_folder
    input_folder = os.getcwd()  # Default to current working directory
    # output_folder = r'C:\Forensics\scripts\python\ObsidianNotebook'  # Default output path
    output_folder = r'ObsidianNotebook'  # Default output path


    # Set input and output folders based on arguments, if provided
    if args.input:
        input_folder = args.input
    if args.output:
        output_folder = args.output

    # Ensure the input folder exists
    if not os.path.exists(input_folder):
        print(f"Input folder {input_folder} doesn't exist.")        
        # logging.error(f"Input folder '{input_folder}' doesn't exist.")
        return 1

    # Ensure the output folder exists, or create it if it doesn’t
    if not os.path.exists(output_folder):
        print(f"output_folder doesn't exist: {output_folder}")
        # os.makedirs(output_folder, exist_ok=True)
        # logging.info(f"Created output folder '{output_folder}'.")
        sys.exit(1)  # Exit the script if output folder cannot be created
        
    if args.convert:
        # logging.info(f'Starting conversion of files in {input_folder} to markdown format in {output_folder}.')
        process_files(input_folder, output_folder)
    else:
        parser.print_help()
        Usage()

    return 0


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

 
def extract_text_from_pdf(pdf_path):
    """
    Function to extract text from a PDF file.
    """
    try:
        doc = fitz.open(pdf_path)
        text = f"\n## file_path: {pdf_path}\n"

        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            text += page.get_text("text")  # Extract text
        return text
    except fitz.EmptyFileError:
        print(f"Cannot open empty or corrupted PDF file: {pdf_path}")
        # logging.error(f"Cannot open empty or corrupted PDF file: {pdf_path}")
        return ""
    except Exception as e:
        print(f"Error reading PDF file '{pdf_path}': {e}")
        # logging.error(f"Error reading PDF file '{pdf_path}': {e}")
        return ""


def extract_text_from_docx(docx_path):
    '''
    Function to extract text from DOCX file
    '''    
    doc = docx.Document(docx_path)
    text = f"\n## file_path: {docx_path}\n"
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text


def extract_text_from_eml(eml_path):
    '''
    Function to extract text from EML file
    ''' 
    with open(eml_path, 'r', encoding='utf-8') as f:
        msg = email.message_from_file(f, policy=default)

    text = f"\n## file_path: {eml_path}\nSubject: {msg['subject']}\nFrom: {msg['from']}\nTo: {msg['to']}\n\n"
    
    # Extract email body
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() == "text/plain":
                text += part.get_content()
    else:
        text += msg.get_content()
    
    return text


def convert_to_markdown(file_path, output_folder):
    '''
    Function to convert the content of .txt, .pdf, .docx, and .eml to Markdown
    ''' 
    file_name = os.path.basename(file_path)
    base_name = os.path.splitext(file_name)[0]  # Remove the file extension
    target_file_path = os.path.join(output_folder, base_name + ".md")
    
    # Extract content based on file type
    if file_path.lower().endswith(".txt") or file_path.lower().endswith(".py"):
    # if file_path.endswith(".txt"):
        # Attempt to open the file with utf-8 encoding first, with fallback
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                # content = f.read()
                text2 = f.read()
                content = (f"\n## file_path: {file_path}\n{text2}\n")

        except UnicodeDecodeError:
            print(f"UTF-8 decoding failed for {file_path}, trying ISO-8859-1 encoding.")
                     
            # logging.warning(f"UTF-8 decoding failed for {file_path}, trying ISO-8859-1 encoding.")
            
            try:
                with open(file_path, 'r', encoding='ISO-8859-1') as f:
                    # content = f.read()
                    text2 = f.read()
                    content = (f"\n## file_path: {file_path}\n{text2}\n")                    
                    
                    # content = (f"\n## file_path: {file_path}\nf.read()")
                    # content = f.read()
            except UnicodeDecodeError:
                print(f"Cannot decode {file_path} with UTF-8 or ISO-8859-1.")                
                # logging.error(f"Cannot decode {file_path} with UTF-8 or ISO-8859-1.")
                return  # Skip this file if decoding fails
    elif file_path.lower().endswith(".pdf"):    
    # elif file_path.endswith(".pdf"):
        content = extract_text_from_pdf(file_path)
    elif file_path.lower().endswith(".docx"):
        content = extract_text_from_docx(file_path)
    elif file_path.lower().endswith(".eml"):
        content = extract_text_from_eml(file_path)
    else:
        print(f"Unsupported file type: {file_path}")
        return
    
    # Write the content to a Markdown file
    with open(target_file_path, 'w', encoding='utf-8') as md_file:
        md_file.write(content)
    
    print(f"Converted '{file_name}' to '{base_name}.md'.")


def get_file_timestamps(file_path):
    """
    Returns the creation, last accessed, and last modified times of a file.

    :param file_path: The file path for which timestamps are needed.
    :return: A tuple containing (creation_time, access_time, modified_time)
    """
    if not os.path.isfile(file_path):
        raise ValueError(f"The path {file_path} is not a valid file.")

    # Get the file stats
    file_stats = os.stat(file_path)

    # Get the creation time (Windows), access time, and modification time
    creation_time = file_stats.st_ctime  # Windows: creation time, Linux: change time
    access_time = file_stats.st_atime
    modified_time = file_stats.st_mtime

    # Convert from timestamps to human-readable format
    creation_time = time.ctime(creation_time)
    access_time = time.ctime(access_time)
    modified_time = time.ctime(modified_time)

    return creation_time, access_time, modified_time
def process_files(root_folder, output_folder):
    '''
    Function to walk through all files and convert
    ''' 

    setup_obsidian(output_folder)   # test

    # Walk through all files and subdirectories in root_folder
    for subdir, _, files in os.walk(root_folder):
        for file in files:
            file_path = os.path.join(subdir, file)
            # Check for file extensions: .txt, .pdf, .docx, .eml
            if file.lower().endswith(('.txt', '.pdf', '.docx', '.eml')):
                convert_to_markdown(file_path, output_folder)

def setup_obsidian(output_folder):
    # output_folder = r'ObsidianNotebook'  # Default output path
    appearance = {
        "theme": "obsidian",
        "accentColor": "#5c6ef5",
        "interfaceFontFamily": "Times New Roman",
        "textFontFamily": "Times New Roman",
        "monospaceFontFamily": "Times New Roman",
        "baseFontSize": 17,
        "baseFontSizeAction": True
    }

    core_plugins = {
        "file-explorer": True,
        "global-search": True,
        "switcher": True,
        "graph": True,
        "backlink": True,
        "canvas": True,
        "outgoing-link": True,
        "tag-pane": True,
        "properties": False,
        "page-preview": True,
        "daily-notes": True,
        "templates": True,
        "note-composer": True,
        "command-palette": True,
        "slash-command": False,
        "editor-status": True,
        "bookmarks": True,
        "markdown-importer": False,
        "zk-prefixer": False,
        "random-note": False,
        "outline": True,
        "word-count": True,
        "slides": False,
        "audio-recorder": False,
        "workspaces": False,
        "file-recovery": True,
        "publish": False,
        "sync": False
    }



    # Path to the .obsidian folder and the appearance.json file
    obsidian_folder = os.path.join(output_folder, '.obsidian')
    appearance_file = os.path.join(obsidian_folder, 'appearance.json')
    core_plugins_file = os.path.join(obsidian_folder, 'core_plugins.json')

    # Check if the .obsidian folder exists, if not, create it
    if not os.path.exists(obsidian_folder):
        os.makedirs(obsidian_folder)
        print(f".obsidian folder created at {obsidian_folder}")
    else:
        print(f".obsidian folder already exists at {obsidian_folder}")

    # Check if appearance.json exists in the .obsidian folder
    if not os.path.exists(appearance_file):
        # Create the appearance.json file and add the appearance data
        with open(appearance_file, 'w') as f:
            json.dump(appearance, f, indent=4)
        print(f"appearance.json created at {appearance_file}")
    if not os.path.exists(core_plugins_file):
        # Create the appearance.json file and add the appearance data
        with open(core_plugins_file, 'w') as f:
            json.dump(core_plugins, f, indent=4)
        print(f"core_plugins.json created at {core_plugins_file}")



def Usage():
    print("\nDescription: " + description)
    print(sys.argv[0] +" Version: %s by %s" % (version, author ))
    #~ print("\nExample:)"
    print("\t" + sys.argv[0] +" -c")
    print("\t" + sys.argv[0] +" -c -I C:\Forensics\scripts\python\Files -O ObsidianNotebook") 
    print("\t" + sys.argv[0] +" -c -O ObsidianNotebook") 
    print("\t" + sys.argv[0] +" -c -I test_files -O ObsidianNotebook") 
         
    

if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>


"""

0.1.2 - create obsidian config files if they don't exist
0.1.1 - Convert the content of .txt, .pdf, .docx, and .eml to Markdown, for use in Obsidian
0.0.9 - converted to template version
0.0.1 - created by ChatGPT
"""


# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
get_file_timestamps = creation_time, access_time, modified_time

if .obsidian doesn't exist in create it
if app.json doesn't exist, create it and add contents
if appearance.json doesn't exist, create it and add contents

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
