#!/usr/bin/python
# coding: utf-8


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>
import os
import re
import sys
import fitz  # pip install PyMuPDF
import json
import docx
import email
import time
import shutil
import hashlib
import exifread
from email.policy import default
from striprtf.striprtf import rtf_to_text   # pip install striprtf

import argparse # for menu system
import msg_parser   # pip install msg_parser, extract-msg

from datetime import datetime
# try:
    # import textract # pip install textract
# except Exception as e:
    # print(f"extract module is not installed': {e}")
from markdownify import markdownify as md   # pip install markdownify

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "Convert the content of various file types to Markdown, for use in Obsidian"
version = '0.2.0'


global file_types
file_types = [
    '.bash', '.bat', '.bmp', '.c', '.cmd', '.cpp', '.cs', '.css', '.csv',
    '.docx', '.drawio', '.eml', '.flac', '.gif', '.go', '.heic', '.heif', '.htm', '.html',
    '.ini', '.java', '.jl', '.jpeg', '.jpg', '.js', '.json', '.kt', '.log',
    '.m', '.m4a', '.md', '.mermaid', '.mkv', '.mov', '.mp3', '.mp4', '.msg',
    '.ogg', '.ogv', '.pdf', '.php', '.png', '.ppt', '.pptx', '.ps1', '.py', '.r', '.rb',
    '.rs', '.rtf', '.sh', '.sql', '.svg', '.swift', '.tif', '.tiff', '.ts', '.tsv', '.txt',
    '.vbs', '.wav', '.webm', '.xlsx', '.xml', '.yaml', '.yml'
]

global files_docs
files_docs = [
    '.csv', '.docx', '.drawio', '.htm', '.html', '.json', '.md', '.mermaid', '.pdf', '.ppt', '.pptx',
    '.ps1', '.py', '.rtf', '.sh', '.tsv', '.txt', '.vbs', '.xlsx', '.xml',
    '.yaml', '.yml'
]

global plain_text
plain_text = [
    ".bat", ".cmd", ".csv", ".ini", ".json", ".log", ".md", 
    ".py", ".sh", ".ps1", ".txt", ".vbs", ".xml", ".yml", ".yaml"
]

global files_media
files_media = [
	'.bmp', '.flac', '.gif', '.heic', '.heif', '.jpeg', '.jpg', '.m4a', '.mkv', '.mov', '.mp3',
    '.mp4', '.ogg', '.ogv', '.png', '.svg', '.tiff', '.tif', '.wav', '.webm', '.webp'
]

global files_scripts
files_scripts = [
    '.bash', '.bat', '.c', '.cmd', '.go', '.java', '.js', '.m', '.php', '.ps1',
    '.py', '.r', '.rb', '.rs', '.sh', '.sql', '.swift', '.vbs', '.yaml', '.yml'
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

if sys.version_info > (3, 7, 9) and os.name == "nt":
    version_info = os.sys.getwindowsversion()
    major_version = version_info.major
    build_version = version_info.build

    if major_version >= 10 and build_version >= 22000: # Windows 11 and above
        import colorama
        from colorama import Fore, Back, Style  
        print(f'{Back.BLACK}') # make sure background is black
        color_red = Fore.RED
        color_yellow = Fore.YELLOW
        color_green = Fore.GREEN
        color_blue = Fore.BLUE
        color_purple = Fore.MAGENTA
        color_reset = Style.RESET_ALL
        
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    """
    Main function to parse arguments and initiate file conversion.
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='Input folder path', required=False)
    parser.add_argument('-O', '--output', help='Output folder path', required=False)
    parser.add_argument('-c', '--convert', help='Convert files to markdown', required=False, action='store_true')
    parser.add_argument('-b', '--blank', help='create a blank obsidian folder', required=False, action='store_true')

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
    elif args.blank:
        setup_obsidian(output_folder)
    else:
        parser.print_help()
        Usage()

    return 0


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def cleanup_description(description):
    # Filter out keys with empty or None values, and exclude 'Icon' and 'Type'
    cleaned_description = {k: v for k, v in description.items() if v not in ('', None) and k not in ('Icon', 'Type', 'Latitude', 'Longitude', 'Time', 'FileTypeExtension', 'Name')}
    
    # Create a formatted string with key-value pairs
    formatted_description = "\n".join(f"{k}: {v}" for k, v in cleaned_description.items())
    
    return formatted_description


def convert_to_decimal(coord, ref):
    """Converts GPS coordinates to decimal format."""
    if not coord:
        return None
    degrees = float(coord[0].num) / float(coord[0].den)
    minutes = float(coord[1].num) / float(coord[1].den) / 60.0
    seconds = float(coord[2].num) / float(coord[2].den) / 3600.0
    decimal = degrees + minutes + seconds
    if ref in ['S', 'W']:
        decimal = -decimal
    return decimal


def convert_to_markdown(file_path, output_folder):
    '''
    Function to convert the content of .txt, .pdf, .docx, and .eml to Markdown 
    ''' 
    file_name = os.path.basename(file_path)
    base_name, extension = os.path.splitext(file_name)
    target_file_path = os.path.join(output_folder, base_name + ".md")


    if file_path.lower().endswith(".pdf"):    
        content = extract_text_from_pdf(file_path) 
    elif extension.lower() in (files_media):
        content = md_for_media(file_path)

    elif extension.lower() in (".docx"):
        content = extract_text_from_docx_with_formatting(file_path)
    elif file_path.lower().endswith(".eml"):
        content = extract_text_from_eml(file_path)
    elif file_path.lower().endswith(".msg"):
        content = extract_text_from_msg(file_path)
    elif file_path.lower().endswith(".rtf"):
        content = convert_rtf(file_path)
    elif file_path.lower().endswith(('.html', '.htm')):
        content = extract_html(file_path)
    elif extension.lower() in (files_scripts):
        content = md_for_scripts(file_path)

    elif extension.lower() in plain_text:
        content = parse_text_file(file_path)

    else:
        print(f"{color_red}Unsupported file type: {color_reset}{file_path}")
        return
    
    # Write the content to a Markdown file
    with open(target_file_path, 'w', encoding='utf-8') as md_file:
        md_file.write(content)
    # try:
        # with open(target_file_path, 'w', encoding='utf-8') as md_file:
            # md_file.write(content)
    # except Exception as e:
        # return f"Error writing {file_path}: {e}"
    
    print(f"{color_green}Converted {color_reset}'{file_name}' to '{base_name}.md'.")


def convert_rtf(file_path):
    """
    Converts an RTF file to Markdown format by extracting the plain text.
    Formatting will be basic (like bold, italics, etc.), but more advanced formatting would need
    manual conversion.
    """
    try:
        hash_md5 = hashfile(file_path)
    except:
        hash_md5 = "" 


    # Read the RTF file
    try:
        with open(file_path, 'r', encoding='utf-8') as rtf_file:
            rtf_content = rtf_file.read()
    except Exception as e:
        return f"Error reading RTF file: {e}"

    # Convert the RTF content to plain text
    plain_text = rtf_to_text(rtf_content)

    # Convert plain text into Markdown:
    text = convert_rtf_to_md2(plain_text)
    (creation_time, access_time, modified_time) = get_file_timestamps(file_path)
    file_name = os.path.basename(file_path)
    
    new_file_path = os.path.join(docs_folder, file_name)
    # Copy the file to the new path
    shutil.copy(file_path, new_file_path)
    
    new_file_path_link = os.path.join('assets/documents/', file_name) 
    new_file_path_link = new_file_path_link.replace("\\", "/")    
    
    
    # text = (f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n{text}")                    
    text = (f"\n## [File]({new_file_path_link}): {file_path}\n"
                f"## Creation: {creation_time}\n"
                f"## Modified: {modified_time}\n"
                f"## MD5: {hash_md5}\n\n"                
                f"{text}\n\n"
                )        
    return text


def convert_rtf_to_md2(plain_text):
    """
    Converts the plain text extracted from the RTF to basic Markdown.
    This function is designed for simple formatting.
    """
    # Example: Convert any instances of **bold** text to Markdown
    markdown_content = plain_text
    markdown_content = re.sub(r'\\b (.*?) \\b0', r'**\1**', markdown_content)  # Bold text
    markdown_content = re.sub(r'\\i (.*?) \\i0', r'*\1*', markdown_content)  # Italics text
    markdown_content = re.sub(r'\\ul (.*?) \\ulnone', r'_\1_', markdown_content)  # Underline text
    markdown_content = re.sub(r'•\s*(.*?)\n', r'- \1\n', markdown_content)  # bullet lists
    markdown_content = re.sub(r'\\pard (.*?) \\pard0', r'\1', markdown_content)  # Paragraphs
    markdown_content = re.sub(r'\\pard\s*(.*?)\\pard0', r'\1\n\n', markdown_content)    # Convert paragraphs
    markdown_content = markdown_content.strip() # Clean up extra spaces and formatting errors

    try:
        markdown_content = re.sub(r'**(.*?)**\n', r'# \1\n', markdown_content)  # simple bold headings
    except:
        pass

    return markdown_content

    
def extract_text_from_docx_with_formatting(file_path):
    '''
    Function to extract formatted text from DOCX file and return in Markdown format with metadata.
    '''
    # Get file timestamps (creation, access, modified)
    creation_time, access_time, modified_time = get_file_timestamps(file_path)

    try:
        hash_md5 = hashfile(file_path)
    except:
        hash_md5 = ""    

    text = ""

    try:
        # Copy the doc to the assets folder

        file_name = os.path.basename(file_path)
        new_file_path_link = os.path.join('assets/documents/', file_name)
        # new_file_path_link = os.path.join(docs_folder, file_name)         
        
        new_file_path_link = new_file_path_link.replace("\\", "/")
        # new_file_path = os.path.join(assets_folder, file_name)
        new_file_path = os.path.join(docs_folder, file_name) # fixme

        # Copy the file to the new path
        shutil.copy(file_path, new_file_path)

        # Initialize the text output with metadata and the file link
        text = (f"\n## [File]({new_file_path_link}): {file_path}\n"
                f"## Creation: {creation_time}\n"
                f"## Modified: {modified_time}\n"
                f"## MD5: {hash_md5}\n\n"                
                )

    except Exception as e:
        print(f"{color_red}Error copying document {color_reset}'{file_path}': {e}")

    try:
        # Load DOCX file
        doc = docx.Document(file_path)
        
        # Add metadata to the extracted text
        # text += f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n"
        
        # Extract formatted text from paragraphs
        for para in doc.paragraphs:
            para_text = handle_paragraph_formatting(para)
            text += para_text + "\n"
    
    except Exception as e:
        print(f"{color_red}Error reading docx file {color_reset}'{file_path}': {e}")
    
    # Return the extracted text with metadata and Markdown formatting
    return text


def extract_text_from_eml(file_path):
    """
    Function to extract text from EML file, including metadata and email content.
    """
    # Get file timestamps
    creation_time, access_time, modified_time = get_file_timestamps(file_path)
    try:
        hash_md5 = hashfile(file_path)
    except:
        hash_md5 = ""  

    # Read the EML file
    with open(file_path, 'r', encoding='utf-8') as f:
        msg = email.message_from_file(f, policy=default)

    # Extract metadata
    text = (
        f"\n## File: {file_path}\n"
        f"## Creation: {creation_time}\n"
        f"## Modified: {modified_time}\n"
        f"## MD5: {hash_md5}\n\n"
        f"Subject: {msg['subject']}\n"
        f"From: {msg['from']}\n"
        f"To: {msg['to']}\n\n"
    )

    # Extract email body
    if msg.is_multipart():
        for part in msg.iter_parts():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                text += part.get_content()
            elif content_type == "text/html":
                text += f"\n[HTML Content Skipped for Readability]\n"
    else:
        text += msg.get_content()

    return text


def extract_html(file_path):

    (creation_time, access_time, modified_time) = get_file_timestamps(file_path)    # test
    try:
        hash_md5 = hashfile(file_path)
    except:
        hash_md5 = "" 
        
    try:
        # Copy the doc to the assets folder

        file_name = os.path.basename(file_path)

        new_file_path_link = os.path.join('assets', file_name) 
        new_file_path_link = new_file_path_link.replace("\\", "/")
        # new_file_path = os.path.join(assets_folder, file_name)
        new_file_path = os.path.join(docs_folder, file_name)

        # Copy the file to the new path
        shutil.copy(file_path, new_file_path)

        # Initialize the text output with metadata and the file link
        text = (f"\n## [File]({new_file_path_link}): {file_path}\n"
                f"## Creation: {creation_time}\n"
                f"## Modified: {modified_time}\n"
                f"## MD5: {hash_md5}\n\n"
                )
    except Exception as e:
        print(f"{color_red}Error copying document {color_reset}'{file_path}': {e}")


    # Read the HTML content from the file
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            html_content = file.read()

        # Convert the HTML content to Markdown using markdownify
        text = (f'({text}\n{md(html_content)}')
 
    except UnicodeDecodeError:
        print(f"{color_red}Cannot decode {color_reset} {file_path}.") 

    return text
  
  
def extract_text_from_msg(file_path):
    """
    Function to extract text from MSG files.
    """
    from extract_msg import Message

    # Get file timestamps
    creation_time, access_time, modified_time = get_file_timestamps(file_path)

    try:
        hash_md5 = hashfile(file_path)
    except:
        hash_md5 = "" 
        
    # Parse the .msg file
    msg = Message(file_path)

    # Extract metadata and message content
    text = (
        f"\n## File: {file_path}\n"
        f"## Creation: {creation_time}\n"
        f"## Modified: {modified_time}\n"
        f"## MD5: {hash_md5}\n\n"        
        f"Subject: {msg.subject}\n"
        f"From: {msg.sender}\n"
        f"To: {msg.to}\n\n"
        f"Body:\n{msg.body}\n"
    )

    # Replace 'From: None' with 'From: '
    text = text.replace('From: None', 'From: ')
    
    return text

    
def extract_text_from_pdf(file_path):
    """
    Function to extract text from a PDF file, copy it to a new location in the assets folder,
    and generate a link to the original file.
    """
    try:
        # Open the PDF file
        doc = fitz.open(file_path)

        # Get file timestamps
        creation_time, access_time, modified_time = get_file_timestamps(file_path)

        # If timestamps are not found, return empty text
        if creation_time is None or access_time is None or modified_time is None:
            return ""

        # Copy the PDF to the assets folder
        if not os.path.exists(assets_folder):
            os.makedirs(assets_folder)  # Create the assets folder if it doesn't exist

        file_name = os.path.basename(file_path)

        new_file_path_link = os.path.join('assets/documents', file_name) 
        new_file_path_link = new_file_path_link.replace("\\", "/")
        new_file_path = os.path.join(docs_folder, file_name)

        # Copy the file to the new path
        shutil.copy(file_path, new_file_path)

        # Initialize the text output with metadata and the file link
        text = (f"\n## [File]({new_file_path_link}): {file_path}\n"
                f"## Creation: {creation_time}\n"
                f"## Modified: {modified_time}\n\n")

        # Extract text from each page
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            page_text = page.get_text("text")  # Extract plain text
            text += page_text

        return text

    except fitz.EmptyFileError:
        print(f"Cannot open empty or corrupted PDF file: {file_path}")
        # logging.error(f"Cannot open empty or corrupted PDF file: {file_path}")
        return ""
    except Exception as e:
        print(f"Error reading PDF file '{file_path}': {e}")
        # logging.error(f"Error reading PDF file '{file_path}': {e}")
        return ""

def get_file_timestamps(file_path):
    """
    Returns the creation, last accessed, and last modified times of a file in YYYY-MM-DD hh:mm:ss format.

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

    # Convert from timestamps to human-readable YYYY-MM-DD hh:mm:ss format
    creation_time = datetime.fromtimestamp(creation_time).strftime('%Y-%m-%d %H:%M:%S')
    access_time = datetime.fromtimestamp(access_time).strftime('%Y-%m-%d %H:%M:%S')
    modified_time = datetime.fromtimestamp(modified_time).strftime('%Y-%m-%d %H:%M:%S')

    return creation_time, access_time, modified_time


def handle_paragraph_formatting(para):
    '''
    Convert paragraph content to Markdown format considering styles.
    '''
    para_text = ""
    
    # Check if the paragraph is a heading (Heading 1, Heading 2, etc.)
    if para.style.name.startswith('Heading'):
        level = int(para.style.name.split()[-1])  # Extract level of the heading
        para_text += "#" * level + " " + para.text
    else:
        # For normal paragraphs, handle runs (bold, italic, etc.)
        for run in para.runs:
            text = run.text
            
            # Bold text
            if run.bold:
                text = f"**{text}**"
                
            # Italic text
            if run.italic:
                text = f"*{text}*"
            
            para_text += text

    return para_text    


def hashfile(file_path):
    """
    Computes and returns the MD5 hash of a file.
    """
    # Create an MD5 hash object
    hash_md5 = hashlib.md5()

    # Open the file in binary read mode
    with open(file_path, "rb") as f:
        # Read the file in chunks to avoid memory overload with large files
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    
    # Return the hexadecimal representation of the hash
    return hash_md5.hexdigest()
    
  
def parse_text_file(file_path):
    """
    Parses plain text files (.txt, .md, .cmd, .py), handles encoding errors,
    and returns the content as a string formatted for Markdown.
    """

    file_name = os.path.basename(file_path)

    creation_time, access_time, modified_time = get_file_timestamps(file_path)
    new_file_path_link = os.path.join('assets/documents/', file_name) 
    new_file_path_link = new_file_path_link.replace("\\", "/")

    new_file_path = os.path.join(docs_folder, file_name)
    
    # Copy the file to the new path
    shutil.copy(file_path, new_file_path)

    header = f"\n## [File]({new_file_path_link}): {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n"

    header = f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n"
    base_name, extension = os.path.splitext(file_path)
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, 'r', encoding='ISO-8859-1') as f:
                content = f.read()
        except UnicodeDecodeError:
            print(f"{color_red}Cannot decode {color_reset} {file_path} with UTF-8 or ISO-8859-1.") 
            # return None

    if extension.lower() in (".py"):
        return header + "\n\n\n" + "```python\n" +content + "```" + "\n\n\n"
        # Check if the .obsidian folder exists, if not, create it
        if not os.path.exists('scripts'):
            os.makedirs('scripts')
            print(f"scripts folder created at scripts")
  
    elif extension.lower() in (".ps1"):
        return header + "\n\n\n" + "```powershell\n" +content + "\n```" + "\n\n\n"
    elif extension.lower() in (".sh"):
        return header + "\n\n\n" + "```bash\n" +content + "\n```" + "\n\n\n"
    elif extension.lower() in (".cmd", ".bat"):
        return header + "\n\n\n" + "```cmd\n" +content + "\n```" + "\n\n\n"

    elif extension.lower() in (".vbs"):
        return header + "\n\n\n" + "```vbs\n" +content + "\n```" + "\n\n\n"
    else:
        return header + content + "\n"

    
def md_for_media(file_path):
    """
    copy media files to a new location in the media folder,
    and generate a link to the original file.
    """
    Description = ''
    try:
        # Get file timestamps
        creation_time, access_time, modified_time = get_file_timestamps(file_path)

        # If timestamps are not found, return empty text
        if creation_time is None or access_time is None or modified_time is None:
            print('error converting time stamp')
            # return ""

    except Exception as e:
        print(f"Error making file '{file_path}': {e}")
        
    try:
        hash_md5 = hashfile(file_path)
    except:
        hash_md5 = ""

    file_name = os.path.basename(file_path)

    new_file_path_link = os.path.join('assets/media/', file_name) 
    new_file_path_link = new_file_path_link.replace("\\", "/")

        
    new_file_path = os.path.join(media_folder, file_name)

    # Copy the file to the new path
    shutil.copy(file_path, new_file_path)

    try:
        if file_path.lower().endswith(('.heic', '.heif', '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.webp')):
            exif_data, Description = read_exif_data(file_path)

    except Exception as e:
        print(f"{color_red}Error getting exif data '{file_path}'{color_reset}: {e}")
        

    if Description != '':
        Description = (f'## Exif Data\n{Description}')

    # Initialize the text output with metadata and the file link
    text = (f"\n### [File]({new_file_path_link}): {file_path}\n"
        f"### Creation: {creation_time}\n"
        f"### Modified: {modified_time}\n"
        f"### MD5: {hash_md5}\n\n"                
        f"### ![[{new_file_path_link}|400]]\n\n"
        f"{Description}\n"
        )

    return text
 

def md_for_scripts(file_path):
    """
    copy scripts files to a new location in the scripts folder,
    and generate a link to the original file.
    """
    try:
        hash_md5 = hashfile(file_path)
    except:
        hash_md5 = "" 

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, 'r', encoding='ISO-8859-1') as f:
                content = f.read()
        except UnicodeDecodeError:
            print(f"{color_red}Cannot decode {color_reset} {file_path} with UTF-8 or ISO-8859-1.") 
            # return None




    try:
        # Get file timestamps
        creation_time, access_time, modified_time = get_file_timestamps(file_path)

        # If timestamps are not found, return empty text
        if creation_time is None or access_time is None or modified_time is None:
            return ""


        base_name, extension = os.path.splitext(file_path)


        file_name = os.path.basename(file_path)

        new_file_path_link = os.path.join('assets/scripts/', file_name) 
        # new_file_path_link = os.path.join(scripts_folder, file_name) 
        new_file_path_link = new_file_path_link.replace("\\", "/")
        
        new_file_path = os.path.join(scripts_folder, file_name)

        # Copy the file to the new path
        shutil.copy(file_path, new_file_path)

        # don't put a link in, it will run the script
        # header = f"\n## [File]({new_file_path_link}): {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n"

        header = f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n"

    
        # new 
        if extension.lower() in (".py"):
            return header + "\n\n\n" + "```python\n" +content + "```" + "\n\n\n"

        elif extension.lower() in (".ps1"):
            return header + "\n\n\n" + "```powershell\n" +content + "\n```" + "\n\n\n"
        elif extension.lower() in (".sh"):
            return header + "\n\n\n" + "```bash\n" +content + "\n```" + "\n\n\n"
        elif extension.lower() in (".cmd", ".bat"):
            return header + "\n\n\n" + "```cmd\n" +content + "\n```" + "\n\n\n"

        elif extension.lower() in (".vbs"):
            return header + "\n\n\n" + "```vbs\n" +content + "\n```" + "\n\n\n"
        else:
            return header + content + "\n"


        print(f'hello world')   # temp

        # Initialize the text output with metadata and the file link
        text = (f"\n## [File]({new_file_path_link}): {file_path}\n"
                f"## Creation: {creation_time}\n"
                f"## Modified: {modified_time}\n"
                f"## MD5: {hash_md5}\n\n"
                f"## header test: \n{header}\n\n"
                
                f"## ![[{new_file_path_link}]]\n\n")

        return text
    except Exception as e:
        print(f"Error making file '{file_path}': {e}")
        return ""    


def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')
    
    
def process_files(root_folder, output_folder):
    '''
    Function to walk through all files and convert
    ''' 
    
    setup_obsidian(output_folder)   # create a default obsidian setup, if it doesn't exist

    # Walk through all files and subdirectories in root_folder
    for subdir, _, files in os.walk(root_folder):
        for file in files:
            file_path = os.path.join(subdir, file)
            base_name, extension = os.path.splitext(file_path)
            # if 1==1:
            if extension.lower() in file_types:
                convert_to_markdown(file_path, output_folder)
    
    msg_blurb = (f'See {output_folder}')
    msg_blurb_square(msg_blurb, color_green)    


def read_exif_data(file_path):
    """Reads selected EXIF data from an image file."""
    with open(file_path, 'rb') as f:
        tags = exifread.process_file(f, stop_tag='UNDEFINED')

    gps_latitude = tags.get('GPS GPSLatitude')
    gps_latitude_ref = tags.get('GPS GPSLatitudeRef')
    gps_longitude = tags.get('GPS GPSLongitude')
    gps_longitude_ref = tags.get('GPS GPSLongitudeRef')
    
    latitude = convert_to_decimal(gps_latitude.values, gps_latitude_ref.values[0]) if gps_latitude and gps_latitude_ref else ""
    longitude = convert_to_decimal(gps_longitude.values, gps_longitude_ref.values[0]) if gps_longitude and gps_longitude_ref else ""

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': tags.get('Image DateTime'),
        'DateTimeOriginal': tags.get('EXIF DateTimeOriginal'),
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': tags.get('EXIF OffsetTime'),
        'DeviceManufacturer': tags.get('Image Make'),
        'LensMake': tags.get('EXIF LensMake'),
        'LensModel': tags.get('EXIF LensModel'),
        'Model': tags.get('Image Model'),
        'Software': tags.get('Image Software'),
        'LensInfo': tags.get('EXIF LensInfo'),
        'HostComputer': tags.get('Image HostComputer'),
        'FileSize': os.path.getsize(file_path),
        'FileType': os.path.splitext(file_path)[1].replace('.', '').upper(),
        'FileTypeExtension': os.path.splitext(file_path)[1].replace('.', '').lower(),
        'Altitude': tags.get('GPS GPSAltitude'),
        'Latitude': latitude,
        'Longitude': longitude,
        'Coordinate': f"{latitude},{longitude}" if latitude and longitude else "",
        'NumberOfImages': tags.get('Exif NumberOfImages'),
        'ExifToolVersion': tags.get('Exif ExifToolVersion'),
        'Icon': 'Images',
        'Type': 'Images',
        'Time': tags.get('EXIF DateTimeOriginal')
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    return exif_data, Description
    
    
def setup_obsidian(output_folder):
    # output_folder = r'ObsidianNotebook'  # Default output path

    
    msg_blurb = (f'See {output_folder}')
    msg_blurb_square(msg_blurb, color_green)


    app = {
      "alwaysUpdateLinks": True,
      "newFileLocation": "current",
      "newLinkFormat": "relative",
      "showUnsupportedFiles": True,
      "attachmentFolderPath": "Assets"
    }

    appearance = {
        "theme": "obsidian",
        "accentColor": "#5cb2f5",
        "interfaceFontFamily": "Times New Roman",
        "textFontFamily": "Times New Roman",
        "monospaceFontFamily": "Times New Roman",
        "baseFontSize": 17,
        "baseFontSizeAction": True
    }

    community_plugins = [
      "obsidian-excalidraw-plugin",
      "dataview",
      "obsidian-icon-folder",
      "obsidian-kanban",
      "calendar",
      "obsidian-tasks-plugin",
      "obsidian-advanced-slides",
      "obsidian-annotator",
      "homepage",
      "omnisearch",
      "periodic-notes",
      "url-into-selection",
      "obsidian-textgenerator-plugin",
      "obsidian-banners",
      "obsidian-plugin-toc",
      "media-extended",
      "settings-search",
      "emoji-shortcodes",
      "smart-connections",
      "obsidian-plugin-update-tracker",
      "janitor"
    ]

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

    templates = {
      "folder": "assets/templates",
      "dateFormat": "YYYY-MM-DD",
      "timeFormat": "HH:mm:ss"
    }


    # Path to the .obsidian folder and the appearance.json file
    obsidian_folder = os.path.join(output_folder, '.obsidian')
    global assets_folder
    assets_folder = os.path.join(output_folder, 'assets')
    global assets_folder2
    assets_folder2 = os.path.join('assets')
    global docs_folder
    docs_folder = os.path.join(output_folder, 'assets\documents')
    global media_folder
    # media_folder = os.path.join(output_folder, 'assets\media')
    media_folder = os.path.join(output_folder, (f'{assets_folder2}\media'))    
    global images_folder

    # images_folder = os.path.join(output_folder, (f'{assets_folder2}\media')) 

    images_folder = os.path.join(output_folder, 'assets\\media\\images')
    global scripts_folder
    scripts_folder = os.path.join(output_folder, 'assets\scripts')
    global templates_folder
    templates_folder = os.path.join(output_folder, 'assets\\templates')

    app_file = os.path.join(obsidian_folder, 'app.json')
    appearance_file = os.path.join(obsidian_folder, 'appearance.json')
    community_plugins_file = os.path.join(obsidian_folder, 'community-plugins.json')
    core_plugins_file = os.path.join(obsidian_folder, 'core-plugins.json')
    template_file = os.path.join(obsidian_folder, 'templates.json')

    # Check if the .obsidian folder exists, if not, create it
    if not os.path.exists(obsidian_folder):
        os.makedirs(obsidian_folder)
        print(f".obsidian folder created at {obsidian_folder}")

    # Check if the assets folder exists, if not, create it
    if not os.path.exists(assets_folder):
        os.makedirs(assets_folder)
        print(f"Assets folder created at {assets_folder}")

    if not os.path.exists(media_folder):
        os.makedirs(media_folder)
        print(f"media folder created at {media_folder}")

    # Check if the templates folder exists, if not, create it
    if not os.path.exists(templates_folder):
        # os.makedirs(templates_folder)
        print(f"templates folder created at {templates_folder}")

    # Check if the Images folder exists, if not, create it
    if not os.path.exists(images_folder):
        os.makedirs(images_folder)
        print(f"Images folder created at {images_folder}")

    # Check if the documents folder exists, if not, create it
    if not os.path.exists(docs_folder):
        os.makedirs(docs_folder)
        print(f"Documents folder created at {docs_folder}")

    if not os.path.exists(templates_folder):
        os.makedirs(templates_folder)
        print(f"Documents folder created at {templates_folder}")


    # Check if the scripts folder exists, if not, create it
    if not os.path.exists(scripts_folder):
        os.makedirs(scripts_folder)
        print(f"scripts folder created at {scripts_folder}")

    if not os.path.exists(app_file):
        with open(app_file, 'w') as f:
            json.dump(app, f, indent=4)
        print(f"app.json created at {app_file}")
    if not os.path.exists(appearance_file):
        with open(appearance_file, 'w') as f:
            json.dump(appearance, f, indent=4)
        print(f"appearance.json created at {appearance_file}")
    if not os.path.exists(community_plugins_file):
        with open(community_plugins_file, 'w') as f:
            json.dump(community_plugins, f, indent=4)
        print(f"community_plugins.json created at {community_plugins_file}")
    if not os.path.exists(core_plugins_file):
        with open(core_plugins_file, 'w') as f:
            json.dump(core_plugins, f, indent=4)
        print(f"core_plugins.json created at {core_plugins_file}")

    if not os.path.exists(template_file):
        with open(template_file, 'w') as f:
            json.dump(templates, f, indent=4)
        print(f"templates.json created at {template_file}")


def Usage():
    file = sys.argv[0].split('\\')[-1]

    print(f'\nDescription: {color_green}{description}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f"    {file} -b")
    print(f"    {file} -c")
    print(f"    {file} -c -I C:\Forensics\scripts\python\Files -O ObsidianNotebook") 
    print(f"    {file} -c -O ObsidianNotebook") 
    print(f"    {file} -c -I test_files -O ObsidianNotebook") 


if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>


"""



extract_html

0.1.5 - add exifdata to the bottom of the .md file
0.1.4 - (creation_time, access_time, modified_time) = get_file_timestamps(file_path)
0.1.3 - -t option to create a default obsidian folder/files
0.1.2 - create obsidian config files if they don't exist
0.1.1 - Convert the content of .txt, .pdf, .docx, and .eml to Markdown, for use in Obsidian
0.0.9 - converted to template version
0.0.1 - created by ChatGPT
"""


# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

export to html -w
convert_pptx(file_path)
pdf convert data and tables.
undo the try except in .eml multi threaded

fix rtf converter - doesn't do all formatting like bold
create seperate module to parse text based files such as .txt, .py, .md, and .cmd

The script’s Windows-specific checks (e.g., file creation time) could potentially be improved for cross-platform compatibility. You might use pathlib for better handling of file paths across platforms.

this will overwrite files with the same name, create a way of if file exists, save it with a unique name


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
