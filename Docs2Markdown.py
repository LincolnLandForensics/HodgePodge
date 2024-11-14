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
from datetime import datetime
try:
    import textract # pip install textract
except Exception as e:
    print(f"extract module is not installed': {e}")


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>
author = 'LincolnLandForensics'
description = "Convert the content of .txt, .pdf, .docx, and .eml to Markdown, for use in Obsidian"
version = '0.1.6'

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

 
def extract_text_from_pdf(file_path):
    """
    Function to extract text from a PDF file.
    """
    try:
        doc = fitz.open(file_path)
        # text = f"\n## file_path: {file_path}\n"
        # text = f"\n## file_path: {file_path}\n"
        (creation_time, access_time, modified_time) = get_file_timestamps(file_path)    # test
        text = (f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n")                    
        
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            text += page.get_text("text")  # Extract text
        return text
    except fitz.EmptyFileError:
        print(f"{color_red}Cannot open empty or corrupted PDF file:{color_reset} {file_path}")
        # logging.error(f"Cannot open empty or corrupted PDF file: {file_path}")
        return ""
    except Exception as e:
        print(f"{color_red}Error reading PDF file {color_reset}'{file_path}': {e}")
        # logging.error(f"Error reading PDF file '{file_path}': {e}")
        return ""


def extract_text_from_doc(file_path):
    """
    Function to extract text from .doc files using textract.
    """
    text = ""
    (creation_time, access_time, modified_time) = get_file_timestamps(file_path)
    
    # print(f"{color_purple}Trying {file_path}{color_reset}")
    
    try:
        # Use textract to process the .doc file
        text = textract.process(file_path).decode('utf-8')
    except Exception as e:
        print(f"{color_red}Error reading .doc file {color_reset}'{file_path}': {e}")

    # Add metadata header
    header = (f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n")
    text = header + text
    
    return text
    
    
def extract_text_from_docx(file_path):
    '''
    Function to extract text from DOCX file
    '''    
    (creation_time, access_time, modified_time) = get_file_timestamps(file_path)
    text = ""
    try:
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            text += para.text + "\n"
    except Exception as e:
        print(f"{color_red}Error reading docx file {color_reset}'{file_path}': {e}")

    text = (f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n{text}")                    

    return text


def extract_text_from_eml(file_path):
    '''
    Function to extract text from EML file
    ''' 
    (creation_time, access_time, modified_time) = get_file_timestamps(file_path)    # test

    with open(file_path, 'r', encoding='utf-8') as f:
        msg = email.message_from_file(f, policy=default)

    text = f"\n## File: {file_path}\n## Creation: {creation_time}\n## Modified: {modified_time}\n\n\nSubject: {msg['subject']}\nFrom: {msg['from']}\nTo: {msg['to']}\n\n"

    # Extract email body
    if msg.is_multipart():
        for part in msg.iter_parts():
            if part.get_content_type() == "text/plain":
                text += part.get_content()
    else:
        text += msg.get_content()
    
    return text

def parse_text_file(file_path):
    """
    Parses plain text files (.txt, .md, .cmd, .py), handles encoding errors,
    and returns the content as a string formatted for Markdown.
    """
    creation_time, access_time, modified_time = get_file_timestamps(file_path)
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
            return None

    if extension.lower() in (".py", ".cmd", ".sh", ".bat", ".ps1", ".vbs"):
        return header + "\n\n\n" + "```" +content + "```" + "\n\n\n"
    else:
        return header + content + "\n"


def convert_to_markdown(file_path, output_folder):
    '''
    Function to convert the content of .txt, .pdf, .docx, and .eml to Markdown
    ''' 
    file_name = os.path.basename(file_path)
    base_name, extension = os.path.splitext(file_name)
    target_file_path = os.path.join(output_folder, base_name + ".md")

    if extension.lower() in (".py", ".txt", ".md", ".cmd", ".sh", ".bat", ".ps1", ".vbs"):
        content = parse_text_file(file_path)
    elif file_path.lower().endswith(".pdf"):    
        content = extract_text_from_pdf(file_path) 
    elif extension.lower() in (".docx"):
        content = extract_text_from_docx(file_path) 
    # elif extension.lower() in (".doc"):
        # content = extract_text_from_doc(file_path)
    elif file_path.lower().endswith(".eml"):
        content = extract_text_from_eml(file_path)
    else:
        print(f"{color_red}Unsupported file type: {color_reset}{file_path}")
        return
    
    # Write the content to a Markdown file
    with open(target_file_path, 'w', encoding='utf-8') as md_file:
        md_file.write(content)
    
    print(f"{color_green}Converted {color_reset}'{file_name}' to '{base_name}.md'.")


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
            if extension.lower() in (".py", ".txt", ".md", ".cmd", ".txt", ".pdf", ".docx", ".eml", ".sh", ".bat", ".ps1", ".vbs"):
                convert_to_markdown(file_path, output_folder)
    
    msg_blurb = (f'See {output_folder}')
    msg_blurb_square(msg_blurb, color_green)    


def setup_obsidian(output_folder):
    # output_folder = r'ObsidianNotebook'  # Default output path

    app = {
      "alwaysUpdateLinks": True,
      "newFileLocation": "current",
      "newLinkFormat": "relative",
      "showUnsupportedFiles": True,
      "attachmentFolderPath": "Images"
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



    # Path to the .obsidian folder and the appearance.json file
    obsidian_folder = os.path.join(output_folder, '.obsidian')
    images_folder = os.path.join(output_folder, 'Images')
    templates_folder = os.path.join(output_folder, 'Templates')

    app_file = os.path.join(obsidian_folder, 'app.json')
    appearance_file = os.path.join(obsidian_folder, 'appearance.json')
    community_plugins_file = os.path.join(obsidian_folder, 'community-plugins.json')

    core_plugins_file = os.path.join(obsidian_folder, 'core-plugins.json')

    # Check if the .obsidian folder exists, if not, create it
    if not os.path.exists(obsidian_folder):
        os.makedirs(obsidian_folder)
        print(f".obsidian folder created at {obsidian_folder}")

    # Check if the Images folder exists, if not, create it
    if not os.path.exists(images_folder):
        os.makedirs(images_folder)
        print(f"Images folder created at {images_folder}")

    # Check if the Templates folder exists, if not, create it
    if not os.path.exists(templates_folder):
        os.makedirs(templates_folder)
        print(f"Templates folder created at {templates_folder}")
        

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

def Usage():
    file = sys.argv[0].split('\\')[-1]

    print(f'\nDescription: {color_green}{description}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print("    {file} -b")
    print("    {file} -c")
    print("    {file} -c -I C:\Forensics\scripts\python\Files -O ObsidianNotebook") 
    print("    {file} -c -O ObsidianNotebook") 
    print("    {file} -c -I test_files -O ObsidianNotebook") 


if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>


"""
0.1.4 - (creation_time, access_time, modified_time) = get_file_timestamps(file_path)
0.1.3 - -t option to create a default obsidian folder/files
0.1.2 - create obsidian config files if they don't exist
0.1.1 - Convert the content of .txt, .pdf, .docx, and .eml to Markdown, for use in Obsidian
0.0.9 - converted to template version
0.0.1 - created by ChatGPT
"""


# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
create seperate module to parse text based files such as .txt, .py, .md, and .cmd

The script’s Windows-specific checks (e.g., file creation time) could potentially be improved for cross-platform compatibility. You might use pathlib for better handling of file paths across platforms.

this will overwrite files with the same name, create a way of if file exists, save it with a unique name

import additional extensions such as .py



"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
