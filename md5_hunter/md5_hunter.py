
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import re
import pdfplumber   # pip install pdfplumber
import docx # pip install python-docx
from datetime import datetime

'''
read pdf's and other specific filetypes and export out unique md5 hashes
and words for password cracking and file hunting.
'''

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "read files and pull out md5 hashes and words"
version = '0.1.2'

input_folder = "md5_hashes"

file_types = [
    '.csv', '.docx', '.eml', '.htm', '.html', '.ini', '.json', '.log',
    '.md', '.msg', '.pdf', '.php', '.pptx', '.py', '.rtf', '.sql', '.tsv', '.txt',
    '.xlsx', '.xml'
]

plain_text = [
    ".bat", ".cmd", ".csv", ".ini", ".json", ".log", ".md", 
    ".py", ".sh", ".ps1", ".txt", ".vbs", ".xml", ".yml", ".yaml"
]



# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def check_folder(folder_path):
    """Check if the folder exists, and create it if it does not."""
    if not os.path.exists(folder_path):
        # os.makedirs(folder_path)
        # print(f"Folder '{folder_path}' was created.")
        print(f"{folder_path} folder doesnt exist. create one, add files and retry this")
        exit()
    else:
        print(f"reading files in {folder_path} folder.")

def extract_text_from_pdf(file_path):
    """Extracts text from a PDF file using pdfplumber."""
    text = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            if page.extract_text():
                text += page.extract_text()
                text += "\n"
            # For tables
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Replace None values with empty strings
                    sanitized_row = [str(cell) if cell is not None else '' for cell in row]
                    text += '\t'.join(sanitized_row) + '\n'
    return text

def extract_text_from_docx(file_path):
    """Extracts text from a DOCX file using python-docx."""
    text = ''
    doc = docx.Document(file_path)
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_text_from_txt(text):
    """Process the extracted text, split by space, clean, and return unique sorted words."""
    words = text.split()  # Split by spaces to get a list of words
    cleaned_words = [re.sub(r'[.,]$', '', word) for word in words]  # Remove commas and periods
    return set(cleaned_words)  # Return a set of unique words


def find_md5_hashes(words):
    """Find and return all MD5 hashes in the list of words."""
    md5_pattern = re.compile(r'\b[a-f0-9]{32}\b', re.IGNORECASE)  # Regex for MD5 hashes
    return {word for word in words if md5_pattern.match(word)}


def process_files(input_folder, file_types, plain_text):
    words = set()  # Set to store unique words

    # Walk through the directory
    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder, filename)

        # Check if the file has a valid extension
        if os.path.isfile(file_path):
            file_extension = os.path.splitext(filename)[1].lower()

            # If the file matches plain text types
            if file_extension in plain_text:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                    text = file.read()
                    words.update(extract_text_from_txt(text))

            # If the file is a PDF
            elif file_extension == '.pdf':
                text = extract_text_from_pdf(file_path)
                words.update(extract_text_from_txt(text))

            # If the file is a DOCX
            elif file_extension == '.docx':
                text = extract_text_from_docx(file_path)
                words.update(extract_text_from_txt(text))

    return sorted(words)


def save_words_to_file(words, filename):
    """Save the words to a file, one word per line."""
    with open(filename, 'w', encoding='utf-8') as f:
        for word in words:
            f.write(word + '\n')

# Check and create the input folder if it doesn't exist
check_folder(input_folder)

# Process all files in the input folder and get sorted unique words
unique_sorted_words = process_files(input_folder, file_types, plain_text)

# Find all MD5 hashes in the unique sorted words
unique_sorted_md5 = find_md5_hashes(unique_sorted_words)

# Sort the MD5 hashes
unique_sorted_md5 = sorted(unique_sorted_md5)

# Get the current date to include in the filenames
current_date = datetime.now().strftime('%Y-%m-%d')

# Save the unique sorted words to a file
words_filename = f"words_{current_date}.txt"
save_words_to_file(unique_sorted_words, words_filename)

# Save the unique sorted MD5 hashes to a file
md5_filename = f"md5_hashes_{current_date}.txt"
save_words_to_file(unique_sorted_md5, md5_filename)

# Print paths of the saved files
print(f"Saved {len(unique_sorted_words)} unique sorted words to: {words_filename}")
print(f"Saved {len(unique_sorted_md5)} unique sorted MD5 hashes to: {md5_filename}")



# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

0.1.1 - working prototype
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
menu system so you can specify a new input_folder
count unique words and md5's , print that to screen.
html parser
test each file type
other hashes like sha256?
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
Is it only checking specfied file types?


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>

test = 'test'

