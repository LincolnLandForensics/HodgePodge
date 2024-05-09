#!/usr/bin/python
# coding: utf-8

'''
This script is used to translate the contents of an Excel spreadsheet from many 
languages to English. It uses the openypyl library to read the input file, requests  
to perform the translation, (googletrans as a backup module) and openypyl to write 
the translated contents to a new Excel file.
'''

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>
import re
import sys
from time import sleep
import os.path
from openpyxl import load_workbook
from requests import get 

import argparse  # for menu system

# from googletrans import Translator  # pip install googletrans
from langdetect import detect   # pip install langdetect


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "Read input_translate.xlsx filled with another language and translate it to english"
version = '1.0.3'

global auto_list
auto_list = ['!','?']

# Colorize section
global color_red
global color_yellow
global color_green
global color_blue
global color_reset
color_red = ''
color_yellow = ''
color_green = ''
color_blue = ''
color_reset = ''

if sys.version_info > (3, 7, 9) and os.name == "nt":
    version_info = os.sys.getwindowsversion()
    major_version = version_info.major
    build_version = version_info.build

    if major_version >= 10 and build_version >= 22000: # Windows 11 and above
        from colorama import Fore, Back, Style  
        print(f'{Back.BLACK}') # make sure background is black
        color_red = Fore.RED
        color_yellow = Fore.YELLOW
        color_green = Fore.GREEN
        color_blue = Fore.BLUE
        color_reset = Style.RESET_ALL

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>
def main():
    output_xlsx = ('translation_.xlsx') 

    check_internet_connection()
    
    parser = argparse.ArgumentParser(description='Translate Excel contents from various languages to English')
    parser.add_argument('-I', '--input', help='Input Excel file', required=False)
    parser.add_argument('-O', '--output', help='Output Excel file', required=False, default=output_xlsx)
    parser.add_argument('-H','--howto', help='help module', required=False, action='store_true')

    args = parser.parse_args()
    
    # global variables
    source_language = '' 
    input_xlsx = ('input_translate.xlsx')
 
    if args.howto:  # this section might be redundant
        parser.print_help()
        usage()
        return 0
        sys.exit() 

    if args.input:  # defaults to input_translate.xlsx
        input_xlsx = args.input
    if args.output:  # defaults to out_english_.xlsx
        output_xlsx = args.output   
    # input_xlsx = args.input
    translate_excel(input_xlsx, output_xlsx, source_language)


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def check_internet_connection():
    try:
        # Try to make a request to a known website
        response = get("http://www.google.com", timeout=5)
        response.raise_for_status()  # Raise an error for any HTTP error status
        msg_blurb = 'Internet connection is available.'

    except:
        msg_blurb = 'Internet connection is not available.'
        msg_blurb_square(msg_blurb, color_red)
        exit(1)  # Exit with error status 1


def detected_language_enhance(detected_language):
    '''
    Convert 2 digit language code to a full name
    '''
    LANGUAGES = {
        'af': 'Afrikaans',
        'sq': 'Albanian',
        'am': 'Amharic',
        'ar': 'Arabic',
        'hy': 'Armenian',
        'az': 'Azerbaijani',
        'eu': 'Basque',
        'be': 'Belarusian',
        'bn': 'Bengali',
        'bs': 'Bosnian',
        'bg': 'Bulgarian',
        'ca': 'Catalan',
        'ceb': 'Cebuano',
        'ny': 'Chichewa',
        'zh-CN': 'Chinese (Simplified)',
        'zh-TW': 'Chinese (Traditional)',
        'co': 'Corsican',
        'hr': 'Croatian',
        'cs': 'Czech',
        'da': 'Danish',
        'nl': 'Dutch',
        'en': 'English',
        'eo': 'Esperanto',
        'et': 'Estonian',
        'tl': 'Filipino',
        'fi': 'Finnish',
        'fr': 'French',
        'fy': 'Frisian',
        'gl': 'Galician',
        'ka': 'Georgian',
        'de': 'German',
        'el': 'Greek',
        'gu': 'Gujarati',
        'ht': 'Haitian Creole',
        'ha': 'Hausa',
        'haw': 'Hawaiian',
        'iw': 'Hebrew',
        'hi': 'Hindi',
        'hmn': 'Hmong',
        'hu': 'Hungarian',
        'is': 'Icelandic',
        'ig': 'Igbo',
        'id': 'Indonesian',
        'ga': 'Irish',
        'it': 'Italian',
        'ja': 'Japanese',
        'jw': 'Javanese',
        'kn': 'Kannada',
        'kk': 'Kazakh',
        'km': 'Khmer',
        'rw': 'Kinyarwanda',
        'ko': 'Korean',
        'ku': 'Kurdish (Kurmanji)',
        'ky': 'Kyrgyz',
        'lo': 'Lao',
        'la': 'Latin',
        'lv': 'Latvian',
        'lt': 'Lithuanian',
        'lb': 'Luxembourgish',
        'mk': 'Macedonian',
        'mg': 'Malagasy',
        'ms': 'Malay',
        'ml': 'Malayalam',
        'mt': 'Maltese',
        'mi': 'Maori',
        'mr': 'Marathi',
        'mn': 'Mongolian',
        'my': 'Myanmar (Burmese)',
        'ne': 'Nepali',
        'no': 'Norwegian',
        'ps': 'Pashto',
        'fa': 'Persian',
        'pl': 'Polish',
        'pt': 'Portuguese',
        'pa': 'Punjabi',
        'ro': 'Romanian',
        'ru': 'Russian',
        'sm': 'Samoan',
        'gd': 'Scots Gaelic',
        'sr': 'Serbian',
        'st': 'Sesotho',
        'sn': 'Shona',
        'sd': 'Sindhi',
        'si': 'Sinhala',
        'sk': 'Slovak',
        'sl': 'Slovenian',
        'so': 'Somali',
        'es': 'Spanish',
        'su': 'Sundanese',
        'sw': 'Swahili',
        'sv': 'Swedish',
        'tg': 'Tajik',
        'ta': 'Tamil',
        'te': 'Telugu',
        'th': 'Thai',
        'tr': 'Turkish',
        'uk': 'Ukrainian',
        'ur': 'Urdu',
        'ug': 'Uyghur',
        'uz': 'Uzbek',
        'vi': 'Vietnamese',
        'cy': 'Welsh',
        'xh': 'Xhosa',
        'yi': 'Yiddish',
        'yo': 'Yoruba',
        'zu': 'Zulu'
    }

    # Check if the detected_language exists in LANGUAGES dictionary
    if detected_language in LANGUAGES:
        # Return the full name of the language
        return LANGUAGES[detected_language]
    else:
        # If the language is not found, return None
        return None

def language_detect(original_content):
    if original_content not in auto_list:
        try:
            detected_language = detect(original_content)
        except:
            detected_language = 'auto'
    else:
        detected_language = 'auto'

    if original_content not in auto_list and detected_language == 'auto':
        auto_list.append(original_content)

    return detected_language

    
def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')

def translate_excel(input_xlsx, output_xlsx, source_language):
    # Regular expression pattern to match words
    word_pattern = re.compile(r'\b\w+\b')
    word_pattern2 = re.compile(r'\b\w+\b', flags=re.UNICODE)
    
    skip_characters = ['!', '?', ':)']  # Define the list of special characters
    
    target_language = 'en'
    file_exists = os.path.exists(input_xlsx)
    if file_exists == True:
        msg_blurb = (f'Reading {input_xlsx}')
        msg_blurb_square(msg_blurb, color_green)
        workbook = load_workbook(input_xlsx)        
        sheet = workbook.active
    else:
        msg_blurb = (f'Create {input_xlsx} and insert foreign language lines in the first column')
        msg_blurb_square(msg_blurb, color_red)  # Using ANSI escape code for color
        sys.exit()
        
    sheet.cell(row=1, column=2, value='english')
    sheet.cell(row=1, column=3, value='language')
    sheet.cell(row=1, column=4, value='note')    

    for row in sheet.iter_rows(min_row=2):
        (translation, note, e) = ('', '', '')
        (detected_language, text, skipper) = ('', '', '')
        original_content = row[0].value
        
        detected_language = language_detect(original_content) 
 
        # if original_content in skip_characters:
            # skipper = 'skip'
            # print(f'skipping {original_content}')

        # if skipper == 'skip':
            # a = '1'
            # print(f'skipping')

        # elif not text.isalnum():
            # print(f'{original_content} is not a word')
        # if 1==1:
            # print(f'')
            # print(f'temp skipping search')  # temp
        if original_content is not None and original_content != '' and detected_language != 'auto':
        # if original_content is not None:

            
            if isinstance(original_content, (int, float)):
                translation = original_content
            # elif len(str(original_content)) <= 1:
                # print(f'this is too small to translate: {original_content}')
            # Check if the text equals any special characters from the list
            # elif original_content in skip_characters:
                # print(f'skipping {original_content}')
            # elif any(original_content == char for char in skip_characters):
                # print('skipping {original_content}')
            # elif re.search(word_pattern2, original_content):
                # print(f'{original_content} is not unicode') 
            elif re.search(word_pattern, original_content):
                (translation, source_language, note) = translate_request(original_content, target_language, note)
                # (translation, source_language, note) = translate_googletrans(text, source_language, target_language, note)
                detected_language = source_language
                # time.sleep(1) #will sleep for a second
                sleep(1)
                
                if not translation:
                    # print(f"Translation failed for: {original_content}")
                    note = "Translation failed"
                    # (translation, source_language, note) = translate_googletrans(text, source_language, target_language, note)
                    # (translation, source_language, note) = translate_googletrans(text, source_language, target_language, note)

                    detected_language = source_language
                    # time.sleep(2) #will sleep for a second
                    sleep(2)
            print(f'{color_blue}{original_content}  {color_yellow}{translation}  {color_green}{detected_language}  {color_red}{note}{color_reset}')
        
        # if detected_language == 'auto': 
            # detected_language = ''
        # Update the translated content and language columns
        sheet.cell(row=row[0].row, column=2, value=translation)
        sheet.cell(row=row[0].row, column=3, value=detected_language)
        sheet.cell(row=row[0].row, column=4, value=note)
    # print(f'detected_language = {detected_language} auto_list = {auto_list}')   # temp
    
    workbook.save(output_xlsx)

    msg_blurb = (f'Saving to {output_xlsx}')
    msg_blurb_square(msg_blurb, color_green)
    
    print(f'\n\t\t\t{color_green}Translation content saved to {output_xlsx}{color_reset}')
    
def translate_googletrans(text, source_language, target_language, note):
    translator = Translator()
    print(f'source_language = {source_language} target_language, = {target_language,}')   # temp
    source_language = 'auto'
    target_language = "en"
    print(f'source_language = {source_language} target_language, = {target_language,}')   # temp


    '''
    use googletrans module, 60% of time, it works every time
    '''
    translation = ('')
    detected_language = source_language
    original_content = text
    retries = 3 # 3
    for _ in range(retries):  
        try:
            # translation_result = translator.translate(original_content, lang_src=detected_language, lang_tgt=target_language)
            translation_result = translator.translate(original_content, src=detected_language, dest='en')
            
            if translation_result and translation_result.text:
                translation = translation_result.text
                break  # Exit the loop on successful translation

        except Exception as e:
            msg_blurb = (f'Error translating: {e}')
            msg_blurb_square(msg_blurb, color_red)
            # print(f"Error translating: {e}")    # Error translating: 'NoneType' object has no attribute 'group'
            # Retry after a short delay
            sleep(2)

    return (translation, source_language, note)
    
def translate_request(text, target_language, note):
    '''
    use requests to translate lan
    '''
    source_language = 'auto'
    
    url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl={}&tl={}&dt=t&q={}".format(
        source_language, target_language, text
    )

    # Define custom user agent
    user_agent = "Mozilla/5.0"

    # Define SSL certificate verification (set to False if you don't want to verify)
    verify_ssl_cert = True  # Change to False if you don't want to verify SSL certificates

    # Define headers with user agent
    headers = {
        "User-Agent": user_agent
    }


    # Send GET request with custom user agent, proxies, and SSL certificate verification
    # response = requests.get(url, headers=headers, proxies=proxies, verify=verify_ssl_cert)



    try:
        # response = requests.get(url, verify=True)   # works
        # response = requests.get(url, headers=headers, verify=True)
        response = get(url, headers=headers, verify=True)
        
        if response.status_code == 200:
            data = response.json()
            translation = data[0][0][0] if data else ""

            source_language = data[2]
            note = ''
        else:
            print("Failed to translate. Status code:", response.status_code)
            note = ("Failed to translate. Status code:", response.status_code)
            source_language = ''
            
    except Exception as e:
        print(f'Error occurred while translating: {color_red}{e}{color_reset}')
        (translation, source_language) = ('', '')
        note = 'Error occurred while translating'
        source_language = ''
    detected_language = detected_language_enhance(source_language)
    source_language = detected_language
    return (translation, source_language, note)
    
def usage():
    file = sys.argv[0].split('\\')[-1]

    print(f'\nDescription: {color_green}{description}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f'\n    {color_yellow}insert your info into input_translate.xlsx')
    print(f'\nExample:')
    # print(f'    {file} -c')
    # print(f'    {file} -f')
    # print(f'    {file} -g')
    # print(f'    {file} -m')
    print(f'    {file}')
    print(f'    {file} -I input_translate.xlsx')


if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

1.0.0 - use requests with googletrans as a backup (non-working) module.
0.2.0 - removed any switch requirements to make the exe version easier
0.1.0 - read xlsx, translate, export to xlsx
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
fix language module ar = arabic
Make sure to handle potential issues like rate limiting, certificate verification, and unexpected input data gracefully. 

if len(original_content) >= 1:
TypeError: object of type 'NoneType' has no len()

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

GoogleTrans is either rate limiting or they are using an API now (tested fine on 10/17/2023


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Copyright (C) 2024 LincolnLandForensics
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License version 2, as published by the
# Free Software Foundation
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details (http://www.gnu.org/licenses/gpl.txt).

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>