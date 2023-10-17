#!/usr/bin/python
# coding: utf-8

'''
This script is used to translate the contents of an Excel spreadsheet from many 
languages to English. It uses the openypyl library to read the input file, googletrans 
to perform the translation, and openypyl to write the translated contents to a new Excel file.
'''


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import sys
import time
import os.path
import openpyxl
import argparse  # for menu system
import googletrans		# pip install googletrans   # redunant?
from googletrans import Translator

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "Read input.xlsx filled with another language and translate it to english"
version = '0.4.7'

LANGUAGES = {
    'arabic': 'ar',
    'chinese': 'zh-CN',
    'french': 'fr',
    'german': 'de',
    'spanish': 'es',
    'multi': ''
}

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
    # print(f'major version = {major_version} Build= {build_version} {version_info}')   # temp

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
    parser = argparse.ArgumentParser(description='Translate Excel contents from various languages to English')
    parser.add_argument('-I', '--input', help='Input Excel file', required=False)
    parser.add_argument('-O', '--output', help='Output Excel file', required=False, default='out_english.xlsx')
    parser.add_argument('-l', '--language', help='Source language for translation', choices=LANGUAGES.keys(), required=False)
    args = parser.parse_args()
    
    # global variables
    # global Row
    # Row = 0  # defines arguments
    # Row = 1  # defines arguments   # if you want to add headers 
    # global source_language
    source_language = '' 
    # global input_xlsx
    input_xlsx = ('input_translate.xlsx')
    # global output_xlsx
    output_xlsx = ('out_english_.xlsx')
    # global sheet_format
    # sheet_format = ''
    # sheet_format = "Translation"    

    if args.input:  # defaults to input_translate.xlsx
        input_xlsx = args.input
    if args.output:  # defaults to out_english_.xlsx
        output_xlsx = args.output   
    if args.language:  # defaults to input_translate.xlsx
        source_language = args.language

    translate_excel(input_xlsx, output_xlsx, source_language)


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def translate_excel(input_xlsx, output_xlsx, source_language):
    translator = Translator()
    print(f'{color_yellow}You selected {source_language} as your language{color_reset}')
    file_exists = os.path.exists(input_xlsx)
    if file_exists == True:
        print(f'{color_green}Reading {input_xlsx} {color_reset}')
        workbook = openpyxl.load_workbook(input_xlsx)
        sheet = workbook.active
    else:
        print(f'{color_red}{input_xlsx} does not exist{color_reset}')
        exit()


    # Add columns for translated content and language. Maintains any extra input from the source
    sheet.insert_cols(2)
    sheet.insert_cols(3)
    sheet.cell(row=1, column=2, value='english')
    sheet.cell(row=1, column=3, value='language')
    sheet.cell(row=1, column=4, value='note')    

    for row in sheet.iter_rows(min_row=2):
        (translation, note, e) = ('', '', '')
        original_content = row[0].value

        try:
            detected_language = LANGUAGES.get(source_language, '')
        except:
            detected_language = source_language # statically assign till detection works
    
        if original_content is not None:
            if isinstance(original_content, (int, float)):
                # print("Original content is a number:", original_content)
                translation = original_content
            elif len(str(original_content)) >= 1:
                retries = 1 # 3
                for _ in range(retries):
                    try:
                        translation_result = translator.translate(original_content, lang_src=detected_language, lang_tgt='en')
                        if translation_result and translation_result.text:
                            translation = translation_result.text
                            break  # Exit the loop on successful translation
                    except Exception as e:
                        print(f"Error translating: {e}")
                        # Retry after a short delay
                        time.sleep(2)

                if not translation:
                    # print(f"Translation failed for: {original_content}")
                    note = "Translation failed"

        print(f'{color_blue}{original_content}  {color_yellow}{translation}  {color_green}{detected_language}  {color_red}{note}{color_reset}')

        time.sleep(5) #will sleep for 5 seconds

        # Update the translated content and language columns
        sheet.cell(row=row[0].row, column=2, value=translation)
        sheet.cell(row=row[0].row, column=3, value=detected_language)
        sheet.cell(row=row[0].row, column=4, value=note)
        
    workbook.save(output_xlsx)
    print(f'{color_green}Translated content saved to {output_xlsx}{color_reset}')
    
def usage():
    file = sys.argv[0].split('\\')[-1]

    print(f'\nDescription: {color_green}{description}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f'\n    {color_yellow}insert your info into input_translate.xlsx')
    print(f'\nExample:')
    print(f'    {file} -c')
    print(f'    {file} -f')
    print(f'    {file} -g')
    print(f'    {file} -m')
    print(f'    {file} -s')


if __name__ == '__main__':
    main()


# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

0.2.0 - removed any switch requirements to make the exe version easier
0.1.0 - read xlsx, translate, export to xlsx
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""

detectedLanguage = translator.detect(original[0])  # fix me

[SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: unable to get local issuer certificate (_ssl.c:1007)

if len(original_content) >= 1:
TypeError: object of type 'NoneType' has no len()

if input is just a number or is blank, skip it

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
if you know it's one of the specified languages, select that, otherwise select -m or no switches for unknown or mixed languages

GoogleTrans is either rate limiting or they are using an API now


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Copyright (C) 2023 LincolnLandForensics
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