#!/usr/bin/python
# coding: utf-8

'''
This script is used to translate the contents of an Excel spreadsheet from many 
languages to English. It uses the xlrd library to read the input file, googletrans 
to perform the translation, and xlsxwriter to write the translated contents to a new Excel file.
'''


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import sys
import xlrd # read xlsx
import os.path
import argparse  # for menu system
import xlsxwriter
import googletrans		# pip install googletrans
from googletrans import Translator


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "Read input.xlsx filled with another language and translate it to english"
version = '0.2.0'


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    # global variables
    global Row
    # Row = 0  # defines arguments
    Row = 1  # defines arguments   # if you want to add headers 
    global detectedLanguage
    detectedLanguage = '' 
    global filename
    filename = ('input.xlsx')
    global spreadsheet
    spreadsheet = ('out_english_.xlsx')
    global sheet_format
    sheet_format = ''
    sheet_format = "Translation"

    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-a', '--arabic', help='arabic 2 english', required=False, action='store_true')
    parser.add_argument('-c', '--chinese', help='chinese 2 english', required=False, action='store_true')
    parser.add_argument('-f', '--french', help='french 2 english', required=False, action='store_true')
    parser.add_argument('-g', '--german', help='german 2 english', required=False, action='store_true')

    parser.add_argument('-m', '--multi', help='multi language 2 english when you dont know', required=False, action='store_true')
    parser.add_argument('-s', '--spanish', help='spanish 2 english', required=False, action='store_true')


    args = parser.parse_args()

    if args.input:  # defaults to index.xlsx
        filename = args.input
    if args.output: # defaults to out_english_.xlsx
        spreadsheet = args.output

    create_xlsx()

    if args.arabic:
        print('Translating Arabic from %s' %(filename))
        detectedLanguage = 'ar'  # arabic?
        read_language()
    elif args.chinese:
        print('Translating Chinese from %s' %(filename))
        detectedLanguage = 'zh-CN'  # chinese (simplified)
        read_language()
    elif args.french:
        print('Translating French from %s' %(filename))
        detectedLanguage = 'fr'  # french
        read_language()
    elif args.german:
        print('Translating German from %s' %(filename))
        detectedLanguage = 'de'  # german
        read_language()
    elif args.spanish: 
        print('Translating Spanish from %s' %(filename))
        detectedLanguage = 'es'  # spanish
        read_language()
    elif args.multi:
        detectedLanguage = ''  # unknown language?
        read_language()
    else:
        print('Translating from %s' %(filename))
        detectedLanguage = ''  # unknown language?
        read_language()        
        
    workbook.close()
    return 0


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def create_xlsx():  
    global workbook
    workbook = xlsxwriter.Workbook(spreadsheet)
    global Sheet1
    Sheet1 = workbook.add_worksheet('Sheet1')
    header_format = workbook.add_format({'bold': True, 'border': 1})
    Sheet1.freeze_panes(1, 1)  # Freeze cells
    Sheet1.set_selection('B2')

    # Excel column width
    Sheet1.set_column(0, 0, 75)  # original
    Sheet1.set_column(1, 1, 75)  # english
    Sheet1.set_column(2, 2, 25)  # detectedLanguage

    # Write column headers
    Sheet1.write(0, 0, 'original', header_format)
    Sheet1.write(0, 1, 'english', header_format)
    Sheet1.write(0, 2, 'language', header_format)


def format_function(bg_color='white'):
    global format
    format = workbook.add_format({
        'bg_color': bg_color
    })

def read_language():
    file_exists = os.path.exists(filename)
    if file_exists == True:
        workbook = xlrd.open_workbook(filename)
    else:
        print('%s does not exist' %(filename))
        exit()

    # Open the worksheet
    worksheet = workbook.sheet_by_index(0)

    # determine how many columns wide this xlsx is
    columns = worksheet.ncols

    # determine how many rows deep this xlsx is 
    rows = worksheet.nrows

    # Iterate the rows and columns

    for r in range(0, rows):   # rows 18
        for c in range(0, columns):   # columns 5
            (original, english) = ('', '')
            original = worksheet.cell_value(r, 0)

        if len(original) >= 1:
            translator = Translator()
            # if detectedLanguage == '':
                # detectedLanguage = translator.detect(original[0])  # fix me
                # print("DetectedLanguage = %s" %(detectedLanguage))
         
            try:
                english = translator.translate(original, lang_src=detectedLanguage, lang_tgt='en').text 
            except TypeError as error:
                print(error)

            print('%s\t=\t%s' %(original, english))

        else:
            print('This line >%s< is too short' %(original))  # temp

        # Write excel
        write_xlsx(original, english, detectedLanguage)

def write_xlsx(original, english, detectedLanguage):
    global Row
    Sheet1.write_string(Row, 0, original)
    Sheet1.write_string(Row, 1, english)
    Sheet1.write_string(Row, 2, detectedLanguage)

    Row += 1

def usage():
    file = sys.argv[0].split('\\')[-1]
    print("\nDescription: " + description)
    print(file + " Version: %s by %s" % (version, author))
    print("\nExample:")

    print("\t" + file + " -a -I input.xlsx -O out_english_.xlsx\t\t")
    print("\t" + file + " -c\t\t")
    print("\t" + file + " -f\t\t")
    print("\t" + file + " -g\t\t")
    print("\t" + file + " -m\t\t")
    print("\t" + file + " -s\t\t")

  
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

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
if you know it's one of the specified languages, select that, otherwise select -m or no switches for unknown or mixed languages


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Copyright (C) 2022 LincolnLandForensics
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
