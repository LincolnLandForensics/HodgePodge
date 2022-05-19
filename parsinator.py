#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import re
import sys
import hashlib
import datetime
import argparse  # for menu system
import xlsxwriter
from subprocess import call

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "parse dirlist"
tech = 'LincolnLandForensics'  # change this to your name
version = '1.0.2'
# output_file = "report_backup.txt"   # temp

# Regex section
regex_md5 = re.compile(r'^([a-fA-F\d]{32})$')  # regex_md5        [a-f0-9]{32}$/gm
regex_pdf = re.compile(r'^.(pdf|xps|oxps)')  # (pdf|xps|oxps)

# Color options
if sys.platform == 'win32' or sys.platform == 'win64':
    # if windows, don't use colors
    (r, o, y, g, b) = ('', '', '', '', '')
else:
    r = '\033[31m'  # red
    o = '\033[0m'  # off
    y = '\033[33m'  # yellow
    g = '\033[32m'  # green
    b = '\033[34m'  # blue


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    global Row
    Row = 1  # defines arguments
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    # parser.add_argument('-l', '--logparse', help='tableau log parser', required=False, action='store_true')
    # parser.add_argument('-r', '--report', help='write report', required=False, action='store_true')
    parser.add_argument('-d', '--dirlist', help='dirlist parser', required=False, action='store_true')
    # parser.add_argument('-s', '--sticker', help='write sticker', required=False, action='store_true')

    args = parser.parse_args()

    if not args.input:  # this section might be redundant
        parser.print_help() 
        usage()
        return 0
    # Choose Sheet format
    global sheet_format
    sheet_format = ''

    # if args.phone:
        # sheet_format = "phone"
        # print('this is a phone report') #temp

    if args.input and args.output:
        global filename
        filename = args.input
        global spreadsheet
        spreadsheet = args.output
        create_xlsx()

        if args.dirlist:
            read_text()
        # elif args.logparse:
            # parse_log()
        # elif args.sticker:
            # write_sticker()
        

    # set linux ownership    
    if sys.platform == 'win32' or sys.platform == 'win64':
        pass
    else:
        call(["chown %s.%s *.xlsx" % (tech.lower(), tech.lower())], shell=True)

    workbook.close()
    return 0


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def create_xlsx():  # BCI output (Default)
    global workbook
    workbook = xlsxwriter.Workbook(spreadsheet)
    global Sheet1
    Sheet1 = workbook.add_worksheet('Forensics')
    header_format = workbook.add_format({'bold': True, 'border': 1})
    Sheet1.freeze_panes(1, 1)  # Freeze cells
    Sheet1.set_selection('B2')

    # Excel column width
 
    Sheet1.set_column(0, 0, 150)  
    Sheet1.set_column(1, 1, 20)  # 
    Sheet1.set_column(2, 2, 10)  # 
    Sheet1.set_column(3, 3, 26)  # 
    Sheet1.set_column(4, 4, 8)  # 
    
    # Write column headers
    Sheet1.write(0, 0, 'filePath', header_format)
    Sheet1.write(0, 1, 'fileType', header_format)
    Sheet1.write(0, 2, 'extension', header_format)
    Sheet1.write(0, 3, 'note', header_format)
    Sheet1.write(0, 4, 'priority', header_format)

def format_function(bg_color='white'):
    global format
    format = workbook.add_format({
        'bg_color': bg_color
    })

def parse_log():
    style = workbook.add_format()
    (header, report, date) = ('', '', '<insert date here>')
    csv_file = open(filename)
    outputFile = "logreport.txt"
    output = open(outputFile, 'w+')
    (caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType) = ('', '', '', '', '', '', '')
    (caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived) = ('', '', '', '', '', '')
    (serial, makeModel, storageLocation, removalDate, exportedEvidence, status) = ('', '', '', '', '', '')
    (analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff) = ('', '', '', '', '', '')
    (notes, attachment, exportComplete, model, hddserial, capacity) = ('', '', '', '', '', '')

    exhibit = str(input("exhibit : ")).strip()
    # read section
    for each_line in csv_file:
    # for each_line in text.splitlines():
        # if each_line[1]:

        if "Task:" in each_line:
            imagingType = re.split("Task: ", each_line, 0)
            imagingType = str(imagingType[1]).strip().lower()
        elif "Status: Ok" in each_line:
            status = 'Imaged'
        elif "Status: Error/Failed" in each_line:
            status = 'Not imaged'
        
        elif "Started:" in each_line:
            imagingStarted = re.split("Started: ", each_line, 0)
            imagingStarted = str(imagingStarted[1]).strip()
            
        elif "Closed:" in each_line:
            imagingFinished = re.split("Closed: ", each_line, 0)
            imagingFinished = str(imagingFinished[1]).strip()
        elif "User: " in each_line:
            forensicExaminer = re.split("User: ", each_line, 0)
            print("forensicExaminer=", forensicExaminer[1].strip())      
            forensicExaminer = str(forensicExaminer[1]).strip()

        elif "Case ID:" in each_line:
            caseNumber = re.split("Case ID:", each_line, 0)
            caseNumber = str(caseNumber[1]).strip()
            caseNumber = caseNumber.replace("<<not entered>>", "")

        elif "Case Notes:" in each_line:
            notes = re.split("Case Notes:", each_line, 0)
            notes = str(notes[1]).strip()
            notes = notes.replace("<<not entered>>", "")
        elif "Imager App: " in each_line:
            imagingTool1 = re.split("Imager App: ", each_line, 0)
            imagingTool1 = str(imagingTool1[1]).strip()

        elif "Imager Ver: " in each_line:
            imagingTool2 = re.split("Imager Ver: ", each_line, 0)
            imagingTool2 = str(imagingTool2[1]).strip()

        elif "Model: " in each_line and len(model) == 0:
            model = re.split("Model: ", each_line, 0)
            model = str(model[1]).strip()
        elif "S/N: " in each_line:
            hddserial = re.split("S/N: ", each_line, 0)
            hddserial = str(hddserial[1]).strip()

        elif "Capacity in bytes reported Pwr-ON: " in each_line:
            capacity = re.split("Capacity in bytes reported Pwr-ON: ", each_line, 0)
            capacity = str(capacity[1]).strip()
            if "(" in capacity:
                capacity = re.split("\(", each_line, 0)
                capacity = str(capacity[1]).strip()
                capacity = capacity.replace(")", "").replace(".0", "")

        elif "Filename of first chunk: " in each_line:
            exportLocation = re.split("Filename of first chunk: ", each_line, 0)
            exportLocation = str(exportLocation[1]).strip()
        elif "Disk MD5:  " in each_line:
            imageMD5 = re.split("Disk MD5:  ", each_line, 0)
            imageMD5 = str(imageMD5[1]).strip()

    imagingTool = ('%s %s' %(imagingTool1.strip(), imagingTool2.strip()))

    notes = ("This had a %s hard drive, model %s, serial # %s. %s" %(capacity, model, hddserial, notes))

    if status == 'Not imaged':
        notes = ("%s This drive could not be imaged." %(notes))

    print(notes)
    print("status = %s" %(status))

    # Write excel
    # insert a color red if status = not imaged
    write_report(caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType,
                caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived,
                serial, makeModel, storageLocation, removalDate, exportedEvidence, status,
                analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff,
                notes,attachment, exportComplete)

    # output.write(footer+'\n')


def read_text():
    '''
    dir /s /b >dirlist.txt
    
    '''
    # global Row    #The magic to pass Row globally
    style = workbook.add_format()
    (header, report, date) = ('', '', '<insert date here>')
    csv_file = open(filename)
    outputFile = "report.txt"
    output = open(outputFile, 'w+')
    (subject, vowel) = ('test', 'aeiou')


    for each_line in csv_file:
        (filePath, fileType, extension, note, priority) = ('', '', '', '', '')
        filePath = each_line.strip()

        (color) = ('white')
        style.set_bg_color('white')  # test
        # each_line = each_line +  "\t" * 27
        # each_line = each_line.split('\t')  # splits by tabs

        # value = each_line
        # note = value

        try:
            extension = each_line.split(".")[-1]
            # print('blah')   # temp
            extension = extension.lower()
        except:pass

        
        if filePath == extension:
            extension = ''
        if len(extension) > 6:
            extension = ''


        if each_line.endswith('.pdf'):
            (extension, fileType) = ('pdf', 'pdf')
        elif each_line.endswith('.xps'):
            (extension, fileType) = ('xps', 'pdf-like document')
        elif each_line.endswith('.oxps'):
            (extension, fileType) = ('oxps', 'pdf-like document')
            
        
        # if re.match(regex_pdf, each_line):	#regex pdf
            # (fileType) = ('pdf-like document')
            # print('blah2')  # temp

        # fileType
        if re.search('pdf|xps|oxps', extension):    
            (fileType) = ('pdf-like document')
            if extension == 'pdf':
                (fileType) = ('pdf')
        elif re.search('xls|xlsx|csv|tsv|xlt|xlm|xlsm|xltx|xltm|xlsb|xla|xlam|xll|xlw|ods|fodp|qpw', extension):    
            (fileType) = ('excel-like document')
        elif re.search('accdb|gub|mdb|dbf|myd|myi|frm|dbf|dat|db|dbf', extension):    
            (fileType) = ('database')
            priority = '3'            
        elif re.search('sql|bak|archive', extension):    
            (fileType) = ('database backup')
            priority = '2'
        elif re.search('7z|xz|bzip2|gzip|tar|zip|wim|ar|arj|cab|chm|cpio|cramfs|dmg|ext|fat|gpt|hfs|ihex|iso|lzh|lzma|mbr|msi|nsis|ntfs|qcow2|rar|rpm|squashfs|udf|uefi|vdi|vhd|vmdk|wim|xar', extension):    
            (fileType) = ('compressed')


            # if extension = 'pdf':
                # (fileType) = ('pdf')
        elif re.search('ai|bmp|bpg|cdr|cpc|eps|exr|flif|gif|heif|ilbm|ima|jp2|j2k|jpf|jpm|jpg|jpeg|jpg2|j2c|jpc|jpx|mj2|jpeg|jpg|jxl|kra|ora|pcx|pgf|pgm|png|pnm|ppm|psb|psd|psp|svg|tga|tiff|webp|xaml|xcf', extension):    
            (fileType) = ('picture files')
        elif re.search('3g2|3gp|amv|asf|avi|drc|flv|f4v|f4p|f4a|f4b|gif|gifv|m4v|mkv|mov|qt|mp4|m4p|mpg|mpeg|m2v|mp2|mpe|mpv|mts|m2ts|ts|mxf|nsv|ogv|ogg|rm|rmvb|roq|svi|viv|vob|webm|wmv|yuv', extension):    
            (fileType) = ('video files')
        elif re.search('doc|docx|docm|dotx|dotm|docb|dot|wbk|odt|fodt|rtf|wp*|tmd', extension):    
            (fileType) = ('word-like document')
        elif re.search('emlx|msg', extension):    
            (fileType) = ('email-like document')

        if re.search('qbw|qba|qbb|qbx', extension):
                priority = '1'
                (fileType) = ('quickbooks')
        elif re.search('qdf|qdb', extension):
                priority = '1'
                (fileType) = ('quicken')

# monthlysales|salestax|revenue.state.il.us|dailyreport|sales_tax_returns


        # priority 
        if 'password' in filePath.lower():  
            note = ('password %s' %(note))
            if re.search('doc|xls|txt', extension):  
                priority = '1'
        elif 'quickbook' in filePath.lower() or 'qbook' in filePath.lower():  
            note = ('quickbook %s' %(note))
            # priority = '2'
        elif re.search('budget|sales|quickbook', filePath.lower()):
            note = ('%s research' %(note))
            if re.search('pdf|xls|csv', extension):  
                priority = '2'
            else:
                priority = '5'
        elif re.search('monthlysales|salestax|revenue.state.il.us|dailyreport|sales_tax_returns', filePath.lower()):
            if len(priority) == 0:
                priority = '3'

        if re.search('google drive|dropbox', filePath.lower()):    
            note = ('cloudStorage %s' %(note))
            if len(priority) == 0:
                priority = '3'            
        if 'mobilesync\\backup' in filePath.lower() and 'Manifest.db' in filePath: 
            note = ('icloudbackup feed this to c-brite %s' %(note))
            priority = '2'

        # print(report)
        output.write(report)

        # Write excel

        write_report(filePath, fileType, extension, note, priority)

    # output.write(footer+'\n')

def write_report(filePath, fileType, extension, note, priority):

    global Row

    Sheet1.write_string(Row, 0, filePath)
    Sheet1.write_string(Row, 1, fileType)
    Sheet1.write_string(Row, 2, extension)
    Sheet1.write_string(Row, 3, note)
    Sheet1.write_string(Row, 4, priority)

    Row += 1

def write_sticker():
    # global Row    #The magic to pass Row globally
    style = workbook.add_format()
    (header, report, date) = ('', '', '<insert date here>')
    csv_file = open(filename)
    outputFile = "report.txt"
    output = open(outputFile, 'w+')


    footer = '''  

    The images of all the devices will be retained. The case agent may request additional analysis or files to be exported if new evidence of probative value is determined, at a future date.
    
Evidence:
    Reports and supporting files were exported and given to the case agent.
    '''

    for each_line in csv_file:
        (caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType) = ('', '', '', '', '', '', '')
        (caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived) = ('', '', '', '', '', '')
        (serial, makeModel, storageLocation, removalDate, exportedEvidence, status) = ('', '', '', '', '', '')
        (analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff) = ('', '', '', '', '', '')
        (notes, attachment, exportComplete) = ('', '', '')


        if phoneNumber != '':
            print('_____________%s is from a phone extract') %(phoneNumber) #temp

        (color) = ('white')
        style.set_bg_color('white')  # test
        each_line = each_line +  "\t" * 27
        each_line = each_line.split('\t')  # splits by tabs

        value = each_line
        note = value

        if each_line[1]:                        ### checks to see if there is an each_line[1] before preceeding
            Value = note
            caseNumber = each_line[0]
            caseName = each_line[1]
            subjectBusinessName = each_line[2]
            caseType = each_line[3]
            caseAgent = each_line[4]
            forensicExaminer = each_line[5]
            exhibit = each_line[6]
            makeModel = each_line[7]
            serial = each_line[8]
            phoneNumber = each_line[9]
            imagingStarted = each_line[10]
            imagingFinished = each_line[11]
            imagingTool = each_line[12]
            imagingType = each_line[13]
            storageLocation = each_line[14]
            dateReceived = each_line[15]
            removalDate = each_line[16]
            exportedEvidence = each_line[17]
            status = each_line[18]
            analysisTool = each_line[19]
            exportLocation = each_line[20]
            imageMD5 = each_line[21]
            locationOfCaseFile = each_line[22]
            reasonForRemoval = each_line[23]
            removalStaff = each_line[24]
            notes = each_line[25]            
            attachement = each_line[26]            
            exportComplete = each_line[27]
            
        header = ('''Case#: %s  Ex: %s
CaseName: %s
Subject: %s
Make: %s 
Serial: %s
Agent: %s
%s
''') %(caseNumber, exhibit, caseName, subjectBusinessName, makeModel, serial, caseAgent, status)

        output.write(header+'\n')


def usage():
    file = sys.argv[0].split('\\')[-1]
    print("\nDescription: " + description)
    print(file + " Version: %s by %s" % (version, author))
    print("\nExample:")
    print("\t" + file + " -d -I dirlist.txt -O output_dirlist_.xlsx\t\t")
    # print("\t" + file + " -l -I input.log -O out_log_.xlsx\t\t")
    # print("\t" + file + " -s -I input.txt -O out_log_.xlsx\t\t")
    
if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
1.0.1 - Created a Tableau log parser
1.0.0 - Created forensic report writer
0.1.2 - converted tabs to 4 spaces for #pep8
0.0.2 - python2to3 conversion
0.0.1 - based on Password_recheckinator.py

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
Create an FTK parser.



"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
 

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
