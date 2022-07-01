#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import re
import sys
import hashlib
import datetime
import argparse  # for menu system
import xlsxwriter
from datetime import date
from subprocess import call

d = date.today()
Day    = d.strftime("%d")
Month = d.strftime("%B")
Year  = d.strftime("%Y")        
todaysDate = d.strftime("%m/%d/%Y")



# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "convert cls to written forensic report, print stickers, and convert imaging logs to xlsx"
tech = 'LincolnLandForensics'  # change this to your name
version = '1.4.1'

# Regex section
regex_md5 = re.compile(r'^([a-fA-F\d]{32})$')  # regex_md5        [a-f0-9]{32}$/gm

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
    parser.add_argument('-r', '--report', help='write report', required=False, action='store_true')
    parser.add_argument('-P', '--phone', help='phone output', required=False, action='store_true')
    parser.add_argument('-s', '--sticker', help='write sticker', required=False, action='store_true')
    parser.add_argument('-l', '--logparse', help='tableau or FTK log parser', required=False, action='store_true')

    args = parser.parse_args()

    if not args.input:  # this section might be redundant
        parser.print_help() 
        usage()
        return 0
    # Choose Sheet format
    global sheet_format
    sheet_format = ''

    if args.phone:
        sheet_format = "phone"
        print('this is a phone report') #temp

    if args.input and args.output:
        global filename
        filename = args.input
        global spreadsheet
        spreadsheet = args.output
        create_xlsx()

        if args.report:
            read_text()
        elif args.logparse:
            parse_log()
        elif args.sticker:
            write_sticker()
        

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
 
    Sheet1.set_column(0, 0, 15)  
    Sheet1.set_column(1, 1, 16)  # 
    Sheet1.set_column(2, 2, 20)  # 
    Sheet1.set_column(3, 3, 16)  # 
    Sheet1.set_column(4, 4, 25)  # 
    Sheet1.set_column(5, 5, 15)  # 
    Sheet1.set_column(6, 6, 7)  # 
    Sheet1.set_column(7, 7, 30)  # 
    Sheet1.set_column(8, 8, 17)  # 
    Sheet1.set_column(9, 9, 16)  # 
    Sheet1.set_column(10, 10, 16)  # 
    Sheet1.set_column(11, 11, 16)  # 
    Sheet1.set_column(12, 12, 30)  # 
    Sheet1.set_column(13, 13, 32)  # 
    Sheet1.set_column(14, 14, 16)  # 
    Sheet1.set_column(15, 15, 16)  # 
    Sheet1.set_column(16, 16, 16)  # 
    Sheet1.set_column(17, 17, 8)  # 
    Sheet1.set_column(18, 18, 26)  # 
    Sheet1.set_column(19, 19, 15)  # 
    Sheet1.set_column(20, 20, 23)  # 
    Sheet1.set_column(21, 21, 18)  # 
    Sheet1.set_column(22, 22, 20)  # 
    Sheet1.set_column(23, 23, 19)  # 
    Sheet1.set_column(24, 24, 15)  # 
    Sheet1.set_column(25, 25, 25)  # 
    Sheet1.set_column(26, 26, 25)  # 
    Sheet1.set_column(27, 27, 25)  #
    Sheet1.set_column(28, 28, 15)  #
    Sheet1.set_column(29, 29, 20)  #
    Sheet1.set_column(30, 30, 25)  #
    Sheet1.set_column(31, 31, 25)  #
    Sheet1.set_column(32, 32, 25)  #
    Sheet1.set_column(33, 33, 15)  #
    Sheet1.set_column(34, 34, 20)  #


    
    # Write column headers

    Sheet1.write(0, 0, 'caseNumber', header_format)
    Sheet1.write(0, 1, 'caseName', header_format)
    Sheet1.write(0, 2, 'subjectBusinessName', header_format)
    Sheet1.write(0, 3, 'caseType', header_format)
    Sheet1.write(0, 4, 'caseAgent', header_format)
    Sheet1.write(0, 5, 'forensicExaminer', header_format)
    Sheet1.write(0, 6, 'exhibit', header_format)
    Sheet1.write(0, 7, 'make/Model', header_format)
    Sheet1.write(0, 8, 'serial#', header_format)
    Sheet1.write(0, 9, 'phoneNumber', header_format)
    Sheet1.write(0, 10, 'imagingStarted', header_format)
    Sheet1.write(0, 11, 'imagingFinished', header_format)
    Sheet1.write(0, 12, 'imagingTool', header_format)
    Sheet1.write(0, 13, 'imagingType', header_format)
    Sheet1.write(0, 14, 'storageLocation', header_format)
    Sheet1.write(0, 15, 'dateReceived', header_format)
    Sheet1.write(0, 16, 'removalDate', header_format)
    Sheet1.write(0, 17, 'exportedEvidence', header_format)
    Sheet1.write(0, 18, 'status', header_format)
    Sheet1.write(0, 19, 'analysisTool', header_format)
    Sheet1.write(0, 20, 'exportLocation', header_format)
    Sheet1.write(0, 21, 'imageMD5', header_format)
    Sheet1.write(0, 22, 'locationOfCaseFile', header_format)
    Sheet1.write(0, 23, 'reasonForRemoval', header_format)
    Sheet1.write(0, 24, 'removalStaff', header_format)
    Sheet1.write(0, 25, 'notes', header_format)
    Sheet1.write(0, 26, 'attachment', header_format)
    Sheet1.write(0, 27, 'tempNotes', header_format)
    Sheet1.write(0, 28, 'inventoryDate', header_format)
    Sheet1.write(0, 29, 'operation', header_format)
    Sheet1.write(0, 30, 'Action', header_format)
    Sheet1.write(0, 31, 'imageSHA256', header_format)
    Sheet1.write(0, 32, 'OS', header_format)
    Sheet1.write(0, 33, 'dateSeized', header_format)    
    Sheet1.write(0, 34, 'summary', header_format)    
    
    
    

def format_function(bg_color='white'):
    global format
    format = workbook.add_format({
        'bg_color': bg_color
    })

def fix_date(date):
    date = date.strip()
    date = date.replace("  ", " ")  # test
    date = date.split(' ')     # Fri Jun 04 07:55:41 2021
    mo = date[1]    # convert month to a number
    mo = mo.replace("Jan", "1").replace("Feb", "2").replace("Mar", "3").replace("Apr", "4")
    mo = mo.replace("May", "5").replace("Jun", "6").replace("Jul", "7").replace("Aug", "8")
    mo = mo.replace("Sep", "9").replace("Oct", "10").replace("Nov", "11").replace("Dec", "12")
    dy = date[2].lstrip('0')
    tm = date[3].lstrip('0')
    yr = date[4]
    date = ('%s/%s/%s %s' %(mo, dy, yr, tm))  # 3/4/2021 9:17
    return date
    
def parse_log():
    style = workbook.add_format()
    (header, report, date) = ('', '', '<insert date here>')
    # csv_file = open(filename)
    csv_file = open(filename, encoding='utf8')
    outputFile = "logreport.txt"
    output = open(outputFile, 'w+')
    (caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType) = ('', '', '', '', '', '', '')
    (caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived) = ('', '', '', '', '', '')
    (serial, makeModel, storageLocation, removalDate, exportedEvidence, status) = ('', '', '', '', '', '')
    (analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff) = ('', '', '', '', '', '')
    (notes, attachment, tempNotes, model, hddserial, capacity) = ('', '', '', '', '', '')
    (size, imagingTool1, imagingTool2) = ('', '', '')
    (inventoryDate, operation, Action, imageSHA256, OS, dateSeized) = ('', '', '', '', '', '')
    (hostname, timezone, os, ip, encryption, summary) = ('', '', '', '', '', '')
    
    exhibit = str(input("exhibit : ")).strip()
    # read section
    for each_line in csv_file:
    # for each_line in text.splitlines():
        # if each_line[1]:

        if "Task:" in each_line:
            imagingType = re.split("Task: ", each_line, 0)
            imagingType = str(imagingType[1]).strip().lower()

        elif " Extraction type " in each_line: #cellebrite xls
            imagingType = re.split(" Extraction type ", each_line, 0)
            imagingType = str(imagingType[1]).strip().lower()
            print(imagingType)  #temp
        elif "Source Type: Physical" in each_line:
            imagingType = "disk to file"
            
        elif "Status: Ok" in each_line:
            status = 'Imaged'
        elif "Status: Error/Failed" in each_line:
            status = 'Not imaged'

        elif "Evidence Number: " in each_line:      #FTK_parse
            exhibit = re.split("Evidence Number: ", each_line, 0)
            exhibit = str(exhibit[1]).strip()
        elif "Exhibit#" in each_line:      #cellebrite
            exhibit = re.split("Exhibit#", each_line, 0)
            exhibit = str(exhibit[1]).strip()
        
        elif "Started:" in each_line:
            imagingStarted = re.split("Started: ", each_line, 0)
            imagingStarted = str(imagingStarted[1]).strip()
            imagingStarted = fix_date(imagingStarted)
        elif "Acquisition started:" in each_line:
            imagingStarted = re.split("Acquisition started: ", each_line, 0)
            imagingStarted = str(imagingStarted[1]).strip()
            imagingStarted = fix_date(imagingStarted)


        elif "Extraction start date/time" in each_line: #cellebrite
            imagingStarted = re.split("time", each_line, 0)
            imagingStarted = str(imagingStarted[1]).strip().replace(" -05:00", "").strip(':').strip().replace("(GMT-5)", "")
            # imagingStarted = fix_date(imagingStarted)
            print(imagingStarted)   #temp

            
        elif "Closed:" in each_line:
            imagingFinished = re.split("Closed: ", each_line, 0)
            imagingFinished = str(imagingFinished[1]).strip()
            imagingFinished = fix_date(imagingFinished)
        elif "Acquisition finished:" in each_line:
            imagingFinished = re.split("Acquisition finished: ", each_line, 0)
            imagingFinished = str(imagingFinished[1]).strip()
            imagingFinished = fix_date(imagingFinished)

        elif "Extraction end date" in each_line:     #cellebrite
            imagingFinished = re.split("Extraction end date", each_line, 0)
            imagingFinished = str(imagingFinished[1]).strip()
            imagingFinished = imagingFinished.replace("/time", "").replace(" -05:00", "").strip(':').strip().replace("(GMT-5)", "")

            print(imagingFinished)   #temp
            # imagingFinished = fix_date(imagingFinished)

        elif "Unique description: " in each_line:
            makeModel = re.split("Unique description: ", each_line, 0)
            print("makeModel=", makeModel[1].strip())      
            makeModel = str(makeModel[1]).strip()

        elif "Device	" in each_line: #cellebrite excel
            makeModel = re.split("Device	", each_line, 0)
            print("makeModel=", makeModel[1].strip())      
            makeModel = str(makeModel[1]).strip()
        elif "Selected device name" in each_line: #cellebrite
            makeModel = re.split("Selected device name", each_line, 0)
            print("makeModel=", makeModel[1].strip())      
            makeModel = str(makeModel[1]).strip()

        elif "Selected Model:" in each_line:
            makeModel = re.split("Selected Model:", each_line, 0)
            print("makeModel=", makeModel[1].strip())      
            makeModel = str(makeModel[1]).strip()

        elif "Model:" in each_line and len(model) == 0:
            model = re.split("Model:", each_line, 0)
            model = str(model[1]).strip()
            # makeModel = model

        elif "Revision:" in each_line: #cellebite
            os = re.split("Revision:", each_line, 0)
            os = str(os[1]).strip()
            if 'iPhone' in makeModel:
                os = ('iOS %s' %(os))

        elif "Serial Number:" in each_line: #cellebrite
            serial = re.split("Serial Number:", each_line, 0)
            print("serial=",serial[1].strip())      
            serial = str(serial[1]).strip()
        elif "Serial " in each_line: #cellebrite
            serial = re.split("Serial ", each_line, 0)
            print("serial=",serial[1].strip())      
            serial = str(serial[1]).strip()

        elif "MSISDN" in each_line: #cellebrite
            phoneNumber = re.split("MSISDN", each_line, 0)
            phoneNumber = str(phoneNumber[1]).strip()
            if ')' in phoneNumber:
                phoneNumber = phoneNumber.replace("+1 (", "1-").replace(") ", "-")
            print("phoneNumber=",phoneNumber)
            (exportedEvidence, status) = ('', 'Imaged')

        elif " Username" in each_line: #cellebrite xls
            phoneNumber = re.split(" Username", each_line, 0)
            phoneNumber = str(phoneNumber[1]).strip()
            if ')' in phoneNumber:
                phoneNumber = phoneNumber.replace("+1 (", "1-").replace(") ", "-")
            print("phoneNumber=",phoneNumber)
            (exportedEvidence, status) = ('', 'Imaged')
 
        elif "User: " in each_line:
            forensicExaminer = re.split("User: ", each_line, 0)
            print("forensicExaminer ", forensicExaminer[1].strip())      
            forensicExaminer = str(forensicExaminer[1]).strip()
        elif "Examiner:" in each_line:
            forensicExaminer = re.split("Examiner:", each_line, 0)
            print("forensicExaminer=", forensicExaminer[1].strip())      
            forensicExaminer = str(forensicExaminer[1]).strip()
            forensicExaminer =forensicExaminer.replace("CIA - ", "")
        elif "Examiner name" in each_line:  #cellebrite
            forensicExaminer = re.split("Examiner name", each_line, 0)
            print("forensicExaminer ", forensicExaminer[1].strip())      
            forensicExaminer = str(forensicExaminer[1]).strip()
            forensicExaminer =forensicExaminer.replace("CIA - ", "")

        elif "Case ID:" in each_line:
            caseNumber = re.split("Case ID:", each_line, 0)
            caseNumber = str(caseNumber[1]).strip()
            caseNumber = caseNumber.replace("<<not entered>>", "")
        elif "Case Number:" in each_line:
            caseNumber = re.split("Case Number:", each_line, 0)
            caseNumber = str(caseNumber[1]).strip()
            caseNumber = caseNumber.replace("<<not entered>>", "")
        elif "CaseNumber" in each_line:   #cellebrite
            caseNumber = re.split("CaseNumber", each_line, 0)
            caseNumber = str(caseNumber[1]).strip()
            print(caseNumber)   #temp

        elif "Case Notes:" in each_line:
            notes = re.split("Case Notes:", each_line, 0)
            notes = str(notes[1]).strip()
            notes = notes.replace("<<not entered>>", "")
        elif "Notes: " in each_line:
            notes = re.split("Notes: ", each_line, 0)
            notes = str(notes[1]).strip()
            notes = notes.replace("<<not entered>>", "")

        elif "Imager App: " in each_line:
            imagingTool1 = re.split("Imager App: ", each_line, 0)
            imagingTool1 = str(imagingTool1[1]).strip()
        elif "Created By AccessData" in each_line:
            imagingTool1 = each_line.replace("Created By AccessData", "").replace("®", "").strip()
            # imagingTool = re.split(" ", each_line, 0)
            # imagingTool = str(imagingTool[5]).strip()
            # imagingTool = ('FTK Imager %s' %(imagingTool))
            
        elif "Imager Ver: " in each_line:
            imagingTool2 = re.split("Imager Ver: ", each_line, 0)
            imagingTool2 = str(imagingTool2[1]).strip()

        elif "UFED Version:	Product Version: " in each_line:    #cellebrite
            imagingTool = re.split("UFED Version:	Product Version: ", each_line, 0)
            imagingTool = str(imagingTool[1]).strip()
            imagingTool = re.split(" ", imagingTool, 0)
            imagingTool = str(imagingTool[0]).strip()
            imagingTool = ('Cellebrite UFED %s' %(imagingTool))
        elif "UFED version" in each_line:    #cellebrite
            imagingTool = re.split("UFED version", each_line, 0)
            imagingTool = str(imagingTool[1]).strip()
            imagingTool = re.split(" ", imagingTool, 0)
            imagingTool = str(imagingTool[0]).strip()
            imagingTool = ('Cellebrite UFED %s' %(imagingTool))

        elif "Cellebrite Physical Analyzer version" in each_line: #cellebrite xls
            imagingTool1 = re.split("Cellebrite Physical Analyzer version", each_line, 0)
            imagingTool1 = str(imagingTool1[1]).strip()
            imagingTool1 = ('Cellebrite Physical Analyzer %s' %(imagingTool1))
            print(imagingTool1) #temp

        elif "Acquired using: ADI3" in each_line:
            imagingTool = "FTK Imager"


        elif "S/N: " in each_line:
            hddserial = re.split("S/N: ", each_line, 0)
            hddserial = str(hddserial[1]).strip()
            # serial = hddserial
        elif "Drive Serial Number:" in each_line:
            hddserial = re.split("Drive Serial Number:", each_line, 0)
            hddserial = str(hddserial[1]).strip()
            # serial = hddserial

        elif "Capacity in bytes reported Pwr-ON: " in each_line:
            capacity = re.split("Capacity in bytes reported Pwr-ON: ", each_line, 0)
            capacity = str(capacity[1]).strip()
            if "(" in capacity:
                capacity = re.split("\(", each_line, 0)
                capacity = str(capacity[1]).strip()
                capcty = capacity.replace(")", "")
                capacity = capcty.split('.')[0]
                if ' ' in capcty:
                    size = capcty.split(' ')[1]
                else:
                    size = ''
                # size = capacity
                capacity = ('%s %s' %(capacity, size))
        elif "Source data size: " in each_line:
            capacity = re.split("Source data size: ", each_line, 0)
            capacity = str(capacity[1]).strip()
            if "(" in capacity:
                capacity = re.split("\(", each_line, 0)
                capacity = str(capacity[1]).strip()
                capcty = capacity.replace(")", "")
                capacity = capcty.split('.')[0]
                if ' ' in capcty:
                    size = capcty.split(' ')[1]
                # size = capacity
                capacity = ('%s %s' %(capacity, size))
                
        elif "Filename of first chunk: " in each_line:
            exportLocation = re.split("Filename of first chunk: ", each_line, 0)
            exportLocation = str(exportLocation[1]).strip()
        elif "Information for " in each_line:       # ftk_parse
            exportLocation = re.split("Information for ", each_line, 0)
            exportLocation = str(exportLocation[1]).strip()


        elif "E01" in each_line:
            exportLocation = each_line.strip()

        elif "Disk MD5:  " in each_line:
            imageMD5 = re.split("Disk MD5:  ", each_line, 0)
            imageMD5 = str(imageMD5[1]).strip()
        elif "MD5 checksum:" in each_line:  # fix me
            imageMD5 = re.split("MD5 checksum:", each_line, 0)
            imageMD5 = str(imageMD5[1]).strip()
            imageMD5 = re.split(": ", each_line, 0)
            imageMD5 = str(imageMD5[1]).strip()
            if "verified" in each_line:
                status = "Imaged"
                imageMD5 = imageMD5.replace(' : verified','')
            else:
                status = 'Not imaged'


        elif "Host Name: " in each_line:
            hostname = re.split("Host Name: ", each_line, 0)
            hostname = str(hostname[1]).strip()
            notes = ("%s The hostname is %s." %(notes, hostname))
        elif "Timezone: " in each_line:
            timezone = re.split("Timezone: ", each_line, 0)
            timezone = str(timezone[1]).strip()
            notes = ("%s The system timezone is set to %s." %(notes, timezone))
        elif "OS Name: " in each_line:
            os = re.split("OS Name: ", each_line, 0)
            os = str(os[1]).strip()
            # notes = ("%s The operating system was %s." %(notes, os)) 
        elif "   IPv4 Address" in each_line:
            ip = re.split("   IPv4 Address. . . . . . . . . . . : ", each_line, 0)
            ip = str(ip[1]).strip()
            notes = ("%s The IP address is %s." %(notes, ip))
        elif "    Lock Status:" in each_line:
            encryption = re.split("    Lock Status:", each_line, 0)
            encryption = str(encryption[1]).strip()
            if 'Locked' in encryption:
                encryption = 'BitLocker Encrypted'
                notes = ("%s BitLocker encryption is enabled." %(notes)) 

    if len(imagingTool1) != 0:
        imagingTool = ('%s %s' %(imagingTool1.strip(), imagingTool2.strip()))
    
    
    if len(capacity) != 0:
        notes = ("This had a %s drive, model %s, serial # %s. %s" %(capacity, model, hddserial, notes))

    if len(OS) != 0 and 'The operating system was' not in notes:
        notes = ("%s The operating system was %s." %(notes, OS)) 


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
                notes,attachment, tempNotes, inventoryDate, operation, Action, imageSHA256, OS, dateSeized, summary)


def read_text():
    # global Row    #The magic to pass Row globally
    style = workbook.add_format()
    (header, report, date) = ('', '', '<insert date here>')
    # csv_file = open(filename) # UnicodeDecodeError: 'charmap' codec can't decode byte 0x9d
    csv_file = open(filename, encoding='utf8')
    # with open(filename, encoding='utf8') as csv_file:
        # html = BeautifulSoup(csv_file, "html.parser")
        # csv_file = csv_file
    
    outputFile = "report.txt"
    output = open(outputFile, 'w+')
    (subject, vowel) = ('test', 'aeiou')

    # footer = '''  

    # The images of all the devices will be retained. The case agent may request additional analysis or files to be exported if new evidence of probative value is determined, at a future date.
    
# Evidence:
    # Reports and supporting files were exported and given to the case agent.
    # '''

    footer = '''  

Evidence:
    All digital images obtained pursuant to this investigation will be maintained on IDOR servers for five years past the date of adjudication and/or case discontinuance. Copies of digital images will be made available upon request. All files copied from the images and provided to the case agent for review are identified as the DIGITAL EVIDENCE FILE and will be included as an exhibit in the case file. 
    '''
    
    
    for each_line in csv_file:
        (caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType) = ('', '', '', '', '', '', '')
        (caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived) = ('', '', '', '', '', '')
        (serial, makeModel, storageLocation, removalDate, exportedEvidence, status) = ('', '', '', '', '', '')
        (analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff) = ('', '', '', '', '', '')
        (notes, attachment, tempNotes, report) = ('', '', '', '')
        (tempNotes,inventoryDate,operation,Action,imageSHA256,OS,dateSeized) = ('', '', '', '', '', '', '')
        (inventoryDate, operation, Action, imageSHA256, summary) = ('', '', '', '', '')
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
            caseType = each_line[3].lower()
            caseAgent = each_line[4]
            forensicExaminer = each_line[5]
            exhibit = each_line[6]
            makeModel = each_line[7].strip()
            serial = each_line[8].strip()
            phoneNumber = each_line[9]
            imagingStarted = each_line[10]
            imagingFinished = each_line[11]
            imagingTool = each_line[12]
            imagingType = each_line[13].lower()
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
            tempNotes = each_line[27]
            inventoryDate = each_line[28]
            operation = each_line[29]
            Action = each_line[30]
            imageSHA256 = each_line[31]
            OS = each_line[32]
            dateSeized = each_line[33]
            summary = each_line[34]
            
            if subject == 'test':
                subject = subjectBusinessName
                # future idea. if subjectBusinessName != subject: Exhibit # <Exhibit> <subjectBusinessName>


        if header == '':
            header = ('''
ACTIVITY REPORT                              BUREAU OF CRIMINAL INVESTIGATIONS
____________________________________________________________________________________

Activity Number:                             Date of Activity:
%s                               %s
____________________________________________________________________________________
Subject of Activity:                         Case Agent:             Typed by:
%s %s                           %s    %s
%s
____________________________________________________________________________________


Executive Summary 
    Special Agent %s of the Illinois Department of Revenue, Bureau of Criminal Investigations, requested an examination of evidence for any information regarding the %s investigation in the %s case. %s
''') %(caseNumber, todaysDate, caseName, subjectBusinessName, caseAgent, forensicExaminer, caseType.lower(), caseAgent, caseType, caseName, summary)

            output.write(header+'\n')
        
        # if len(summary) != 0:
            # summarize = ('''
    # Summary
# ''')            
        
        
        # report = ('''
        
# Exhibit # %s
    # ''') %(exhibit)

        # report = ('''
        
# Exhibit # %s, Case # %s %s
    # ''') %(exhibit, caseNumber, )

        # report = ('''
        
# Case # %s Exhibit %s (%s)
    # ''') %(caseNumber, exhibit, subjectBusinessName)

        # report = ('''
        
# Case # %s Exhibit %s
    # ''') %(caseNumber, exhibit)

        report = ('''
        
Exhibit %s
    ''') %(exhibit)

        if makeModel != '':
            if makeModel[0].lower() in vowel:
                report = ('''%sAn %s''') %(report, makeModel)
            else:
                report = ('''%sA %s''') %(report, makeModel)

        if len(OS) != 0:
            if OS[0].lower() in vowel:
                report = ("%s, with an %s OS" %(report, OS))
            else:
                report = ("%s, with a %s OS" %(report, OS))          

        if len(serial) != 0:
            report = ("%s, serial # %s" %(report, serial))
        if len(dateReceived) != 0:
            report = ("%s, was received on %s" %(report, dateReceived.replace(" ", " at ", 1)))
        else:
            report = ("%s, was received" %(report))
        report = ("%s." %(report))
        
        # if len(imagingStarted) != 0:
        if len(imagingStarted) != 0 and status != "Not imaged":
            report = ("%s On %s," %(report, imagingStarted.replace(" ", " at ", 1)))
        report = ("%s Digital Forensic Examiner %s" %(report, forensicExaminer))
        if len(imagingTool) != 0 and imagingType != '':
            report = ("%s used %s to conduct a %s" %(report, imagingTool, imagingType.lower()))  
        elif imagingTool != '':
            report = ("%s used %s to conduct " %(report, imagingTool))  

        elif imagingType != '' and exportedEvidence != "N":
            report = ("%s conducted a %s" %(report, imagingType))  
        elif exportedEvidence == "N":
            report = ("%s did not conduct a" %(report))  
        else:
            report = ("%s conducted a" %(report))  

            
        if phoneNumber != '':
            report = ("%s phone extraction." %(report))
            if phoneNumber.lower() != 'unknown':
                report = ("%s The mobile Station International Subscriber Number (MSISDN) was %s." %(report, phoneNumber))
        else:
            report = ("%s forensic extraction." %(report))

        if len(imageMD5) != 0 and exportLocation != '':
            report = ("%s The image, which had a MD5 hash of % s, was saved as %s." %(report, imageMD5, exportLocation.split('\\')[-1])) 

        # if len(imageSHA256) != 0 and exportLocation != '':
        if len(imageSHA256) != 0:
            report = ("%s The image had a SHA256 hash of % s." %(report, imageSHA256))

        
        if analysisTool != '':
            report = ("%s The image was processed with %s." %(report, analysisTool))
        
        if notes != '':
            report = ("%s %s" %(report, notes))

        if exportedEvidence == "Y" and 'elevant files were exported' not in notes:
            # report = ("%s Relevant files were exported." %(report.strip()))
            report = ("%s Relevant files were exported." %(report.rstrip()))
        elif exportedEvidence == "N" and 'search for relevant files was made and no files were found' not in notes:
            report = ("%s A search for relevant files was made and no files were found." %(report.rstrip()))
        
        report = report.replace("    , was received. ", "    ")
        report = report.replace("This was a DVR system was not imaged.","This was a DVR system and was not imaged.")
        report = report.replace("Digital Forensic Examiner Casey Karaffa did not conduct a forensic extraction.","This was not imaged.")
        report = report.replace("The image was processed with copy.","Pertinent files were copied.")
        report = report.replace("This had a  drive, model , serial # .","") # fixme     
        notes = notes.replace("This had a  drive, model , serial # .","")  # fixme
        report = report.replace(", serial # .",".") # fixme 
        notes = notes.replace(", serial # .",".") # fixme 
        
        print(report)
        output.write(report)

        # Write excel
        write_report(caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType,
                caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived,
                serial, makeModel, storageLocation, removalDate, exportedEvidence, status,
                analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff,
                notes,attachment, tempNotes, inventoryDate, operation, Action, imageSHA256, OS, dateSeized, summary)

    output.write(footer+'\n')

def write_report(caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType,
                caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived,
                serial, makeModel, storageLocation, removalDate, exportedEvidence, status,
                analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff,
                notes,attachment, tempNotes, inventoryDate, operation, Action, imageSHA256, OS, dateSeized, summary):

    global Row

    Sheet1.write_string(Row, 0, caseNumber)
    Sheet1.write_string(Row, 1, caseName)
    Sheet1.write_string(Row, 2, subjectBusinessName)
    Sheet1.write_string(Row, 3, caseType)
    Sheet1.write_string(Row, 4, caseAgent)
    Sheet1.write_string(Row, 5, forensicExaminer)
    Sheet1.write_string(Row, 6, exhibit)
    Sheet1.write_string(Row, 7, makeModel)
    Sheet1.write_string(Row, 8, serial)
    Sheet1.write_string(Row, 9, phoneNumber)
    Sheet1.write_string(Row, 10, imagingStarted)
    Sheet1.write_string(Row, 11, imagingFinished)
    Sheet1.write_string(Row, 12, imagingTool)
    Sheet1.write_string(Row, 13, imagingType)
    Sheet1.write_string(Row, 14, storageLocation)
    Sheet1.write_string(Row, 15, dateReceived) 
    Sheet1.write_string(Row, 16, removalDate)
    Sheet1.write_string(Row, 17, exportedEvidence)
    Sheet1.write_string(Row, 18, status)
    Sheet1.write_string(Row, 19, analysisTool)
    Sheet1.write_string(Row, 20, exportLocation)
    Sheet1.write_string(Row, 21, imageMD5)
    Sheet1.write_string(Row, 22, locationOfCaseFile)
    Sheet1.write_string(Row, 23, reasonForRemoval)
    Sheet1.write_string(Row, 24, removalStaff)
    try:
        Sheet1.write_string(Row, 25, notes)
    except:pass
    Sheet1.write_string(Row, 26, attachment)
    Sheet1.write_string(Row, 27, tempNotes)  
    Sheet1.write_string(Row, 28, inventoryDate)
    Sheet1.write_string(Row, 29, operation)
    Sheet1.write_string(Row, 30, Action)
    Sheet1.write_string(Row, 31, imageSHA256)
    Sheet1.write_string(Row, 32, OS)
    Sheet1.write_string(Row, 33, dateSeized)
    Sheet1.write_string(Row, 34, summary)

    Row += 1

def write_sticker():
    # global Row    #The magic to pass Row globally
    style = workbook.add_format()
    (header, report, date) = ('', '', '<insert date here>')
    # csv_file = open(filename)
    csv_file = open(filename, encoding='utf8')
    outputFile = "report.txt"
    output = open(outputFile, 'w+')


    # footer = '''  

# The images of all the devices will be retained. The case agent may request additional analysis or files to be exported if new evidence of probative value is determined, at a future date.
    
# Evidence:
    # Reports and supporting files were exported and given to the case agent.
    # '''

    footer = '''  

All digital images obtained pursuant to this investigation will be maintained on IDOR servers for five years past the date of adjudication and/or case discontinuance. Copies of digital images will be made available upon request. All files copied from the images and provided to the case agent for review are identified as the DIGITAL EVIDENCE FILE and will be included as an exhibit in the case file. 
    '''
    
    for each_line in csv_file:
        (caseNumber, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType) = ('', '', '', '', '', '', '')
        (caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived) = ('', '', '', '', '', '')
        (serial, makeModel, storageLocation, removalDate, exportedEvidence, status) = ('', '', '', '', '', '')
        (analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff) = ('', '', '', '', '', '')
        (notes, attachment, tempNotes) = ('', '', '')
        (inventoryDate, operation, Action, imageSHA256, OS, dateSeized, summary) = ('', '', '', '', '', '', '')

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
            tempNotes = each_line[27]
            inventoryDate = each_line[28]
            operation = each_line[29]
            Action = each_line[30]
            imageSHA256 = each_line[31]
            OS = each_line[32]
            dateSeized = each_line[33]
            summary = each_line[34]
            
        header = ('''Case#: %s  Ex: %s
CaseName: %s
Subject: %s
Make: %s 
Serial: %s
Agent: %s
%s
''') %(caseNumber, exhibit, caseName, subjectBusinessName, makeModel, serial, caseAgent, status)
        header = header.strip()

# write it one line at at time. If phone isn't blank, include it

        # Write excel
        write_report(header, exhibit, imagingStarted, imagingFinished, caseName, subjectBusinessName, caseType,
                caseAgent, forensicExaminer, imagingTool, imagingType, phoneNumber, dateReceived,
                serial, makeModel, storageLocation, removalDate, exportedEvidence, status,
                analysisTool, exportLocation, imageMD5, locationOfCaseFile, reasonForRemoval, removalStaff,
                notes,attachment, tempNotes, inventoryDate, operation, Action, imageSHA256, OS, dateSeized, summary)

        output.write(header+'\n')


def usage():
    file = sys.argv[0].split('\\')[-1]
    print("\nDescription: " + description)
    print(file + " Version: %s by %s" % (version, author))
    print("\nExample:")
    # print("\t" + file + " -f -I input.txt -O out_log_.xlsx\t\t")
    print("\t" + file + " -r -I input.txt -O out_cases_.xlsx\t\t")
    print("\t" + file + " -s -I input.txt -O out_log_.xlsx\t\t")
    print("\t" + file + " -l -I input.txt -O out_log_.xlsx\t\t")
    
if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
1.0.1 - Created a Tableau log parser
1.0.0 - Created forensic report writer
0.1.2 - converted tabs to 4 spaces for #pep8
0.0.2 - python2to3 conversion
1.3.6 - Added summary and OS column

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
output to pdf or Doc?
if ', serial # .', replace with .

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

