#!/usr/bin/python2
# coding: utf-8

############################        Imports        ############################
 
import re            # regular expression
import os
import sys
import time
import random
import argparse        # for menu system
import xlsxwriter
from datetime import date
from subprocess import call

d = date.today()
Day    = d.strftime("%d")
Month = d.strftime("%B")
Year  = d.strftime("%Y")        
TodaysDate = d.strftime("%m/%d/%Y")

# # Should be copied to /usr/lib/python2.7/TSU.py
# TSU.py translates IP's to Agencies
try:
    from TSU import check_agency, factory, check_port, password_defaults
except:
    print ('you are missing /usr/lib/python2.7/TSU.py')
    
############################       Pre-Sets        ############################

Author        = 'LincolnLandForensics'
description     = "Read a file or pdf and export them various formats";
tech            = 'LincolnLandForensics'        # change this to your name
version        = '2.2.7'

# regex section

regex_bitcoin = re.compile(r'/^(bc1|[13])[a-zA-HJ-NP-Z0-9]{25,39}$')    #FixMe



regex_email = re.compile(r'/^[a-z0-9.!#$%&\'*+\\/=?^_`{|}~-]+@[a-z0-9-]+(?:\\.[a-z0-9-]+)*$/i')    #regex_email

regex_host = re.compile(r'\b((?:(?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+(?i)(?!exe|php|dll|doc' \
        '|docx|txt|rtf|odt|xls|xlsx|ppt|pptx|bin|pcap|ioc|pdf|mdb|asp|html|xml|jpg|gif$|png' \
        '|lnk|log|vbs|lco|bat|shell|quit|pdb|vbp|bdoda|bsspx|save|cpl|wav|tmp|close|ico|ini' \
        '|sleep|run|dat$|scr|jar|jxr|apt|w32|css|js|xpi|class|apk|rar|zip|hlp|cpp|crl' \
        '|cfg|cer|plg|lxdns|cgi|xn$)(?:xn--[a-zA-Z0-9]{2,22}|[a-zA-Z]{2,13}))(?:\s|$)')

regex_md5 = re.compile(r'^([a-fA-F\d]{32})$')    #regex_md5        [a-f0-9]{32}$/gm
regex_number = re.compile(r'^(^\d)$')    #regex_number    #Beta
regex_phone = re.compile(r'/^[+]?(?=(?:[^\\dx]*\\d){7})(?:\\(\\d+(?:\\.\\d+)?\\)|\\d+(?:\\.\\d+)?)(?:[ -]?(?:\\(\\d+(?:\\.\\d+)?\\)|\\d+(?:\\.\\d+)?))*(?:[ ]?(?:x|ext)\\.?[ ]?\\d{1,5})?$/')    #regex_phone    #Beta
 

regex_sha1 = re.compile(r'^([a-fA-F\d]{40})$')    #regex_sha1
regex_sha256 = re.compile(r'^([a-fA-F\d]{64})$')#regex_sha256
regex_sha512 = re.compile(r'^([a-fA-F\d]{128})$')#regex_sha256

   
regex_ipv4 = re.compile('(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}' +
                 '(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)')
regex_ipv6 = re.compile('(S*([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}S*|S*(' +
                 '[0-9a-fA-F]{1,4}:){1,7}:S*|S*([0-9a-fA-F]{1,4}:)' +
                 '{1,6}:[0-9a-fA-F]{1,4}S*|S*([0-9a-fA-F]{1,4}:)' +
                 '{1,5}(:[0-9a-fA-F]{1,4}){1,2}S*|S*([0-9a-fA-F]' +
                 '{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}S*|S*(' +
                 '[0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}S*' +
                 '|S*([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4})' +
                 '{1,5}S*|S*[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4})' +
                 '{1,6})S*|S*:((:[0-9a-fA-F]{1,4}){1,7}|:)S*|::(ffff' +
                 '(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}' +
                 '[0-9]){0,1}[0-9]).){3,3}(25[0-5]|(2[0-4]|1{0,1}[' +
                 '0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[' +
                 '0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]).){3,3}(25[' +
                 '0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))')    


# regex_vin = re.compile(r'/^(?<wmi>[A-HJ-NPR-Z\\d]{3})(?<vds>[A-HJ-NPR-Z\\d]{5})(?<check>[\\dX])(?<vis>(?<year>[A-HJ-NPR-Z\\d])(?<plant>[A-HJ-NPR-Z\\d])(?<seq>[A-HJ-NPR-Z\\d]{6}))$/') #regex_vin

#Colors
if sys.platform == 'win32' or sys.platform == 'win64':    
    # if windows, don't use colors
    (r,o,y,g,b) = ('','','','','')
else:
    r             = '\033[31m'     #red
    o             = '\033[0m'     #off
    y             = '\033[33m'     #yellow
    g             = '\033[32m'     #green
    b             = '\033[34m'     #blue


############################         Menu         #############################

def main():                            
    global Row
    Row = 1    
    
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I','--input', help='', required=False)
    parser.add_argument('-O','--output', help='', required=False)
    parser.add_argument('-a','--anamoli', help='Anomoli Output', required=False, action='store_true')
    parser.add_argument('-f','--file', help='file import', required=False, action='store_true')
    parser.add_argument('-Q','--qradar', help='Qradar Output', required=False, action='store_true')
    parser.add_argument('-T','--test', help='test import', required=False, action='store_true')
    parser.add_argument('-t','--threatconnect', help='threatconnect Output', required=False, action='store_true')
    
    args = parser.parse_args()

    cls()
    print_logo()

    if not(args.output):                                     # this section might be redundant
        parser.print_help()
        usage()
        return 0

    global SheetFormat


    if args.output:
        global filename
        global spreadsheet
        spreadsheet = args.output
        global output_file
        output_file = open('out_backup_.csv', 'w')    #Backup csv file
        if args.qradar:
            SheetFormat = "qradar"
            create_xlsx()
        elif args.anamoli:
            SheetFormat = "anamoli"
            create_xlsx()
        elif args.threatconnect:
            SheetFormat = "threatconnect"
            create_xlsx_tc()
        else:
            SheetFormat = "basic"
            create_xlsx_basic()

        try:
            global filename
            filename = args.input

            # InputFile = args.input
        except:
            print("missing input file so using combo.txt")
            # InputFile = 'combo.txt'
            filename = 'combo.txt'
            filename = args.input
            filename = str(filename)    #FixMe
            
        # output_file = open(args.output, 'w')
        
        if args.file:                                 
            # ReadFile(InputFile,output_file)
            read_file()
        elif args.test:                                 
            # ReadTest(InputFile,output_file)
            read_test()
            
    if sys.platform == 'win32' or sys.platform == 'win64':
        pass
    else:
        call(["chown %s.%s *.csv" %(tech.lower(), tech.lower())],shell=True)
        call(["chown %s.%s *.xlsx" %(tech.lower(), tech.lower())],shell=True)

    workbook.close()
    return 0
    exit()


############################      Sub-Routines     ############################

def cls():
    linux = 'clear'
    windows = 'cls'
    os.system([linux, windows][os.name == 'nt'])

def create_xlsx():        # with hidden columns
    global workbook
    workbook = xlsxwriter.Workbook(spreadsheet)
    global sheet1
    sheet1 = workbook.add_worksheet('data')
    HeaderFormat = workbook.add_format({'bold': True,'border': 1})
    sheet1.freeze_panes(1, 1)    #Freeze cells
    sheet1.set_selection('B2')
    
#Excel column width
    sheet1.set_column(0,0,16) #log_source
    sheet1.set_column(1,1,14) #event_name
    sheet1.set_column(2,2,8) #reason
    sheet1.set_column(3,3,15) #first_seen
    sheet1.set_column(4,4,16) #src_ip
    sheet1.set_column(5,5,8) #src_port
    sheet1.set_column(6,6,16) #dst_ip
    sheet1.set_column(7,7,8) #dst_port
    sheet1.set_column(8,8,8) #src_user
    sheet1.set_column(9,9,20) #dns_domain
    sheet1.set_column(10,10,70) #url
    sheet1.set_column(11,11,8) #Referer/urlcategory
    sheet1.set_column(12,12,8) #urlReputation/type
    sheet1.set_column(13,13,8) #src_hostname
    sheet1.set_column(14,14,12) #dst_hostname
    sheet1.set_column(15,15,5) #count
    sheet1.set_column(16,16,4) #Srccountry
    sheet1.set_column(17,17,4) #Dstcountry/Continent
    sheet1.set_column(18,18,15) #log_src
    sheet1.set_column(19,19,8) #log_srcGroup/Flowtype
    sheet1.set_column(20,20,15) #category
    sheet1.set_column(21,21,5) #src_packets
    sheet1.set_column(22,22,5) #dst_packets
    sheet1.set_column(23,23,40) #payload
    sheet1.set_column(24,24,8) #src_hex
    sheet1.set_column(25,25,8) #src_b64
    sheet1.set_column(26,26,8) #dst_payload
    sheet1.set_column(27,27,8) #dst_hex
    sheet1.set_column(28,28,8) #dst_b64
    sheet1.set_column(29,29,6) #src_agency
    sheet1.set_column(30,30,6) #dst_agency
    sheet1.set_column(31,31,5) #src_state
    sheet1.set_column(32,32,5) #dst_state
    sheet1.set_column(33,33,8) #tech
    sheet1.set_column(34,34,8) #ticket
    sheet1.set_column(35,35,15) #action
    sheet1.set_column(36,36,16) #note
    sheet1.set_column(37,37,8) #notes2
    sheet1.set_column(38,38,20) #os_url
    sheet1.set_column(39,39,28) #title_url
    sheet1.set_column(40,40,20) #page_status
    sheet1.set_column(41,41,17) #internal_ip
    sheet1.set_column(42,42,20) #internal_hostname
    sheet1.set_column(43,43,30) #client
    sheet1.set_column(44,44,8) #severity
    sheet1.set_column(45,45,9) #resolution

#Hidden columns
    sheet1.set_column('E:E', None, None, {'hidden': 1})    # src_ip
    sheet1.set_column('F:F', None, None, {'hidden': 1})    # src_port
    sheet1.set_column('I:I', None, None, {'hidden': 1})    # src_user
    sheet1.set_column('L:L', None, None, {'hidden': 1})    # Referer
    sheet1.set_column('M:M', None, None, {'hidden': 1})    # urlReputation
    sheet1.set_column('N:N', None, None, {'hidden': 1})    # src_hostname
    sheet1.set_column('O:O', None, None, {'hidden': 1})    # dst_hostname
    sheet1.set_column('P:P', None, None, {'hidden': 1})    # count
    sheet1.set_column('Q:Q', None, None, {'hidden': 1})    # Srccountry
    sheet1.set_column('R:R', None, None, {'hidden': 1})    # Dstcountry
    sheet1.set_column('S:S', None, None, {'hidden': 1})    # log_src
    sheet1.set_column('T:T', None, None, {'hidden': 1})    # log_srcGroup
    sheet1.set_column('U:U', None, None, {'hidden': 1})    # category
    sheet1.set_column('V:V', None, None, {'hidden': 1})    # src_packets
    sheet1.set_column('W:W', None, None, {'hidden': 1})    # dst_packets
    sheet1.set_column('X:X', None, None, {'hidden': 1})    # payload
    sheet1.set_column('Y:Y', None, None, {'hidden': 1})    # src_hex
    sheet1.set_column('Z:Z', None, None, {'hidden': 1})    # src_b64
    sheet1.set_column('AA:AA', None, None, {'hidden': 1})    # dst_payload
    sheet1.set_column('AB:AB', None, None, {'hidden': 1})    # dst_hex
    sheet1.set_column('AB:AB', None, None, {'hidden': 1})    # dst_b64
    sheet1.set_column('AC:AC', None, None, {'hidden': 1})    # src_state
    sheet1.set_column('AD:AD', None, None, {'hidden': 1})    # src_agency
    # sheet1.set_column('AE:AE', None, None, {'hidden': 1})    # dst_agency
    sheet1.set_column('AF:AF', None, None, {'hidden': 1})    # src_state
    sheet1.set_column('AG:AG', None, None, {'hidden': 1})    # dst_state
    sheet1.set_column('AH:AH', None, None, {'hidden': 1})    # tech
    sheet1.set_column('AI:AI', None, None, {'hidden': 1})    # ticket
    # sheet1.set_column('AJ:AJ', None, None, {'hidden': 1})    # action
    # sheet1.set_column('AK:AK', None, None, {'hidden': 1})    # note
    sheet1.set_column('AL:AL', None, None, {'hidden': 1})    # notes2
    sheet1.set_column('AM:AM', None, None, {'hidden': 1})    # os_url
    # sheet1.set_column('AN:AN', None, None, {'hidden': 1})    # title_url
    sheet1.set_column('AP:AP', None, None, {'hidden': 1})    # internal_ip
    sheet1.set_column('AQ:AQ', None, None, {'hidden': 1})    # internal_hostname
    sheet1.set_column('AR:AR', None, None, {'hidden': 1})    # client
    
#write column headers    

    sheet1.write(0,0,'log_source', HeaderFormat)
    sheet1.write(0,1,'event_name', HeaderFormat)
    sheet1.write(0,2,'reason', HeaderFormat)
    sheet1.write(0,3,'first_seen', HeaderFormat)
    sheet1.write(0,4,'src_ip', HeaderFormat)
    sheet1.write(0,5,'src_port', HeaderFormat)
    sheet1.write(0,6,'dst_ip', HeaderFormat)
    sheet1.write(0,7,'dst_port', HeaderFormat)
    sheet1.write(0,8,'src_user', HeaderFormat)
    sheet1.write(0,9,'dns_domain', HeaderFormat)
    sheet1.write(0,10,'url', HeaderFormat)
    sheet1.write(0,11,'Referer/urlcategory', HeaderFormat)
    sheet1.write(0,12,'urlReputation/type', HeaderFormat)
    sheet1.write(0,13,'src_hostname', HeaderFormat)
    sheet1.write(0,14,'dst_hostname', HeaderFormat)
    sheet1.write(0,15,'count', HeaderFormat)
    sheet1.write(0,16,'Srccountry', HeaderFormat)
    sheet1.write(0,17,'Dstcountry/Continent', HeaderFormat)
    sheet1.write(0,18,'log_src', HeaderFormat)
    sheet1.write(0,19,'log_srcGroup/Flowtype', HeaderFormat)
    sheet1.write(0,20,'category', HeaderFormat)
    sheet1.write(0,21,'src_packets', HeaderFormat)
    sheet1.write(0,22,'dst_packets', HeaderFormat)
    sheet1.write(0,23,'payload', HeaderFormat)
    sheet1.write(0,24,'src_hex', HeaderFormat)
    sheet1.write(0,25,'src_b64', HeaderFormat)
    sheet1.write(0,26,'dst_payload', HeaderFormat)
    sheet1.write(0,27,'dst_hex', HeaderFormat)
    sheet1.write(0,28,'dst_b64', HeaderFormat)
    sheet1.write(0,29,'src_agency', HeaderFormat)
    sheet1.write(0,30,'dst_agency', HeaderFormat)
    sheet1.write(0,31,'src_state', HeaderFormat)
    sheet1.write(0,32,'dst_state', HeaderFormat)
    sheet1.write(0,33,'tech', HeaderFormat)
    sheet1.write(0,34,'ticket', HeaderFormat)
    sheet1.write(0,35,'action', HeaderFormat)
    sheet1.write(0,36,'note', HeaderFormat)
    sheet1.write(0,37,'notes2', HeaderFormat)
    sheet1.write(0,38,'os_url', HeaderFormat)
    sheet1.write(0,39,'title_url', HeaderFormat)
    sheet1.write(0,40,'page_status', HeaderFormat)
    sheet1.write(0,41,'internal_ip', HeaderFormat)
    sheet1.write(0,42,'internal_hostname', HeaderFormat)
    sheet1.write(0,43,'client', HeaderFormat)
    sheet1.write(0,44,'severity', HeaderFormat)
    sheet1.write(0,45,'resolution', HeaderFormat)

def create_xlsx_basic():    #basic Output
    global workbook
    workbook = xlsxwriter.Workbook(spreadsheet)
    global sheet1
    sheet1 = workbook.add_worksheet('data')
    HeaderFormat = workbook.add_format({'bold': True,'border': 1})
    sheet1.freeze_panes(1, 1)    #Freeze cells
    sheet1.set_selection('B2')

#Excel column width     

    sheet1.set_column(0,0,135) #value
    sheet1.set_column(1,1,15) # description
    sheet1.set_column(2,2,7) #length
    sheet1.set_column(3,3,15) #complexity
    sheet1.set_column(4,4,15) #type
    sheet1.set_column(5,5,8) #priority
    sheet1.set_column(6,6,50) #revrse

#write column headers

    sheet1.write(0,0,'value', HeaderFormat)
    sheet1.write(0,1,'description', HeaderFormat)
    sheet1.write(0,2,'length', HeaderFormat)
    sheet1.write(0,3,'complexity', HeaderFormat)    
    sheet1.write(0,4,'type', HeaderFormat)
    sheet1.write(0,5,'priority', HeaderFormat)
    sheet1.write(0,6,'reverse', HeaderFormat)

def create_xlsx_tc():    #ThreatConnect Output
    global workbook
    workbook = xlsxwriter.Workbook(spreadsheet)
    global sheet1
    sheet1 = workbook.add_worksheet('sheet1')
    HeaderFormat = workbook.add_format({'bold': True,'border': 1})
    sheet1.freeze_panes(1, 1)    #Freeze cells
    sheet1.set_selection('B2')
    

#Excel column width

    sheet1.set_column(0,0,15) #type
    sheet1.set_column(1,1,135) #value
    sheet1.set_column(2,2,8) #threat_rating
    sheet1.set_column(3,3,15) #confidence
    sheet1.set_column(4,4,16) #source
    sheet1.set_column(5,5,8) #description
    sheet1.set_column(6,6,16) #dns
    sheet1.set_column(7,7,8) #whois

#write column headers

    sheet1.write(0,0,'type', HeaderFormat)
    sheet1.write(0,1,'value', HeaderFormat)
    sheet1.write(0,2,'threat_rating', HeaderFormat)
    sheet1.write(0,3,'confidence', HeaderFormat)
    sheet1.write(0,4,'source', HeaderFormat)
    sheet1.write(0,5,'description', HeaderFormat)
    sheet1.write(0,6,'dns', HeaderFormat)
    sheet1.write(0,7,'whois', HeaderFormat)


def FormatFunction(bg_color = 'white'):
    global Format
    Format=workbook.add_format({
    'bg_color' : bg_color
    })       

def print_logo():
    clear = "\x1b[0m"
    colors = [36, 32, 34, 35, 31, 37]

    x = """
 _____ _____   _______ ___  _   _  __  ___ __    __  
|_   _| __\ \_/ /_   _(_  || | | |/__\| _ \ _\ /' _/ 
  | | | _| > , <  | |  / / | 'V' | \/ | v / v |`._`. 
  |_| |___/_/ \_\ |_| |___|!_/ \_!\__/|_|_\__/ |___/ 
                                                    """
    for N, line in enumerate(x.split("\n")):
        sys.stdout.write("\x1b[1;%dm%s%s\n" % (random.choice(colors), line, clear))
        time.sleep(0.05)

def read_test():    
    output_file.write('%s\t%s\t%s\t%s\t%s\n' %('type','value','Email','password','Domain'))
    
    if InputFile.lower().endswith('.pdf'):
        print('%s endswith .pdf' %(InputFile))    #temp
    else:
        print('fix pdf checker')
        
    # input_file = open(InputFile)
    input_file = open(filename, encoding='utf8')
    
    
    source = InputFile
    for each_line in input_file:
        (type,value,Email,password,Domain,Skip) =  ('','','','','','')
        each_line = each_line.strip()
        value = each_line

        # if re.search("[\w\.\+\-]+\@[\w\.\-]+\.[\w\.\-]+", each_line):   # works
        if re.search("[a-z0-9.!#$%&\'*+\\/=?^_`{|}~-]+@[a-z0-9-]+(?:\\.[a-z0-9-]+)", each_line):   # works
        
            # type = Email
            try:
                Email = re.search("[\w\.\+\-]+\@[\w\.\-]+\.[\w\.\-]+", each_line).group(0)    #JohnMagic .group(0)
            except:
                print(("%sError found with     %s") %(r,each_line,y))
            Email = str(Email)
            password = each_line.replace(Email, "")    #Pull out email from password
            password = password.lstrip('')
            password = password.lstrip('\t')
            password = password.lstrip(':')
            password = password.lstrip('|')
        if "\t" in password:    # replace tabs with ;
            password = password.replace('\t',';')
            password = password.lstrip(' ')
            password = password.lstrip(';')
            
        if re.search(r'@', Email) :
            if bool(re.search(r"[\w\.\+\-]+\@[\w\.\-]+\.[\w\.\-]+", Email)):    #regex email new1
                pass
                
                
            if 'illinois.gov' in Email.lower():
                type = '_Email Sender'
            Domain = Email.lower().split('@')[1]    #Domain split


            if any(re.findall(r'\.gov|\.mil|\.edu|\.k12\.', Domain, re.IGNORECASE)):
                print(('%s%s%s%s%s') %(g, Domain, b, Email, o))
                type = Domain

        
        if Email.lower().endswith('.cn') or Email.lower().endswith('.tw'):
            Skip = 'Skip'
        
        output_file.write('%s\t%s\t%s\t%s\t%s\n'%(type,value,Email,password,Domain))
    
    
def read_file():    
    global Row    #The magic to pass Row globally
    Style = workbook.add_format()
    Color = 'white'

    CsvFile = open(filename, encoding='utf8')   # UnicodeDecodeError: 'charmap' codec can't decode byte 0x9d

    global event_name
    global first_seen
    global log_source
    global threat_rating
    global confidence
    
    (log_source,event_name,reason,first_seen,src_ip,src_port) = ('','','','','','')
    (dst_ip,dst_port,src_user,dns_domain,url,Referer) = ('','','','','','')
    (urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry) = ('','','','','','')
    (log_source,log_srcGroup,category,src_packets,dst_packets,payload) = ('','','','','','')
    (src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency) = ('','','','','','')
    (dst_agency,src_state,dst_state,tech,ticket,action) = ('','','','','','')
    (note,notes2,os_url,title_url,page_status,internal_ip) = ('','','','','','')
    (internal_hostname,client,severity,resolution) = ('','','','')
    reverse = []
    (UniqRow,MyDump) = ('','')
    (threat_rating, confidence) = ('','')
    # input_file = open(InputFile)

    # if InputFile.lower().endswith('.pdf'):
    if filename.lower().endswith('.pdf'):
        
        
        
        (input_file) = ('')

        import pyPdf    # pip install pyPdf

        pdf = pyPdf.PdfFileReader(open(filename, "rb"))

        # pdf = pyPdf.PdfFileReader(open(InputFile, "rb"))
        for page in pdf.pages:
            input_file = ('%s    %s' %(input_file,page.extractText()))
            # print page.extractText()    #temp
            # print '/n/n'

        # print(input_file)    #temp    #success
        # print(input_file)    #temp
    else:
        # input_file = open(filename)    #works
        input_file = open(filename, encoding='utf8')
        
        # import codecs
        # with codecs.open(InputFile, "r",encoding='utf-8', errors='ignore') as input_file:        #test
        
        
        # print("input_file = %s" %(input_file))    #temp
        (log_src, event_name, threat_rating) = ('', '', '')
    if sys.version_info[0] >= 3:    #check for python3 
        if SheetFormat != 'basic':
            log_src = input('log_src= (ex. Threatconnect)(Default = <none>)')
            event_name = input('event_name= (ex. Emotet)(Default = <none>)')
            threat_rating = input('threat_rating= (pick a # between 1 & 5)(Default = <none>)')

        if SheetFormat == 'qradar':
            print(('%s ') % (y))
            first_seen = input('Date= (ex. 9/19/2017)(Default = <TodaysDate>)') 
            
            if first_seen == '': first_seen = TodaysDate    #temp
            print ('\n')
        elif SheetFormat == 'threatconnect':
            confidence = input('confidence= (pick a # between 1 & 100)(Default = <none>)')
   
    else:
        log_src = raw_input('log_src= (ex. Threatconnect)(Default = <none>)')
        event_name = raw_input('event_name= (ex. Emotet)(Default = <none>)')
        threat_rating = raw_input('threat_rating= (pick a # between 1 & 5)(Default = <none>)')

        if SheetFormat == 'qradar':
            print(('%s ') % (y))
            first_seen = raw_input('Date= (ex. 9/19/2017)(Default = <TodaysDate>)') 
            
            if first_seen == '': first_seen = TodaysDate    #temp
            print ('\n')
        elif SheetFormat == 'threatconnect':
            confidence = raw_input('confidence= (pick a # between 1 & 100)(Default = <none>)')

    print(o)
        
    for EachRow in input_file:
        # EachRow = unicode(EachRow, errors='replace')    #test
        # EachRow = unicode(EachRow, errors='ignore')
        
        if EachRow not in UniqRow:
            EachRow = EachRow.strip()
            UniqRow = ('%s\n%s') %(UniqRow,EachRow)
    
    # print(UniqRow)    #temp
    if filename.lower().endswith('.pdf'):
        MyList = input_file
    else:
        MyList = UniqRow
    
    
    MyList = MyList.replace('\n', ' ')        #FixMe group these into one find and replace  maybe | them
    MyList = MyList.replace('\t', ' ')
    MyList = MyList.replace(',', ' ')
    MyList = MyList.replace('\"', ' ')    #fixme
    MyList = MyList.split(' ')    #splits on each space

    try:
        MyList = MyList.strip()
    except:
        pass
    MyList = set(MyList)
    MyList = list(MyList)
    MyList.sort()

    for EachRow in MyList:
        print(EachRow)    #temp
        try:
            EachRow = EachRow.encode('utf8')     #test
        except:
            EachRow = ''
        
        MyDump = ('%s\n%s') %(MyDump,EachRow)


#Color
    if 'yahoo' in url:        #
        FormatFunction(bg_color = 'orange')
        # print('Yahoo')    #temp
    elif 'google' in url:        #
        FormatFunction(bg_color = 'green')
        # print('Google')    #temp
    else:
        FormatFunction(bg_color = 'white')

#write Qradar excel
    # write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution)

    # if SheetFormat == 'qradar':
    #write XLSX
        # write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution)
        # print('%s%s    %s%s    %s    %s%s    %s' %(b,note,g,os_url,r,title_url,o,page_status))
        
    #write backup csv    
        # output_file.write('%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\n' %('log_source','event_name','reason','first_seen','src_ip','src_port','dst_ip','dst_port','src_user','dns_domain','url','Referer/urlcategory','urlReputation/type','src_hostname','dst_hostname','count','Srccountry','Dstcountry/Continent','log_src','log_srcGroup/Flowtype','category','src_packets','dst_packets','payload','src_hex','src_b64','dst_payload','dst_hex','dst_b64','src_agency','dst_agency','src_state','dst_state','tech','ticket','action','note','notes2','os_url','title_url','page_status','internal_ip','internal_hostname','client','severity','resolution'))
        # print('%s%s    %s%s    %s    %s%s    %s' %(b,note,g,os_url,r,title_url,o,page_status))
        # try:
            # output_file.write('%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\n' %('log_source','event_name','reason','first_seen','src_ip','src_port','dst_ip','dst_port','src_user','dns_domain','url','Referer/urlcategory','urlReputation/type','src_hostname','dst_hostname','count','Srccountry','Dstcountry/Continent','log_src','log_srcGroup/Flowtype','category','src_packets','dst_packets','payload','src_hex','src_b64','dst_payload','dst_hex','dst_b64','src_agency','dst_agency','src_state','dst_state','tech','ticket','action','note','notes2','os_url','title_url','page_status','internal_ip','internal_hostname','client','severity','resolution'))
        # except TypeError as error:
            # print(error)
            
    # elif SheetFormat == 'anamoli':
        # # write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution)
        # print('%s%s    %s%s    %s    %s%s    %s' %(b,note,g,os_url,r,title_url,o,page_status))
        # try:
            # output_file.write('%s\t%s\t%s\n' %('value','itype','tags'))
        # except TypeError as error:
            # print(error)

    # else:
        # # write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution)
        # print('%s%s    %s%s    %s    %s%s    %s' %(b,note,g,os_url,r,title_url,o,page_status))
        # try:
            # output_file.write('%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n' %('type','value','threat_rating','confidence','source','description','dns','whois'))
        # except TypeError as error:
            # print(error)

    value = sentences(MyDump,output_file,log_src,event_name,url,threat_rating,confidence,first_seen)    
    
def sentences(input_file,output_file,log_src,event_name,url,threat_rating,confidence,first_seen):    
    (UniqRow) = ('')    

    input_file = input_file.replace('\n', ' ')    
    input_file = input_file.replace('\r', ' ')        #TestMe CR
    input_file = input_file.replace('\t', ' ')    
    input_file = input_file.replace(',', ' ')    
    input_file = input_file.replace(';', ' ')    
    input_file = input_file.split(' ')    
    for EachRow in input_file:
        if EachRow.startswith('#')  or EachRow.startswith('\/\/ ')  or EachRow.startswith('; '):    #toss out rem lines
                EachRow = ''
        if EachRow not in UniqRow:
            EachRow = EachRow.strip()
            UniqRow = ('%s\n%s') %(UniqRow,EachRow)
            
    MyList = UniqRow
    MyList = MyList.replace('\n', ' ')        #FixMe group these into one find and replace  maybe | them
    MyList = MyList.replace('\t', ' ')
    MyList = MyList.replace(',', ' ')
    MyList = MyList.replace('\"', ' ')    #fixme
    MyList = MyList.split(' ')    #splits on each space

    try:
        MyList = MyList.strip()
    except:
        pass
    MyList = set(MyList)
    MyList = list(MyList)
    MyList.sort()

    for each_line in MyList:
        if each_line.startswith('b\''):        #test for python3 b'
            try:
                each_line = str(str(each_line.lstrip('b')).strip('\''))        #Fix me
            except:
                pass
            
        (source) = ('')
        # if url != '':
            # source = url
        if log_src != '':
            source = log_src
        # else:
            # source = ''
        
        (log_source,reason,src_ip,src_port,notes2) = ('','','','','')
        (dst_ip,dst_port,src_user,dns_domain,url,Referer) = ('','','','','','')
        (urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry) = ('','','','','','')
        (log_srcGroup,category,src_packets,dst_packets,payload) = ('','','','','')
        (src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency) = ('','','','','','')
        (dst_agency,src_state,dst_state,tech,ticket,action) = ('','','','','','')
        (note,os_url,title_url,page_status,internal_ip,FileExt) = ('','','','','','')
        (internal_hostname,client,severity,resolution) = ('','','','')
        (Tags,Itype) = ('','')
        (length,complexity,priority,revrse) = ('','','','')
        
        (type,value,description) = ('','','')
        (dns,whois,UniqRow,dst_hostname) = ('','','','')
        (Modified,Sort,dst_agency) = ('','','')

        each_line = each_line.rstrip('\.')
        each_line = each_line.rstrip('\?')    
        each_line = each_line.rstrip(':')    #FixMe
        each_line = each_line.rstrip('=')    #FixMe
        each_line = each_line.lstrip('\(')
        each_line = each_line.rstrip('\)')
        each_line = each_line.lstrip('\[')
        each_line = each_line.rstrip('\]')
        
        if each_line.lower().startswith('md5:') :    
            each_line = each_line.replace('MD5:', '')
        elif each_line.startswith('SHA1:') :    
            each_line = each_line.replace('SHA1:', '')
        elif each_line.startswith('SHA256:') :    
            each_line = each_line.replace('SHA256:', '')
        
        if '[.]' in each_line:        #fix obfusacted IP's
            Temp = each_line
            Modified = each_line.replace('[.]', '.')    # replace [.] with .
            each_line = Modified
            Modified = Temp

        if '[:]' in each_line:        #fix obfusacted urls
            each_line = each_line.replace('[:]', ':')
        if '(.)' in each_line:        # todo FixMefix me
            each_line = each_line.replace('(.)', '.')
        if '[at]' in each_line:        #fix obfusacted emails
            each_line = each_line.replace('[at]', '@')
        if '(at)' in each_line:        #fix obfusacted emails
            each_line = each_line.replace('(at)', '@')

        if each_line.startswith('hxxps://') or each_line.startswith('hxxp://'):    #UnRedact Item
            each_line = each_line.replace('hxxp', 'http')


#FileExt finder
        if '.' in each_line:
            FileName, FileExt = os.path.splitext(each_line.lower())

        if each_line.lower().startswith('c:\\') :    #fixme
            if SheetFormat == 'qradar': log_source = 'IOC-FilePath'
            description = 'FilePath'
            type = '_FilePath'
            notes2 = FileExt

        elif each_line.startswith('https://') or each_line.startswith('http://'):    #UnRedact Item
            (description, type, log_source, url) = ('Url', 'Url', 'IOC-url-Full', each_line)
        elif 'https://' in each_line or 'http://' in each_line: 
            (description, type, log_source, url) = ('Url?', 'Url?', 'IOC-url-Full?', each_line)

        elif re.search(r'@', each_line) :
            if bool(re.search(r"^[\w\.\+\-]+\@[\w]+\.[a-z]{2,3}$", each_line)):    #regex email
                (type,log_source,description) = ('Emailaddress','IOC-Email','Email')
            if 'illinois.gov' in each_line.lower():
                log_source = 'IOC-Victim-Email'
                type = '_Email Sender'
                dst_agency = 'SOI'

        elif re.match(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', each_line):        #FixMe
            (description, log_source, type, url) = ('url','IOC-url-Full','url',each_line)
            if 'illinois.gov' in each_line.lower():
                log_source = 'IOC-Victim-Url'
                type = '_Url'
                dst_agency = 'SOI'


        elif  re.match(regex_bitcoin, each_line):   #FixMe
            print(each_line)
            print('bitcoin test')
            (type,log_source,description) = ('File-bitcoin','IOC-Bitcoin','Bitcoin Address')


#IP section
        elif re.match(regex_ipv4, each_line):    #regex ip
            (description, type, log_source) = ('IP', 'address-IPv4', 'IOC-IP')
            each_line = each_line.rstrip('.in-addr.arpa')
            dst_ip  = each_line
            if ":" in each_line:
                dst_ip = each_line.split(':')[0]
                src_ip = dst_ip
                try:
                    dst_port = each_line.split(':')[1]
                    src_port = dst_port
                except:pass

            try:
                (dst_agency, switch_port, publicfacing) = check_agency(each_line)
                if dst_agency != '':
                    source = dst_agency
            except:
                pass

            if dst_agency != '' or dst_ip.startswith('10.') or dst_ip.startswith('163.191.'):
                (description, type) = ('IP', '_address-IPv4')
                log_source = 'IOC-IP-Victim'
                
            if each_line == '127.0.0.1' or each_line == '172.20.10.2' or each_line == '172.30.254.234':    #whitelist localhost
                type = ''
            elif re.search('[a-zA-Z]',each_line):    #change IP's with Alpha characters to hosts
                (description, type, log_source, url, src_ip, dst_ip) = ('Host', 'Host', 'IOC-dns_domain?', each_line, '', '')
                if '.com.' in each_line:
                    description = 'Host_reverse'
                else:
                    description = 'Host'
                    
            elif re.match('^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])(\/([0-9]|[1-2][0-9]|3[0-2]))$', each_line):    #regex CIDR    #works
                (log_source,description,src_ip,dst_ip) = ('IOC-IP-CIDR','IP CIDR block','','')
                

        elif re.match(regex_ipv6, each_line):    #regex ipv6
            (type,log_source) = ('address-IPv6','IOC-IPv6')
            (dst_ip,src_ip) = (each_line,each_line)

#hash section        
        elif len(each_line) == 32:
            if  re.match(regex_md5, each_line):    #regex md5 hash    #Works
                (type,log_source,description) = ('File-MD5','IOC-Hash-MD5','MD5 hash')

        elif len(each_line) == 40:
            if  re.match(regex_sha1, each_line):    #regex SHA1 hash
                (type,log_source,description) = ('File-SHA1','IOC-Hash-SHA1','SHA1 hash')

        elif len(each_line) == 64:
            if  re.match(regex_sha256, each_line):    #regex SHA256 hash
                (type,log_source,description) = ('File-SHA256','IOC-Hash-SHA256','SHA256 hash')
                # Sort = '..'
        elif len(each_line) == 128:
            if  re.match(regex_sha512, each_line):    #regex sha512 hash
                (type,log_source,description) = ('File-SHA512','IOC-Hash-SHA512','SHA512 hash')

#host section
        elif re.search(regex_host, each_line):    #regex url
            if each_line.lower().startswith('http'):
                (description, type, log_source, url) = ('Host', 'Host', 'IOC-url-Full', each_line)
            else:
                (description, type, log_source, dns_domain, url) = ('Host', 'Host', 'IOC-dns_domain', each_line, '')
            if each_line.lower().endswith('.l.google.com') or value == 'www.malwaredomains.com' or value == 'http://freedns.afraid.org' or value == 'abuse.ch': 
                type = '_Host'
                description = 'System-wide exclusion list'

#misc section
        elif  re.match('^\d{3}-\d{2}-\d{4}$', each_line) or re.match('^\d{9}$', each_line):    #regex ssn
            description = 'SSN'
            action = 'SSN'
            
        elif  re.match('^[a-zA-Z]+$', each_line):    #regex ssn
            description = 'Word'
            
        elif  re.match('(\d{3}) \W* (\d{3}) \W* (\d{4}) \W* (\d*)', each_line):    #regex phone    #fixMe
            description = 'Phone number'
            print(('%sPhone Number    %s%s%s') %(b,g,each_line,o))    #temp
            
            if re.search('[A-Za-z]',each_line):        # regex alpha characters
                (description, type) = ('Host?', 'Host')


        elif  re.match('(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})', each_line) and not re.search('[A-Za-z]',each_line):    #regex phone
            description = 'Phone number'

        elif  re.match('^\d+$', each_line):    #regex number
            description = 'Number'
            type = 'Number'


#FileExt section
        elif re.search(r'.exe|.doc|.pdf|.dll|.txt|.vbs|.bat|.bin|.docx|.js|.aspx|.zip|.gzip|.lnk|.ini|.ps1|.pnf', FileExt):
            log_source = 'IOC-File'
            description = 'File'
            type = '_File'
            notes2 = FileExt

        elif each_line.lower().startswith('cve-20') or  each_line.lower().startswith('cve-19'):    #fixme CVE checker
            description = 'CVE'
            type = '_CVE'

        elif each_line.lower().startswith('hklm\\')  or each_line.lower().startswith('hkcu\\'):    #Registry key
            description = 'Registry string'
            type = '_RegistryString'

        # elif  re.match('([0-9a-fA-F]:?){12}', each_line):    #regex MAC address
            # description = 'MAC address'

        (Email,password) = ('','')

        if re.search("[\w\.\+\-]+\@[\w\.\-]+\.[\w\.\-]+", each_line):
            
            Email = re.search("[\w\.\+\-]+\@[\w\.\-]+\.[\w\.\-]+", each_line).group(0)    #JohnMagic .group(0)
            Email = str(Email)
            password = each_line.replace(Email, "")
            password = password.lstrip(':')
            password = password.lstrip('|')

        # if password == '':
            # print(('%s%s%s%s' %(r,each_line,b,o))     #temp
        # else:
            # print (g, Email, g, "password=", b, password, y, "Line=", o, each_line)        #temp

        # value = str(each_line)
        value = each_line
        # print('%soriginal value = %s%s' %(b,value,o))    #temp
        
        try:
            # value = (value.encode('utf8'))    # 'ascii' codec can't decode byte
            # value  = str(value)    #test
            # print('%s is utf8 encoded' %(value))    #temp
            # value = value.decode('utf-8')    #convert bytes to Unicode    #FixMe    # fixes python3 error : 'str' does not support the buffer interface
            # value = value.decode(encoding)    #test
            # print('%sconverted to unicode Temp %s%s' %(g,value,o))    #temp
            pass
        except:
            value = ''
            notes2 = 'ascii codec error'
            print('ascii codec error')    #temp
            pass
            
        severity = threat_rating
        
        note = value
        src_user = Sort



        if event_name != '' and log_source != '':
            description = ('%s %s') %(event_name,description)


        reason = description
        action = type
#revrse
        if 'com.' in value:
            # for o in value.split('.'):
                # print("test", o)
                # revrse.insert(0,o)  # insert value into revrse list
            revrse = value #temp
            # print(value.revrse())
            # print(type(revrse))
            # blah = reverse
            # revrse = value.split('/.')
            # revrse = value.split('.')
            # revrse = revrse.reverse()
            

            
#length
        # if description == 'Word' or description == 'Number' or description == 'Phone number'   :
        if 1==1:
            try:
                # print("%s,%s,%s" %(value,description,len(value)))
                length = str(len(value))
            except TypeError as error: print(error)
            
#convert url to a domain Name
        if url != '':    
            dns_domain = url
            dns_domain = dns_domain.lower()
            dns_domain = dns_domain.replace('http://','')
            dns_domain = dns_domain.replace('https://','')
            dns_domain = dns_domain.split('/')[0]
            if re.match(regex_ipv4, dns_domain):    #regex ip
                dst_ip = dns_domain
                dns_domain = ''
                if ":" in dst_ip:
                    dst_port = dst_ip.split(':')[1]
                    dst_ip = dst_ip.split(':')[0]
            
        severity = threat_rating

        # itype = type    #FixMe
# print section
#Color
        if SheetFormat == "basic":
            if type == 'Emailaddress' or 'Url' in type or 'File' in description or description == 'IP':
                FormatFunction(bg_color = 'orange')
            elif description == 'Phone number':
                FormatFunction(bg_color = 'yellow')
            elif description == 'Host' and value.endswith('.com'):
                FormatFunction(bg_color = 'orange')
            elif description == 'Number':
                if length == '4' or length == '6':
                    FormatFunction(bg_color = 'orange')
            else:
                FormatFunction(bg_color = 'white')        
        elif log_source == '' and reason == '':        #
            FormatFunction(bg_color = 'red')
        elif log_source == '' or note == '':        #
            FormatFunction(bg_color = 'yellow')
        elif 'pastebin.com' in dns_domain or 'paste.cryptoaemus.com' in dns_domain:        #
            FormatFunction(bg_color = 'green')
        else:
            FormatFunction(bg_color = 'white')

#write Qradar xlsx
        # write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution)
        # print('%s%s    %s%s    %s    %s%s    %s' %(b,note,g,os_url,r,title_url,o,page_status))

        if value != '' and SheetFormat == 'qradar':
            write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution)
            
            output_file.write('%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\t\n' %(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution))
        elif value != '' and SheetFormat == 'anamoli':
            write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution)
            output_file.write('%s\t%s\t%s\n'%(value,Itype,Tags))
        elif value != '' and SheetFormat == 'threatconnect':
            write_threatconnect(type,value,threat_rating,confidence,source,description,dns,whois)
            output_file.write('%s\t%s\t%s\t%s\t%s\t%s\t%s\t%s\n'%(type,value,threat_rating,confidence,log_src,description,dns,whois))
        elif value != '':
            write_basic(value, description, length, complexity, type, priority, revrse)
            output_file.write('%s\t%s\t%s\t%s\t%s\t%s\n'%(value, description, length, complexity, type, priority))

    return (value)    #temp

def write_basic(value, description, length, complexity, type, priority, revrse):
    global Row

    sheet1.write_string(Row,0,value,Format)
    sheet1.write_string(Row,1,description,Format)
    sheet1.write_string(Row,2,length,Format)
    sheet1.write_string(Row,3,complexity,Format)
    sheet1.write_string(Row,4,type,Format)
    sheet1.write_string(Row,5,priority,Format)
    sheet1.write_string(Row,6,revrse,Format)
    
    Row += 1

def write_qradar(log_source,event_name,reason,first_seen,src_ip,src_port,dst_ip,dst_port,src_user,dns_domain,url,Referer,urlReputation,src_hostname,dst_hostname,count,Srccountry,Dstcountry,log_src,log_srcGroup,category,src_packets,dst_packets,payload,src_hex,src_b64,dst_payload,dst_hex,dst_b64,src_agency,dst_agency,src_state,dst_state,tech,ticket,action,note,notes2,os_url,title_url,page_status,internal_ip,internal_hostname,client,severity,resolution):
    global Row

    sheet1.write_string(Row,0,log_source,Format)
    sheet1.write_string(Row,1,event_name,Format)
    sheet1.write_string(Row,2,reason,Format)
    sheet1.write_string(Row,3,first_seen,Format)
    sheet1.write_string(Row,4,src_ip,Format)
    sheet1.write_string(Row,5,src_port,Format)
    sheet1.write_string(Row,6,dst_ip,Format)
    sheet1.write_string(Row,7,dst_port,Format)
    sheet1.write_string(Row,8,src_user,Format)
    sheet1.write_string(Row,9,dns_domain,Format)
    sheet1.write_string(Row,10,url,Format)
    sheet1.write_string(Row,11,Referer,Format)
    sheet1.write_string(Row,12,urlReputation,Format)
    sheet1.write_string(Row,13,src_hostname,Format)
    sheet1.write_string(Row,14,dst_hostname,Format)
    sheet1.write_string(Row,15,count,Format)
    sheet1.write_string(Row,16,Srccountry,Format)
    sheet1.write_string(Row,17,Dstcountry,Format)
    sheet1.write_string(Row,18,log_src,Format)
    sheet1.write_string(Row,19,log_srcGroup,Format)
    sheet1.write_string(Row,20,category,Format)
    sheet1.write_string(Row,21,src_packets,Format)
    sheet1.write_string(Row,22,dst_packets,Format)
    try:
        sheet1.write_string(Row,23,payload,Format)
    except TypeError as error: print(error)
    sheet1.write_string(Row,24,src_hex,Format)
    sheet1.write_string(Row,25,src_b64,Format)
    sheet1.write_string(Row,26,dst_payload,Format)
    sheet1.write_string(Row,27,dst_hex,Format)
    sheet1.write_string(Row,28,dst_b64,Format)
    sheet1.write_string(Row,29,src_agency,Format)
    sheet1.write_string(Row,30,dst_agency,Format)
    sheet1.write_string(Row,31,src_state,Format)
    sheet1.write_string(Row,32,dst_state,Format)
    sheet1.write_string(Row,33,tech,Format)
    sheet1.write_string(Row,34,ticket,Format)
    sheet1.write_string(Row,35,action,Format)
    sheet1.write_string(Row,36,note,Format)
    sheet1.write_string(Row,37,notes2,Format)
    sheet1.write_string(Row,38,os_url,Format)
    sheet1.write_string(Row,39,title_url,Format)
    try:
        sheet1.write_string(Row,40,page_status,Format)
    except TypeError as error: print(error)
    sheet1.write_string(Row,41,internal_ip,Format)
    sheet1.write_string(Row,42,internal_hostname,Format)
    sheet1.write_string(Row,43,client,Format)
    sheet1.write_string(Row,44,severity,Format)
    sheet1.write_string(Row,45,resolution,Format)

    Row += 1

def write_threatconnect(type,value,threat_rating,confidence,source,description,dns,whois):
    global Row

    sheet1.write_string(Row,0,type,Format)
    sheet1.write_string(Row,1,value,Format)
    sheet1.write_string(Row,2,threat_rating,Format)
    sheet1.write_string(Row,3,confidence,Format)
    sheet1.write_string(Row,4,source,Format)
    sheet1.write_string(Row,5,description,Format)
    sheet1.write_string(Row,6,dns,Format)
    sheet1.write_string(Row,7,whois,Format)

    Row += 1
    
def usage():                                            # -u will give you some usage examples
    File = sys.argv[0].split('\\')[-1]
    print(y + File +" version: %s by %s" % (version, Author ))
    print("\nExample:")
    print("\t" + File +" -f -I input.txt -O out_text2words.xlsx")    
    print("\t" + File +" -f -I input.pdf -O out_text2words.xlsx")   
    print("\t" + File +" -f -I input.txt -O out_text2words.xlsx -Q")    
    print("\t" + File +" -f -I input.txt -O out_text2words_TC.xlsx -T") 
    print(o)

main()    #GoTo Main Menu

############################   Revision History   #############################

"""
2.2.4 - made a DOR basic output as default
2.2.0 - converted to pep8
2.0.5 - Python2 and 3 compatibility, added write_threatconnect(
2.0.0 - Outputs to Xlsx, fixed url vs dnsomain check, colorized,split ip and port
1.8.0 - PDF input module added (pdf2text)
1.6.9 - python2to3 conversion
0.2.6 - Added input questions
0.2.2 - John added a SSN regex. Kudos John.
0.1.1 - based on password_recheckinator.py
"""


############################   Future Wishlist     ############################

"""

reverse anything with .com. in it.  com.apple.ipay should be ipay.apple.com as the dns
complexity column

Search and fix all lines with #FixMe
-
fix phone numbers with alpha characters (bad regex)
add a hosts regex so hosts and user names don't get mixed together

"""


############################          notes          #############################

"""

10.11.12.13
10.51.34.26
1234
172.217.4.238
217-867-5309
3E9E15AE174E7E16D94270CA06E52838817A9D55
444444444
555-55-5555
A5F31A4B77F4A2169C2AE7ECE08DFCC6
CD17CE11DF9DE507AF025EF46398CFDCB99D3904B2B5718BFF2DC0B01AEAE38C
CVE-2018-1038
HKCU\SOFTWARE\Classes\AppX3
ImAWord
MALICIOUS.EXE
agr.state.il.us
c:\windows\system.exe
google.com
http://34.224.250.219
http://duckduckgo.com
http://www.sameip.org/34.224.250.219
https://www.google.com/robots.txt
paste.cryptoaemus.com
pastebin.com
com.panerabread.mobile.minneapolis
WMWSY9C59ET123456
com.apple.help

"""

############################        Copyright      ############################

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




############################        The End        ############################





