#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<     Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>
try:
    from bs4 import BeautifulSoup
except:
    print(f'install missing modules:    pip install -r requirements_identity_hunt.txt')
    exit()

import os
import re
import sys
import json
import time
from datetime import datetime
import random
import openpyxl
import requests # pip install requests

from docx import Document   # pip install python-docx

from openpyxl import load_workbook, Workbook    # pip install openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

from tkinter import * 
from tkinter import messagebox

import socket
import argparse  # for menu system

# <<<<<<<<<<<<<<<<<<<<<<<<<<     Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description2 = "OSINT: track people down by username, email, ip, phone and website"
tech = 'LincolnLandForensics'  # change this to your name if you are using Linux
version = '3.1.8'

headers_intel = [
    "query", "ranking", "fullname", "url", "email", "user", "phone",
    "business", "fulladdress", "city", "state", "country", "note", "AKA",
    "DOB", "SEX", "info", "misc", "firstname", "middlename", "lastname",
    "associates", "case", "sosfilenumber", "owner", "president", "sosagent",
    "managers", "Time", "Latitude", "Longitude", "Coordinate",
    "original_file", "Source", "Source file information", "Plate", "VIS", "VIN",
    "VYR", "VMA", "LIC", "LIY", "DLN", "DLS", "content", "referer", "osurl",
    "titleurl", "pagestatus", "ip", "dnsdomain", "Tag", "Icon", "Type"
]

headers_locations = [
    "#", "Time", "Latitude", "Longitude", "Address", "Group", "Subgroup"
    , "Description", "Type", "Source", "Deleted", "Tag", "Source file information"
    , "Service Identifier", "Carved", "Name", "business", "number", "street"
    , "city", "county", "state", "zipcode", "country", "fulladdress", "query"
    , "Sighting State", "Plate", "Capture Time", "Capture Network", "Highway Name"
    , "Coordinate", "Capture Location Latitude", "Capture Location Longitude"
    , "Container", "Sighting Location", "Direction", "Time Local", "End time"
    , "Category", "Manually decoded", "Account", "PlusCode", "Time Original", "Timezone"
    , "Icon", "original_file", "case", "Index"
    ]
    

# Regex section
# regex_host = re.compile(r'\b((?:(?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+(?i)(?!exe|php|dll|doc' \
                        # '|docx|txt|rtf|odt|xls|xlsx|ppt|pptx|bin|pcap|ioc|pdf|mdb|asp|html|xml|jpg|gif$|png' \
                        # '|lnk|log|vbs|lco|bat|shell|quit|pdb|vbp|bdoda|bsspx|save|cpl|wav|tmp|close|ico|ini' \
                        # '|sleep|run|dat$|scr|jar|jxr|apt|w32|css|js|xpi|class|apk|rar|zip|hlp|cpp|crl' \
                        # '|cfg|cer|plg|lxdns|cgi|xn$)(?:xn--[a-zA-Z0-9]{2,22}|[a-zA-Z]{2,13}))(?:\s|$)')

# regex_email = re.compile(
    # r'(([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)(\s*;\s*|\s*$))*') # test
# terrible. matches emails and phone numbers

regex_host = re.compile(
    r'(?i)\b((?:(?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+'
    '(?!exe|php|dll|doc|docx|txt|rtf|odt|xls|xlsx|ppt|pptx|bin|pcap|ioc|pdf|mdb|asp|html|xml|jpg|gif$|png'
    '|lnk|log|vbs|lco|bat|shell|quit|pdb|vbp|bdoda|bsspx|save|cpl|wav|tmp|close|ico|ini'
    '|sleep|run|dat$|scr|jar|jxr|apt|w32|css|js|xpi|class|apk|rar|zip|hlp|cpp|crl'
    '|cfg|cer|plg|lxdns|cgi|xn$)'
    '(?:xn--[a-zA-Z0-9]{2,22}|[a-zA-Z]{2,13}))(?:\s|$)')

# regex_url = re.compile(
    # r'(https?:\/\/)?(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()!@:%_\+.~#?&\/\/=]*)') # test
# failed


regex_md5 = re.compile(r'^([a-fA-F\d]{32})$')  # regex_md5        [a-f0-9]{32}$/gm
regex_sha1 = re.compile(r'^([a-fA-F\d]{40})$')  # regex_sha1
regex_sha256 = re.compile(r'^([a-fA-F\d]{64})$')  # regex_sha256
regex_sha512 = re.compile(r'^([a-fA-F\d]{128})$')  # regex_sha512

regex_number = re.compile(r'^(^\d)$')  # regex_number    #Beta
regex_number_fb = re.compile(r'^\d{9,15}$')  # regex_number    #to match facebook user id

# regex_ipv4 = re.compile('(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}' +
                        # '(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)') # works

regex_ipv4 = re.compile('([1-2]?[0-9]?[0-9]\.[1-2]?[0-9]?[0-9]\.[1-2]?[0-9]?[0-9]\.[1-2]?[0-9]?[0-9])') # test


# regex_ipv6 = re.compile('(S*([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}S*|S*(' +
                        # '[0-9a-fA-F]{1,4}:){1,7}:S*|S*([0-9a-fA-F]{1,4}:)' +
                        # '{1,6}:[0-9a-fA-F]{1,4}S*|S*([0-9a-fA-F]{1,4}:)' +
                        # '{1,5}(:[0-9a-fA-F]{1,4}){1,2}S*|S*([0-9a-fA-F]' +
                        # '{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}S*|S*(' +
                        # '[0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}S*' +
                        # '|S*([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4})' +
                        # '{1,5}S*|S*[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4})' +
                        # '{1,6})S*|S*:((:[0-9a-fA-F]{1,4}){1,7}|:)S*|::(ffff' +
                        # '(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}' +
                        # '[0-9]){0,1}[0-9]).){3,3}(25[0-5]|(2[0-4]|1{0,1}[' +
                        # '0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[' +
                        # '0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]).){3,3}(25[' +
                        # '0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))')    # works

regex_ipv6 = re.compile('(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))')    # test




# regex_phone = re.compile(
    # r'^(?:(?:\+?1\s*(?:[.-]\s*)?)?(?:\(\s*([2-9][0-8][0-9])\s*\)|([2-9][0-8][0-9]))\s*(?:[.-]\s*)?)?'
    # r'([2-9][0-9]{2})\s*(?:[.-]\s*)?([0-9]{4})$|^(\d{10})$|^1\d{10}$')  # works

regex_phone = re.compile(r'^(\+\d{1,2}\s)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$')  # test

regex_phone11 = re.compile(r'^1\d{10}$')
regex_phone2 = re.compile(r'(\d{3}) \W* (\d{3}) \W* (\d{4}) \W* (\d*)$')

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


# <<<<<<<<<<<<<<<<<<<<<<<<<<     Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():

    # check internet status
    status = internet()
    status2 = is_running_in_virtual_machine()

    if status2 == True:
        print(f'{color_yellow}This is a virtual machine. Not checking for internet connectivity{color_reset}')
        # apparently when running from a VM (and maybe behind a proxy) it says internet isn't connected
    elif status == False:
        noInternetMsg()
        input(f'{color_red}CONNECT TO THE INTERNET FIRST. Hit Enter to exit...{color_reset}')
        exit()
    else:
        print(color_green + '\nINTERNET IS CONNECTED\n' + color_reset)

    # global section
    global data
    global filename
    filename = 'input.txt'
    global input_xlsx
    input_xlsx = 'Intel.xlsx'
    # global inputDetails
    # inputDetails = 'no'

    global output_xlsx
        
    global row
    row = 1

    global emails
    global ips
    global phones
    global users
    global dnsdomains
    global websites
    global data
    input_file_type = ''
    data = []
    emails = []
    ips = []
    phones = []
    users = []
    dnsdomains = []
    websites = [] 
    
    parser = argparse.ArgumentParser(description=description2)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-c','--convert', help='convert from old headers', required=False, action='store_true')

    parser.add_argument('-E','--emailmodules', help='email modules', required=False, action='store_true')
    parser.add_argument('-b','--blurb', help='write ossint blurb', required=False, action='store_true')
    parser.add_argument('-B','--blank', help='create blank intel sheet', required=False, action='store_true')

    parser.add_argument('-H','--howto', help='help module', required=False, action='store_true')
    parser.add_argument('-i','--ips', help='ip modules', required=False, action='store_true')
    parser.add_argument('-l','--locations', help='convert intel 2 locations format', required=False, action='store_true')
    parser.add_argument('-p','--phonestuff', help='phone modules', required=False, action='store_true')
    parser.add_argument('-s','--samples', help='print sample inputs', required=False, action='store_true')
    parser.add_argument('-t','--test', help='testing individual modules', required=False, action='store_true')
    parser.add_argument('-U','--usersmodules', help='username modules', required=False, action='store_true')
    parser.add_argument('-w','--websitetitle', help='websites titles', required=False, action='store_true')    

    parser.add_argument('-W','--websites', help='websites modules', required=False, action='store_true')    
    

    args = parser.parse_args()
    cls()
    print_logo()

    if args.samples:  
        samples()
        return 0 
        sys.exit()


    if args.howto:  # this section might be redundant
        parser.print_help()
        usage()
        return 0
        sys.exit()



# default input
    if not args.input: 
        input_file_type = 'txt'

# txt input
    elif '.txt' in args.input:
        input_file_type = 'txt'
        filename = args.input
        emails,dnsdomains,ips,users,phones,websites = read_text(filename)

# xlsx input        
    elif '.xlsx' in args.input:
        input_file_type = 'xlsx'
        input_xlsx = args.input
    else:
        input_xlsx = args.input 
        # input_xlsx = args.input
        
        
# output xlsx
    if not args.output:     # openpyxl conversion
        output_xlsx = "Intel__DRAFT_V1.xlsx"          
    else:
        output_xlsx = args.output


    if args.blank:  
        # write_intel_basic(data, output_xlsx)
        write_intel(data)
        return 0 
        sys.exit()



# if text file input
    if input_file_type == 'txt':
        if not os.path.exists(filename):
            input(f"{color_red}{filename} doesnt exist.{color_reset}")
            sys.exit()
        elif os.path.getsize(filename) == 0:
            input(f'{color_red}{filename} is empty. Fill it with username, email, ip, phone and/or websites.{color_reset}')
            sys.exit()
        elif os.path.isfile(filename):
   
            emails,dnsdomains,ips,users,phones,websites = read_text(filename)
            
            inputfile = open(filename)

# if xlsx input
    elif input_file_type == 'xlsx':
        if not os.path.exists(input_xlsx):
            input(f"{color_red}{input_xlsx} doesnt exist.{color_reset}")
            sys.exit()
        elif os.path.getsize(input_xlsx) == 0:
            input(f'{color_red}{input_xlsx} is empty. Fill it with username, email, ip, phone and/or websites.{color_reset}')
            sys.exit()
        elif os.path.isfile(input_xlsx):
            data = read_xlsx(input_xlsx)
            # data = read_xlsx_basic(input_xlsx)
            if args.convert:
                file_exists = os.path.exists(input_xlsx)
                if file_exists == True:
                    
                    print(f' converting {input_xlsx}')    # temp
                    write_intel_basic(data, output_xlsx)
                    
                else:
                    print(f'{color_red}{input_xlsx} does not exist{color_reset}')
                    exit()
                sys.exit()
                
            if args.blurb:
                write_blurb()
                sys.exit()

            if args.locations:
                data = []
                # data = read_xlsx(input_xlsx)    # never finishes
                data = read_xlsx_basic(input_xlsx)    # works
                write_locations(data)   # works
                # write_locations_basic(data, output_xlsx)    # works
                sys.exit()

            data = read_xlsx(input_xlsx)
  
    # Check if no arguments are entered
    if len(sys.argv) == 1:
        print(f"{color_yellow}You didn't select any options so I'll run the major options{color_reset}")
        print(f'{color_yellow}try -h for a listing of all menu options{color_reset}')
        args.emailmodules = True
        args.ips = True
        args.phonestuff = True
        args.usersmodules = True
        args.websites = True
   

        
        
    if args.emailmodules and len(emails) > 0:  
        print(f'Emails = {emails}') # temp
        main_email()    # 
        carrot_email()
        cyberbackground_email() # beta
        epios_email()   # beta
        # digitalfootprintcheckemail()
        ghunt()  # this is overwriting data
        google_calendar()     #
        have_i_been_pwned()    #
        holehe_email()    #
        osintIndustries_email()    #
        thatsthememail()    #
        truepeople_email()
        ## twitteremail()    # auth required    
        veraxity()
        wordpresssearchemail()  # requires auth
        
    if args.ips and len(ips) > 0:  
        print(f'IPs = {ips}')
        arinip()    # alpha
        main_ip()
        ## geoiptool() # works but need need to rate limit; expired certificate breaks this
        resolverRS()    #? 
        # thatsthemip() # broken
        whoisip()    #
        whatismyip()    #
        
    ### phone modules
    if args.phonestuff and len(phones) > 0:
        print(f'phones = {phones}')
        main_phone()
        familytreephone()    #
        thatsthemphone()   #
        reversephonecheck()    #
        # spydialer()    #
        validnumber()    #
        whitepagesphone()    #
        # whocalld()    # works
        
    if args.test:  
        print(f' using test module')
        ham_radio()
        

    if args.usersmodules and len(users) > 0:  
        print(f'users = {users}')    
        main_user()
        about()    #
        bitbucket()    #
        blogspot_users()    # test
        bsky()
        cashapp()
        disqus()    # test
        # ebay()  # all false positives due to captcha
        etsy()  # task
        facebook()     #
        familytree()    #
        # fiverr()    # test
        flickr()   # add photo, note, name, info
        freelancer()    #
        # friendfinder()  # add (fullname, city, country, note, DOB, SEX)
        foursquare()    #
        garmin()    #
        github()    #
        gravatar()  # grab bonus urls
        ham_radio() # manual but works
        imageshack()    #
        instagram()    #
        instantusername() # test
        instructables()     #
        inteltechniques()   #test
        keybase()   # add location and twitter, and photo
        kik()    #
        linkedin()    # cookie test
        # massageanywhere()   # broken ssl query
        mastadon()    #
        myfitnesspal()    #
        myshopify()    #
        myspace_users()    #
        paypal()  # needs work
        patreon()    #
        pinterest()
        poshmark()    #    
        public()    #
        reddit()
        # rumble()  # test
        roblox()
        sherlock()    #
        # signal()    # test
        # slack() # blows up on some users first.last#
        snapchat()    # must manually verify
        spotify()    #
        ## telegram()# crashes the script
        threads()    #
        tiktok()    #
        tinder() # add DOB, schools
        truthSocial()   # false positives
        twitter()   # needs auth
        venvmo()
        # vimeo()   # test
        whatnot()
        whatsmyname()    #
        wordpress()    #
        wordpress_profiles()    #  
        youtube()    #
        

    if args.websitetitle and len(websites) > 0:  
        print(f'websites = {websites}')   
        main_website()
        # titles()    # alpha
        
    if args.websites and len(websites) > 0:  
        print(f'websites = {websites}')    
        centralops()
        main_website()
        # redirect_detect() # timed out and crashed
        robtex()
        # titles()    # alpha
        viewdnsdomain()
        whoiswebsite()

    write_intel(data)



    # set linux ownership    
    if sys.platform == 'win32' or sys.platform == 'win64':
        pass
    else:
        call(["chown %s.%s *.xlsx" % (tech.lower(), tech.lower())], shell=True)

    # workbook.close()
    if not args.blurb:
        input(f"See '{input_xlsx}' for output. Hit Enter to exit...")

    return 0
    
    # sys.exit()  # this code is unreachable


# <<<<<<<<<<<<<<<<<<<<<<<<<<  Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>


def about(): # testuser = kevinrose
    """
        parses each user
        creates a url based on the username
        if the webpage exists,it writes it to the output sheet
    """
    
    print(f'{color_yellow}\n\t<<<<< about.me {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - about.me')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (note, city, country, fullname, titleurl, pagestatus) = ('', '', '', '', '', '')
        (lastname, firstname) = ('', '')
        user = user.rstrip()
        url = (f'https://about.me/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        for eachline in content.split("\n"):
            if " | about.me" in eachline:        
                fullname = eachline.strip().replace(" | about.me","") # .split("-")(0)

                if ' - ' in fullname:
                    note = fullname.split(' - ')[1]
                    fullname = fullname.split(' - ')[0]  
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                fullname = ''




        if '404' not in pagestatus:
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["note"] = note
            row_data["user"] = user
            row_data["city"] = city
            row_data["country"] = country
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname

            data.append(row_data)

def arinip():    # testuser=    77.15.67.232
    from subprocess import call, Popen, PIPE
    print(f"{color_yellow}\n\t<<<<< arin {color_blue}IP's{color_yellow} >>>>>{color_reset}")
    for ip in ips:    
        row_data = {}
        (query, note) = (ip, '')
        (city, business, country, zipcode, state) = ('', '', '', '', '')
        (Latitude, Longitude) = ('', '')

        (content, titleurl, pagestatus) = ('', '', '')
        (email, phone, fullname, entity, fulladdress) = ('', '', '', '', '') 
        url = (f'https://search.arin.net/rdap/?query={ip}')        
        
        if sys.platform == 'win32' or sys.platform == 'win64':    
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
            
            
            if '403 Forbidden' in content:
                pagestatus = '403 Forbidden'
                content = ''
            for eachline in content.split("\n"):
                
                if "www.ip-adress.com/legal-notice" in eachline:
                    for eachline in content.split("<tr><th>"):
                        if "Country<td>" in eachline:
                            country = eachline.strip().split("Country<td>")[1]
                        elif "City<td>" in eachline:
                            city = eachline.strip().split("City<td>")[1]
                        elif "ISP</abbr><td>" in eachline:
                            business = eachline.strip().split("ISP</abbr><td>")[1]
                        elif "Postal Code<td>" in eachline:
                            zipcode = eachline.strip().split("Postal Code<td>")[1]
                        elif "State<td>" in eachline:
                            state = eachline.strip().split("State<td>")[1]
               
                elif "<tr><th>Country</th><td>" in eachline:
                    country = eachline.strip().split("<tr><th>Country</th><td>")[1]
                    country = country.split("<")[0]
                elif "City: " in eachline:
                    city = eachline.strip().split("<tr><th>City</th><td>")[1]
                    city = city.split("<")[0]            
                elif "<tr><th>Postal Code</th><td>" in eachline:
                    zipcode = eachline.strip().split("<tr><th>Postal Code</th><td>")[1]
                    zipcode = zipcode.split("<")[0] 
            # time.sleep(3) #will sleep for 30 seconds
            if "Fail" in pagestatus:
                pagestatus = 'fail'

        print(f'{color_green}{ip}{color_yellow}	{country}	{city}	{zipcode}{color_reset}')

        if '.' in ip:
            ranking = '9 - ARIN'
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["ip"] = ip
            row_data["note"] = note
            row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["email"] = email
            row_data["phone"] = phone
            row_data["note"] = entity
            row_data["fulladdress"] = fulladdress
            row_data["query"] = query
            
            row_data["city"] = city
            row_data["country"] = country
            row_data["state"] = state
            row_data["zipcode"] = zipcode            
            row_data["Latitude"] = Latitude
            row_data["Longitude"] = Longitude
          
           
            data.append(row_data)


def bitbucket(): # testuser = rick
    print(f'{color_yellow}\n\t<<<<< bitbucket {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '4 - bitbucket')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (city, country, fullname, titleurl, pagestatus) = ('', '', '', '', '')
        user = user.rstrip()
        url = (f'https://bitbucket.org/{user}/')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("\n"):
            if "display_name" in eachline:
                pattern = r'"display_name":\s*"([^"]+)"'
                match = re.search(pattern, eachline)

                if match:
                    fullname = match.group(1)

        if ' ' in fullname:
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
        else:
            fullname = ''

        if '404' not in pagestatus:
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname            
            row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["user"] = user
            row_data["city"] = city
            row_data["country"] = country

            data.append(row_data)

def blogspot_users(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< blogspot {color_blue}users{color_yellow} >>>>>{color_reset}')
    
    for user in users:
        row_data = {}
        (query, ranking) = (user, '4 - blogspot')
    
        url = f"https://{user}.blogspot.com"

        (content, referer, osurl, titleurl, pagestatus) = request_url(url)
        (fullname, firstname, lastname, middlename) = ('', '', '', '')

        if 'Success' in pagestatus:
            titleurl = titleurl_og(content)
            fullname = titleurl

            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                fullname = ''

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname 
            row_data["url"] = url
            # row_data["titleurl"] = titleurl
            row_data["user"] = user
            
            data.append(row_data)


def bsky(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< bsky {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '5 - bsky')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.strip()
        url = (f'https://bsky.app/profile/{user}.bsky.social')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1]
                fullname = fullname.split(" (")[0]
            elif "og:description" in eachline:
                note = eachline.strip().split("\"")[1]
                print(f'note = {note}')

        if '@' in titleurl:
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user  
            row_data["note"] = note 

            data.append(row_data)  


def carrot_email(): 
    print(f'{color_yellow}\n\t<<<<< carrot2 {color_blue}emails{color_yellow} >>>>>{color_reset}')
    
    for email in emails:
        row_data = {}
        (query, content, note) = (email, '', '')
        url = f'https://search.carrot2.org/#/search/web/{email}/folders'
        
        # Uncomment this line when request() function is implemented
        # (content, referer, osurl, titleurl, pagestatus) = request(url)

        if '@' in email.lower():
            if 'Enable JavaScript to run this app' in content:
                ranking = '9 - carrot'
            elif 'All retrieved results' in content:
                note = 'Events shown in time zone'
                ranking = '4 - carrot2'
                # print(f'{color_green} {email}{color_reset}')                
            else:
                ranking = '9 - carrot2'
                # print(f'{color_yellow} {email}  {color_reset}')

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["email"] = email
            row_data["note"] = note
            data.append(row_data)
            # print(f'row_data = {row_data}') # temp

def cashapp(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< cash.app {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - cashapp')
        (fullname, firstname, lastname, middlename, country) = ('', '', '', '', '')
        user = user.rstrip()
        url = (f'https://cash.app/${user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        if '404' not in pagestatus:
            pattern = r'"display_name":"([^"]+)"'
            match = re.search(pattern, content)
            if match:
                fullname = match.group(1)

            if fullname.count(' ') == 1:
                firstname, lastname = fullname.split(' ')
                firstname = firstname.title()
                lastname = lastname.upper()
                fullname = (f'{firstname} {lastname}')
                
            elif fullname.count(' ') == 2:
                firstname, middlename, lastname = fullname.split(' ')
                firstname = firstname.title()
                lastname = lastname.upper()
                middlename = middlename.title()
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            
            
            pattern2 = r'"country_code":"([^"]+)"'
            match2 = re.search(pattern2, content)
            if match:
                country = match2.group(1)        
            
            print(f'{color_green}{url}\t{fullname}\t{country}{color_yellow}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname
            row_data["country"] = country
                      
            data.append(row_data)
     

def centralops(): 
    if len(websites) > 0:
        row_data = {}
        ranking = '9 - manual'
        url = ('https://centralops.net/co/')

        row_data["ranking"] = ranking
        row_data["url"] = url
        data.append(row_data)    

def cls():
    """
        Clears the screen    
    """
    
    linux = 'clear'
    windows = 'cls'
    os.system([linux, windows][os.name == 'nt'])

def convert_timestamp(timestamp, time_orig, timezone):
    timezone = timezone or ''
    time_orig = time_orig or ''

    timestamp = str(timestamp)

    if re.match(r'\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2}:\d{2}\.\d{3} (AM|PM)', timestamp):
        # Define the expected format
        expected_format = "%m/%d/%Y %I:%M:%S.%f %p"

        # Parse the string into a datetime object
        dt_obj = datetime.strptime(timestamp, expected_format)

        # Remove microseconds
        dt_obj = dt_obj.replace(microsecond=0)

        # Format the datetime object back into a string with the specified format
        timestamp = dt_obj.strftime("%Y/%m/%d %I:%M:%S %p")


    # Regular expression to find all timezones
    timezone_pattern = r"([A-Za-z ]+)$"
    matches = re.findall(timezone_pattern, timestamp)

    # Extract the last timezone match
    if matches:
        timezone = matches[-1]
        timestamp = timestamp.replace(timezone, "").strip()
    else:
        timezone = ''
        
    if time_orig == "":
        time_orig = timestamp
    else:
        timezone = ''

    # timestamp = timestamp.replace(' at ', ' ')
    if "(" in timestamp:
        timestamp = timestamp.split('(')
        timezone = timestamp[1].replace(")", '')
        timestamp = timestamp[0]
    elif " CDT" in timestamp:
        timezone = "CDT"
        timestamp = timestamp.replace(" CDT", "")
    elif " CST" in timestamp:
        timezone = "CST"
        timestamp = timestamp.replace(" CST", "")

    formats = [
        "%B %d, %Y, %I:%M:%S %p %Z",    # June 13, 2022, 9:41:33 PM CDT (Flock)
        "%Y:%m:%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M:%S.%f %p", # '3/8/2024 11:06:47.358 AM  # task
        "%m/%d/%Y %I:%M %p",  # timestamps without seconds
        "%m/%d/%Y %H:%M:%S",  # timestamps in military time without seconds
        "%m-%d-%y at %I:%M:%S %p %Z", # test 09-10-23 at 4:29:12 PM CDT
        "%m-%d-%y %I:%M:%S %p",
        "%B %d, %Y at %I:%M:%S %p %Z",
        "%B %d, %Y at %I:%M:%S %p",
        "%B %d, %Y %I:%M:%S %p %Z",
        "%B %d, %Y %I:%M:%S %p",
        "%B %d, %Y, %I:%M:%S %p %Z",
        "%Y-%m-%dT%H:%M:%SZ",  # ISO 8601 format with UTC timezone
        "%Y/%m/%d %H:%M:%S",  # 2022/06/13 21:41:33
        "%d-%m-%Y %I:%M:%S %p",  # 13-06-2022 9:41:33 PM
        "%d/%m/%Y %H:%M:%S",  # 13/06/2022 21:41:33
        "%Y-%m-%d %I:%M:%S %p",  # 2022-06-13 9:41:33 PM
        "%Y%m%d%H%M%S",  # 20220613214133
        "%Y%m%d %H%M%S",  # 20220613 214133
        "%m/%d/%y %H:%M:%S",  # 06/13/22 21:41:33
        "%d-%b-%Y %I:%M:%S %p",  # 13-Jun-2022 9:41:33 PM
        "%d/%b/%Y %H:%M:%S",  # 13/Jun/2022 21:41:33
        "%Y/%b/%d %I:%M:%S %p",  # 2022/Jun/13 9:41:33 PM
        "%d %b %Y %H:%M:%S",  # 13 Jun 2022 21:41:33
        "%A, %B %d, %Y %I:%M:%S %p %Z",  # Monday, June 13, 2022 9:41:33 PM CDT ?
        "%A, %B %d, %Y %I:%M:%S %p"     # Monday, June 13, 2022 9:41:33 PM CDT
    ]

    for fmt in formats:
        try:
            dt_obj = datetime.strptime(timestamp.strip(), fmt)
            timestamp = dt_obj
            return timestamp, time_orig, timezone
                        
        except ValueError:
            pass

    # raise ValueError(f"{time_orig} Timestamp format not recognized")
    return timestamp, time_orig, timezone
    
def cyberbackground_email():# testEmail= kevinrose@gmail.com    
    print(f'{color_yellow}\n\t<<<<< cyberbackground {color_blue}emails{color_yellow} >>>>>{color_reset}')
    
    for email in emails:
        row_data = {}
        (query, content, note) = (email, '', '')
        url = (f'https://www.cyberbackgroundchecks.com/email/{email}')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)        

        if 1==1:
            if ('results for') in content: 
                note = 'results for'
                ranking = '8 - cyberbackground'

                print(f'{color_green} {email}   {url}{color_reset}')
          
            else:
            
                ranking = '9 - cyberbackground'

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["email"] = email
            row_data["note"] = note
            row_data["content"] = content
                        
            data.append(row_data)

            
def digitalfootprintcheckemail():    # testuser=    kevinrose@gmail.com 
    print(f'{color_yellow}\n\t<<<<< digitalfootprintcheck {color_blue}emails{color_yellow} >>>>>{color_reset}')    
    
    for email in emails:
        row_data = {}
        (query, content, note) = (email, '', '')
        
        url = (f'https://www.digitalfootprintcheck.com/free-checker.html?q={email}')

        ranking = '9 - digitalfootprintcheck'
                        
        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["url"] = url
        row_data["email"] = email                        
        data.append(row_data)

def disqus(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< discus {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '5 - discus')
        
        (city, country, fullname, titleurl, pagestatus) = ('', '', '', '', '')
        user = user.rstrip()
        # url = (f'http://disqus.com/{user}')
        # url = (f'http://disqus.com/by/{user}')  
        url = (f'http://disqus.com/by/{user}/about/')  

        
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        
        if '404' not in pagestatus:
            print(f'{color_green}{url}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["user"] = user

            data.append(row_data)

def ebay(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< ebay {color_blue}users{color_yellow} >>>>>{color_reset}')
    print(f'{color_yellow}\n\tthis can take a while >>>>>{color_reset}')

    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - ebay')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (city, country, fullname, titleurl, pagestatus, note) = ('', '', '', '', '', '')
        user = user.rstrip()
        url = (f'https://www.ebay.com/str/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        if "200" in pagestatus:
            titleurl = titleurl.replace(' | eBay Stores', '')
            fullname = titleurl
            
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                fullname = ''          
            
            
            print(f'{color_green}{url}{color_reset}') 
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["user"] = user
            row_data["note"] = note            
           

            data.append(row_data)

        time.sleep(5) #will sleep for 5 seconds


def epios_email():# testEmail= kevinrose@gmail.com    
    print(f'{color_yellow}\n\t<<<<< epios {color_blue}emails{color_yellow} >>>>>{color_reset}')
    
    for email in emails:
        row_data = {}
        (query, content, note) = (email, '', '')
        url = (f'https://epieos.com/?q={email}&t=email')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)        

        if 1==1:
            if ('results for') in content: 
                note = 'results for'
                ranking = '8 - epios'

                print(f'{color_green} {email}   {url}{color_reset}')
          
            else:
            
                ranking = '9 - epios'

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["email"] = email
            row_data["note"] = note
            # row_data["content"] = content
                        
            data.append(row_data)
            
            
def etsy(): # testuser = kevinrose https://www.etsy.com/people/kevinrose
    print(f'{color_yellow}\n\t<<<<< etsy {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '6 - etsy')
        (firstname, lastname) = ('', '')
        (city, country, fullname, titleurl, pagestatus) = ('', '', '', '', '')
        user = user.rstrip()
        url = (f'https://www.etsy.com/people/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        if '404' not in pagestatus:
            # grab display_name = fullname
            titleurl = titleurl.replace("'s favorite items - Etsy",'')

            fullname = titleurl
            
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                fullname = ''

            if fullname == 'etsy.com':
                ranking = '9 - etsy'
                fullname = ''

            if ranking == '4 - etsy':
                print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["fullname"] = fullname
                row_data["firstname"] = firstname            
                row_data["lastname"] = lastname
                row_data["url"] = url
                row_data["user"] = user
                row_data["city"] = city            
         

                data.append(row_data)

        time.sleep(5) #will sleep for 5 seconds

def facebook(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< facebook {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '4 - Facebook')

        (fullname,lastname,firstname) = ('','','')
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        user = user.rstrip()
        url = (f'https://facebook.com/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        fullname = titleurl.strip()
        if ' ' in fullname:
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)        
        else:
            fullname = ''

        if 'This content isn' not in content:
            print(f'{color_green}{user}{color_yellow}	{fullname}{color_reset}') 
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname            
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user

            data.append(row_data)        


def familytree(): 
    print(f'\n\t{color_yellow}<<<<< Manually check familytreenow.com >>>>>{color_reset}')
    row_data = {}
    (query, ranking, note) = ('', '9 - manual', 'See FamilyTree link for Possible Relatives and Possible Associates')
    url = ('https://www.familytreenow.com/search/')

    # row_data["query"] = query
    row_data["ranking"] = ranking
    row_data["url"] = url
    row_data["note"] = note
    data.append(row_data)


def familytreephone():# DROP THE LEADING 1
    print(f'{color_yellow}\n\t<<<<< familytree {color_blue}phone numbers{color_yellow} >>>>>{color_reset}')
    for phone in phones:
        row_data = {}
        (query, ranking) = (phone, '8 - familytree')

        (country, city, state, zipcode, case, note) = ('', '', '', '', '', 'See FamilyTree link for Possible Relatives and Possible Associates')
        (fullname, content, referer, osurl, titleurl, pagestatus)  = ('', '', '', '', '', '')

        url = (f'https://www.familytreenow.com/search/genealogy/results?phoneno=%s' %(phone.lstrip('1')))
        # url = (f'https://www.familytreenow.com/search/genealogy/results?phoneno={phone.lstrip('1')}')

        state = phone_state_check(phone, state).replace('?', '')
        
        if 1==1:        
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["phone"] = phone
            row_data["state"] = state
            row_data["note"] = note
            data.append(row_data)

def findwhocallsyou():# testPhone= 
    # note: need to add dishes if there are none
    print(f'{color_yellow}\n\t<<<<< findwhocallsyou {color_blue}phone numbers{color_yellow} >>>>>{color_reset}')
    for phone in phones:
        phone = phone.replace('-', '')
 
        row_data = {}
        (query, ranking) = (phone, '9 - findwhocallsyou')
  
        (country, city, zipcode, case, note) = ('', '', '', '', '')

        url = (f'https://findwhocallsyou.com/{phone}')    

        if 1==1:
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["phone"] = phone
            row_data["note"] = note

            data.append(row_data)

def fiverr():    # testuser=    kevinrose
    print(f'{color_yellow}\n\t<<<<< fiverr {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - fiverr')
        (fullname, firstname, lastname, middlename, note, DOB)  = ('','','','', '', '')
        (misc) = ('')

        url = (f'https://www.fiverr.com/{user}')
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')

        try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
            content = content.strip()
            titleurl = titleurl.strip()

        except:
            pass
            
        # time.sleep(1) # will sleep for 1 seconds
        if 1==1:
        # if 'alternate' in content:
            print(f'{color_green}{url}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname
            row_data["note"] = note


            data.append(row_data)


def flickr(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< flickr {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '4 - flickr')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.rstrip()
        url = (f'https://www.flickr.com/people/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1]

        if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            
            data.append(row_data)           

def foursquare():    # testuser=    john
    print(f'{color_yellow}\n\t<<<<< foursquare {color_blue}users>>>>>{color_reset}')    

    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - foursquare')
        (fullname, note, firstname, lastname, middlename) = ('', '', '', '', '')
        url = (f'https://foursquare.com/{user}')
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)

            content = content.strip()
            titleurl = titleurl.strip()
       
        except:
            pass

        if ' on Foursquare' in titleurl:
            pattern = r'<meta\s+content=["\'](.*?)["\']\s+property="og:description"'
            match = re.search(pattern, content)

            if match:
                note = match.group(1)
                if '"' in note:
                    note = note.split('"')[1]
                # print(note)

            fullname = titleurl.rstrip(' on Foursquare')

            if '' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
        
            print(f'{color_green}{url}{color_yellow} {fullname}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            row_data["note"] = note 
            # row_data["titleurl"] = titleurl            
            # row_data["content"] = content                          
            data.append(row_data)  

            

def freelancer(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< freelancer {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '5 - freelancer')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        user = user.rstrip()
        url = (f'https://www.freelancer.com/u/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        titleurl = titleurl.replace(' Profile | Freelancer','')
        if '404' not in pagestatus:
           
            if ' ' in titleurl:
                fullname = titleurl
            
            if fullname.lower() == user.lower():
                fullname = ''

            if '' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if 'Browser ' not in fullname:
                print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 
                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["fullname"] = fullname
                row_data["firstname"] = firstname
                row_data["lastname"] = lastname
                row_data["url"] = url
                row_data["user"] = user            
                # row_data["titleurl"] = titleurl            
                            
                data.append(row_data)  

def friendfinder():    # testuser=  kevinrose
    print(f'{color_yellow}\n\t<<<<< friendfinder {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - friendfinder')
        (fullname, city, country, note, DOB, SEX) = ('', '', '', '', '', '')
        (firstname, middlename, lastname) = ('', '', '')
        url = (f'https://www.friendfinder-x.com/profile/{user}')

        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        # (fullname, info, note) = ('', '', '')
        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
      
        except:
            pass

        if 'Register to Find' not in titleurl:
            print(f'{color_green}{url}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user 
            row_data["city"] = city            
            row_data["country"] = country            
            row_data["note"] = note            
            row_data["DOB"] = DOB            
            row_data["SEX"] = SEX            
            # row_data["content"] = content            


            data.append(row_data) 


def fullname_parse(fullname):
    (firstname, lastname, middlename) = ('', '', '')
    fullname = fullname.strip()
    if ' ' in fullname:
        fullname_parts = fullname.split(' ')
        firstname = fullname_parts[0].title()
        lastname = fullname_parts[-1].upper()
        
        if len(fullname_parts) == 2:
            fullname = (f'{firstname} {lastname}')
        elif len(fullname_parts) == 3:
            middlename = fullname_parts[1].title()
            fullname = (f'{firstname} {middlename} {lastname}')

    return fullname, firstname, middlename, lastname


def garmin(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< garmin {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - garmin')
        # (fullname, lastname, firstname, case, SEX) = ('','','','','')
        user = user.rstrip()

        url = (f'https://connect.garmin.com/modern/profile/{user}')

        (content, referer, osurl, titleurl, pagestatus) = request(url)

        if 'twitter:card' not in content:
        
            fullname = titleurl
            fullname = fullname.split(" (")[0]
            fullname = fullname.replace("Garmin Connect","").strip()

            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            # row_data["titleurl"] = titleurl            
            # row_data["city"] = city            
            # row_data["country"] = country            
            # row_data["note"] = note            
            # row_data["DOB"] = DOB            
            # row_data["SEX"] = SEX            
            # row_data["content"] = content            

            data.append(row_data) 

def ghunt():    # testEmail= kevinrose@gmail.com
    for email in emails:
        row_data = {}
        (query, note) = (email, '')
        note = (f'cd C:\Forensics\scripts\python\git-repo\GHunt && ghunt email {email}')

        if '@' in email.lower():
            if email.endswith('gmail.com'):
                ranking = ('9 - ghunt')
                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["email"] = email
                row_data["note"] = note
                data.append(row_data)


def github(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< github {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - github')
        (city, country, fullname, titleurl, pagestatus, content, info) = ('', '', '', '', '', '', '')
        (note) = ('')
        user = user.strip()
        url = (f'https://github.com/{user}')

        (content, referer, osurl, titleurl, pagestatus) = request(url)
        if "(" in titleurl:
            fullname = titleurl.strip()
            if "(" in fullname:
                fullname = fullname.split("(")[1]
                fullname = fullname.strip().split(")")[0]

        for eachline in content.split("  <"):

                
            if "og:description" in eachline:
                note = eachline.strip()
                note = note.strip().split("\"")[1]
            elif "twitter:title" in eachline:
                info = eachline.strip()
                info = info.strip().split("\"")[1].replace(' - Overview', '')
                info = (f'https://x.com/{info}')
  
        if 'GitHub is where ' in note:
            note = ''

        if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["middlename"] = middlename            
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user 
            row_data["note"] = note             
            row_data["info"] = info
            data.append(row_data)  
     

def gravatar(): # testuser = kevinrose      https://en.gravatar.com/kevinrose
    print(f'{color_yellow}\n\t<<<<< gravatar {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - gravatar')
        (city, country, fullname, titleurl, pagestatus, info) = ('', '', '', '', '', '')
        (info, lastname, firstname, note, otherurls, misc) = ('', '','', '', '', '')
        user = user.rstrip()
        url = (f' https://gravatar.com/{user}.json')        

        if any(char.isalpha() for char in user):
            (content, referer, osurl, titleurl, pagestatus) = request(url)

            (parsed_data) = []
            if 'ail' not in pagestatus:
                try:            
                    parsed_data = json.loads(content)
                except TypeError as error:
                    print(f'{color_red}{error}{color_reset}')

                
                info = parsed_data['entry'][0]['photos'][0]['value']


                if 'familyName' in content:
                    fullname = parsed_data['entry'][0]['name']['formatted']

                if ' ' in fullname:
                    (fullname, firstname, lastname, middlename) = fullname_parse(fullname)


                if 'aboutMe' in content:
                    note = parsed_data['entry'][0]['aboutMe']
                    note = note.replace('&amp', '&')

                url = (f'http://en.gravatar.com/{user}')
                print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 
                
                if fullname != '' or otherurls != '' or note != '': 
                    ranking = '3 - gravatar'

                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["fullname"] = fullname
                row_data["firstname"] = firstname
                row_data["lastname"] = lastname
                row_data["url"] = url
                row_data["user"] = user            
                row_data["info"] = info 
                row_data["misc"] = misc                 
                row_data["note"] = note 
                row_data["titleurl"] = titleurl            
                row_data["city"] = city            
                row_data["country"] = country            
                row_data["note"] = note            
                # row_data["content"] = content            
                data.append(row_data)

def google_calendar():# testEmail= kevinrose@gmail.com    
    print(f'{color_yellow}\n\t<<<<< google calendar {color_blue}emails{color_yellow} >>>>>{color_reset}')
    
    for email in emails:
        row_data = {}
        (query, content, note) = (email, '', '')
        url = (f'https://calendar.google.com/calendar/u/0/embed?src={email}&pli=1')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)        

        if 'gmail.com' in email.lower():
            if ('you do not have the permission to view') in content: 
                note = 'you do not have the permission to view'
                ranking = '8 - calendar'

                print(f'{color_green} {email}   {url}{color_reset}')
            elif ('Events shown in time zone') in content:
                note = 'Events shown in time zone'
                ranking = '4 - calendar'             
            else:
                ranking = '9 - calendar'

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["email"] = email
            row_data["note"] = note
            data.append(row_data)
            # print(f'row_data = {row_data}') # temp


def ham_radio(): # testuser = K9CYC
    '''
    make a list of all users that might also be FCC Ham radio call signs.
    Ham radio call signs can lead to verified full names, address, etc.
    '''
    print(f'{color_yellow}\n\t<<<<< ham radio {color_blue}users{color_yellow} >>>>>{color_reset}')
    callsigns = []
    call_sign_pattern = r'([BFGKIMNRW]|[0-9][A-Z]|[A-Z][0-9]|[A-Z][A-Z])[0-9][0-9A-Z]{0,2}[A-Z]?'

    for user in users:    
        row_data = {}
        (query, ranking, info) = (user, '5 - ham radio', '')
        (note) = ('https://www.arrl.org/advanced-call-sign-search')
        user = user.rstrip().upper()
        url = (f'https://www.radioqth.net/lookup')        

        if len(user) <= 6 and re.match(call_sign_pattern, user):
            callsigns.append(user)
            
            
    info = f'{", ".join(callsigns)}'  # Join all call signs with commas and append to info string
    
    if callsigns:
            print(f'{color_green}{url}{color_yellow}	{info}{color_reset}') 
            row_data["query"] = info
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["note"] = note
            data.append(row_data)


def have_i_been_pwned(): 
    if len(emails) > 0:
        row_data = {}
        ranking = '8 - manual'
        url = ('https://haveibeenpwned.com')

        row_data["ranking"] = ranking
        row_data["url"] = url
        data.append(row_data)


def holehe_email(): # testEmail= kevinrose@gmail.com
    print(f'{color_yellow}\n\t<<<<< holehe {color_blue}emails{color_yellow} >>>>>{color_reset}')    # temp
    for email in emails:
        row_data = {}
        (query, ranking) = (email, '9 - manual')
        note = (f'cd C:\Forensics\scripts\python\git-repo\holehe && holehe -NP --no-color --no-clear --only-used {email}')

        if '@' in email.lower():
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["email"] = email
            row_data["note"] = note
            data.append(row_data)

def imageshack(): # testuser = ToddGilbert

    print(f'{color_yellow}\n\t<<<<< imageshack {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '4 - imageshack')
        user = user.rstrip()
        url = (f'https://imageshack.com/user/{user}')

        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
        except:
            pass

        if 's Images' in titleurl:
            # fullname = titleurl
            print(f'{color_green}{url}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            data.append(row_data)
            
def instagram():    # testuser=    kevinrose     # add info
    print(f'{color_yellow}\n\t<<<<< instagram {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - instagram')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        url = (f'https://instagram.com/{user}/')
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        (fullname, info, note) = ('', '', '')
        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)

            # Regular expression to find the profile_id
            pattern = r'"profile_id":"(\d+)"'

            content = content.strip()
            titleurl = titleurl.strip()
            for eachline in content.split("\n"):
                if "@context" in eachline:
                    content = eachline.strip()
                elif 'og:title' in eachline and 'content=\"' in eachline:
                    fullname = eachline.split('\"')[1].split(' (')[0]
                elif 'ProfilePage\",\"description' in eachline:
                    info = eachline
                    # Load the JSON data
                    datatemp = json.loads(eachline)

                    # Extract the description value and print it
                    note = datatemp['description']
                elif 'profile_id' in eachline:

                    # Search for the pattern in the content
                    match = re.search(pattern, content)

                    # Extract and print the profile_id if found
                    if match:
                        misc = match.group(1)   # user id
                    else:
                        misc = ''
        except:
            pass

        if '@' in titleurl:
            fullname = titleurl.split(" (")[0]
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                ranking = '3 - instagram'
            else:
                fullname = ''
            if "@" in fullname:
                (fullname, firstname, lastname, middlename) = ('', '', '', '')
        if '@' in titleurl:
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset} {misc}')   
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            # row_data["middlename"] = middlename
            row_data["lastname"] = lastname
            row_data["note"] = note
            row_data["misc"] = misc

            data.append(row_data)

def instagramtwo(): #alpha
    # from lib.colors import red,white,green,reset
    self = 'kevinrose'

    response = self.session.get(self.url)
    if response.status_code != 200:
        exit(f'{color_red}[-] instagram: user not found{color_reset}')
    response = response.json()
    user = response['graphql']['user']
           
    data = {'Profile photo': user['profile_pic_url_hd'],
                 'Username': user['username'],
                 'User ID': user['id'],
                 'External URL': user['external_url'],
                 'Bio': user['biography'],
                 'Followers': user['edge_followed_by']['count'],
                 'Following': user['edge_follow']['count'],
                 'Pronouns': user['pronouns'],
                 'Images': user['edge_owner_to_timeline_media']['count'],
                 'Videos': user['edge_felix_video_timeline']['count'],
                 'Reels': user['highlight_reel_count'],
                 'Is private?': user['is_private'],
                 'Is verified?': user['is_verified'],
                 'Is business account?': user['is_business_account'],
                 'Is professional account?': user['is_professional_account'],
                 'Is recently joined?': user['is_joined_recently'],
                 'Business category': user['business_category_name'],
                 'Category': user['category_enum'],
                 'Has guides?': user['has_guides']
    }
    print
    
    print(f"\n{user['full_name']} | Instagram{reset}")
    for key, value in data.items():
       print(f"├─ {key}: {value}")


def instantusername(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< instantusername {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking, note) = (user, '9 - instantusername', '')
        
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        user = user.rstrip()
        url = (f'https://instantusername.com/?q={user}')
        
        try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
            row_data["query"] = query
            row_data["ranking"] = ranking
            # row_data["fullname"] = fullname                
            row_data["url"] = url
            row_data["user"] = user
            # row_data["note"] = note
            data.append(row_data)

        except TypeError as error:
            print(f'{color_red}{error}{color_reset}')


def instructables(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< instructables {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - instructables')
        (city, country, fullname, titleurl, pagestatus) = ('', '', '', '', '')
        user = user.rstrip()
        url = (f'https://www.instructables.com/member/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        if '404' not in pagestatus:
            if "'" in titleurl:
                titleurl = titleurl.split("'")[0]
            fullname = titleurl
            
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                fullname = ''           
            
            
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 
            
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            data.append(row_data)
            
def inteltechniques(): 
    if len(users) > 0:
        row_data = {}
        ranking = '9 - manual'
        url = ('https://inteltechniques.com/tools/')

        row_data["ranking"] = ranking
        # row_data["user"] = user
        data.append(row_data)                

def internet(host="8.8.8.8", port=53, timeout=3):
    """
    Host: 8.8.8.8 (google-public-dns-a.google.com)
    OpenPort: 53/tcp
    Service: domain (DNS/TCP)
    """
    try:
        socket.setdefaulttimeout(timeout)
        socket.socket(socket.AF_INET, socket.SOCK_STREAM).connect((host, port))
        return True
    except socket.error as ex:
        print(f'this is an error')  # temp
        print(ex)
        return False

def ip_address(dnsdomain):
    (ip) = ('')
    """
    Ping the URL and return the IP address
    """
    try:
        ip = socket.gethostbyname(dnsdomain)
    except socket.gaierror:
        ip = ''
    return ip


def is_running_in_virtual_machine():
    # Check for common virtualization artifacts
    virtualization_artifacts = [
        "/dev/virtio-ports",
        "/dev/vboxguest",
        "/dev/vmware",
        "/dev/qemu",
        "/sys/class/dmi/id/product_name",
        "/proc/scsi/scsi",
    ]

    for artifact in virtualization_artifacts:
        if os.path.exists(artifact):
            return True
            print('This is running in a virtual machine')
    return False

def keybase():    # testuser=    kevin
    print(f'{color_yellow}\n\t<<<<< keybase.io {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - keybase')
        (fullname, firstname, lastname, middlename) = ('','','','')
        url = (f'https://keybase.io/{user}')
        
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        (fullname, info, note) = ('', '', '')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        try:

            content = content.strip()
            titleurl = titleurl.strip()
            fullname = titleurl
            if " (" in fullname:
                fullname = fullname.split(" (")[1].split(")")[0]
                if ' ' in fullname:
                    (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                else:
                    fullname = ''
            for eachline in content.split("\n"):
                if "@context" in eachline:
                    content = eachline.strip()
                elif 'og:title' in eachline and 'content=\"' in eachline:
                    fullname = eachline.split('\"')[1].split(' (')[0]
                elif 'ProfilePage\",\"description' in eachline:
                    info = eachline
                    # Load the JSON data
                    datatemp = json.loads(eachline)

                    # Extract the description value and print it
                    note = datatemp['description']
       
        except:
            pass
            
        # time.sleep(1) # will sleep for 1 seconds
        # if 'what you are looking for...it does not exist' not in content:
        # if 'Your conversation will be end-to-end encrypted' in content:
        if 'Following' in content:
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["middlename"] = middlename
            row_data["lastname"] = lastname
            row_data["info"] = info
            row_data["note"] = note
            data.append(row_data)



    
def kik(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< kik {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '4 - kik')
        (fullname, titleurl, pagestatus, content) = ('', '', '', '')
        (note, firstname, lastname, photo, misc, lastseen) = ('', '', '', '', '', '')
        (otherurl, info, misc) = ('', '', '')
        user = user.rstrip()
        url = (f'https://ws2.kik.com/user/{user}')
        
        misc = (f'https://kik.me/{user}')
               
        
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        for eachline in content.split(","):
            if "firstName" in eachline:
                firstname = eachline.strip().split(":")[1]
                firstname = firstname.strip('\"')
                firstname = firstname.title()
            elif "lastName" in eachline:
                lastname = eachline.strip().split(":")[1]
                lastname = lastname.strip('\"')
                lastname = lastname.upper()
                lastname = lastname.replace('"}', '')
            elif "displayPicLastModified" in eachline:
                note = eachline.strip().split(":")[1]
             
                
            elif "displayPic\"" in eachline:
                photo = eachline.strip().split(":\"")[1].split("\"")[0].replace("\\","")
                
            fullname = (f'{firstname} {lastname}')
            fullname = fullname.replace("\"}","")
            # if ' ' in fullname:
                # (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            # else:
                # fullname = ''
            
            
        if '404' not in pagestatus:
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}')

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = misc
            row_data["note"] = note
            row_data["misc"] = url
            row_data["info"] = photo        
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname


            data.append(row_data)


def linkedin(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< linkedin {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - linkedin')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        try:
            user = user.strip()
        except:
            pass        
        url = (f'https://www.linkedin.com/in/{user}')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "og:title" in eachline:
                try:
                    fullname = eachline.strip().split("\"")[1]
                except:
                    pass
        if 1==1:
        # if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            

            data.append(row_data)   

def mastadon(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< mastadon {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - mastadon')
        (fullname, lastname, firstname) = ('','','')
        (note, info, content, pagestatus) = ('', '', '', '')
    
        user = user.rstrip()
        url = (f'https://mastodon.social/@{user}')
        note = (f'https://mastodon.social/api/v2/search?q={user}')

        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("\n"):
            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1].split(' (')[0]

        if 'accounts\":[{' in content:        
            datatemp = json.loads(content)              # Convert JSON data to Python dictionary
            fullname = datatemp["accounts"][0]["display_name"]
            info = datatemp['accounts'][0]['avatar']

        if user == fullname:
            fullname = ''
               
        if ' ' in fullname:
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)

        if "uccess" in pagestatus and 'This resource could not be found' not in content:
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["note"] = note
            row_data["info"] = info            
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            # row_data["middlename"] = middlename
            row_data["lastname"] = lastname

            data.append(row_data)

def myfitnesspal(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< myfitnesspal {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - myfitnesspal')
        
        user = user.rstrip()
        url = (f'https://www.myfitnesspal.com/profile/{user}')

        (content, referer, osurl, titleurl, pagestatus) = request(url)

        if "uccess" in pagestatus and 'This resource could not be found' not in content:
            print(f'{color_green}{url}{color_reset}') 
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url

            data.append(row_data)


def myshopify():    # testuser=    rothys
    print(f'{color_yellow}\n\t<<<<< myshopify {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '5 - myshopify')
        url = (f'https://{user}.myshopify.com/')
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        (fullname, info, note) = ('', '', '')
        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
            content = content.strip()
            titleurl = titleurl.strip()
            for eachline in content.split("\n"):
                if "@context" in eachline:
                    content = eachline.strip()
                elif 'og:title' in eachline and 'content=\"' in eachline:
                    info = eachline.split('\"')[1].split(' (')[0]
                elif 'ProfilePage\",\"description' in eachline:
                    info = eachline
                    # # Load the JSON data
                    datatemp = json.loads(eachline)

                    # # Extract the description value and print it
                    note = datatemp['description']

        except:
            pass
            
        time.sleep(1) # will sleep for 1 seconds
        if 'Success' in pagestatus:

            response = requests.get(url)

            if response.history:
                otherurls = response.url
                note = (f'redirects to {otherurls}')

            print(f'{color_green}{url}	{color_yellow}{note}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["info"] = info
            row_data["note"] = note           
            row_data["fullname"] = fullname

            data.append(row_data)


def myspace_users():
    print(f'{color_yellow}\n\t<<<<< myspace {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:
        row_data = {}
        (query, ranking) = (user, '4 - myspace')
        url = f"https://myspace.com/{user}"
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (note) = ("")
        (content, referer, osurl, titleurl, pagestatus) = request_url(url)


        if 'Success' in pagestatus and ('Your search did not return any results') not in content:
            fullname = titleurl

            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	   {fullname}{color_reset}')

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["middlename"] = middlename
            row_data["lastname"] = lastname
            row_data["note"] = note            

            data.append(row_data)

def main_email(): 

    for email in emails:
        row_data = {}
        (query, ranking) = (email, '1 - main')

        if '@' in email.lower():
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["email"] = email
            data.append(row_data)

def main_ip(): 

    for ip in ips:
        row_data = {}
        (query, ranking) = (ip, '1 - main')
        if 1==1:
        # if '.' in ip.lower():
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["ip"] = ip
            data.append(row_data)
            
def main_phone(): 

    for phone in phones:
        row_data = {}
        (query, ranking, state) = (phone, '1 - main', '')
        state = phone_state_check(phone, state)
    
        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["phone"] = phone
        row_data["state"] = state
        data.append(row_data)            


def main_user():
    for user in users:
        row_data = {}
        (query, ranking) = (user, '1 - main')

        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["user"] = user
        data.append(row_data)      


def main_website():
    for website in websites:
        row_data = {}
        (query, ranking) = (website, '1 - main')

        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["website"] = website
        data.append(row_data) 

def massageanywhere():    # testuser=   Misty0427
    print(f'{color_yellow}\n\t<<<<< massageanywhere {color_blue}users{color_yellow} >>>>>{color_reset}')

    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - massageanywhere')
        url = (f'https://www.massageanywhere.com/profile/{user}')
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        (info, note, fulladdress) = ('', '', '')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (content, referer, osurl, titleurl, pagestatus) = request_url(url)
        
        
        try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
            # (content, referer, osurl, titleurl, pagestatus) = request_url(url)
                        
            
            
            
            
            print(f' temp test')
        except:
            pass

        if 1==1:
        # if 'Profile for' in titleurl:
            if 1==1:
            # if 'MassageAnywhere.com Profile for ' in titleurl:  
                titleurl = titleurl.replace('MassageAnywhere.com Profile for ','')
                if ' of ' in titleurl:
                    # titleurl = titleurl.split(' of ')
                    fullname = titleurl.split(' of ')[0]
                    fulladdress = titleurl.split(' of ')[1]

                    if ' ' in fullname:
                        (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                    else:
                        (fullname, firstname, lastname, middlename) = ('', '', '', '')

            print(f'{color_green}{url}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname

            
            # row_data["phone"] = phone
            # row_data["note"] = note
            
            # row_data["city"] = city
            row_data["content"] = content
            row_data["fulladdress"] = fulladdress
            row_data["titleurl"] = titleurl            
            row_data["pagestatus"] = pagestatus            
                        
            data.append(row_data)

def message_square(message, color):
    horizontal_line = f"+{'-' * (len(message) + 2)}+"
    empty_line = f"| {' ' * (len(message))} |"

    print(color + horizontal_line)
    print(empty_line)
    print(f"| {message} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')

def noInternetMsg():
    '''
    prints a pop-up that says "Connect to the Internet first"
    '''
    window = Tk()
    window.geometry("1x1")
      
    w = Label(window, text ='Translate-Inator', font = "100") 
    w.pack()
    messagebox.showwarning("Warning", "Connect to the Internet first") 

def osintIndustries_email(): 
    if len(emails) > 0:
        row_data = {}
        ranking = '9 - manual'
        url = ('https://app.osint.industries/')
        # app.osint.industries
        
        row_data["ranking"] = ranking
        row_data["url"] = url
        data.append(row_data)

def patreon(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< patreon {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '5 - patreon')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        user = user.rstrip()
        url = (f'https://www.patreon.com/{user}/creators')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        if '404' not in pagestatus:
            print(f'{color_green}{url}{color_yellow}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname
            
            data.append(row_data)


def paypal(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< paypal {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - paypal')

        (city, country, fullname, titleurl, pagestatus, state) = ('', '', '', '', '', '')
        (fulladdress, lastname, firstname, photo) = ('', '', '', '')
        (email, phone, info, misc) = ('', '', '', '')
        user = user.rstrip()
        url = (f'https://www.paypal.com/paypalme/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        (note) = ('')
        if '404' not in pagestatus:
            for eachline in content.split("\n"):
                if eachline == "": pass                                             # skip blank lines
                else:
                    # define the regular expression pattern
                    pattern = r'{"userInfo":{(.*?)}}'

                    # match the pattern to the input string
                    match = re.search(pattern, content)

                    # extract the data variable from the match object
                    if match:
                        # datatemp = match
                        datatemp = match.group(1)
                        note = datatemp
                        note = note.replace("null",'\"\"')
        # else:
            # print(f'{color_red}{user}{color_reset}') 
        if ':' in note:
            titleurl = titleurl.replace('PayPal.Me','').strip() # task
            # print(f'titleurl = {titleurl}   hello world')   # temp   
            # fullname = titleurl       

            # Extract variables using regex
            try:
                fullname = re.search(r'"displayName":"(.*?)"', datatemp).group(1)

                if ' ' in fullname:
                    (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            except:pass                
            try:    
                email = re.search(r'"displayEmail":(null|".*?")', datatemp).group(1)
            except:pass                
            if "null" in email:
                email = ''
            try:    
                phone = re.search(r'"displayMobilePhone":(null|".*?")', datatemp).group(1)
            except:pass                
            if "null" in phone:
                phone = ''

            try:    
                photo = re.search(r'"coverPhotoUrl":(null|".*?")', datatemp).group(1)
                photo = photo.replace('"', '')
            except:pass                
            if "null" in photo:
                photo = ''
            
            try:    
                city = re.search(r'"displayAddress":"(.*?)"', datatemp).group(1)
                if ", " in city:
                   temp = city.split(", ")
                   city = temp[0]
                   state = temp[1]                
            except:pass                
            try:    
                info = re.search(r'"website":(null|".*?")', datatemp).group(1)
            except:pass
            if "null" in info:
                info = ''
                
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["phone"] = phone 
            row_data["email"] = email 
            row_data["city"] = city            
            row_data["state"] = state
            row_data["fulladdress"] = fulladdress            
            row_data["fullname"] = fullname
            # row_data["info"] = photo
            row_data["misc"] = misc           
            row_data["firstname"] = firstname
            row_data["middlename"] = middlename
            row_data["lastname"] = lastname
            row_data["note"] = note            
  
            data.append(row_data)

def phone_dashes(phone):
    # Remove any non-digit characters from the phone number
    cleaned_number = ''.join(filter(str.isdigit, phone))
    
    # Add dashes at appropriate positions
    formatted_number = '-'.join([cleaned_number[:3], cleaned_number[3:6], cleaned_number[6:]])
    
    return formatted_number


def phone_state_check(phone, state):
    """
    Checks if the phone number starts with any U.S. state area code and if the state is empty ('').
    If both conditions are met, sets the state accordingly.

    Args:
        phone (str): The phone number as a string.
        state (str): The state abbreviation.

    Returns:
        str: Updated state abbreviation (state based on area code if conditions are met, otherwise original state).
    """
    phone = phone.lstrip('1').strip()

    area_codes_by_state = {
        "AL?": ["205", "251", "256", "334", "938"],
        "AK?": ["907"],
        "AZ?": ["480", "520", "602", "623", "928"],
        "AR?": ["479", "501", "870"],
        "CA?": ["209", "213", "279", "310", "323", "341", "408", "415", "424", "442", "510", "530", "559", "562", "619", "626", "628", "650", "657", "661", "707", "714", "747", "760", "805", "818", "820", "831", "858", "909", "916", "925", "949", "951"],
        "CO?": ["303", "719", "720", "970"],
        "CT?": ["203", "475", "860", "959"],
        "DE?": ["302"],
        "FL?": ["239", "305", "321", "352", "386", "407", "561", "689", "727", "754", "772", "786", "813", "850", "863", "904", "941", "954"],
        "GA?": ["229", "404", "470", "478", "678", "706", "762", "770", "912"],
        "HI?": ["808"],
        "ID?": ["208", "986"],
        "IL?": ["217", "224", "309", "312", "331", "447", "464", "618", "630", "708", "730", "773", "779", "815", "847", "861", "872"], 
        "IN?": ["219", "260", "317", "463", "574", "765", "812", "930"],
        "IA?": ["319", "515", "563", "641", "712"],
        "KS?": ["316", "620", "785", "913"],
        "KY?": ["270", "364", "502", "606", "859"],
        "LA?": ["225", "318", "337", "504", "985"],
        "ME?": ["207"],
        "MD?": ["240", "301", "410", "443", "667"],
        "MA?": ["339", "351", "413", "508", "617", "774", "781", "857", "978"],
        "MI?": ["231", "248", "269", "313", "517", "586", "616", "734", "810", "906", "947", "989"],
        "MN?": ["218", "320", "507", "612", "651", "763", "952"],
        "MS?": ["228", "601", "662", "769"],
        "MO?": ["314", "417", "557", "573", "636", "660", "816", "975"],
        "MT?": ["406"],
        "NE?": ["308", "402", "531"],
        "NV?": ["702", "725", "775"],
        "NH?": ["603"],
        "NJ?": ["201", "551", "609", "640", "732", "848", "856", "862", "908", "973"],
        "NM?": ["505", "575"],
        "NY?": ["212", "315", "329", "332", "347", "363", "516", "518", "585", "607", "631", "646", "680", "716", "718", "838", "845", "914", "917", "929", "934"], 
        "NC?": ["252", "336", "704", "743", "828", "910", "919", "980", "984"],
        "ND?": ["701"],
        "OH?": ["216", "220", "234", "283", "326", "330", "380", "419", "440", "513", "567", "614", "740", "937"],
        "OK?": ["405", "539", "580", "918"],
        "OR?": ["458", "503", "541", "971"],
        "PA?": ["215", "223", "267", "272", "412", "445", "484", "570", "610", "717", "724", "814", "878"],
        "RI?": ["401"],
        "SC?": ["803", "843", "854", "864"],
        "SD?": ["605"],
        "TN?": ["423", "615", "629", "731", "865", "901", "931"],
        "TX?": ["210", "214", "254", "281", "325", "346", "361", "409", "430", "432", "469", "512", "682", "713", "737", "806", "817", "830", "832", "903", "915", "936", "940", "956", "972", "979"],
        "UT?": ["385", "435", "801"],
        "VT?": ["802"],
        "VA?": ["276", "434", "540", "571", "703", "757", "804"],
        "WA?": ["206", "253", "360", "425", "509", "564"],
        "WV?": ["304", "681"],
        "WI?": ["262", "414", "534", "608", "715", "920"],
        "WY?": ["307"]
    }

    if state == "":
        for state_code, area_codes in area_codes_by_state.items():
            if any(phone.startswith(code) for code in area_codes):
                return state_code
    
    return state


def pinterest():    # testuser=    kevinrose     # add city
    print(f'{color_yellow}\n\t<<<<< pinterest {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - pinterest')
        (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '')
        (country, email, fullname,lastname,firstname) = ('', '', '','','')
        (success, note, photo, website, city, otherurls) = ('','','','','', '')

        url = (f'https://www.pinterest.com/{user}/')
        otherurls = (f'https://pinterest.com/search/users/?q={user}')

        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
            parts = titleurl.split(' (', 1)

            if len(parts) > 1:
                titleurl = parts[0]
            fullname = titleurl
            if ' ' in fullname:
                ranking = '4 - pinterest'
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                
            else:
                fullname = ''
  
        except:
            pass
        fullname = fullname.replace('User AVATAR', '')  # test
        if 'Success' in pagestatus:
            if titleurl != 'None':
                print(f'{color_green} {url}{color_yellow}	   {fullname}	{note}{color_reset}')
                
                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["user"] = user
                row_data["url"] = url
                row_data["fullname"] = fullname
                row_data["firstname"] = firstname
                # row_data["middlename"] = middlename
                row_data["lastname"] = lastname
                row_data["note"] = note            

                data.append(row_data)                


def poshmark():    # testuser=    kevinrose
    print(f'{color_yellow}\n\t<<<<< poshmark {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - poshmark')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        url = (f'https://poshmark.com/closet/{user}')
        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
        except:
            pass

        if 'Success' in pagestatus:
            for eachline in content.split("\n"):
                if eachline == "": pass                                             # skip blank lines
                elif "og:title" in eachline:
                    fullname = eachline.strip().split("\"")[1].replace('\'s Closet', '')
                    
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                fullname = ''                    
                    
                    # firstname = fullname

            if 1==1:

                print(f'{color_green} {url}{color_yellow}	   {fullname}{color_reset}')

                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["user"] = user
                row_data["url"] = url
                row_data["firstname"] = firstname
                row_data["fullname"] = fullname
                # row_data["middlename"] = middlename
                row_data["lastname"] = lastname


                data.append(row_data)   

def print_logo():
    
    art = """
 ___    _            _   _ _         _   _             _   
|_ _|__| | ___ _ __ | |_(_) |_ _   _| | | |_   _ _ __ | |_ 
 | |/ _` |/ _ \ '_ \| __| | __| | | | |_| | | | | '_ \| __|
 | | (_| |  __/ | | | |_| | |_| |_| |  _  | |_| | | | | |_ 
|___\__,_|\___|_| |_|\__|_|\__|\__, |_| |_|\__,_|_| |_|\__|
                               |___/                       

  """
    print(f'{color_blue}{art}{color_reset}')

def public():    # testuser=    kevinrose
    print(f'{color_yellow}\n\t<<<<< public {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '8 - public')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '','')
        url = (f'https://public.com/@{user}')
        # try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
        # except:
            # pass

        if 'Success' in pagestatus:
            for eachline in content.split("\n"):
                if eachline == "": pass                                             # skip blank lines
                elif "og:title" in eachline:
                    fullname = eachline.strip().split("\"")[1]
                    fullname = fullname.split(" (")[0]
                    if ' ' in fullname:
                        (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                    else:
                        fullname = ''
            if 1==1:
                print(f'{color_green} {url}{color_yellow}	   {fullname}{color_reset}')

                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["user"] = user
                row_data["url"] = url
                row_data["fullname"] = fullname
                row_data["firstname"] = firstname
                # row_data["middlename"] = middlename
                row_data["lastname"] = lastname

                data.append(row_data)   
 

def read_text(filename):
    """
        Reads the input file
        parses the data into lists
        exports it to the outputfile
    """

    message = (f'Reading {filename}')
    message_square(message, color_green)

    if not os.path.exists(filename):
        input(f"{color_red}{filename} doesnt exist.{color_reset}")
        sys.exit()
    elif os.path.getsize(filename) == 0:
        input(f'{color_red}{filename} is empty. Fill it with username, email, ip, phone and/or websites.{color_reset}')
        sys.exit()
    elif os.path.isfile(filename):
        inputfile = open(filename)
    else:
        input(f'{color_red}See {filename} does not exist. Hit Enter to exit...{color_reset}')
        sys.exit()
        
    for eachline in inputfile:
        (query, ranking, fullname, url, email, user) = ('', '', '', '', '', '')
        (phone, business, fulladdress, city, state, country) = ('', '', '', '', '', '')
        (note, AKA, DOB, SEX, info, misc) = ('', '', '', '', '', '')
        (firstname, middlename, lastname, associates, case, sosfilenumber) = ('', '', '', '', '', '')
        (owner, president, sosagent, managers, Time, Latitude) = ('', '', '', '', '', '')
        (Longitude, Coordinate, original_file, Source, Source_file_information, Plate) = ('', '', '', '', '', '')
        (VIS, VIN, VYR, VMA, LIC, LIY) = ('', '', '', '', '', '')
        (DLN, DLS, content, referer, osurl, titleurl) = ('', '', '', '', '', '')
        (pagestatus, ip, dnsdomain) = ('', '', '')

        eachline = eachline + "\t" * 52
        eachline = eachline.split('\t')  # splits by tabs

        query = (eachline[0].strip())
        if ranking == '':
            ranking = '1 - main'
        fullname = (eachline[2].strip())
        url = (eachline[3].strip())
        email = (eachline[4].strip())
        user = (eachline[5].strip())
        phone = (eachline[6].strip())
        busines = (eachline[7].strip())
        fulladdress = (eachline[8].strip())
        city = (eachline[9].strip())
        state = (eachline[10].strip())
        country = (eachline[11].strip())
        note = (eachline[12].strip())
        AKA = (eachline[13].strip())
        DOB = (eachline[14].strip())
        SEX = (eachline[15].strip())
        info = (eachline[16].strip())
        misc = (eachline[17].strip())
        firstname = (eachline[18].strip())
        middlename = (eachline[19].strip())
        lastname = (eachline[20].strip())
        associates = (eachline[21].strip())
        case = (eachline[22].strip())
        sosfilenumber = (eachline[23].strip())
        owner = (eachline[24].strip())
        president = (eachline[25].strip())
        sosagent = (eachline[26].strip())
        managers = (eachline[27].strip())
        Time = (eachline[28].strip())
        Latitude = (eachline[29].strip())
        Longitude = (eachline[30].strip())
        Coordinate = (eachline[31].strip())
        original_file = (eachline[32].strip())
        Source = (eachline[33].strip())
        Source_file_information = (eachline[34].strip())
        Plate = (eachline[35].strip())
        VIS = (eachline[36].strip())
        VIN = (eachline[37].strip())
        VYR = (eachline[38].strip())
        VMA = (eachline[39].strip())
        LIC = (eachline[40].strip())
        LIY = (eachline[41].strip())
        DLN = (eachline[42].strip())
        DLS = (eachline[43].strip())
        content = (eachline[44].strip())
        referer = (eachline[45].strip())
        osurl = (eachline[46].strip())
        titleurl = (eachline[47].strip())
        pagestatus = (eachline[48].strip())
        ip = (eachline[49].strip())
        dnsdomain = (eachline[50].strip())

        # Regex data type
        # if re.search(regex_email, query):    # terrible. matches emails and phone numbers
        if bool(re.search(r"^[\w\.\+\-]+\@[\w]+\.[a-z]{2,3}$", query)):  # regex email    # works
            email = query
            user = email.split('@')[0]
            temp1 = [email]
            if query.lower() not in emails:            # don't add duplicates
                emails.append(email)
        # elif re.search(regex_url, query):   # test
            # url = query

        elif re.search(regex_host, query):  # regex_host (Finds url and dnsdomain) # False positives for emails    # todo
            url = query

            if url.lower().startswith('http'):
                if url.lower() not in websites:            # don't add duplicates
                    websites.append(url)            
            else:
                logsource = 'IOC-dnsdomain'
                dnsdomain = query
            url = url.rstrip('/')
            if "@" not in url and url.lower() not in websites:            # don't add duplicates
                websites.append(url)            
            dnsdomain = url.lower()
            dnsdomain = dnsdomain.replace("https://", "")
            dnsdomain = dnsdomain.replace("http://", "")
            dnsdomain = dnsdomain.split('/')[0]
            notes2 = dnsdomain.split('.')[-1]
            dnsdomain = dnsdomain.rstrip('/')
            if dnsdomain.lower() not in dnsdomains:            # don't add duplicates
                dnsdomains.append(dnsdomain)
            
        elif re.search(regex_ipv4, query):  # regex_ipv4
            (ip) = (query)
            if query.lower() not in ips:            # don't add duplicates
                ips.append(ip)

        elif re.search(regex_ipv6, query):  # regex_ipv6
            (ip) = (query)
            if query.lower() not in ips:            # don't add duplicates
                ips.append(ip)

        elif re.search(regex_phone, query) or re.search(regex_phone11, query) or re.search(regex_phone2, query):  # regex_phone
            (phone) = (query)
            phone = re.sub(r'[\+\- \(\)]', '', phone).strip()    # E165 standard
            # phone = phone.replace("-", '')  # E165 standard 
            # phone = phone.replace('(','').replace(')','').replace(' ','')   # E165 standard 
            # phone = phone.replace("+", "")          # E165 standard doesn't have a + (E.164 has a +)      
            # phone = phone.lstrip('1') # E.165 standard 16365551212

            if phone not in phones:            # don't add duplicates
                phones.append(phone)

            
        elif query.lower().startswith('http'):
            url = query
            if url.lower() not in websites:            # don't add duplicates
                websites.append(url)            
        elif query.strip() == '':
            print(f'{color_red}blank input found{color_reset}')
        else:
            user = query
            if query.lower() not in users:            # don't add duplicates
                users.append(user)

    return emails,dnsdomains,ips,users,phones,websites


def read_xlsx(input_xlsx):

    """Read data from an xlsx file and return as a list of dictionaries.
    Read XLSX Function: The read_xlsx() function reads data from the input 
    Excel file using the openpyxl library. It extracts headers from the 
    first row and then iterates through the data rows, creating dictionaries 
    for each row with headers as keys and cell values as values.
    
    """
    message = (f'Reading {input_xlsx}')
    message_square(message, color_green)
 
    wb = openpyxl.load_workbook(input_xlsx, read_only=True, data_only=True, keep_links=False)
    ws = wb.active
    data = [] 

    # get header values from first row
    headers = [cell.value for cell in ws[1]]

    # get data rows
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_data = {}   # test
        row_data = dict(zip(headers, row))

    # dnsdomains = []

        # url
        url = (row_data.get("url") or '').strip()

        if url.lower() not in [u.lower() for u in websites]:
            websites.append(url)

        # dnsdomain
        dnsdomain = (row_data.get("dnsdomain") or '').strip()

        # Only append non-empty and non-duplicate values
        if dnsdomain and dnsdomain.lower() not in [d.lower() for d in dnsdomains]:
            dnsdomains.append(dnsdomain)
            
        # user
        user = (row_data.get("user") or '').strip()

        if user and user.lower() not in [u.lower() for u in users]:
            users.append(user)
            
        # ip
        ip = (row_data.get("ip") or '').replace('\n', '').strip()

        if ip and ip.lower() not in [i.lower() for i in ips]:
            ips.append(ip)
 
        # email
        email = (row_data.get("email") or '').strip()

        if '@' in email and email.lower() not in [e.lower() for e in emails]:
            emails.append(email)
            
        # phone
        phone = (row_data.get("phone") or '').strip()

        if phone:
            # Remove unwanted characters
            phone = re.sub(r'[^\d+]', '', phone)

            # Optional: validate E.164 format
            if re.match(r'^\+?[1-9]\d{1,14}$', phone) and phone not in phones:
                phones.append(phone)


        # business
        business = (
            row_data.get("business")
            or row_data.get("business/entity")
            or row_data.get("Business")
            or ''
        ).strip()

        # owner
        owner = (row_data.get("owner") or '').strip()

        # AKA
        AKA = (
            row_data.get("AKA") or
            row_data.get("aka") or
            row_data.get("alias") or
            ''
        ).strip()


        # city
        city = (row_data.get("city") or '').strip().title()
    
        # state
        state = (row_data.get("state") or '').strip()

        # DOB
        DOB = (
            row_data.get("DOB") or
            row_data.get("dob") or
            ''
        ).strip()

        # associates
        associates = (
            row_data.get("associates") or
            row_data.get("friend") or
            ''
        ).strip()

        # SEX
        SEX = (row_data.get("SEX") or row_data.get("gender") or '').strip()
        
        # firstname
        firstname = (row_data.get("firstname") or '').strip().title()

        # lastname
        lastname = (row_data.get("lastname") or '').strip().upper()


        # middlename
        middlename = row_data.get("middlename", "")

        # fullname
        fullname = (row_data.get("fullname") or '').strip()

        if not fullname:
            if firstname and lastname and middlename:
                fullname = f'{firstname} {middlename} {lastname}'
            elif firstname and lastname:
                fullname = f'{firstname} {lastname}'


        # timestamp
        timestamp = (row_data.get("Time") or '').strip()
        timestamp, time_orig, timezone = convert_timestamp(timestamp, '', '')




        row_data["user"] = user
        row_data["url"] = url
        row_data["email"] = email
        row_data["phone"] = phone
        row_data["fullname"] = fullname
        row_data["business"] = business
        row_data["city"] = city 
        row_data["state"] = state 
        row_data["DOB"] = DOB  
        row_data["SEX"] = SEX         
        row_data["AKA"] = AKA
        row_data["firstname"] = firstname
        row_data["middlename"] = middlename
        row_data["lastname"] = lastname
        row_data["Time"] = timestamp
        row_data["associates"] = associates
        row_data["owner"] = owner
        row_data["ip"] = ip
        row_data["dnsdomain"] = dnsdomain

        # data.append(row_data)
        try:
            data.append(row_data)
        except Exception as e:
            print(f"{color_red}Error appending data: {str(e)}{color_reset}")

    return data

def read_xlsx_basic_old(input_xlsx):
    message = (f'Reading basic intel: {input_xlsx}')
    message_square(message, color_green)

    wb = openpyxl.load_workbook(input_xlsx)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):  # Assuming headers are in the first row
        entry = dict(zip(headers_intel, row))
        reordered_entry = {key: entry[key] for key in headers_intel}  # Reorder the entry based on headers_intel
        data.append(reordered_entry)

    return data


def read_xlsx_basic(input_xlsx):
    message = (f'Reading basic intel: {input_xlsx}')
    message_square(message, color_green)

    # data = []
    wb = openpyxl.load_workbook(input_xlsx)
    ws = wb.active

    headers = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))  # Read headers from the first row

    for row in ws.iter_rows(min_row=2, values_only=True):  # Start iterating from the second row
        entry = dict(zip(headers, row))  # Use the headers read from the file
        filtered_entry = {key: entry[key] for key in headers_intel if key in headers}  # Filter out only the relevant headers
        data.append(filtered_entry)

    return data


def read_xlsx_basic_location(input_xlsx):
    message = (f'Reading basic intel: {input_xlsx}')
    message_square(message, color_green)

    # data = []
    wb = openpyxl.load_workbook(input_xlsx)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):  # Assuming headers are in the first row
        entry = dict(zip(headers_intel, row))
        data.append(entry)

    return data
    
    
def reddit(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< reddit {color_blue}users{color_yellow} >>>>>{color_reset}')

    for user in users:    
        row_data = {}
        (query, ranking) = (user, '6 - reddit')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.rstrip()
        url = (f'https://www.reddit.com/user/{user}/')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "nobody on Reddit goes by that name" in eachline or "This account has been suspended" in eachline or "This user has deleted their account" in eachline:
                ranking = '9 - reddit'
            elif "hasn't posted yet" in eachline :
                ranking = '8 - reddit'
                note = "hasn't posted yet"

        if '9' not in ranking:
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["user"] = user            
            row_data["note"] = note

            data.append(row_data)           


def redirect_detect():  # https://goo.gle
    print(f'{color_yellow}\n\t<<<<< redirected {color_blue}websites{color_yellow} >>>>>{color_reset}')
    for website in websites:    
        row_data = {}
        (query, ranking) = (website, '7 - redirect')
        (ip) = ('')    
        (final_url, dnsdomain, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '', '')
        url = website
        url = url.replace("http://", "https://")
        if "http" not in url.lower():
            url = (f'https://{website}')
            
        referer = url.lower().strip()
        try:
            response = requests.get(url)

            final_url = response.url

        except TypeError as error:

            pass
        
        dnsdomain = url.lower()
        dnsdomain = dnsdomain.replace("https://", "")
        dnsdomain = dnsdomain.replace("http://", "")
        dnsdomain = dnsdomain.split('/')[0]

        ip = ip_address(dnsdomain)
        
        if dnsdomain not in final_url:
            print(f'{color_green}{url} redirects to {final_url}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = final_url
            row_data["ip"] = ip
            row_data["titleurl"] = titleurl
            row_data["referer"] = referer
            row_data["dnsdomain"] = dnsdomain

            data.append(row_data)


def request_url(url):
    
    fake_referer = 'https://www.google.com/'
    headers_url = {'Referer': fake_referer}

    
    (content, referer, osurl, titleurl, pagestatus)= ('blank', '', '', '', '')
    (response) = ('')
    
    if url.lower().startswith('http'):
        blah = ''
    else:
        url = ("https://" +url)

    try:
        response = requests.get(url, verify=False, headers=headers_url)        
        response.raise_for_status()
        pagestatus  = response.status_code
        content = response.content.decode()
        content = BeautifulSoup(content, 'html.parser')
        


    except requests.exceptions.RequestException as e:
        # print(f"Could not fetch URL {url}: {str(e)}")
        # raise requests.exceptions.RequestException(str(e))
        (pagestatus) = ('Fail')

# osurl
    try:
        osurl = response.headers['Server']
    except:
    # except KeyError:
        pass

# titleurl
    try:
        titleurl = titleurl_get(content)
        # titleurl = titleurl_og(content)
    except KeyError:
        titleurl = ''


#pagestatus
    
    if str(pagestatus).startswith('2') :    
        pagestatus = (f'Success - pagestatus')
    elif str(pagestatus).startswith('3') :    
        pagestatus = (f'Redirect - {pagestatus}')
    elif str(pagestatus).startswith('4') :    
        pagestatus = (f'Fail - {pagestatus}')
    elif str(pagestatus).startswith('5') :    
        pagestatus = (f'Fail - {pagestatus}')
    elif str(pagestatus).startswith('1') :    
        pagestatus = (f'Info - {pagestatus}')

    pagestatus = pagestatus.strip()    

    return (content, referer, osurl, titleurl, pagestatus)



def resolverRS():# testIP= 77.15.67.232
    print(f'{color_yellow}\n\t<<<<< resolverRS {color_blue}ip{color_yellow} >>>>>{color_reset}')
   
    for ip in ips:
        row_data = {}
        (query) = (ip)
        (country, city, zipcode, case, note, state) = ('', '', '', '', '', '')
        (misc, info) = ('', '')
        
        url = (f'https://resolve.rs/ip/geolocation.html?ip={ip}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("\n"):
            if "403 ERROR" in eachline:
                pagestatus = '403 Error'
                content = ''

            elif "\"code\": \"" in eachline :   # and zipcode != ''
                zipcode = eachline.split("\"")[3]
            elif "\"en\": \"" in eachline :   # city
                # print(f'')
                if city == '':
                    city = eachline.split("\"")[3]
                elif misc == '' and city != '': # continent
                    misc = eachline.split("\"")[3]

                elif misc != '' and country == '' and city != '': # country
                    country = eachline.split("\"")[3]
                elif info == '' and misc != '' and country != '' and city != '':    # registered country
                    info = eachline.split("\"")[3]
                elif state == '' and info != '' and misc != '' and country != '' and city != '':    # state
                    state = eachline.split("\"")[3]
            elif "COMCAST" in eachline :   # isp  >ASN</a>
                note = 'COMCAST'


        # pagestatus = ''                
        if url != '':
            print(f'{color_green}{url}{color_reset}') 
            ranking = '6 - resolve.rs'
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["ip"] = ip
            row_data["note"] = note
            row_data["city"] = city
            row_data["country"] = country
               
            data.append(row_data)

    # return data

def request(url):
    (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
    (string) = ('')
    fake_referer = 'https://www.google.com/'
    headers_url = {'Referer': fake_referer}
    url = url.replace("http://", "https://")     # test

    page  = requests.get(url, headers=headers_url)

    pagestatus  = page.status_code
    soup = BeautifulSoup(page.content, 'html.parser')
    content = soup.prettify()
    try:
        osurl = page.headers['Server']
    except:pass
    
    try:
        titleurl = soup.title.string
    except:pass
    

#pagestatus
    
    if str(pagestatus).startswith('2') :    
        pagestatus = (f'Success - {pagestatus}')
    elif str(pagestatus).startswith('3') :    
        pagestatus = (f'Redirect - {pagestatus}')
    elif str(pagestatus).startswith('4') :    
        pagestatus = (f'Fail - {pagestatus}')
    elif str(pagestatus).startswith('5') :    
        pagestatus = (f'Fail - {pagestatus}')
    elif str(pagestatus).startswith('1') :    
        pagestatus = (f'Info - {pagestatus}')
    try:
        pagestatus = pagestatus.strip()
    except Exception as e:
        print(f"{color_red}Error striping pagestatus: {str(e)}{color_reset}")
# titleurl

    try:
        title_tag = content.find('title')
        if title_tag is not None:
            title = title_tag.text.strip()
            # The full name is often included in parentheses after the username
            # Example: "Tom Anderson (myspacetom)"
            # We'll split the title at the first occurrence of '(' to extract the full name
            parts = title.split(' (', 1)
            if len(parts) > 1:
                titleurl = parts[0]
    except Exception as e:
        # print(f"Error parsing title: {str(e)}")
        pass

    if titleurl !="":
        try:
            meta_tags = content.find_all('meta')
            for tag in meta_tags:
   
                if tag.get('property') == 'og:title':
                    titleurl = tag.get('content')
                    titleurl =  title.split(' (')[0]
        except Exception as e:
            # print(' ')
            # print(f'{color_red}Error parsing metadata: {str(e)}{color_reset}')
            pass

    try:
        titleurl = str(titleurl)    #test
        titleurl = (titleurl.encode('utf8'))    # 'ascii' codec can't decode byte
        titleurl = (titleurl.decode('utf8'))    # get rid of bytes b''
    except TypeError as error:
        print(f'{color_red}{error}{color_reset}')
        titleurl = ''

    titleurl = titleurl.strip()
    content = content.strip()
    
    return (content, referer, osurl, titleurl, pagestatus)   

def reversephonecheck():# testPhone= 
    print(f'{color_yellow}\n\t<<<<< reversephonecheck {color_blue}phone numbers{color_yellow} >>>>>{color_reset}')
    
    for phone in phones:
        row_data = {}
        (query) = (phone)
        ranking = '9 - reversephonecheck'
        (fulladdress, country, city, state, case, note) = ('', '', '', '', '', '')
        (content, referer, osurl, titleurl, pagestatus)  = ('', '', '', '', '')
        (areacode, prefix, line, count, match, match2) = ('', '', '', 1, '', '')
        (url) = ('')

        if len(phone.lstrip('1')) == 10:
            # print(f' {phone} has 10 digits')    # temp
            phone = phone.lstrip('1')
            phone = (phone[:3] + "-" + phone[3:6] + "-" + phone[6:])
            # print(f'{color_yellow}phone {color_reset} {phone}') # temp
            
        (line2) = ('')
        if "-" in phone:
            phone2 = phone.split("-")
            areacode = phone2[0]
            prefix = phone2[1]
            try:
                line = phone2[2]
                line2 = line
                line = line[:2]
                
            except:
                pass

        url = (f'https://www.reversephonecheck.com/1-{areacode}/{prefix}/{line}/#{phone[-2:]}' )

        (content, referer, osurl, titleurl, pagestatus) = request(url) 
        match = (f"{prefix} - {line2}")
        
        phone = phone.replace('-', '')
        for eachline in content.split("\n"):
            if match in eachline:
                pagestatus = 'research'
                ranking = '4 - reversephonecheck'
                count += 1

        if '404' in pagestatus:
            ranking = '99 - reversephonecheck'
        elif pagestatus == 'research' and count == 2:
            print(f'{color_green}{url}{color_reset} {phone}')
            ranking = '5 - reversephonecheck'
        else:
            print(f'{color_red}{url}{color_reset} {phone}')
            ranking = '9 - reversephonecheck'

        if state == '':
            state = phone_state_check(phone, state).replace('?', '')


        if '404' not in pagestatus:
        # if 1==1:    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["phone"] = phone
            row_data["note"] = note
            
            row_data["city"] = city
            row_data["state"] = state
            
            
            row_data["country"] = country
            row_data["fulladdress"] = fulladdress
      
            data.append(row_data)

def roblox(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< roblox {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - roblox')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.strip()
        url = (f'https://www.roblox.com/user.aspx?username={user}')

        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1]

        if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            # row_data["fullname"] = fullname
            # row_data["firstname"] = firstname
            # row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            # row_data["titleurl"] = titleurl
            # row_data["pagestatus"] = pagestatus
            # row_data["content"] = content
            
            data.append(row_data)  

def rumble(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< rumble {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - rumble')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.strip()
        # url = (f'https://rumble.com/user/{user}/about')
        url = (f'https://rumble.com/c/{user}/about')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1]

        if 1==1:
        # if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            row_data["titleurl"] = titleurl
            row_data["pagestatus"] = pagestatus
            row_data["content"] = content
            
            data.append(row_data)  


def samples():
    print(f'''{color_yellow}    
Alain_THIERRY
anh_usa
frizz925
gurkha_love
Hmei7
itdoesnthavetohappen
JLight29
kevinrose
KC3ELO
kevinwagner
kobegal
luckyjames844
maverick3819
merz356
Misty0427
MR-JEFF
N3tAtt4ck3r
Pattycakes98
rothys
ryanlwatkins
thekevinrose
williger
zazenergy
nullcrew
realDonaldTrump
Their1sn0freakingwaythisisreal12JT4321
        
77.15.67.232
92.20.236.78
255.255.255.255

2607:f0d0:1002:51::4
2607:f0d0:1002:0051:0000:0000:0000:0004

annemconnor@yahoo.com
kandyem@yahoo.com
kevinrose@gmail.com
craig@craigslist.org
ceo@zappos.com
lnd_whitaker@yahoo.com
gsmstocks@gmail.com
lydianorman1@hotmail.com
soniraj388@gmail.com
tanderson09@gmail.com
tin_max87@yahoo.com
Their1sn0freakingwaythisisreal12344321@fakedomain.com

385-347-1531
312-999-9999
5596833344
15596833344
{color_reset}
'''
)    


def sherlock():    # testuser=    kevinrose
    print(f'\n\t{color_yellow}<<<<< Manually check Sherlock users >>>>>{color_reset}')
    
    for user in users:    
        note = (f'cd C:\Forensics\scripts\python\git-repo\sherlock && python sherlock {user}')
        row_data = {}
        (query, ranking) = (user, '9 - manual')

        if 1==1:

            row_data["query"] = query
            row_data["ranking"] = ranking
            # row_data["url"] = url
            row_data["note"] = note
      
            data.append(row_data)


def signal(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< signal {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - signal')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.strip()
        url = (f'https://www.signal.com/people/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1]


        if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            row_data["titleurl"] = titleurl
            row_data["pagestatus"] = pagestatus
            row_data["content"] = content
            
            data.append(row_data)   
            

def slack(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< slack {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '6 - slack')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.strip()
        url = (f'https://{user}.slack.com')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        # for eachline in content.split("  <"):
            # if "og:title" in eachline:
                # fullname = eachline.strip().split("\"")[1]


        # Regex to extract teamName
        fullname_match = re.search(r'"teamName":"(.*?)"', content)
        fullname = fullname_match.group(1) if fullname_match else ""

        # Regex to extract formattedEmailDomains
        dnsdomain_match = re.search(r'"formattedEmailDomains":"(.*?)"', content)
        dnsdomain = dnsdomain_match.group(1) if dnsdomain_match else ""
        dnsdomain = dnsdomain.lstrip('@')


        if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            # row_data["titleurl"] = titleurl
            row_data["dnsdomain"] = dnsdomain
            # row_data["content"] = content
            
            data.append(row_data)   
            

def snapchat(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< snapchat {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - snapchat')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        user = user.rstrip()
        url = (f'https://www.snapchat.com/add/{user}?')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("\n"):

            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1].replace(' on Snapchat','')
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                ranking = ('5 - snapchat')
            else:
                fullname = ''

        if 'name=\"description' in content:

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["middlename"] = middlename
            row_data["lastname"] = lastname
            row_data["user"] = user
            data.append(row_data)


def spotify(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< spotify {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - spotify')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        # (city, country, fullname, titleurl, pagestatus) = ('', '', '', '', '')
        user = user.rstrip()
        url = (f'https://open.spotify.com/user/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        if '404' not in pagestatus:
            titleurl = titleurl.replace(" on Spotify","").strip()
            fullname = titleurl
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                ranking = ('5 - spotify')
            else:
                fullname = ''
            
            
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname

            
            # row_data["phone"] = phone
            # row_data["note"] = note
            
            # row_data["city"] = city
            # row_data["country"] = country
            # row_data["fulladdress"] = fulladdress
            # row_data["titleurl"] = titleurl            
            # row_data["pagestatus"] = pagestatus            
                        
            data.append(row_data)


def spydialer():# testPhone= 
    print(f'{color_yellow}\n\t<<<<< spydialer {color_blue}users{color_yellow} >>>>>{color_reset}')

    for phone in phones:
        row_data = {}
        (query, pagestatus, state) = (phone, 'research', '')
        state = phone_state_check(phone, state) 
        url = ('https://www.spydialer.com')
        print(f'{color_yellow}{phone}{color_reset}')

        ranking = '3 - spydialer'
        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["url"] = url
        row_data["phone"] = phone
        row_data["state"] = state        
        # row_data["pagestatus"] = pagestatus            
                        
        data.append(row_data)

  
def thatsthememail():   # testEmail= smooth8101@yahoo.com 
    print(f'{color_yellow}\n\t<<<<< thatsthem {color_blue}emails{color_yellow} >>>>>{color_reset}')
     
    for email in emails:
        # print(f'{color_red}{email}{color_reset}')
        row_data = {}
        (query, content, note) = (email, '', '')
        (country, city, zipcode, case, note) = ('', '', '', '', '')
        
        url = (f'https://thatsthem.com/email/{email}')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("\n"):
            if "Found 0 results for your query" in eachline and case == '':
                # print(f'{color_red}not found{color_reset}')  # temp
                url = ('')

        pagestatus = ''                
        if url != '':
            # print(f'{color_green}{url}{color_yellow}	{email}{color_reset}') 
            ranking = '8 - thatsthem'
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["email"] = email
            # row_data["note"] = note
            
            # row_data["city"] = city
            # row_data["country"] = country
            # row_data["referer"] = city
            # row_data["titleurl"] = titleurl            
            # row_data["pagestatus"] = pagestatus            
                        
            data.append(row_data)
            # print(f'row_data = {row_data}') # temp


        else:
            print(f'{color_red}{url}{color_yellow}	{email}{color_reset}') 

def thatsthemip():# testIP= 8.8.8.8
    print(f'{color_yellow}\n\t<<<<< thatsthem {color_blue}ip{color_yellow} >>>>>{color_reset}')
       
    for ip in ips:
        row_data = {}
        (country, city, zipcode, case, note, state) = ('', '', '', '', '', '')
        
        url = (f'https://thatsthem.com/ip/{ip}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("\n"):
            if "located in " in eachline:
                state = eachline
                note = eachline

            elif "403 ERROR" in eachline:
                pagestatus = '403 Error'
                content = ''
            elif "Found 0 results for your query" in eachline:
                print(f'{color_red}Not found{color_reset}')  # temp
                url = ('')
        # pagestatus = ''                
        if url != '':
            ranking = '9 - thatsthem'
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["ip"] = ip
            # row_data["note"] = note
            
            # row_data["city"] = city
            # row_data["country"] = country
            # row_data["referer"] = city
            # row_data["titleurl"] = titleurl            
            # row_data["pagestatus"] = pagestatus            
                        
            data.append(row_data)

def thatsthemphone():# testPhone=   
    # note: need to add dishes if there are none
    print(f'{color_yellow}\n\t<<<<< thatsthem {color_blue}phone numbers{color_yellow} >>>>>{color_reset}')
    for phone in phones:
        
        row_data = {}
        (query, ranking) = (phone, '9 - thatsthem')
        
        if '-' not in phone:
            phone = phone_dashes(phone.lstrip('1'))
        
        (country, city, zipcode, case, note, state) = ('', '', '', '', '', '')
        (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '')
        
        time.sleep(2) # will sleep for 10 seconds
        url = ('https://thatsthem.com/phone/%s' %(phone.lstrip('1')))    # https://thatsthem.com/reverse-phone-lookup


        if "Found 0 results for your query" in content or "The request could not be satisfied" in content:
            url = ('')
            note = ('captcha protected')

        for eachline in content.split("\n"):
            if "Found 0 results for your query" in eachline and case == '':
                print(f'{color_red}Not found{color_reset}')  # temp
                url = ('')

        phone = phone.replace('-', '')
        
        if note == '':
            ranking = '6 - thatsthem'
            state = phone_state_check(phone, state).replace('?', '')
            
        else:   
            ranking = '9 - thatsthem'

        if 1==1:
        # if url != '':
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["phone"] = phone
            row_data["state"] = state
            row_data["note"] = note
            data.append(row_data)

def telegram(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< telegram {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking, note) = (user, '7 - telegram', '')
        
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', 'research', '')
        user = user.rstrip()
        url = (f'https://t.me/{user}')
        
        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
            # (content, referer, osurl, titleurl, pagestatus) = request_url(url)

            for eachline in content.split("\n"):
                if "og:title" in eachline:
                    fullname = eachline.strip().split("\"")[1]

            if 'Telegram' not in fullname:
                print(f'{color_green}{url}{color_yellow}	{titleurl}{color_reset}') 

                row_data["query"] = query
                row_data["ranking"] = ranking
                row_data["fullname"] = fullname                
                row_data["url"] = url
                row_data["user"] = user
                row_data["note"] = note
                data.append(row_data)

        except TypeError as error:
            print(f'{color_red}{error}{color_reset}')


def threads():    # testuser=    kevinrose     # add info
    print(f'{color_yellow}\n\t<<<<< threads {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '3 - threads')
        (fullname, firstname, lastname, middlename, note)  = ('','','','', '')

        url = (f'https://www.threads.net/@{user}')

        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
 
            content = content.strip()
            titleurl = titleurl.strip()
            for eachline in content.split("\n"):
                if "@context" in eachline:
                    content = eachline.strip()
                elif 'og:title' in eachline and 'content=\"' in eachline:
                    fullname = eachline.split('\"')[1].split(' (')[0]
                elif "og:description" in eachline:
                    note = eachline.strip()
                    note = note.replace("\" property=\"og:description\"/>",'').replace("<meta content=\"",'')

        except:
            pass
        
        if ' ' in fullname:
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
        else:
            fullname = ''
        
        if 'on Threads' in titleurl:
            print(f'{color_green}{url}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["middlename"] = middlename
            row_data["fullname"] = fullname
            row_data["note"] = note
             
            data.append(row_data)


def tiktok(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< tiktok {color_blue}users{color_yellow}>>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - tiktok')
        (fullname, firstname, lastname, middlename, note)  = ('','','','', '')


        # (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', 'research', '')
        user = user.rstrip()
        url = (f'https://tiktok.com/@{user}?')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        
        if 'uccess' in pagestatus:
            fullname = titleurl
            fullname = fullname.split(' (')[0]
            if fullname == user:
                fullname = ''
            elif 'Make Your Day' in fullname:
                ranking = '8 - tiktok'
                fullname = ''
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
                ranking = '4 - tiktok'
            else:
                fullname = ''            
         
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["user"] = user
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname
            # row_data["titleurl"] = titleurl
            # row_data["pagestatus"] = pagestatus
            
                                   
            data.append(row_data)

def tinder():    # testuser=    john
    print(f'{color_yellow}\n\t<<<<< tinder {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '7 - tinder')
        (fullname, firstname, lastname, middlename, note, DOB)  = ('','','','', '', '')
        (misc) = ('')

        url = (f'https://tinder.com/@{user}')

        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
            
            
            
            content = content.strip()
            titleurl = titleurl.strip()
            for eachline in content.split("\n"):
                if "@context" in eachline:
                    content = eachline.strip()
                elif 'og:title' in eachline and 'content=\"' in eachline:
                    firstname = eachline.split('\"')[1].split(' (')[0]
                    
                elif 'schools\"' in eachline:
                # elif 'schools\":\[{\"name' in eachline:
                    datatemp = json.loads(eachline)
                    note = datatemp["schools"][0]["name"]
                    # Extract the description value and print it
                    # note = data['description']
                    misc = eachline
                    
                    DOB= datatemp["webProfile"]["user"]["birth_date"]
        except:
            pass
            
        # time.sleep(1) # will sleep for 1 seconds
        if 'alternate' in content:
            print(f'{color_green}{url}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname
            row_data["note"] = note
            row_data["DOB"] = DOB
            row_data["misc"] = misc

            data.append(row_data)
 
def titleurl_get(content):

    titleurl = ''
    try:
        # soup = BeautifulSoup(html, 'html.parser')
        title_tag = content.find('title')
        if title_tag is not None:
            title = title_tag.text.strip()
            # The full name is often included in parentheses after the username
            # Example: "Tom Anderson (myspacetom)"
            # We'll split the title at the first occurrence of '(' to extract the full name
            parts = title.split(' (', 1)
            if len(parts) > 1:
                titleurl = parts[0]
    except Exception as e:
        # print(f'Error parsing title: {str(e)}')
        pass


    return titleurl

def titleurl_og(content):
    (titleurl) = ('')

    try:
        meta_tags = content.find_all('meta')
        for tag in meta_tags:
            if tag.get('property') == 'og:title':
                titleurl = tag.get('content')
                titleurl =  title.split(' (')[0]
    except Exception as e:
        pass
    return titleurl


def tripadvisor(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< tripadvisor {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - tripadvisor')
        (city, country, fullname, titleurl, pagestatus, content) = ('', '', '', '', '', '')
        (note) = ('')
        user = user.strip()
        url = (f'https://www.tripadvisor.com/Profile/{user}')

        (content, referer, osurl, titleurl, pagestatus) = request(url)

        for eachline in content.split("  <"):
            if "og:title" in eachline:
                fullname = eachline.strip().split("\"")[1]

        if 1==1:
        # if '404' not in pagestatus and 'ail' not in pagestatus:
            if fullname.lower() == user.lower():
                fullname = ''
            
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            if " " not in fullname:
                fullname = ''

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
            row_data["user"] = user            
            row_data["titleurl"] = titleurl
            row_data["pagestatus"] = pagestatus
            row_data["content"] = content
            
            data.append(row_data)  
            
            
def truepeople_email(): 
    '''
    this is protected by javascript, cookies and cloud flair
    it would require using selenium and a hard coded webdriver
    '''
    
    print(f'{color_yellow}\n\t<<<<< truepeoplesearch {color_blue}emails{color_yellow} >>>>>{color_reset}')
    
    for email in emails:
        row_data = {}
        (query, content, note) = (email, '', '')
        url = f'https://www.truepeoplesearch.com/resultemail?email={email}'
        # (content, referer, osurl, titleurl, pagestatus) = request(url)
        if '@' in email.lower():

            if 'Enable JavaScript and cookies to continue' in content:
                print(f'blah')
            # if 'We could not find any records for that search criteria' in content:
                ranking = '99 - truepeoplesearch'    # needs work
            elif 'All retrieved results' in content:
                note = 'Events shown in time zone'
                ranking = '4 - truepeoplesearch'
                # print(f'{color_green} {email}{color_reset}')                
            else:
                ranking = '9 - truepeoplesearch'
                # print(f'{color_yellow} {email}  {color_reset}')

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["email"] = email
            row_data["note"] = note
            # row_data["content"] = content
            data.append(row_data)
            # print(f'row_data = {row_data}') # temp


def truthSocial(): # testuser = realdonaldtrump https://truthsocial.com/@realDonaldTrump
    print(f'{color_yellow}\n\t<<<<< truthsocial {color_blue}users{color_yellow} >>>>>{color_reset}')
    print(f'{color_yellow}\n\t\t\tThis one one takes a while{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - truthsocial')

        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (note) = ('')
        (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '')
        user = user.rstrip()
        url = (f'https://truthsocial.com/@{user}')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)
        (fullname,lastname,firstname, email, name, country) = ('','','', '', '', '')
        pagestatus = ''
        # time.sleep(3) #will sleep for 3 seconds
        for eachline in content.split("  <"):
            if 'This resource could not be found' in eachline:
                pagestatus = '404'
            elif "og:title" in eachline:
                titleurl = eachline.strip().split("\"")[1]
                fullname = titleurl.split(" (")[0]
                pagestatus = '200'
                ranking = '9 - truthsocial'
                if titleurl == 'Truth Social':
                    pagestatus = '404'
                else:
                    pagestatus = '200'
                    ranking = '3 - truthsocial'
            elif "og:description" in eachline:
                note = eachline.strip().split("\"")[1]

        if ' ' in fullname:
            (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
        else:
            fullname = ''

        if 1==1:
        # if '@' in titleurl: 
            print(f'{color_yellow}{url}{color_yellow}	{fullname}{color_reset}') 
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["user"] = user
            # row_data["lastname"] = lastname
            # row_data["firstname"] = firstname
            row_data["fullname"] = fullname
            row_data["note"] = note
            row_data["titleurl"] = titleurl
            row_data["pagestatus"] = pagestatus
            data.append(row_data)

   
def twitter():    # testuser=    kevinrose     # add info
    print(f'{color_yellow}\n\t<<<<< twitter {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - X')

        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        url = (f'https://x.com/{user}')
        (content, referer, osurl, titleurl, pagestatus) = ('','','', '', '')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)
        # print(titleurl, url, pagestatus)  # temp
        try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
            # print(titleurl)  # temp
            titleurl = titleurl.replace(") on Twitter","")
            titleurl = titleurl.lower().replace(User.lower(),"")
            titleurl = titleurl.replace(" (","")
            fullname = titleurl
            fullname = fullname.replace(" account suspended","")
            fullname = fullname.replace("twitter /","")
            titleurl = titleurl.lower().replace(fullname.lower(),"")

            print(f'{color_green}{url}{color_yellow}	   {fullname}	{titleurl}{color_reset}')

            ranking = '5 - X'
        except:
            # print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            print(f'{color_yellow}{url}{color_yellow}	   {fullname}	{titleurl}{color_reset}')
            ranking = '9 - X'


        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["user"] = user
        row_data["url"] = url
        row_data["lastname"] = lastname
        row_data["firstname"] = firstname
        row_data["fullname"] = fullname
        # row_data["note"] = note
        row_data["titleurl"] = titleurl
        row_data["pagestatus"] = pagestatus

        data.append(row_data)

        # time.sleep(10) #will sleep for 10 seconds


def veraxity(): 
    if len(emails) > 0:
        row_data = {}
        ranking = '8 - manual'
        url = ('https://intel.veraxity.org')
        note = ('https://breachbase.com/')
        row_data["ranking"] = ranking
        row_data["url"] = url
        row_data["note"] = note
        data.append(row_data)
            

def vimeo():    # testuser=    kevinrose
    print(f'{color_yellow}\n\t<<<<< vimeo {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - vimeo')
        (fullname, firstname, lastname, middlename, note, DOB)  = ('','','','', '', '')
        (misc) = ('')

        url = (f'https://vimeo.com/{user}')
        (content, referer, osurl, titleurl, pagestatus) = ('','','','','')

        try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)

            content = content.strip()
            titleurl = titleurl.strip()

        except:
            pass
            
        # time.sleep(1) # will sleep for 1 seconds
        if 1==1:
        # if 'alternate' in content:
            print(f'{color_green}{url}{color_reset}')    

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["lastname"] = lastname
            row_data["firstname"] = firstname
            row_data["fullname"] = fullname
            row_data["note"] = note
            # row_data["DOB"] = DOB
            # row_data["misc"] = misc
            row_data["content"] = content
            row_data["pagestatus"] = pagestatus
            row_data["titleurl"] = titleurl

            data.append(row_data)

            

def whatismyip():    # testuser= 77.15.67.232  
    print(f"{color_yellow}\n\t<<<<< whatismyipaddress.com {color_blue}IP's{color_yellow} >>>>>{color_reset}")
    for ip in ips:
        row_data = {}
        (query, city, state, zipcode, pagestatus, title) = (ip, '', '', '', '', '')
        url = (f'https://whatismyipaddress.com/ip/{ip}')

        ranking = '9 - whatismyipaddress'
        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["url"] = url
        row_data["ip"] = ip
        data.append(row_data)

def whatnot(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< whatnot {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - whatnot')
        (fullname, firstname, lastname, middlename) = ('', '', '', '')
        (city, country, fullname, titleurl, pagestatus) = ('', '', '', '', '')
        user = user.rstrip()
        url = (f'https://www.whatnot.com/user/{user}')
        
        # (content, referer, osurl, titleurl, pagestatus) = request(url)
        (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '')
        
        
        if '404' not in pagestatus:
            titleurl = titleurl.replace("Just a moment...","").strip()

            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["titleurl"] = titleurl
            data.append(row_data)


def whatsmyname():    # testuser=   kevinrose
    print(f'\n\t{color_yellow}<<<<< Manually check whatsmyname users >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        query = user
        url = ('https://whatsmyname.app/')
        
        note = (f'cd C:\Forensics\scripts\python\git-repo\WhatsMyName && python web_accounts_list_checker.py -u {user} -of C:\Forensics\scripts\python\output_{user}txt') 
        ranking = '9 - manual'
        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["url"] = url
        row_data["note"] = note
        data.append(row_data)

def whitepagesphone():# testuser=    210-316-9435
    print(f'{color_yellow}\n\t<<<<< whitepages {color_blue}phone numbers{color_yellow} >>>>>{color_reset}')
    for phone in phones:    
        row_data = {}
        (query, ranking, state) = (phone, '9 - whitepages', '')
        url = ('https://www.whitepages.com/phone/1-%s' %(phone.lstrip('1')))
        state = phone_state_check(phone, state)
        # (content, referer, osurl, titleurl, pagestatus) = request(url)    # access denied cloudflare
       
        row_data["query"] = query
        row_data["ranking"] = ranking
        row_data["url"] = url
        row_data["phone"] = phone
        row_data["state"] = state
        data.append(row_data)    

def whocalld():# testPhone=  DROP THE LEADING 1
    print(f'{color_yellow}\n\t<<<<< whocalld {color_blue}phone numbers{color_yellow} >>>>>{color_reset}')

    # https://whocalld.com/+17083728101
    for phone in phones:
        row_data = {}
        (query, note) = (phone, '')
        (country, city, state, zipcode, case, note) = ('', '', '', '', '', '')
        (fullname, content, referer, osurl, titleurl, pagestatus)  = ('', '', '', '', '', '')
 
        url = ('https://whocalld.com/+1%s' %(phone.lstrip('1')))
        
        (content, referer, osurl, titleurl, pagestatus) = request(url)    # protected by cloudflare

        for eachline in content.split("\n"):
            if "Not found" in eachline and case == '':

                url = ('')
            elif "This seems to be" in eachline:
                if ' in ' in eachline:
                    note = eachline.replace(". </p>",'').replace("<p>",'').strip().replace("This",phone)
                    city = eachline.split(" in ")[1].replace(". </p>",'').replace("<p>",'').strip()
                    if ", " in city:
                        state = city.split(", ")[1].replace(".",'')
                        city = city.split(", ")[0]
                    note = (f'According to {url} {note}')
            elif "The name of this caller seemed to be " in eachline:
                note = eachline
                fullname = eachline.replace("The name of this caller seemed to be ",'').split(",")[0].strip()
                if ' in ' in eachline:
                    # try:
                        # note = eachline.replace(". </p>",'').replace("<p>",'').strip().replace("This",phone)
                        # city = eachline.split(" in ")[2].replace(". </p>",'').replace("<p>",'').strip()
                    # except TypeError as error:
                        # print(f'{color_red}{error}{color_reset}')

                    if ", " in city:
                        state = city.split(", ")[1].replace(".",'')
                        city = city.split(", ")[0]
        if state == '':
            state = phone_state_check(phone, state).replace('?', '')

        pagestatus = ''        
                
        if url != '':
            print(f'{color_green}{url}{color_yellow}	{fullname}  {city}  {state}{color_reset}') 
            
            ranking = '3 - spydialer'
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = 'https://www.spydialer.com'
            # row_data["url"] = url            
            row_data["note"] = note
            row_data["fullname"] = fullname
            row_data["phone"] = phone
            row_data["note"] = note
            row_data["city"] = city
            row_data["country"] = country
            row_data["state"] = state
            row_data["zipcode"] = zipcode            
            data.append(row_data)

        else:
            ranking = '4 - spydialer'
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = 'https://www.spydialer.com'
            row_data["phone"] = phone
            row_data["state"] = state
            data.append(row_data)

        time.sleep(2) 

def whoisip():    # testuser=    77.15.67.232   only gets 403 Forbidden
    from subprocess import call, Popen, PIPE
    print(f"{color_yellow}\n\t<<<<< whois {color_blue}IP's{color_yellow} >>>>>{color_reset}")
    for ip in ips:    
        row_data = {}
        (query, ranking, note) = (ip, '9 - whois', '')
        (city, business, country, zipcode, state) = ('', '', '', '', '')
        (Latitude, Longitude) = ('', '')

        (content, titleurl, pagestatus) = ('', '', '')
        (email, phone, fullname, entity, fulladdress) = ('', '', '', '', '') 
        url = (f'https://www.ipaddress.com/ipv4/{ip}')      

        if regex_ipv6.match(ip):
            url = f'https://www.ipaddress.com/ipv6/{ip}'

        if sys.platform == 'win32' or sys.platform == 'win64':    
            (content, referer, osurl, titleurl, pagestatus) = request(url)
            if '403 Forbidden' in content:
                pagestatus = '403 Forbidden'
                content = ''

            for eachline in content.split("\n"):
                 #
                if " is located in " in eachline:
                    (ranking) = ('4 - whois')
                    meta_match = re.search(r'content="([^"]+)"', eachline)
                    if meta_match:
                        note = meta_match.group(1)
                    note = note.replace(' | Public IP Address', '')
                if 'Reserved IP Address' in eachline:
                    note = (f'{ip} is a reserved IP address.')



                    
            time.sleep(3) #will sleep for 30 seconds
            if "Fail" in pagestatus:
                pagestatus = 'fail'
        else:
            WhoisArgs = (f'whois {ip}')
            response= Popen(WhoisArgs, shell=True, stdout=PIPE)
            for line in response.stdout:
                line = line.decode("utf-8")
                if ':' in line and "# " not in line and len(line) > 2:
                    line = line.strip()
                    content = (f'{content}\n{line}')
                if email == '':
                    if line.startswith('RAbuseEmail:'):
                        try:
                            email = (line.split(': ')[1].lstrip())
                        except:pass    
                    elif line.lower().startswith('abuse-mailbox:'):email = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('orgabuseemail:'):email = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('Orgtechemail:'):email = (line.split(': ')[1].lstrip())
                
                if phone == '':
                    if line.lower().startswith('rabusephone:'):phone = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('orgabusephone:'):phone = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('phone:'):phone = (line.split(': ')[1].lstrip())
                    phone = phone.replace("+", "")
                if line.lower().startswith('rtechname:'):fullname = (line.split(': ')[1].lstrip())
                elif line.lower().startswith('person:'):fullname = (line.split(': ')[1].lstrip())                
                
                if line.lower().startswith('country:'):country = (line.split(': ')[1].lstrip())
                if line.lower().startswith('city:'):city = (line.split(': ')[1].lstrip())
                if line.lower().startswith('address:'):fulladdress = ('%s %s' %(fulladdress, line.split(': ')[1].lstrip()))
                if line.lower().startswith('stateprov:'):state = (line.split(': ')[1].lstrip())
                if line.lower().startswith('postalcode:'):zipcode = (line.split(': ')[1].lstrip())
                if line.lower().startswith('orgname:'):entity = (line.split(': ')[1].lstrip())
                elif line.lower().startswith('org-name:'):entity = (line.split(': ')[1].lstrip())

        print(f'{color_green}{ip}{color_yellow}	{country}	{city}	{zipcode}{color_reset}')
        if note == '':
            note = entity


        if '.' in ip or ':' in ip:
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["ip"] = ip
            row_data["note"] = note
            row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["email"] = email
            row_data["phone"] = phone
            # row_data["note"] = entity
            row_data["fulladdress"] = fulladdress
            row_data["query"] = query
            
            row_data["city"] = city
            row_data["country"] = country
            row_data["state"] = state
            row_data["zipcode"] = zipcode            
            row_data["Latitude"] = Latitude
            row_data["Longitude"] = Longitude
            # row_data["content"] = content # temp
            # row_data["pagestatus"] = pagestatus # temp
           
           
            data.append(row_data)

def wordpress(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< wordpress {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - wordpress')


        (Success, fullname, lastname, firstname, case, SEX) = ('','','','','','')
        (photo, country, website, email, language, username) = ('','','','','','')
        (city, note) = ('', '')
        user = user.rstrip()
        url = (f'https://wordpress.org/support/users/{user}/')
        note = (f'https://{user}.wordspress.com')        
        
        
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        
        try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
            (content, referer, osurl, titleurl, pagestatus) = ('','','','','') # too many false positives
        except socket.error as ex:
            print(f'{color_red}{ex}{color_reset}')

        if 'That page can' not in content:
        # if 'Do you want to register' not in content:
            titleUrl = titleurl.replace("'s Profile | WordPress.org","").strip()
            fullname = titleurl
            fullname = fullname.split(" (")[0]
            print(f'{color_green}{url}{color_yellow}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["note"] = note
            row_data["user"] = user
            row_data["note"] = note
            data.append(row_data)


def wordpress_profiles(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< wordpress {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '8 - wordpress')


        (Success, fullname, lastname, firstname, case, SEX) = ('','','','','','')
        (photo, country, website, email, language, username) = ('','','','','','')
        (city) = ('')
        user = user.rstrip()
        url = (f'https://profiles.wordpress.org/{user}/')
        
        try:
            # (content, referer, osurl, titleurl, pagestatus) = request(url)
            (content, referer, osurl, titleurl, pagestatus) = ('','','','','')
        except:
            pass
        if '404' not in pagestatus:
            fullname = titleurl
            fullname = fullname.split(" (")[0]
            if fullname.lower() in titleurl.lower():
                (fullname, titleurl) = ('', '')
            
            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["fullname"] = fullname
            row_data["url"] = url

            data.append(row_data)



def wordpresssearchemail():    # testuser=    kevinrose@gmail.com 
    print(f'{color_yellow}\n\t<<<<< wordpressemail {color_blue}emails{color_yellow} >>>>>{color_reset}')    
    
    for email in emails:
        row_data = {}
        (query, content, note) = (email, '', '')
        
        url = (f'http://en.search.wordpress.com/?q={email}')

        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
        except:
            (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '')
            pass
        if 'No sites found' not in content:
            content = content.split('\n') 
            for eachline in content:

                if eachline == "": pass                                             # skip blank lines
                else:
                    if 'post-title' in eachline:
                        # print(eachline) # temp
                        eachline = eachline.split('\"')
                        url = eachline[3]
                        ranking = '9 - wordpress'
                        
                        row_data["query"] = query
                        row_data["ranking"] = ranking
                        row_data["url"] = url
                        row_data["email"] = email                        
                        print(f'{color_green}{url}{color_reset}')    

def write_blurb():
    '''
        read Intel.xlsx and write OSINT_.docx to describe what you found.
    '''
    fullname_column_index = ''
    docx_file = "OSINT__DRAFT_V1.docx"

    message = (f'Writing blurb from {input_xlsx}')
    message_square(message, color_green)

    if not os.path.exists(input_xlsx):
        input(f"{color_red}{input_xlsx} doesnt exist.{color_reset}")
        sys.exit()
    elif os.path.getsize(input_xlsx) == 0:
        input(f'{color_red}{input_xlsx} is empty. Fill it with intel you found.{color_reset}')
        sys.exit()
    elif os.path.isfile(input_xlsx):
        message = (f'Reading {input_xlsx} for blurb')
        message_square(message, color_green)
       
        # write_blurb(input_xlsx, docx_file)
    else:
        input(f'{color_red}See {input_xlsx} does not exist. Hit Enter to exit...{color_reset}')
        sys.exit()


    # Open the Excel file
    wb = openpyxl.load_workbook(input_xlsx)
    sheet = wb.active
    
    # Find the column headers
    header_row = sheet[1]
    column_names = [cell.value for cell in header_row]    
    
    # Columns to skip
    columns_to_skip = ["query", "ranking", "content", "referer", "osurl", "titleurl", "pagestatus", "city", "state", "firstname", "lastname", "Latitude", "Longitude", "Coordinate", "original_file", "Icon"]

    
    for idx, cell in enumerate(header_row, start=1):
        if cell.value == "fullname":
            fullname_column_index = idx
            # print(f'Fullname: {fullname_column_index}')
            break

    if fullname_column_index is None:
        print("Fullname column not found in the Excel file.")
        return

    # Create a new Word document
    doc = Document()
 
    sentence = (f'An open-source search revealed the following details.\n\n')
    print(f'{sentence}')  
    doc.add_paragraph(sentence)    
    # Loop through rows in the Excel file and write to Word document
    for row in sheet.iter_rows(min_row=2, values_only=True):
        sentence = "\n".join(f"{column}: {value}" for column, value in zip(column_names, row) if column not in columns_to_skip and value is not None)
        doc.add_paragraph(sentence)
        doc.add_paragraph("")  # Add an empty line between rows


    # Save the Word document
    doc.save(docx_file)

    message = (f'Data written to {docx_file}')
    message_square(message, color_green)


    

def write_intel(data):
    '''
    The write_locations() function receives the processed data as a list of 
    dictionaries and writes it to a new Excel file using openpyxl. 
    It defines the column headers, sets column widths, and then iterates 
    through each row of data, writing it into the Excel worksheet.
    '''
    message = (f'Writing {output_xlsx}')
    message_square(message, color_green)

    try:
        data = sorted(data, key=lambda x: (x.get("ranking", ""), x.get("fullname", ""), x.get("query", "")))
        print(f'sorted by ranking')
    except TypeError as error:

        print(f'{color_red}{error}{color_reset}')

    global workbook
    workbook = Workbook()
    global worksheet
    worksheet = workbook.active

    worksheet.title = 'Intel'
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'B2'  # Freeze cells
    worksheet.selection = 'B2'

    log_headers = [
        "Date", "Subject", "Requesting Agency", "Requesting Agent", "Case"
        , "Summary of Findings", "Source", "Notes"
    ]


    # Write headers to the first row
    for col_index, header in enumerate(headers_intel):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [3, 4, 5, 6, 49, 50]: 
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # orange?
            cell.fill = fill
        elif col_index in [7,8, 13, 14, 15, 29, 30, 35, 36, 37, 38, 39, 40, 41, 42, 43]:  # yellow headers
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Use yellow color
            cell.fill = fill
        # elif col_index == 27:  # Red for column 27
            # fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Red color
            # cell.fill = fill




    ## Excel column width

    worksheet.column_dimensions['A'].width = 15 # query
    worksheet.column_dimensions['B'].width = 20 # ranking
    worksheet.column_dimensions['C'].width = 20 # fullname
    worksheet.column_dimensions['D'].width = 25 # url
    worksheet.column_dimensions['E'].width = 25 # email
    worksheet.column_dimensions['F'].width = 15 # user
    worksheet.column_dimensions['G'].width = 14 # phone
    worksheet.column_dimensions['H'].width = 16 # business
    worksheet.column_dimensions['I'].width = 24 # fulladdress
    worksheet.column_dimensions['J'].width = 12 # city
    worksheet.column_dimensions['K'].width = 10 # state
    worksheet.column_dimensions['L'].width = 8 # country
    worksheet.column_dimensions['M'].width = 20 # note
    worksheet.column_dimensions['N'].width = 14 # AKA
    worksheet.column_dimensions['O'].width = 11 # DOB
    worksheet.column_dimensions['P'].width = 5 # SEX
    worksheet.column_dimensions['Q'].width = 20 # info
    worksheet.column_dimensions['R'].width = 20 # misc
    worksheet.column_dimensions['S'].width = 10 # firstname
    worksheet.column_dimensions['T'].width = 11 # middlename
    worksheet.column_dimensions['U'].width = 10 # lastname
    worksheet.column_dimensions['V'].width = 10 # associates
    worksheet.column_dimensions['W'].width = 10 # case
    worksheet.column_dimensions['X'].width = 13 # sosfilenumber
    worksheet.column_dimensions['Y'].width = 10 # owner
    worksheet.column_dimensions['Z'].width = 10 # president
    worksheet.column_dimensions['AA'].width = 10 # sosagent
    worksheet.column_dimensions['AB'].width = 10 # managers
    worksheet.column_dimensions['AC'].width = 15 # Time
    worksheet.column_dimensions['AD'].width = 12 # Latitude
    worksheet.column_dimensions['AE'].width = 12 # Longitude
    worksheet.column_dimensions['AF'].width = 22 # Coordinate
    worksheet.column_dimensions['AG'].width = 12 # original_file
    worksheet.column_dimensions['AH'].width = 12 # Source
    worksheet.column_dimensions['AI'].width = 12 # Source file information
    worksheet.column_dimensions['AJ'].width = 10 # Plate
    worksheet.column_dimensions['AK'].width = 10 # VIS
    worksheet.column_dimensions['AL'].width = 10 # VIN
    worksheet.column_dimensions['AM'].width = 10 # VYR
    worksheet.column_dimensions['AN'].width = 10 # VMA
    worksheet.column_dimensions['AO'].width = 10 # LIC
    worksheet.column_dimensions['AP'].width = 10 # LIY
    worksheet.column_dimensions['AQ'].width = 10 # DLN
    worksheet.column_dimensions['AR'].width = 10 # DLS
    worksheet.column_dimensions['AS'].width = 10 # content
    worksheet.column_dimensions['AT'].width = 10 # referer
    worksheet.column_dimensions['AU'].width = 10 # osurl
    worksheet.column_dimensions['AV'].width = 10 # titleurl
    worksheet.column_dimensions['AW'].width = 12 # pagestatus
    worksheet.column_dimensions['AX'].width = 16 # ip
    worksheet.column_dimensions['AY'].width = 15 # dnsdomain

    for i in range(len(data)):
        if data[i] is None:
            data[i] = ''

    for row_index, row_data in enumerate(data):

        for col_index, col_name in enumerate(headers_intel):
            try:
                cell_data = row_data.get(col_name)
                worksheet.cell(row=row_index+2, column=col_index+1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error printing line: {str(e)}{color_reset}")

    # Create a new worksheet for color codes
    color_worksheet = workbook.create_sheet(title='ColorCode')
    color_worksheet.freeze_panes = 'B2'  # Freeze cells

    # Excel column width
    color_worksheet.column_dimensions['A'].width = 14# Color
    color_worksheet.column_dimensions['B'].width = 20# Description


    # Excel row height
    color_worksheet.row_dimensions[2].height = 22  # Adjust the height as needed
    color_worksheet.row_dimensions[3].height = 22
    color_worksheet.row_dimensions[4].height = 23
    color_worksheet.row_dimensions[5].height = 23
    color_worksheet.row_dimensions[6].height = 40   # truck

    color_worksheet.cell(row=1, column=1).value = 'Color'
    color_worksheet.cell(row=1, column=2).value = 'description'
    color_worksheet.cell(row=2, column=1).value = 'Red'
    color_worksheet.cell(row=3, column=1).value = 'Orange'
    color_worksheet.cell(row=4, column=1).value = 'Green'
    color_worksheet.cell(row=5, column=1).value = 'Yellow'

    color_worksheet.cell(row=7, column=1).value = 'ABBREVIATIONS'
    color_worksheet.cell(row=8, column=1).value = 'AKA'
    color_worksheet.cell(row=9, column=1).value = 'DOB'
    color_worksheet.cell(row=10, column=1).value = 'VIS'
    color_worksheet.cell(row=11, column=1).value = 'VIN'
    color_worksheet.cell(row=12, column=1).value = 'VYR'
    color_worksheet.cell(row=13, column=1).value = 'VMA'
    color_worksheet.cell(row=14, column=1).value = 'LIC'
    color_worksheet.cell(row=15, column=1).value = 'LIY'
    color_worksheet.cell(row=16, column=1).value = 'DLN'
    color_worksheet.cell(row=17, column=1).value = 'DLS'

       
    color_worksheet.cell(row=2, column=2).value = 'Bad Intel or dead link'
    color_worksheet.cell(row=3, column=2).value = 'Research'
    color_worksheet.cell(row=4, column=2).value = 'Good Intel'
    color_worksheet.cell(row=5, column=2).value = 'Highlighted'

    color_worksheet.cell(row=8, column=2).value = 'Also Known As (Alias)'
    color_worksheet.cell(row=9, column=2).value = 'Date of Birth'
    color_worksheet.cell(row=10, column=2).value = 'Vehicle State'
    color_worksheet.cell(row=11, column=2).value = 'Vehicle Identification Number'
    color_worksheet.cell(row=12, column=2).value = 'Vehicle Year'
    color_worksheet.cell(row=13, column=2).value = 'Vehicle Make'
    color_worksheet.cell(row=14, column=2).value = 'License'
    color_worksheet.cell(row=15, column=2).value = 'License Year'
    color_worksheet.cell(row=16, column=2).value = 'Drivers License Number'
    color_worksheet.cell(row=17, column=2).value = 'Drivers License State'


    # colored fills
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
    green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')


    # Apply the orange fill to the cell in row 2, column 2
    color_worksheet.cell(row=2, column=2).fill = red_fill
    color_worksheet.cell(row=3, column=2).fill = orange_fill
    color_worksheet.cell(row=4, column=2).fill = green_fill
    color_worksheet.cell(row=5, column=2).fill = yellow_fill


    # Create a new worksheet for logs
    log_worksheet = workbook.create_sheet(title='Log')
    log_worksheet.freeze_panes = 'B2'  # Freeze cells

# Date, Subject, Requesting Agency, Requesting Agent, Case, Summary of Findings, Source, Notes, Requestor

    # Excel column width
    log_worksheet.column_dimensions['A'].width = 14# Date
    log_worksheet.column_dimensions['B'].width = 20# Subject
    log_worksheet.column_dimensions['C'].width = 24# Requesting Agency
    log_worksheet.column_dimensions['D'].width = 20# Requesting Agent
    log_worksheet.column_dimensions['E'].width = 14# Case
    log_worksheet.column_dimensions['F'].width = 20# Summary of Findings
    log_worksheet.column_dimensions['G'].width = 14# Source
    log_worksheet.column_dimensions['H'].width = 25# Notes

    log_worksheet.cell(row=1, column=1).value = 'Date'
    log_worksheet.cell(row=1, column=2).value = 'Subject'
    log_worksheet.cell(row=1, column=3).value = 'Requesting Agency'
    log_worksheet.cell(row=1, column=4).value = 'Requesting Agent'
    log_worksheet.cell(row=1, column=5).value = 'Case'
    log_worksheet.cell(row=1, column=6).value = 'Summary of Findings'
    log_worksheet.cell(row=1, column=7).value = 'Notes'



    workbook.save(output_xlsx)

# Save the workbook
# wb.save('output.xlsx')

def write_intel_basic(data, output_xlsx):
    message = (f'Writing intel to {output_xlsx}')
    message_square(message, color_green)

    wb = openpyxl.Workbook()
    ws = wb.active

    # ws.append(headers)  # Writing headers
    ws.append(headers_intel)  # Writing headers


    for row_data in data:
        row = [row_data.get(header, '') for header in headers_intel]
        ws.append(row)

    wb.save(output_xlsx)

def write_locations_basic(data, output_xlsx):
    message = (f'Writing locations to {output_xlsx}')
    message_square(message, color_green)

    wb = openpyxl.Workbook()
    ws = wb.active

    ws.append(headers_locations)  # Writing location headers

    for row_data in data:
        row = [row_data.get(header_l, '') for header_l in headers_locations]
        ws.append(row)

    wb.save(output_xlsx)

def write_locations(data):
    '''
    The write_locations() function receives the processed data as a list of 
    dictionaries and writes it to a new Excel file using openpyxl. 
    It defines the column headers, sets column widths, and then iterates 
    through each row of data, writing it into the Excel worksheet.
    '''
    message = (f'Writing locations to {output_xlsx}')
    message_square(message, color_green)

    global workbook
    workbook = Workbook()
    global worksheet
    worksheet = workbook.active

    worksheet.title = 'Locations'
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'B2'  # Freeze cells
    worksheet.selection = 'B2'

    headers_locations = [
        "#", "Time", "Latitude", "Longitude", "Address", "Group", "Subgroup"
        , "Description", "Type", "Source", "Deleted", "Tag", "Source file information"
        , "Service Identifier", "Carved", "Name", "business", "number", "street"
        , "city", "county", "state", "zipcode", "country", "fulladdress", "query"
        , "Sighting State", "Plate", "Capture Time", "Capture Network", "Highway Name"
        , "Coordinate", "Capture Location Latitude", "Capture Location Longitude"
        , "Container", "Sighting Location", "Direction", "Time Local", "End time"
        , "Category", "Manually decoded", "Account", "PlusCode", "Time Original", "Timezone"
        , "Icon", "original_file", "case", "Index"

    ]

    # Write headers to the first row
    for col_index, header in enumerate(headers_locations):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [2, 3, 4]: 
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # orange?
            cell.fill = fill
        elif col_index in [1, 5, 6, 7, 8, 9, 15, 16, 24, 30, 31, 36, 38]:  # yellow headers
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Use yellow color
            cell.fill = fill
        elif col_index == 27:  # Red for column 27
            fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Red color
            cell.fill = fill
    try:
        ## Excel column width
        worksheet.column_dimensions['A'].width = 8# #
        worksheet.column_dimensions['B'].width = 19# Time
        worksheet.column_dimensions['C'].width = 18# Latitude
        worksheet.column_dimensions['D'].width = 18# Longitude
        worksheet.column_dimensions['E'].width = 45# Address
        worksheet.column_dimensions['F'].width = 14# Group
        worksheet.column_dimensions['G'].width = 13# Subgroup
        worksheet.column_dimensions['H'].width = 17# Description
        worksheet.column_dimensions['I'].width = 9# Type
        worksheet.column_dimensions['J'].width = 10# Source
        worksheet.column_dimensions['K'].width = 10# Deleted
        worksheet.column_dimensions['L'].width = 11# Tag
        worksheet.column_dimensions['M'].width = 20# Source file information
        worksheet.column_dimensions['N'].width = 15# Service Identifier
        worksheet.column_dimensions['O'].width = 7# Carved
        worksheet.column_dimensions['P'].width = 15# Name
        
        ## bonus
        worksheet.column_dimensions['Q'].width = 20# business 
        worksheet.column_dimensions['R'].width = 10# number
        worksheet.column_dimensions['S'].width = 20# street 
        worksheet.column_dimensions['T'].width = 15# city   
        worksheet.column_dimensions['Y'].width = 25# county    
        worksheet.column_dimensions['V'].width = 12# state   
        worksheet.column_dimensions['W'].width = 8# zipcode     
        worksheet.column_dimensions['X'].width = 6# country    
        worksheet.column_dimensions['Y'].width = 26# FullAddress   
        worksheet.column_dimensions['Z'].width = 26# query

        ##  Flock
        worksheet.column_dimensions['AA'].width = 11# Sighting State
        worksheet.column_dimensions['AB'].width = 11# Plate
        worksheet.column_dimensions['AC'].width = 22# Capture Time
        worksheet.column_dimensions['AD'].width = 15# Capture Network
        worksheet.column_dimensions['AE'].width = 21# Highway Name
        worksheet.column_dimensions['AF'].width = 30# Coordinate
        worksheet.column_dimensions['AG'].width = 20# Capture Location Latitude
        worksheet.column_dimensions['AH'].width = 20# Capture Location Longitude

        ##
        worksheet.column_dimensions['AI'].width = 10# Container
        worksheet.column_dimensions['AJ'].width = 14# Sighting Location
        worksheet.column_dimensions['AK'].width = 10# Direction
        worksheet.column_dimensions['AL'].width = 11# Time Local
        worksheet.column_dimensions['AM'].width = 25# End time
        worksheet.column_dimensions['AN'].width = 10# Category
        worksheet.column_dimensions['AO'].width = 18# Manually decoded
        worksheet.column_dimensions['AP'].width = 10# Account
        worksheet.column_dimensions['AQ'].width = 25 # PlusCode
        worksheet.column_dimensions['AR'].width = 21 # Time Original
        worksheet.column_dimensions['AS'].width = 9 # Timezone
        worksheet.column_dimensions['AT'].width = 10 # Icon   
        worksheet.column_dimensions['AU'].width = 20 # original_file
        worksheet.column_dimensions['AV'].width = 10 # case
        worksheet.column_dimensions['AW'].width = 6 # Index
    except:pass
    
    for row_index, row_data in enumerate(data):

        for col_index, col_name in enumerate(headers_locations):
            cell_data = row_data.get(col_name)
            try:
                worksheet.cell(row=row_index+2, column=col_index+1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error printing line: {str(e)}{color_reset}")


    # Create a new worksheet for color codes
    color_worksheet = workbook.create_sheet(title='Icons')
    color_worksheet.freeze_panes = 'B2'  # Freeze cells

    # Excel column width
    color_worksheet.column_dimensions['A'].width = 8# Icon sample
    color_worksheet.column_dimensions['B'].width = 9# Name
    color_worksheet.column_dimensions['C'].width = 29# Description

    # Excel row height
    color_worksheet.row_dimensions[2].height = 22  # Adjust the height as needed
    color_worksheet.row_dimensions[3].height = 22
    color_worksheet.row_dimensions[4].height = 23
    color_worksheet.row_dimensions[5].height = 23
    color_worksheet.row_dimensions[6].height = 40   # truck
    color_worksheet.row_dimensions[7].height = 6
    color_worksheet.row_dimensions[8].height = 24
    color_worksheet.row_dimensions[9].height = 22
    color_worksheet.row_dimensions[10].height = 22
    color_worksheet.row_dimensions[11].height = 22
    color_worksheet.row_dimensions[12].height = 23
    color_worksheet.row_dimensions[13].height = 23
    color_worksheet.row_dimensions[14].height = 25
    color_worksheet.row_dimensions[15].height = 25
    color_worksheet.row_dimensions[16].height = 23
    color_worksheet.row_dimensions[17].height = 6
    color_worksheet.row_dimensions[18].height = 38
    color_worksheet.row_dimensions[19].height = 38
    color_worksheet.row_dimensions[20].height = 38
    color_worksheet.row_dimensions[21].height = 38
    color_worksheet.row_dimensions[22].height = 38
    color_worksheet.row_dimensions[23].height = 6
    color_worksheet.row_dimensions[24].height = 15
    color_worksheet.row_dimensions[25].height = 6
    color_worksheet.row_dimensions[26].height = 15


    
    # Define color codes
    color_worksheet['A1'] = ' '
    color_worksheet['B1'] = 'Icon'
    color_worksheet['C1'] = 'Icon Description'

    icon_data = [

        ('', 'Car', 'Lpr red car (License Plate Reader)'),
        ('', 'Car2', 'Lpr yellow car'),
        ('', 'Car3', 'Lpr greeen car with circle'),
        ('', 'Car4', 'Lpr red car with circle'),
        ('', 'Truck', 'Lpr truck'),         
        ('', '', ''),
        ('', 'Calendar', 'Calendar'), 
        ('', 'Home', 'Home'),                
        ('', 'Images', 'Photo'),
        ('', 'Intel', 'I'),  
        ('', 'Locations', 'Reticle'),  
        ('', 'default', 'Yellow flag'),  
        ('', 'Office', 'Office'),         
        ('', 'Searched', 'Searched Item'),          
        ('', 'Videos', 'Video clip'),        
        ('', '', ''),
        ('', 'Toll', 'Blue square'), 
        ('', 'N', 'Northbound blue arrow'),
        ('', 'E', 'Eastbound blue arrow'),
        ('', 'S', 'Southbound blue arrow'),
        ('', 'W', 'Westbound blue arrow'),
        ('', '', ''),
        ('', 'Yellow font', 'Tagged'),
        ('', 'Chats', 'Chats'),   # 


        ('', '', ''),
        ('', 'NOTE', 'visit https://earth.google.com/ <file><Import KML> select gps.kml <open>'),
    ]

    for row_index, (icon, tag, description) in enumerate(icon_data):
        color_worksheet.cell(row=row_index + 2, column=1).value = icon
        color_worksheet.cell(row=row_index + 2, column=2).value = tag
        color_worksheet.cell(row=row_index + 2, column=3).value = description

    car_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon15.png'   # red car
    car2_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon47.png'  # yellow car
    car3_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon54.png'  # green car with circle
    car4_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon7.png'  # red car with circle
    truck_icon = 'https://maps.google.com/mapfiles/kml/shapes/truck.png'    # blue truck
    default_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon13.png'   # yellow flag
    calendar_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon23.png' # paper
    chat_icon = 'https://maps.google.com/mapfiles/kml/shapes/post_office.png' # email
    locations_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon28.png'    # yellow paddle
    home_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon56.png'
    images_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon46.png'
    intel_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon44.png'
    office_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon21.png'
    searched_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon0.png'  #  
    toll_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-none.png'
    videos_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon30.png'
    n_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-0.png'
    e_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-4.png'
    s_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-8.png'
    w_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-12.png'

    try:
        # Insert graphic from URL into cell of color_worksheet

        response = requests.get(car_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A2')

        response = requests.get(car2_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A3')

        response = requests.get(car3_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A4')

        response = requests.get(car4_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A5')

        response = requests.get(truck_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A6')

        response = requests.get(calendar_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A8')
        
        response = requests.get(home_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A9')

        response = requests.get(images_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A10')

        response = requests.get(intel_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A11')

        response = requests.get(locations_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A12')

        response = requests.get(default_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A13')

        response = requests.get(office_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A14')        

        response = requests.get(searched_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A15')

        response = requests.get(videos_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A16')
        
        response = requests.get(toll_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A18')

        response = requests.get(n_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A19')

        response = requests.get(e_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A20')

        response = requests.get(s_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A21')

        response = requests.get(w_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A22')

        response = requests.get(chat_icon)
        img = Image(io.BytesIO(response.content))
        color_worksheet.add_image(img, 'A24')     
        
    except:
        pass

    
    workbook.save(output_xlsx)




def youtube(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< youtube {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '4 - youtube')

        (fullname) = ('')
        user = user.rstrip()
        url = (f'https://www.youtube.com/{user}')
        (content, referer, osurl, titleurl, pagestatus) = request(url)
        titleurl = titleurl.replace(' - YouTube','')
        if '404' not in pagestatus:
            fullname = titleurl
            
            if fullname.lower() == user.lower():
                fullname = ''
            if ' ' in fullname:
                (fullname, firstname, lastname, middlename) = fullname_parse(fullname)
            else:
                (fullname, firstname, lastname, middlename) = ('', '', '', '')


            print(f'{color_green}{url}{color_yellow}	{fullname}{color_reset}')

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname

            data.append(row_data)


def robtex():
    print(f'{color_yellow}\n\t<<<<<robtex dns lookup >>>>>{color_reset}')    

    for website in websites:    
        row_data = {}
        (query, ranking) = (website, '9 - robtexDNS-lookup')

        (final_url, dnsdomain, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '', '')
        (otherurl, ip) = ('', '')
        url = website
        url = url.replace("http://", "https://")
        if "http" not in url.lower():
            url = (f'https://{url}')
        
        dnsdomain = url.lower()
        dnsdomain = dnsdomain.replace("https://", "")
        dnsdomain = dnsdomain.replace("http://", "")
        dnsdomain = dnsdomain.split('/')[0]

        ip = ip_address(dnsdomain)

        note = url  

        url = (f'https://www.robtex.com/dns-lookup/{dnsdomain}#quick')
        
        if 1==1:
        # if dnsdomain not in final_url:
            print(f'{color_green}{website}{color_yellow}	{ip}{color_reset}')

            row_data["query"] = query
            row_data["ranking"] = ranking
            # row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["ip"] = ip            
          
            row_data["note"] = note            
            row_data["dnsdomain"] = dnsdomain    
            row_data["referer"] = referer   
            row_data["osurl"] = osurl              
            row_data["titleurl"] = titleurl              
            row_data["pagestatus"] = pagestatus              

            data.append(row_data)    


def titles():    # testsite= google.com
    from subprocess import call, Popen, PIPE
    print(f'{color_yellow}\n\t<<<<< Titles grab {color_blue}Website\'s{color_yellow} >>>>>{color_reset}')
    for website in websites:    
        row_data = {}
        (query, ranking) = (website, '7 - website')

        (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '')
        (ip, note) = ('', '')
        
        url = website

        fake_referer = 'https://www.google.com/'
        headers_url = {'Referer': fake_referer}
        url = url.replace("http://", "https://")     # test
        if "http" not in url.lower():
            url = (f'https://{url}')

        url = url.replace("https://", "http://")

        try:
            (content, referer, osurl, titleurl, pagestatus) = request(url)
        except TypeError as error:
            print(f'{color_red}{error}{color_reset}')
        
        # dnsdomain
        dnsdomain = url.lower()
        dnsdomain = dnsdomain.replace("https://", "")
        dnsdomain = dnsdomain.replace("http://", "")
        dnsdomain = dnsdomain.split('/')[0]
        
        # ip
        ip = ip_address(dnsdomain)
        print(f'{color_green}{website}{color_yellow}	   {pagestatus}	{color_blue}{titleurl}{color_reset}')
        if 1==1:

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["ip"] = ip            
          
            row_data["note"] = note            
            row_data["dnsdomain"] = dnsdomain    
            row_data["referer"] = referer   
            row_data["osurl"] = osurl              
            row_data["titleurl"] = titleurl              
            row_data["pagestatus"] = pagestatus              

            data.append(row_data) 

def validnumber():# testPhone= 
    print(f'{color_yellow}\n\t<<<<< validnumber {color_blue}phone numbers{color_yellow} >>>>>{color_reset}')

    # https://validnumber.com/phone-number/3124377966/
    for phone in phones:
        row_data = {}
        (query, ranking) = (phone, '9 - validnumber')
        (country, city, state, zipcode, case, note) = ('', '', '', '', '', '')
        (content, referer, osurl, titleurl, pagestatus)  = ('', '', '', '', '')
        (query) = (phone)
 
        url = ('https://validnumber.com/phone-number/%s/' %(phone.lstrip("1")))
        
        (content, referer, osurl, titleurl, pagestatus) = request(url)    # protected by cloudflare
        pagestatus = ''
        for eachline in content.split("\n"):
            if "No name associated with this number" in eachline and case == '':
                print(f'{color_red}not found{color_reset}')  # temp
                
                # url = ('')
            elif "Find out who owns" in eachline:
                if 'This device is registered in ' in eachline:
                    ranking = '5 - validnumber'
                    note = eachline.split('\"')[1]
                    note = note.split('Free owner details for')[0]
                    city = eachline.split("This device is registered in ")[1].split("Free owner details")[0]
                    state = city.split(',')[1].strip()
                    city = city.split(',')[0]
        if city != '':        
            city = city.title()
            if city == "Directory Assistance":
                ranking = '8 - validnumber'
        state = state.replace('..','').replace('Illinois','IL')
        if state == '':
            state = phone_state_check(phone, state).replace('?', '')

        if url != '':        

            print(f'{color_green}{url}{color_reset}') 

            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["url"] = url
            row_data["phone"] = phone            
            row_data["note"] = note            
            row_data["state"] = state    
            row_data["city"] = city   

            data.append(row_data) 


def venvmo(): # testuser = kevinrose
    print(f'{color_yellow}\n\t<<<<< venmo {color_blue}users{color_yellow} >>>>>{color_reset}')
    for user in users:    
        row_data = {}
        (query, ranking) = (user, '9 - venmo')

        (fullname) = ('')
        user = user.rstrip()
        url = (f'https://account.venmo.com/u/{user}')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)
        # titleurl = titleurl.replace(' - venmo','')
        # if '404' not in pagestatus:
        if 1 ==1:
            row_data["query"] = query
            row_data["ranking"] = ranking
            row_data["user"] = user
            row_data["url"] = url
            data.append(row_data)
            

def viewdnsdomain():
    print(f'{color_yellow}\n\t<<<<<viewdns lookup >>>>>{color_reset}')    

    for website in websites:    
        row_data = {}
        (query, ranking) = (website, '9 - viewdns')

        (final_url, dnsdomain, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '', '')
        (content, referer, osurl, titleurl, pagestatus) = ('', '', '', '', '')
        (otherurl, ip) = ('', '')
        url = website
        url = url.replace("http://", "https://")
        if "http" not in url.lower():
            url = (f'https://{url}')
        
        dnsdomain = url.lower()
        dnsdomain = dnsdomain.replace("https://", "")
        dnsdomain = dnsdomain.replace("http://", "")
        dnsdomain = dnsdomain.split('/')[0]

        ip = ip_address(dnsdomain)

        note = url  

        url = (f'https://viewdns.info/whois/?domain={dnsdomain}')
        # (content, referer, osurl, titleurl, pagestatus) = request(url)
        # time.sleep(10) #will sleep for 10 seconds
        if 1==1:
        # if dnsdomain not in final_url:
            print(f'{color_green}{website}{color_yellow}	{ip}{color_reset}')

            row_data["query"] = query
            row_data["ranking"] = ranking
            # row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["ip"] = ip            
          
            row_data["note"] = note            
            row_data["dnsdomain"] = dnsdomain    
            row_data["referer"] = referer   
            row_data["osurl"] = osurl              
            row_data["titleurl"] = titleurl              
            row_data["pagestatus"] = pagestatus              

            data.append(row_data)    


def whoiswebsite():    # testsite= google.com
    from subprocess import call, Popen, PIPE
    print(f'{color_yellow}\n\t<<<<< whois {color_blue}Website\'s{color_yellow} >>>>>{color_reset}')

    for dnsdomain in dnsdomains:    
        row_data = {}
        (query, ranking) = (dnsdomain, '7 - whois')

        # query = website
        # website = website.replace('http://','')
        # website = website.replace('https://','')
        # website = website.replace('www.','')
        
        url = (f'https://www.ip-adress.com/website/{dnsdomain}')
        url2 = ('https://whois.domaintools.com/%s' %(dnsdomain.replace('www.','')))
        (email,phone,fullname,country,city,state) = ('','','','','','')
        (city, country, zipcode, state, ip) = ('', '', '', '', '')
        (content, titleurl, pagestatus) = ('', '', '')
        (email, phone, fullname, entity, fulladdress) = ('', '', '', '', '') 

        if sys.platform == 'win32' or sys.platform == 'win64':    
            print(f'skipping whois query from windows')  # temp

            row_data["query"] = query
            row_data["ranking"] = '9 - whois.domaintools.com'
            row_data["fullname"] = fullname
            row_data["url"] = url
            row_data["city"] = city            
            row_data["country"] = country            
            row_data["state"] = state            
            row_data["dnsdomain"] = dnsdomain            

            data.append(row_data) 

        elif dnsdomain.endswith('.com') or dnsdomain.endswith('.edu') or dnsdomain.endswith('.net'):
            WhoisArgs = (f'whois {dnsdomain}')
            response= Popen(WhoisArgs, shell=True, stdout=PIPE)
            for line in response.stdout:
                line = line.decode("utf-8")
                if ':' in line and "# " not in line and len(line) > 2:
                    line = line.strip()
                    content = (f'{content}\n{line}')
                if email == '':
                    if line.startswith('RAbuseEmail:'):
                        try:
                            email = (line.split(': ')[1].lstrip())
                        except:pass    
                    elif line.lower().startswith('abuse-mailbox:'):email = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('orgabuseemail:'):email = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('Orgtechemail:'):email = (line.split(': ')[1].lstrip())
                
                if phone == '':
                    if line.lower().startswith('rabusephone:'):phone = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('orgabusephone:'):phone = (line.split(': ')[1].lstrip())
                    elif line.lower().startswith('phone:'):phone = (line.split(': ')[1].lstrip())
                    phone = phone.replace("+", "")
                if line.lower().startswith('rtechname:'):fullname = (line.split(': ')[1].lstrip())
                elif line.lower().startswith('person:'):fullname = (line.split(': ')[1].lstrip())                
                
                if line.lower().startswith('country:'):country = (line.split(': ')[1].lstrip())
                if line.lower().startswith('city:'):city = (line.split(': ')[1].lstrip())
                if line.lower().startswith('address:'):fulladdress = ('%s %s' %(fulladdress, line.split(': ')[1].lstrip()))
                if line.lower().startswith('stateprov:'):state = (line.split(': ')[1].lstrip())
                if line.lower().startswith('postalcode:'):zipcode = (line.split(': ')[1].lstrip())
                if line.lower().startswith('orgname:'):entity = (line.split(': ')[1].lstrip())
                elif line.lower().startswith('org-name:'):entity = (line.split(': ')[1].lstrip())

            print(f'{color_green}"whois {dnsdomain}{color_yellow}	   {email}	{color_blue}{phone}{color_reset}')

            row_data["query"] = query
            row_data["ranking"] = '7 - whois'
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
         
            row_data["titleurl"] = titleurl            
            row_data["city"] = city            
            row_data["country"] = country            
            row_data["note"] = note            
            row_data["state"] = state            
            row_data["SEX"] = SEX            
            row_data["zipcode"] = zipcode            
            row_data["dnsdomain"] = dnsdomain            
            row_data["titleurl"] = titleurl            
            row_data["pagestatus"] = pagestatus            


            data.append(row_data) 


        else:
            print(f'{color_red}{dnsdomain} not an edu net or edu site?{color_reset}')

            row_data["query"] = query
            row_data["ranking"] = '7 - whois'
            row_data["fullname"] = fullname
            row_data["firstname"] = firstname
            row_data["lastname"] = lastname
            row_data["url"] = url
         
            row_data["titleurl"] = titleurl            
            row_data["city"] = city            
            row_data["country"] = country            
            row_data["note"] = note            
            row_data["state"] = state            
            row_data["SEX"] = SEX            
            row_data["zipcode"] = zipcode            
            row_data["dnsdomain"] = dnsdomain            
            row_data["titleurl"] = titleurl            
            row_data["pagestatus"] = pagestatus            


            data.append(row_data) 

        time.sleep(7)



def usage():
    '''
        Prints out examples of syntax
    '''
    file = sys.argv[0].split('\\')[-1]
    print(f'\nDescription2: {color_green}{description2}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f'\n    {color_yellow}insert your input into input.txt')
    print(f'\nExample:')
    print(f'    {file} -b -I Intel_test.xlsx')
    print(f'    {file} -B -O Intel_.xlsx')    
    print(f'    {file} -c -I Intel_test.xlsx    # alpha')
    print(f'    {file} -E')
    print(f'    {file} -i')
    print(f'    {file} -l -O locations_.xlsx -I intel_test.xlsx # alpha')
    print(f'    {file} -t')
    print(f'    {file} -s')
    print(f'    {file} -p')
    print(f'    {file} -U')
    print(f'    {file} -W')
    print(f'    {file} -E -i -p -U -I input.txt')
    print(f'    {file} -E -i -p -U -I Intel_test.xlsx')

if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<<Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
3.0.2 - slack, roblox, ham_radio
3.0.1 - Convert area code to state if it doesn't exist, instantusername
3.0.0 - switched to openpyxl, added log sheet
2.8.7 - fixed regex_phone, skip internet check if it's a virtual machine
2.8.6 - made the .py and .exe version dummy proof. Just double click and it runs
2.7.6 - internet checker, removed -I and -O requirement
2.8.0 - kik
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
https://www.deviantart.com/kevinrose/gallery

implement convert_timestamp()


https://start.me/p/0Pqbdg/osint-500-tools   # reviewed

https://start.me/p/1kJKR9/commandergirl-s-suggestions



# phone : whatsapp, haveibeenpwned, group me, true call, weibo, chime, qq, crickwick,
discord, foursquare, facebook, walkie talkie, apple, marco polo, okru, 


https://www.textnow.com/
https://www.talkatone.com/
https://www.pinger.com/
https://www.tumblr.com/login
https://www.tumblr.com/kevinrose
https://www.tumblr.com/search/kevinrose?src=typed_query

add timestamp to log sheet
currently only reads input.txt. add input.xlsx input.
python identityhunt.py  -E -i -p -U -I Intel_test.xlsx  # works
fix -c convert

populate log sheet with todays date


NAM = last name(comma) first name (space) Middle initial


add ID and photo , phone2 column

tkinter purely gui interface
.replace isn't working in several modules
instagramtwo()
create a new identity_hunt with xlsx and requests instead of urllib2


https://opengovus.com/search?q=kevinrose%2C+LLC
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<     notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
if 1- main, then write a report of your findings.
change the order of columns by modifying headers in the write module

Protected by cookies: dailymotion, linkedin, trello, xboxgamertag, twitch.tv, telegram, tripadvisor
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<     The End        >>>>>>>>>>>>>>>>>>>>>>>>>>