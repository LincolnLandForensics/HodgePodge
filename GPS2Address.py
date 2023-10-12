#!/usr/bin/env python3
# coding: utf-8
'''
read GPS coordinates and convert them to addresses
or
read addresses and convert them to coordinates

'''
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import re
import sys
import time
import openpyxl

import argparse  # for menu system
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill

from geopy.geocoders import Nominatim
geolocator = Nominatim(user_agent="GeoTraxer")

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

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "convert GPS coordinates to addresses or visa versa"
version = '0.9.1'


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    global row
    row = 0  # defines arguments
    # Row = 1  # defines arguments   # if you want to add headers 
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-c', '--create', help='create blank input sheet', required=False, action='store_true')
    
    parser.add_argument('-r', '--read', help='read xlsx', required=False, action='store_true')

    args = parser.parse_args()

    global input_xlsx
    global outuput_xlsx
    
    if not args.input: 
        input_xlsx = "locations.xlsx"        
    else:
        input_xlsx = args.input

    if not args.output: 
        outuput_xlsx = "locations2addresses_.xlsx"        
    else:
        outuput_xlsx = args.output
    if args.create:
        data = []
        print(f'{color_green}Writing to {outuput_xlsx} {color_reset}')
        write_xlsx(data)
    elif args.read:
        file_exists = os.path.exists(input_xlsx)
        if file_exists == True:
            print(f'{color_green}Reading {input_xlsx} {color_reset}')
            
            data = read_xlsx(input_xlsx)
            # write_xlsx(data)
            workbook.close()
            print(f'{color_green}Writing to {outuput_xlsx} {color_reset}')
        else:
            print(f'{color_red}{input_xlsx} does not exist{color_reset}')
            exit()

    else:
        usage()
    
    return 0


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def read_xlsx(input_xlsx):

    """Read data from an xlsx file and return as a list of dictionaries.
    Read XLSX Function: The read_xlsx() function reads data from the input 
    Excel file using the openpyxl library. It extracts headers from the 
    first row and then iterates through the data rows, creating dictionaries 
    for each row with headers as keys and cell values as values.
    
    """
    zip_code_pattern = r'\b\d{5}(?:-\d{4})?\b' 
    
    
    wb = openpyxl.load_workbook(input_xlsx)
    ws = wb.active
    data = []

    # get header values from first row
    global headers
    headers = []
    for cell in ws[1]:
        headers.append(cell.value)

    # get data rows
    for row in ws.iter_rows(min_row=2, values_only=True):
        row_data = {}
        for header, value in zip(headers, row):
            row_data[header] = value
        data.append(row_data)

    for row_index, row_data in enumerate(data):
        (zipcode, business, number, street, city, county) = ('', '', '', '', '', '')
        (state, fulladdress, Latitude, Longitude, query) = ('', '', '', '', '')

        name_data = row_data.get("Name")

        lat_data = row_data.get("Latitude")
        long_data = row_data.get("Longitude")
        address_data = row_data.get("Address")        

        # Check if address_data is None
        if address_data is None:
            print('')

        else:
            # Check the length of address_data
            if len(address_data) > 3:
                
                try:
                    location = geolocator.geocode(address_data)
                except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset}")
                    
                try:
                    fulladdress = location.address
                    row_data["fulladdress"] = fulladdress                

                    # Check if there are exactly 7 commas
                    if fulladdress.count(',') == 7:
                        # Split the address by commas
                        address_parts = fulladdress.split(', ')
                        business = address_parts[0]
                        number = address_parts[1]
                        street = address_parts[2]
                        city = address_parts[3]
                        county = address_parts[4]
                        state = address_parts[5]
                        
                        row_data["business"] = business 
                        row_data["number"] = number 
                        row_data["street"] = street
                        row_data["city"] = city 
                        row_data["county"] = county 
                        row_data["state"] = state 
                    elif fulladdress.count(',') == 5:
                        # Split the address by commas
                        address_parts = fulladdress.split(', ')
                        # business = address_parts[0]
                        # number = address_parts[1]
                        street = address_parts[0]
                        city = address_parts[1]
                        county = address_parts[2]
                        state = address_parts[3]
                        
                        row_data["business"] = business 
                        row_data["number"] = number 
                        row_data["street"] = street
                        row_data["city"] = city 
                        row_data["state"] = state 
                        
                    
                    if "United States" in fulladdress:
                        row_data["country"] = "US" 
                    if ", Illinois," in fulladdress:
                        row_data["state"] = "IL" 
                    elif ", Missouri," in fulladdress:
                        row_data["state"] = "MO"                                  

                    # Find the first ZIP code in the address
                    match = re.search(zip_code_pattern, fulladdress)

                    # Check if a ZIP code was found # task
                    # if match:
                        # zipcode = match.group() # Error : name 'zipzode' is not defined
                        # row_data["zipcode"] = zipzode 
                    # else:
                        # zipcode = ('')
                    
  
                except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset}")                

                # try:
                    # Find the first ZIP code in the address
                    # match = re.search(zip_code_pattern, fulladdress)

                    # Check if a ZIP code was found
                    # if match:
                        # zipcode = match.group()
                        # row_data["zipcode"] = zipzode 
                    # else:
                        # zipcode = ('')                        
                # except Exception as e:
                    # print(f"{color_red}Error : {str(e)}{color_reset}")

                # Sleep for x seconds
                time.sleep(8)

                # if print(location.latitude, location.longitude)
                if long_data is None or lat_data is None:
                    try:
                        row_data["Longitude"] = location.longitude 
                        row_data["Latitude"] = location.latitude
                        query = (f'{location.latitude}, {location.longitude}')
                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 

        if long_data is None or lat_data is None:
            print('')
            # print(f'{color_red}there is no lat or long{color_reset} {lat_data} - {long_data}')
        else:
            print('')
            # print(f'{color_yellow}trying {color_reset} {lat_data} {long_data}')            
            if len(lat_data) > 3:
                if len(fulladdress) < 2:
                    
                    # print(f'{color_yellow}there is no fulladdress{color_reset} but there is lat long {lat_data} - {long_data}')
                    query = (f'{lat_data}, {long_data}')
                    row_data["query"] = query
                    try:
                        location = geolocator.reverse(query)
                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 
                    try:
                        fulladdress = location.address
                        row_data["fulladdress"] = fulladdress                

                        # Check if there are exactly 7 commas
                        if fulladdress.count(',') == 7:
                            # Split the address by commas
                            address_parts = fulladdress.split(', ')
                            business = address_parts[0]
                            number = address_parts[1]
                            street = address_parts[2]
                            city = address_parts[3]
                            county = address_parts[4]
                            state = address_parts[5]
                            
                            row_data["business"] = business 
                            row_data["number"] = number 
                            row_data["street"] = street
                            row_data["city"] = city 
                            row_data["state"] = state 
                        elif fulladdress.count(',') == 5:
                            # Split the address by commas
                            address_parts = fulladdress.split(', ')
                            # business = address_parts[0]
                            # number = address_parts[1]
                            street = address_parts[0]
                            city = address_parts[1]
                            county = address_parts[2]
                            state = address_parts[3]
                            
                            row_data["business"] = business 
                            row_data["number"] = number 
                            row_data["street"] = street
                            row_data["city"] = city 
                            row_data["state"] = state 
                            
                        
                        if "United States" in fulladdress:
                            row_data["country"] = "US" 
                        if ", Illinois," in fulladdress:
                            row_data["state"] = "IL" 
                        elif ", Missouri," in fulladdress:
                            row_data["state"] = "MO"                                  

                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 
                    
                    # try:
                        # Find the first ZIP code in the address
                        # match = re.search(zip_code_pattern, fulladdress)

                        # Check if a ZIP code was found
                        # if match:
                            # zipcode = match.group()
                            # row_data["zipcode"] = zipzode 
                    # except Exception as e:
                        # print(f"{color_red}Error : {str(e)}{color_reset}")
                    
                    # Sleep x seconds
                    time.sleep(8) 

        print(f'{query}\t{color_yellow}{fulladdress}{color_reset}')     
    # workbook.save(outuput_xlsx)

    write_xlsx(data)
    # return data

def write_xlsx(data):
    '''
    The write_xlsx() function receives the processed data as a list of 
    dictionaries and writes it to a new Excel file using openpyxl. 
    It defines the column headers, sets column widths, and then iterates 
    through each row of data, writing it into the Excel worksheet.
    '''

    global workbook
    workbook = Workbook()
    global worksheet
    worksheet = workbook.active

    worksheet.title = 'Sheet1'
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'B2'  # Freeze cells
    worksheet.selection = 'B2'

    # headers = data[0].keys()  # Get the keys (headers) from the first row of data

    headers = [
        "Name", "Description", "Time", "End time", "Category", "Latitude"
        , "Longitude", "Map Address", "Address", "Type", "Source", "Account"
        , "Deleted", "Tag Note", "Source file information", "Carved"
        , "Manually decoded", "business", "number", "street", "city", "county"
        , "state", "zipcode", "country", "fulladdress", "query"
    ]

    # Write headers to the first row
    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [5, 6, 8]: 
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            cell.fill = fill

    # Excel column width
    worksheet.column_dimensions['A'].width = 15# Name
    worksheet.column_dimensions['B'].width = 7# Description
    worksheet.column_dimensions['C'].width = 16# Time
    worksheet.column_dimensions['D'].width = 20# End time
    worksheet.column_dimensions['E'].width = 20# Category
    worksheet.column_dimensions['F'].width = 20# Latitude
    worksheet.column_dimensions['G'].width = 20# Longitude
    worksheet.column_dimensions['H'].width = 20# Map Address
    worksheet.column_dimensions['I'].width = 45# Address
    worksheet.column_dimensions['J'].width = 10# Type
    worksheet.column_dimensions['K'].width = 10# Source
    worksheet.column_dimensions['L'].width = 10# Account
    worksheet.column_dimensions['M'].width = 10# Deleted
    worksheet.column_dimensions['N'].width = 10# Tag Note
    worksheet.column_dimensions['O'].width = 20# Source file information
    worksheet.column_dimensions['P'].width = 10# Carved
    worksheet.column_dimensions['Q'].width = 10# Manually decoded
    worksheet.column_dimensions['R'].width = 20# business   
    worksheet.column_dimensions['S'].width = 10# number    
    worksheet.column_dimensions['T'].width = 20# street    
    worksheet.column_dimensions['Y'].width = 20# city    
    worksheet.column_dimensions['V'].width = 25# county    
    worksheet.column_dimensions['W'].width = 15# state    
    
      
        
    worksheet.column_dimensions['Z'].width = 26# FullAddress

    for row_index, row_data in enumerate(data):

        for col_index, col_name in enumerate(headers):
            cell_data = row_data.get(col_name)
            try:
                worksheet.cell(row=row_index+2, column=col_index+1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error printing line: {str(e)}{color_reset}")
    
    workbook.save(outuput_xlsx)


def usage():
    '''
    working examples of syntax
    '''
    file = sys.argv[0].split('\\')[-1]
    print(f'\nDescription: {color_green}{description}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f'\n    {color_yellow}insert your input into input.xlsx')
    print(f'\nExample:')
    print(f'    {file} -c -O input_blank.xlsx') 
    print(f'    {file} -r')
    print(f'    {file} -r -I locations.xlsx -O locations2addresses_.xlsx') 
 
                
if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
export a temp copy to output.txt
if it's less than 3000 skip the sleep timer

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
connection timeout after about 4000 attempts
with the sleep timer set to 10 (sec) it doesn't crap out.

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>

