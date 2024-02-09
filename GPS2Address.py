 #!/usr/bin/env python3
# coding: utf-8
'''
read GPS coordinates and convert them to addresses
or
read addresses and convert them to coordinates
and
read GPS coordinates (xlsx) and convert them to KML
'''
# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import io
import requests
from openpyxl.drawing.image import Image

import os
import re
import sys
import time
import openpyxl
import simplekml
import geohash2
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import PatternFill

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
description2 = "convert GPS coordinates to addresses or visa versa & create a KML file"
version = '1.1.8'

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    global row
    row = 0  # defines arguments
    # Row = 1  # defines arguments   # if you want to add headers 
    parser = argparse.ArgumentParser(description=description2)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-c', '--create', help='create blank input sheet', required=False, action='store_true')
    parser.add_argument('-i', '--intel', help='convert intel sheet to locations', required=False, action='store_true')
  
    parser.add_argument('-k', '--kml', help='xlsx to kml with nothing else', required=False, action='store_true')
        
    parser.add_argument('-r', '--read', help='read xlsx', required=False, action='store_true')

    args = parser.parse_args()

    # global input_xlsx
    global outuput_xlsx
    global output_kml
    output_kml = 'gps.kml'




    
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
        write_locations(data)

    elif args.kml:
        data = []
        file_exists = os.path.exists(input_xlsx)
        if file_exists == True:
            print(f'{color_green}Reading {input_xlsx} {color_reset}')
            
            # data = read_xlsx(input_xlsx)
            data = read_locations(input_xlsx)

            # create kml file
            write_kml(data)
            
            
            # workbook.close()
            # print(f'{color_green}Writing to {outuput_xlsx} {color_reset}')
            print(f'{color_blue}Writing to gps.kml {color_reset}')
            print(f'''\n\n{color_yellow}
            visit https://earth.google.com/
            <file><Import KML> select gps.kml <open>
            {color_reset}\n''')
        else:
            print(f'{color_red}{input_xlsx} does not exist{color_reset}')
            exit()
    elif args.intel:
        data = []
        file_exists = os.path.exists(input_xlsx)
        if file_exists == True:
            print(f'{color_green}Reading {input_xlsx} {color_reset}')

            data = read_intel(input_xlsx)
            # write_kml(data)
            write_locations(data)
            print(f'{color_green}Writing to {outuput_xlsx} {color_reset}')
        else:
            print(f'{color_red}{input_xlsx} does not exist{color_reset}')
            exit()
            
    elif args.read:
        data = []
        file_exists = os.path.exists(input_xlsx)
        if file_exists == True:
            print(f'{color_green}Reading {input_xlsx} {color_reset}')
            
            # data = read_xlsx(input_xlsx)
            data = read_locations(input_xlsx)
            data = read_gps(data)
            write_locations(data)
            write_kml(data)

            workbook.close()
            print(f'{color_green}Writing to {outuput_xlsx} {color_reset}')
            print(f'''\n\n{color_yellow}
            visit https://earth.google.com/
            <file><Import KML> select gps.kml <open>
            {color_reset}\n''')
        else:
            print(f'{color_red}{input_xlsx} does not exist{color_reset}')
            exit()

    else:
        usage()
    
    return 0


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>


def convert_timestamp(timestamp, time_orig, timezone):
    if timezone != "" or timezone is None:
        timezone = ''
    
    time_data = timestamp
    timestamp = str(timestamp)
    if time_orig == "":
        time_orig = timestamp
    timestamp = timestamp.replace(' at ', ' ')
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

    # %B: Full month name (e.g., January, February, etc.)
    # %d: Day of the month as a zero-padded decimal number (01 to 31)
    # %Y: Year with century as a decimal number (e.g., 2023)
    # %I: Hour (12-hour clock) as a zero-padded decimal number (01 to 12)
    # %M: Minute as a zero-padded decimal number (00 to 59)
    # %S: Second as a zero-padded decimal number (00 to 59)
    # %p: AM or PM designation
    # %Z: Time zone name or abbreviation    
    
    formats = ["%Y-%m-%d %H:%M:%S",
               "%m/%d/%Y %I:%M:%S %p",
               "%m/%d/%Y %I:%M:%S %p",
               "%m/%d/%Y %I:%M %p",  # timestamps without seconds
               "%m/%d/%Y %H:%M:%S",  # timestamps in military time without seconds
               "%m/%d/%Y %I:%M %p",  # "07/26/2020 06:08 AM"
               "%B %d, %Y at %I:%M:%S %p %Z",  # timestamp with month name
               "%B %d, %Y at %I:%M:%S %p CST",  
               "%B %d, %Y at %I:%M:%S %p",  # "December 15, 2023 at 12:20:09 PM CST"
               "%B %d, %Y %I:%M:%S %p %Z",  # "December 15, 2023 12:20:09 PM CST" 
               "%B %d, %Y %I:%M:%S %p"]  # "December 15, 2023 12:20:09 PM"               
    for fmt in formats:
        try:
            # print(f'trying timestamp: {timestamp} with fmt:{fmt}')  # temp
            dt_obj = datetime.strptime(timestamp, fmt)
            time_data = dt_obj  # Assigning the datetime object to time_data
            return dt_obj, time_orig, timezone  # Return datetime object, original timestamp, and timezone
        except Exception as e:
            # print(f"Error3 : {str(e)}") 
            pass

                        


    # If no format matches, raise ValueError
    raise ValueError(f"{time_orig} Timestamp format not recognized")

    
def read_gps(data): 

    """Read data and return as a list of dictionaries.
    It extracts headers from the 
    first row and then iterates through the data rows, creating dictionaries 
    for each row with headers as keys and cell values as values.
    
    """

    for row_index, row_data in enumerate(data):
        print(f'\n{row_index + 2} ______________________\n')
        (street, Coordinate, fulladdress_data) = ('', '', '')
        
        (zipcode, business, number, street, city, county) = ('', '', '', '', '', '')
        (state, Latitude, Longitude, query, Coordinate) = ('', '', '', '', '')
        (Index, country, lat_data, long_data, PlusCode) = ('', '', '', '', '')
        (county, query, type_data) = ('', '', '')
        (location, skip, address_data) = ('', '', '')
        fulladdress_data = row_data.get("fulladdress")

        name_data = row_data.get("Name")
        lat_data = row_data.get("Latitude") # works
        lat_data = str(lat_data)
        if lat_data is None:
            lat_data = ''        
        long_data = row_data.get("Longitude") # works
        long_data = str(long_data)
        if long_data is None:
            long_data = ''        
        address_data = row_data.get("Address") # works
        type_data = row_data.get("Type") # works  
        business = row_data.get("business")
        number = row_data.get("number")
        city = row_data.get("city")
        county = row_data.get("county")
        state = row_data.get("state")
        zipcode = row_data.get("zipcode")
        query = row_data.get("query")
        country = row_data.get("country")

        coordinate_data = row_data.get("Coordinate")
        row_data["Index"] = (row_index + 2)
        PlusCode = row_data.get("PlusCode")


# skip lines with fulladdress
        if len(fulladdress_data) > 2 or fulladdress_data == 'Soul Buoy':        
            fulladdress_data = ''
            skip == 'skip'
# GPS to fulladdress
        elif lat_data != '' and isinstance(lat_data, str) and len(lat_data) > 5:

            if lat_data != '' and isinstance(lat_data, str) and len(lat_data) > 3:

                ##  if no fulladdress
                if len(fulladdress_data) < 2:
                    query = (f'{lat_data}, {long_data}') # backwards

                    try:
                        # location = geolocator.reverse((long_data, lat_data), language='en')

                        location = geolocator.reverse((lat_data, long_data), language='en') #lat/long
                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 
                    try:
                        fulladdress_data = location.address

                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 

                    if fulladdress_data == 'Soul Buoy':        
                        fulladdress_data = ''
  

                    time.sleep(8)   ## Sleep x seconds

# address to gps / Full address
        elif address_data != '':
            if len(address_data) > 3:

                try:
                    location = geolocator.geocode(address_data)

                    if location:
                        lat_data = location.latitude
                        long_data = location.longitude
                        query = (f'{address_data}')
                    else:
                        print(f"Location not found for address: {address_data}")

                except Exception as e:
                    print(f"Error: {str(e)}")

            try:
                location = geolocator.reverse((lat_data, long_data), language='en')
                fulladdress_data = location.address

            except Exception as e:
                print(f"{color_red}Error: {str(e)}{color_reset}")

            if fulladdress_data == 'Soul Buoy':        
                fulladdress_data = ''
                skip == 'skip'

    # coordinate        
            if lat_data != '' and long_data != '':
                coordinate_data = (f'{lat_data},{long_data}')

                time.sleep(8)   # Sleep for x seconds

        elif PlusCode != '':
            print(f'testing PlusCode {PlusCode}')
            
            try:
                decoded_location = geohash2.decode(PlusCode)
                lat_data = decoded_location[0]
                long_data = decoded_location[1]

            
                print(f"Coordinates for Geohash '{PlusCode}': {decoded_location[0]}, {decoded_location[1]}")
                
                
                
                # location = geolocator.geocode(PlusCode)
                if decoded_location is not None:
                    # lat_data, long_data = location.latitude, location.longitude
                    print(f'PlusCode {PlusCode} = {lat_data}, {long_data}')
                    fulladdress_data = location.address
                else:
                    print(f"Unable to find coordinates for address: {PlusCode}")
            except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset} PlusCode = <{PlusCode}>")  

        else:
            print(f'{color_red}none of the above{color_reset}')

# parse fulladdress
        if len(fulladdress_data) > 2 and skip != 'skip':
            try:
                address_parts = fulladdress_data.split(', ')
                country = address_parts[-1]
                zipcode = address_parts[-2]
                state = address_parts[-3]
                county = address_parts[-4]
                
                if "Township" in address_parts[-5]:
                    city = address_parts[-6]
                else:    
                    county = address_parts[-5]
                    city = address_parts[-6]
                    
                if fulladdress_data.count(',') == 7:    # Check if there are exactly 7 commas
                    if 'Township' in address_parts[3]:    # task
                             
                        address_parts = fulladdress_data.split(', ')    # Split the address by commas
                        business = address_parts[0]
                        number = address_parts[0]
                        street = address_parts[1]

                    else:
                        address_parts = fulladdress_data.split(', ')    # Split the address by commas
                        business = address_parts[0]
                        number = address_parts[1]
                        street = address_parts[2]
    
                elif fulladdress_data.count(',') == 5:
                    if 'Township' in address_parts[1]:    # task
                        address_parts = fulladdress_data.split(', ')    # Split the address by commas
                        business = address_parts[0]
                        number = address_parts[1]
                        street = address_parts[2]

                    else:
                        address_parts = fulladdress_data.split(', ')    # Split the address by commas
                        business = address_parts[0]
                        number = address_parts[1]
                        street = address_parts[2]

                if "United States" in country:
                    country = "US" 

                if street.endswith(" Township"):
                    street == ''                

                try:
                    if business == '' and address_parts[1].isdigit():
                        business = address_parts[0]
                except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset} Business = <{business}>")  


                try:
                    if business.isdigit():
                        business = ''
                    elif business.endswith(" Street") or business.endswith(" Road") or business.endswith(" Tollway") or business.endswith(" Avenue"): 
                        business = ''
                except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset} Business2 = <{business}>")  


                if address_parts[0].isdigit():
                    number = address_parts[0]
                    if street == '':
                        street = address_parts[1]
                if number is None:
                    number = ''
                    
                try:        
                    number = number if number.isdigit() else ''
                except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset} number = <{number}>")  


            except Exception as e:
                print(f"{color_red}Error : {str(e)}{color_reset} Business = <{business}> Full address =<{address_parts}>")  

# Icon    
        Icon = ''
        Icon = row_data.get("Icon")
        if Icon is None:
            Icon = ''

        if Icon != "":
            Icon = Icon        
        elif type_data == "Calendar":
            Icon = "Calendar"
        elif type_data == "LPR":
            Icon = "LPR"
        elif type_data == "Images":
            Icon = "Images"
        elif type_data == "Intel":
            Icon = "Intel"
        elif type_data == "Locations":
            Icon = "Locations"
        elif type_data == "Searched Items":
            Icon = "Searched"
        elif type_data == "Toll":
            Icon = "Toll"
        elif type_data == "Videos":
            Icon = "Videos"




# write rows to data
        row_data["Latitude"] = lat_data
        row_data["Longitude"] = long_data 
        row_data["Address"] = address_data
        row_data["business"] = business 
        row_data["number"] = number 
        row_data["street"] = street
        row_data["city"] = city 
        row_data["county"] = county 
        row_data["state"] = state 
        row_data["zipcode"] = zipcode
        row_data["country"] = country 
        row_data["fulladdress"] = fulladdress_data
        row_data["query"] = query
        row_data["Coordinate"] = coordinate_data
        row_data["PlusCode"] = PlusCode
        row_data["Icon"] = Icon
        
        print(f'\nName: {name_data}\nCoordinate: {coordinate_data}\naddress = {address_data}\nbusiness = {business}\nfulladdress_data = {fulladdress_data}\n')

    return data

def read_intel(input_xlsx):
    """Read intel_.xlsx sheet and convert it to locations format.    
    """

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

    if not data:
        print(f"{color_red}No data found in the Excel file.{color_reset}")
        return None

    for row_index, row_data in enumerate(data):

        (fullname, phone, business, fulladdress_data, city, state) = ('', '', '', '', '', '')
        (note, sosagent, Time, Latitude, Longitude, Coordinate) = ('', '', '', '', '', '')
        (Source, description_data, type_data, Name, tag, source) = ('', '', 'Intel', '', '', '')
        (group, subgroup, number, street, county, zipcode) = ('', '', '', '', '', '')
        (country, query, plate_data, capture_time, hwy_data, direction_data) = ('', '', '', '', '', '')
        (end_time_data, category_data, time_orig, timezone, PlusCode, Icon) = ('', '', '', '', '', '')

# Time
        time_data = ''
        time_data = row_data.get("Time")
        if time_data is None:
            time_data = ''

# Latitude
        lat_data = ''
        lat_data = row_data.get("Latitude")
        lat_data = str(lat_data)
        if lat_data is None or lat_data == 'None':
            lat_data = ''        

# Latitude
        long_data = ''
        long_data = row_data.get("Longitude")
        long_data = str(long_data)
        if long_data is None or long_data == 'None':
            long_data = ''        

# Coordinate    
        Coordinate = ''
        Coordinate = row_data.get("Coordinate")
        if Coordinate is None:
            Coordinate = ''
            
# address
        address_data = ''    
        address_data = row_data.get("fulladdress")
        if address_data is None:
            address_data = ''    
            
# Name    
        Name = ''
        Name = row_data.get("fullname")
        if Name is None:
            Name = ''

# source file
        source_file = row_data.get("Source file information")
        if source_file is None and input_xlsx != 'intel_.xlsx':
            source_file = input_xlsx

# sosagent    
        sosagent = ''
        sosagent = row_data.get("sosagent")
        if sosagent is None:
            sosagent = ''
            
# phone    
        phone = ''
        phone = row_data.get("phone")
        if phone is None:
            phone = ''
        else:
            description_data = (f'{description_data}\nPhone:{phone}')

# business    
        business = ''
        business = row_data.get("business")
        if business is None:
            business = ''
        else:
            description_data = (f'{description_data}\nBusiness:{business}')

        description_data = description_data.strip()

# city    
        city = ''
        city = row_data.get("city")
        if city is None:
            city = ''

# state    
        state = ''
        state = row_data.get("state")
        if state is None:
            state = ''

# Icon    
        Icon = ''
        Icon = row_data.get("Icon")
        if Icon is None:
            Icon = ''

      
# write rows to data
        row_data["Time"] = time_data
        row_data["Latitude"] = lat_data
        row_data["Longitude"] = long_data 
        row_data["Address"] = address_data
        # row_data["Group"] = group
        # row_data["Subgroup"] = subgroup
        row_data["Description"] = description_data
        row_data["Type"] = type_data
        row_data["Tag"] = tag        
        row_data["Source"] = source
        row_data["Source file information"] = source_file
        row_data["Name"] = Name
        row_data["business"] = business 
        # row_data["number"] = number 
        # row_data["street"] = street
        row_data["city"] = city 
        # row_data["county"] = county 
        row_data["state"] = state 
        # row_data["zipcode"] = zipcode
        # row_data["country"] = country 
        row_data["fulladdress"] = fulladdress_data
        # row_data["query"] = query
        # row_data["Plate"] = plate_data
        # row_data["Capture Time"] = capture_time
        # row_data["Highway Name"] = hwy_data
        row_data["Coordinate"] = Coordinate
        # row_data["Direction"] = direction_data
        # row_data["End time"] = end_time_data
        # row_data["Category"] = category_data
        # row_data["Time Original"] = time_orig
        # row_data["Timezone"] = timezone
        # row_data["PlusCode"] = PlusCode
        row_data["Icon"] = Icon
        # index
     
     
    return data



def read_locations(input_xlsx):

    """Read data from an xlsx file and return as a list of dictionaries.
    Read XLSX Function: The read_xlsx() function reads data from the input 
    Excel file using the openpyxl library. It extracts headers from the 
    first row and then iterates through the data rows.    
    """

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

    if not data:
        print(f"{color_red}No data found in the Excel file.{color_reset}")
        return None

    for row_index, row_data in enumerate(data):
        (zipcode, business, number, street, city, county) = ('', '', '', '', '', '')
        (state, fulladdress_data, Latitude, Longitude, query, Coordinate) = ('', '', '', '', '', '')
        (Index, country, capture_time, PlusCode, time_orig, Icon) = ('', '', '', '', '', '')
        (description_data, group, subgroup, source, source_file, tag) = ('', '', '', '', '', '')

        name_data = ''  # in case there is no Name column
        name_data = row_data.get("Name")
        if name_data is None:
            name_data = ''

# Description    
        description_data = ''
        description_data = row_data.get("Description")
        if description_data is None:
            description_data = ''

# Time
        time_data = ''
        time_data = row_data.get("Time")
        if time_data is None:
            time_data = ''

# time_orig
        time_orig = row_data.get("Time Original")
        if time_orig is None:
            time_orig = ''

# Capture Time
        capture_time = ''
        capture_time  = row_data.get("Capture Time") 
        if capture_time is None:
            capture_time = ''     

        if time_data == '' or time_data is None:
            if capture_time != '':
                time_data = capture_time


# timezone
        timezone = ''
        timezone = row_data.get("timezone")

# convert time
        output_format = "%m/%d/%Y %H:%M:%S"  # Changed to military time
        
        if time_orig == "" and time_data != '':
            time_orig = time_data
        try:
            (time_data, time_orig, timezone) = convert_timestamp(time_data, time_orig, timezone)
            time_data = time_data.strftime(output_format)
        except ValueError as e:
            # print(f"Error time2: {e} - {time_data}")
            time_data = ''
            
   
# End time
        end_time_data = ''
        end_time_data = row_data.get("End time")
        if end_time_data is None:
            end_time_data = ''        

# Category
        category_data = ''
        category_data = row_data.get("Category")
        if category_data is None:
            category_data = ''        

# gps
        lat_data = ''
        lat_data = row_data.get("Latitude")
        lat_data = str(lat_data)
        if lat_data is None or lat_data == 'None':
            lat_data = ''        

        long_data = ''
        long_data = row_data.get("Longitude")
        long_data = str(long_data)
        if long_data is None or long_data == 'None':
            long_data = ''        

        if lat_data == '': 
            lat_data = row_data.get("Capture Location Latitude")
            lat_data = str(lat_data)
            if lat_data is None or lat_data == 'None':
                lat_data = ''        

        if long_data == '':
            long_data = row_data.get("Capture Location Longitude")
            long_data = str(long_data)
            if long_data is None or long_data == 'None':
                long_data = ''        

# coordinate        
        coordinate_data = ''

        if lat_data != '' and long_data != '':
            coordinate_data = (f'{lat_data},{long_data}')
        elif row_data.get("Coordinate") != None:
            coordinate_data = row_data.get("Coordinate")
            
        elif row_data.get("Capture Location") != None:
            coordinate_data = row_data.get("Capture Location")
        elif row_data.get("Capture Location (Latitude,Longitude)") != None:
            coordinate_data = row_data.get("Capture Location (Latitude,Longitude)")

        if len(coordinate_data) > 6:
           
            coordinate_data = coordinate_data.replace('(', '').replace(')', '')
            if ',' in coordinate_data:
                coordinate_data = coordinate_data.split(',')
                lat_data = coordinate_data[0].strip()
                long_data = coordinate_data[1].strip()
                long_data = coordinate_data[1].strip()
                coordinate_data = (f'{lat_data},{long_data}')

# address
        address_data = ''    
        address_data = row_data.get("Address")
        if address_data is None:
            address_data = ''        

# group
        group = row_data.get("Group")
        if group is None:
            group = ''

# subgroup
        subgroup = row_data.get("Subgroup")
        if subgroup is None:
            subgroup = ''


# type
        type_data = ''
        type_data = row_data.get("Type")
        if type_data is None:
            type_data = ''  

# tag
        tag = ''
        tag = row_data.get("Tag")
        if tag is None:
            tag = ''

# source file
        source_file = row_data.get("Source file information")
        if source_file is None and input_xlsx != 'locations.xlsx':
            source_file = input_xlsx

# business  
        business = row_data.get("business")
        if business is None:
            business = ''
        # else:
            # try:
                # business = business.strip()
            # except Exception as e:
                # print(f"{color_red}Error stripping {business}: {str(e)}{color_reset}")

# fulladdress
        fulladdress_data  = row_data.get("fulladdress")
        if fulladdress_data is None:
            fulladdress_data = ''       

# query
        query = ''
        query  = row_data.get("query")
        if query is None:
            query = ''     

# Plate
        plate_data = ''
        plate_data  = row_data.get("Plate")         # red
        if plate_data is None:
            plate_data = ''     

        if plate_data != '' and type_data == '':
            type_data = 'LPR'
            


# country
        country = ''
        country = row_data.get("country")
        if country is None:
            country = ''     

# hwy
        hwy_data = ''
        hwy_data  = row_data.get("Highway Name")
        if hwy_data is None:
            hwy_data = ''           

        if hwy_data == '':
            hwy_data  = row_data.get("Capture Camera")
            if hwy_data is None:
                hwy_data = ''           

# Direction
        direction_data = ''
        direction_data  = row_data.get("Direction")
        if direction_data is None:
            direction_data = ''    

# PlusCode
        PlusCode = ''
        PlusCode  = row_data.get("PlusCode")
        if PlusCode is None:
            PlusCode = ''  

# Icon    
        Icon = ''
        Icon = row_data.get("Icon")
        if Icon is None:
            Icon = ''
        if Icon != "":
            Icon = Icon[0].upper() + Icon[1:].lower()
        
# write rows to data
        row_data["Time"] = time_data
        row_data["Latitude"] = lat_data
        row_data["Longitude"] = long_data 
        row_data["Address"] = address_data
        row_data["Group"] = group
        row_data["Subgroup"] = subgroup
        row_data["Description"] = description_data
        row_data["Type"] = type_data
        row_data["Tag"] = tag        
        row_data["Source"] = source
        row_data["Source file information"] = source_file
        row_data["Name"] = name_data
        row_data["business"] = business 
        row_data["number"] = number 
        row_data["street"] = street
        row_data["city"] = city 
        row_data["county"] = county 
        row_data["state"] = state 
        row_data["zipcode"] = zipcode
        row_data["country"] = country 
        row_data["fulladdress"] = fulladdress_data
        row_data["query"] = query
        row_data["Plate"] = plate_data
        row_data["Capture Time"] = capture_time
        row_data["Highway Name"] = hwy_data
        row_data["Coordinate"] = coordinate_data
        row_data["Direction"] = direction_data
        row_data["End time"] = end_time_data
        row_data["Category"] = category_data
        row_data["Time Original"] = time_orig
        row_data["Timezone"] = timezone
        row_data["PlusCode"] = PlusCode
        row_data["Icon"] = Icon        
        # index
     
     
    return data


def write_kml(data):
    '''
    The write_kml() function receives the processed data as a list of 
    dictionaries and writes it to a kml using simplekml. 
    '''

    # Create KML object
    kml = simplekml.Kml()

    # Define different default icons
    # square_icon = 'http://maps.google.com/mapfiles/kml/shapes/square.png'
    # triangle_icon = 'http://maps.google.com/mapfiles/kml/shapes/triangle.png'
    # star_icon = 'http://maps.google.com/mapfiles/kml/shapes/star.png'
    # polygon_icon = 'http://maps.google.com/mapfiles/kml/shapes/polygon.png'
    # circle_icon = 'http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png'
    # yellow_circle_icon = 'http://maps.google.com/mapfiles/kml/paddle/ylw-circle.png'
    # red_circle_icon = 'http://maps.google.com/mapfiles/kml/paddle/red-circle.png'
    # white_circle_icon = 'http://maps.google.com/mapfiles/kml/paddle/wht-circle.png'



    default_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon13.png'   # yellow flag

    car_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon15.png'   # red car
    car2_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon47.png'  # yellow car
    car3_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon54.png'  # green car with circle
    car4_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon7.png'  # red car with circle

    truck_icon = 'https://maps.google.com/mapfiles/kml/shapes/truck.png'    # blue truck

    calendar_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon23.png' # paper
    locations_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon28.png'    # yellow paddle
    home_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon56.png'
    images_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon46.png'
    intel_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon44.png'
    office_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon21.png'
    searched_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon9.png'
    toll_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-none.png'
    videos_icon = 'https://maps.google.com/mapfiles/kml/pal2/icon30.png'
    n_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-0.png'
    e_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-4.png'
    s_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-8.png'
    w_icon = 'https://earth.google.com/images/kml-icons/track-directional/track-12.png'
      
    

    for row_index, row_data in enumerate(data):
        index_data = row_index + 2  # excel row starts at 2, not 0
        
        time_data = row_data.get("Time") #
        lat_data = row_data.get("Latitude") #
        long_data = row_data.get("Longitude") #
        address_data = row_data.get("Address") #
        group_data = row_data.get("Group") 
        subgroup_data = row_data.get("Subgroup")   
        description_data = row_data.get("Description")
        type_data = row_data.get("Type")#
        tag = row_data.get("Tag")
        source_file = row_data.get("Source file information") #
        name_data = row_data.get("Name")
        business = row_data.get("business")
        fulladdress_data  = row_data.get("fulladdress")
        plate_data  = row_data.get("Plate")   
        hwy_data  = row_data.get("Highway Name") #
        coordinate_data = row_data.get("Coordinate")
        direction_data  = row_data.get("Direction")
        end_time_data = row_data.get("End time")
        category_data = row_data.get("Category")
        Icon = row_data.get("Icon")


        if name_data != '':
            (description_data) = (f'{description_data}\nNAME: {name_data}')
        
        if time_data != '':
            (description_data) = (f'{description_data}\nTIME: {time_data}')

        if end_time_data != '':
            (description_data) = (f'{description_data}\nEnd Time: {end_time_data}')

        if address_data != '':
            (description_data) = (f'{description_data}\n{address_data}')

        elif fulladdress_data != '':
            (description_data) = (f'{description_data}\nADDRESS: {fulladdress_data}')

        if hwy_data != '':
            (description_data) = (f'{description_data}\nHWY NAME: {hwy_data}')
            
        if direction_data != '':
            (description_data) = (f'{description_data}\nDIRECTION: {direction_data}')

        # if source_file != '' and source_file != None:
            # (description_data) = (f'{description_data}\nSOURCE: {source_file}')

        if tag != '':
            (description_data) = (f'{description_data}\nTAG: {tag}')    # test

        if type_data != '':
            (description_data) = (f'{description_data}\nTYPE: {type_data}')

        if group_data != '':
            (description_data) = (f'{description_data} / {group_data}')

        if subgroup_data != '' and subgroup_data != 'Unknown':
            (description_data) = (f'{description_data} / {subgroup_data}')

        if business != '':
            (description_data) = (f'{description_data}\nBusiness: {business}')
            
        if plate_data != '':
            (description_data) = (f'{description_data}\nPLATE: {plate_data}')

        point = ''  # Initialize point variable outside the block
        
        if lat_data == '' or long_data == '' or lat_data == None or long_data == None:
            print(f'skipping row {index_data} - No GPS')

        elif Icon == "Lpr" or Icon == "Car":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = car_icon  # red car

        elif Icon == "Car2":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = car2_icon  # yellow car

        elif Icon == "Car3":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = car3_icon  # green car

        elif Icon == "Car4":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = car4_icon  # red car (with circle)

        elif Icon == "Truck":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = truck_icon

        elif Icon == "Calendar":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = calendar_icon

        elif Icon == "Home":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = home_icon

        elif Icon == "Images":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = images_icon

        elif Icon == "Intel":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = intel_icon

        elif Icon == "Office":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = office_icon

        elif Icon == "Searched":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = searched_icon

        elif Icon == "Videos":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = videos_icon

        elif Icon == "Locations":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = locations_icon    # yellow paddle

        elif Icon == "Toll":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = toll_icon

        elif Icon == "N":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = n_icon

        elif Icon == "E":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = e_icon

        elif Icon == "S":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = s_icon

        elif Icon == "W":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = w_icon

        else:
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.icon.href = default_icon    # orange paddle


        
        if tag != '':   # mark label yellow if tag is not blank
            point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            # point.style.labelstyle.scale = 1.2  # Adjust label scale if needed    # task


    
    kml.save(output_kml)    # Save the KML document to the specified output file

    print(f"KML file '{output_kml}' created successfully!")


def write_locations(data):
    '''
    The write_locations() function receives the processed data as a list of 
    dictionaries and writes it to a new Excel file using openpyxl. 
    It defines the column headers, sets column widths, and then iterates 
    through each row of data, writing it into the Excel worksheet.
    '''
    print(f'write_locations')   # temp
    global workbook
    workbook = Workbook()
    global worksheet
    worksheet = workbook.active

    worksheet.title = 'Locations'
    header_format = {'bold': True, 'border': True}
    worksheet.freeze_panes = 'B2'  # Freeze cells
    worksheet.selection = 'B2'

    headers = [
        "#", "Time", "Latitude", "Longitude", "Address", "Group", "Subgroup"
        , "Description", "Type", "Source", "Deleted", "Tag", "Source file information"
        , "Service Identifier", "Carved", "Name", "business", "number", "street"
        , "city", "county", "state", "zipcode", "country", "fulladdress", "query"
        , "Sighting State", "Plate", "Capture Time", "Capture Network", "Highway Name"
        , "Coordinate", "Capture Location Latitude", "Capture Location Longitude"
        , "Container", "Sighting Location", "Direction", "Time Local", "End time"
        , "Category", "Manually decoded", "Account", "PlusCode", "Time Original", "Timezone"
        , "Icon", "Index"

    ]

    # Write headers to the first row
    for col_index, header in enumerate(headers):
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
    worksheet.column_dimensions['L'].width = 4# Tag
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
    worksheet.column_dimensions['AT'].width = 6 # Icon   
    worksheet.column_dimensions['AU'].width = 6 # Index

    
    for row_index, row_data in enumerate(data):

        for col_index, col_name in enumerate(headers):
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

    color_data = [

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


        ('', '', ''),
        ('', 'NOTE', 'visit https://earth.google.com/ <file><Import KML> select gps.kml <open>'),
    ]

    for row_index, (icon, tag, description) in enumerate(color_data):
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
    locations_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon28.png'    # yellow paddle
    home_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon56.png'
    images_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon46.png'
    intel_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon44.png'
    office_icon = 'https://maps.google.com/mapfiles/kml/pal3/icon21.png'
    searched_icon = 'https://maps.google.com/mapfiles/kml/pal4/icon9.png'
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
     
        
    except:
        pass

    
    workbook.save(outuput_xlsx)


def usage():
    '''
    working examples of syntax
    '''
    file = sys.argv[0].split('\\')[-1]
    print(f'\nDescription: {color_green}{description2}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f'\n    {color_yellow}insert your input into locations.xlsx')
    print(f'\nExample:')
    print(f'    {file} -c -O input_blank.xlsx') 
    print(f'    {file} -k -I locations.xlsx  # xlsx 2 kml with no internet processing')     
    print(f'    {file} -r')
    print(f'    {file} -r -I locations.xlsx -O locations2addresses_.xlsx') 
    print(f'    {file} -i -I intel_.xlsx -O intel2locations_.xlsx')  
                
if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
1.1.0 - color code sheet
1.0.1 - Color coded pins for gps.kml
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
run -r a second time and fulladdress and Tag are blank.

if address / fulldata only, get lat/long
export a temp copy to output.txt
if it's less than 3000 skip the sleep timer

Add Group and Subgroup, color
if coordinate but no address or gps. convert coordinate to lat/long
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
connection timeout after about 4000 attempts
with the sleep timer set to 10 (sec) it doesn't crap out.

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>

