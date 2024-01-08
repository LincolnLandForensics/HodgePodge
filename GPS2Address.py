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

import os
import sys
import time
import openpyxl
import simplekml

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
description = "convert GPS coordinates to addresses or visa versa & create a KML file"
version = '1.1.0'

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    global row
    row = 0  # defines arguments
    # Row = 1  # defines arguments   # if you want to add headers 
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='', required=False)
    parser.add_argument('-O', '--output', help='', required=False)
    parser.add_argument('-c', '--create', help='create blank input sheet', required=False, action='store_true')
    parser.add_argument('-k', '--kml', help='xlsx to kml with nothing else', required=False, action='store_true')
        
    parser.add_argument('-r', '--read', help='read xlsx', required=False, action='store_true')

    args = parser.parse_args()

    global input_xlsx
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
        write_xlsx(data)

    elif args.kml:
        data = []
        file_exists = os.path.exists(input_xlsx)
        if file_exists == True:
            print(f'{color_green}Reading {input_xlsx} {color_reset}')
            
            # data = read_xlsx(input_xlsx)
            data = read_xlsx_basic(input_xlsx)

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

    elif args.read:
        data = []
        file_exists = os.path.exists(input_xlsx)
        if file_exists == True:
            print(f'{color_green}Reading {input_xlsx} {color_reset}')
            
            data = read_xlsx(input_xlsx)

            # create kml file
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


def read_xlsx(input_xlsx):

    """Read data from an xlsx file and return as a list of dictionaries.
    Read XLSX Function: The read_xlsx() function reads data from the input 
    Excel file using the openpyxl library. It extracts headers from the 
    first row and then iterates through the data rows, creating dictionaries 
    for each row with headers as keys and cell values as values.
    
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
        (state, fulladdress, Latitude, Longitude, query, Coordinate) = ('', '', '', '', '', '')
        (Index, country) = ('', '')
        
        name_data = row_data.get("Name")
        coordinate_data = row_data.get("Coordinate")
        lat_data = row_data.get("Latitude")
        long_data = row_data.get("Longitude")
        address_data = row_data.get("Address")  # fulladdress
        fulladdress_data = row_data.get("fulladdress") 
        row_data["Index"] = (row_index + 2)

        if address_data is None:
            if fulladdress_data is None:
                fulladdress_data = ''

            address_data = fulladdress_data
            row_data["Address"] = address_data
            print(f'address_data = {address_data}') # temp
            
        # Check if address_data is None
        if address_data is None or address_data == '':
            print(f'')

        else:
            # Check the length of address_data
            if len(address_data) > 3:
                
                # turn fulladdress into lat/long
                try:
                    location = geolocator.geocode(address_data)
                except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset}")
                    
                try:
                    fulladdress = location.address
                    row_data["fulladdress"] = fulladdress                
                    address_parts = fulladdress.split(', ')
                    country = address_parts[-1]
                    row_data["country"] = country
                    
                    zipcode = address_parts[-2]
                    print(f'zipcode = {zipcode}')   # temp
                    row_data["zipcode"] = zipcode
                    
                    print(f'country = {county}')    # temp
                    
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
                   
  
                except Exception as e:
                    print(f"{color_red}Error : {str(e)}{color_reset}")                


            try:
                location = geolocator.reverse(query)
                fulladdress = location.address
                row_data["fulladdress"] = fulladdress

                # ... (rest of the code remains unchanged)

            except Exception as e:
                print(f"{color_red}Error: {str(e)}{color_reset}")



                # Sleep for x seconds
                time.sleep(8)
                # Check if latitude and longitude are float values
                # if isinstance(lat_data, float) and isinstance(long_data, float):
                    # query = f"{lat_data},{long_data}"
                    # row_data["query"] = query

                # if print(location.latitude, location.longitude)
                if isinstance(lat_data, float) and isinstance(long_data, float):
                # if long_data is None or lat_data is None:
                    try:
                        row_data["Longitude"] = location.longitude 
                        row_data["Latitude"] = location.latitude
                        query = (f'{location.latitude}, {location.longitude}')
                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 

        # check lat long
        if long_data is None or lat_data is None:
            # print('')
            print(f'{color_red}there is no lat or long{color_reset} {lat_data} - {long_data}')
        else:
            # print('')
            print(f'{color_yellow}trying {color_reset} {lat_data} {long_data}')   # temp         
            
            # Check if there is latitude data and it is a string
            if lat_data is not None and isinstance(lat_data, str) and len(lat_data) > 3:
                print(f'') # temp
                # b = b
            # if len(lat_data) > 3:
                if len(fulladdress) < 2:
                    
                    if address_data is None:
                        address_data = fulladdress_data
                        row_data["Address"] = address_data

                    print(f'{color_yellow}there is no fulladdress{color_reset} but there is lat long {lat_data} - {long_data}')
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
                        
                        if ", " in fulladdress:
                            
                            address_parts = fulladdress.split(', ')
                            country = address_parts[-1]
                            zipcode = address_parts[-2]
                            state = address_parts[-3]
                            county = address_parts[-4]
                            # city = address_parts[-5]

                        if fulladdress.count(',') == 7:
                            # Split the address by commas
                            address_parts = fulladdress.split(', ')
                            business = address_parts[0]
                            number = address_parts[1]
                            street = address_parts[2]
                            city = address_parts[3]

                        elif fulladdress.count(',') == 5:
                            # Split the address by commas
                            address_parts = fulladdress.split(', ')
                            business = address_parts[0]
                            number = address_parts[1]
                            street = address_parts[0]
                            city = address_parts[1]

                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 
                    
                    row_data["business"] = business 
                    row_data["number"] = number 
                    row_data["street"] = street
                    row_data["city"] = city  
                    row_data["county"] = county 
                    if state == "Illinois":
                        state = "IL" 
                    elif state == "Missouri":
                        state = "Mo"                         
                    row_data["state"] = state 
                    row_data["zipcode"] = zipcode 
                    if country == "United States":
                        country = "US"
                    row_data["country"] = country
                    
                    # Sleep x seconds
                    time.sleep(8) 

        # if 1 == 1:
            # print("hello world")
            # if 
       
        print(f'{query}\t{color_yellow}{fulladdress}{color_reset}')     
    # workbook.save(outuput_xlsx)

    write_xlsx(data)
    # write_kml(data)
    
    return data

def read_xlsx_basic(input_xlsx):

    """Read data from an xlsx file and return as a list of dictionaries.
    Read XLSX Function: The read_xlsx() function reads data from the input 
    Excel file using the openpyxl library. It extracts headers from the 
    first row and then iterates through the data rows, creating dictionaries 
    for each row with headers as keys and cell values as values.
    
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
        (state, fulladdress, Latitude, Longitude, query, Coordinate) = ('', '', '', '', '', '')
        (Index, country) = ('', '')

    
    return data



def write_kml(data):
    '''
    The write_kml() function receives the processed data as a list of 
    dictionaries and writes it to a kml using simplekml. 
    '''

    # Create KML object
    kml = simplekml.Kml()
    
    print(f'testing write_kml') # temp
    headers = [
        "Name", "Description", "Time", "End time", "Category", "Latitude"
        , "Longitude", "Map Address", "Address", "Type", "Source", "Account"
        , "Deleted", "Tag Note", "Source file information", "Carved"
        , "Manually decoded", "business", "number", "street", "city", "county"
        , "state", "zipcode", "country", "fulladdress", "query", "Plate"
        , "Container", "Sighting State", "Sighting Location", "Coordinate"
        , "Highway Name", "Direction", "Time (Local)", "Index"
    ]

    for row_index, row_data in enumerate(data):
        (description_data) = ('')
        
        name_data = row_data.get("Name")
        description_data = row_data.get("Description")
        time_data = row_data.get("Time")
        end_time_data = row_data.get("End time")
        coordinate_data = row_data.get("Coordinate")
        lat_data = row_data.get("Latitude")
        long_data = row_data.get("Longitude")
        address_data = row_data.get("Address")
        type_data = row_data.get("Type")
        business_data = row_data.get("business")    # test   
        # business_data = business_data.strip()   # test
        
        # if business_data is not None:
            # business_data = business_data.strip()

        full_address_data  = row_data.get("fulladdress")
        plate_data  = row_data.get("Plate")         # red
        hwy_data  = row_data.get("Highway Name")
        direction_data  = row_data.get("Direction")
                
        # index_data = row_data.get("Index")
        index_data = row_index  # task
  
  
        if description_data == None:
            description_data = ''       
        
        # (description_data) = (f'{index_data}, {description_data}')

        if name_data != None:
            (description_data) = (f'{description_data}\nNAME: {name_data}')
        
        if time_data != None:
            (description_data) = (f'{description_data}\nTIME: {time_data}')

        if end_time_data != None:
            (description_data) = (f'{description_data}\nendTime: {end_time_data}')

        if address_data != None:
            (description_data) = (f'{description_data}\n{address_data}')

        elif full_address_data != None:
            (description_data) = (f'{description_data}\nADDRESS: {full_address_data}')

        if hwy_data != None:
            (description_data) = (f'{description_data}\nHwy NAME: {hwy_data}')
            
        if direction_data != None:
            (description_data) = (f'{description_data}\nDIRECTION: {direction_data}')

        if type_data != None:
            (description_data) = (f'{description_data}\nTYPE: {type_data}')
            
        if business_data != None:
            (description_data) = (f'{description_data}\nBusiness: {business_data}')


        if plate_data != None:
            (description_data) = (f'{description_data}\nPLATE: {plate_data}')

            
        print(f'______________________{index_data + 2} description = {description_data}')

        point = None  # Initialize point variable outside the block

        if lat_data is not None and long_data is not None and type_data == "LPR":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.red
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != None:
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color

        elif lat_data is not None and long_data is not None and type_data == "Images":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.blue
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != None or business_data != "":
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color

        elif lat_data is not None and long_data is not None and type_data == "Videos":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.yellow
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != None or business_data != "":
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color


        # elif lat_data is not None and long_data is not None and type_data == "Videos":    # test
            # kml.newpoint(name=f"{index_data}", description=f"{description_data}", coords=[(long_data, lat_data)]).style.iconstyle.color = simplekml.Color.yellow  # lon, lat

        # elif lat_data is not None and long_data is not None and type_data == "Locations":    # test
            # print(f'index_data = {index_data}')   # temp
            # kml.newpoint(name=f"{index_data}", description=f"{description_data}", coords=[(long_data, lat_data)]).style.iconstyle.color = simplekml.Color.purple  # lon, lat

        elif lat_data is not None and long_data is not None and type_data == "Locations":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.purple
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != None:   
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color


        # elif lat_data is not None and long_data is not None:
            # kml.newpoint(name=f"{description_data}", coords=[(long_data, lat_data)])    #should it be lat, long?
            # kml.newpoint(name=f"{index_data}", description=f"{description_data}", coords=[(long_data, lat_data)])  # lon, lat
            # kml.newpoint(name=f"{index_data}", description=f"{description_data}", coords=[(long_data, lat_data)]).style.iconstyle.color = simplekml.Color.orange  # lon, lat

        elif lat_data is not None and long_data is not None:
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.orange
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != None:
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color




            # color options: red, orange, yellow, green, blue, purple, white, black, gray
            
        elif coordinate_data is not None:
            # kml.newpoint(name=f"{description_data}", coords=[(long_data, lat_data)])    #should it be lat, long?
            # kml.newpoint(name=f"{index_data}", description=f"{description_data}", coords=[(coordinate_data)])  # lon, lat
            print(f'fix coordinate_data {coordinate_data}')
            

    # Save the KML document to the specified output file
    kml.save(output_kml)

    print(f"KML file '{output_kml}' created successfully!")


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
        , "state", "zipcode", "country", "fulladdress", "query", "Plate"
        , "Container", "Sighting State", "Sighting Location", "Coordinate"
        , "Highway Name", "Direction", "Time (Local)", "Index", "#"
    ]

    # Write headers to the first row
    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [5, 6, 8]: 
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
            cell.fill = fill
        elif col_index in [0, 1, 2, 3, 9, 17, 25, 31, 32, 33]:  # yellow headers
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Use yellow color
            cell.fill = fill
        elif col_index == 27:  # Red for column 27
            fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Red color
            cell.fill = fill

    # Excel column width
    worksheet.column_dimensions['A'].width = 15# Name
    worksheet.column_dimensions['B'].width = 17# Description
    worksheet.column_dimensions['C'].width = 16# Time
    worksheet.column_dimensions['D'].width = 20# End time
    worksheet.column_dimensions['E'].width = 20# Category       or Group or Subgroup
    worksheet.column_dimensions['F'].width = 20# Latitude
    worksheet.column_dimensions['G'].width = 20# Longitude
    worksheet.column_dimensions['H'].width = 20# Map Address
    worksheet.column_dimensions['I'].width = 45# Address
    worksheet.column_dimensions['J'].width = 10# Type
    worksheet.column_dimensions['K'].width = 10# Source
    worksheet.column_dimensions['L'].width = 10# Account
    worksheet.column_dimensions['M'].width = 10# Deleted
    worksheet.column_dimensions['N'].width = 10# Tag Note       or Tag
    worksheet.column_dimensions['O'].width = 20# Source file information
    worksheet.column_dimensions['P'].width = 10# Carved
    worksheet.column_dimensions['Q'].width = 10# Manually decoded
    worksheet.column_dimensions['R'].width = 20# business   
    worksheet.column_dimensions['S'].width = 10# number    
    worksheet.column_dimensions['T'].width = 20# street    
    worksheet.column_dimensions['Y'].width = 20# city    
    worksheet.column_dimensions['V'].width = 25# county    
    worksheet.column_dimensions['W'].width = 15# state    
    worksheet.column_dimensions['X'].width = 8# zipcode   
    worksheet.column_dimensions['Y'].width = 9# country   
    worksheet.column_dimensions['Z'].width = 26# FullAddress
    worksheet.column_dimensions['AA'].width = 26# query
    # Flock
    worksheet.column_dimensions['AB'].width = 11# Plate
    worksheet.column_dimensions['AC'].width = 10# Container
    worksheet.column_dimensions['AD'].width = 17# Sighting State
    worksheet.column_dimensions['AE'].width = 17# Sighting Location
    worksheet.column_dimensions['AF'].width = 17# Coordinate
    worksheet.column_dimensions['AG'].width = 26# Highway Name
    worksheet.column_dimensions['AH'].width = 12# Direction
    worksheet.column_dimensions['AI'].width = 29# Time (LOCAL)
    worksheet.column_dimensions['AJ'].width = 8# Index
    worksheet.column_dimensions['AK'].width = 8# #
    
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
    print(f'\n    {color_yellow}insert your input into locations.xlsx')
    print(f'\nExample:')
    print(f'    {file} -c -O input_blank.xlsx') 
    print(f'    {file} -k -I locations.xlsx  # xlsx 2 kml with no internet processing')     
    print(f'    {file} -r')
    print(f'    {file} -r -I locations.xlsx -O locations2addresses_.xlsx') 
 
                
if __name__ == '__main__':
    main()

# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>

"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
if address / fulldata only, get lat/long
export a temp copy to output.txt
if it's less than 3000 skip the sleep timer

Add Group and Subgroup, color
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
connection timeout after about 4000 attempts
with the sleep timer set to 10 (sec) it doesn't crap out.

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>

