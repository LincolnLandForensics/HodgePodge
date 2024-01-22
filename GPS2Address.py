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
from datetime import datetime



from openpyxl import Workbook   # test
from openpyxl.styles import PatternFill # test


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
version = '1.1.5'

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
            
            # data = read_xlsx(input_xlsx)
            data = read_xlsx_basic(input_xlsx)
            data = read_gps(data)
            write_xlsx(data)
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

def convert_date_format(input_date):
    (output_date) = ('')
    input_date = input_date.replace(' at ',' ')
    date_parts = input_date.split(' ')
    updated_date_parts = date_parts[:-1]
    updated_date_string = ' '.join(updated_date_parts)
    input_date = updated_date_string

    if input_date.count(',') == 1:
        input_format = "%B %d, %Y %I:%M:%S %p"

        try:
            dt_object = datetime.strptime(input_date, input_format)
        except Exception as e:
            print(f"Error : {str(e)}") 
        
        try:
            output_format = "%m/%d/%Y %I:%M:%S %p"
            output_date = dt_object.strftime(output_format)
        except Exception as e:
            print(f"Error : {str(e)}") 
    elif input_date.count(',') == 2:
        input_format = "%B %d, %Y, %I:%M:%S %p"

        try:
            dt_object = datetime.strptime(input_date, input_format)
        except Exception as e:
            print(f"Error : {str(e)}") 
        
        try:
            output_format = "%m/%d/%Y %I:%M:%S %p"
            output_date = dt_object.strftime(output_format)
        except Exception as e:
            print(f"Error : {str(e)}") 
        
    return output_date


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
        (Index, country, lat_data, long_data) = ('', '', '', '')
        (county, query) = ('', '')
        (location, skip) = ('', '')
        # print(f'fulladdress = {row_data.get("fulladdress")}')    # temp
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

# skip lines with fulladdress
        if len(fulladdress_data) > 2:        
            skip = 'skip'

            print(f'{color_green} full address =  {fulladdress_data}{color_reset}')    # temp
# GPS to fulladdress
        elif lat_data != '' and isinstance(lat_data, str) and len(lat_data) > 5:
            print(f'{color_yellow} Coordinate =  {lat_data}, {long_data}{color_reset}')    # temp

            if lat_data != '' and isinstance(lat_data, str) and len(lat_data) > 3:
                print(f'') # temp
                ##  if no fulladdress
                if len(fulladdress_data) < 2:
                    query = (f'{lat_data}, {long_data}') # backwards

                    try:
                        location = geolocator.reverse((lat_data, long_data), language='en')
                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 
                    try:
                        fulladdress_data = location.address

                    except Exception as e:
                        print(f"{color_red}Error : {str(e)}{color_reset}") 

                    time.sleep(8)   ## Sleep x seconds

# address to gps / Full address
        elif address_data != '':
            print(f'{color_blue} address =  {address_data}{color_reset}')    # temp
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

                time.sleep(8)   # Sleep for x seconds
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
                        # print(f'township = {address_parts[1]}') # temp   
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

                
                if business == '' and address_parts[1].isdigit():
                    business = address_parts[0]
                
                if business.isdigit():
                    business = ''
                elif business.endswith(" Street") or business.endswith(" Road") or business.endswith(" Tollway") or business.endswith(" Avenue"): 
                    business = ''

                if address_parts[0].isdigit():
                    number = address_parts[0]
                    if street == '':
                        street = address_parts[1]
                        
                number = number if number.isdigit() else ''

            except Exception as e:
                print(f"{color_red}Error : {str(e)}{color_reset}")  

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

        
        print(f'\nName: {name_data}\nLat\Long: {lat_data}, {long_data}\naddress = {address_data}\nbusiness = {business}\nfulladdress_data = {fulladdress_data}\n')

    return data

def read_xlsx_basic(input_xlsx):

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
        (Index, country, capture_time) = ('', '', '')
        (description_data) = ('')
        
        ## replace all None values with '' 
        name_data = ''  # in case there is no Name column
        name_data = row_data.get("Name")
        if name_data is None:
            name_data = ''
        row_data["Name"] = name_data
        
        description_data = ''
        description_data = row_data.get("Description")
        if description_data is None:
            description_data = ''
        row_data["Description"] = description_data

        time_data = ''
        time_data = row_data.get("Time")
        if time_data is None:
            time_data = ''
        row_data["Time"] = time_data
        
        end_time_data = ''
        end_time_data = row_data.get("End time")
        if end_time_data is None:
            end_time_data = ''        
        row_data["End time"] = end_time_data

        category_data = ''
        category_data = row_data.get("Category")
        if category_data is None:
            category_data = ''        
        row_data["Category"] = category_data

        lat_data = ''
        lat_data = row_data.get("Latitude")
        lat_data = str(lat_data)
        if lat_data is None or lat_data == 'None':
            lat_data = ''        
        row_data["Latitude"] = lat_data

        long_data = ''
        long_data = row_data.get("Longitude")
        long_data = str(long_data)
        if long_data is None or long_data == 'None':
            long_data = ''        
        row_data["Longitude"] = long_data

        if lat_data == '': 
            lat_data = row_data.get("Capture Location Latitude")
            lat_data = str(lat_data)
            if lat_data is None or lat_data == 'None':
                lat_data = ''        
            row_data["Latitude"] = lat_data

        if long_data == '':
            long_data = ''
            long_data = row_data.get("Capture Location Longitude")
            long_data = str(long_data)
            if long_data is None or long_data == 'None':
                long_data = ''        
            row_data["Longitude"] = long_data

        address_data = ''    
        address_data = row_data.get("Address")
        if address_data is None:
            address_data = ''        
        row_data["Address"] = address_data

        type_data = ''
        type_data = row_data.get("Type")
        if type_data is None:
            type_data = ''  
        row_data["Type"] = type_data
        
        business_data = ''
        business_data = row_data.get("business")    # test   
        if business_data is None:
            business_data = ''
        else:
            try:
                business_data = business_data.strip()
            except Exception as e:
                print(f"{color_red}Error stripping {business_data}: {str(e)}{color_reset}")
        row_data["business"] = business_data    

        fulladdress_data  = row_data.get("fulladdress")
        if fulladdress_data is None:
            fulladdress_data = ''       
        row_data["fulladdress"] = fulladdress_data

        query = ''
        query  = row_data.get("query")
        if query is None:
            query = ''     
        row_data["query"] = query


        plate_data = ''
        plate_data  = row_data.get("Plate")         # red
        if plate_data is None:
            plate_data = ''     
        row_data["Plate"] = plate_data

        capture_time = ''
        capture_time  = row_data.get("Capture Time") 
        if capture_time is None:
            capture_time = ''     
        row_data["Capture Time"] = capture_time

        if time_data == '' or time_data is None:
            if capture_time != '':
                # capture_time = "January 13, 2022, 9:41:33 PM CDT"
                converted_date = convert_date_format(capture_time)
                row_data["Time"] = converted_date
                # print(f'Converted Capture time to new format {converted_date}')


        country = ''
        country = row_data.get("country")
        if country is None:
            country = ''     
        row_data["country"] = country


        
        lat_data = row_data.get("Coordinate")
        lat_data = row_data.get("Latitude")
        if lat_data is None:
            lat_data = ''  
        
        coordinate_data = ''
        coordinate_data = row_data.get("Coordinate")
        if coordinate_data is None:
            coordinate_data = ''        
        if coordinate_data == '':
            if lat_data != '' and long_data != '':
                coordinate_data = (f'{long_data}, {lat_data}')
        row_data["Coordinate"] = coordinate_data

        if coordinate_data == '':
            coordinate_data = row_data.get("Capture Location")
            if coordinate_data is None:
                coordinate_data = '' 
            coordinate_data = coordinate_data.replace('(', '').replace(')', '')

            row_data["Coordinate"] = coordinate_data

        hwy_data = ''
        hwy_data  = row_data.get("Highway Name")
        if hwy_data is None:
            hwy_data = ''           
        row_data["Highway Name"] = hwy_data

        if hwy_data == '':
            hwy_data  = row_data.get("Capture Camera")
            if hwy_data is None:
                hwy_data = ''           
            row_data["Highway Name"] = hwy_data

        direction_data = ''
        direction_data  = row_data.get("Direction")
        if direction_data is None:
            direction_data = ''    
        row_data["Direction"] = direction_data
        
    return data


def write_kml(data):
    '''
    The write_kml() function receives the processed data as a list of 
    dictionaries and writes it to a kml using simplekml. 
    '''

    # Create KML object
    kml = simplekml.Kml()
    
    print(f'testing write_kml') # temp

    for row_index, row_data in enumerate(data):
        index_data = row_index + 2  # excel row starts at 2, not 0
        
        time_data = row_data.get("Time")
        lat_data = row_data.get("Latitude")
        long_data = row_data.get("Longitude")
        address_data = row_data.get("Address")
        group_data = row_data.get("Group") 
        subgroup_data = row_data.get("Subgroup")   
        description_data = row_data.get("Description")
        type_data = row_data.get("Type")
        source_data = row_data.get("Source")
        name_data = row_data.get("Name")
        business_data = ''
        business_data = row_data.get("business")    # test   
        fulladdress_data  = row_data.get("fulladdress")
        plate_data  = row_data.get("Plate")   
        hwy_data  = row_data.get("Highway Name")
        coordinate_data = row_data.get("Coordinate")
        direction_data  = row_data.get("Direction")
        end_time_data = row_data.get("End time")
        category_data = row_data.get("Category")

        if name_data != '':
            (description_data) = (f'{description_data}\nNAME: {name_data}')
        
        if time_data != '':
            (description_data) = (f'{description_data}\nTIME: {time_data}')

        if end_time_data != '':
            (description_data) = (f'{description_data}\nendTime: {end_time_data}')

        if address_data != '':
            (description_data) = (f'{description_data}\n{address_data}')

        elif fulladdress_data != '':
            (description_data) = (f'{description_data}\nADDRESS: {fulladdress_data}')

        if hwy_data != '':
            (description_data) = (f'{description_data}\nHwy NAME: {hwy_data}')
            
        if direction_data != '':
            (description_data) = (f'{description_data}\nDIRECTION: {direction_data}')

        if source_data != '' and source_data != None:
            (description_data) = (f'{description_data}\nSOURCE: {source_data}')

        if type_data != '':
            (description_data) = (f'{description_data}\nTYPE: {type_data}')

        if group_data != '':
            (description_data) = (f'{description_data} / {group_data}')

        if subgroup_data != '' and subgroup_data != 'Unknown':
            (description_data) = (f'{description_data} / {subgroup_data}')

        if business_data != '':
            (description_data) = (f'{description_data}\nBusiness: {business_data}')

        if plate_data != '':
            (description_data) = (f'{description_data}\nPLATE: {plate_data}')

        point = ''  # Initialize point variable outside the block
        
        if lat_data == '' or long_data == '' or lat_data == None or long_data == None:
            print(f'skipping row {index_data} - No GPS')

        elif type_data == "LPR":
        # if lat_data != '' and long_data != '' and type_data == "LPR":
            
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.red
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed

            if business_data != '':
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color

        elif type_data == "Images":
        # elif lat_data != '' and long_data != '' and type_data == "Images":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.blue
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != "":
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color

        elif type_data == "Videos":
        # elif lat_data != '' and long_data != '' and type_data == "Videos":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.yellow
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != "":
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color

        elif type_data == "Locations":
        # elif lat_data != '' and long_data != '' and type_data == "Locations":
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.purple
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != '':   
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color

        elif lat_data != '' and long_data != '':
            point = kml.newpoint(
                name=f"{index_data}",
                description=f"{description_data}",
                coords=[(long_data, lat_data)]
            )
            point.style.iconstyle.color = simplekml.Color.orange
            point.style.labelstyle.scale = 0.8  # Adjust label scale if needed
            point.style.labelstyle.text = "Videos"  # Set the label text

            if business_data != '':
                point.style.labelstyle.color = simplekml.Color.yellow  # Set label text color
            else:   
                point.style.labelstyle.color = simplekml.Color.white  # Set label text color

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

    # headers_old = [
        # "Name", "Description", "Time", "End time", "Category", "Latitude"
        # , "Longitude", "Map Address", "Address", "Type", "Source", "Account"
        # , "Deleted", "Tag Note", "Source file information", "Carved"
        # , "Manually decoded", "business", "number", "street", "city", "county"
        # , "state", "zipcode", "country", "fulladdress", "query", "Plate"
        # , "Container", "Sighting State", "Sighting Location", "Coordinate"
        # , "Highway Name", "Direction", "Time (Local)", "Index", "#"
    # ]
    
    headers = [
        "#", "Time", "Latitude", "Longitude", "Address", "Group", "Subgroup"
        , "Description", "Type", "Source", "Deleted", "Tag", "Source file information"
        , "Service Identifier", "Carved", "Name", "business", "number", "street"
        , "city", "county", "state", "zipcode", "country", "fulladdress", "query"
        , "Sighting State", "Plate", "Capture Time", "Capture Network", "Highway Name"
        , "Coordinate", "Capture Location Latitude", "Capture Location Longitude"
        , "Index", "Container", "Sighting Location", "Highway Name", "Direction"
        , "Time Local", "End time", "Category", "Manually decoded", "Account"
    ]



    # Write headers to the first row
    for col_index, header in enumerate(headers):
        cell = worksheet.cell(row=1, column=col_index + 1)
        cell.value = header
        if col_index in [2, 3, 4]: 
            fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # orange?
            cell.fill = fill
        elif col_index in [1, 5, 6, 7, 8, 9, 15, 16, 24, 30, 31, 37, 38, 40]:  # yellow headers
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Use yellow color
            cell.fill = fill
        elif col_index == 27:  # Red for column 27
            fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Red color
            cell.fill = fill

    ## Excel column width
    worksheet.column_dimensions['A'].width = 8# #
    worksheet.column_dimensions['B'].width = 16# Time
    worksheet.column_dimensions['C'].width = 20# Latitude
    worksheet.column_dimensions['D'].width = 20# Longitude
    worksheet.column_dimensions['E'].width = 45# Address  Category       or Group or Subgroup
    worksheet.column_dimensions['F'].width = 15# Group
    worksheet.column_dimensions['G'].width = 15# Subgroup
    worksheet.column_dimensions['H'].width = 17# Description
    worksheet.column_dimensions['I'].width = 10# Type
    worksheet.column_dimensions['J'].width = 10# Source
    worksheet.column_dimensions['K'].width = 10# Source
    worksheet.column_dimensions['L'].width = 10# Tag
    worksheet.column_dimensions['M'].width = 20# Source file information
    worksheet.column_dimensions['N'].width = 15# Service Identifier
    worksheet.column_dimensions['O'].width = 10# Carved
    worksheet.column_dimensions['P'].width = 15# Name
    
    ## bonus
    worksheet.column_dimensions['Q'].width = 20# business 
    worksheet.column_dimensions['R'].width = 10# number
    worksheet.column_dimensions['S'].width = 20# street 
    worksheet.column_dimensions['T'].width = 20# city   
    worksheet.column_dimensions['Y'].width = 25# county    
    worksheet.column_dimensions['V'].width = 15# state   
    worksheet.column_dimensions['W'].width = 8# zipcode     
    worksheet.column_dimensions['X'].width = 8# country    
    worksheet.column_dimensions['Y'].width = 26# FullAddress   
    worksheet.column_dimensions['Z'].width = 26# query

    ##  Flock
    worksheet.column_dimensions['AA'].width = 11# Sighting State
    worksheet.column_dimensions['AB'].width = 11# Plate
    worksheet.column_dimensions['AC'].width = 20# Capture Time
    worksheet.column_dimensions['AD'].width = 15# Capture Network
    worksheet.column_dimensions['AE'].width = 26# Highway Name
    worksheet.column_dimensions['AF'].width = 17# Coordinate
    worksheet.column_dimensions['AG'].width = 26# Capture Location Latitude
    worksheet.column_dimensions['AH'].width = 26# Capture Location Longitude

    ##
    worksheet.column_dimensions['AI'].width = 6# Index
    worksheet.column_dimensions['AJ'].width = 10# Container
    worksheet.column_dimensions['AK'].width = 18# Sighting Location
    worksheet.column_dimensions['AL'].width = 14# Highway Name
    worksheet.column_dimensions['AM'].width = 10# Direction
    worksheet.column_dimensions['AN'].width = 11# Time Local
    worksheet.column_dimensions['AO'].width = 11# End time
    worksheet.column_dimensions['AP'].width = 10# Category
    worksheet.column_dimensions['AQ'].width = 18# Manually decoded
    worksheet.column_dimensions['AR'].width = 10# Account

    
    for row_index, row_data in enumerate(data):

        for col_index, col_name in enumerate(headers):
            cell_data = row_data.get(col_name)
            try:
                worksheet.cell(row=row_index+2, column=col_index+1).value = cell_data
            except Exception as e:
                print(f"{color_red}Error printing line: {str(e)}{color_reset}")


    # Create a new worksheet for color codes
    color_worksheet = workbook.create_sheet(title='GPS Pin Color Codes')
    color_worksheet.freeze_panes = 'B2'  # Freeze cells

    # Excel column width
    color_worksheet.column_dimensions['A'].width = 16# Name
    color_worksheet.column_dimensions['B'].width = 60# Description 21
    
    # Define color codes
    color_worksheet['A1'] = 'GPS Pin Colors'
    color_worksheet['B1'] = 'Description'

    color_data = [
        ('Red', 'LPR (License Plate Reader)'),
        ('Orange', 'Default pin color'),  
        ('Yellow', 'Videos'),        
        ('Black', 'Images'),
        ('Purple', 'Locations'),
        ('Yellow font', 'Business'),
        ('', ''),
        ('NOTE', 'visit https://earth.google.com/ <file><Import KML> select gps.kml <open>'),
    ]

    for row_index, (color, description) in enumerate(color_data):
        color_worksheet.cell(row=row_index + 2, column=1).value = color
        color_worksheet.cell(row=row_index + 2, column=2).value = description
    
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
1.1.0 - color code sheet
1.0.1 - Color coded pins for gps.kml
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>

"""
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

