#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import sys
import exifread
import openpyxl
from datetime import datetime

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "exif data reader"
version = '0.1.0'

global image_types
image_types = ['.heic', '.jpg', '.jpeg', '.png', '.tiff', '.tif', '.webp']

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

    if major_version >= 10 and build_version >= 22000:  # Windows 11 and above
        import colorama
        from colorama import Fore, Back, Style  
        print(f'{Back.BLACK}')  # make sure background is black
        color_red = Fore.RED
        color_yellow = Fore.YELLOW
        color_green = Fore.GREEN
        color_blue = Fore.BLUE
        color_purple = Fore.MAGENTA
        color_reset = Style.RESET_ALL

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu           >>>>>>>>>>>>>>>>>>>>>>>>>>


def main():
    # Folder containing photos
    photos_folder = 'photos'  # Change this to your folder path
    timestamp = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    output_file = f'exif_data_{timestamp}.xlsx'

    # Create a new workbook and add basic formatting
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "FileMetadata"

    # Define headers based on EXIF data keys
    headers = [
        'Name', 'DateCreated', 'DateTimeOriginal', 'FileCreateDate', 'FileModifyDate', 'Timezone',
        'DeviceManufacturer', 'ExifToolVersion', 'FileSize', 'FileType', 'FileTypeExtension', 'LensMake',
        'Software', 'HostComputer', 'LensInfo', 'LensModel', 'Model', 'NumberOfImages', 
        'Altitude', 'Latitude', 'Longitude', 'Coordinate', 'Time', 'Type', 'Icon', 'Description'
    ]
    sheet.append(headers)

    # Format headers as bold and set column widths
    for col, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col, value=header).font = openpyxl.styles.Font(bold=True)
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15

    # Freeze the top row at cell B2
    sheet.freeze_panes = 'B2'

    # Check if the path exists
    if not os.path.exists(photos_folder):
        msg_blurb = (f"The '{photos_folder}' folder does not exist.")
        msg_blurb_square(msg_blurb, color_red)
        sys.exit(1)  # Exit the program with an error code
    else:
        msg_blurb = (f'photos analyzed from {photos_folder} folder')
        msg_blurb_square(msg_blurb, color_green)

    # Iterate over files in the photos folder
    for filename in os.listdir(photos_folder):
        base_name, extension = os.path.splitext(filename)
        if extension.lower() in image_types:
            file_path = os.path.join(photos_folder, filename)
            try:
                exif_data, Description = read_exif_data(file_path)
                # Append the EXIF data to the sheet
                print(f"{color_green}{filename}{color_reset}")
                sheet.append([exif_data.get(header, "") for header in headers])
            except Exception as e:
                print(f"{color_red}Error processing {filename}: {e}{color_reset}")
                exif_data = {'Name': os.path.basename(file_path)}
                sheet.append([exif_data.get(header, "") for header in headers])

    # Save the workbook
    workbook.save(output_file)

    msg_blurb = (f"File metadata exported to {output_file}")
    msg_blurb_square(msg_blurb, color_green)


# <<<<<<<<<<<<<<<<<<<<<<<<<<   Sub-Routines   >>>>>>>>>>>>>>>>>>>>>>>>>>

def cleanup_description(description):
    cleaned_description = {k: v for k, v in description.items() if v not in ('', None) and k not in ('Icon', 'Type', 'Latitude', 'Longitude')}
    formatted_description = "\n".join(f"{k}: {v}" for k, v in cleaned_description.items())
    return formatted_description


def convert_to_decimal(coord, ref):
    """Converts GPS coordinates to decimal format."""
    if not coord:
        return None
    degrees = float(coord[0].num) / float(coord[0].den)
    minutes = float(coord[1].num) / float(coord[1].den) / 60.0
    seconds = float(coord[2].num) / float(coord[2].den) / 3600.0
    decimal = degrees + minutes + seconds
    if ref in ['S', 'W']:
        decimal = -decimal
    return decimal


def get_file_timestamps(file_path):
    """Returns the creation, last accessed, and last modified times of a file."""
    if not os.path.isfile(file_path):
        raise ValueError(f"The path {file_path} is not a valid file.")

    file_stats = os.stat(file_path)
    creation_time = datetime.fromtimestamp(file_stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
    access_time = datetime.fromtimestamp(file_stats.st_atime).strftime('%Y-%m-%d %H:%M:%S')
    modified_time = datetime.fromtimestamp(file_stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
    return creation_time, access_time, modified_time


def read_exif_data(file_path):
    """Reads selected EXIF data from an image file."""
    with open(file_path, 'rb') as f:
        tags = exifread.process_file(f, stop_tag='UNDEFINED')

    gps_latitude = tags.get('GPS GPSLatitude')
    gps_latitude_ref = tags.get('GPS GPSLatitudeRef')
    gps_longitude = tags.get('GPS GPSLongitude')
    gps_longitude_ref = tags.get('GPS GPSLongitudeRef')
    
    latitude = convert_to_decimal(gps_latitude.values, gps_latitude_ref.values[0]) if gps_latitude and gps_latitude_ref else ""
    longitude = convert_to_decimal(gps_longitude.values, gps_longitude_ref.values[0]) if gps_longitude and gps_longitude_ref else ""

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': tags.get('Image DateTime'),
        'DateTimeOriginal': tags.get('EXIF DateTimeOriginal'),
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': tags.get('EXIF OffsetTime'),
        'DeviceManufacturer': tags.get('Image Make'),
        'FileSize': os.path.getsize(file_path),
        'FileType': os.path.splitext(file_path)[1].replace('.', '').upper(),
        'FileTypeExtension': os.path.splitext(file_path)[1].replace('.', '').lower(),
        'LensMake': tags.get('EXIF LensMake'),
        'Software': tags.get('Image Software'),
        'HostComputer': tags.get('Image HostComputer'),
        'LensInfo': tags.get('EXIF LensInfo'),
        'LensModel': tags.get('EXIF LensModel'),
        'Model': tags.get('Image Model'),
        'Altitude': tags.get('GPS GPSAltitude'),
        'Latitude': latitude,
        'Longitude': longitude,
        'Coordinate': f"{latitude},{longitude}" if latitude and longitude else "",
        'NumberOfImages': tags.get('Exif NumberOfImages'),
        'ExifToolVersion': tags.get('Exif ExifToolVersion'),
        'Icon': 'Images',
        'Type': 'Images',
        'Time': tags.get('EXIF DateTimeOriginal')
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    return exif_data, Description


def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"
    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')


if __name__ == "__main__":
    main()




# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>


"""

0.0.1 - exif data to xlsx
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>


"""
add a menu system / usage


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>


"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
