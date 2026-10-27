#!/usr/bin/python
# coding: utf-8

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

import os
import sys
import exifread   # pip install exifread
import hashlib   # pip install hashlib
import openpyxl   # pip install openpyx
import argparse # for menu system
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import zipfile
import xml.etree.ElementTree as ET
import pypdf

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "forensic metadata and exif reader"
version = '1.0.1'

global image_types
global doc_types
global opendoc_types
global pdf_types
global rtf_types
global epub_types
global geo_types
global media_types

image_types = [
    '.jpg', '.jpeg', '.png', '.heic', '.heif', '.tiff', '.tif', '.webp', '.avif', '.gif', '.bmp',
    '.cr2', '.cr3', '.nef', '.arw', '.dng', '.orf', '.rw2', '.raf'
]
doc_types = ['.docx', '.xlsx', '.pptx']
opendoc_types = ['.odt', '.ods', '.odp']
pdf_types = ['.pdf']
rtf_types = ['.rtf']
epub_types = ['.epub']
geo_types = ['.gpx', '.kml', '.kmz']
media_types = [
    '.mp4', '.m4v', '.mov', '.3gp', '.3g2', '.avi', '.mkv', '.mts', '.m2ts', 
    '.m4a', '.wav', '.mp3', '.flac', '.ogg'
]


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

# <<<<<<<<<<<<<<<<<<<<<<<<<<      GUI Class      >>>>>>>>>>>>>>>>>>>>>>>>>>

class ExifGui:
    def __init__(self, root):
        self.root = root
        script_name = os.path.basename(sys.argv[0])
        self.root.title(f"{script_name} {version}")
        self.root.geometry("700x550")
        
        # Set Vista theme
        self.style = ttk.Style()
        try:
            self.style.theme_use('vista')
        except:
            pass # Fallback to default if vista is not available

        self.setup_gui()
        self.set_defaults()

    def setup_gui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title/Description
        ttk.Label(main_frame, text=f"{description}", font=("Helvetica", 12, "bold")).pack(pady=5)

        # Input Folder
        input_frame = ttk.LabelFrame(main_frame, text="Input Folder", padding="5")
        input_frame.pack(fill=tk.X, pady=5)
        
        self.input_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.input_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).pack(side=tk.RIGHT)

        # Output File
        output_frame = ttk.LabelFrame(main_frame, text="Output Excel File", padding="5")
        output_frame.pack(fill=tk.X, pady=5)
        
        self.output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.output_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_frame, text="Browse", command=self.browse_output).pack(side=tk.RIGHT)

        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)

        # Status/Log Window
        self.status_box = ScrolledText(main_frame, height=15, state='disabled')
        self.status_box.pack(fill=tk.BOTH, expand=True, pady=5)

        # Extract Button
        self.btn_extract = ttk.Button(main_frame, text="Extract MetaData", command=self.start_processing)
        self.btn_extract.pack(pady=10)

    def set_defaults(self):
        self.input_var.set("photos")
        timestamp = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
        self.output_var.set(f"exif_data_{timestamp}.xlsx")

    def browse_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_var.set(folder)

    def browse_output(self):
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.output_var.set(file)

    def log(self, message):
        self.status_box.config(state='normal')
        self.status_box.insert(tk.END, message + "\n")
        self.status_box.see(tk.END)
        self.status_box.config(state='disabled')
        self.root.update_idletasks()

    def start_processing(self):
        input_folder = self.input_var.get()
        output_file = self.output_var.get()

        if not os.path.exists(input_folder):
            messagebox.showerror("Error", f"Input folder '{input_folder}' does not exist.")
            return

        self.btn_extract.config(state='disabled')
        self.progress_var.set(0)
        self.status_box.config(state='normal')
        self.status_box.delete(1.0, tk.END)
        self.status_box.config(state='disabled')
        
        self.log(f"Input Folder: {os.path.abspath(input_folder)}")

        # Start processing in a new thread
        thread = threading.Thread(target=self.run_process, args=(input_folder, output_file))
        thread.daemon = True
        thread.start()

    def run_process(self, input_folder, output_file):
        try:
            process_exif(input_folder, output_file, self.log, self.update_progress)
            self.log(f"Output File: {output_file}")
            self.log("\nDone.")
        except Exception as e:
            self.log(f"Critical Error: {e}")
        finally:
            self.btn_extract.config(state='normal')

    def update_progress(self, current, total):
        if total > 0:
            percent = (current / total) * 100
            self.progress_var.set(percent)
        else:
            self.progress_var.set(100)


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Menu/Logic     >>>>>>>>>>>>>>>>>>>>>>>>>>

def process_exif(photos_folder, output_file, status_callback=None, progress_callback=None):
    """
    Core logic to extract EXIF data and save to Excel.
    """
    # Ensure the input folder exists
    if not os.path.exists(photos_folder):
        msg = f"Input folder {photos_folder} doesn't exist."
        if status_callback: status_callback(msg)
        msg_blurb_square(msg, color_red)
        return False

    # Create a new workbook and add basic formatting
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "FileMetadata"

    # Define headers
    headers = [
        'Name', 'DateCreated', 'DateTimeOriginal', 'FileCreateDate', 'FileModifyDate', 'Timezone'
        ,'DeviceManufacturer', 'ExifToolVersion', 'FileSize', 'FileType', 'FileTypeExtension', 'LensMake'
        ,'Software', 'HostComputer', 'LensInfo', 'LensModel', 'Model', 'NumberOfImages'
        ,'Altitude', 'Latitude', 'Longitude', 'Coordinate', 'Time', 'Type', 'Icon', 'Description', 'MD5'
        , 'Copyright', 'Creator', 'Author', 'LastModifiedBy', 'Path'
    ]
    sheet.append(headers)

    for col, header in enumerate(headers, start=1):
        sheet.cell(row=1, column=col, value=header).font = openpyxl.styles.Font(bold=True)
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15

    sheet.freeze_panes = 'B2'

    files = [f for f in os.listdir(photos_folder) if os.path.isfile(os.path.join(photos_folder, f))]
    total_files = len(files)
    
    msg = f'Analyzing {total_files} files from {photos_folder} folder'
    if status_callback: status_callback(msg)
    msg_blurb_square(msg, color_green)

    for i, filename in enumerate(files, 1):
        file_path = os.path.join(photos_folder, filename)
        try:
            exif_data, Description = read_exif_data(file_path)
            print(f"{color_green}{filename}{color_reset}")
            if status_callback: status_callback(filename)
            sheet.append([exif_data.get(header, "") for header in headers])
        except Exception as e:
            err_msg = f"Error processing {filename}: {e}"
            print(f"{color_red}{err_msg}{color_reset}")
            if status_callback: status_callback(err_msg)
            exif_data = {'Name': os.path.basename(file_path)}
            sheet.append([exif_data.get(header, "") for header in headers])
        
        if progress_callback:
            progress_callback(i, total_files)

    workbook.save(output_file)
    msg = f"File metadata exported to {output_file}"
    if status_callback: status_callback(msg)
    msg_blurb_square(msg, color_green)
    return True

    
def main():
    """
    Main function to parse arguments and initiate file conversion.
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='Input folder', required=False)
    parser.add_argument('-e', '--exif', help='extract exif data from photos', required=False, action='store_true')

    args = parser.parse_args()

    # If no arguments are provided, launch GUI
    if len(sys.argv) == 1:
        root = tk.Tk()
        gui = ExifGui(root)
        root.mainloop()
        return

    # CLI Mode
    photos_folder = args.input if args.input else 'photos'
    timestamp = datetime.now().strftime("%Y-%m-%d_%H_%M_%S")
    output_file = f'exif_data_{timestamp}.xlsx'

    process_exif(photos_folder, output_file)


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


def hashfile(file_path):
    """
    Computes and returns the MD5 hash of a file.
    """
    # Create an MD5 hash object
    hash_md5 = hashlib.md5()

    # Open the file in binary read mode
    with open(file_path, "rb") as f:
        # Read the file in chunks to avoid memory overload with large files
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    
    # Return the hexadecimal representation of the hash
    return hash_md5.hexdigest()
    

def get_ooxml_metadata(file_path):
    """
    Parses metadata from OOXML files (.docx, .xlsx, .pptx) without external dependencies.
    """
    metadata = {}
    
    def find_element(root, local_name):
        # Standard namespaces
        ns = {
            'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'dc': 'http://purl.org/dc/elements/1.1/',
            'dcterms': 'http://purl.org/dc/terms/',
            'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
        }
        for prefix in ns:
            elem = root.find(f"{prefix}:{local_name}", ns)
            if elem is not None:
                return elem
        # Fallback to tag suffix matching
        for elem in root.iter():
            tag = elem.tag
            if tag.endswith('}' + local_name) or tag == local_name:
                return elem
        return None

    def get_val(root, local_name):
        elem = find_element(root, local_name)
        return elem.text if elem is not None else None

    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            # Parse core.xml
            if 'docProps/core.xml' in zip_ref.namelist():
                core_xml = zip_ref.read('docProps/core.xml')
                root = ET.fromstring(core_xml)
                metadata['creator'] = get_val(root, 'creator')
                metadata['lastModifiedBy'] = get_val(root, 'lastModifiedBy')
                metadata['created'] = get_val(root, 'created')
                metadata['modified'] = get_val(root, 'modified')
                metadata['title'] = get_val(root, 'title')
                metadata['subject'] = get_val(root, 'subject')
                metadata['description'] = get_val(root, 'description')
                metadata['keywords'] = get_val(root, 'keywords')

            # Parse app.xml
            if 'docProps/app.xml' in zip_ref.namelist():
                app_xml = zip_ref.read('docProps/app.xml')
                root = ET.fromstring(app_xml)
                metadata['application'] = get_val(root, 'Application')
                metadata['appVersion'] = get_val(root, 'AppVersion')
                metadata['company'] = get_val(root, 'Company')
                metadata['template'] = get_val(root, 'Template')
                metadata['totalTime'] = get_val(root, 'TotalTime')
                metadata['pages'] = get_val(root, 'Pages')
                metadata['words'] = get_val(root, 'Words')
                metadata['characters'] = get_val(root, 'Characters')
    except Exception as e:
        print(f"Error parsing OOXML file {file_path}: {e}")
    return metadata


def parse_mp4_timestamps(file_path):
    """
    Extracts embedded creation and modification timestamps from QuickTime/ISO BMFF 
    containers (MP4, MOV, 3GP, M4V, M4A) by reading the 'mvhd' atom directly.
    """
    creation_time = None
    modification_time = None
    try:
        with open(file_path, 'rb') as f:
            data = f.read(1024 * 1024) # check first 1MB
            idx = data.find(b'mvhd')
            if idx != -1:
                version = data[idx + 4]
                time_offset = idx + 8
                if version == 0:
                    c_time = int.from_bytes(data[time_offset : time_offset + 4], 'big')
                    m_time = int.from_bytes(data[time_offset + 4 : time_offset + 8], 'big')
                elif version == 1:
                    c_time = int.from_bytes(data[time_offset : time_offset + 8], 'big')
                    m_time = int.from_bytes(data[time_offset + 8 : time_offset + 16], 'big')
                else:
                    return None, None
                
                epoch_1904 = datetime(1904, 1, 1)
                if c_time > 0:
                    creation_time = (epoch_1904 + timedelta(seconds=c_time)).strftime('%Y-%m-%d %H:%M:%S')
                if m_time > 0:
                    modification_time = (epoch_1904 + timedelta(seconds=m_time)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        print(f"Error parsing QuickTime timestamps: {e}")
    return creation_time, modification_time


def get_pdf_metadata(file_path):
    """Parses PDF metadata using the pypdf library."""
    metadata = {}
    try:
        reader = pypdf.PdfReader(file_path)
        meta = reader.metadata
        if meta:
            metadata['creator'] = meta.creator
            metadata['producer'] = meta.producer
            metadata['author'] = meta.author
            metadata['title'] = meta.title
            metadata['subject'] = meta.subject
            metadata['creation_date'] = meta.creation_date
            metadata['modification_date'] = meta.modification_date
    except Exception as e:
        print(f"Error parsing PDF metadata: {e}")
    return metadata


def get_gpx_metadata(file_path):
    """Parses GPX tracking files."""
    metadata = {}
    try:
        with open(file_path, 'rb') as f:
            root = ET.fromstring(f.read())
        ns = {'g': 'http://www.topografix.com/GPX/1/1'}
        def get_val(local_name):
            elem = root.find(f".//g:{local_name}", ns) or root.find(f".//{local_name}")
            return elem.text if elem is not None else None
        
        metadata['name'] = get_val('name')
        metadata['creator'] = root.get('creator')
        metadata['time'] = get_val('time')
        
        trks = root.findall(".//g:trkpt", ns) or root.findall(".//trkpt")
        wpts = root.findall(".//g:wpt", ns) or root.findall(".//wpt")
        metadata['trackpoints_count'] = len(trks)
        metadata['waypoints_count'] = len(wpts)
    except Exception as e:
        print(f"Error parsing GPX metadata: {e}")
    return metadata


def get_kml_metadata(file_path):
    """Parses KML and KMZ tracking containers."""
    metadata = {}
    try:
        kml_content = None
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.kmz':
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for name in zip_ref.namelist():
                    if name.endswith('.kml'):
                        kml_content = zip_ref.read(name)
                        break
        else:
            with open(file_path, 'rb') as f:
                kml_content = f.read()
                
        if kml_content:
            root = ET.fromstring(kml_content)
            ns = {'k': 'http://www.opengis.net/kml/2.2'}
            def get_val(local_name):
                elem = root.find(f".//k:{local_name}", ns) or root.find(f".//{local_name}")
                return elem.text if elem is not None else None
            
            metadata['name'] = get_val('name')
            metadata['description'] = get_val('description')
            
            placemarks = root.findall(".//k:Placemark", ns) or root.findall(".//Placemark")
            metadata['placemarks_count'] = len(placemarks)
    except Exception as e:
        print(f"Error parsing KML/KMZ metadata: {e}")
    return metadata


def get_epub_metadata(file_path):
    """Parses EPUB book metadata."""
    metadata = {}
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            opf_file = None
            for name in zip_ref.namelist():
                if name.endswith('.opf'):
                    opf_file = name
                    break
            if opf_file:
                opf_content = zip_ref.read(opf_file)
                root = ET.fromstring(opf_content)
                metadata_elem = root.find('.//{http://www.idpf.org/2007/opf}metadata') or root.find('.//metadata')
                if metadata_elem is not None:
                    def get_val(local_name):
                        elem = metadata_elem.find(f".//{{http://purl.org/dc/elements/1.1/}}{local_name}") or metadata_elem.find(f".//{local_name}")
                        return elem.text if elem is not None else None
                    
                    metadata['creator'] = get_val('creator')
                    metadata['title'] = get_val('title')
                    metadata['date'] = get_val('date')
                    metadata['publisher'] = get_val('publisher')
                    metadata['language'] = get_val('language')
                    metadata['subject'] = get_val('subject')
    except Exception as e:
        print(f"Error parsing EPUB metadata: {e}")
    return metadata


def get_opendoc_metadata(file_path):
    """Parses OpenDocument meta.xml metadata."""
    metadata = {}
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            if 'meta.xml' in zip_ref.namelist():
                meta_xml = zip_ref.read('meta.xml')
                root = ET.fromstring(meta_xml)
                ns = {
                    'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
                    'meta': 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
                    'dc': 'http://purl.org/dc/elements/1.1/',
                }
                office_meta = root.find('office:meta', ns)
                if office_meta is not None:
                    def get_val(path):
                        elem = office_meta.find(path, ns)
                        return elem.text if elem is not None else None
                    
                    metadata['creator'] = get_val('dc:creator')
                    metadata['creation_date'] = get_val('meta:creation-date')
                    metadata['date'] = get_val('dc:date')
                    metadata['generator'] = get_val('meta:generator')
                    metadata['editing_duration'] = get_val('meta:editing-duration')
                    
                    stats_elem = office_meta.find('meta:document-statistic', ns)
                    if stats_elem is not None:
                        metadata['words'] = stats_elem.get(f"{{{ns['meta']}}}word-count")
                        metadata['pages'] = stats_elem.get(f"{{{ns['meta']}}}page-count")
                        metadata['characters'] = stats_elem.get(f"{{{ns['meta']}}}character-count")
    except Exception as e:
        print(f"Error parsing OpenDocument metadata: {e}")
    return metadata


def read_image_metadata(file_path, hash_md5):
    """Reads EXIF data from an image file."""
    tags = {}
    try:
        with open(file_path, 'rb') as f:
            tags = exifread.process_file(f, stop_tag='UNDEFINED')
    except Exception as e:
        print(f"Error parsing EXIF with exifread: {e}")

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
        'Time': tags.get('EXIF DateTimeOriginal'),
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': tags.get('Copyright'),
        'Creator': tags.get('Image Artist'),
        'LastModifiedBy': '',
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_docx_metadata(file_path, hash_md5):
    """Reads metadata from an OOXML file (.docx, .xlsx, .pptx)."""
    meta = get_ooxml_metadata(file_path)
    
    def format_ooxml_date(date_str):
        if not date_str:
            return ""
        try:
            date_str_clean = date_str.replace('Z', '')
            if '.' in date_str_clean:
                date_str_clean = date_str_clean.split('.')[0]
            dt = datetime.strptime(date_str_clean, '%Y-%m-%dT%H:%M:%S')
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
            return date_str

    date_created = format_ooxml_date(meta.get('created'))
    date_modified = format_ooxml_date(meta.get('modified'))

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': date_created,
        'DateTimeOriginal': date_modified,
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': os.path.splitext(file_path)[1].replace('.', '').upper(),
        'FileTypeExtension': os.path.splitext(file_path)[1].replace('.', '').lower(),
        'LensMake': '',
        'Software': meta.get('application', ''),
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': 'Documents',
        'Type': 'Documents',
        'Time': date_modified if date_modified else date_created,
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        
        # Additional fields captured in Description
        'Creator': meta.get('creator', ''),
        'LastModifiedBy': meta.get('lastModifiedBy', ''),
        'AppVersion': meta.get('appVersion', ''),
        'Company': meta.get('company', ''),
        'Template': meta.get('template', ''),
        'TotalTime': meta.get('totalTime', ''),
        'Pages': meta.get('pages', ''),
        'Words': meta.get('words', ''),
        'Characters': meta.get('characters', ''),
        'Title': meta.get('title', ''),
        'Subject': meta.get('subject', ''),
        'Keywords': meta.get('keywords', ''),
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_opendoc_metadata(file_path, hash_md5):
    """Reads metadata from OpenDocument formats (.odt, .ods, .odp)."""
    ext = os.path.splitext(file_path)[1].lower()
    meta = get_opendoc_metadata(file_path)
    
    def format_opendoc_date(date_str):
        if not date_str:
            return ""
        try:
            date_str_clean = date_str.replace('Z', '')
            if 'T' in date_str_clean:
                if '.' in date_str_clean:
                    date_str_clean = date_str_clean.split('.')[0]
                dt = datetime.strptime(date_str_clean, '%Y-%m-%dT%H:%M:%S')
                return dt.strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
            pass
        return date_str

    date_created = format_opendoc_date(meta.get('creation_date'))
    date_modified = format_opendoc_date(meta.get('date'))

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': date_created,
        'DateTimeOriginal': date_modified,
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': ext.replace('.', '').upper(),
        'FileTypeExtension': ext.replace('.', '').lower(),
        'LensMake': '',
        'Software': meta.get('generator', ''),
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': 'Documents',
        'Type': 'Documents',
        'Time': date_modified if date_modified else date_created,
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        'Creator': meta.get('creator', ''),
        'LastModifiedBy': '',
        'EditingDuration': meta.get('editing_duration', ''),
        'Pages': meta.get('pages', ''),
        'Words': meta.get('words', ''),
        'Characters': meta.get('characters', '')
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_epub_metadata(file_path, hash_md5):
    """Reads metadata from EPUB electronic books."""
    ext = os.path.splitext(file_path)[1].lower()
    meta = get_epub_metadata(file_path)
    
    date_val = meta.get('date', '')
    if date_val and 'T' in date_val:
        try:
            date_val_clean = date_val.replace('Z', '').split('.')[0]
            dt = datetime.strptime(date_val_clean, '%Y-%m-%dT%H:%M:%S')
            date_val = dt.strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
            pass

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': date_val,
        'DateTimeOriginal': '',
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': 'EPUB',
        'FileTypeExtension': 'epub',
        'LensMake': '',
        'Software': meta.get('publisher', ''),
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': 'Documents',
        'Type': 'Documents',
        'Time': date_val if date_val else datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        'Creator': meta.get('creator', ''),
        'LastModifiedBy': '',
        'Title': meta.get('title', ''),
        'Publisher': meta.get('publisher', ''),
        'Language': meta.get('language', ''),
        'Subject': meta.get('subject', '')
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_pdf_metadata(file_path, hash_md5):
    """Reads metadata from PDF documents using pypdf."""
    ext = os.path.splitext(file_path)[1].lower()
    meta = get_pdf_metadata(file_path)
    
    def format_pdf_date(dt):
        if not dt:
            return ""
        if isinstance(dt, datetime):
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        dt_str = str(dt)
        if dt_str.startswith('D:'):
            try:
                raw_dt = dt_str[2:16]
                parsed = datetime.strptime(raw_dt, '%Y%m%d%H%M%S')
                return parsed.strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                pass
        return dt_str

    date_created = format_pdf_date(meta.get('creation_date'))
    date_modified = format_pdf_date(meta.get('modification_date'))

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': date_created,
        'DateTimeOriginal': date_modified,
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': 'PDF',
        'FileTypeExtension': 'pdf',
        'LensMake': '',
        'Software': meta.get('producer', ''),
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': 'Documents',
        'Type': 'Documents',
        'Time': date_modified if date_modified else date_created,
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        'Creator': meta.get('creator', '') or meta.get('author', ''),
        'LastModifiedBy': '',
        'Title': meta.get('title', ''),
        'Subject': meta.get('subject', ''),
        'Producer': meta.get('producer', ''),
        'Author': meta.get('author', '')
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_rtf_metadata(file_path, hash_md5):
    """Reads metadata from RTF documents by regex searching the info block."""
    ext = os.path.splitext(file_path)[1].lower()
    author, company, title, subject = '', '', '', ''
    c_date, r_date = '', ''
    
    try:
        with open(file_path, 'rb') as f:
            header_data = f.read(32768)
            import re
            
            def extract_field(tag):
                match = re.search(rb'\\' + tag.encode() + rb'\s+([^\\}]+)', header_data)
                return match.group(1).decode('utf-8', errors='ignore').strip() if match else ''

            author = extract_field('author')
            company = extract_field('company')
            title = extract_field('title')
            subject = extract_field('subject')
            
            def extract_date(date_tag):
                date_match = re.search(rb'\\' + date_tag.encode() + rb'\\yr(\d+)\\mo(\d+)\\dy(\d+)\\hr(\d+)\\min(\d+)', header_data)
                if date_match:
                    parts = [int(p) for p in date_match.groups()]
                    try:
                        dt = datetime(parts[0], parts[1], parts[2], parts[3], parts[4])
                        return dt.strftime('%Y-%m-%d %H:%M:%S')
                    except Exception:
                        pass
                return ''

            c_date = extract_date('creatim')
            r_date = extract_date('revtim')
    except Exception as e:
        print(f"Error parsing RTF {file_path}: {e}")

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': c_date,
        'DateTimeOriginal': r_date,
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': 'RTF',
        'FileTypeExtension': 'rtf',
        'LensMake': '',
        'Software': '',
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': 'Documents',
        'Type': 'Documents',
        'Time': r_date if r_date else c_date,
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        'Creator': author,
        'LastModifiedBy': '',
        'Title': title,
        'Subject': subject,
        'Company': company
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_geo_metadata(file_path, hash_md5):
    """Reads metadata from GPX and KML/KMZ geolocation files."""
    ext = os.path.splitext(file_path)[1].lower()
    name, description_str, creator = '', '', ''
    time_val = ''
    trackpoints_count, waypoints_count, placemarks_count = 0, 0, 0
    
    if ext == '.gpx':
        meta = get_gpx_metadata(file_path)
        name = meta.get('name', '')
        creator = meta.get('creator', '')
        time_val = meta.get('time', '')
        trackpoints_count = meta.get('trackpoints_count', 0)
        waypoints_count = meta.get('waypoints_count', 0)
    elif ext in ['.kml', '.kmz']:
        meta = get_kml_metadata(file_path)
        name = meta.get('name', '')
        description_str = meta.get('description', '')
        placemarks_count = meta.get('placemarks_count', 0)

    if time_val and 'T' in time_val:
        try:
            time_val_clean = time_val.replace('Z', '').split('.')[0]
            dt = datetime.strptime(time_val_clean, '%Y-%m-%dT%H:%M:%S')
            time_val = dt.strftime('%Y-%m-%d %H:%M:%S')
        except Exception:
            pass

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': time_val,
        'DateTimeOriginal': '',
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': ext.replace('.', '').upper(),
        'FileTypeExtension': ext.replace('.', '').lower(),
        'LensMake': '',
        'Software': '',
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': 'Tracking',
        'Type': 'Tracking',
        'Time': time_val if time_val else datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        'Creator': creator,
        'LastModifiedBy': '',
        'GeoName': name,
        'GeoDescription': description_str,
        'TrackPointsCount': trackpoints_count,
        'WayPointsCount': waypoints_count,
        'PlacemarksCount': placemarks_count
    }

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_media_metadata(file_path, hash_md5):
    """Reads metadata from video or audio files."""
    ext = os.path.splitext(file_path)[1].lower()
    c_time, m_time = None, None
    if ext in ['.mp4', '.m4v', '.mov', '.3gp', '.3g2', '.m4a']:
        c_time, m_time = parse_mp4_timestamps(file_path)

    is_audio = ext in ['.m4a', '.wav', '.mp3', '.flac', '.ogg']
    icon_type = 'Audio' if is_audio else 'Video'

    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': c_time if c_time else '',
        'DateTimeOriginal': m_time if m_time else '',
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': ext.replace('.', '').upper(),
        'FileTypeExtension': ext.replace('.', '').lower(),
        'LensMake': '',
        'Software': '',
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': icon_type,
        'Type': icon_type,
        'Time': m_time if m_time else (c_time if c_time else datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')),
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        'Creator': '',
        'LastModifiedBy': ''
    }

    if ext == '.mp3':
        try:
            with open(file_path, 'rb') as f:
                f.seek(-128, 2)
                tag_data = f.read(128)
                if tag_data[:3] == b'TAG':
                    title = tag_data[3:33].strip(b'\x00\r\n\t ').decode('utf-8', errors='ignore')
                    artist = tag_data[33:63].strip(b'\x00\r\n\t ').decode('utf-8', errors='ignore')
                    album = tag_data[63:93].strip(b'\x00\r\n\t ').decode('utf-8', errors='ignore')
                    if title: exif_data['Title'] = title
                    if artist: exif_data['Creator'] = artist
                    if album: exif_data['Album'] = album
        except Exception:
            pass

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_generic_metadata(file_path, hash_md5):
    """Reads basic metadata from any generic file."""
    ext = os.path.splitext(file_path)[1].lower()
    
    exif_data = {
        'Name': os.path.basename(file_path),
        'DateCreated': '',
        'DateTimeOriginal': '',
        'FileCreateDate': datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'FileModifyDate': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Timezone': '',
        'DeviceManufacturer': '',
        'FileSize': os.path.getsize(file_path),
        'FileType': ext.replace('.', '').upper() if ext else "UNKNOWN",
        'FileTypeExtension': ext.replace('.', '').lower() if ext else "",
        'LensMake': '',
        'Software': '',
        'HostComputer': '',
        'LensInfo': '',
        'LensModel': '',
        'Model': '',
        'Altitude': '',
        'Latitude': '',
        'Longitude': '',
        'Coordinate': '',
        'NumberOfImages': '',
        'ExifToolVersion': '',
        'Icon': 'Files',
        'Type': 'Files',
        'Time': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
        'Description': '',        
        'MD5': hash_md5,
        'Copyright': '',
        'Creator': '',
        'Author': '',        
        'LastModifiedBy': ''
    }

    if ext == '.json':
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                import json
                jdata = json.load(f)
                snippet = {k: v for i, (k, v) in enumerate(jdata.items()) if i < 30}
                for k, v in snippet.items():
                    exif_data[f"JSON_{k}"] = str(v)
        except Exception:
            pass

    for key, value in exif_data.items():
        exif_data[key] = str(value) if value else ""
        
    Description = cleanup_description(exif_data)
    exif_data['Description'] = Description
    return exif_data, Description


def read_exif_data(file_path):
    """Unified entry point to read metadata from any file type."""
    try:
        hash_md5 = hashfile(file_path)
    except Exception as e:
        print(f"Error hashing '{file_path}': {e}")        
        hash_md5 = ""

    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in doc_types:
        return read_docx_metadata(file_path, hash_md5)
    elif ext in opendoc_types:
        return read_opendoc_metadata(file_path, hash_md5)
    elif ext in epub_types:
        return read_epub_metadata(file_path, hash_md5)
    elif ext in pdf_types:
        return read_pdf_metadata(file_path, hash_md5)
    elif ext in rtf_types:
        return read_rtf_metadata(file_path, hash_md5)
    elif ext in geo_types:
        return read_geo_metadata(file_path, hash_md5)
    elif ext in media_types:
        return read_media_metadata(file_path, hash_md5)
    elif ext in image_types:
        return read_image_metadata(file_path, hash_md5)
    else:
        return read_generic_metadata(file_path, hash_md5)


def msg_blurb_square(msg_blurb, color):
    horizontal_line = f"+{'-' * (len(msg_blurb) + 2)}+"
    empty_line = f"| {' ' * (len(msg_blurb))} |"
    print(color + horizontal_line)
    print(empty_line)
    print(f"| {msg_blurb} |")
    print(empty_line)
    print(horizontal_line)
    print(f'{color_reset}')


def Usage():
    file = sys.argv[0].split('\\')[-1]

    print(f'\nDescription: {color_green}{description}{color_reset}')
    print(f'{file} Version: {version} by {author}')
    print(f"    {file}")
    print(f"    {file} -I photos")



if __name__ == "__main__":
    main()




# <<<<<<<<<<<<<<<<<<<<<<<<<< Revision History >>>>>>>>>>>>>>>>>>>>>>>>>>


"""
0.2.0 - added menu and usage
0.0.1 - exif data to xlsx
"""

# <<<<<<<<<<<<<<<<<<<<<<<<<< Future Wishlist  >>>>>>>>>>>>>>>>>>>>>>>>>>


"""

hash each file

"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>


"""


"""

# <<<<<<<<<<<<<<<<<<<<<<<<<<      The End        >>>>>>>>>>>>>>>>>>>>>>>>>>
