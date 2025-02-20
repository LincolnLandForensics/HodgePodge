import sys
import os
import gzip
import shutil

def process_wigle_file(filename):
    if filename.endswith('.gz'):
        unzipped_filename = filename[:-3]  # Remove .gz extension
        try:
            with gzip.open(filename, 'rb') as f_in:
                with open(unzipped_filename, 'wb') as f_out:
                    shutil.copyfileobj(f_in, f_out)
            print(f"Unzipped: {filename} -> {unzipped_filename}")
            filename = unzipped_filename
        except Exception as e:
            print(f"Error unzipping file: {e}")
            return
    
    if not os.path.isfile(filename):
        print(f"Error: File '{filename}' not found or is not a valid file.")
        return
    
    if not filename.startswith('WigleWifi') or not filename.endswith('.csv'):
        print("Invalid file: Filename must start with 'WigleWifi' and end with '.csv'")
        return
    
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        if lines and "WigleWifi" in lines[0]:
            lines = lines[1:]  # Remove the first line
        
        # Insert the new header
        if lines and "Altitudetemp" not in lines[0]:
            lines.insert(0, "Name,#,Description,Time,Group,Subgroup,Sighting State,Latitude,Longitude,Altitudetemp,RadiusTemp,RCOIs,Source file information,Icon\n")

        if lines and "FirstSeen" in lines[1]:
            del lines[1]  # Remove the second line (index 1)
  
        with open(filename, 'w', encoding='utf-8') as f:
            f.writelines(lines)

        print(f"Processed: {filename}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python WigleWifi_parser.py <filename>")
        print("Example: python WigleWifi_parser.py WigleWifi_sample3.csv")
        sys.exit(1)

    filename = sys.argv[1]
    process_wigle_file(filename)


'''
python WigleWifi_parser.py WigleWifi_sample3.csv

'''