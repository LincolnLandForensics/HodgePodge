import os
import sqlite3
import tkinter as tk
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

output_file = "Aleap_latlong.xlsx"

def convert_db_to_xlsx(db_path, xlsx_path, message_box):
    if not db_path:
        db_path = os.path.join(os.getcwd(), "_latlong.db")
    if not xlsx_path:
        xlsx_path = os.path.join(os.getcwd(), output_file)

    message_box.insert(tk.END, f"📁 Converting {db_path} to {xlsx_path}\n")

    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT timestamp AS time, latitude AS Latitude, longitude AS Longitude,
                   activity AS Type, timestamp AS 'Time Original'
            FROM data;
        """)
        rows = cursor.fetchall()
        headers = ["time", "Latitude", "Longitude", "Type", "Time Original",
                   "original_file", "Icon", "group", "Subgroup"]
        # headers = ["#", "Time", "Latitude", "Longitude", "Address", "Group", "Subgroup"
        # , "Description", "Type", "Source", "Deleted", "Tag", "Source file information"
        # , "Service Identifier", "Carved", "Name", "business", "number", "street", "city"
        # , "county", "state", "zipcode", "country", "fulladdress", "query", "Sighting State"
        # , "Plate", "Capture Time", "Capture Network", "Highway Name", "Coordinate"
        # , "Capture Location Latitude", "Capture Location Longitude", "Container"
        # , "Sighting Location", "Direction", "Time Local", "End time", "Category"
        # , "Manually decoded", "Account", "PlusCode", "Time Original", "Timezone", "Icon"
        # , "original_file", "case", "Origin Latitude", "Origin Longitude", "Start Time"
        # , "Azimuth", "Radius", "Altitude", "Location", "time_orig_start", "timezone_start"
        # , "Index", "speed", "parked", "MAC"                   
# ]
                                      
                   
    except Exception as e:
        message_box.insert(tk.END, f"❌ Failed to read database: {e}")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "LatLong"
    ws.freeze_panes = "B2"

    # Write headers
    ws.append(headers)
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="top", horizontal="left")

    # Write data rows with conditional logic
    for row in rows:
        time, lat, lon, Type, time_original = row
        if "." in time:
            time = time.split('.')[0]

        original_file = "_latlong.db from ALEAPP"
        Icon = ""
        group = ""
        Subgroup = ""

        if "Google Photos" in Type:
            Icon = "Images"
            group = "Media"
            Subgroup = "Media"
            if "Remote" in Type:
                group = "Media remote"
        elif "Searched" in Type or "Searches" in Type:
            Icon = "Searched"
            Subgroup = "SearchedPlaces"
        elif "Google Maps Last Trip" in Type:
            Icon = "Car"
            # group = "Media"
            # Subgroup = "Media"



        extended_row = [time, lat, lon, Type, time_original, original_file, Icon, group, Subgroup]
        ws.append(extended_row)
        for col_idx in range(1, len(extended_row) + 1):
            ws.cell(row=ws.max_row, column=col_idx).alignment = Alignment(vertical="top", horizontal="left")

    # Set column widths
    column_widths = [30, 20, 20, 30, 30, 30, 15, 15, 15]
    for i, width in enumerate(column_widths, start=1):
        col_letter = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[col_letter].width = width

    try:
        wb.save(xlsx_path)
        message_box.insert(tk.END, f"✅ Saved to {xlsx_path}")
    except Exception as e:
        message_box.insert(tk.END, f"❌ Failed to save XLSX: {e}")

def start_conversion():
    db_path = input_entry.get().strip()
    xlsx_path = output_entry.get().strip()
    message_box.delete(0, tk.END)
    convert_db_to_xlsx(db_path, xlsx_path, message_box)

# GUI setup
root = tk.Tk()
root.title("LatLong DB to XLSX Converter")

tk.Label(root, text="📎 Convert _latlong.db to Aleap_latlong.xlsx", font=("Arial", 10, "bold")).pack(pady=(5, 0))

frame = tk.Frame(root)
frame.pack(pady=5)

tk.Label(frame, text="Input .db File:").grid(row=0, column=0, sticky="e")
input_entry = tk.Entry(frame, width=50)
input_entry.insert(0, os.path.join(os.getcwd(), "_latlong.db"))
input_entry.grid(row=0, column=1)
tk.Button(frame, text="Browse", command=lambda: input_entry.delete(0, tk.END) or input_entry.insert(0, filedialog.askopenfilename(filetypes=[("SQLite DB", "*.db")]))).grid(row=0, column=2)

tk.Label(frame, text="Output .xlsx File:").grid(row=1, column=0, sticky="e")
output_entry = tk.Entry(frame, width=50)
output_entry.insert(0, os.path.join(os.getcwd(), output_file))
output_entry.grid(row=1, column=1)
tk.Button(frame, text="Browse", command=lambda: output_entry.delete(0, tk.END) or output_entry.insert(0, filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel File", "*.xlsx")]))).grid(row=1, column=2)

tk.Button(root, text="Convert", command=start_conversion, bg="lightblue").pack(pady=10)

message_box = tk.Listbox(root, width=80, height=10)
message_box.pack(pady=5)

root.mainloop()