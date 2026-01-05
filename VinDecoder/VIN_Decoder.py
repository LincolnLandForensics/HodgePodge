import requests # pip install requests
from openpyxl import Workbook   # pip install openpyxl
import time
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

version = '1.0'

# Constants
VIN_API_ENDPOINT = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValuesExtended/{vin}?format=json"
INPUT_FILE = "vins.txt"
OUTPUT_FILE = "vins.xlsx"

# Fields of interest from NHTSA response
FIELDS_TO_EXTRACT = [
    "VIN", "Make", "Model", "ModelYear", "BodyClass",
    "VehicleType", "EngineCylinders", "DisplacementL",
    "FuelTypePrimary", "TransmissionStyle", "PlantCountry"
]

# print(f'try https://berla.co/vehicle-lookup')

class VINDecoderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"VIN Decoder {version}")
        self.root.geometry("600x520")

        # Label
        lbl_intro = tk.Label(root, text="add NCMEC PDFs or Zips into a folder (Defaults to NCMEC folder)", font=("Arial", 10))
        lbl_intro.pack(pady=10)

        # Input File
        frame_input = tk.Frame(root)
        frame_input.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_input, text="Input File:").pack(side=tk.LEFT)
        self.entry_input = tk.Entry(frame_input)
        self.entry_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        current_dir = os.getcwd()
        default_input = os.path.join(current_dir, "vins.txt")
        self.entry_input.insert(0, default_input)
        tk.Button(frame_input, text="Browse", command=self.browse_input).pack(side=tk.LEFT)

        # Output File
        frame_output = tk.Frame(root)
        frame_output.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_output, text="Output File:").pack(side=tk.LEFT)
        self.entry_output = tk.Entry(frame_output)
        self.entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        default_output = os.path.join(current_dir, "vins.xlsx")
        self.entry_output.insert(0, default_output)
        tk.Button(frame_output, text="Browse", command=self.browse_output).pack(side=tk.LEFT)

        # Progress Bar
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=100, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=10)

        # Decode Button
        self.btn_decode = tk.Button(root, text="Decode", command=self.start_thread, bg="green", fg="white", font=("Arial", 12, "bold"))
        self.btn_decode.pack(pady=5)

        # Message Window
        self.text_area = scrolledtext.ScrolledText(root, state='disabled', height=15)
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def browse_input(self):
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select Input File", filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
        if filename:
            self.entry_input.delete(0, tk.END)
            self.entry_input.insert(0, filename)
            # Suggest output filename
            base, _ = os.path.splitext(filename)
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, base + ".xlsx")

    def browse_output(self):
        filename = filedialog.asksaveasfilename(initialdir=os.getcwd(), title="Select Output File", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")), defaultextension=".xlsx")
        if filename:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, filename)

    def start_thread(self):
        input_file = self.entry_input.get()
        output_file = self.entry_output.get()
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"Input file does not exist:\n{input_file}")
            return
            
        self.btn_decode.config(state='disabled')
        self.progress['value'] = 0
        self.log(f"Starting decoding from: {input_file}")
        
        t = threading.Thread(target=self.run_process, args=(input_file, output_file))
        t.daemon = True
        t.start()

    def run_process(self, input_file, output_file):
        try:
            process_vins(input_file, output_file, log_func=self.log, progress_func=self.update_progress)
            self.log("Done")
            self.root.after(0, lambda: messagebox.showinfo("Finished", "Done"))
        except Exception as e:
            self.log(f"An error occurred: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, self.enable_button)

    def enable_button(self):
        self.btn_decode.config(state='normal')

    def update_progress(self, value, maximum=None):
        def _update():
            if maximum:
                self.progress['maximum'] = maximum
            self.progress['value'] = value
        self.root.after(0, _update)

    def log(self, message):
        def _log():
            self.text_area.config(state='normal')
            self.text_area.insert(tk.END, str(message) + "\n")
            self.text_area.see(tk.END)
            self.text_area.config(state='disabled')
        self.root.after(0, _log)

def read_vins(file_path):
    """Reads VINs from a file, one per line."""
    with open(file_path, "r") as file:
        vins = [line.strip() for line in file if line.strip()]
    return vins

def decode_vin(vin):
    """Sends a request to NHTSA VIN decoder and extracts vehicle data."""
    url = VIN_API_ENDPOINT.format(vin=vin)
    response = requests.get(url)
    if response.status_code != 200:
        return {"VIN": vin, "Error": f"HTTP {response.status_code}"}
    
    data = response.json()
    if not data.get("Results"):
        return {"VIN": vin, "Error": "No result returned"}

    result = data["Results"][0]
    decoded = {field: result.get(field, "") for field in FIELDS_TO_EXTRACT}
    return decoded

def write_to_excel(data, output_path):
    """Writes list of dictionaries to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "VIN Data"

    # Write header
    ws.append(FIELDS_TO_EXTRACT)

    # Write data rows
    for entry in data:
        row = [entry.get(field, "") for field in FIELDS_TO_EXTRACT]
        ws.append(row)

    wb.save(output_path)

def process_vins(input_file, output_file, log_func=print, progress_func=None):
    vins = read_vins(input_file)
    total_vins = len(vins)
    if progress_func:
        progress_func(0, total_vins)
        
    results = []
    for i, vin in enumerate(vins):
        log_func(f"Decoding VIN: {vin}")
        try:
            decoded_data = decode_vin(vin)
            results.append(decoded_data)
        except Exception as e:
            log_func(f"Error decoding {vin}: {e}")
            results.append({"VIN": vin, "Error": str(e)})
            
        if progress_func:
            progress_func(i + 1)
            
        time.sleep(1)  # Respectful delay to avoid hammering the API

    write_to_excel(results, output_file)
    log_func(f"VIN decoding complete. Results saved to {output_file}")


def main():
    root = tk.Tk()
    app = VINDecoderGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()


# for item in results:
    # print(f"{item['Variable']}: {item['Value']}")

'''
Certainly. Below is a Python script designed to ingest a plaintext file vins.txt containing a list of 
Vehicle Identification Numbers (VINs), one per line, perform detailed VIN decoding using the 
NHTSA (National Highway Traffic Safety Administration) API, and export the parsed information to an 
Excel file vins.xlsx using the openpyxl library.

This script is engineered with robustness in mind, including response validation and basic fault tolerance.

'''