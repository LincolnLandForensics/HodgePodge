import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook

class CsvToXlsxConverter:
    def __init__(self, root):
        self.root = root
        self.version = "v1.1"
        self.root.title(f"Csv to xlsx conversions {self.version}")
        self.root.geometry("500x450")

        # Set default paths
        self.default_input = os.getcwd()
        self.default_output = os.path.join(self.default_input, "Converted_XLSX")

        self.create_widgets()

    def create_widgets(self):
        # Header
        tk.Label(self.root, text="Convert csv to xlsx files from a folder", font=("Arial", 12, "bold")).pack(pady=10)

        # Input Folder
        tk.Label(self.root, text="Input Folder:").pack(anchor="w", padx=20)
        self.input_entry = tk.Entry(self.root, width=50)
        self.input_entry.insert(0, self.default_input)
        self.input_entry.pack(padx=20, pady=2)
        tk.Button(self.root, text="Browse", command=self.browse_input).pack(anchor="e", padx=20)

        # Output Folder
        tk.Label(self.root, text="Output Folder:").pack(anchor="w", padx=20)
        self.output_entry = tk.Entry(self.root, width=50)
        self.output_entry.insert(0, self.default_output)
        self.output_entry.pack(padx=20, pady=2)
        tk.Button(self.root, text="Browse", command=self.browse_output).pack(anchor="e", padx=20)

        # Delimiter
        tk.Label(self.root, text="Delimiter:").pack(anchor="w", padx=20)
        self.delim_entry = tk.Entry(self.root, width=5)
        self.delim_entry.insert(0, ",")
        self.delim_entry.pack(anchor="w", padx=20, pady=5)

        # Progress Bar
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=10)

        # Log Window
        self.log_text = tk.Text(self.root, height=8, width=55, font=("Consolas", 9))
        self.log_text.pack(padx=20, pady=5)

        # Convert Button
        tk.Button(self.root, text="Convert", bg="#4CAF50", fg="white", height=2, width=15, command=self.start_conversion).pack(pady=10)

    def browse_input(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, folder)

    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def start_conversion(self):
        input_dir = self.input_entry.get()
        output_dir = self.output_entry.get()
        delimiter = self.delim_entry.get()

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Filter for CSV files
        files = [f for f in os.listdir(input_dir) if f.lower().endswith('.csv')]
        
        if not files:
            messagebox.showinfo("No Files", "No CSV files found in the input folder.")
            return

        self.progress["maximum"] = len(files)
        self.progress["value"] = 0
        self.log_text.delete(1.0, tk.END)
        self.log(f"Found {len(files)} files. Starting conversion...")

        for idx, filename in enumerate(files):
            try:
                # 1. Clean file name: Replace spaces with underscore
                clean_name = filename.replace(" ", "_")
                base_name = os.path.splitext(clean_name)[0]
                output_path = os.path.join(output_dir, f"{base_name}.xlsx")

                # 2. Create Workbook and read CSV data
                wb = Workbook()
                ws = wb.active
                ws.title = base_name[:31]  # Excel tab limit is 31 chars

                csv_path = os.path.join(input_dir, filename)
                
                # Use 'utf-8-sig' to handle files with BOM (common in Excel CSVs)
                with open(csv_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    for row in reader:
                        ws.append(row)

                # 3. Freeze panes by cell B2
                ws.freeze_panes = "B2"

                # 4. Save file
                wb.save(output_path)

                self.log(f"Success: {filename} -> {base_name}.xlsx")
                self.progress["value"] = idx + 1
            except Exception as e:
                self.log(f"Error processing {filename}: {str(e)}")

        self.log("--- Done ---")
        messagebox.showinfo("Status", "Done")

if __name__ == "__main__":
    root = tk.Tk()
    app = CsvToXlsxConverter(root)
    root.mainloop()