import os
import zipfile
import tkinter as tk
from tkinter import filedialog, scrolledtext
from eml_parser_core import process_eml_folder, deduplicate_by_sha256, write_xlsx

def extract_zip_recursive(zip_path, extract_to):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)
    except Exception as e:
        print(f"Error extracting {zip_path}: {e}")

def find_email_files(folder):
    email_files = []
    for root, _, files in os.walk(folder):
        for file in files:
            if file.lower().endswith(('.eml', '.mbox', '.json')):
                full_path = os.path.join(root, file)
                email_files.append(full_path)
    return email_files

def run_parser(input_folder, output_file, progress_display, dedup_enabled):

    temp_dir = os.path.join(input_folder, "_unzipped")
    print(f'saving unzipped files to {temp_dir}')   # temp
    os.makedirs(temp_dir, exist_ok=True)

    # Extract ZIPs
    for file in os.listdir(input_folder):
        if file.lower().endswith('.zip'):
            zip_path = os.path.join(input_folder, file)
            print(f'unzipping {zip_path}')  # temp
            zip_extract_path = os.path.join(temp_dir, os.path.splitext(file)[0])
            extract_zip_recursive(zip_path, zip_extract_path)
            progress_display.insert(tk.END, f"📦 Extracted ZIP: {file}\n")
            progress_display.see(tk.END)


    # Collect all email files
    all_emails = []
    for path in [input_folder, temp_dir]:
        all_emails.extend(find_email_files(path))

    if not all_emails:
        progress_display.insert(tk.END, "⚠️ No email files found.\n")
        progress_display.see(tk.END)
        return

    progress_display.insert(tk.END, f"📨 Found {len(all_emails)} email files\n")
    progress_display.see(tk.END)

    # Run parser
    # try:
        # process_eml_folder(all_emails, output_file, progress_display)
        # progress_display.insert(tk.END, f"\n✅ Export complete: {output_file}\n")
        # progress_display.see(tk.END)
    # except Exception as e:
        # progress_display.insert(tk.END, f"❌ Error during parsing: {e}\n")
        # progress_display.see(tk.END)

    try:
        # Step 1: parse all files
        # raw_data, raw_contacts = process_eml_folder(all_emails, output_file=None)
        raw_data, raw_contacts = process_eml_folder(all_emails, output_file=None, progress_display=progress_display)
        # Step 2: deduplicate if needed
        if dedup_enabled:
            print(f'DeDuplicating')   # temp
            raw_data, raw_contacts = deduplicate_by_sha256(raw_data, raw_contacts)
            progress_display.insert(tk.END, f"🧹 Deduplicated to {len(raw_data)} unique emails\n")

        # Step 3: write to Excel
        write_xlsx(raw_data, raw_contacts, output_file)
        progress_display.insert(tk.END, f"\n✅ Export complete: {output_file}\n")
        progress_display.see(tk.END)
    except Exception as e:
        progress_display.insert(tk.END, f"❌ Error during parsing: {e}\n")
        progress_display.see(tk.END)

def launch_gui():
    root = tk.Tk()
    root.title("EML to XLSX Parser")

    # Input folder
    tk.Label(root, text="Input Folder:").grid(row=0, column=0, sticky="w")
    input_entry = tk.Entry(root, width=50)
    input_entry.insert(0, os.path.abspath("email"))
    input_entry.grid(row=0, column=1)



    def browse_input():
        folder = filedialog.askdirectory()
        if folder:
            input_entry.delete(0, tk.END)
            input_entry.insert(0, folder)

    tk.Button(root, text="Browse", command=browse_input).grid(row=0, column=2)

    # Output file
    tk.Label(root, text="Output File:").grid(row=1, column=0, sticky="w")
    output_entry = tk.Entry(root, width=50)
    output_entry.insert(0, "email.xlsx")
    output_entry.grid(row=1, column=1)

    # ✅ DeDuplicate checkbox
    # dedup_var = tk.IntVar(value=0)
    dedup_var = tk.IntVar(value=1)    
    tk.Checkbutton(root, text="DeDuplicate", variable=dedup_var).grid(row=2, column=0, sticky="w")

    # Progress display
    progress_display = scrolledtext.ScrolledText(root, width=80, height=20)
    progress_display.grid(row=3, column=0, columnspan=3, pady=10)

    # Start button
    def start():
        input_folder = input_entry.get()
        output_file = output_entry.get()
        dedup_enabled = dedup_var.get() == 1  # ✅ Now dedup_var is defined
        progress_display.delete(1.0, tk.END)
        run_parser(input_folder, output_file, progress_display, dedup_enabled)

    tk.Button(root, text="Start Parsing", command=start).grid(row=2, column=1, pady=5)

    root.mainloop()

if __name__ == "__main__":
    launch_gui()