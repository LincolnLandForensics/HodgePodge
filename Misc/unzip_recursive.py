import zipfile
import os
import io
import sys
import argparse
import threading
import time
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Imports        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Color support for Windows 11+
color_red = color_yellow = color_green = color_blue = color_purple = color_reset = ''
if sys.version_info > (3, 7, 9) and os.name == "nt":
    try:
        from colorama import Fore, Back, Style
        color_red, color_yellow, color_green = Fore.RED, Fore.YELLOW, Fore.GREEN
        color_blue, color_purple, color_reset = Fore.BLUE, Fore.MAGENTA, Style.RESET_ALL
    except ImportError:
        pass

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Pre-Sets       >>>>>>>>>>>>>>>>>>>>>>>>>>

author = 'LincolnLandForensics'
description = "unzip zip's inside of zips"
version = '0.1.2'

# <<<<<<<<<<<<<<<<<<<<<<<<<<      GUI Class      >>>>>>>>>>>>>>>>>>>>>>>>>>

class UnzipGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(f"unzip_recursive.py {version}")
        self.root.geometry("650x500")
        
        try:
            self.root.tk.call('ttk::style', 'theme', 'use', 'vista')
        except tk.TclError:
            pass  # Fallback to default if vista not available

        # Description
        desc_label = tk.Label(root, text=description, font=("Arial", 10), wraplength=600)
        desc_label.pack(pady=10)

        # Input File
        input_frame = tk.Frame(root)
        input_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(input_frame, text="Input File (Zip):").pack(anchor=tk.W)
        
        self.input_entry = tk.Entry(input_frame)
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        tk.Button(input_frame, text="Browse", command=self.browse_input).pack(side=tk.LEFT)

        # Output Folder
        output_frame = tk.Frame(root)
        output_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(output_frame, text="Output Folder:").pack(anchor=tk.W)
        
        self.output_entry = tk.Entry(output_frame)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        tk.Button(output_frame, text="Browse", command=self.browse_output).pack(side=tk.LEFT)

        # Progress Bar
        self.progress = ttk.Progressbar(root, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=10, pady=10)

        # Status Window
        tk.Label(root, text="Status:").pack(fill=tk.X, padx=10)
        self.status_window = scrolledtext.ScrolledText(root, height=12)
        self.status_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Buttons
        btn_frame = tk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        self.unzip_btn = tk.Button(btn_frame, text="Unzip", command=self.start_unzip_thread, height=2, width=20, bg="#dddddd")
        self.unzip_btn.pack()

        # Set default output when input changes
        self.input_entry.bind('<KeyRelease>', self.update_output_path)

    def browse_input(self):
        filename = filedialog.askopenfilename(filetypes=[("Zip files", "*.zip"), ("All files", "*.*")])
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)
            self.update_output_path()

    def browse_output(self):
        dirname = filedialog.askdirectory()
        if dirname:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, dirname)

    def update_output_path(self, event=None):
        input_path = self.input_entry.get()
        if input_path and not self.output_entry.get():
            # Suggest output folder based on input filename
            base_name = os.path.splitext(os.path.basename(input_path))[0]
            parent_dir = os.path.dirname(input_path)
            suggested_output = os.path.join(parent_dir, base_name)
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, suggested_output)

    def log(self, message):
        print(message)
        self.root.after(0, self._log_internal, message)

    def _log_internal(self, message):
        self.status_window.insert(tk.END, message + "\n")
        self.status_window.see(tk.END)

    def start_unzip_thread(self):
        input_file = self.input_entry.get()
        output_folder = self.output_entry.get()

        if not input_file or not os.path.exists(input_file):
            self.log("Error: Input file does not exist.")
            return
        
        if not output_folder:
            self.log("Error: Output folder is required.")
            return

        self.unzip_btn.config(state=tk.DISABLED)
        self.progress.start()
        
        thread = threading.Thread(target=self.run_unzip, args=(input_file, output_folder))
        thread.daemon = True
        thread.start()

    def run_unzip(self, input_file, output_folder):
        try:
            self.log(f"Starting unzip process...")
            self.log(f"Input: {input_file}")
            self.log(f"Output: {output_folder}")
            
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                self.log(f"Created output directory: {output_folder}")

            extract_nested_zip(input_file, output_folder, log_callback=self.log)
            
            self.log("-" * 30)
            self.log(f"Done! Output saved to: {output_folder}")
            
        except Exception as e:
            self.log(f"Error occurred: {str(e)}")
        finally:
            self.root.after(0, self.processing_complete)

    def processing_complete(self):
        self.progress.stop()
        self.unzip_btn.config(state=tk.NORMAL)


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Logic          >>>>>>>>>>>>>>>>>>>>>>>>>>

def extract_nested_zip(zip_source, extract_to, log_callback=None):
    # zip_source can be a file path or a BytesIO object
    
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    try:
        with zipfile.ZipFile(zip_source, 'r') as zip_file:
            zip_file.extractall(extract_to)
            
            # If zip_source is a file path (string), log it. If it's BytesIO, we skip logging the name here to avoid spamming for inner files too much, 
            # unless we want to track every nested file.
            if isinstance(zip_source, str):
                log(f"Extracted: {os.path.basename(zip_source)}")

            for file_name in zip_file.namelist():
                if file_name.endswith('.zip'):
                    nested_extract_to = os.path.join(extract_to, file_name.replace('.zip', ''))
                    os.makedirs(nested_extract_to, exist_ok=True)
                    
                    log(f"Found nested zip: {file_name}")

                    with zip_file.open(file_name) as nested_zip_file:
                        nested_bytes = io.BytesIO(nested_zip_file.read())
                        extract_nested_zip(nested_bytes, nested_extract_to, log_callback)
    except zipfile.BadZipFile:
        log(f"Error: Bad zip file encountered.")
    except Exception as e:
        log(f"Error extracting zip: {e}")

def msg_blurb_square(msg, color):
    border = f"+{'-' * (len(msg) + 2)}+"
    print(f"{color}{border}\n| {msg} |\n{border}{color_reset}")

# <<<<<<<<<<<<<<<<<<<<<<<<<<      Main           >>>>>>>>>>>>>>>>>>>>>>>>>>

def main():
    # If no arguments are provided, launch GUI
    if len(sys.argv) == 1:
        launch_gui()
        return

    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('-I', '--input', help='Input zip file', required=False)
    parser.add_argument('-O', '--output', help='Output folder', required=False)
    parser.add_argument('-u', '--unzip', help='unzip file and sub zips', required=False, action='store_true')

    args = parser.parse_args()

    input_file = args.input if args.input else "sample.zip"
    output_folder = args.output if args.output else input_file.replace('.zip','')

    if args.unzip:
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        file_exists = os.path.exists(input_file)

        if file_exists:
            msg_blurb = (f'Unzipping {input_file} to {output_folder} folder')
            msg_blurb_square(msg_blurb, color_green)    
            extract_nested_zip(input_file, output_folder)
            print(f"\nDone! Output saved to: {output_folder}")
        else:
            msg_blurb = (f'{input_file} does not exist')
            msg_blurb_square(msg_blurb, color_red)      
            exit()
    else:
        # If arguments are provided but not -u, and not empty enough to trigger GUI (caught above), show usage
        usage()
    
    return 0

def launch_gui():
    root = tk.Tk()
    app = UnzipGUI(root)
    root.mainloop()

def usage():
    print(f"Usage: {sys.argv[0]} -u [-I sample.zip] [-O output]")
    print(f"       {sys.argv[0]} (Run without arguments for GUI)")
    print("Example:")
    print(f"    {sys.argv[0]} -u")
    print(f"    {sys.argv[0]} -u -I sample.zip -O sample")

if __name__ == '__main__':
    main()
