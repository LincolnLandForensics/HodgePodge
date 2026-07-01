import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import Image
import pillow_heif

# CRITICAL FIX: Globally disable all metadata, EXIF, and XMP parsing 
# This forces libheif to completely skip reading the broken metadata blocks.
pillow_heif.options.HEIF_THUMBNAILS = False
pillow_heif.options.SAVE_EXIF = False
pillow_heif.options.DECODE_THREADS = 4

# Register HEIF opener with Pillow for native decoding fallback
pillow_heif.register_heif_opener()

class HEICConverterGUI:
    def __init__(self, root):
        self.root = root
        
        # Script Metadata variables
        self.script_name = "HEIC to JPG Converter"
        self.version = "v2.2"
        self.description = "Batch convert Apple iPhone .heic photos to standard .jpg images effortlessly."
        
        self.root.title(f"{self.script_name} {self.version}")
        self.root.geometry("620x520")
        self.root.minsize(550, 450)
        
        # Configure a clean look using the native system theme
        self.style = ttk.Style()
        if "vista" in self.style.theme_names():
            self.style.theme_use("vista")
        else:
            self.style.theme_use("clam")

        # String Variables for dynamic tracking
        self.input_dir_var = tk.StringVar(value=r"C:\Forensics\scripts\python\images")
        self.output_dir_var = tk.StringVar()
        
        self._update_output_path()
        self.input_dir_var.trace_add("write", lambda *args: self._update_output_path())
        
        self._build_ui()

    def _update_output_path(self):
        """Automatically updates the output path when the input path changes."""
        inp = self.input_dir_var.get()
        if inp:
            self.output_dir_var.set(os.path.join(inp, "images_converted"))
        else:
            self.output_dir_var.set("")

    def _build_ui(self):
        # Main Container with padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header Section
        # title_label = ttk.Label(main_frame, text=f"{self.script_name} {self.version}", font=("Segoe UI", 14, "bold"))
        # title_label.pack(anchor=tk.W, pady=(0, 2))
        
        desc_label = ttk.Label(main_frame, text=self.description, font=("Segoe UI", 9, "italic"), wraplength=580)
        desc_label.pack(anchor=tk.W, pady=(0, 15))

        # --- Folder Selection Frame ---
        folder_frame = ttk.LabelFrame(main_frame, text=" Directories ", padding="10")
        folder_frame.pack(fill=tk.X, pady=(0, 15))

        # Input Row
        ttk.Label(folder_frame, text="Input Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(folder_frame, textvariable=self.input_dir_var, width=50).grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(folder_frame, text="Browse...", command=self._browse_input).grid(row=0, column=2, padx=2, pady=5)

        # Output Row
        ttk.Label(folder_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(folder_frame, textvariable=self.output_dir_var, width=50).grid(row=1, column=1, padx=5, pady=5, sticky=tk.EW)
        ttk.Button(folder_frame, text="Browse...", command=self._browse_output).grid(row=1, column=2, padx=2, pady=5)
        
        folder_frame.columnconfigure(1, weight=1)

        # --- Progress Bar & Action Button ---
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))


        # self.convert_btn = ttk.Button(main_frame, text='Convert iPhone .heic photos to .jpg', command=self._start_conversion_thread)

        self.convert_btn = ttk.Button(main_frame, text='Convert photos', command=self._start_conversion_thread)
        self.convert_btn.pack(fill=tk.X, ipady=5, pady=(0, 15))

        # --- Embedded Status Terminal Window ---
        terminal_frame = ttk.LabelFrame(main_frame, text=" Terminal Console Log ", padding="5")
        terminal_frame.pack(fill=tk.BOTH, expand=True)

        self.terminal_log = tk.Text(terminal_frame, wrap=tk.WORD, height=12, state=tk.DISABLED, bg="#1e1e1e", fg="#f1f1f1", font=("Consolas", 9))
        self.terminal_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(terminal_frame, orient=tk.VERTICAL, command=self.terminal_log.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.terminal_log.configure(yscrollcommand=scrollbar.set)

    def _browse_input(self):
        selected = filedialog.askdirectory(initialdir=self.input_dir_var.get())
        if selected:
            self.input_dir_var.set(os.path.normpath(selected))

    def _browse_output(self):
        selected = filedialog.askdirectory(initialdir=self.output_dir_var.get())
        if selected:
            self.output_dir_var.set(os.path.normpath(selected))

    def _log_to_terminal(self, text):
        """Appends logs to the UI console and prints natively to the system terminal window."""
        try:
            print(text)
        except UnicodeEncodeError:
            # Fallback for Windows consoles with non-UTF-8 encodings
            try:
                import sys
                encoding = sys.stdout.encoding or 'utf-8'
                print(text.encode(encoding, errors='replace').decode(encoding))
            except Exception:
                print(text.encode('ascii', errors='replace').decode('ascii'))
        self.terminal_log.config(state=tk.NORMAL)
        self.terminal_log.insert(tk.END, text + "\n")
        self.terminal_log.see(tk.END)
        self.terminal_log.config(state=tk.DISABLED)

    def _start_conversion_thread(self):
        self.convert_btn.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        worker = threading.Thread(target=self._process_conversion, daemon=True)
        worker.start()

    def _process_conversion(self):
        input_folder = self.input_dir_var.get()
        output_folder = self.output_dir_var.get()

        if not input_folder or not os.path.exists(input_folder):
            self._log_to_terminal(f"CRITICAL ERROR: Input directory '{input_folder}' does not exist.")
            self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))
            return

        os.makedirs(output_folder, exist_ok=True)
        
        all_files = os.listdir(input_folder)
        heic_files = [f for f in all_files if f.lower().endswith(".heic")]
        total_files = len(heic_files)

        self._log_to_terminal(f"--- Initialization Context ---")
        self._log_to_terminal(f"Scanning target folder: {input_folder}")
        self._log_to_terminal(f"Discovered matching targets: {total_files} .heic file(s)\n")

        if total_files == 0:
            self._log_to_terminal("Execution Halted: No targets detected matching extension parameters.")
            self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))
            return

        self.root.after(0, lambda: self.progress_bar.config(maximum=total_files))

        converted_count = 0
        for index, filename in enumerate(heic_files):
            heic_file_path = os.path.join(input_folder, filename)
            jpg_filename = os.path.splitext(filename)[0] + ".jpg"
            jpg_file_path = os.path.join(output_folder, jpg_filename)
            
            self._log_to_terminal(f"[Processing Input]: {filename}")
            
            try:
                # Open image using raw bytes loading sequence
                heif_file = pillow_heif.open_heif(heic_file_path, convert_hdr_to_8bit=True)
                
                # Double down by actively clearing container metadata details
                heif_file.info.clear()
                
                img = Image.frombytes(
                    heif_file.mode,
                    heif_file.size,
                    heif_file.data,
                    "raw",
                    heif_file.mode,
                    heif_file.stride,
                )
                img.convert("RGB").save(jpg_file_path, "JPEG", quality=90)
                self._log_to_terminal(f"   ↳ [Success Output Created]: {jpg_filename}")
                converted_count += 1
            except Exception as e:
                self._log_to_terminal(f"   ↳ [Error]: Structural container failure on file: {str(e)}")

            self.root.after(0, lambda v=index+1: self.progress_bar.config(value=v))

        self._log_to_terminal(f"\n--- Batch Job Complete ---")
        self._log_to_terminal(f"Successfully finalized {converted_count} image conversions.")
        self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))

if __name__ == "__main__":
    root = tk.Tk()
    app = HEICConverterGUI(root)
    root.mainloop()