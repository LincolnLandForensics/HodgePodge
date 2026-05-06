import os
import sys
import subprocess
import shutil
import platform
import filetype
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

description = '''
Process print job files into standard file formats.
'''

version = "1.0.1"

# Default folders
DEFAULT_INPUT = 'input'
DEFAULT_OUTPUT = 'output'

# GUI globals
gui_active = False
root = None
entry_input = None
entry_output = None
text_status = None
progress = None
btn_convert = None

# MIME → extension map for normalization
MIME_EXTENSION_MAP = {
    "application/pdf": ".pdf",
    "application/postscript": ".ps",
    "application/vnd.cups-raster": ".ras",
    "application/vnd.apple.raster": ".ur",
    "application/octet-stream": ".bin",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
    "image/tiff": ".tiff",
    "image/x-portable-bitmap": ".pbm",
    "image/x-portable-graymap": ".pgm",
    "image/x-portable-pixmap": ".ppm",
    "image/x-portable-anymap": ".pnm",
    "image/x-sgi-rgb": ".rgb",
    "text/plain": ".txt",
}

RASTER_MIME_TYPES = {
    "image/jpeg",
    "image/jpg",
    "image/png",
    "image/gif",
    "image/tiff",
    "image/x-portable-bitmap",
    "image/x-portable-graymap",
    "image/x-portable-pixmap",
    "image/x-portable-anymap",
    "image/x-sgi-rgb",
    "application/vnd.cups-raster",
    "application/vnd.apple.raster",
}


def extension_for_mime(mime):
    if not mime:
        return ".bin"
    return MIME_EXTENSION_MAP.get(mime, ".bin")


# ---------------------------------------------------------
# LOGGING
# ---------------------------------------------------------
def log(msg):
    """Log to terminal and GUI status window."""
    print(msg)
    if gui_active and text_status is not None:
        def append():
            text_status.insert(tk.END, msg + "\n")
            text_status.see(tk.END)
        root.after(0, append)


# ---------------------------------------------------------
# WINDOWS IMAGEMAGICK VALIDATION
# ---------------------------------------------------------
def ensure_imagemagick_on_windows():
    if platform.system() == "Windows":
        exe = shutil.which("magick")
        if exe is None:
            print("ERROR: ImageMagick is not installed or not in PATH on Windows.")
            sys.exit(1)

        try:
            out = subprocess.check_output(["magick", "-version"], stderr=subprocess.STDOUT)
            text = out.decode().lower()
            if "imagemagick" not in text:
                raise Exception("Not real ImageMagick")
        except Exception:
            print("ERROR: 'magick' was found, but it is NOT ImageMagick.")
            print("Install ImageMagick (64‑bit Q16-HDRI) and enable:")
            print("  - Add application directory to PATH")
            print("  - Install legacy utilities (optional)")
            sys.exit(1)


ensure_imagemagick_on_windows()


# ---------------------------------------------------------
# FILE TYPE DETECTION
# ---------------------------------------------------------
def detect_file_type(path):
    """Detect type using filetype library, fallback to `file` command."""
    kind = filetype.guess(path)
    if kind:
        return kind.mime

    try:
        out = subprocess.check_output(
            ["file", "--mime-type", "-b", path]
        ).decode().strip()
        return out
    except Exception:
        return None


def cupsfilter_available():
    return shutil.which("cupsfilter") is not None


def imagemagick_cmd():
    return "magick" if platform.system() == "Windows" else "convert"


# ---------------------------------------------------------
# JPG CONVERSION
# ---------------------------------------------------------
def convert_to_jpg(input_path, output_path):
    mime = detect_file_type(input_path)
    log(f"Detected MIME type: {mime}")

    if cupsfilter_available() and platform.system() != "Windows":
        log("Using CUPS cupsfilter...")
        try:
            subprocess.check_call(
                ["cupsfilter", "-m", "image/jpeg", input_path, "-o", output_path]
            )
            log("Conversion successful via cupsfilter.")
            return
        except Exception as e:
            log(f"cupsfilter failed, falling back to ImageMagick: {e}")

    magick = imagemagick_cmd()

    if mime in ("application/pdf", "application/postscript"):
        subprocess.check_call([magick, "-density", "300", input_path, output_path])
        log("Conversion successful via ImageMagick.")
        return

    subprocess.check_call([magick, input_path, output_path])
    log("Conversion successful via ImageMagick (raster fallback).")


# ---------------------------------------------------------
# FOLDER PROCESSING
# ---------------------------------------------------------
def process_folder(in_folder, out_folder):
    if not os.path.isdir(in_folder):
        log(f"Input folder does not exist: {in_folder}")
        return

    os.makedirs(out_folder, exist_ok=True)

    files = sorted(os.listdir(in_folder))
    if not files:
        log("No files found in input folder.")
        return

    for name in files:
        in_path = os.path.join(in_folder, name)
        if not os.path.isfile(in_path):
            continue

        log(f"Processing input file: {in_path}")

        base, ext = os.path.splitext(name)

        # CASE 1 — File has extension → copy as-is
        if ext:
            out_path = os.path.join(out_folder, name)
            shutil.copy2(in_path, out_path)
            log(f"Copied to output file: {out_path}")
            continue

        # CASE 2 — No extension → detect MIME and normalize
        mime = detect_file_type(in_path)
        log(f"Detected MIME type: {mime}")

        # PDFs and PS → copy with proper extension
        if mime in ("application/pdf", "application/postscript"):
            out_ext = extension_for_mime(mime)
            out_path = os.path.join(out_folder, base + out_ext)
            shutil.copy2(in_path, out_path)
            log(f"Copied document to output file: {out_path}")
            continue

        # Raster/image/CUPS raster → normalize to JPG
        if mime in RASTER_MIME_TYPES:
            out_path = os.path.join(out_folder, base + ".jpg")
            try:
                convert_to_jpg(in_path, out_path)
                log(f"Converted raster to JPG: {out_path}")
            except Exception as e:
                log(f"Conversion failed for {in_path}: {e}")
            continue

        # Text → copy as .txt
        if mime == "text/plain":
            out_path = os.path.join(out_folder, base + ".txt")
            shutil.copy2(in_path, out_path)
            log(f"Copied text to output file: {out_path}")
            continue

        # Unknown/other → copy with best-guess extension
        out_ext = extension_for_mime(mime)
        out_path = os.path.join(out_folder, base + out_ext)
        shutil.copy2(in_path, out_path)
        log(f"Copied unknown type to output file: {out_path}")

    log(f"All done. Output folder: {out_folder}")


# ---------------------------------------------------------
# THREADING
# ---------------------------------------------------------
def start_processing_thread():
    in_folder = entry_input.get().strip()
    out_folder = entry_output.get().strip()

    if not in_folder:
        in_folder = DEFAULT_INPUT
    if not out_folder:
        out_folder = DEFAULT_OUTPUT

    btn_convert.config(state=tk.DISABLED)
    progress.start(10)

    def worker():
        try:
            process_folder(in_folder, out_folder)
        finally:
            def done():
                progress.stop()
                btn_convert.config(state=tk.NORMAL)
            root.after(0, done)

    threading.Thread(target=worker, daemon=True).start()


# ---------------------------------------------------------
# GUI
# ---------------------------------------------------------
def launch_gui():
    global gui_active, root, entry_input, entry_output, text_status, progress, btn_convert

    gui_active = True
    root = tk.Tk()

    script_name = os.path.basename(sys.argv[0])
    root.title(f"{script_name} {version}")
    root.geometry("700x500")

    style = ttk.Style()
    if 'vista' in style.theme_names():
        style.theme_use('vista')

    lbl_desc = tk.Label(root, text=description.strip(), font=("Arial", 12, "bold"), justify="left")
    lbl_desc.pack(pady=10, anchor="w", padx=20)

    frame_input = tk.Frame(root)
    frame_input.pack(fill=tk.X, padx=20, pady=5)

    tk.Label(frame_input, text="Input Folder:", width=15, anchor="e").pack(side=tk.LEFT)
    entry_input = tk.Entry(frame_input)
    entry_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    entry_input.insert(0, DEFAULT_INPUT)

    def browse_input():
        folder = filedialog.askdirectory()
        if folder:
            entry_input.delete(0, tk.END)
            entry_input.insert(0, folder)

    tk.Button(frame_input, text="Browse", command=browse_input).pack(side=tk.LEFT)

    frame_output = tk.Frame(root)
    frame_output.pack(fill=tk.X, padx=20, pady=5)

    tk.Label(frame_output, text="Output Folder:", width=15, anchor="e").pack(side=tk.LEFT)
    entry_output = tk.Entry(frame_output)
    entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    entry_output.insert(0, DEFAULT_OUTPUT)

    def browse_output():
        folder = filedialog.askdirectory()
        if folder:
            entry_output.delete(0, tk.END)
            entry_output.insert(0, folder)

    tk.Button(frame_output, text="Browse", command=browse_output).pack(side=tk.LEFT)

    btn_convert = tk.Button(
        root,
        text="Convert",
        command=start_processing_thread,
        font=("Arial", 10, "bold"),
        bg="#4CAF50",
        fg="white"
    )
    btn_convert.pack(pady=15)

    progress = ttk.Progressbar(root, mode='indeterminate')
    progress.pack(fill=tk.X, padx=20, pady=5)

    tk.Label(root, text="Status Output:").pack(anchor="w", padx=20)
    text_status = scrolledtext.ScrolledText(root, height=15)
    text_status.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    root.mainloop()


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if __name__ == "__main__":
    launch_gui()
