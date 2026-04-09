import os
import threading
import sys
from pypdf import PdfReader, PdfWriter
import tkinter as tk
from tkinter import ttk, filedialog

version = "1.0.1"
description = "PDF password remover"

SCRIPT_NAME = os.path.splitext(os.path.basename(sys.argv[0]))[0]


def build_output_path(input_path: str) -> str:
    base, ext = os.path.splitext(input_path)
    base = base.replace(' ', '_')
    if ext.lower() != ".pdf":
        ext = ".pdf"
    return f"{base}_unprotected{ext}"


def unprotect_pdf(input_pdf: str, output_pdf: str, password: str, status_var, progress_bar):
    if not input_pdf:
        status_var.set("No input file selected.")
        progress_bar.stop()
        return

    try:
        status_var.set("Reading PDF...")
        reader = PdfReader(input_pdf)

        if password:
            reader.decrypt(password)

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        status_var.set("Writing unprotected PDF...")
        with open(output_pdf, "wb") as f:
            writer.write(f)

        status_var.set(f"Done.\nInput file: {input_pdf}\nOutput file: {output_pdf}")
        print(f"Input file: {input_pdf}")
        print(f"Output file: {output_pdf}")

    except Exception as e:
        status_var.set(f"Error: {e}")
    finally:
        progress_bar.stop()


def start_unprotect_thread(input_var, password_var, status_var, progress_bar):
    input_pdf = input_var.get().strip()
    password = password_var.get()
    if not input_pdf:
        status_var.set("Please select an input PDF.")
        return

    output_pdf = build_output_path(input_pdf)
    status_var.set("Processing...")
    progress_bar.start()

    t = threading.Thread(
        target=unprotect_pdf,
        args=(input_pdf, output_pdf, password, status_var, progress_bar),
        daemon=True,
    )
    t.start()


def browse_input_file(input_var, status_var):
    filename = filedialog.askopenfilename(
        title="Select PDF",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    if filename:
        input_var.set(filename)
        status_var.set(f"Selected: {os.path.basename(filename)}")


def main():
    root = tk.Tk()

    # Try to use Vista theme if available
    style = ttk.Style()
    try:
        style.theme_use("vista")
    except tk.TclError:
        # Fallback to default theme
        pass

    root.title(f"{SCRIPT_NAME} {version}")

    main_frame = ttk.Frame(root, padding=10)
    main_frame.grid(row=0, column=0, sticky="nsew")

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    # Description label
    desc_label = ttk.Label(main_frame, text=description)
    desc_label.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

    # Input file
    input_var = tk.StringVar()
    ttk.Label(main_frame, text="Input PDF:").grid(row=1, column=0, sticky="e", padx=(0, 5))
    input_entry = ttk.Entry(main_frame, textvariable=input_var, width=50)
    input_entry.grid(row=1, column=1, sticky="we")
    browse_btn = ttk.Button(
        main_frame,
        text="Browse...",
        command=lambda: browse_input_file(input_var, status_var),
    )
    browse_btn.grid(row=1, column=2, sticky="w", padx=(5, 0))

    # Password
    password_var = tk.StringVar(value="")
    ttk.Label(main_frame, text="Password:").grid(row=2, column=0, sticky="e", padx=(0, 5), pady=(5, 0))
    password_entry = ttk.Entry(main_frame, textvariable=password_var, show="*", width=50)
    password_entry.grid(row=2, column=1, sticky="we", pady=(5, 0))
    # Empty spacer in column 2
    ttk.Label(main_frame, text="").grid(row=2, column=2)

    # Progress bar
    progress_bar = ttk.Progressbar(main_frame, mode="indeterminate")
    progress_bar.grid(row=3, column=0, columnspan=3, sticky="we", pady=(10, 5))

    # Status label
    status_var = tk.StringVar(value="Idle.")
    status_label = ttk.Label(main_frame, textvariable=status_var)
    status_label.grid(row=4, column=0, columnspan=3, sticky="w")

    # Unprotect button
    unprotect_btn = ttk.Button(
        main_frame,
        text="Unprotect PDF",
        command=lambda: start_unprotect_thread(input_var, password_var, status_var, progress_bar),
    )
    unprotect_btn.grid(row=5, column=0, columnspan=3, pady=(10, 0))

    main_frame.columnconfigure(1, weight=1)

    root.mainloop()


if __name__ == "__main__":
    main()
