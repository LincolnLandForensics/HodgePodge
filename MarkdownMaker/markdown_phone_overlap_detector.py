import os
import re
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext

version = "v1.0"
description = (
    "Point this at a folder full of markdown contacts. It will create new markdown files "
    "with only the lines that have phone numbers that are also in at least one of the "
    "other markdown files."
)

input_folder = 'Contacts'
output_folder = 'Overlapping_Contacts'

# Updated regex: matches [[##########]] and extracts only the digits
regex_phone = re.compile(r"\[\[(\d{7,17})\]\]")

root = None
entry_input = None
entry_output = None
text_status = None
progress = None
btn_start = None
gui_active = False


def log_terminal(msg: str):
    print(msg)


def log_status(msg: str):
    if gui_active and text_status is not None:
        def _append():
            text_status.insert(tk.END, msg + "\n")
            text_status.see(tk.END)
        root.after(0, _append)
    log_terminal(msg)


def find_md_files(folder: str):
    if not os.path.isdir(folder):
        return []
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith(".md") and os.path.isfile(os.path.join(folder, f))
    ]


def build_phone_index(md_files):
    phone_index = {}      # phone -> set(files)
    file_lines = {}       # file -> list of (phone, line)

    for path in md_files:
        file_lines[path] = []
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    stripped = line.strip()
                    m = regex_phone.search(stripped)
                    if m:
                        phone = m.group(1)
                        file_lines[path].append((phone, line.rstrip("\n")))
                        phone_index.setdefault(phone, set()).add(path)
        except Exception as e:
            log_status(f"Error reading {path}: {e}")

    return phone_index, file_lines


def process_contacts(input_dir: str, output_dir: str):
    md_files = find_md_files(input_dir)

    if md_files:
        log_status(f"Input folder: {input_dir}")
    else:
        log_status(f"No .md files found in input folder: {input_dir}")

    if not os.path.isdir(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    phone_index, file_lines = build_phone_index(md_files)

    for path in md_files:
        overlaps = set()

        for phone, line in file_lines.get(path, []):
            files_with_phone = phone_index.get(phone, set())

            # Must appear in at least one other file
            if len(files_with_phone) > 1:
                if any(other != path for other in files_with_phone):
                    overlaps.add(line)

        base_name = os.path.basename(path)
        new_name = f"overlaps_{base_name}"
        out_path = os.path.join(output_dir, new_name)

        if overlaps:
            sorted_lines = sorted(overlaps)

            try:
                with open(out_path, "w", encoding="utf-8") as out_f:

                    # Build header based on original filename
                    xlsx_name = os.path.splitext(base_name)[0] + ".xlsx"
                    out_f.write(f"Contact Summary from {input_dir}/{xlsx_name}\n")
                    out_f.write("fullname\tphone\tuser\tfulladdress\tcase\n")

                    # Write overlapping lines
                    for line in sorted_lines:
                        out_f.write(line + "\n")

                    # Footer
                    out_f.write("End of overlapping contacts\n")

                log_status(f"Wrote overlapping contacts: {out_path}")

            except Exception as e:
                log_status(f"Error writing {out_path}: {e}")

        else:
            log_status(f"No overlapping phone numbers for: {path}")

    log_status(f"Output folder: {output_dir}")
    log_terminal(f"Output folder: {output_dir}")


def start_processing():
    input_dir = entry_input.get().strip() or input_folder
    output_dir = entry_output.get().strip() or output_folder

    btn_start.config(state=tk.DISABLED)
    progress.start(10)
    log_status("Starting processing...")

    def worker():
        try:
            process_contacts(input_dir, output_dir)
        finally:
            def done():
                progress.stop()
                btn_start.config(state=tk.NORMAL)
                log_status("Processing complete.")
            root.after(0, done)

    threading.Thread(target=worker, daemon=True).start()


def browse_input():
    folder = filedialog.askdirectory()
    if folder:
        entry_input.delete(0, tk.END)
        entry_input.insert(0, folder)


def browse_output():
    folder = filedialog.askdirectory()
    if folder:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, folder)


def launch_gui():
    global gui_active, text_status, progress, btn_start
    global entry_input, entry_output, root

    gui_active = True
    root = tk.Tk()

    script_name = os.path.basename(sys.argv[0])
    root.title(f"{script_name} {version}")
    root.geometry("700x500")

    style = ttk.Style()
    if 'vista' in style.theme_names():
        style.theme_use('vista')

    lbl_desc = tk.Label(root, text=description, font=("Arial", 11, "bold"),
                        wraplength=650, justify="left")
    lbl_desc.pack(pady=10, padx=10)

    # Input Folder
    frame_input = tk.Frame(root)
    frame_input.pack(fill=tk.X, padx=20, pady=5)
    tk.Label(frame_input, text="Input Folder:", width=15, anchor="e").pack(side=tk.LEFT)
    entry_input = tk.Entry(frame_input)
    entry_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    entry_input.insert(0, input_folder)
    tk.Button(frame_input, text="Browse", command=browse_input).pack(side=tk.LEFT)

    # Output Folder
    frame_output = tk.Frame(root)
    frame_output.pack(fill=tk.X, padx=20, pady=5)
    tk.Label(frame_output, text="Output Folder:", width=15, anchor="e").pack(side=tk.LEFT)
    entry_output = tk.Entry(frame_output)
    entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
    entry_output.insert(0, output_folder)
    tk.Button(frame_output, text="Browse", command=browse_output).pack(side=tk.LEFT)

    # Start Button
    btn_start = tk.Button(root, text="Start", command=start_processing,
                          font=("Arial", 10, "bold"), bg="#4CAF50", fg="white")
    btn_start.pack(pady=15)

    # Progress Bar
    progress = ttk.Progressbar(root, mode='indeterminate')
    progress.pack(fill=tk.X, padx=20, pady=5)

    # Status Window
    tk.Label(root, text="Status Output:").pack(anchor="w", padx=20)
    text_status = scrolledtext.ScrolledText(root, height=15)
    text_status.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

    root.mainloop()


if __name__ == "__main__":
    launch_gui()
