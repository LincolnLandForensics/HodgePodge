#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
Pinginator.py - Network Forensic Tool
Rewritten for Python 3 and PEP 8 compliance.
"""

import os
import re
import csv
import sys
import time
import socket
import random
import argparse
import subprocess
import urllib.request
import urllib.error
import urllib.parse
import struct

import ipaddress
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
from tkinter import ttk
import threading

import ssl
from html.parser import HTMLParser

# Pre-Sets
AUTHOR = 'LincolnLandForensics'
DESCRIPTION = "A rapid network‑sweep utility that discovers live host in subnet(s) and/or performs targeted port scans to fingerprint and identify each device."
TECH = 'kali'        # change this to your name
VERSION = '2.0.4'
INPUT_FILENAME = 'nodes.txt'

# Pre-Sets removed globals


# Regex Compilation
REGEX_HOST = re.compile(
    r'(?i)\b((?:(?!-)[a-zA-Z0-9-]{1,63}(?<!-)\.)+(?!exe|php|dll|doc'
    r'|docx|txt|rtf|odt|xls|xlsx|ppt|pptx|bin|pcap|ioc|pdf|mdb|asp|html|xml|jpg|gif$|png'
    r'|lnk|log|vbs|lco|bat|shell|quit|pdb|vbp|bdoda|bsspx|save|cpl|wav|tmp|close|ico|ini'
    r'|sleep|run|dat$|scr|jar|jxr|apt|w32|css|js|xpi|class|apk|rar|zip|hlp|cpp|crl'
    r'|cfg|cer|plg|lxdns|cgi|xn$)(?:xn--[a-zA-Z0-9]{2,22}|[a-zA-Z]{2,13}))(?:\s|$)'
)

REGEX_IPV4 = re.compile(
    r'(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}'
    r'(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)'
)
REGEX_IPV6 = re.compile(
    r'(S*([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}S*|S*('
    r'[0-9a-fA-F]{1,4}:){1,7}:S*|S*([0-9a-fA-F]{1,4}:)'
    r'{1,6}:[0-9a-fA-F]{1,4}S*|S*([0-9a-fA-F]{1,4}:)'
    r'{1,5}(:[0-9a-fA-F]{1,4}){1,2}S*|S*([0-9a-fA-F]'
    r'{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}S*|S*('
    r'[0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}S*'
    r'|S*([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4})'
    r'{1,5}S*|S*[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4})'
    r'{1,6})S*|S*:((:[0-9a-fA-F]{1,4}){1,7}|:)S*|::(ffff'
    r'(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}'
    r'[0-9]){0,1}[0-9]).){3,3}(25[0-5]|(2[0-4]|1{0,1}['
    r'0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25['
    r'0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]).){3,3}(25['
    r'0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))'
)

def cls():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')

def main():
    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument('-I', '--input', help='Input file', required=False, default='nodes.txt')
    parser.add_argument('-o', '--output', help='Output file', required=False, default='output.csv')
    parser.add_argument('-N', '--nmap', help='nmap scan', required=False, action='store_true')
    parser.add_argument('-S', '--subnet', help='subnet scan', required=False, action='store_true')
    
    args = parser.parse_args()
    
    # If no arguments provided, launch GUI
    if len(sys.argv) == 1:
        launch_gui()
        return

    cls()

    # Prepare input lines
    global INPUT_FILENAME
    try:
        with open(args.input, 'r', encoding='utf-8', errors='ignore') as f:
            lines = [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"Error: Input file '{args.input}' not found.")
        return
    # Prepare Input
    INPUT_FILENAME = args.input

    # Prepare Output
    output_filename = args.output
    if not output_filename.endswith('.csv'):
        output_filename += '.csv'
    
    # Initialize CSV
    outfile = open(output_filename, 'w', newline='', encoding='utf-8')
    writer = csv.writer(outfile)

    # Execution Routing
    if args.subnet:
        subnet_scan(writer=writer)

    if args.nmap:
        nmap_scan(lines, writer) 

    # if args.hostname:
        # hostname(lines, writer)  

    # Close file
    try:
        outfile.close() 
        print(f"Done. Saved to {output_filename}")
    except PermissionError:
        print(f"Error: Permission denied writing to {output_filename}. Is it open?")
    except Exception as e:
        print(f"Error saving file: {e}")

    # Set ownership on Linux
    if sys.platform.startswith('linux'):
        try:
            os.system(f"chown {TECH.lower()}:{TECH.lower()} *.csv")
        except Exception:
            pass


def nmap_scan(lines, writer, logger=None, progress_callback=None):
    """Comprehensive Nmap Scan."""
    writer.writerow([
        'IP','hostname','notes','OS','MAC','Manufacturer','DeviceType'
        ,'Http','Https','Smb','SMB','Rdp','Ftp','Ssh','Telnet','Smtp'
        ,'Vnc','Dns','Rshell','Sql','DB2','Mysql','Oracle','JetDirect'
        ,'Other','os_url','title_url','page_status'
    ])
    
    current_hostname = socket.gethostname()
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)

    # print(f'    Reading {writer}\n')
    log(f'    Running NMap... This can take a while (go get some coffee)\n')

    # nmap_ports = "80,443,445,139" # Common
    nmap_ports = "80,443,139,445,3389,21,22,23,25,5900,53,8081" # MEDIUM set
    # nmap_ports = "80,443,139,445,3389,21,22,23,25,5900,53,514,1433,523,3306,1521,8081,9100" # big set
    log(f'    Checking {nmap_ports}\n')

    total_lines = len(lines)
    for idx, line in enumerate(lines, 1):
        if progress_callback:
             progress_callback(int((idx / total_lines) * 100))

        ip, hostname, notes, OS, MAC, Manufacturer, DeviceType = [''] * 7
        Http, Https, Smb, SMB, Rdp, Ftp, Ssh, Telnet, Smtp = [''] * 9
        Vnc, Dns, Rshell, Sql, DB2, Mysql, Oracle, JetDirect, Other = [''] * 9
        os_url, title_url, page_status = [''] * 3

        # Check if hostname
        if re.search(REGEX_HOST, line):
            hostname = line

        if re.match(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', line):
            ip = line

            try:
                # socket.gethostbyaddr(ip) returns (hostname, aliaslist, ipaddrlist)
                dns = socket.gethostbyaddr(ip)
                hostname = dns[0]
                if current_hostname in hostname:
                    notes = "scanner" 

            except socket.error as e:
                msg = str(e)
                notes = msg
                if "host not found" in notes:
                    notes= ""
            
            if notes:
                 log(f"    Note: {notes}")

            nmap_args = f'nmap -sS --open -p {nmap_ports} {ip}'
            
            try:
                proc = subprocess.Popen(nmap_args, shell=True, stdout=subprocess.PIPE)
                stdout, _ = proc.communicate()
                output = stdout.decode('utf-8', errors='ignore')
                
                for output_line in output.splitlines():
                    if "80/tcp" in output_line: 
                        Http = '80'
                        os_url = f"http://{ip}:80"
                    elif "443/tcp" in output_line: 
                        Https = '443'
                        if ip.endswith('254'):
                            DeviceType = "Router"
                    elif "1433/tcp" in output_line: Sql = '1433'
                    elif "523/tcp" in output_line: DB2 = '523'
                    elif "3306/tcp" in output_line: Mysql = '3306'
                    elif "1521/tcp" in output_line: Oracle = '1521'
                    elif "21/tcp" in output_line: Ftp = '21'
                    elif "22/tcp" in output_line: Ssh = '22'
                    elif "23/tcp" in output_line: Telnet = '23'
                    elif "25/tcp" in output_line: Smtp = '25'
                    elif "139/tcp" in output_line: 
                        Smb = '139'
                        DeviceType = "Windows"  
                    elif "445/tcp" in output_line: 
                        SMB = '445'
                        DeviceType = "Windows"
                    elif "53/tcp" in output_line: Dns = '53'
                    elif "3389/tcp" in output_line: Rdp = '3389'
                    elif "514/tcp" in output_line: Rshell = '514'
                    elif "5900/tcp" in output_line: Vnc = '5900'
                    elif "8081/tcp" in output_line: 
                        Other = f"{Other} 8081".strip()
                    elif "9100/tcp" in output_line: 
                        JetDirect = '9100'
                    elif "/tcp" in output_line:
                        temp = output_line.replace('/tcp', '').replace(' open', '').strip()
                        # regex to clean valid port?
                        other = f"{other} {temp}".strip()
                    elif "MAC Address: " in output_line:
                        MAC = output_line.replace("MAC Address: ", "").strip().replace(' (Unknown)', '')
                        if ' (' in MAC:
                            Manufacturer = MAC.split(' (')[1].replace(')', '')
                            MAC = MAC.split(' (')[0]

                    if DeviceType == "" and hostname.startswith(
                        ("amazon-", "blink-", "esp_", "gemodule", "google-home")
                        ):
                        DeviceType = "IoT"

                if os_url != "":
                    content, referer, os_url, title_url, page_status = request(os_url)
                

            except Exception as e:
                log(f"Nmap error: {e}")
            
            log(f'{ip}\t ({hostname}) [{DeviceType}] {Http} {Https} {Smb} {SMB}')

            writer.writerow([
                ip,hostname,notes,OS,MAC,Manufacturer,DeviceType,Http,Https,Smb,SMB,Rdp,Ftp,Ssh,Telnet,Smtp,Vnc,Dns,Rshell,Sql,DB2,Mysql,Oracle,JetDirect,Other,os_url,title_url,page_status])

def get_subnets():
    """Return a list of (ip, mask, cidr) tuples found on Windows or Linux."""
    subnets = []
    
    # Windows
    if sys.platform.startswith('win'):
        try:
            output = subprocess.check_output(
                ["ipconfig"],
                text=True,
                encoding="utf-8",
                errors="ignore"
            )
            ips = re.findall(r"IPv4 Address[^\d]+(\d+\.\d+\.\d+\.\d+)", output)
            masks = re.findall(r"Subnet Mask[^\d]+(\d+\.\d+\.\d+\.\d+)", output)

            for ip, mask in zip(ips, masks):
                if ip.startswith("127."):
                    continue
                cidr = sum(bin(int(o)).count("1") for o in mask.split("."))
                subnets.append((ip, mask, cidr))
        except Exception as e:
            print(f"Error running ipconfig: {e}")

    # Linux
    else:
        try:
            # Use ip -o -4 addr show for easier parsing
            output = subprocess.check_output(
                ["ip", "-o", "-4", "addr", "show"],
                text=True,
                encoding="utf-8",
                errors="ignore"
            )
            # Example line: 2: eth0    inet 192.168.1.10/24 brd 192.168.1.255 scope global eth0
            for line in output.splitlines():
                parts = line.split()
                if len(parts) >= 4:
                    ip_cidr = parts[3]
                    if '/' in ip_cidr:
                        ip, cidr = ip_cidr.split('/')
                        if ip.startswith("127."):
                            continue
                        # Calculate netmask from CIDR
                        mask = socket.inet_ntoa(struct.pack('!I', (1 << 32) - (1 << (32 - int(cidr)))))
                        subnets.append((ip, mask, int(cidr)))
        except Exception as e:
            print(f"Error running ip addr: {e}")

    return subnets

# Simple HTML title extractor using only stdlib
class TitleParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_title = False
        self.title = ""

    def handle_starttag(self, tag, attrs):
        if tag.lower() == "title":
            self.in_title = True

    def handle_endtag(self, tag):
        if tag.lower() == "title":
            self.in_title = False

    def handle_data(self, data):
        if self.in_title:
            self.title += data.strip()


def request(os_url):
    content = ""
    referer = ""
    title_url = ""
    page_status = ""

    # Ignore SSL certificate errors
    ssl_ctx = ssl._create_unverified_context()

    try:
        req = urllib.request.Request(os_url)
        req.add_header("User-Agent", "Mozilla/5.0")
        req.add_header("Referer", referer)

        with urllib.request.urlopen(req, context=ssl_ctx, timeout=5) as response:
            page_status = response.getcode()
            raw = response.read()

            try:
                content = raw.decode("utf-8", errors="replace")
            except:
                content = str(raw)

            # Parse title manually
            parser = TitleParser()
            parser.feed(content)
            title_url = parser.title or ""

    except urllib.error.HTTPError as e:
        page_status = e.code
    except urllib.error.URLError:
        page_status = "Fail"
    except Exception:
        page_status = "Fail"

    # Normalize status text
    ps = str(page_status)
    if ps.startswith("2"):
        page_status = f"Success - {ps}"
    elif ps.startswith("3"):
        page_status = f"Redirect - {ps}"
    elif ps.startswith("4") or ps.startswith("5"):
        page_status = f"Fail - {ps}"
    elif ps.startswith("1"):
        page_status = f"Info - {ps}"
    else:
        page_status = "Fail"

    return content, referer, os_url, title_url, page_status

def parse_selection(selection, max_index):
    """Parse input like '1,3-5' into a list of integers."""
    chosen = set()
    parts = selection.split(",")

    for part in parts:
        part = part.strip()
        if "-" in part:
            start, end = part.split("-")
            try:
                start = int(start)
                end = int(end)
                for i in range(start, end + 1):
                    if 1 <= i <= max_index:
                        chosen.add(i)
            except ValueError:
                pass
        else:
            try:
                num = int(part)
                if 1 <= num <= max_index:
                    chosen.add(num)
            except ValueError:
                pass

    return sorted(chosen)


    return sorted(chosen)


def subnet_scan(writer=None, logger=None, progress_callback=None, gui_selection_callback=None, chosen_indices=None):
    def log(msg):
        if logger:
            logger(msg)
        else:
            print(msg)

    subnets = get_subnets()

    if not subnets:
        log("No IPv4 subnets detected.")
        return

    if chosen_indices is not None:
         # Use pre-provided indices
         pass
    # If only one subnet, use it automatically
    elif len(subnets) == 1:
        ip, mask, cidr = subnets[0]
        log(f"[+] One subnet detected: {ip}/{cidr}")
        chosen_indices = [1]
    else:
        if gui_selection_callback:
            # interactive callback for GUI (Note: should be called from main thread)
             chosen_indices = gui_selection_callback(subnets)
        else:
            log("\nMultiple network ranges detected:")
            for idx, (ip, mask, cidr) in enumerate(subnets, start=1):
                log(f"  {idx}. {ip}/{cidr}   (Mask: {mask})")

            # Prompt for multi-selection
            while True:
                selection = input(
                    "\nSelect which subnet(s) to scan (e.g., 1,3-4): "
                ).strip()

                chosen_indices = parse_selection(selection, len(subnets))
                if chosen_indices:
                    break

                log("Invalid selection. Try again.")
    
    if not chosen_indices:
         log("No subnets selected.")
         return

    # Expand chosen subnets
    chosen_subnets = [subnets[i - 1] for i in chosen_indices]

    log("\n[+] Subnets selected for scanning:")
    for ip, mask, cidr in chosen_subnets:
        log(f"    - {ip}/{cidr}")

    # Scan each subnet
    live_hosts = []

    total_subnets = len(chosen_subnets)
    for idx, (ip, mask, cidr) in enumerate(chosen_subnets, 1):
        subnet_cidr = f"{ip}/{cidr}"
        log(f"\n[+] Scanning subnet: {subnet_cidr}")

        try:
            nmap_output = subprocess.check_output(
                ["nmap", "-sn", subnet_cidr],
                text=True
            )
        except Exception as e:
            log(f"Error running nmap on {subnet_cidr}: {e}")
            continue

        for line in nmap_output.splitlines():
            if "Nmap scan report for" in line:
                ip_match = re.search(r'(\d+\.\d+\.\d+\.\d+)', line)
                if ip_match:
                    live_hosts.append(ip_match.group(1))
        
        if progress_callback:
            progress_callback(int((idx / total_subnets) * 100))

    # Write results
    global INPUT_FILENAME
    if writer: 
       # If writer provided (CSV), this functions differently? 
       # Original logic wrote to nodes.txt. 
       # If called from GUI with writer, maybe we want to output to CSV?
       # The user request said "The PingSubnet toggle selects the subnet_scan() option".
       # But typically subnet scan populates nodes.txt for the NEXT scan.
       # Existing code writes to nodes.txt. I will keep that behavior but maybe also log.
       pass

    with open(INPUT_FILENAME, "w") as f:
        for host in live_hosts:
            f.write(host + "\n")

    log(f"\n[+] Found {len(live_hosts)} live hosts across all selected subnets")
    log(f"[+] Saved to {INPUT_FILENAME}")


def usage():
    File = sys.argv[0].split('\\')[-1]
    print("\ndescription: " + DESCRIPTION)
    print(File +" %s by %s" % (VERSION, AUTHOR))
    print("\nExample:")
    print("\tpython " + File +" -H -I nodes.txt -o out_hostname.csv")
    print("\tpython " + File +" -N -I nodes.txt -o out_portscan.csv")


class GuiApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Pinginator {VERSION}")
        
        # Variables
        self.input_file = tk.StringVar(value="nodes.txt")
        self.output_file = tk.StringVar(value="output.csv")
        self.ping_subnet_var = tk.BooleanVar()
        self.port_scan_var = tk.BooleanVar()

        # Layout
        frame_top = tk.Frame(root, padx=10, pady=10)
        frame_top.pack(fill=tk.X)
        
        # Description Label
        tk.Label(frame_top, text=DESCRIPTION, wraplength=550, justify="left", font=("Helvetica", 10, "italic")).grid(row=0, column=0, columnspan=3, pady=(0, 15), sticky="w")

        # Input Checkbox & Entry
        tk.Label(frame_top, text="Input File:").grid(row=1, column=0, sticky="w")
        tk.Entry(frame_top, textvariable=self.input_file, width=40).grid(row=1, column=1, padx=5)
        tk.Button(frame_top, text="Browse", command=self.browse_input).grid(row=1, column=2)

        # Output Checkbox & Entry
        tk.Label(frame_top, text="Output File:").grid(row=2, column=0, sticky="w")
        tk.Entry(frame_top, textvariable=self.output_file, width=40).grid(row=2, column=1, padx=5)
        tk.Button(frame_top, text="Browse", command=self.browse_output).grid(row=2, column=2)

        # Toggles
        tk.Checkbutton(frame_top, text="PingSubnet (Scan network for live hosts)", variable=self.ping_subnet_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=5)
        tk.Checkbutton(frame_top, text="PortScan (Nmap scan input file)", variable=self.port_scan_var).grid(row=4, column=0, columnspan=2, sticky="w")

        # Scan Button
        tk.Button(root, text="Start Scan", command=self.start_scan_thread, bg="#cccccc", width=20).pack(pady=10)

        # Progress Bar
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(pady=5)
        
        # Output Window
        self.log_area = scrolledtext.ScrolledText(root, width=80, height=20)
        self.log_area.pack(padx=10, pady=10)

    def browse_input(self):
        filename = filedialog.askopenfilename(initialdir=os.getcwd(), title="Select Input File")
        if filename:
            self.input_file.set(filename)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(initialdir=os.getcwd(), title="Select Output File",
                                                  defaultextension=".csv",
                                                  filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if filename:
            self.output_file.set(filename)
    
    def log(self, msg):
        self.log_area.insert(tk.END, str(msg) + "\n")
        self.log_area.see(tk.END)

    def update_progress(self, value):
        self.progress['value'] = value

    def start_scan_thread(self):
        self.progress['value'] = 0
        self.log_area.delete(1.0, tk.END)

        chosen_indices = None
        if self.ping_subnet_var.get():
            subnets = get_subnets()
            if len(subnets) > 1:
                # Get selection on main thread
                msg = "Multiple subnets found:\n"
                for idx, (ip, mask, cidr) in enumerate(subnets, 1):
                    msg += f"{idx}. {ip}/{cidr} (Mask: {mask})\n"
                msg += "\nEnter indices (e.g. 1,2):"
                
                res = simpledialog.askstring("Subnet Selection", msg, parent=self.root)
                if res:
                    chosen_indices = parse_selection(res, len(subnets))
                else:
                    self.log("Scan cancelled: No subnets selected.")
                    return
            elif not subnets:
                self.log("Error: No subnets detected.")
                return

        threading.Thread(target=self.run_scan, args=(chosen_indices,), daemon=True).start()

    def run_scan(self, chosen_subnet_indices=None):
        global INPUT_FILENAME
        input_path = self.input_file.get()
        INPUT_FILENAME = input_path
        output_path = self.output_file.get()

        # PingSubnet
        if self.ping_subnet_var.get():
            self.log("--- Starting Subnet Scan ---")
            
            subnet_scan(logger=self.log, progress_callback=self.update_progress, chosen_indices=chosen_subnet_indices)
            self.log("--- Subnet Scan Complete ---")
            self.log(f"Nodes saved to {INPUT_FILENAME} (used as input for PortScan if selected)")
            # If we just ran subnet scan, logic usually implies we might want to use that as input
            # If input file was set to 'nodes.txt' (default), then PortScan will naturally pick it up.
            
        # PortScan
        if self.port_scan_var.get():
            self.log("--- Starting Port Scan ---")
            try:
                # Read input
                lines = []
                with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = [line.strip() for line in f if line.strip()]

                # Setup Output
                outfile = open(output_path, 'w', newline='', encoding='utf-8')
                writer = csv.writer(outfile)
                
                nmap_scan(lines, writer, logger=self.log, progress_callback=self.update_progress)
                
                outfile.close()
                self.log(f"--- Port Scan Complete. Saved to {output_path} ---")
            except Exception as e:
                self.log(f"Error during PortScan: {e}")
                
        self.log("Done.")
        self.progress['value'] = 100


def launch_gui():
    root = tk.Tk()
    app = GuiApp(root)
    root.mainloop()



if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print(f"\nAborted by user.")
        sys.exit(0)
