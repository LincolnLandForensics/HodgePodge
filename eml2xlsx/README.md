## 🕵️ EML to XLSX Parser

A forensic-grade GUI tool for parsing .eml, (.emlx), .mbox, .ost, .pst and .json email files into structured Excel workbooks. Designed for investigators, fraud analysts, and digital forensic workflows.

📦 Features
- ✅ GUI interface built with Tkinter
- ✅ Recursively extracts ZIP archives
- ✅ Parses .eml, .mbox, .ost, .pst and .json files
- ✅ Optional deduplication by SHA256 hash
- ✅ Exports to Excel with two sheets: Eml and Contacts

🖼️ Screenshots
![sample output](images/eml2xlsx.png)


🚀 Installation
```
git clone https://github.com/LincolnLandForensics/HodgePodge.git
```
cd eml2xlsx


🧪 modules installation
```
pip install -r requirements_eml2xlsx.txt
```


🧪 Usage
```
python eml2xlsx.py
```

- Select your input folder containing .eml, .mbox, .json, and/or .zip files
- Choose an output filename (e.g., email.xlsx)
- (Optional) Check the DeDuplicate box to ed-duplicate SHA256 hashes (Skip if it's a file with more than one email in it)
- Click <Start Parsing> and monitor progress in the log window

📁 Output Structure
The Excel file contains:
Eml Sheet
- Time, From, To, Subject, Body, Attachments, Labels, Tags, Source, SHA256, and more
Contacts Sheet
- Extracted names, emails, and linkage to original files


📜 License
This project is licensed under the MIT License.
