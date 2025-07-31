
## eml2xlsx.py 
Parses .eml or .mbox email files from a folder, extracts metadata, and exports to Excel.

Doesn't currently support zip files like RLEAP.

Installation:
```
python pip install -r requirements_eml2xlsx.txt
```

help menu
```
python eml2xlsx.py
```

Examples:

eml2xlsx.py -E [-I input_folder] [-O output.xlsx] 

    eml2xlsx.py -E

    eml2xlsx.py -E -I C:\emails -O parsed_emails.xlsx

    eml2xlsx.py -M -I emails
	
	
	