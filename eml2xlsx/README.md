
## eml2xlsx.py 
Parses .eml and .mbox files in a folder, extracts messages / non-spam contacts, and exports to Excel.

NDCAC GoogleReturnViewer.exe does convert .json files to .eml, but then mostly crashes. This was my interum solution.

Installation:
```
python pip install -r requirements_eml2xlsx.txt
```

help menu
```
python eml2xlsx.py -h
```

Examples:

eml2xlsx.py -E [-I input_folder] [-O output.xlsx] 

    eml2xlsx.py -E

    eml2xlsx.py -E -I C:\emails -O parsed_emails.xlsx

	
	
	