
## cellebrite_parser.py 
convert Cellebrite contacts, account, web history, chats and call exports to intel format

note: the spreadsheet has to have the top row deleted so that the columns become the first row.


Installation:
```
python pip install -r requirements_cellebrite_parser.txt
```

Usage:\
process call logs
```
python cellebrite_parser.py -c -I calls.xlsx
```

create blank input_blank.xlsx
```
python cellebrite_parser.py -b -O input_blank.xlsx
```

help menu
```
python cellebrite_parser.py
```

Example:

    cellebrite_parser.py -b -O input_blank.xlsx
	
    cellebrite_parser.py -C -I Accounts.xlsx
	
    cellebrite_parser.py -C -I Calls.xlsx
	
    cellebrite_parser.py -C -I Chats.xlsx
	
    cellebrite_parser.py -C -I Contacts.xlsx
	
    cellebrite_parser.py -C -I SearchedItems.xlsx
	
    cellebrite_parser.py -C -I WebHistory.xlsx
	
	
	
![sample output](images/Intel_Contacts_Sample.png)	
	


## CellebriteEmailxlsx2xlsx.py

Read a Cellebrite email export parse it and export it out.

This example reads SmilePOS emails and parses out the details.

You can re-arrange the headers to meet your needs.


Example:

'''
   python  CellebriteEmailxlsx2xlsx.py -r
   
   python CellebriteEmailxlsx2xlsx.py -r -I Cellebrite_Emails.xlsx -O Cellebrite_Emails_Parsed.xlsx
'''



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
	
	
	