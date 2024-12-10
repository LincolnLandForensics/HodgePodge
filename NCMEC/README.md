
## NCMEC_PDF_parser.py 

Dump all of your NCMEC .zip or .pdf files into the NCMEC folder.
it will export the emails, ip's, md5's, phone numbers and users into _output folder

Example:
    python NCMEC_PDFs_parser.py

Note:
	try md5_hunter.py for one that just does md5's.


Installation:
```
python pip install -r requirements_NCMEC.txt
```

Usage:


```
python NCMEC_PDF_parser.py
```


Flowchart:
	
![sample output](images/NCMEC_flowchart.png)



Sample output: 

NCMEC_PDFs_parser.py

5 files unzipped in NCMEC folder

Processing PDFs in NCMEC folder....

        NCMEC\compressed_pdf_unlocked\test.pdf

        NCMEC\compressed_pdf_unlocked\tes2.pdf

        NCMEC\compressed_pdf_unlocked\tester.pdf

        NCMEC\compressed_pdf_unlocked_2024-12-10\test.pdf

        NCMEC\compressed_pdf_unlocked_2024-12-10\test2.pdf

See text files in NCMEC\_output folder.