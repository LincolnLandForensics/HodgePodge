
## vcf2xlsx.py

Convert a folder of .vcf files to an intel sheet, or visa versa


Installation:
```
python pip install -r requirements_vcf2xlsx.txt
```

options:

  -h, --help            show this help message and exit

  -I INPUT, --input INPUT

  -O OUTPUT, --output OUTPUT

  -B, --blank           create blank intel sheet

  -c, --contacts        Read contacts from .vcf files and create an Excel sheet

  -x, --xlsx            Read contacts from .xlsx files and create .vcf files


Usage:

```
python vcf2xlsx.py -c
python vcf2xlsx.py -c -I LogsVCF -O contacts_Apple.xlsx
python vcf2xlsx.py -B # create a blank
python vcf2xlsx.py -x -I contacts_Apple.xlsx -O LogsVCF	
```

note:

A vCard is a digital file format for storing and sharing contact information. The file typically has a .vcf extension and is widely used for exchanging contact details across various platforms and devices, such as smartphones, email clients, and CRM systems.