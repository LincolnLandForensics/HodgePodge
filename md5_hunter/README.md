
## md5_hunter.py 
read pdf's and other specific filetypes and export out unique md5 hashes
and words for password cracking and file hunting.

filetypes: .pdf, .txt, .docx

throw lots of files into the 'pdfs' folder and run the script.

Exe version available upon request.



Installation:
```
python pip install -r requirements_md5_hunter.txt
```

Usage:


```
python md5_hunter.py
```


Note: the output is saved as words_{date}.txt and md5_hashes_{date}.txt 


If you throw in an old list of md5's it will combine them all with a unique sort.


Example output:

reading files in pdfs folder.

Saved 908 unique sorted words to: words_2024-12-04.txt

Saved 153 unique sorted MD5 hashes to: md5_hashes_2024-12-04.txt


