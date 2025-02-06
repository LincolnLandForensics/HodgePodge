# This one is a beta version that needs a lot of work



## memprocFS_process.py 
https://github.com/ufrisk/MemProcFS

cd C:\Forensics\scripts\python\git-repo\MemProcFS_files_and_binaries_v5.8.1-win_x64-20230910

.\MemProcFS.exe -device C:\temp\memdump.raw  -forensic 1

maps it all as M:\


Once M: is mapped, run this script.


pulls out all the cool bits and throw them into a spreadsheet for a quick triage and case notes

Installation:
```
python pip install -r requirements_memprocFS.txt
```
