Print Job Normalizer

(super beta version)

A GUI tool for processing CUPS print job spool files and normalizing them into standard, readable formats.

The script detects each file’s MIME type, assigns the correct extension, and converts raster or spool formats into JPEG images.

It is designed for forensic workflows, print job analysis, and environments where raw CUPS spool files need to be interpreted.



Features

✔ MIME type detection for extensionless print job files



✔ Normalization of all file types



PDFs → .pdf



PostScript → .ps



Raster/spool formats → .jpg



Text → .txt



Unknown → .bin



✔ ImageMagick powered conversion (with Windows validation)



✔ CUPS cupsfilter support on Linux



Installation

1\. Install Python 3.8+

Windows, macOS, or Linux are supported.



2\. Install required Python packages

bash

pip install filetype

3\. Install ImageMagick

Windows

Download from:

https://imagemagick.org/script/download.php (imagemagick.org in Bing)



During installation, enable:



Add application directory to your system PATH



Install legacy utilities (convert) (optional)



The script verifies ImageMagick on startup and exits with an error if missing.



Ubuntu / Debian / WSL

bash

sudo apt update

sudo apt install imagemagick

macOS (Homebrew)

bash

brew install imagemagick

Usage

1\. Prepare folders

The script defaults to:



Code

input/

output/

You may change these in the GUI.



2\. Run the script

bash

python3 print\_job\_normalizer.py

3\. Use the GUI

Select Input Folder



Select Output Folder



Click Convert



The script will:



Scan all files in the input folder



Detect MIME types



Normalize each file to the correct extension



Convert raster/spool formats to JPEG



Write results to the output folder



Display progress and status updates in the GUI and terminal



How It Works

Files with extensions

These are copied directly to the output folder.



Files without extensions

The script:



Detects MIME type



Determines the correct extension



Normalizes the file:



MIME Type Category	Output

PDF	.pdf

PostScript	.ps

Raster / CUPS Raster / Apple Raster	.jpg (converted)

Image formats	.jpg (converted)

Text	.txt

Unknown	.bin





This ensures every file becomes readable and properly labeled.



Why This Exists

CUPS print spool files often lack extensions and may contain:



PDF



PostScript



CUPS Raster



Apple Raster



URF



PBM/PGM/PPM/PNM



TIFF



Raw binary



This tool makes them human readable and easy to analyze by automatically normalizing them.



License

MIT License

