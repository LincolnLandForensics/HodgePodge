


## GPS2Address.py
read GPS coordinates and convert them to addresses

or

read addresses and convert them to coordinates

and

read GPS coordinates (xlsx) and convert them to KML for review in earth.google.com

additional red lines for trips within a reasonable time. (travel_path_SAMPLE.kml)

Usage:
```
python GPS2Address.py -r
```
insert your GPS or addresses into locations.xlsx
```
Example:
    GPS2Address.py -c -O input_blank.xlsx
    GPS2Address.py -k -I locations.xlsx  # xlsx 2 kml with no internet processing
    GPS2Address.py -r
    GPS2Address.py -r -I locations.xlsx -O locations2addresses_.xlsx
```    
*   Visit earth.google.com, File,Import KML 

![Case Example](images/GPS.png)


![XLSX Example](images/GPS_xlsx.png)

Icons
*   Car -Lpr red car (License Plate Reader)
*   Car2 -Lpr yellow car
*   Car3 -Lpr greeen car with circle
*   Car4 -Lpr red car with circle
*   Truck -Lpr truck
	
*   Calendar
*   Home
*   Images -Photo
*   Intel -I
*   Locations -Reticle
*   default -Yellow flag
*   Office
*   Searched -Searched Item
*   Videos -Video clip
	
*   Toll -Blue square
*   N -Northbound blue arrow
*   E -Eastbound blue arrow
*   S -Southbound blue arrow
*   W -Westbound blue arrow
	
*   Chats
*   Tower -Bullseye

*   Yellow font -Tagged
*   blue lines -trips with a start and end
*   red circles -indicate radius of the signal and/or accuracy of the point

---


## WarDriveParser.py

Convert wigle .gz or .csv exports to gps2address.py locations format or convert HackRf logs.

Note: Wigle(dot)net can be used to query MAC address and SSID's. Wigle Wifi is an Android app that captures Bluetooth, Wifi, & Cell Tower info. HackRF can be used to sniff Bluetooth and more.


Usage:
```
python GPS2Address.py -h
```


```
options:
  -h, --help            show this help message and exit
  -I INPUT, --input INPUT
  -O OUTPUT, --output OUTPUT
  -b, --blank           create blank sheet
  -C, --clear           clear logs off the HackRF
  -L, --logs            log grabber (HackRF)
  -p, --parseHackRF     parse HackRF text
  -w, --wigleparse      parse wigle file csv
```



```
Example:
  python WarDriveParser.py -b      # create a blank sheet
  python WarDriveParser.py -C      # clear logs off the HackRF
  python WarDriveParser.py -L      # log grabber (HackRF)
  python WarDriveParser.py -p      # parse HackRF text
  python WarDriveParser.py -p -I logs -O WarDrive_.xlsx
  python WarDriveParser.py -w -I WigleWifi_Neighborhood.csv.gz     # parse wigle log
```    


![Syntax Example](images/WigleWiFI2.jpg)

*   Visit earth.google.com, File,Import KML 


![Example](images/WigleWiFI.jpg)

---
