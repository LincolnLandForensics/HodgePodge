
# WarDriveParser.py

Python script to parse and visualize Wigle(.)net and HackRF data. 

Converts Wigle (.gz/.csv) and HackRF logs to gps2address.py format, retaining the strongest signals for precise geolocation. 

Maps MAC addresses to company names.


## Example:
  python WarDriveParser.py -b      # create a blank sheet

  python WarDriveParser.py -C      # clear logs off the HackRF

  python WarDriveParser.py -L      # log grabber (HackRF)

  python WarDriveParser.py -p      # parse HackRF text

  python WarDriveParser.py -p -I logs -O WarDrive_.xlsx

  python WarDriveParser.py -w -I WigleWifi_Neighborhood.csv.gz     # parse wigle log

  python WarDriveParser.py -w -I WigleWifi_sample.csv

## Installation:
```
python pip install -r requirements_WarDriveParser.txt
```
or 
```
pip install openpyxl
```


## Usage:


```
python WarDriveParser.py
```



## Note:

When the Wigle app is used in conjunction with GPS2Address.py, it can create google earth files (.KML) like this:

![sample KML](Images/Wigle_Wifi1.png)


Wigle(.)net has an app for Android devices.

Bluetooth was captured with HackRF Porta pack h4

Install a Comet antenna (insert correct length here)
turn on HackRF
(Receive)(BLE RX)

Logs are saved in: H:\BLERX\Lists\*.csv

and

Logs are saved in: H:\BLERX\Logs\*.txt

additional logs to parse:
H:\LOGS\ADSB.TXT
H:\LOGS\AIS.TXT
H:\LOGS\random.TXT
H:\LOGS\TPMS.TXT
H:\LOGS\APRS.TXT




![sample output](Images/HackRF_BLE_RX.jpg)


[Wiki](https://github.com/portapack-mayhem/mayhem-firmware/wiki/Bluetooth-Low-Energy-Receiver)


# Wigle_Query.py

Read MAC addresses from xlsx

For all Types that are WIFI, BT or BLE it will check wigle(.)net for matches

Requires a wigle API key (user:password)


## Example:

   python Wigle_Query.py -b -O input_blank.xlsx
   
   python Wigle_Query.py -Q -I input_.xlsx

## Installation:
```
python pip install -r requirements_WarDriveParser.txt
```


## Usage:


```
python WarDriveParser.py
```


## Note:

MAC addresses must accompany the type. (WIFI, BLT, BT - currently supported)

