
# HackRF_parser.py

parse HackRF Mac Address logs

HackRF can track the signal strength. The closer the dB is to zero, the closer you are.

You can filter by mac address to track a specific device.

This is where this parser might come in handy, by identifying the company that made the device.

Eventually I might be able to decode the device type (ex. phone)

whitelist out known Mac's so you don't track something that you brought to the party.


## Example:
    python HackRF_parser.py -p -I H:\BLERX\Lists -O output_BleRX.xlsx


## Installation:
```
python pip install -r requirements_BleRX_parser.txt
```
or 
```
pip install openpyxl
```


## Usage:


```
python HackRF_parser.py
```



## Note:

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
