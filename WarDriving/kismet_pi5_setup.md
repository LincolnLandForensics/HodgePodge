## Kismet Auto-Start Wi-Fi + Bluetooth + GPS + Wigle Logging on Kali (Raspberry Pi 5)

This document describes the one-time manual setup required to deploy a Raspberry Pi 5 running Kali Linux as an autonomous Kismet sensor. The configuration includes:

Wi-Fi capture using Panda PAU0F AXE3000 (wlan1)

Bluetooth/BLE capture using internal Pi 5 Bluetooth (hci0)

GPS capture via gpsd (/dev/ttyACM0)

Wigle-compatible CSV auto-logging

Auto-start on boot via systemd

Field-friendly log directory structure

Drop-in configuration files

## 1. Install Kismet + Plugins

```
sudo apt update
sudo apt install kismet kismet-plugins
```

## 1.1 Create kismet user

```
sudo useradd -r -g kismet -s /usr/sbin/nologin kismet
```

## 2. Prepare Log Directory

```
sudo mkdir -p /var/log/kismet
sudo chown kismet:kismet /var/log/kismet
```

Kismet will write all logs here, including Wigle CSVs.

## 3. Configure Wi-Fi Capture (wlan1)

Your Panda PAU0F AXE3000 will enumerate as wlan1.

Kismet handles monitor mode automatically; no manual airmon-ng steps are required.

## 4. Configure Bluetooth Capture (hci0)

Enable Bluetooth:

```
sudo systemctl enable bluetooth
sudo systemctl start bluetooth
sudo hciconfig hci0 up
```

## 5. Configure GPSD (GPS on /dev/ttyACM0)

Install GPSD:

```
sudo apt install gpsd gpsd-clients
```

Enable GPSD:

```
sudo systemctl enable gpsd
sudo systemctl start gpsd
```



Test GPS:

```
cgps
```

If coordinates update, Kismet will automatically tag devices with GPS.

## 6. Install Drop-In Kismet Config Files


```
sudo nano /etc/kismet/kismet.conf
```


	# Wi-Fi capture source (Panda PAU0F AXE3000)

	source=wlan1:name=wifi0

	# Bluetooth capture source (internal Pi 5 Bluetooth)

	source=bluetooth:hci0

	# Log directory

	log_prefix=/var/log/kismet/

	gps=gpsdP:host=localhost,port=2947

```
sudo nano /etc/kismet/kismet_logging.conf
```


	# Core Kismet logs

	log_types=pcapng,netxml,nettxt

	# Wigle-compatible CSV export

	log_types+=wiglecsv

	# Timestamped filenames for chain-of-custody clarity

	log_prefix=kismet-$(date +%Y%m%d-%H%M%S)

## 7. Create Systemd Service for Auto‑Start and Dedicated User



### Create systemd service file

```
sudo nano /etc/systemd/system/kismet.service
```

Paste:

[Unit]


Description=Kismet Wireless Scanner

After=network.target bluetooth.target gpsd.service

[Service]


User=kismet

Group=kismet

ExecStart=/usr/bin/kismet

WorkingDirectory=/var/log/kismet

Restart=always

[Install]


WantedBy=multi-user.target



## Enable and start service

```
sudo systemctl daemon-reload
sudo systemctl enable kismet
sudo systemctl start kismet
```



Kismet will now automatically:

Run as the dedicated kismet user

Start on boot regardless of user login

Capture Wi-Fi + Bluetooth

Use GPSD

Write Wigle CSV logs

Store everything in /var/log/kismet/

8. GPSD Sanity-Check Script

Create:

```
sudo nano /usr/local/bin/gps-check.sh
```

Paste:

```
#!/bin/bash
echo "\[+] Checking GPSD status..."
systemctl status gpsd --no-pager

echo "\[+] Checking GPS device..."
ls -l /dev/ttyACM0

echo "\[+] Testing GPS feed..."
timeout 5 cgps || echo "\[-] GPS not responding"
```



Make executable:

```
sudo chmod +x /usr/local/bin/gps-check.sh
```



9. Final Verification

Wi-Fi source:

```
iwconfig wlan1
```



Bluetooth source:

```
hciconfig hci0
```



GPS:

```
cgps
```



Kismet logs:

```
ls /var/log/kismet/
```

## manually start kismet


```
kismet -c wlan1
```


```
https://localhost:2501
```










You should see:

kismet-20260220-1240-1.wiglecsv
kismet-20260220-1240-1.netxml
kismet-20260220-1240-1.pcapng

Deployment Complete

Your Raspberry Pi 5 now operates as a fully autonomous Kismet sensor:

Wi-Fi + Bluetooth scanning

GPS tagging

Wigle-compatible CSV logging

Auto-start on boot

Forensic-grade timestamped logs



## 10. convert kismet to kml

```
kismetdb_to_kml --in kismet-20240220-1530-1.kismet --out kismet.kml
```




