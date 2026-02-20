# Kismet Auto Start Wi Fi + Bluetooth + GPS + Wigle Logging on Kali (Raspberry Pi 5)

This document describes the **one time manual setup** required to deploy a Raspberry Pi 5 running Kali Linux as an autonomous Kismet sensor. The configuration includes:

- Wi Fi capture using **Panda PAU0F AXE3000** (`wlan1`)
    
- Bluetooth/BLE capture using **internal Pi 5 Bluetooth** (`hci0`)
    
- GPS capture via **gpsd** (`/dev/ttyACM0`)
    
- Wigle compatible CSV auto logging
    
- Auto start on boot via systemd
    
- Field friendly log directory structure
    
- Drop in configuration files
    

---

# 1. Install Kismet + Plugins

```
sudo apt update
sudo apt install kismet kismet-plugins
```

---

# 2. Prepare Log Directory

```
sudo mkdir -p /var/log/kismet
sudo chown kismet:kismet /var/log/kismet
```

Kismet will write all logs here, including Wigle CSVs.

---

# 3. Configure Wi Fi Capture (wlan1)

Your Panda PAU0F AXE3000 will enumerate as `wlan1`.

Kismet handles monitor mode automatically; no manual `airmon-ng` steps are required.

---

# 4. Configure Bluetooth Capture (hci0)

Enable Bluetooth:

```
sudo systemctl enable bluetooth
sudo systemctl start bluetooth
sudo hciconfig hci0 up
```

---

# 5. Configure GPSD (GPS on /dev/ttyACM0)

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

---

# 6. Install Drop In Kismet Config Files

## /etc/kismet/kismet.conf

```
# Wi Fi capture source (Panda PAU0F AXE3000)
source=wlan1:name=wifi1

# Bluetooth capture source (internal Pi 5 Bluetooth)
source=bluetooth:hci0

# Log directory
log_prefix=/var/log/kismet/kismet
```

---

## /etc/kismet/kismet_logging.conf

```
# Core Kismet logs
log_types=pcapng,netxml,nettxt

# Wigle-compatible CSV export
log_types+=wiglecsv

# Timestamped filenames for chain-of-custody clarity
log_prefix=kismet-$(date +%Y%m%d-%H%M%S)
```

---

# 7. Create Systemd Service for Auto Start

Create:

```
sudo nano /etc/systemd/system/kismet.service
```

Paste:

```
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
```

Enable + start:

```
sudo systemctl enable kismet
sudo systemctl start kismet
```

Kismet will now automatically:

- Start on boot
    
- Capture Wi Fi + Bluetooth
    
- Use GPSD
    
- Write Wigle CSV logs
    
- Store everything in `/var/log/kismet/`
    

---

# 8. GPSD Sanity Check Script

Create:

```
sudo nano /usr/local/bin/gps-check.sh
```

Paste:

```
#!/bin/bash
echo "[+] Checking GPSD status..."
systemctl status gpsd --no-pager

echo "[+] Checking GPS device..."
ls -l /dev/ttyACM0

echo "[+] Testing GPS feed..."
timeout 5 cgps || echo "[-] GPS not responding"
```

Make executable:

```
sudo chmod +x /usr/local/bin/gps-check.sh
```

---

# 9. Final Verification

### Wi Fi source:

```
iwconfig wlan1
```

### Bluetooth source:

```
hciconfig hci0
```

### GPS:

```
cgps
```

### Kismet logs:

```
ls /var/log/kismet/
```

You should see:

```
kismet-20260220-1240-1.wiglecsv
kismet-20260220-1240-1.netxml
kismet-20260220-1240-1.pcapng
```

---

# 10. Export Kismet Logs to KML with Strongest Signal MAC Addresses

Kismet includes a tool called `kismet_log_to_kml` that can export device data from Kismet logs to KML format. This tool supports options to filter devices by minimum signal strength and to plot points based on the strongest signal detected.

To generate a KML file containing only MAC addresses using the strongest signal detected, use the following options:

- `--in` to specify the input `.kismet` log file
    
- `--out` to specify the output `.kml` file
    
- `--strongest-point` to plot points based on the strongest signal seen for each device
    
- `--min-signal` to filter devices by minimum signal strength (optional)
    

Example usage:

```
kismet_log_to_kml --in /var/log/kismet/kismet-20260220-1240-1.netxml --out /var/log/kismet/devices.kml --strongest-point
```

This will produce a KML file with device locations tagged by their MAC addresses, using the strongest signal point per device.

Ensure your Kismet logs include location data (GPS) for accurate mapping.

---

# Deployment Complete

Your Raspberry Pi 5 now operates as a fully autonomous Kismet sensor:

- Wi Fi + Bluetooth scanning
    
- GPS tagging
    
- Wigle compatible CSV logging
    
- Auto start on boot
    
- Forensic grade timestamped logs