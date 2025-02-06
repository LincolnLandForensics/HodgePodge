#!/bin/bash
# version 1.0.6

# output_dir="$(pwd)"

triage_file="TRIAGE__$(hostname).txt"

# Create the triage_file
touch "$triage_file"

echo "TriageStartTime:  " $(date) >"$triage_file"
echo >> "$triage_file"

is_linux() {
    # Check if the operating system is Linux
    if [ "$(uname -s)" = "Linux" ]; then
        return 0  # Return 0 for true
    else
        return 1  # Return 1 for false
    fi
}

is_mac() {
    # Check if the operating system is macOS (Darwin)
    if [ "$(uname -s)" = "Darwin" ]; then
        return 0  # Return 0 for true
    else
        return 1  # Return 1 for false
    fi
}


# this functions from : https://github.com/vm32/Digital-Forensics-Script-for-Linux
write_output() {
    command=$1
    filename=$2
    echo >> "$triage_file"
    echo "<========   $command ========>" >>"$triage_file"
    if $command >> "$triage_file" 2>&1; then
    # if $command > "$output_dir/$filename" 2>&1; then

        echo >> "$triage_file"
    # else
        # echo >> "$triage_file"
    fi
}

echo "<======== TRIAGE INFORMATION ========>" >TRIAGE__$(hostname).txt
echo "<======== TRIAGE INFORMATION ========>"

echo >>TRIAGE__$(hostname).txt
echo "TriageTime:                " $(date) >> TRIAGE__$(hostname).txt
echo "CurrentUnixTime:           " $(date +%s) >> TRIAGE__$(hostname).txt
echo "Host Name:                 " $(hostname) >>TRIAGE__$(hostname).txt
echo "Host Name:                 " $(hostname)
echo "Current User:              " $(whoami) >>TRIAGE__$(hostname).txt
echo "Current User:              " $(whoami)
echo "Users:                     " $(users) >>TRIAGE__$(hostname).txt
echo "Timezone:                  " $(cat /etc/timezone)
echo "Timezone:                  " $(cat /etc/timezone) >>TRIAGE__$(hostname).txt
echo "Uptime:                    " $(uptime -p) >>TRIAGE__$(hostname).txt
echo "SystemStartupTime:         " $(uptime -s) >>TRIAGE__$(hostname).txt



# OS detection
if is_linux; then
    echo "OS System:                  Linux"
	# write_output "hostnamectl | grep 'Operating System' | sed 's/  Operating System: /OS Name: /g'" "$triage_file"	# todo
	hostnamectl | grep 'Operating System' | sed 's/  Operating System: /OS Name: /g' >>TRIAGE__$(hostname).txt
	hostnamectl | grep 'Operating System' | sed 's/  Operating System: /OS Name: /g'

elif is_mac; then
    echo "OS System:                  macOS."
	write_output "sw_vers" "$triage_file"
else

    echo "The system is not running Linux or macOS."
fi


write_output "cat /etc/hosts" "$triage_file"

write_output "ifconfig |grep 'inet'" "$triage_file"

write_output "ip addr |grep 'inet'" "$triage_file"

write_output "arp -a" "$triage_file"

write_output "lsof -i" "$triage_file"

write_output "who '-a'" "$triage_file"	#test

write_output "w" "$triage_file"

write_output "env" "$triage_file"

write_output "df -h" "$triage_file"


# test me

if crontab -l &>/dev/null; then
    write_output "crontab -l" "$triage_file"
else
    echo "<========   crontab -l &>/dev/null ========>	No crontab for root" >>"$triage_file"
fi


if command -v hwclock &>/dev/null; then
    write_output "hwclock -r" "$triage_file"
else
    echo "hwclock command not found" >> "$triage_file"
fi

# Operating System Installation Date
write_output "df -P /" "$triage_file"
write_output "ls -l /var/log/installer" "$triage_file"
write_output "tune2fs -l /dev/sda1" "$triage_file" # Check for correct root partition


# Installed Programs

write_output "rpm -qa" "$triage_file"

# sudo user specific

# Check if the script has root or equivalent privileges
# if groups | grep -q 'sudo'; then
if [ "$EUID" -eq 0 ]; then
# if groups | grep -q '\bsudo\b'; then
    echo "$(whoami) has sudo privileges"
    echo "$(whoami) has sudo privileges" >> "$triage_file"

    echo "<======== sudo cat ~/.ssh/authorized_keys ========> " >> TRIAGE__$(hostname).txt
    sudo cat ~/.ssh/authorized_keys >> TRIAGE__$(hostname).txt

    echo "<======== sudo cat ~/.ssh/known_hosts ========> " >> TRIAGE__$(hostname).txt
    sudo cat ~/.ssh/known_hosts >> TRIAGE__$(hostname).txt

    echo "<======== sudo cat ~/.ssh/config ========> " >> TRIAGE__$(hostname).txt
    sudo cat ~/.ssh/config >> TRIAGE__$(hostname).txt
	
	if is_linux; then
		echo "<======== sudo dmidecode ========> " >TRIAGE__$(hostname)_hardware.txt	#NotMacOS
		echo $(sudo dmidecode) >>TRIAGE__$(hostname)_hardware.txt		#NotMacOS#
		echo "Make:                      " $(sudo dmidecode -s system-manufacturer)
		echo "Make:                      "	$(sudo dmidecode -s system-manufacturer) >>TRIAGE__$(hostname).txt		#NotMacOS

		echo "Version:                   " $(sudo dmidecode -s system-version)
		echo "Version:                   " $(sudo dmidecode -s system-version) >>TRIAGE__$(hostname).txt		#NotMacOS

		echo "product-name:              " $(dmidecode -s system-product-name)
		echo "product-name:              " $(dmidecode -s system-product-name) >>TRIAGE__$(hostname).txt		#NotMacOS

		echo "Serial:                    " $(dmidecode -s system-serial-number)
		echo "Serial:                    " $(dmidecode -s system-serial-number) >>TRIAGE__$(hostname).txt		#NotMacOS

		echo "<======== sudo dmidecode ========>" >> TRIAGE__$(hostname)_dmidecode.txt
		echo $(sudo dmidecode >> TRIAGE__$(hostname)_dmidecode.txt)		#NotMacOS
		# write_output "dmidecode" "$triage_file"

else
    echo "$(whoami) isn't root or in SUDO mode" >> "$triage_file"
    echo "$(whoami) isn't root or in SUDO mode"  
fi
 
# OS specific commands
if is_linux; then
    echo "The system is running Linux."

	write_output "lsblk -io KTYPE,TYPE,SIZE,MODEL,FSTYPE,UUID,MOUNTPOINT" "$triage_file"
	write_output "cat /etc/passwd" "$triage_file"
	write_output "cat /etc/group" "$triage_file"
	
	# echo "<======== journalctl ========> " >TRIAGE__$(hostname)_journalctl.txt	# too big
	# echo $(journalctl) >>TRIAGE__$(hostname)_journalctl.txt	#NotMacOS
	# write_output "journalctl" "$triage_file"	#NotMacOS

	if [ "$EUID" -eq 0 ]; then
		echo "$(whoami) is root or in SUDO mode" >> "$triage_file"
		echo "<======== sudo cat /etc/shadow ========> " >> TRIAGE__$(hostname).txt
		sudo cat /etc/shadow >> TRIAGE__$(hostname).txt

		echo "<======== sudo cat /etc/sudoers ========> " >> TRIAGE__$(hostname).txt
		sudo cat /etc/sudoers >> TRIAGE__$(hostname).txt

		echo "<======== lshw -class disk ========> " >> TRIAGE__$(hostname).txt
		sudo lshw -class disk | sed s/'       serial: /Drive Serial Number: /g' | sed s/'       size:/Source data size:/g' | sed s/'       product:/Model:/g'>> TRIAGE__$(hostname).txt
	else
		echo "$(whoami) isn't root or in SUDO mode" >> "$triage_file"
	fi
elif is_mac; then
    echo "The system is running macOS."
	echo "Uptime:                    " $(uptime) >>TRIAGE__$(hostname).txt
	write_output "diskutil ap list" "$triage_file"
	write_output "security list-keychains" "$triage_file"

else
    echo "The system is not running Linux or macOS."
fi

    
# MACOS specific
# echo $(diskutil ap list >> TRIAGE__$(hostname).txt)
# write_output "diskutil ap list" "$triage_file"

# security list-keychains >> TRIAGE__$(hostname).txt
# write_output "security list-keychains" "$triage_file"



# output is too big, stick at the bottom


# System Logs and Usage
write_output "ls -lah /var/log/" "$triage_file"

write_output "netstat -nat | grep 'tcp'" "$triage_file"

# write_output "netstat -i" "$triage_file"	# test

write_output "netstat -rn" "$triage_file"

write_output "last" "$triage_file"
write_output "lshw -short" "$triage_file"

# Hardware Information
write_output "lspci" "$triage_file"

write_output "dpkg -l" "$triage_file" # Replaced 'apt' with 'dpkg -l'   # installed programs


# echo "<======== cat ~/.bash_history | nl ========> " > TRIAGE__$(hostname)_bash_history.txt
# cat ~/.bash_history | nl >> TRIAGE__$(hostname)_bash_history.txt

for user_home in /home/*; do
    username=$(basename "$user_home")
    echo "$username" >>"$triage_file"
    
    if [ -f "$user_home/.bash_history" ]; then
        write_output "cat $user_home/.bash_history" "$triage_file"
    else
        echo "No .bash_history for $username" >> "$triage_file"
    fi
    
    if [ -f "$user_home/.zsh_history" ]; then
        write_output "cat $user_home/.zsh_history" "$triage_file"
    else
        echo "No .zsh_history for $username" >> "$triage_file"
    fi
    
    write_output "cat $user_home/.local/share/recently-used.xbel" "$triage_file"
done

echo "<======== arp -an ========>" >> TRIAGE__$(hostname)_arp_an.txt
echo $(arp -an >> TRIAGE__$(hostname)_arp_an.txt)
# write_output "arp -an" "$triage_file"

# echo "<======== netstat -an ========>" >> TRIAGE__$(hostname)_netstat_an.txt
# echo $(netstat -an >> TRIAGE__$(hostname)_netstat_an.txt)
write_output "netstat -an" "$triage_file"

# write_output "lsof" "$triage_file"
echo "<======== lsof ========> " > TRIAGE__$(hostname)_lsof.txt
echo $(lsof >> TRIAGE__$(hostname)_lsof.txt)

write_output "ps aux" "$triage_file"


echo "<======== Done! See $triage_file. ========> "
echo "TriageEndTime:                   " $(date) >> TRIAGE__$(hostname).txt

echo "<======== The End ========>" >> TRIAGE__$(hostname).txt

# os version

# version and serial number
# /System/Library/CoreServices/SystemVersion.plist


# <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

# Copyright (C) 2023 LincolnLandForensics
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License version 2, as published by the
# Free Software Foundation
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details (http://www.gnu.org/licenses/gpl.txt).


