# Triage
Various Triage scripts


<======== triage_windows.cmd ========>

Purpose: To quickly collect pertienent information from a Windows systems such as 
Hostname, user domain, email, current user, timezone,  system model, logon server, onedrive info, net share, bitlocker encryption status, systeminfo, ipconfig. The uniq outputfile has a start and end timestamp for you lab records.
The output file also records the command that was used to capture the information for posterity's sake.
Running as admin is preferred but not required.

This grabs SAM and SYSTEM for password cracking. Group policies can alert on "Credential harvesting" in an Enterprise environment. Replace your domain name for the sample one (ILLINOIS) if you want to skip this step on an your enterpise system.

Steps: 
1. R.click the script and select "run as administrator"

It spits very pertinent stuff to the screen to help fill out your evidence recovery log.
It sames the output to _TRIAGE__<computername>.txt in your current working directory. This allows uniq output files in case of multiple Triage's.


<======== triage_linux_mac.sh ========>
  
Purpose: To quickly collect pertienent information from a Linux or MAC system 
Hostname, OS, current user, timezone, uptime, ifconfig, netstat, arp, ssh keys, bash history, /etc/shadow. The uniq outputfile has a start and end timestamp for you lab records.
The output file also records the command that was used to capture the information for posterity's sake.
Running as admin is preferred but not required.

Command:
open up a terminal window: 
sudo triage_linux_mac.sh



<Future wishlist>
Let me know if you think I've missed something important.
