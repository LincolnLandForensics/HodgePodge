@echo off
rem version 1.0.9

rem this will triage a windows PC (run as administrator, if you can)

echo ======== TRIAGE INFORMATION ======== >_TRIAGE__%computername%.txt
echo ======== TRIAGE INFORMATION ========

echo StartTime >>_TRIAGE__%computername%.txt
powershell invoke-command -scr {(get-date -format "{MMM-dd-yyyy HH:mm:ss})} >>_TRIAGE__%computername%.txt

echo. >>_TRIAGE__%computername%.txt

echo Host Name:                 %computername% >>_TRIAGE__%computername%.txt
echo Host Name:                 %computername%
echo UserDomain:                %UserDomain% >>_TRIAGE__%computername%.txt
echo Currentuser:              %username% >>_TRIAGE__%computername%.txt
echo Currentuser:              %username%
echo Email:                     %email% >>_TRIAGE__%computername%.txt

rem determine Timezone
setlocal
for /f "tokens=*" %%f in ('tzutil /g') do (echo Timezone:                  %%f >>_TRIAGE__%computername%.txt)
for /f "tokens=*" %%f in ('tzutil /g') do (echo Timezone:                  %%f )
endlocal

REM echo OS Name:                   >>_TRIAGE__%computername%.txt
REM ver |find "Version" >>_TRIAGE__%computername%.txt

wmic computersystem get totalphysicalmemory >>_TRIAGE__%computername%.txt
wmic computersystem get totalphysicalmemory

echo System Model:                >>_TRIAGE__%computername%.txt
echo System Model:
wmic csproduct get name | find "Name " /v >>_TRIAGE__%computername%.txt
wmic csproduct get name | find "Name " /v 

echo Homedrive:                 %HOMEDRIVE% >>_TRIAGE__%computername%.txt
echo LogonServer:               %LogonServer% >>_TRIAGE__%computername%.txt
echo OneDrive:                  %OneDrive% >>_TRIAGE__%computername%.txt

echo ======== powershell invoke-command -scr {(get-ciminstance -classname win32_logicaldisk)} >>_TRIAGE__%computername%.txt
powershell invoke-command -scr {(get-ciminstance -classname win32_logicaldisk)} >>_TRIAGE__%computername%.txt

echo ======== net share ======== >>_TRIAGE__%computername%.txt
wmic share get Caption, Description, Name, Path, Status >>_TRIAGE__%computername%.txt
wmic share get Caption, Description, Name, Path, Status

rem this requires administrator access
echo ======== powershell invoke-command -scr {(Get-BitLockerVolume)} >>_TRIAGE__%computername%.txt
powershell invoke-command -scr {(Get-BitLockerVolume)} >>_TRIAGE__%computername%.txt
powershell invoke-command -scr {(Get-BitLockerVolume)}

rem Backup BitLocker Recovery Key in Command Prompt
echo ======== manage-bde -protectors -get c: ======== >>_TRIAGE__%computername%.txt
rem  grab bitlocker recovery key for each drive (or at least c-f)
manage-bde -protectors -get c:  >>_TRIAGE__%computername%.txt
manage-bde -protectors -get d:  >>_TRIAGE__%computername%.txt
manage-bde -protectors -get e:  >>_TRIAGE__%computername%.txt
manage-bde -protectors -get f:  >>_TRIAGE__%computername%.txt


echo ======== manage-bde -status c: ======== >>_TRIAGE__%computername%.txt
rem get encryption status
manage-bde -status c:  >>_TRIAGE__%computername%.txt
manage-bde -status d:  >>_TRIAGE__%computername%.txt
manage-bde -status e:  >>_TRIAGE__%computername%.txt
manage-bde -status f:  >>_TRIAGE__%computername%.txt


echo ======== systeminfo ======== >>_TRIAGE__%computername%.txt
systeminfo >>_TRIAGE__%computername%.txt
systeminfo


echo ======== fsutil fsinfo ntfsinfo c: ======== >>_TRIAGE__%computername%.txt
fsutil fsinfo ntfsinfo c:  >>_TRIAGE__%computername%.txt
fsutil fsinfo ntfsinfo d:  >>_TRIAGE__%computername%.txt
fsutil fsinfo ntfsinfo e:  >>_TRIAGE__%computername%.txt
fsutil fsinfo ntfsinfo f:  >>_TRIAGE__%computername%.txt


echo. >>_TRIAGE__%computername%.txt
echo ======== ipconfig ======== >>_TRIAGE__%computername%.txt
ipconfig | find "IPv" | find ":" >>_TRIAGE__%computername%.txt


echo ======== net user ======== >>_TRIAGE__%computername%.txt
wmic useraccount get Disabled,FullName,Name,PasswordRequired,SID >>_TRIAGE__%computername%.txt
wmic useraccount get Disabled,FullName,Name,PasswordRequired,SID
rem net user | find "command completed" /v

echo ======== net localgroup administrators ======== >>_TRIAGE__%computername%.txt

net localgroup administrators | find "command completed" /v | find "Members" /v | find "Comment" /v >>_TRIAGE__%computername%.txt
net localgroup administrators | find "command completed" /v | find "Members" /v | find "Comment" /v


echo ======== netstat -nat ======== >>_TRIAGE__%computername%.txt
netstat -nat | find "TCP"  >>_TRIAGE__%computername%.txt

echo ======== arp -a ======== >>_TRIAGE__%computername%.txt
arp -a >>_TRIAGE__%computername%.txt


echo ======== ipconfig /displaydns ======== >>_TRIAGE__%computername%.txt
ipconfig /displaydns >>_TRIAGE__%computername%.txt
echo.

echo ======== netsh wlan show all ======== >>_TRIAGE__%computername%.txt
netsh wlan show all >>_TRIAGE__%computername%.txt

echo ======== wmic startup get caption,command ======== >>_TRIAGE__%computername%.txt
echo startup commands:               >>_TRIAGE__%computername%.txt
wmic startup get caption,command >>_TRIAGE__%computername%.txt
echo.

echo ======== powershell.exe -command "Get-History" ======== >>_TRIAGE__%computername%.txt
rem powershell.exe -command "Get-History" >>_TRIAGE__%computername%.txt
powershell invoke-command -scr {(Get-History)} >>_TRIAGE__%computername%.txt

rem this will light up an incident response alert as "credential dumping"
rem this will skip this task if you are in the illinois domain
if %UserDomain% == ILLINOIS (echo ILLINOIS) else (reg save hklm\sam SAM_%computername%)
if %UserDomain% == ILLINOIS (echo ILLINOIS) else (reg save hklm\system SYSTEM_%computername%)



rem How well patched is the system?
rem wmic qfe get Caption,Description,HotFixID,InstalledOn

rem copy a setup file with plaintext passwords if it exists
IF EXIST C:\Windows\Panther\Unattend\Unattend.xml COPY C:\Windows\Panther\Unattend\Unattend.xml Unattend__%computername%.xml

echo ======== Done! Files are located in your working directory. ======== 
echo ======== The End ======== >> _TRIAGE__%computername%.txt
echo EndTime:                  >>_TRIAGE__%computername%.txt
powershell invoke-command -scr {(get-date -format "{MMM-dd-yyyy HH:mm:ss})} >>_TRIAGE__%computername%.txt


echo.
echo see _TRIAGE__%computername%.txt for output

echo ======== netsh wlan export profile key=clear folder=. ======== 
netsh wlan export profile key=clear folder=.
echo see WIFI-*.xml for Wifi profiles

pause


REM # <<<<<<<<<<<<<<<<<<<<<<<<<<      Future Wishlist        >>>>>>>>>>>>>>>>>>>>>>>>>>
rem powershell to grab Chrome saved credentials
https://sushant747.gitbooks.io/total-oscp-guide/content/privilege_escalation_windows.html

REM # <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

REM # Copyright (C) 2022 LincolLandForensics
REM #
REM # This program is free software; you can redistribute it and/or modify it under
REM # the terms of the GNU General Public License version 2, as published by the
REM # Free Software Foundation
REM #
REM # This program is distributed in the hope that it will be useful, but WITHOUT
REM # ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
REM # FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
REM # details (http://www.gnu.org/licenses/gpl.txt).

