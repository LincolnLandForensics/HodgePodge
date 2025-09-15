@echo off
setlocal enabledelayedexpansion
REM color 07	# green font, black background

rem version 1.1.2

rem this will triage a windows PC (run as administrator, if you can)

set "logfile=_TRIAGE__%computername%.txt"

echo #

echo ======== TRIAGE INFORMATION ======== >%logfile%
echo ======== TRIAGE INFORMATION ========

:: Extract date components
for /f "tokens=2-4 delims=/ " %%a in ("%date%") do (
    set "mm=%%a"
    set "dd=%%b"
    set "yyyy=%%c"
)

:: Extract time components
set "hh=%time:~0,2%"
set "mn=%time:~3,2%"
set "ss=%time:~6,2%"

:: Handle single-digit hour (e.g., " 9" becomes "09")
if "%hh:~0,1%"==" " set "hh=0%hh:~1,1%"

:: Construct StartTime
set "StartTime=%yyyy%-%mm%-%dd% %hh%:%mn%:%ss%"

:: Echo result
echo StartTime = %StartTime% >>%logfile%





REM echo StartTime >>%logfile%
powershell invoke-command -scr {(get-date -format "{MMM-dd-yyyy HH:mm:ss})} >>%logfile%

echo. >>%logfile%

echo Host Name:                 %computername% >>%logfile%
echo Host Name:                 %computername%
echo UserDomain:                %UserDomain% >>%logfile%
echo Currentuser:              %username% >>%logfile%
echo Currentuser:              %username%
echo Email:                     %email% >>%logfile%

rem determine Timezone
setlocal
for /f "tokens=*" %%f in ('tzutil /g') do (echo Timezone:                  %%f >>%logfile%)
for /f "tokens=*" %%f in ('tzutil /g') do (echo Timezone:                  %%f )
endlocal

REM echo OS Name:                   >>%logfile%
REM ver |find "Version" >>%logfile%

wmic computersystem get totalphysicalmemory >>%logfile%
wmic computersystem get totalphysicalmemory

echo System Model:                >>%logfile%
echo System Model:
wmic csproduct get name | find "Name " /v >>%logfile%
wmic csproduct get name | find "Name " /v 

echo Homedrive:                 %HOMEDRIVE% >>%logfile%
echo LogonServer:               %LogonServer% >>%logfile%
echo OneDrive:                  %OneDrive% >>%logfile%

echo ======== powershell invoke-command -scr {(get-ciminstance -classname win32_logicaldisk)} >>%logfile%
powershell invoke-command -scr {(get-ciminstance -classname win32_logicaldisk)} >>%logfile%

echo ======== net share ======== >>%logfile%
powershell -Command "Get-WmiObject Win32_Share | Select Name, Path, Description, Status | Format-Table -AutoSize" >> "%logfile%"
powershell -Command "Get-WmiObject Win32_Share | Select Name, Path, Description, Status | Format-Table -AutoSize"



rem this requires administrator access
echo ======== powershell invoke-command -scr {(Get-BitLockerVolume)} >>%logfile%
powershell invoke-command -scr {(Get-BitLockerVolume)} >>%logfile%
powershell invoke-command -scr {(Get-BitLockerVolume)}

rem Backup BitLocker Recovery Key in Command Prompt
echo ======== manage-bde -protectors -get c: ======== >>%logfile%
rem  grab bitlocker recovery key for each drive (or at least c-f)
manage-bde -protectors -get c:  >>%logfile%
manage-bde -protectors -get d:  >>%logfile%
manage-bde -protectors -get e:  >>%logfile%
manage-bde -protectors -get f:  >>%logfile%


echo ======== manage-bde -status c: ======== >>%logfile%
rem get encryption status
manage-bde -status c:  >>%logfile%
manage-bde -status d:  >>%logfile%
manage-bde -status e:  >>%logfile%
manage-bde -status f:  >>%logfile%


echo ======== systeminfo ======== >>%logfile%
systeminfo >>%logfile%
systeminfo


echo ======== fsutil fsinfo ntfsinfo c: ======== >>%logfile%
fsutil fsinfo ntfsinfo c:  >>%logfile%
fsutil fsinfo ntfsinfo d:  >>%logfile%
fsutil fsinfo ntfsinfo e:  >>%logfile%
fsutil fsinfo ntfsinfo f:  >>%logfile%

echo. >>%logfile%

echo ======== ipconfig ======== >>%logfile%
ipconfig | find "IPv" | find ":" >>%logfile%


echo ======== net user ======== >>%logfile%

net user >>%logfile%

echo ==== User Account Audit (%date% %time%) ==== >> "%logfile%"

REM wmic useraccount get Disabled,FullName,Name,PasswordRequired,SID >>%logfile%
REM wmic useraccount get Disabled,FullName,Name,PasswordRequired,SID
rem net user | find "command completed" /v


for /f "skip=1 tokens=1,2,3,4,5*" %%A in ('wmic useraccount get Disabled^,FullName^,Name^,PasswordRequired^,SID') do (
    set "Disabled=%%A"
    set "FullName=%%B"
    set "Name=%%C"
    set "PasswordRequired=%%D"
    set "SID=%%E"
    echo User: !Name! ^| FullName: !FullName! ^| Disabled: !Disabled! ^| PasswordRequired: !PasswordRequired! ^| SID: !SID! >> "%logfile%"
)

echo ======== net localgroup administrators ======== >>%logfile%

net localgroup administrators | find "command completed" /v | find "Members" /v | find "Comment" /v >>%logfile%
net localgroup administrators | find "command completed" /v | find "Members" /v | find "Comment" /v


echo ======== netstat -nat ======== >>%logfile%
netstat -nat | find "TCP"  >>%logfile%

echo ======== arp -a ======== >>%logfile%
arp -a >>%logfile%


echo ======== ipconfig /displaydns ======== >>%logfile%
ipconfig /displaydns >>%logfile%
echo.

echo ======== netsh wlan show all ======== >>%logfile%
netsh wlan show all >>%logfile%

echo ======== wmic startup get Caption,Command ======== >>%logfile%
echo wmic startup get Caption,Command               >>%logfile%
REM wmic startup get Caption,Command >>%logfile%

wmic startup get Caption,Command | findstr /v /c:"Caption" >> "%logfile%"


echo.

echo ======== powershell.exe -command "Get-History" ======== >>%logfile%
rem powershell.exe -command "Get-History" >>%logfile%
powershell invoke-command -scr {(Get-History)} >>%logfile%

rem this will light up an incident response alert as "credential dumping"
rem this will skip this task if you are in the illinois domain
if %UserDomain% == ILLINOIS (echo ILLINOIS) else (reg save hklm\sam SAM_%computername%)
if %UserDomain% == ILLINOIS (echo ILLINOIS) else (reg save hklm\system SYSTEM_%computername%)



rem How well patched is the system?
rem wmic qfe get Caption,Description,HotFixID,InstalledOn

rem copy a setup file with plaintext passwords if it exists
IF EXIST C:\Windows\Panther\Unattend\Unattend.xml COPY C:\Windows\Panther\Unattend\Unattend.xml Unattend__%computername%.xml

echo ======== Done! Files are located in your working directory. ======== 
echo ======== The End ======== >> %logfile%
echo EndTime:                  >>%logfile%
powershell invoke-command -scr {(get-date -format "{MMM-dd-yyyy HH:mm:ss})} >>%logfile%


echo.
echo see %logfile% for output

echo ======== netsh wlan export profile key=clear folder=. ======== 
netsh wlan export profile key=clear folder=.
echo see WIFI-*.xml for Wifi profiles


>> wifi_passwords.txt echo WIFI_NAME: !name!
>> wifi_passwords.txt echo PASSWORD : !password!
>> wifi_passwords.txt echo -----------------------------




pause






REM # <<<<<<<<<<<<<<<<<<<<<<<<<<      Future Wishlist        >>>>>>>>>>>>>>>>>>>>>>>>>>
rem powershell to grab Chrome saved credentials
rem https://sushant747.gitbooks.io/total-oscp-guide/content/privilege_escalation_windows.html

REM # <<<<<<<<<<<<<<<<<<<<<<<<<<      Copyright        >>>>>>>>>>>>>>>>>>>>>>>>>>

REM # Copyright (C) 2025 LincolLandForensics
REM #
REM # This program is free software; you can redistribute it and/or modify it under
REM # the terms of the GNU General Public License version 2, as published by the
REM # Free Software Foundation
REM #
REM # This program is distributed in the hope that it will be useful, but WITHOUT
REM # ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
REM # FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
REM # details (http://www.gnu.org/licenses/gpl.txt).
