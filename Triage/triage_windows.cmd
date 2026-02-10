@echo off
setlocal enabledelayedexpansion

rem version 1.4.0 - Refactored for PowerShell consistency and WMIC deprecation

REM Check for admin rights
net session >nul 2>&1
if %errorlevel% == 0 (
    set adminlevel=admin
) else (
    set adminlevel=user
)

echo You are running as: !adminlevel!

set "logfile=_TRIAGE__%computername%.txt" 

echo ======== TRIAGE INFORMATION ======== >%logfile%
echo ======== TRIAGE INFORMATION ========

REM Extract date and time
for /f "tokens=2-4 delims=/ " %%a in ("%date%") do (
    set "mm=%%a"
    set "dd=%%b"
    set "yyyy=%%c"
)


set "hh=%time:~0,2%"
set "mn=%time:~3,2%"
set "ss=%time:~6,2%"
if "%hh:~0,1%"==" " set "hh=0%hh:~1,1%"
set "StartTime=%yyyy%-%mm%-%dd% %hh%:%mm%:%ss%"
echo StartTime:                 %StartTime% >>%logfile%

REM powershell -Command "& { Get-Date -Format 'yyyy-MM-dd HH:mm:ss' }" >> %logfile%
echo. >>%logfile%

REM System Info
echo Host Name:                 %computername% >>%logfile%
echo Host Name:                 %computername%
echo UserDomain:                %UserDomain% >>%logfile%
echo Currentuser:               %username% >>%logfile%
echo Currentuser:               %username%
echo Email:                     %email% >>%logfile%
echo Email:                     %email%
for /f "tokens=*" %%f in ('tzutil /g') do (
    echo Timezone:                  %%f >>%logfile%
)

REM Memory and Model
echo Memory: >> %logfile%
powershell -Command "& { (Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory }" >> %logfile%
echo Model: >> %logfile%
powershell -Command "& { (Get-CimInstance -ClassName Win32_ComputerSystem).Model }" >> %logfile%

REM Legacy fallback (commented)
REM IF EXIST wmic.exe (
REM     wmic computersystem get totalphysicalmemory >>%logfile%
REM     wmic csproduct get name | find "Name " /v >>%logfile%
REM ) ELSE (
REM     echo WMIC does not exist. >>%logfile%
REM )

echo Homedrive:                 %HOMEDRIVE% >>%logfile%
echo LogonServer:               %LogonServer% >>%logfile%
echo OneDrive:                  %OneDrive% >>%logfile%

REM Logical Disks
echo Logical Disks: >>%logfile%
if "!adminlevel!"=="admin" (
    powershell -Command "& { Get-CimInstance -ClassName Win32_LogicalDisk }" >> %logfile%
) else (
    echo Skipped LogicalDisk (not admin) >>%logfile%
)

REM Network Shares
echo Network Shares: >>%logfile% 
powershell -Command "& { Get-WmiObject Win32_Share | Select Name, Path, Description, Status | Format-Table -AutoSize }" >> %logfile%

REM BitLocker Info
echo BitLocker Info Protectors: (manage-bde -protectors) >>%logfile% 
powershell -Command "& { Get-BitLockerVolume }" >> %logfile%

for /f "tokens=*" %%d in ('powershell -Command "Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID"') do (
    echo BitLocker Protectors for %%d >>%logfile%
    manage-bde -protectors -get %%d >>%logfile%
)

echo BitLocker Info:
echo BitLocker Info Status: (manage-bde -status) >>%logfile%
for /f "tokens=*" %%d in ('powershell -Command "Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID"') do (
    echo Status for %%d >>%logfile%
    manage-bde -status %%d >>%logfile%
    manage-bde -status %%d
)

REM System Info
echo System Info: >>%logfile% 
systeminfo >>%logfile%


REM NTFS Info
REM fsutil fsinfo ntfsinfo c: >>%logfile%
REM fsutil fsinfo ntfsinfo d: >>%logfile%
REM fsutil fsinfo ntfsinfo e: >>%logfile%

echo NTFS Info: >>%logfile%
for /f "tokens=*" %%d in ('powershell -Command "Get-CimInstance -ClassName Win32_LogicalDisk | Select-Object -ExpandProperty DeviceID"') do (
    echo NTFS Info for %%d >>%logfile%
    fsutil fsinfo ntfsinfo %%d >>%logfile%
)


REM IP and Network
echo IP and Network: (ipconfig | find "IPv" | find ":") >>%logfile%
echo ipconfig | find "IPv" | find ":" >>%logfile%
ipconfig | find "IPv" | find ":" >>%logfile%

echo Users: (net user) >>%logfile%
net user >>%logfile%

REM User Audit
powershell -Command "& { Get-CimInstance Win32_UserAccount | Select Name, FullName, Disabled, PasswordRequired, SID | ForEach-Object { Write-Output ('User: {0} | FullName: {1} | Disabled: {2} | PasswordRequired: {3} | SID: {4}' -f $_.Name, $_.FullName, $_.Disabled, $_.PasswordRequired, $_.SID) } }" >> %logfile%


REM Legacy fallback
REM IF EXIST wmic.exe (
REM     wmic useraccount get Disabled,FullName,Name,PasswordRequired,SID >>%logfile%
REM )

REM Admin Group
REM echo Admin Group:
echo Admin Group: (net localgroup administrators) >>%logfile%
net localgroup administrators | find "command completed" /v | find "Members" /v | find "Comment" /v >>%logfile%

REM Netstat, ARP, DNS
echo Netstat: (netstat -nat | find "TCP") >>%logfile%
netstat -nat | find "TCP" >>%logfile%

echo ARP: (arp -a) >>%logfile%
arp -a >>%logfile%

echo DNS: (ipconfig /displaydns) >>%logfile%
ipconfig /displaydns >>%logfile%

REM WLAN Info
echo WLAN Info:
echo WLAN Info: (netsh wlan show all) >>%logfile%
netsh wlan show all >>%logfile%

REM Startup Programs
echo Startup Programs: (powershell -Command "& { Get-CimInstance Win32_StartupCommand | Select Caption, Command }") >>%logfile%
powershell -Command "& { Get-CimInstance Win32_StartupCommand | Select Caption, Command }" >> %logfile%

REM Legacy fallback
REM IF EXIST wmic.exe (
REM     wmic startup get Caption,Command | findstr /v /c:"Caption" >> %logfile%
REM )

REM PowerShell History
echo PowerShell History: (powershell -Command "& { Get-History }") >>%logfile%
powershell -Command "& { Get-History }" >> %logfile%

REM Registry Dump (if not ILLINOIS domain)
if %UserDomain% == ILLINOIS (
    echo Skipped registry dump for ILLINOIS domain
) else (
    reg save hklm\sam SAM_%computername% /y
    reg save hklm\system SYSTEM_%computername% /y
)

REM Unattended Setup File
IF EXIST C:\Windows\Panther\Unattend\Unattend.xml (
    COPY C:\Windows\Panther\Unattend\Unattend.xml Unattend__%computername%.xml
)

echo ======== netsh wlan export profile key=clear folder=. ======== >> %logfile%
netsh wlan export profile key=clear folder=.

for /f "tokens=2 delims=:" %A in ('netsh wlan show profiles ^| findstr "All User Profile"') do @echo Profile:%A & for /f "tokens=2 delims=:" %B in ('netsh wlan show profile name=%A key=clear ^| findstr "Key Content"') do @echo Password:%B & echo. >> %logfile%

REM Final Timestamp
echo EndTime: >>%logfile%
powershell -Command "& { Get-Date -Format 'yyyy-MM-dd HH:mm:ss' }" >> %logfile%

echo ======== Done! Files are located in your working directory. ========
echo see %logfile% for output

pause