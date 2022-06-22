@echo off

rem https://blog.digital-forensics.it/2021/03/triaging-modern-android-devices-aka.html
rem install android sdk
rem this script has to be run from the Android SDK folder with adb.exe in it.
rem cd C:\Users\<user>\AppData\Local\Android\Sdk\platform-tools

rem version 1.0.2

echo ==== TRIAGE INFORMATION ==== >TRIAGE_.txt
echo.
echo ==== adb shell pm list users ==== Phone Number >>TRIAGE_.txt
adb shell pm list users |find "UserInfo" >>TRIAGE_.txt

rem setlocal
rem for /f %%f in ('adb shell getprop ro.product.manufacturer') do (echo Manufacturer:                  %%f >>TRIAGE_.txt
rem endlocal

echo ====  adb shell getprop ro.product.manufacturer  ==== Device Manufacturer  >>TRIAGE_.txt
adb shell getprop ro.product.manufacturer >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell getprop ro.product.model  ==== Device Product  >>TRIAGE_.txt
adb shell getprop ro.product.model >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell getprop ril.serialnumber  ==== Android Serial Number  >>TRIAGE_.txt
adb shell getprop ril.serialnumber >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell getprop persist.sys.timezone  ==== Timezone  >>TRIAGE_.txt
adb shell getprop persist.sys.timezone >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell settings get secure android_id  ==== Android ID  >>TRIAGE_.txt
adb shell settings get secure android_id >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell settings get secure bluetooth_name  ==== Bluetooth Name  >>TRIAGE_.txt
adb shell settings get secure bluetooth_name >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell settings get secure bluetooth_address  ==== Bluetooth Address  >>TRIAGE_.txt
adb shell settings get secure bluetooth_address >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell getprop ro.build.version.release  ==== Android Version  >>TRIAGE_.txt
adb shell getprop ro.build.version.release >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys clipboard ==== >>TRIAGE_.txt
adb shell dumpsys clipboard >>TRIAGE_.txt
echo.
echo ====  adb shell getprop ro.build.fingerprint  ==== Android Fingerprint  >>TRIAGE_.txt
adb shell getprop ro.build.fingerprint >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell getprop ro.build.date  ==== Build date  >>TRIAGE_.txt
adb shell getprop ro.build.date >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell getprop persist.sys.usb.config  ==== USB Configuration  >>TRIAGE_.txt
adb shell getprop persist.sys.usb.config >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ====  adb shell getprop storage.mmc.size  ==== Storage Size  >>TRIAGE_.txt
adb shell getprop storage.mmc.size >>TRIAGE_.txt
echo. >>TRIAGE_.txt

echo ==== adb shell content query --uri content://com.android.contacts/contacts ==== Contacts>>TRIAGE_.txt
adb shell content query --uri content://com.android.contacts/contacts
echo.

rem echo ====  adb shell date  ==== Device Time  >>TRIAGE_.txt
rem adb shell date >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt

rem echo ==== adb shell id ==== >>TRIAGE_.txt
rem adb shell id >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt
echo ==== adb shell cat /proc/version ==== >>TRIAGE_.txt
adb shell cat /proc/version >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell df ==== >>TRIAGE_.txt
adb shell df >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell ip address show wlan0 ==== >>TRIAGE_.txt
adb shell ip address show wlan0 >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell ifconfig -a ==== >>TRIAGE_.txt
adb shell ifconfig -a >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell netstat -an ==== >>TRIAGE_.txt
adb shell netstat -an >>TRIAGE_.txt
echo. >>TRIAGE_.txt

echo ==== adb shell service list ==== >>TRIAGE_.txt
adb shell service list >>TRIAGE_.txt
echo. >>TRIAGE_.txt

echo ==== adb shell dumpsys notification --noredact |find "String" ====  >>TRIAGE_.txt
echo adb shell dumpsys notification --noredact |find "String" >>TRIAGE_.txt
adb shell dumpsys notification --noredact |find "String" >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys account --noredact ====  >>TRIAGE_.txt
adb shell dumpsys account --noredact >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys input ==== >>TRIAGE_.txt
adb shell dumpsys input >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys procstats ==== >>TRIAGE_.txt
adb shell dumpsys procstats >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys appops ==== >>TRIAGE_.txt
adb shell dumpsys appops >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys batterystats ==== >>TRIAGE_.txt
adb shell dumpsys batterystats >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys bluetooth_manager ==== >>TRIAGE_.txt
adb shell dumpsys bluetooth_manager >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys cpuinfo ==== >>TRIAGE_.txt
adb shell dumpsys cpuinfo >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys dbinfo -v ==== >>TRIAGE_.txt
adb shell dumpsys dbinfo -v >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys diskstats ==== >>TRIAGE_.txt
adb shell dumpsys diskstats >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys package ==== >>TRIAGE_.txt
rem adb shell dumpsys package |find "dumpDir" >>TRIAGE_.txt
rem adb shell dumpsys package |find "codePath" >>TRIAGE_.txt
rem adb shell dumpsys package |find "install permissions" >>TRIAGE_.txt
adb shell dumpsys package >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys usagestats ==== >>TRIAGE_.txt
adb shell dumpsys usagestats >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys vibrator ==== >>TRIAGE_.txt
adb shell dumpsys vibrator >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell dumpsys wifi ==== >>TRIAGE_.txt
adb shell dumpsys wifi >>TRIAGE_.txt
rem find " SSID:"
rem find "creation time"

rem echo ==== adb bugreport ==== >>TRIAGE_.txt
rem adb bugreport >>TRIAGE_.txt

echo ==== adb shell pm list packages -f ==== prints all packages including their associated APKs >>TRIAGE_.txt
adb shell pm list packages -f >>TRIAGE_.txt
echo. >>TRIAGE_.txt
echo ==== adb shell pm list permissions -f ==== prints all permissions including all related information >>TRIAGE_.txt
adb shell pm list permissions -f >>TRIAGE_.txt
echo. >>TRIAGE_.txt

REM echo ====  adb shell pm get-max-users  ==== >>TRIAGE_.txt
REM adb shell pm get-max-users >>TRIAGE_.txt
REM echo.
REM echo ====  adb shell pm list users  ==== >>TRIAGE_.txt
REM adb shell pm list users >>TRIAGE_.txt
REM echo.
echo ====  adb shell pm list features  ==== >>TRIAGE_.txt
adb shell pm list features >>TRIAGE_.txt
echo.
REM echo ====  adb shell pm list instrumentation  ==== >>TRIAGE_.txt
REM adb shell pm list instrumentation >>TRIAGE_.txt
REM echo.
echo ====  adb shell pm list libraries -f  ==== >>TRIAGE_.txt
adb shell pm list libraries -f >>TRIAGE_.txt
echo.
echo ====  adb shell pm list packages -f  ==== >>TRIAGE_.txt
adb shell pm list packages -f >>TRIAGE_.txt
echo.
echo ====  adb shell pm list packages -f -u  ==== >>TRIAGE_.txt
adb shell pm list packages -f -u >>TRIAGE_.txt
echo.
echo ====  adb shell pm list permissions -f  ==== >>TRIAGE_.txt
adb shell pm list permissions -f >>TRIAGE_.txt
echo.
echo ====  adb shell cat /data/system/uiderrors.txt  ==== >>TRIAGE_.txt
adb shell cat /data/system/uiderrors.txt >>TRIAGE_.txt
echo.

REM echo ==== adb bugreport ==== >>TRIAGE_.txt
REM adb bugreport >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys  ==== >>TRIAGE_.txt
REM adb shell dumpsys  >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys account ==== >>TRIAGE_.txt
REM adb shell dumpsys account >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys activity ==== >>TRIAGE_.txt
REM adb shell dumpsys activity >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys alarm ==== >>TRIAGE_.txt
REM adb shell dumpsys alarm >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys appops ==== >>TRIAGE_.txt
REM adb shell dumpsys appops >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys audio ==== >>TRIAGE_.txt
REM adb shell dumpsys audio >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys autofill ==== >>TRIAGE_.txt
REM adb shell dumpsys autofill >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys backup ==== >>TRIAGE_.txt
REM adb shell dumpsys backup >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys battery ==== >>TRIAGE_.txt
REM adb shell dumpsys battery >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys batteryproperties ==== >>TRIAGE_.txt
REM adb shell dumpsys batteryproperties >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys batterystats ==== >>TRIAGE_.txt
REM adb shell dumpsys batterystats >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys carrier_config ==== >>TRIAGE_.txt
REM adb shell dumpsys carrier_config >>TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys connectivity ==== >>TRIAGE_.txt
adb shell dumpsys connectivity >>TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys content ==== >>TRIAGE_.txt
REM adb shell dumpsys content >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys cpuinfo ==== >>TRIAGE_.txt
REM adb shell dumpsys cpuinfo >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys dbinfo ==== >>TRIAGE_.txt
REM adb shell dumpsys dbinfo >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys device_policy ==== >>TRIAGE_.txt
REM adb shell dumpsys device_policy >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys devicestoragemonitor ==== >>TRIAGE_.txt
REM adb shell dumpsys devicestoragemonitor >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys display ==== >>TRIAGE_.txt
REM adb shell dumpsys display >>TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys dropbox ==== >>TRIAGE_.txt
adb shell dumpsys dropbox >>TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys gfxinfo ==== >>TRIAGE_.txt
REM adb shell dumpsys gfxinfo >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys iphonesubinfo ==== >>TRIAGE_.txt
REM adb shell dumpsys iphonesubinfo >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys jobscheduler ==== >>TRIAGE_.txt
REM adb shell dumpsys jobscheduler >>TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys location  --noredact ==== >>TRIAGE_.txt
adb shell dumpsys location --noredact >>TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys meminfo -a ==== >>TRIAGE_.txt
REM adb shell dumpsys meminfo -a >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys mount ==== >>TRIAGE_.txt
REM adb shell dumpsys mount >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys netpolicy ==== >>TRIAGE_.txt
REM adb shell dumpsys netpolicy >>TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys netstats |find "networkId=" ==== >>TRIAGE_.txt
adb shell dumpsys netstats |find "networkId=" >>TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys network_management ==== >>TRIAGE_.txt
REM adb shell dumpsys network_management >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys network_score ==== >>TRIAGE_.txt
REM adb shell dumpsys network_score >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys phone ==== >>TRIAGE_.txt
REM adb shell dumpsys phone >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys procstats --full-details ==== >>TRIAGE_.txt
REM adb shell dumpsys procstats --full-details >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys restriction_policy ==== >>TRIAGE_.txt
REM adb shell dumpsys restriction_policy >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys sdhms ==== >>TRIAGE_.txt
REM adb shell dumpsys sdhms >>TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys sec_location ==== >>TRIAGE_.txt
adb shell dumpsys sec_location >>TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys secims ==== >>TRIAGE_.txt
REM adb shell dumpsys secims >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys search ==== >>TRIAGE_.txt
REM adb shell dumpsys search >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys sensorservice ==== >>TRIAGE_.txt
REM adb shell dumpsys sensorservice >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys settings ==== >>TRIAGE_.txt
REM adb shell dumpsys settings >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys shortcut ==== >>TRIAGE_.txt
REM adb shell dumpsys shortcut >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys stats ==== >>TRIAGE_.txt
REM adb shell dumpsys stats >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys statusbar ==== >>TRIAGE_.txt
REM adb shell dumpsys statusbar >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys storaged ==== >>TRIAGE_.txt
REM adb shell dumpsys storaged >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys telecom ==== >>TRIAGE_.txt
REM adb shell dumpsys telecom >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys usagestats ==== >>TRIAGE_.txt
REM adb shell dumpsys usagestats >>TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys user ==== Last logged in >>TRIAGE_.txt
adb shell dumpsys user >>TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys usb ==== >>TRIAGE_.txt
REM adb shell dumpsys usb >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys wallpaper ==== >>TRIAGE_.txt
REM adb shell dumpsys wallpaper >>TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys window ==== >>TRIAGE_.txt
REM adb shell dumpsys window >>TRIAGE_.txt
REM echo.



rem echo ====  adb shell getprop ro.product.device  ==== Product Device  >>TRIAGE_.txt
rem adb shell getprop ro.product.device >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt
rem echo ====  adb shell getprop ro.product.name  ==== Product Name  >>TRIAGE_.txt
rem adb shell getprop ro.product.name >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt
rem echo ====  adb shell getprop ro.chipname  ==== Chipname  >>TRIAGE_.txt
rem adb shell getprop ro.chipname >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt
rem echo ==== adb backup -all -shared -system -keyvalue -apk -f backup.ab ==== >>TRIAGE_.txt
rem adb backup -all -shared -system -keyvalue -apk -f backup.ab

rem echo ====  adb shell ro.build.id  ==== Build ID  >>TRIAGE_.txt
rem adb shell ro.build.id >>TRIAGE_.txt

rem echo ====  adb shell ro.boot.bootloader  ==== Bootloader Version  >>TRIAGE_.txt
rem adb shell ro.boot.bootloader >>TRIAGE_.txt

rem echo ====  adb shell ro.build.version.security_patch  ==== Security Patch  >>TRIAGE_.txt
rem adb shell ro.build.version.security_patch >>TRIAGE_.txt

rem echo ====  adb shell getprop ro.product.code  ==== Product Code  >>TRIAGE_.txt
rem adb shell getprop ro.product.code >>TRIAGE_.txt

rem echo ==== adb shell uname -a ==== >>TRIAGE_.txt
rem adb shell uname -a >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell uptime ==== >>TRIAGE_.txt
rem adb shell uptime >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell printenv ==== >>TRIAGE_.txt
rem adb shell printenv >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/partitions ==== >>TRIAGE_.txt
rem adb shell cat /proc/partitions >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/cpuinfo ==== >>TRIAGE_.txt
rem adb shell cat /proc/cpuinfo >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/diskstats ==== >>TRIAGE_.txt
rem adb shell cat /proc/diskstats >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell df -ah ==== >>TRIAGE_.txt
rem adb shell df -ah >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell mount ==== >>TRIAGE_.txt
rem adb shell mount >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell lsof ==== >>TRIAGE_.txt
rem adb shell lsof >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell ps -ef ==== >>TRIAGE_.txt
rem adb shell ps -ef >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell top -n 1 ==== >>TRIAGE_.txt
rem adb shell top -n 1 >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/sched_debug ==== >>TRIAGE_.txt
rem adb shell cat /proc/sched_debug >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell vmstat ==== >>TRIAGE_.txt
rem adb shell vmstat >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell sysctl -a ==== >>TRIAGE_.txt
rem adb shell sysctl -a >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell ime list ==== >>TRIAGE_.txt
rem adb shell ime list >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell logcat -S -b all ==== >>TRIAGE_.txt
rem adb shell logcat -S -b all >>TRIAGE_.txt
rem echo.
rem echo ==== adb shell logcat -d -b all V:* ==== >>TRIAGE_.txt
rem adb shell logcat -d -b all V:* >>TRIAGE_.txt
rem echo.
rem echo ====  adb shell getprop ro.config.notification_sound  ==== Notification Sound  >>TRIAGE_.txt
rem adb shell getprop ro.config.notification_sound >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt
rem echo ====  adb shell getprop ro.config.alarm_alert  ==== Alerm Alert  >>TRIAGE_.txt
rem adb shell getprop ro.config.alarm_alert >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt
rem echo ====  adb shell getprop ro.config.ringtone  ==== Ringtone  >>TRIAGE_.txt
rem adb shell getprop ro.config.ringtone >>TRIAGE_.txt
rem echo. >>TRIAGE_.txt
rem echo ====  adb shell getprop rro.config.media_sound  ==== Media Sound  >>TRIAGE_.txt
rem adb shell getprop rro.config.media_sound >>TRIAGE_.txt
rem echo.

echo collection saved in TRIAGE_.txt