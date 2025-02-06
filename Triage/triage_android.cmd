@echo off

rem https://blog.digital-forensics.it/2021/03/triaging-modern-android-devices-aka.html
rem install android sdk
rem this script has to be run from the Android SDK folder with adb.exe in it.
rem cd C:\Users\<user>\AppData\Local\Android\Sdk\platform-tools

rem version 1.0.5

echo ==== TRIAGE INFORMATION ==== >_TRIAGE_.txt
echo.
echo ==== adb shell pm list users ==== Phone Number >>_TRIAGE_.txt
adb shell pm list users |find "UserInfo" >>_TRIAGE_.txt

rem setlocal
rem for /f %%f in ('adb shell getprop ro.product.manufacturer') do (echo Manufacturer:                  %%f >>_TRIAGE_.txt
rem endlocal

echo ====  adb shell getprop ro.product.manufacturer  ==== Device Manufacturer  >>_TRIAGE_.txt
adb shell getprop ro.product.manufacturer >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell getprop ro.product.model  ==== Device Product  >>_TRIAGE_.txt
adb shell getprop ro.product.model >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell getprop ril.serialnumber  ==== Android Serial Number  >>_TRIAGE_.txt
adb shell getprop ril.serialnumber >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell getprop persist.sys.timezone  ==== Timezone  >>_TRIAGE_.txt
adb shell getprop persist.sys.timezone >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell settings get secure android_id  ==== Android ID  >>_TRIAGE_.txt
adb shell settings get secure android_id >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell getprop persist.sys.usb.config  ==== USB Configuration  >>_TRIAGE_.txt
adb shell getprop persist.sys.usb.config >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys user ==== Last logged in >>_TRIAGE_.txt
adb shell dumpsys user >>_TRIAGE_.txt
echo.
echo ====  adb shell getprop ro.build.version.release  ==== Android Version  >>_TRIAGE_.txt
adb shell getprop ro.build.version.release >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell settings get secure bluetooth_name  ==== Bluetooth Name  >>_TRIAGE_.txt
adb shell settings get secure bluetooth_name >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell settings get secure bluetooth_address  ==== Bluetooth Address  >>_TRIAGE_.txt
adb shell settings get secure bluetooth_address >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt

echo ==== adb shell dumpsys clipboard ==== >>_TRIAGE_.txt
adb shell dumpsys clipboard >>_TRIAGE_.txt
echo.
echo ====  adb shell getprop ro.build.fingerprint  ==== Android Fingerprint  >>_TRIAGE_.txt
adb shell getprop ro.build.fingerprint >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ====  adb shell getprop ro.build.date  ==== Build date  >>_TRIAGE_.txt
adb shell getprop ro.build.date >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt

echo ====  adb shell getprop storage.mmc.size  ==== Storage Size  >>_TRIAGE_.txt
adb shell getprop storage.mmc.size >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt

echo ==== adb shell content query --uri content://com.android.contacts/contacts ==== Contacts>>_TRIAGE_.txt
adb shell content query --uri content://com.android.contacts/contacts
echo.

rem echo ====  adb shell date  ==== Device Time  >>_TRIAGE_.txt
rem adb shell date >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt

rem echo ==== adb shell id ==== >>_TRIAGE_.txt
rem adb shell id >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt
echo ==== adb shell cat /proc/version ==== >>_TRIAGE_.txt
adb shell cat /proc/version >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell df ==== >>_TRIAGE_.txt
adb shell df >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell ip address show wlan0 ==== >>_TRIAGE_.txt
adb shell ip address show wlan0 >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell ifconfig -a ==== >>_TRIAGE_.txt
adb shell ifconfig -a >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell netstat -an ==== >>_TRIAGE_.txt
adb shell netstat -an >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt

echo ==== adb shell service list ==== >>_TRIAGE_.txt
adb shell service list >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt

echo ==== adb shell dumpsys notification --noredact |find "String" ====  >>_TRIAGE_.txt
echo adb shell dumpsys notification --noredact |find "String" >>_TRIAGE_.txt
adb shell dumpsys notification --noredact |find "String" >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys account --noredact ====  >>_TRIAGE_.txt
adb shell dumpsys account --noredact >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys input ==== >>_TRIAGE_.txt
adb shell dumpsys input >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys procstats ==== >>_TRIAGE_.txt
adb shell dumpsys procstats >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys appops ==== >>_TRIAGE_.txt
adb shell dumpsys appops >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys batterystats ==== >>_TRIAGE_.txt
adb shell dumpsys batterystats >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys bluetooth_manager ==== >>_TRIAGE_.txt
adb shell dumpsys bluetooth_manager >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys cpuinfo ==== >>_TRIAGE_.txt
adb shell dumpsys cpuinfo >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys dbinfo -v ==== >>_TRIAGE_.txt
adb shell dumpsys dbinfo -v >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys diskstats ==== >>_TRIAGE_.txt
adb shell dumpsys diskstats >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys package ==== >>_TRIAGE_.txt
rem adb shell dumpsys package |find "dumpDir" >>_TRIAGE_.txt
rem adb shell dumpsys package |find "codePath" >>_TRIAGE_.txt
rem adb shell dumpsys package |find "install permissions" >>_TRIAGE_.txt
adb shell dumpsys package >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys usagestats ==== >>_TRIAGE_.txt
adb shell dumpsys usagestats >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys vibrator ==== >>_TRIAGE_.txt
adb shell dumpsys vibrator >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell dumpsys wifi ==== >>_TRIAGE_.txt
adb shell dumpsys wifi >>_TRIAGE_.txt
rem find " SSID:"
rem find "creation time"

rem echo ==== adb bugreport ==== >>_TRIAGE_.txt
rem adb bugreport >>_TRIAGE_.txt

echo ==== adb shell pm list packages -f ==== prints all packages including their associated APKs >>_TRIAGE_.txt
adb shell pm list packages -f >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt
echo ==== adb shell pm list permissions -f ==== prints all permissions including all related information >>_TRIAGE_.txt
adb shell pm list permissions -f >>_TRIAGE_.txt
echo. >>_TRIAGE_.txt

REM echo ====  adb shell pm get-max-users  ==== >>_TRIAGE_.txt
REM adb shell pm get-max-users >>_TRIAGE_.txt
REM echo.
REM echo ====  adb shell pm list users  ==== >>_TRIAGE_.txt
REM adb shell pm list users >>_TRIAGE_.txt
REM echo.
echo ====  adb shell pm list features  ==== >>_TRIAGE_.txt
adb shell pm list features >>_TRIAGE_.txt
echo.
REM echo ====  adb shell pm list instrumentation  ==== >>_TRIAGE_.txt
REM adb shell pm list instrumentation >>_TRIAGE_.txt
REM echo.
echo ====  adb shell pm list libraries -f  ==== >>_TRIAGE_.txt
adb shell pm list libraries -f >>_TRIAGE_.txt
echo.
echo ====  adb shell pm list packages -f  ==== >>_TRIAGE_.txt
adb shell pm list packages -f >>_TRIAGE_.txt
echo.
echo ====  adb shell pm list packages -f -u  ==== >>_TRIAGE_.txt
adb shell pm list packages -f -u >>_TRIAGE_.txt
echo.
echo ====  adb shell pm list permissions -f  ==== >>_TRIAGE_.txt
adb shell pm list permissions -f >>_TRIAGE_.txt
echo.
echo ====  adb shell cat /data/system/uiderrors.txt  ==== >>_TRIAGE_.txt
adb shell cat /data/system/uiderrors.txt >>_TRIAGE_.txt
echo.

REM echo ==== adb bugreport ==== >>_TRIAGE_.txt
REM adb bugreport >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys  ==== >>_TRIAGE_.txt
REM adb shell dumpsys  >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys account ==== >>_TRIAGE_.txt
REM adb shell dumpsys account >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys activity ==== >>_TRIAGE_.txt
REM adb shell dumpsys activity >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys alarm ==== >>_TRIAGE_.txt
REM adb shell dumpsys alarm >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys appops ==== >>_TRIAGE_.txt
REM adb shell dumpsys appops >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys audio ==== >>_TRIAGE_.txt
REM adb shell dumpsys audio >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys autofill ==== >>_TRIAGE_.txt
REM adb shell dumpsys autofill >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys backup ==== >>_TRIAGE_.txt
REM adb shell dumpsys backup >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys battery ==== >>_TRIAGE_.txt
REM adb shell dumpsys battery >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys batteryproperties ==== >>_TRIAGE_.txt
REM adb shell dumpsys batteryproperties >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys batterystats ==== >>_TRIAGE_.txt
REM adb shell dumpsys batterystats >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys carrier_config ==== >>_TRIAGE_.txt
REM adb shell dumpsys carrier_config >>_TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys connectivity ==== >>_TRIAGE_.txt
adb shell dumpsys connectivity >>_TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys content ==== >>_TRIAGE_.txt
REM adb shell dumpsys content >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys cpuinfo ==== >>_TRIAGE_.txt
REM adb shell dumpsys cpuinfo >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys dbinfo ==== >>_TRIAGE_.txt
REM adb shell dumpsys dbinfo >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys device_policy ==== >>_TRIAGE_.txt
REM adb shell dumpsys device_policy >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys devicestoragemonitor ==== >>_TRIAGE_.txt
REM adb shell dumpsys devicestoragemonitor >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys display ==== >>_TRIAGE_.txt
REM adb shell dumpsys display >>_TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys dropbox ==== >>_TRIAGE_.txt
adb shell dumpsys dropbox >>_TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys gfxinfo ==== >>_TRIAGE_.txt
REM adb shell dumpsys gfxinfo >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys iphonesubinfo ==== >>_TRIAGE_.txt
REM adb shell dumpsys iphonesubinfo >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys jobscheduler ==== >>_TRIAGE_.txt
REM adb shell dumpsys jobscheduler >>_TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys location  --noredact ==== >>_TRIAGE_.txt
adb shell dumpsys location --noredact >>_TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys meminfo -a ==== >>_TRIAGE_.txt
REM adb shell dumpsys meminfo -a >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys mount ==== >>_TRIAGE_.txt
REM adb shell dumpsys mount >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys netpolicy ==== >>_TRIAGE_.txt
REM adb shell dumpsys netpolicy >>_TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys netstats |find "networkId=" ==== >>_TRIAGE_.txt
adb shell dumpsys netstats |find "networkId=" >>_TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys network_management ==== >>_TRIAGE_.txt
REM adb shell dumpsys network_management >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys network_score ==== >>_TRIAGE_.txt
REM adb shell dumpsys network_score >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys phone ==== >>_TRIAGE_.txt
REM adb shell dumpsys phone >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys procstats --full-details ==== >>_TRIAGE_.txt
REM adb shell dumpsys procstats --full-details >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys restriction_policy ==== >>_TRIAGE_.txt
REM adb shell dumpsys restriction_policy >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys sdhms ==== >>_TRIAGE_.txt
REM adb shell dumpsys sdhms >>_TRIAGE_.txt
REM echo.
echo ==== adb shell dumpsys sec_location ==== >>_TRIAGE_.txt
adb shell dumpsys sec_location >>_TRIAGE_.txt
echo.
REM echo ==== adb shell dumpsys secims ==== >>_TRIAGE_.txt
REM adb shell dumpsys secims >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys search ==== >>_TRIAGE_.txt
REM adb shell dumpsys search >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys sensorservice ==== >>_TRIAGE_.txt
REM adb shell dumpsys sensorservice >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys settings ==== >>_TRIAGE_.txt
REM adb shell dumpsys settings >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys shortcut ==== >>_TRIAGE_.txt
REM adb shell dumpsys shortcut >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys stats ==== >>_TRIAGE_.txt
REM adb shell dumpsys stats >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys statusbar ==== >>_TRIAGE_.txt
REM adb shell dumpsys statusbar >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys storaged ==== >>_TRIAGE_.txt
REM adb shell dumpsys storaged >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys telecom ==== >>_TRIAGE_.txt
REM adb shell dumpsys telecom >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys usagestats ==== >>_TRIAGE_.txt
REM adb shell dumpsys usagestats >>_TRIAGE_.txt
REM echo.

REM echo ==== adb shell dumpsys usb ==== >>_TRIAGE_.txt
REM adb shell dumpsys usb >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys wallpaper ==== >>_TRIAGE_.txt
REM adb shell dumpsys wallpaper >>_TRIAGE_.txt
REM echo.
REM echo ==== adb shell dumpsys window ==== >>_TRIAGE_.txt
REM adb shell dumpsys window >>_TRIAGE_.txt
REM echo.



rem echo ====  adb shell getprop ro.product.device  ==== Product Device  >>_TRIAGE_.txt
rem adb shell getprop ro.product.device >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt
rem echo ====  adb shell getprop ro.product.name  ==== Product Name  >>_TRIAGE_.txt
rem adb shell getprop ro.product.name >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt
rem echo ====  adb shell getprop ro.chipname  ==== Chipname  >>_TRIAGE_.txt
rem adb shell getprop ro.chipname >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt
rem echo ==== adb backup -all -shared -system -keyvalue -apk -f backup.ab ==== >>_TRIAGE_.txt
rem adb backup -all -shared -system -keyvalue -apk -f backup.ab

rem echo ====  adb shell ro.build.id  ==== Build ID  >>_TRIAGE_.txt
rem adb shell ro.build.id >>_TRIAGE_.txt

rem echo ====  adb shell ro.boot.bootloader  ==== Bootloader Version  >>_TRIAGE_.txt
rem adb shell ro.boot.bootloader >>_TRIAGE_.txt

rem echo ====  adb shell ro.build.version.security_patch  ==== Security Patch  >>_TRIAGE_.txt
rem adb shell ro.build.version.security_patch >>_TRIAGE_.txt

rem echo ====  adb shell getprop ro.product.code  ==== Product Code  >>_TRIAGE_.txt
rem adb shell getprop ro.product.code >>_TRIAGE_.txt

rem echo ==== adb shell uname -a ==== >>_TRIAGE_.txt
rem adb shell uname -a >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell uptime ==== >>_TRIAGE_.txt
rem adb shell uptime >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell printenv ==== >>_TRIAGE_.txt
rem adb shell printenv >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/partitions ==== >>_TRIAGE_.txt
rem adb shell cat /proc/partitions >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/cpuinfo ==== >>_TRIAGE_.txt
rem adb shell cat /proc/cpuinfo >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/diskstats ==== >>_TRIAGE_.txt
rem adb shell cat /proc/diskstats >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell df -ah ==== >>_TRIAGE_.txt
rem adb shell df -ah >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell mount ==== >>_TRIAGE_.txt
rem adb shell mount >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell lsof ==== >>_TRIAGE_.txt
rem adb shell lsof >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell ps -ef ==== >>_TRIAGE_.txt
rem adb shell ps -ef >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell top -n 1 ==== >>_TRIAGE_.txt
rem adb shell top -n 1 >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell cat /proc/sched_debug ==== >>_TRIAGE_.txt
rem adb shell cat /proc/sched_debug >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell vmstat ==== >>_TRIAGE_.txt
rem adb shell vmstat >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell sysctl -a ==== >>_TRIAGE_.txt
rem adb shell sysctl -a >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell ime list ==== >>_TRIAGE_.txt
rem adb shell ime list >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell logcat -S -b all ==== >>_TRIAGE_.txt
rem adb shell logcat -S -b all >>_TRIAGE_.txt
rem echo.
rem echo ==== adb shell logcat -d -b all V:* ==== >>_TRIAGE_.txt
rem adb shell logcat -d -b all V:* >>_TRIAGE_.txt
rem echo.
rem echo ====  adb shell getprop ro.config.notification_sound  ==== Notification Sound  >>_TRIAGE_.txt
rem adb shell getprop ro.config.notification_sound >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt
rem echo ====  adb shell getprop ro.config.alarm_alert  ==== Alerm Alert  >>_TRIAGE_.txt
rem adb shell getprop ro.config.alarm_alert >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt
rem echo ====  adb shell getprop ro.config.ringtone  ==== Ringtone  >>_TRIAGE_.txt
rem adb shell getprop ro.config.ringtone >>_TRIAGE_.txt
rem echo. >>_TRIAGE_.txt
rem echo ====  adb shell getprop rro.config.media_sound  ==== Media Sound  >>_TRIAGE_.txt
rem adb shell getprop rro.config.media_sound >>_TRIAGE_.txt
rem echo.

echo collection saved in _TRIAGE_.txt

rem <<<<<<<<<<<<<<<<<<<<<<<<<<      notes            >>>>>>>>>>>>>>>>>>>>>>>>>>
rem adb backup -all -shared -system -apk -f backup.ab
