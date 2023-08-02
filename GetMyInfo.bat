@echo off
setlocal enabledelayedexpansion

REM Get PC name
for /f "tokens=2 delims==" %%A in ('wmic computersystem get name /value') do (
    set "pcname=%%A"
    goto :break_pcname
)
:break_pcname

REM Get device serial number
for /f "tokens=2 delims==" %%A in ('wmic bios get serialnumber /value') do (
    set "deviceserial=%%A"
    goto :break_deviceserial
)
:break_deviceserial

REM Get IP address
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4 Address"') do (
    set "ipaddress=%%A"
    goto :break_ipaddress
)
:break_ipaddress

REM Get Windows serial number
for /f "tokens=2 delims==" %%A in ('wmic path SoftwareLicensingService get OA3xOriginalProductKey /value') do (
    set "windowsserial=%%A"
    goto :break_windowsserial
)
:break_windowsserial

REM Get MS Office serial number
for /f "tokens=3 delims=: " %%A in ('cscript "%ProgramFiles(x86)%\Microsoft Office\Office16\OSPP.VBS" /dstatus') do (
    set "officeversion=%%A"
    goto :break_officeversion
)
:break_officeversion

echo PC Name: %pcname%
echo Device Serial: %deviceserial%
echo IP Address: %ipaddress%
echo Windows Serial: %windowsserial%
echo MS Office Serial: %officeversion%

endlocal
