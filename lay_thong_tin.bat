@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

echo Lay thong tin may tinh...

set "HOSTNAME=%COMPUTERNAME%"

for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_Processor).Name"') do set "CPU=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_ComputerSystem).Manufacturer"') do set "HANG=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_ComputerSystem).Model"') do set "MODEL=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "[math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB,2)"') do set "RAM=%%a GB"
for /f "delims=" %%a in ('powershell -NoProfile -Command "[math]::Round(((Get-CimInstance Win32_DiskDrive|Measure-Object Size -Sum).Sum)/1GB,2)"') do set "DISK=%%a GB"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_BIOS).SerialNumber"') do set "SERIAL=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_OperatingSystem).Caption"') do set "OS=%%a"

for /f "delims=" %%a in ('powershell -NoProfile -Command "Get-NetIPConfiguration | Where {$_.IPv4DefaultGateway} | Select -First 1 -Expand IPv4Address | Select -Expand IPAddress"') do set "IP=%%a"

for /f "delims=" %%a in ('powershell -NoProfile -Command "Get-NetIPConfiguration | Where {$_.IPv4DefaultGateway} | Select -First 1 -Expand NetAdapter | Select -Expand MacAddress"') do set "MAC=%%a"

if not defined IP set "IP=Khong tim thay"
if not defined MAC set "MAC=Khong tim thay"

for /f "delims=" %%a in ('powershell -NoProfile -Command "if(Test-Path ''HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration''){''Microsoft Office ClickToRun''} elseif(Test-Path ''HKLM:\SOFTWARE\Microsoft\Office''){''Microsoft Office''}"') do set "OFFICE=%%a"

if not defined OFFICE set "OFFICE=Khong tim thay"

for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance SoftwareLicensingService).OA3xOriginalProductKey"') do set "OEMKEY=%%a"

if defined OEMKEY (
set "LICENSEWIN=OEM Key: %OEMKEY%"
) else (
set "LICENSEWIN=Khong co OEM Key"
)

set "LICENSEOFF=Khong tim thay"

for /f "delims=" %%a in ('powershell -NoProfile -Command "Get-ChildItem ''C:\Program Files'',''C:\Program Files (x86)'' -Filter ospp.vbs -Recurse -ErrorAction SilentlyContinue | Select -First 1 -Expand FullName"') do set "OSPP=%%a"

if defined OSPP (
for /f "delims=" %%a in ('cscript //nologo "%OSPP%" /dstatus ^| findstr /i "Last 5"') do set "LICENSEOFF=%%a"
)

set "URL=https://vietxuanvnp-hash.github.io/qltb_sonla"

powershell -NoProfile -Command "$p=@{hostname='%HOSTNAME%';cpu='%CPU%';hang='%HANG%';model='%MODEL%';ram='%RAM%';disk='%DISK%';serial='%SERIAL%';os='%OS%';ip='%IP%';mac='%MAC%';loaiMay='May tinh';office='%OFFICE%';antivirus='Windows Defender';licenseWin='%LICENSEWIN%';licenseOff='%LICENSEOFF%'}; $q=$p.GetEnumerator()|ForEach-Object{$_.Key+'='+[uri]::EscapeDataString([string]$_.Value)}; Start-Process ('%URL%?' + ($q -join '&'))"

echo Hoan tat.
pause


