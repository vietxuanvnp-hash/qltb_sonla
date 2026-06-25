@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

echo Lay thong tin may tinh tu dong...

set "HOSTNAME=%COMPUTERNAME%"

for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_Processor).Name"') do set "CPU=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_ComputerSystem).Manufacturer"') do set "HANG=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_ComputerSystem).Model"') do set "MODEL=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "[math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory/1GB,2)"') do set "RAM=%%a GB"
for /f "delims=" %%a in ('powershell -NoProfile -Command "[math]::Round(((Get-CimInstance Win32_DiskDrive|Measure-Object Size -Sum).Sum)/1GB,2)"') do set "DISK=%%a GB"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_BIOS).SerialNumber"') do set "SERIAL=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance Win32_OperatingSystem).Caption"') do set "OS=%%a"

for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-NetIPConfiguration | Where {$_.IPv4DefaultGateway} | Select -First 1 -Expand IPv4Address).IPAddress"') do set "IP=%%a"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-NetIPConfiguration | Where {$_.IPv4DefaultGateway} | Select -First 1 -Expand NetAdapter).MacAddress"') do set "MAC=%%a"

if not defined IP set "IP=N/A"
if not defined MAC set "MAC=N/A"

set "OFFICE=Khong tim thay"
for /f "delims=" %%a in ('powershell -NoProfile -Command "$p='HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'; if(Test-Path $p){$x=Get-ItemProperty $p; 'Office '+$x.ClientVersionToReport}"') do set "OFFICE=%%a"

set "LICENSEWIN=Khong tim thay"
for /f "delims=" %%a in ('powershell -NoProfile -Command "(Get-CimInstance SoftwareLicensingService).OA3xOriginalProductKey"') do set "LICENSEWIN=OEM %%a"

set "LICENSEOFF=Khong tim thay"

set "URL=https://vietxuanvnp-hash.github.io/qltb_sonla"

powershell -NoProfile -Command "$d=@{hostname='%HOSTNAME%';cpu='%CPU%';hang='%HANG%';model='%MODEL%';ram='%RAM%';disk='%DISK%';serial='%SERIAL%';os='%OS%';ip='%IP%';mac='%MAC%';loaiMay='May tinh';office='%OFFICE%';licenseWin='%LICENSEWIN%';licenseOff='%LICENSEOFF%';antivirus='Windows Defender'}; $q=$d.GetEnumerator()|ForEach-Object{$_.Key+'='+[uri]::EscapeDataString([string]$_.Value)}; Start-Process ('%URL%?'+($q -join '&'))"

echo Hoan tat.
echo Hay Chon trinh duyet web de vao hoan thanh phieu..
pause
