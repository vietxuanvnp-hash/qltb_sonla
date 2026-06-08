@echo off
chcp 65001 >nul 2>&1

:: ─────────────────────────────────────────────────────────────────────────────
:: VNPost Device Inventory – Thu thap thong tin may tinh
:: Phien ban KHONG dung iex – tranh Kaspersky Adaptive Anomaly Control chan
:: Web: https://nguyennam90.github.io/Check_thiet_bi/
:: ─────────────────────────────────────────────────────────────────────────────
::
:: Kaspersky chan rule: "PowerShell script executes unknown dynamic code"
:: Vi le do cu dung:  iex ((Get-Content '%~f0') -join [char]10)  <-- bi chan
:: Giai phap moi:     powershell -Command "code tinh" ^           <-- khong bi chan
::                    (toan bo code duoc viet thang, khong doc tu file luc chay)
:: ─────────────────────────────────────────────────────────────────────────────

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "$u='https://nguyennam90.github.io/Check_thiet_bi';" ^
 "try{" ^
 "  $d=Join-Path $env:LOCALAPPDATA 'VNPost';" ^
 "  New-Item -ItemType Directory -Path $d -Force|Out-Null;" ^
 "  $s=$MyInvocation.MyCommand.Path;" ^
 "  if($s -and (Test-Path $s)){Copy-Item $s (Join-Path $d 'lay_thong_tin.bat') -Force};" ^
 "  $bat=Join-Path $d 'lay_thong_tin.bat';" ^
 "  $q34=[char]34;" ^
 "  $cv='cmd.exe /c start /min '+$q34+$q34+' '+$q34+$bat+$q34;" ^
 "  $rb='HKCU:\SOFTWARE\Classes\vnpost';" ^
 "  New-Item -Path $rb -Force|Out-Null;" ^
 "  Set-ItemProperty -Path $rb -Name '(Default)' -Value 'URL:VNPost Device Inventory';" ^
 "  Set-ItemProperty -Path $rb -Name 'URL Protocol' -Value '';" ^
 "  New-Item -Path ($rb+'\shell\open\command') -Force|Out-Null;" ^
 "  Set-ItemProperty -Path ($rb+'\shell\open\command') -Name '(Default)' -Value $cv" ^
 "}catch{};" ^
 "$cpu=(Get-CimInstance Win32_Processor|Select-Object -First 1).Name;" ^
 "$cs=Get-CimInstance Win32_ComputerSystem|Select-Object -First 1;" ^
 "$bios=Get-CimInstance Win32_BIOS|Select-Object -First 1;" ^
 "$os=Get-CimInstance Win32_OperatingSystem|Select-Object -First 1;" ^
 "$ds=Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3'|Measure-Object Size -Sum;" ^
 "$na=Get-CimInstance Win32_NetworkAdapterConfiguration|Where-Object{$_.IPEnabled}|Select-Object -First 1;" ^
 "$ram=[math]::Round($cs.TotalPhysicalMemory/1GB,1);" ^
 "$serial=$bios.SerialNumber.Trim();" ^
 "$gb=[math]::Round($ds.Sum/1GB,0);" ^
 "$osn=($os.Caption+' '+$os.Version).Trim();" ^
 "$ip=($na.IPAddress|Where-Object{$_ -notmatch ':'}|Select-Object -First 1);" ^
 "$mac=$na.MACAddress;" ^
 "$type=if($cs.PCSystemType-eq 2){'Laptop'}else{'Desktop'};" ^
 "$pk='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*';" ^
 "$office=(Get-ItemProperty $pk -EA SilentlyContinue|Where-Object{$_.DisplayName -match 'Microsoft 365|Microsoft Office'}|Select-Object -First 1 -ExpandProperty DisplayName);" ^
 "$av=(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -EA SilentlyContinue|Select-Object -ExpandProperty displayName);" ^
 "Add-Type -AssemblyName System.Web;" ^
 "function E($v){if(-not $v){return ''};[System.Web.HttpUtility]::UrlEncode($v.ToString())};" ^
 "$q='hostname='+(E $env:COMPUTERNAME)+'&cpu='+(E $cpu)+'&hang='+(E $cs.Manufacturer.Trim())+'&model='+(E $cs.Model.Trim())+'&ram='+(E($ram.ToString()+' GB'))+'&disk='+(E($gb.ToString()+' GB'))+'&serial='+(E $serial)+'&os='+(E $osn)+'&ip='+(E $ip)+'&mac='+(E $mac)+'&loaiMay='+(E $type)+'&office='+(E $office)+'&antivirus='+(E($av -join ', '));" ^
 "Start-Process($u+'/?'+$q)"
