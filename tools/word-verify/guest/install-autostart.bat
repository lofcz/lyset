@echo off
rem One-time setup inside the Windows VM, run from the shared folder:
rem   1. copy the service and its launcher into %LOCALAPPDATA%\lyset-word-verify
rem   2. allow inbound TCP 47400 (asks for elevation once). The Winboat compose
rem      file must forward it: USER_PORTS += 47400, ports += 127.0.0.1:47400:47400
rem   3. start the service at every logon of this user, and start it now
rem Undo: schtasks /delete /tn "lyset word-verify" /f
rem       netsh advfirewall firewall delete rule name=lyset-word-verify
setlocal
set "LOCAL=%LOCALAPPDATA%\lyset-word-verify"
if not exist "%LOCAL%" mkdir "%LOCAL%"
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'word-verify.exe' -or $_.CommandLine -match 'word-verify.*\\start\.bat' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
copy /y "%~dp0start.bat" "%LOCAL%\start.bat" >nul
copy /y "%~dp0word-verify.exe" "%LOCAL%\word-verify.exe" >nul
copy /y "%~dp0token.txt" "%LOCAL%\token.txt" >nul
> "%LOCAL%\share.txt" echo %~dp0
rem Replace the rule every time so a port change never leaves a stale opening.
powershell -NoProfile -Command "Start-Process cmd -Verb RunAs -Wait -WindowStyle Hidden -ArgumentList '/c netsh advfirewall firewall delete rule name=lyset-word-verify & netsh advfirewall firewall add rule name=lyset-word-verify dir=in action=allow protocol=TCP localport=47400'"
schtasks /create /tn "lyset word-verify" /sc onlogon /rl limited /f /tr "cmd /c start \"word-verify\" /min \"%LOCAL%\start.bat\""
start "word-verify" /min "%LOCAL%\start.bat"
