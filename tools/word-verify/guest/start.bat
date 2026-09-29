@echo off
rem word-verify service loop. Installed into %LOCALAPPDATA%\lyset-word-verify by
rem install-autostart.bat and run at logon. POST /update stages
rem word-verify.new.exe and exits, and this loop swaps it in.
setlocal
set "LOCAL=%~dp0"
set /p SHARE=<"%LOCAL%share.txt"
:loop
rem The share can serve a stale listing for a while, so keep retrying the token copy.
if not exist "%LOCAL%token.txt" copy /y "%SHARE%token.txt" "%LOCAL%token.txt" >nul 2>&1
if exist "%LOCAL%word-verify.new.exe" move /y "%LOCAL%word-verify.new.exe" "%LOCAL%word-verify.exe" >nul
if not exist "%LOCAL%word-verify.exe" copy /y "%SHARE%word-verify.exe" "%LOCAL%word-verify.exe" >nul
"%LOCAL%word-verify.exe" serve --listen 0.0.0.0:47400 --token-file "%LOCAL%token.txt" 2>> "%LOCAL%service.log"
timeout /t 2 /nobreak >nul
goto loop
