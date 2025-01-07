@echo off
echo Installing KWS-helper
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

echo Stopping current AutoHotkey process
%PWS% Stop-Process -Name AutoHotkey64

echo Downloading newest version from Github
%PWS% -command "Invoke-WebRequest https://github.com/CVanmarcke/KWS-helper/archive/refs/heads/main.zip -O P:\uzldownloads\KWSHelper.zip"

echo Extracting to P:\KWS-helper-main
%PWS% -command "Expand-Archive -Force 'P:\uzldownloads\KWSHelper.zip' 'P:\'"

set TARGET='P:\KWS-helper-main\AutoHotkey64.exe'
set SHORTCUT='P:\uzlsystem\StartMenu\Programs\Startup\kws-helper.lnk'
%PWS% -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut(%SHORTCUT%); $S.TargetPath = %TARGET%; $S.Save()"

start /D P:\KWS-helper-main\ AutoHotkey64.exe