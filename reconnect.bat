@echo off
taskkill /im notepad.exe /f
cls
echo Reconnecting Personal Email...

REM  This is will launch a tool to provide location TMs that have used the
REM  MW User Setup Tool.
REM
REM  The tool checks backed up registry files against active AD accounts to
REM  provide a more precise check of actual users that require running of
REM  the Undo-UST PowerShell.
REM  
REM  Place this .bat file in the same folder as the PowerShell file.
REM  

powershell.exe -STA -nologo -file "%~dp0reconnect.ps1"
cls

echo (Press ENTER to exit this window)

TIMEOUT /T 5

REM  Forcing to open the Powershell in Admin. mode.  -  TMs will be prompted for elevated credentials
REM  PowerShell.exe -NoProfile -Command "& {Start-Process PowerShell.exe -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%~dp0reconnect.ps1""' -Verb RunAs}"

exit