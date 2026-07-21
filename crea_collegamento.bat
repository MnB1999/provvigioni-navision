@echo off
rem Da eseguire UNA SOLA VOLTA su ogni PC
rem Crea sul desktop il collegamento "Generatore Excel provvigioni"

set "PROGETTO=%~dp0"
set "PROGETTO=%PROGETTO:~0,-1%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "$ErrorActionPreference='Stop';" ^
 "$pyw=(Get-Command pyw.exe).Source;" ^
 "$desktop=[Environment]::GetFolderPath('Desktop');" ^
 "$shell=New-Object -ComObject WScript.Shell;" ^
 "$lnk=$shell.CreateShortcut((Join-Path $desktop 'Generatore Excel provvigioni.lnk'));" ^
 "$lnk.TargetPath=$pyw;" ^
 "$lnk.Arguments=[char]34+'%PROGETTO%\app.py'+[char]34;" ^
 "$lnk.WorkingDirectory='%PROGETTO%';" ^
 "$lnk.IconLocation='%PROGETTO%\icona.ico,0';" ^
 "$lnk.Description='Generatore Excel provvigioni';" ^
 "$lnk.Save();" ^
 "Write-Host 'Collegamento creato sul desktop con nome: Generatore Excel provvigioni'"

pause
