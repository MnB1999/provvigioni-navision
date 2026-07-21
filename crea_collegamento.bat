@echo off
rem Da eseguire UNA SOLA VOLTA su ogni PC
rem Crea sul desktop il collegamento "Generatore Excel provvigioni"

set "PROGETTO=%~dp0"
set "PROGETTO=%PROGETTO:~0,-1%"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "$ErrorActionPreference='Stop';" ^
 "$exe='%PROGETTO%\.venv\Scripts\provvigioni-navision.exe';" ^
 "if (-not (Test-Path $exe)) { throw \"Non trovato: $exe. Esegui prima: .venv\Scripts\pip install -e .\" };" ^
 "$desktop=[Environment]::GetFolderPath('Desktop');" ^
 "$shell=New-Object -ComObject WScript.Shell;" ^
 "$lnk=$shell.CreateShortcut((Join-Path $desktop 'Generatore Excel provvigioni.lnk'));" ^
 "$lnk.TargetPath=$exe;" ^
 "$lnk.Arguments='';" ^
 "$lnk.WorkingDirectory='%PROGETTO%';" ^
 "$lnk.IconLocation='%PROGETTO%\icona.ico,0';" ^
 "$lnk.Description='Generatore Excel provvigioni';" ^
 "$lnk.Save();" ^
 "Write-Host 'Collegamento creato sul desktop con nome: Generatore Excel provvigioni'"

pause
