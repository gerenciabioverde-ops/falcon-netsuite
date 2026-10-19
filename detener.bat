@echo off
REM Detiene el servidor de Falcon NetSuite (solo el de esta carpeta; no toca
REM Falcon Contable, Tesoreria ni PDB aunque usen otros puertos).
cd /d "%~dp0"
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*%~dp0.venv\Scripts\python.exe*' -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>&1
echo Servidor de Falcon NetSuite detenido ^(si estaba en ejecucion^).
pause
