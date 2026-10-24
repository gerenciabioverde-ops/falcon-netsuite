@echo off
setlocal
REM ============================================================
REM   FALCON NETSUITE - Actualizar e iniciar
REM   Un solo clic: baja la ultima version (actualizar.bat) y,
REM   si todo salio bien, arranca el servidor y abre el navegador
REM   (iniciar.bat). Es lo que abre el acceso directo del escritorio.
REM ============================================================
cd /d "%~dp0"
title Falcon NetSuite - Actualizar e iniciar
call "%~dp0actualizar.bat" --no-pause
if errorlevel 1 exit /b 1
call "%~dp0iniciar.bat"
exit /b
