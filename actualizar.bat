@echo off
setlocal
REM ============================================================
REM   FALCON NETSUITE - Actualizar
REM   Baja la ultima version desde GitHub (rama main), actualiza
REM   las dependencias y deja todo listo para iniciar.bat.
REM   No toca tus cuentas homologadas (carpeta datos_usuario).
REM ============================================================
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Falcon NetSuite - Actualizar

if not exist .git (
    echo [ERROR] Esta carpeta no es un clon de git. Clona de nuevo el repositorio.
    pause
    exit /b 1
)

echo Cerrando el servidor de Falcon NetSuite si esta abierto...
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*%~dp0.venv\Scripts\python.exe*' -and $_.ProcessId -ne $PID } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>&1

echo.
echo Bajando la ultima version...
REM La conexion con GitHub a veces se corta sola (schannel / close_notify): se reintenta hasta 4 veces
set INTENTO=0
:REINTENTAR
set /a INTENTO+=1
git fetch origin
if errorlevel 1 goto :CORTE
git checkout main
if errorlevel 1 goto :FALLO
REM Con lo que ya se bajo en el fetch (sin volver a conectarse): si la conexion es inestable, un segundo acceso a GitHub fallaba
git merge --ff-only origin/main
if errorlevel 1 goto :FALLO
goto :BAJADO

:CORTE
if %INTENTO% GEQ 4 goto :FALLO
echo.
echo La conexion con GitHub se corto. Reintentando (intento %INTENTO% de 4)...
timeout /t 5 /nobreak >nul
goto :REINTENTAR

:BAJADO

if not exist ".venv\Scripts\python.exe" goto :LISTO
echo.
echo Actualizando dependencias...
".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :FALLO

:LISTO
echo.
git log --oneline -1
echo.
echo [OK] Listo. Ahora ejecuta iniciar.bat
pause
exit /b 0

:FALLO
echo.
echo [ERROR] No se pudo actualizar. Copia los mensajes de arriba y enviamelos.
pause
exit /b 1
