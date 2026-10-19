@echo off
setlocal
cd /d "%~dp0"
title Falcon NetSuite CSV

where python >nul 2>nul
if errorlevel 1 (
    echo No se encontro Python. Instale Python 3.10 o superior y vuelva a intentar.
    pause
    exit /b 1
)

rem Crea o repara el entorno virtual si falta o esta roto (por ejemplo, tras mover la carpeta)
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys" >nul 2>nul
    if errorlevel 1 rmdir /s /q ".venv"
)
if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual...
    python -m venv .venv
    if errorlevel 1 goto :error
)

rem Instala dependencias solo si faltan
".venv\Scripts\python.exe" -c "import fastapi, uvicorn, openpyxl, jinja2, multipart" >nul 2>nul
if errorlevel 1 (
    echo Instalando dependencias...
    ".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt
    if errorlevel 1 goto :error
)

rem Busca el primer puerto libre desde el 8002. Se saltan los reservados por otras apps Falcon:
rem   8000 = Falcon Contable, 8001 = Falcon Tesoreria, 8003 = Falcon PDB. Agregue aqui otras apps.
set "RESERVADOS= 8000 8001 8003 "
set PUERTO=
for /l %%P in (8002,1,8099) do if not defined PUERTO call :probar %%P
if not defined PUERTO (
    echo No se encontro un puerto libre entre 8002 y 8099.
    goto :fin
)

echo.
echo Servidor en http://localhost:%PUERTO%  (cierre esta ventana o pulse Ctrl+C para detenerlo)
start "" /b cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:%PUERTO%"
set FALCON_REINICIAR=1
:servidor
".venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port %PUERTO%
rem Codigo 75 = el boton Actualizar pidio reiniciar: reinstala dependencias y vuelve a arrancar
if "%errorlevel%"=="75" (
    echo.
    echo Aplicando actualizacion...
    ".venv\Scripts\python.exe" -m pip install -q --disable-pip-version-check -r requirements.txt
    timeout /t 2 /nobreak >nul
    goto :servidor
)
goto :fin

:error
echo.
echo Ocurrio un error durante la preparacion. Revise los mensajes anteriores.
:fin
pause
exit /b

:probar
echo %RESERVADOS% | find " %1 " >nul && goto :eof
netstat -ano | findstr /R /C:":%1 *0\.0\.0\.0:0" /C:":%1 *\[::\]:0" >nul && goto :eof
set PUERTO=%1
goto :eof
