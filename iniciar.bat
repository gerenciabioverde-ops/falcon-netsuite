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
    ".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
    if errorlevel 1 goto :error
)

echo.
echo Servidor en http://localhost:8000  (cierre esta ventana o pulse Ctrl+C para detenerlo)
start "" /b cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:8000"
".venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8000
goto :fin

:error
echo.
echo Ocurrio un error durante la preparacion. Revise los mensajes anteriores.
:fin
pause
