@echo off
setlocal
REM ============================================================
REM   FALCON NETSUITE - Crear acceso directo en el escritorio
REM   Crea el icono "Falcon NetSuite" que actualiza e inicia la
REM   aplicacion con un solo doble clic. Ejecutar una sola vez
REM   (y de nuevo solo si se mueve la carpeta).
REM ============================================================
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=[Environment]::GetFolderPath('Desktop'); $s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\Falcon NetSuite.lnk'); $s.TargetPath='%~dp0actualizar_e_iniciar.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='shell32.dll,13'; $s.Description='Actualiza e inicia Falcon NetSuite'; $s.Save(); Write-Host ('Acceso directo creado en: ' + $d)"
if errorlevel 1 (
    echo [ERROR] No se pudo crear el acceso directo.
    pause
    exit /b 1
)
echo.
echo Listo: usa el icono "Falcon NetSuite" del escritorio para actualizar e iniciar.
pause
