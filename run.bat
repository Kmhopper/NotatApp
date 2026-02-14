@echo off
setlocal EnableExtensions
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Fant ikke .venv. Kjorer setup forst...
    call setup.bat
    if errorlevel 1 exit /b 1
)

".venv\Scripts\python.exe" app.py
if errorlevel 1 (
    echo.
    echo Appen avsluttet med feil.
    pause
    exit /b 1
)

exit /b 0
