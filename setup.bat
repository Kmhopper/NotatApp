@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ======================================
echo   Notat Overlay - Setup
echo ======================================
echo.

set "PY_CMD="
where py >nul 2>nul
if %errorlevel%==0 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PY_CMD=python"
    )
)

if not defined PY_CMD (
    echo [FEIL] Fant ikke Python.
    echo Installer Python 3.11+ og prov igjen.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Oppretter virtuelt miljo (.venv)...
    %PY_CMD% -m venv .venv
    if errorlevel 1 (
        echo [FEIL] Klarte ikke opprette .venv
        pause
        exit /b 1
    )
) else (
    echo Fant eksisterende .venv
)

echo.
echo Oppgraderer pip...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
    echo [FEIL] Pip-oppgradering feilet.
    pause
    exit /b 1
)

echo.
echo Installerer avhengigheter fra requirements.txt...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [FEIL] Installasjon av avhengigheter feilet.
    pause
    exit /b 1
)

echo.
echo [OK] Setup ferdig.
echo Start appen med: run.bat
pause
exit /b 0
