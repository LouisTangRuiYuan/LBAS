@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo Registration to Registry Launcher
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this computer.
    echo Please install Python first, then run this file again.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
)

echo Installing required library...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo Failed to install required library.
    pause
    exit /b 1
)

echo.
echo Running script...
".venv\Scripts\python.exe" registration_to_registry.py
set EXITCODE=%ERRORLEVEL%

echo.
if %EXITCODE%==0 (
    echo Script finished.
) else (
    echo Script ended with an error code: %EXITCODE%
)

echo.
pause
exit /b %EXITCODE%
