@echo off
title Excel Automation Toolkit
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found.
    echo Please install Python from https://www.python.org/downloads/
    echo and tick "Add Python to PATH" during installation.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\pythonw.exe" (
    echo First run: setting things up, please wait about 1 minute...
    python -m venv .venv || goto :error
    ".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
    ".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt || goto :error
)

start "" ".venv\Scripts\pythonw.exe" excel_toolkit.py
exit /b 0

:error
echo.
echo Setup failed. Check your internet connection and try again.
pause
exit /b 1
