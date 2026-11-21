@echo off
REM ---------------------------------------------------------------
REM  FOR THE SELLER: builds dist\ExcelToolkit.exe (run on Windows)
REM  Not needed by customers.
REM ---------------------------------------------------------------
cd /d "%~dp0"
python -m venv .build-venv || goto :error
".build-venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
".build-venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt pyinstaller || goto :error
".build-venv\Scripts\pyinstaller.exe" --noconfirm --onefile --windowed ^
    --name ExcelToolkit excel_toolkit.py || goto :error
echo.
echo Done! Your app is at: dist\ExcelToolkit.exe
pause
exit /b 0

:error
echo Build failed.
pause
exit /b 1
