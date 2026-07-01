@echo off
echo Building sap-cases.exe ...

where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: python not found. Install Python 3.10+ from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo Installing dependencies...
pip install pyinstaller playwright requests openpyxl pywin32 >nul 2>&1
playwright install chromium

echo Running PyInstaller...
pyinstaller --onefile --name sap-cases --clean sap_cases.py

if errorlevel 1 (
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo Done! Executable: dist\sap-cases.exe
echo.
echo Distribute to colleagues:
echo   dist\sap-cases.exe  +  config.json  (same folder)
echo.
echo First run on each machine:
echo   sap-cases.exe --customer 450162 --relogin
echo   sap-cases.exe setup-cron
pause
