@echo off
echo ============================================
echo  ALIO Collector - Setup
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo Install Python 3.10+ from https://www.python.org
    pause
    exit /b 1
)

python --version
echo.
echo Installing packages...
echo.

pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [ERROR] Installation failed.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Done! Run run.bat to start the program.
echo ============================================
pause
