@echo off
cd /d "%~dp0"

if not exist "output" mkdir output
if not exist "upload" mkdir upload

for /f "tokens=*" %%i in ('py -3 -c "import sys; print(sys.executable)"') do set PYTHON_PATH=%%i

if "%PYTHON_PATH%"=="" (
    echo Error: Python not found
    echo Please install Python with PATH enabled
    pause
    exit /b 1
)

echo.
echo Checking environment...
"%PYTHON_PATH%" --version
echo.

echo Installing packages...
"%PYTHON_PATH%" -m pip install -r requirements.txt
echo.

echo Running program...
"%PYTHON_PATH%" run.py

echo.
start "" "%CD%\output"

pause
