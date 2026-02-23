@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "output" mkdir output
if not exist "upload" mkdir upload

echo.
echo ===== 환경 확인 =====
python --version
echo.

echo ===== 패키지 설치 =====
pip install -r requirements.txt
echo.

echo ===== 프로그램 실행 =====
python run.py

echo.
echo ===== 결과 폴더 열기 =====
start "" "%CD%\output"

pause
