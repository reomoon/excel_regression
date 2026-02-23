@echo off
REM 한글 인코딩 설정 (UTF-8)
chcp 65001 >nul

REM 배치 파일이 위치한 폴더로 자동 이동
cd /d "%~dp0"

REM Python 설치 여부 확인
echo.
echo ===== 설치 상태 확인 =====
python --version
if %errorlevel% neq 0 (
    echo [ERROR] Python: 설치되지 않음
    echo.
    pause
    exit /b 1
)

REM openpyxl 설치 여부 확인
python -c "import openpyxl" >nul 2>&1
if %errorlevel% equ 0 (
    echo [OK] openpyxl: 설치됨
) else (
    echo [INFO] openpyxl: 설치 진행 중...
)
echo.

REM 필요한 패키지 자동 설치
echo ===== 패키지 자동 설치 =====
pip install -r requirements.txt
echo [OK] 패키지 설치 완료
echo.

REM Python 스크립트 실행
echo ===== 프로그램 실행 =====
python run.py

REM 생성된 output 폴더 자동으로 열기
start "" "%CD%\output"

REM 완료 후 창이 닫히지 않도록 유지 (사용자가 결과 확인 가능)
pause
