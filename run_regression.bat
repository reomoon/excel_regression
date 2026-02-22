@echo off
REM 한글 인코딩 설정 (UTF-8)
chcp 65001 >nul

REM 배치 파일이 위치한 폴더로 자동 이동
cd /d "%~dp0"

REM Python 스크립트 실행
python run.py

REM 생성된 output 폴더 자동으로 열기
start "" "%CD%\output"

REM 완료 후 창이 닫히지 않도록 유지 (사용자가 결과 확인 가능)
pause
