@echo off
chcp 65001 > nul
cd /d "%~dp0"

if exist "%LOCALAPPDATA%\Python\pythoncore-3.14-64\Scripts" (
    set "PATH=%LOCALAPPDATA%\Python\pythoncore-3.14-64;%LOCALAPPDATA%\Python\pythoncore-3.14-64\Scripts;%PATH%"
)

echo [1/3] 기존 포트(8504) 충돌 프로세스를 확인 및 정리합니다...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 8504 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess | Where-Object { $_ -gt 0 } | Stop-Process -Force -ErrorAction SilentlyContinue"

echo [2/3] 기본 웹 브라우저에서 http://localhost:8504 를 엽니다...
start "" "http://localhost:8504"

echo [3/3] 러셀대치 강의계획서 검수 툴 서버를 시작합니다 (http://localhost:8504)...
streamlit run app.py --server.port 8504 --server.address 0.0.0.0 --server.headless true
pause
