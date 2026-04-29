@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PROJECT_ROOT=%~dp0.."
set "PYEXE=%PROJECT_ROOT%\.venv\Scripts\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"

echo [1/2] Streamlit 백엔드 기동 (127.0.0.1:8501) ...
start "yoon-streamlit" cmd /k ""%PYEXE%" -m streamlit run "%PROJECT_ROOT%\webapp.py" --server.port 8501 --server.address 127.0.0.1 --server.headless true --browser.gatherUsageStats false"

REM streamlit이 포트를 잡을 시간을 잠깐 줌
timeout /t 4 /nobreak >nul

echo [2/2] nginx 리버스 프록시 기동 (0.0.0.0:8084) ...
cd /d "%~dp0nginx"
start "" /b nginx.exe

echo.
echo ===============================================
echo   yoon billing app
echo   ▶ http://localhost:8084
echo ===============================================
echo.
echo  종료: deploy\stop.bat
endlocal
