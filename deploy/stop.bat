@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0nginx"

echo [1/2] nginx 종료 ...
nginx.exe -s quit 2>nul
if errorlevel 1 (
    echo   graceful quit 실패 - 강제 종료
    taskkill /F /IM nginx.exe >nul 2>&1
)

echo [2/2] Streamlit 종료 ...
REM 콘솔 창(yoon-streamlit) 종료
taskkill /F /FI "WINDOWTITLE eq yoon-streamlit*" >nul 2>&1
REM streamlit를 띄운 python.exe 프로세스 직접 정리
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -match 'streamlit.*webapp\.py' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" 2>nul

echo Done.
endlocal
