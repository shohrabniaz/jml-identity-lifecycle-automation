@echo off
cd /d "%~dp0"
title Project 3 JML Lab - FREE Entra ID only
echo ============================================================
echo  SCI700 Project 3 - Full Lab Runner (100%% FREE)
echo  Tenant: niazkhanghostgmail.onmicrosoft.com
echo ============================================================
echo.
echo A NEW PowerShell window will open.
echo Sign in when your browser opens (GLOBAL ADMIN account).
echo.
pause
start "JML Lab" /wait powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-AllLab.ps1"
if errorlevel 1 (
    echo.
    echo ERROR: Lab script failed. See the PowerShell window for details.
    echo.
    echo If browser sign-in failed, try: CONNECT_DEVICE.bat
    pause
    exit /b 1
)
echo.
python "%~dp0..\_update_evaluation.py"
echo.
echo SUCCESS: Lab complete. Evidence in: %~dp0evidence\
pause
