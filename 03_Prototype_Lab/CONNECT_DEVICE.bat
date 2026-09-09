@echo off
cd /d "%~dp0"
title Project 3 - Device Code Sign-In + Lab (fallback)
echo ============================================================
echo  FALLBACK: Device code sign-in + full lab
echo  Use this if CLICK_TO_RUN_LAB browser sign-in failed.
echo ============================================================
echo.
echo 1. PowerShell will show a code (e.g. ABCD-1234)
echo 2. Open https://login.microsoft.com/device
echo 3. Enter the code and sign in with GLOBAL ADMIN
echo 4. Lab runs automatically after sign-in succeeds
echo.
pause
start "JML Lab Device Auth" /wait powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { Set-Location '%~dp0'; . .\Connect-Lab.ps1; Connect-JmlLabTenant -UseDeviceAuth; . .\Run-AllLab.ps1 -SkipConnect }"
if errorlevel 1 (
    echo ERROR: See PowerShell window for details.
    pause
    exit /b 1
)
python "%~dp0..\_update_evaluation.py"
echo SUCCESS: Lab complete.
pause
