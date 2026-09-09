@echo off
cd /d "%~dp0"
echo Exporting live lab state from Entra (requires sign-in)...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Export-LabEvidence.ps1"
if errorlevel 1 (
    echo Export failed - sign in first or use JSON-only PNG render.
    goto render
)
:render
echo Rendering PNG evidence files...
python "%~dp0..\_render_screenshot_pngs.py"
echo.
echo Screenshots folder: %~dp0..\05_Evaluation\screenshots\
dir "%~dp0..\05_Evaluation\screenshots\*.png" 2>nul
pause
