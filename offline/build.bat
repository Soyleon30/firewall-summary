@echo off
setlocal
cd /d "%~dp0"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_offline.ps1"
if errorlevel 1 (
    echo.
    echo Build failed. See the error message above.
    pause
    exit /b 1
)

echo.
echo Build completed successfully.
echo Output: %~dp0output\FirewallSummary.exe
pause
