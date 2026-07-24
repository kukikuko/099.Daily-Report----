@echo off
setlocal

cd /d "%~dp0"

echo Starting build from main.spec...

where py >nul 2>&1
if %errorlevel%==0 (
    py -3 -m PyInstaller main.spec
) else (
    python -m PyInstaller main.spec
)

if errorlevel 1 (
    echo.
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo Build complete: "%~dp0dist\main.exe"
pause
