@echo off
cd /d "%~dp0"
title Rebuild Most.Traded.Universe.xlsx

echo.
echo ============================================================
echo   Rebuilding Most.Traded.Universe.xlsx
echo ============================================================
echo.

python build_universe.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed — see messages above.
    pause
    exit /b 1
)

echo.
echo Opening workbook...
start "" "Workbook\Most.Traded.Universe.xlsx"

echo.
echo Done.
pause
