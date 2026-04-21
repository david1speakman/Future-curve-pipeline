@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title London Morning Email — DEMO (no Bloomberg)

echo.
echo ============================================================
echo   London Morning Email  --  DEMO MODE  (no Bloomberg needed)
echo ============================================================
echo.

python morning_email.py SFR ER SFI --no-bloomberg

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Script failed — check Python is installed.
    pause
    exit /b 1
)

for /f "delims=" %%f in ('dir /b /od "output\morning_*.html" 2^>nul') do set LATEST=%%f
if defined LATEST (
    echo.
    echo Opening email draft...
    start "" "output\!LATEST!"
)

echo.
echo Done.
pause
