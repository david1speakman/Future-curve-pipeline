@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title London Morning Email — All Currencies

echo.
echo ============================================================
echo   London Morning Email  --  ALL (SFR / SFI / ER / IR / COR)
echo ============================================================
echo.
echo  PRE-REQUISITE: Most.Traded.Universe.xlsx must be open in Excel
echo  with Bloomberg data fully loaded before continuing.
echo.
pause

python morning_email.py --all

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Pipeline failed — see messages above.
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
