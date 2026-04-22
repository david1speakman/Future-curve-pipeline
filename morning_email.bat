@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title London Morning Email — SFR / ER / SFI

echo.
echo ============================================================
echo   London Morning Email  --  SFR / ER / SFI
echo ============================================================
python morning_email.py SFR ER SFI

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Pipeline failed — see messages above.
    pause
    exit /b 1
)

rem Open the generated HTML draft in the default browser / Outlook
for /f "delims=" %%f in ('dir /b /od "output\morning_*.html" 2^>nul') do set LATEST=%%f
if defined LATEST (
    echo.
    echo Opening email draft...
    start "" "output\!LATEST!"
)

echo.
echo Done.
pause
