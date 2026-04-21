@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

rem ── Usage: morning_custom.bat PRODUCT [PRODUCT ...] [--ib-paste]
rem
rem    morning_custom.bat SFR
rem    morning_custom.bat SFR IR COR
rem    morning_custom.bat ER SFI --ib-paste
rem
rem    Valid products: SFR  SFI  ER  IR  COR
rem    Add --ib-paste to automatically paste the text table into
rem    Bloomberg IB (requires Bloomberg IB window to be open).

if "%~1"=="" (
    echo.
    echo  Usage: morning_custom.bat PRODUCT [PRODUCT ...] [--ib-paste]
    echo.
    echo  Products : SFR  SFI  ER  IR  COR
    echo.
    echo  Examples:
    echo    morning_custom.bat SFR
    echo    morning_custom.bat SFR IR COR
    echo    morning_custom.bat ER SFI --ib-paste
    echo.
    pause
    exit /b 0
)

title London Morning Email — %*

echo.
echo ============================================================
echo   London Morning Email  --  %*
echo ============================================================
echo.
echo  PRE-REQUISITE: Most.Traded.Universe.xlsx must be open in Excel
echo  with Bloomberg data fully loaded before continuing.
echo.
pause

python morning_email.py %*

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
