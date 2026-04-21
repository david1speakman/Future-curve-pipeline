@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
title London Morning Email — TEST SEND

echo.
echo ============================================================
echo   London Morning Email  --  TEST SEND to ds@coexpartners.com
echo   (demo data — no Bloomberg or Excel needed)
echo ============================================================
echo.

python morning_email.py SFR ER SFI --no-bloomberg --send-to ds@coexpartners.com

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Script failed — see messages above.
    echo         Make sure Outlook is running and you are logged in.
    pause
    exit /b 1
)

echo.
echo Done — check your inbox at ds@coexpartners.com
pause
