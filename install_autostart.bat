@echo off
setlocal enabledelayedexpansion

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_PATH=%SCRIPT_DIR%parsec_monitor.py"

:: Find pythonw.exe
for /f "delims=" %%i in ('where pythonw 2^>nul') do (
    set "PYTHONW=%%i"
    goto :found
)
for /f "delims=" %%i in ('where python 2^>nul') do (
    set "PYTHON=%%i"
    goto :tryconvert
)
echo [ERROR] Python not found.
pause
exit /b 1

:tryconvert
set "PYTHONW=%PYTHON:python.exe=pythonw.exe%"

:found
echo [OK] Found pythonw: %PYTHONW%
echo [OK] Script path:   %SCRIPT_PATH%

set "REG_KEY=HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
set "REG_NAME=ParsecMonitor"

reg add "%REG_KEY%" /v "%REG_NAME%" /t REG_SZ /d "\"%PYTHONW%\" \"%SCRIPT_PATH%\"" /f

if %errorlevel% equ 0 (
    echo.
    echo [OK] Added to autostart successfully.
    echo      Will launch on next Windows login.
) else (
    echo [ERROR] Failed to add registry key.
)

pause
