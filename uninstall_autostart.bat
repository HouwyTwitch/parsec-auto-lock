@echo off
set "REG_KEY=HKCU\Software\Microsoft\Windows\CurrentVersion\Run"
set "REG_NAME=ParsecMonitor"

reg delete "%REG_KEY%" /v "%REG_NAME%" /f

if %errorlevel% equ 0 (
    echo [OK] Parsec Monitor removed from autostart.
) else (
    echo [INFO] Entry was not found in autostart.
)

pause
