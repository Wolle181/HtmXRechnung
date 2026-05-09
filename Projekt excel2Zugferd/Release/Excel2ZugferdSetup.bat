@echo off
chcp 65001 > nul
echo Excel2ZUGFeRD – Setup
echo =======================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Excel2ZugferdSetup.ps1"
echo.
if %ERRORLEVEL% EQU 0 (
    echo Setup erfolgreich abgeschlossen.
) else (
    echo FEHLER beim Setup. Bitte Fehlermeldung oben beachten.
)
echo.
pause
