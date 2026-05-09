@echo off
chcp 65001 > nul
echo Excel2ZUGFeRD AddIn – Deinstallation
echo ========================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-Excel2Zugferd.ps1"

echo.
if %ERRORLEVEL% EQU 0 (
    echo Deinstallation erfolgreich abgeschlossen.
) else (
    echo FEHLER bei der Deinstallation. Bitte Fehlermeldung oben beachten.
)
echo.
pause
