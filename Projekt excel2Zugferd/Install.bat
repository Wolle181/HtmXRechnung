@echo off
chcp 65001 > nul
echo Excel2ZUGFeRD AddIn – Installation
echo =====================================
echo.

:: PowerShell-Skript im gleichen Verzeichnis wie diese Batch-Datei ausfuehren
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Excel2Zugferd.ps1" -XlamSource "%~dp0Excel2Zugferd.xlam"

echo.
if %ERRORLEVEL% EQU 0 (
    echo Installation erfolgreich abgeschlossen.
) else (
    echo FEHLER bei der Installation. Bitte Fehlermeldung oben beachten.
)
echo.

REM pause