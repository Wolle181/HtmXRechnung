#Requires -Version 5.0
# Installiert Excel2ZUGFeRD nach C:\Rechnungen\Excel2Zugferd
# und richtet das Excel-AddIn fuer den aktuellen Benutzer ein.

$ErrorActionPreference = "Stop"

$installDir = "C:\Rechnungen\Excel2Zugferd"
$payloadDir = Join-Path $PSScriptRoot "Excel2ZugferdSetupPayload"

Write-Host "=== Excel2ZUGFeRD Setup ===" -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $payloadDir)) {
    Write-Host "FEHLER: Payload-Verzeichnis nicht gefunden:" -ForegroundColor Red
    Write-Host "  $payloadDir" -ForegroundColor Red
    Write-Host "Bitte sicherstellen, dass Excel2ZugferdSetup.bat und der Ordner" -ForegroundColor Red
    Write-Host "Excel2ZugferdSetupPayload\ im selben Verzeichnis liegen." -ForegroundColor Red
    exit 1
}

# Installationsverzeichnis anlegen
Write-Host "1. Erstelle Installationsverzeichnis: $installDir"
New-Item -ItemType Directory -Path $installDir -Force | Out-Null
Write-Host "   OK." -ForegroundColor Green

# Programmdateien kopieren
Write-Host ""
Write-Host "2. Kopiere Programmdateien..."
Copy-Item "$payloadDir\*" $installDir -Recurse -Force
Write-Host "   OK." -ForegroundColor Green

# Excel-AddIn installieren
Write-Host ""
Write-Host "3. Installiere Excel-AddIn..."
$installBat = Join-Path $installDir "Install.bat"
Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$installBat`"" -WorkingDirectory $installDir -Wait

Write-Host ""
Write-Host "=== Setup abgeschlossen! ===" -ForegroundColor Cyan
Write-Host "Excel2ZUGFeRD ist unter $installDir installiert."
