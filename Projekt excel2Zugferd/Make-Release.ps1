#Requires -Version 5.0
# Erstellt das Release-Paket:
#   Release\Excel2ZugferdSetupPayload\  – alle Programmdateien
#   Release\Excel2ZugferdSetup.bat      – Endanwender-Installer (Doppelklick)
#   Release\Excel2ZugferdSetup.ps1      – eigentliche Setup-Logik

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$releaseDir = Join-Path $root "Release"
$payloadDir = Join-Path $releaseDir "Excel2ZugferdSetupPayload"

# ---------------------------------------------------------------------------
# 1. Payload-Verzeichnis anlegen und befuellen
# ---------------------------------------------------------------------------
Write-Host "=== Excel2ZUGFeRD Release erstellen ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Erstelle Payload-Verzeichnis..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null
Write-Host "   $payloadDir" -ForegroundColor Green

$items = @(
    "excel2zugferd.exe",
    "Excel2Zugferd.xlam",
    "Install-Excel2Zugferd.ps1",
    "Install.bat",
    "Uninstall-Excel2Zugferd.ps1",
    "Uninstall.bat",
    "_internal"
)

Write-Host ""
Write-Host "2. Kopiere Dateien in Payload-Verzeichnis..." -ForegroundColor Cyan
foreach ($item in $items) {
    $src = Join-Path $root $item
    if (-not (Test-Path $src)) {
        Write-Host "   WARNUNG: '$item' nicht gefunden, wird uebersprungen." -ForegroundColor Yellow
        continue
    }
    if (Test-Path $src -PathType Container) {
        $dst = Join-Path $payloadDir $item
        if (Test-Path $dst) { Remove-Item $dst -Recurse -Force }
        Copy-Item $src $dst -Recurse -Force
        Write-Host "   Ordner kopiert:  $item\" -ForegroundColor Green
    }
    else {
        Copy-Item $src $payloadDir -Force
        Write-Host "   Datei kopiert:   $item" -ForegroundColor Green
    }
}

# ---------------------------------------------------------------------------
# 2. Release\Excel2ZugferdSetup.ps1 schreiben
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "3. Erstelle Setup-Script..." -ForegroundColor Cyan

$setupPs1Path = Join-Path $releaseDir "Excel2ZugferdSetup.ps1"

$setupPs1 = @'
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
'@

$setupPs1 | Set-Content -Path $setupPs1Path -Encoding UTF8
Write-Host "   $setupPs1Path" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. Release\Excel2ZugferdSetup.bat schreiben
# ---------------------------------------------------------------------------
$setupBatPath = Join-Path $releaseDir "Excel2ZugferdSetup.bat"

$setupBat = @'
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
'@

$setupBat | Set-Content -Path $setupBatPath -Encoding Default
Write-Host "   $setupBatPath" -ForegroundColor Green

$src = Join-Path $root "doc\README Endanwender.html"
Copy-Item $src $releaseDir -Force
$src = Join-Path $root "doc\README Endanwender.pdf"
Copy-Item $src $releaseDir -Force

# ---------------------------------------------------------------------------
# Zusammenfassung
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "=== Release-Paket fertig! ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Struktur:" -ForegroundColor White
Write-Host "  Release\"
Write-Host "  ├── Excel2ZugferdSetup.bat       <- Doppelklick zum Installieren"
Write-Host "  ├── Excel2ZugferdSetup.ps1"
Write-Host "  ├── README Endanwender.pdf"
Write-Host "  ├── README Endanwender.html"
Write-Host "  └── Excel2ZugferdSetupPayload\"
foreach ($item in $items) {
    $src = Join-Path $root $item
    if (Test-Path $src -PathType Container) {
        Write-Host "      ├── $item\"
    }
    elseif (Test-Path $src) {
        Write-Host "      ├── $item"
    }
}
Write-Host ""
Write-Host "Zum Verteilen den gesamten Ordner Release\ weitergeben."
Write-Host "Endanwender starten: Excel2ZugferdSetup.bat (Doppelklick)"
