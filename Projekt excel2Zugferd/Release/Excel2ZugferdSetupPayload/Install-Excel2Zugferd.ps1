#Requires -Version 5.0
# Installiert Excel2Zugferd.xlam als Excel-AddIn fuer den aktuellen Benutzer.
# Aufruf: powershell -ExecutionPolicy Bypass -File Install-Excel2Zugferd.ps1

param(
    [string]$XlamSource = "$PSScriptRoot\Excel2Zugferd.xlam"
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
function Register-ViaRegistry([string]$addinPath) {
    # Office-Version erkennen: HKCU-Schluessel werden sowohl bei MSI-
    # als auch bei ClickToRun-Installationen angelegt.
    $officeKey = $null
    foreach ($ver in @("16.0","17.0","15.0","14.0")) {
        $excelBase = "HKCU:\Software\Microsoft\Office\$ver\Excel"
        if (Test-Path $excelBase) {
            $candidate = "$excelBase\Options"
            if (-not (Test-Path $candidate)) {
                New-Item $candidate -Force | Out-Null
            }
            $officeKey = $candidate
            Write-Host "  Erkannte Office-Version: $ver" -ForegroundColor Green
            break
        }
    }

    if (-not $officeKey) {
        Write-Host "FEHLER: Keine unterstuetzte Office-Version gefunden." -ForegroundColor Red
        exit 1
    }

    # Naechsten freien OPEN-Schluessel finden (OPEN, OPEN1, OPEN2, ...)
    $valueName = "OPEN"
    $i = 1
    while ($true) {
        $val = Get-ItemProperty -Path $officeKey -Name $valueName -ErrorAction SilentlyContinue
        if (-not $val) { break }
        if ($val.$valueName -like "*Excel2Zugferd*") {
            Write-Host "  AddIn war bereits in der Registry eingetragen." -ForegroundColor Green
            return
        }
        $valueName = "OPEN$i"; $i++
    }

    Set-ItemProperty -Path $officeKey -Name $valueName -Value "/R `"$addinPath`""
    Write-Host "  Registry-Eintrag gesetzt ($valueName)." -ForegroundColor Green
}

function Register-ViaCOM([string]$addinPath) {
    $excel = $null
    try {
        $excel               = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false

        $existing = $excel.AddIns | Where-Object { $_.FullName -eq $addinPath }
        if ($existing) {
            $existing.Installed = $true
        } else {
            $addin           = $excel.AddIns.Add($addinPath, $false)
            $addin.Installed = $true
        }
        Write-Host "  AddIn per COM registriert und aktiviert." -ForegroundColor Green
    } finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
        }
    }
}

# ---------------------------------------------------------------------------
Write-Host "=== Excel2ZUGFeRD AddIn Installation ===" -ForegroundColor Cyan

# Quelldatei pruefen
if (-not (Test-Path $XlamSource)) {
    Write-Host "FEHLER: Datei nicht gefunden: $XlamSource" -ForegroundColor Red
    Write-Host "Install.bat und Excel2Zugferd.xlam muessen im gleichen Verzeichnis liegen." -ForegroundColor Red
    exit 1
}

# XLAM in den Excel-AddIns-Ordner kopieren
$targetDir  = "$env:APPDATA\Microsoft\AddIns"
$targetPath = Join-Path $targetDir "Excel2Zugferd.xlam"

if (-not (Test-Path $targetDir)) { New-Item $targetDir -ItemType Directory | Out-Null }

Write-Host "Kopiere AddIn nach: $targetPath"
Copy-Item $XlamSource $targetPath -Force
Write-Host "  Kopiert." -ForegroundColor Green

# Registrieren: zuerst COM, bei Fehler Registry-Fallback
Write-Host "Registriere AddIn in Excel..."
$registered = $false

try {
    Register-ViaCOM $targetPath
    $registered = $true
} catch {
    Write-Host "  COM-Zugriff nicht moeglich (Excel evtl. schon offen), nutze Registry-Fallback..." -ForegroundColor Yellow
}

if (-not $registered) {
    Register-ViaRegistry $targetPath
}

Write-Host ""
Write-Host "=== Installation abgeschlossen! ===" -ForegroundColor Cyan
Write-Host "Der Tab 'Excel2ZUGFeRD' erscheint beim naechsten Excel-Start im Ribbon."
