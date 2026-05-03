#Requires -Version 5.0
# Deinstalliert Excel2Zugferd.xlam vollstaendig:
#   1. Deaktivierung per COM (falls Excel geschlossen werden kann)
#   2. Registry-Eintraege (OPEN/OPENx) entfernen
#   3. XLAM-Datei loeschen

$ErrorActionPreference = "Stop"

$addinName  = "Excel2Zugferd.xlam"
$targetPath = "$env:APPDATA\Microsoft\AddIns\$addinName"

# ---------------------------------------------------------------------------
function Unregister-ViaCOM([string]$addinPath) {
    $excel = $null
    try {
        $excel               = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false

        $existing = $excel.AddIns | Where-Object { $_.FullName -eq $addinPath -or $_.Name -eq $addinName }
        if ($existing) {
            $existing.Installed = $false
            Write-Host "  AddIn per COM deaktiviert." -ForegroundColor Green
        } else {
            Write-Host "  AddIn war in Excel nicht als aktiv registriert." -ForegroundColor Yellow
        }
    } finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            [GC]::Collect()
        }
    }
}

function Unregister-ViaRegistry() {
    $found = $false
    foreach ($ver in @("16.0","17.0","15.0","14.0")) {
        $optionsKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
        if (-not (Test-Path $optionsKey)) { continue }

        $props = Get-ItemProperty -Path $optionsKey -ErrorAction SilentlyContinue
        if (-not $props) { continue }

        $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } | ForEach-Object {
            if ($_.Value -like "*Excel2Zugferd*") {
                Remove-ItemProperty -Path $optionsKey -Name $_.Name -ErrorAction SilentlyContinue
                Write-Host "  Registry-Eintrag entfernt: Office $ver -> $($_.Name) = $($_.Value)" -ForegroundColor Green
                $found = $true
            }
        }
    }
    if (-not $found) {
        Write-Host "  Kein Registry-Eintrag fuer Excel2Zugferd gefunden." -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
Write-Host "=== Excel2ZUGFeRD AddIn Deinstallation ===" -ForegroundColor Cyan

# 1. COM-Deaktivierung
Write-Host "Deaktiviere AddIn in Excel..."
try {
    Unregister-ViaCOM $targetPath
} catch {
    Write-Host "  COM-Zugriff nicht moeglich (Excel evtl. offen), ueberspringe COM-Schritt." -ForegroundColor Yellow
    Write-Host "  Bitte Excel schliessen und danach pruefen, ob das AddIn noch aktiv ist." -ForegroundColor Yellow
}

# 2. Registry bereinigen
Write-Host "Bereinige Registry-Eintraege..."
Unregister-ViaRegistry

# 3. XLAM-Datei loeschen
Write-Host "Loesche AddIn-Datei..."
if (Test-Path $targetPath) {
    Remove-Item $targetPath -Force
    Write-Host "  Datei geloescht: $targetPath" -ForegroundColor Green
} else {
    Write-Host "  Datei nicht vorhanden: $targetPath" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Deinstallation abgeschlossen! ===" -ForegroundColor Cyan
Write-Host "Excel2ZUGFeRD ist vollstaendig entfernt. Bitte Excel neu starten."
