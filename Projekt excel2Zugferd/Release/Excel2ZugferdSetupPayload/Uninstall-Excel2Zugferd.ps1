#Requires -Version 5.0
# Deinstalliert Excel2Zugferd.xlam vollstaendig:
#   1. Deaktivierung per COM (falls Excel geschlossen werden kann)
#   2. Registry-Eintraege (OPEN/OPENx, Add-in Manager, AddInLoadTimes) entfernen
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

function Remove-RegistryValuesByName([string]$keyPath) {
    # Loescht alle Werte in keyPath, deren *Name* "Excel2Zugferd" enthaelt.
    if (-not (Test-Path $keyPath)) { return $false }
    $props = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue
    if (-not $props) { return $false }
    $found = $false
    $props.PSObject.Properties |
        Where-Object { $_.Name -like "*Excel2Zugferd*" } |
        ForEach-Object {
            Remove-ItemProperty -Path $keyPath -Name $_.Name -ErrorAction SilentlyContinue
            Write-Host "    Entfernt: $($_.Name)" -ForegroundColor Green
            $found = $true
        }
    return $found
}

function Unregister-ViaRegistry() {
    foreach ($ver in @("16.0","17.0","15.0","14.0")) {
        $excelBase = "HKCU:\Software\Microsoft\Office\$ver\Excel"
        if (-not (Test-Path $excelBase)) { continue }

        Write-Host "  Office $ver gefunden – bereinige Keys..." -ForegroundColor Cyan

        # OPEN/OPENx – Auto-Lade-Eintraege
        $optKey = "$excelBase\Options"
        if (Test-Path $optKey) {
            $props = Get-ItemProperty -Path $optKey -ErrorAction SilentlyContinue
            $props.PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' -and $_.Value -like "*Excel2Zugferd*" } |
                ForEach-Object {
                    Remove-ItemProperty -Path $optKey -Name $_.Name -ErrorAction SilentlyContinue
                    Write-Host "    Options\$($_.Name) entfernt." -ForegroundColor Green
                }
        }

        # Add-in Manager – die sichtbare AddIn-Liste im Dialog
        Write-Host "  Add-in Manager..." -NoNewline
        if (Remove-RegistryValuesByName "$excelBase\Add-in Manager") {
            # Bereits ausgegeben
        } else {
            Write-Host " nichts gefunden." -ForegroundColor Yellow
        }

        # AddInLoadTimes – Lade-Timing-Cache
        Write-Host "  AddInLoadTimes..." -NoNewline
        if (Remove-RegistryValuesByName "$excelBase\AddInLoadTimes") {
            # Bereits ausgegeben
        } else {
            Write-Host " nichts gefunden." -ForegroundColor Yellow
        }
    }
}

# ---------------------------------------------------------------------------
Write-Host "=== Excel2ZUGFeRD AddIn Deinstallation ===" -ForegroundColor Cyan
Write-Host "Bitte sicherstellen, dass Excel geschlossen ist."
Write-Host ""

# 1. COM-Deaktivierung
Write-Host "1. Deaktiviere AddIn in Excel..."
try {
    Unregister-ViaCOM $targetPath
} catch {
    Write-Host "  COM-Zugriff nicht moeglich (Excel evtl. offen), ueberspringe COM-Schritt." -ForegroundColor Yellow
}

# 2. Registry bereinigen
Write-Host ""
Write-Host "2. Bereinige Registry-Eintraege..."
Unregister-ViaRegistry

# 3. XLAM-Datei loeschen
Write-Host ""
Write-Host "3. Loesche AddIn-Datei(en)..."
$deleted = $false
foreach ($path in @(
    $targetPath,
    "C:\WORK\$addinName"
)) {
    if (Test-Path $path) {
        Remove-Item $path -Force
        Write-Host "  Geloescht: $path" -ForegroundColor Green
        $deleted = $true
    }
}
if (-not $deleted) {
    Write-Host "  Keine XLAM-Datei gefunden." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== Deinstallation abgeschlossen! ===" -ForegroundColor Cyan
Write-Host "Excel2ZUGFeRD ist vollstaendig entfernt. Excel neu starten um zu bestaetigen."
