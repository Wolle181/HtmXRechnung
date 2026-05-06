# Build & Deploy Script für Excel2ZugFeRD
#
# Verwendung:
# 1. PowerShell als Administrator öffnen
# 2. cd C:\Users\Charis\Projekte\excel2zugferd
# 3. .\build_and_deploy.ps1

param(
    [string]$ProjectDir = "C:\Users\Charis\Projekte\excel2zugferd",
    [string]$TargetDir = "C:\Program Files (x86)\Excel2ZUGFeRD",
    [string]$SpecFile = "Excel2ZugFeRD.spec",
    [string]$PythonExe = "C:\Users\Charis\AppData\Roaming\uv\python\cpython-3.14.3-windows-x86_64-none\python.exe",
    [switch]$SkipInstallDeps
)

Set-Location -Path $ProjectDir

if (-not $SkipInstallDeps) {
    Write-Host "=== Abhängigkeiten installieren ===" -ForegroundColor Cyan
    & "$PythonExe" -m pip install --break-system-packages --upgrade pyinstaller pillow pandas openpyxl fpdf2 qrcode drafthorse pywin32
}

Write-Host "=== Beende ggf. laufende excel2zugferd Prozesse ===" -ForegroundColor Cyan
try {
    Get-Process -Name "excel2zugferd" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
} catch {
    Write-Host "Keine laufenden excel2zugferd Prozesse oder Stop fehlgeschlagen" -ForegroundColor Yellow
}

Write-Host "=== Clean PyInstaller Temp ===" -ForegroundColor Cyan
Remove-Item -Force -Recurse .\build, .\dist -ErrorAction SilentlyContinue

Write-Host "=== EXE neu bauen mittels PyInstaller ===" -ForegroundColor Cyan
& "$PythonExe" -m PyInstaller --clean --noconfirm -y `
    -n excel2zugferd `
    --paths="$ProjectDir" `
    --collect-all encodings `
    --collect-all drafthorse `
    --collect-all src `
    --add-data="_internal/Fonts;Fonts" `
    --hidden-import=openpyxl `
    --hidden-import=openpyxl.workbook `
    --hidden-import=openpyxl.worksheet `
    --hidden-import=xlrd `
    --hidden-import=et_xmlfile `
    --hidden-import=lxml `
    --hidden-import=lxml.etree `
    --hidden-import=pandas `
    --hidden-import=pandas._libs.tslibs.np_datetime `
    --hidden-import=pandas._libs.tslibs.nattype `
    --hidden-import=numpy `
    --hidden-import=tkinter `
    --hidden-import=tkinter.messagebox `
    --hidden-import=tkinter.filedialog `
    --hidden-import=PIL `
    --hidden-import=PIL.Image `
    --hidden-import=drafthorse `
    --hidden-import=pypdf `
    --hidden-import=src `
    --hidden-import=src.handle_pdf `
    --hidden-import=src.handle_zugferd `
    --hidden-import=src.handle_girocode `
    --hidden-import=src.handle_ini_file `
    --hidden-import=src.middleware `
    --hidden-import=src.steuerung `
    --hidden-import=src.oberflaeche_base `
    --hidden-import=src.oberflaeche_excel2zugferd `
    --hidden-import=src.oberflaeche_ini `
    --hidden-import=src.oberflaeche_steuerung `
    --hidden-import=src.oberflaeche_excelpositions `
    --hidden-import=src.oberflaeche_excelsteuerung `
    --hidden-import=src.invoice_collection `
    --hidden-import=src.invoice `
    --hidden-import=src.excel_content `
    --hidden-import=src.adresse `
    --hidden-import=src.constants `
    --hidden-import=src.konto `
    --hidden-import=src.kunde `
    --hidden-import=src.lieferant `
    --hidden-import=src.stammdaten `
    --hidden-import=src.windowseventlog `
    excel2zugferd.py

$distPath = Join-Path -Path $ProjectDir -ChildPath "dist\excel2zugferd\excel2zugferd.exe"
if (-not (Test-Path $distPath)) {
    Write-Error "Build fehlgeschlagen: $distPath nicht gefunden"
    exit 1
}

Write-Host "=== Zielordner kopieren ===" -ForegroundColor Cyan
if (-not (Test-Path $TargetDir)) {
    New-Item -Path $TargetDir -ItemType Directory -Force
}

$distFolder = Join-Path -Path $ProjectDir -ChildPath "dist\excel2zugferd"
try {
    Write-Host "Zielordner wird aktualisiert..." -ForegroundColor Yellow
    try {
        Stop-Process -Name "excel2zugferd" -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Kein laufender excel2zugferd Prozess gefunden oder Stop fehlgeschlagen" -ForegroundColor Yellow
    }
    # Alten _internal-Ordner entfernen, damit keine veralteten Dateien bleiben
    $oldInternal = Join-Path $TargetDir "_internal"
    if (Test-Path $oldInternal) {
        Remove-Item -Path $oldInternal -Recurse -Force -ErrorAction Stop
    }
    # Gesamten dist-Ordner (exe + _internal) ins Zielverzeichnis kopieren
    Copy-Item -Path "$distFolder\*" -Destination $TargetDir -Recurse -Force -ErrorAction Stop
} catch {
    Write-Error "Kopieren fehlgeschlagen: $_"
    Write-Host "Wenn die Datei noch geöffnet ist, schließe sie und starte das Script erneut." -ForegroundColor Red
    Write-Host "(Oder beende Excel2ZUGFeRD mit Task-Manager, dann Skript erneut ausführen)" -ForegroundColor Red
    exit 1
}

Write-Host "=== Fertig ✅ ===" -ForegroundColor Green
Write-Host "Neue EXE liegt jetzt in: $TargetDir"
