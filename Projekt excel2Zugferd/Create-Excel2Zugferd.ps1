#Requires -Version 5.0
# Erstellt Excel2Zugferd.xlam mit VBA-Makro, Ribbon-Button und Pferd-Icon.

$OutputPath = "C:\WORK\Excel2Zugferd.xlam"

# =============================================================================
# [0/2]  Pferd-Icon erzeugen: Schachspringer U+265E aus System-Font als BMP
# =============================================================================
Write-Host "=== Excel2ZUGFeRD AddIn Generator ===" -ForegroundColor Cyan
Write-Host "`n[0/2] Generiere Pferd-Icon (Schachspringer)..." -ForegroundColor White

Add-Type -AssemblyName System.Drawing

$iconSize   = 32
$bmp        = New-Object System.Drawing.Bitmap($iconSize, $iconSize,
                  [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
$g          = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
$g.Clear([System.Drawing.Color]::White)

$knightChar = [char]0x265E          # ♞ BLACK CHESS KNIGHT

# Font-Suche: bevorzuge Symbol-Fonts mit dem Springer-Zeichen
$usedFontName = $null
foreach ($fname in @("Segoe UI Symbol", "Segoe UI Emoji", "Arial Unicode MS")) {
    try {
        $ff = New-Object System.Drawing.FontFamily($fname)
        $usedFontName = $fname
        break
    } catch { }
}

if ($usedFontName) {
    $font  = New-Object System.Drawing.Font($usedFontName, 26,
                 [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(31, 78, 121))
    $sf    = New-Object System.Drawing.StringFormat
    $sf.Alignment     = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($knightChar, $font, $brush, (New-Object System.Drawing.RectangleF(0,0,32,32)), $sf)
    $font.Dispose(); $brush.Dispose()
    Write-Host "    Springer mit Font '$usedFontName' gerendert." -ForegroundColor Green
} else {
    # Fallback: "E2Z"-Text
    $font  = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(31,78,121))
    $g.DrawString("E2Z", $font, $brush, 2, 11)
    $font.Dispose(); $brush.Dispose()
    Write-Host "    Kein Symbol-Font gefunden, Fallback 'E2Z'." -ForegroundColor Yellow
}
$g.Dispose()

# BMP in MemoryStream -> Base64
$ms      = New-Object System.IO.MemoryStream
$bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Bmp)
$bmp.Dispose()
$iconB64 = [Convert]::ToBase64String($ms.ToArray())
$ms.Dispose()
Write-Host "    $($iconB64.Length) Base64-Zeichen generiert." -ForegroundColor Green

# Base64 in 80-Zeichen-Haeppchen fuer VBA-Stringliterale aufteilen
$b64Lines = [System.Collections.Generic.List[string]]::new()
$b64Lines.Add("    Dim s As String")
for ($i = 0; $i -lt $iconB64.Length; $i += 80) {
    $chunk = $iconB64.Substring($i, [Math]::Min(80, $iconB64.Length - $i))
    $b64Lines.Add("    s = s & `"$chunk`"")
}
$b64Lines.Add("    HorseIconB64 = s")
$horseFuncLines = $b64Lines -join "`r`n"

# =============================================================================
# VBA-Code: statischer Teil (single-quoted here-string -> keine Interpolation)
# =============================================================================
$VBAStatic = @'
Option Explicit

' Pfad zum Verzeichnis mit excel2zugferd.exe
' ".\" = gleiches Verzeichnis wie die geoeffnete Excel-Datei (Standard)
' Fuer absoluten Pfad: z.B. "C:\Tools\Excel2ZUGFeRD\"
Const E2ZPFAD As String = ".\"

' Direkt per Alt+F8 aufrufbar; onAction im Ribbon zeigt direkt auf diese Sub
Public Sub RunMake()
    Dim tabsheetNummer As Long
    Dim sheetName     As String
    Dim excelDateiPfad As String
    Dim exePfad       As String
    Dim befehl        As String
    Dim wsh           As Object

    On Error GoTo ErrHandler
    Application.Cursor = xlWait

    ' Tabsheet-Position (0-basiert: erstes Sheet = 0)
    tabsheetNummer = ActiveSheet.Index - 1
    sheetName      = ActiveSheet.Name

    ' Vollstaendiger Pfad inkl. Dateiendung der geoeffneten Excel-Datei
    excelDateiPfad = ActiveWorkbook.FullName

    ' exe-Pfad: E2ZPFAD relativ zum Verzeichnis der Excel-Datei
    exePfad = ActiveWorkbook.Path & "\" & E2ZPFAD & "excel2zugferd.exe"

    ' Kommandozeile: "exePfad" TABSHEET_NUMMER "EXCELDATEIPFAD"
    befehl = """" & exePfad & """ " & tabsheetNummer & " """ & excelDateiPfad & """"

    Set wsh = CreateObject("WScript.Shell")
    wsh.Run befehl, 0, False
    Set wsh = Nothing

    Application.Cursor = xlDefault
    MsgBox "ZUGFeRD-Rechnung fuer Tabellenblatt """ & sheetName & """ wurde erzeugt.", _
           vbInformation, "Excel2ZUGFeRD"
    Exit Sub

ErrHandler:
    Application.Cursor = xlDefault
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Excel2ZUGFeRD"
End Sub

' Ribbon-Callback: liefert das Icon fuer image="horse"
' Rueckgabe als Object (nicht IPictureDisp) vermeidet Cast-Probleme in manchen Excel-Versionen
Public Function LoadImage(id As String) As Object
    On Error Resume Next
    If id = "horse" Then
        Set LoadImage = BmpFromBase64(HorseIconB64())
    End If
    On Error GoTo 0
End Function

' Dekodiert Base64-BMP -> Tempdatei -> Object (StdPicture via LoadPicture)
Private Function BmpFromBase64(b64 As String) As Object
    Dim xml   As Object
    Dim node  As Object
    Dim bytes() As Byte
    Dim tmp   As String
    Dim f     As Integer

    ' Base64 -> Byte-Array via MSXML (auf jedem Windows verfuegbar)
    Set xml  = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64
    bytes = node.nodeTypedValue

    ' Bytes als BMP in Tempverzeichnis schreiben
    tmp = Environ("TEMP") & "\e2z_icon.bmp"
    f = FreeFile
    Open tmp For Binary As #f
    Put #f, , bytes
    Close #f

    ' Laden und aufraumen
    Dim pic As Object
    Set pic = LoadPicture(tmp)
    Kill tmp
    Set BmpFromBase64 = pic
End Function

'@

# Dynamischer Teil: HorseIconB64() mit eingebettetem Base64
$VBADynamic = "Private Function HorseIconB64() As String`r`n" +
              $horseFuncLines +
              "`r`nEnd Function"

$VBACode = $VBAStatic + "`r`n" + $VBADynamic

# =============================================================================
# Ribbon-XML: loadImage-Callback + image="horse" statt imageMso
# =============================================================================
$CustomUIXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" loadImage="LoadImage">
  <ribbon>
    <tabs>
      <tab id="tabExcel2Zugferd" label="Excel2ZUGFeRD">
        <group id="grpExcel2Zugferd" label="ZUGFeRD">
          <button id="btnMake"
                  label="Excel2Zugferd"
                  image="horse"
                  size="large"
                  onAction="RunMake"
                  screentip="Excel zu ZUGFeRD konvertieren"
                  supertip="Ruft excel2zugferd.exe fuer das aktuelle Tabellenblatt auf." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>'

# =============================================================================
# [1/2]  XLAM via Excel COM erzeugen
# =============================================================================
Write-Host "`n[1/2] Erstelle XLAM via Excel..." -ForegroundColor White

try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Error "Excel konnte nicht gestartet werden."
    exit 1
}
$excel.Visible       = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Add()

$vbaOK = $false
try {
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = "Excel2ZugferdMakro"
    $mod.CodeModule.AddFromString($VBACode)
    $vbaOK = $true
    Write-Host "    VBA-Modul eingefuegt (inkl. LoadImage + HorseIconB64)." -ForegroundColor Green
} catch {
    Write-Host "    WARNUNG: VBA-Projektzugriff verweigert." -ForegroundColor Yellow
    Write-Host "    Datei > Optionen > Sicherheitscenter > Makroeinstellungen >" -ForegroundColor Yellow
    Write-Host "    'Zugriff auf VBA-Projektobjektmodell vertrauen' aktivieren." -ForegroundColor Yellow
}

# Zieldatei evtl. von Excel gesperrt -> zuerst in Tempdatei speichern
$TempPath = $OutputPath + ".new"
if (Test-Path $TempPath) { Remove-Item $TempPath -Force }
$wb.SaveAs($TempPath, 55)
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
Start-Sleep -Milliseconds 800
Write-Host "    XLAM in Tempdatei gespeichert." -ForegroundColor Green

# =============================================================================
# [2/2]  Ribbon-XML in ZIP einbetten
# =============================================================================
Write-Host "`n[2/2] Bettet Ribbon-XML ein..." -ForegroundColor White

Add-Type -AssemblyName System.IO.Compression.FileSystem
$enc = New-Object System.Text.UTF8Encoding($false)
$zip = [System.IO.Compression.ZipFile]::Open($TempPath, [System.IO.Compression.ZipArchiveMode]::Update)

try {
    # [Content_Types].xml
    $ctE = $zip.GetEntry("[Content_Types].xml")
    $r   = New-Object System.IO.StreamReader($ctE.Open(), $enc)
    $ct  = $r.ReadToEnd(); $r.Close()
    if ($ct -notmatch "customUI") {
        $ct = $ct -replace '</Types>',
            '<Override PartName="/customUI/customUI.xml" ContentType="application/xml"/></Types>'
        $ctE.Delete()
        $w = New-Object System.IO.StreamWriter($zip.CreateEntry("[Content_Types].xml").Open(), $enc)
        $w.Write($ct); $w.Flush(); $w.Close()
        Write-Host "    [Content_Types].xml aktualisiert." -ForegroundColor Green
    }

    # _rels/.rels
    $rE  = $zip.GetEntry("_rels/.rels")
    $r   = New-Object System.IO.StreamReader($rE.Open(), $enc)
    $rel = $r.ReadToEnd(); $r.Close()
    if ($rel -notmatch "extensibility") {
        $rel = $rel -replace '</Relationships>',
            '<Relationship Id="rIdUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/></Relationships>'
        $rE.Delete()
        $w = New-Object System.IO.StreamWriter($zip.CreateEntry("_rels/.rels").Open(), $enc)
        $w.Write($rel); $w.Flush(); $w.Close()
        Write-Host "    _rels/.rels aktualisiert." -ForegroundColor Green
    }

    # customUI/customUI.xml
    $ex = $zip.GetEntry("customUI/customUI.xml")
    if ($ex) { $ex.Delete() }
    $w = New-Object System.IO.StreamWriter($zip.CreateEntry("customUI/customUI.xml").Open(), $enc)
    $w.Write($CustomUIXml); $w.Flush(); $w.Close()
    Write-Host "    customUI/customUI.xml angelegt." -ForegroundColor Green

} finally {
    $zip.Dispose()
}

# Tempdatei -> finale XLAM (evtl. gesperrte Original-Datei ersetzen)
$replaced = $false
try {
    if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
    Move-Item $TempPath $OutputPath
    $replaced = $true
    Write-Host "    Finale XLAM erstellt: $OutputPath" -ForegroundColor Green
} catch {
    Write-Host "    HINWEIS: Originaldatei ist gesperrt (Excel hat sie offen)." -ForegroundColor Yellow
    Write-Host "    Neue Version liegt unter: $TempPath" -ForegroundColor Yellow
    Write-Host "    -> Excel schliessen, dann '$TempPath' in '$OutputPath' umbenennen." -ForegroundColor Yellow
}

# =============================================================================
Write-Host "`n=== Fertig! ===" -ForegroundColor Cyan
Write-Host "AddIn-Datei: $OutputPath"
Write-Host ""
Write-Host "Installation:" -ForegroundColor White
Write-Host "  1. Excel oeffnen"
Write-Host "  2. Datei > Optionen > Add-Ins"
Write-Host "  3. Verwalten: 'Excel-Add-Ins'  -->  'Gehe zu...'"
Write-Host "  4. 'Durchsuchen' -> $OutputPath auswaehlen"
Write-Host "  5. Haken setzen -> OK"
if (-not $vbaOK) {
    Write-Host ""
    Write-Host "WICHTIG: VBA-Projektzugriff aktivieren + Skript erneut ausfuehren!" -ForegroundColor Red
}
