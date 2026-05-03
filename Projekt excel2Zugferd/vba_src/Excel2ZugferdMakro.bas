'Attribute VB_Name = "Excel2ZugferdMakro"
Option Explicit

' Pfad zum Verzeichnis mit excel2zugferd.exe
' ".\" = gleiches Verzeichnis wie die geoeffnete Excel-Datei (Standard)
' Fuer absoluten Pfad: z.B. "C:\Tools\Excel2ZUGFeRD\"
Const E2ZPFAD As String = ".\"

' Direkt per Alt+F8 aufrufbar; onAction im Ribbon zeigt direkt auf diese Sub
Public Sub RunMake(Optional control As IRibbonControl = Nothing)
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
    sheetName = ActiveSheet.Name

    ' Kleine Prüfung, ob Inhalt des aktuellen Sheets überhaupt geeignet ist für eix XL2Zugferd-Rechnung
    If ActiveSheet.Range("A1").Value <> "An:" Then
        Application.Cursor = xlDefault
        MsgBox "Das aktuelle Sheet scheint keinen Excel2Zugferd-Inhalt zu haben!"
        Exit Sub
    End If

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
    MsgBox "ZUGFeRD-Rechnung f?r Tabellenblatt """ & sheetName & """ wurde erzeugt.", _
           vbInformation, "Excel2ZUGFeRD"
    Exit Sub

ErrHandler:
    Application.Cursor = xlDefault
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, _
           vbCritical, "Excel2ZUGFeRD"
End Sub




