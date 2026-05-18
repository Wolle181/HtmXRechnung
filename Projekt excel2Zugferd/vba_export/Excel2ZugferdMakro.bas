Attribute VB_Name = "Excel2ZugferdMakro"
Option Explicit

' API-Funktion zum Erstellen tiefer Pfadstrukturen
#If VBA7 Then
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#Else
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
#End If

' Pfad zum Verzeichnis mit excel2zugferd.exe
' ".\" = gleiches Verzeichnis wie die geoeffnete Excel-Datei (Standard)
' Fuer absoluten Pfad: z.B. "C:\Tools\Excel2ZUGFeRD\"
'Const E2ZPFAD As String = ".\"  => Dan weiter unten relativen Pfad zu absolutem machen:  exePfad = ActiveWorkbook.Path & "\" & E2ZPFAD & "excel2zugferd.exe"
Const E2ZPFAD As String = "C:\Rechnungen\Excel2Zugferd\"

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

    ' Kleine Pruefung, ob Inhalt des aktuellen Sheets ueberhaupt geeignet ist fuer eine XL2Zugferd-Rechnung
    If ActiveSheet.Range("A1").Value <> "An:" Then
        Application.Cursor = xlDefault
        MsgBox "Das aktuelle Sheet scheint keinen Excel2Zugferd-Inhalt zu haben!"
        Exit Sub
    End If

    ' Vollstaendiger Pfad inkl. Dateiendung der geoeffneten Excel-Datei
    excelDateiPfad = ActiveWorkbook.FullName

    ' exe-Pfad: E2ZPFAD relativ zum Verzeichnis der Excel-Datei
    exePfad = E2ZPFAD & "excel2zugferd.exe"
    If CreateDeepPath(E2ZPFAD) Then
        Debug.Print "Erfolg: Pfad ist bereit."
    Else
        Application.Cursor = xlDefault
        MsgBox "Fehler: Pfad konnte nicht erstellt werden.", vbCritical
        Exit Sub
    End If

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


Public Function CreateDeepPath(ByVal ZielPfad As String) As Boolean
    ' Die API-Funktion benoetigt zwingend einen abschliessenden Backslash
    If Right(ZielPfad, 1) <> "\" Then ZielPfad = ZielPfad & "\"

    ' Rueckgabewert der API ist 1 bei Erfolg, 0 bei Fehler
    If MakeSureDirectoryPathExists(ZielPfad) <> 0 Then
        CreateDeepPath = True
    Else
        CreateDeepPath = False
    End If
End Function

' getImage-Callback fuer den Ribbon-Button: laedt horse.bmp aus dem AddIn-Verzeichnis
Public Sub GetHorseImage(control As IRibbonControl, ByRef returnedVal)
    Dim imgPath As String
    imgPath = ThisWorkbook.Path & "\horse.bmp"
    On Error Resume Next
    If Dir(imgPath) <> "" Then
        Set returnedVal = LoadPicture(imgPath)
    End If
End Sub

