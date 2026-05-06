' Excel2ZUGFeRD Add-In Module
' Button callback um ZUGFeRD PDF zu erstellen

Option Explicit

Sub CreateZugFeRDPDF(control As Object)
    ' Onclick Handler für "ZgFeRD pdf erstellen" Button
    
    Const EXCEL2ZUGFERD_EXE = "C:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\excel2zugferd.exe"
    Const BLATT_NR = "0"  ' Erstes Tabellenblatt (0-indexiert)
    
    Dim ws As Worksheet
    Dim excelFilePath As String
    Dim shell As Object
    Dim command As String
    Dim result As Long
    
    ' Sicherheitscheck
    On Error Resume Next
    
    ' Get active workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "Bitte öffnen Sie zuerst eine Excel-Datei!", vbExclamation, "ZUGFeRD PDF erstellen"
        Exit Sub
    End If
    
    ' Workbook muss gespeichert sein
    If ActiveWorkbook.Saved = False Then
        If MsgBox("Die Datei wurde noch nicht gespeichert. Jetzt speichern?", vbQuestion + vbYesNo, "ZUGFeRD PDF erstellen") = vbYes Then
            ActiveWorkbook.Save
        Else
            Exit Sub
        End If
    End If
    
    ' Check ob EXE existiert
    If Dir(EXCEL2ZUGFERD_EXE) = "" Then
        MsgBox "Die Datei '" & EXCEL2ZUGFERD_EXE & "' wurde nicht gefunden!" & vbCrLf & _
               "Bitte installieren Sie Excel2ZUGFeRD zuerst." & vbCrLf & _
               "Download: https://github.com/Lkammer/excel2zugferd/releases", _
               vbCritical, "ZUGFeRD PDF erstellen"
        Exit Sub
    End If
    
    excelFilePath = ActiveWorkbook.FullName
    
    ' Command zusammenstellen
    command = """" & EXCEL2ZUGFERD_EXE & """ -" & BLATT_NR & " """ & excelFilePath & """"
    
    ' Shell Objekt erstellen
    Set shell = CreateObject("WScript.Shell")
    
    ' Command ausführen (hidden window)
    On Error Resume Next
    result = shell.Run(command, 0, True)
    On Error GoTo 0
    
    If result = 0 Then
        MsgBox "ZUGFeRD PDF erfolgreich erstellt! Die Datei finden Sie im selben Verzeichnis wie diese Excel-Datei.", _
               vbInformation, "ZUGFeRD PDF erstellen"
    Else
        MsgBox "Ein Fehler ist aufgetreten. Bitte überprüfen Sie die Windows Ereignisanzeige für Details.", _
               vbCritical, "ZUGFeRD PDF erstellen"
    End If
    
    Set shell = Nothing
    
End Sub

Sub CreateZugFeRDPDFWithCustomSheet()
    ' Alternative: Mit benutzerdefinierten Blattindizes
    ' Kann erweitert werden für Multi-Sheet Auswahl
    Const EXCEL2ZUGFERD_EXE = "C:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\excel2zugferd.exe"
    
    Dim ws As Worksheet
    Dim sheetIndex As Integer
    Dim shell As Object
    
    On Error Resume Next
    
    If ActiveWorkbook Is Nothing Then Exit Sub
    
    ' Get aktuelle Blattposition (0-indexiert)
    sheetIndex = ActiveSheet.Index - 1
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run """" & EXCEL2ZUGFERD_EXE & """ -" & CStr(sheetIndex) & " """ & ActiveWorkbook.FullName & """", 0, False
    Set shell = Nothing
    
End Sub
