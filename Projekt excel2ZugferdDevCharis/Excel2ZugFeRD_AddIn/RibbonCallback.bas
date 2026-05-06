' Excel2ZugFeRD Add-In mit programmatischem Button
' Der Button wird beim Laden des Add-Ins automatisch ins Menueband eingefuegt

Option Explicit

' Globale Variable fuer die Ribbon Referenz
Public g_ribbon As Object

' OnLoad-Callback (wird aufgerufen, wenn das Add-In geladen wird)
Sub OnLoad(ribbon As Object)
    Set g_ribbon = ribbon
    ' Der Button wird durch customUI definiert
End Sub

' Haupt-Funktion: ZUGFeRD PDF erstellen
Sub CreateZugFeRDPDF(control As Object)
    Const EXE_PATH = "c:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\excel2zugferd.exe"
    
    Dim ws As Object
    Dim exePath As String
    Dim shell As Object
    Dim result As Long
    
    On Error Resume Next
    
    ' Sicherheitscheck
    If ActiveWorkbook Is Nothing Then
        MsgBox "Bitte oeffnen Sie zuerst eine Excel-Datei!", vbExclamation, "ZgFeRD pdf erstellen"
        Exit Sub
    End If
    
    ' Arbeitsmappe speichern
    If ActiveWorkbook.Saved = False Then
        If MsgBox("Die Datei wurde noch nicht gespeichert. Jetzt speichern?", vbQuestion + vbYesNo, "ZgFeRD pdf erstellen") = vbYes Then
            ActiveWorkbook.Save
        Else
            Exit Sub
        End If
    End If
    
    ' Check EXE
    If Dir(EXE_PATH) = "" Then
        MsgBox "Excel2ZugFeRD nicht gefunden: " & EXE_PATH, vbCritical, "Fehler"
        Exit Sub
    End If
    
    ' Starte Prozess
    exePath = ActiveWorkbook.FullName
    Set shell = CreateObject("WScript.Shell")
    result = shell.Run("""" & EXE_PATH & """ -0 """ & exePath & """", 0, True)
    
    If result = 0 Then
        MsgBox "Erfolg! PDF erstellt." & vbCrLf & ActiveWorkbook.Path, vbInformation, "ZgFeRD pdf erstellen"
    Else
        MsgBox "Fehler beim Erstellen. Code: " & result, vbCritical, "ZgFeRD pdf erstellen"
    End If
    
    Set shell = Nothing
End Sub
