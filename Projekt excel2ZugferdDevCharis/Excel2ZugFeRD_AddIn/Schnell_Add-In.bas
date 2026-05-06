' ════════════════════════════════════════════════════════════════════════════
' SCHNELL-ANLEITUNG: Excel Add-In in 5 Minuten erstellen
' ════════════════════════════════════════════════════════════════════════════
'
' SCHRITT 1: Diese Datei-Liste in Excel öffnen
' SCHRITT 2: Menü Entwickler → Code-Editor (Alt+F11)
' SCHRITT 3: Module → neues Modul hinzufügen
' SCHRITT 4: Kompletten VBA-Code unten einfügen
' SCHRITT 5: Datei → Speichern unter → Dateityp: Excel-Add-In (*.xlam)
'            Name: "Excel2ZugFeRD_AddIn"
'            Speicherort: C:\Users\[USERNAME]\AppData\Roaming\Microsoft\AddIns\
' SCHRITT 6: Excel neu starten
' SCHRITT 7: Datei → Optionen → Add-Ins → Verwalten → Add-In durchsuchen
'
' ════════════════════════════════════════════════════════════════════════════

Option Explicit

' ────────────────────────────────────────────────────────────────────────────
' HAUPTFUNKTION: ZUGFeRD PDF erstellen
' ────────────────────────────────────────────────────────────────────────────
Sub CreateZugFeRDPDF()
    '
    ' Konfiguration
    '
    Const EXCEL2ZUGFERD_EXE As String = "C:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\excel2zugferd.exe"
    Const BLATT_NR As String = "0"  ' 0 = erstes Blatt, 1 = zweites Blatt, etc.
    
    Dim shell As Object
    Dim excelFilePath As String
    Dim command As String
    Dim result As Long
    Dim msg As String
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 1. Fehlerbehandlung aktivieren
    ' ─────────────────────────────────────────────────────────────────────────
    On Error Resume Next
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 2. Sicherheitscheck: Ist eine Arbeitsmappe offen?
    ' ─────────────────────────────────────────────────────────────────────────
    If ActiveWorkbook Is Nothing Then
        MsgBox "Bitte öffnen Sie zuerst eine Excel-Datei!", vbExclamation, "ZgFeRD pdf erstellen"
        Exit Sub
    End If
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 3. Arbeitsmappe speichern (notwendig für ZugFeRD-Verarbeitung)
    ' ─────────────────────────────────────────────────────────────────────────
    If ActiveWorkbook.Saved = False Then
        If MsgBox("Die Datei wurde noch nicht gespeichert." & vbCrLf & _
                  "Jetzt speichern?", vbQuestion + vbYesNo, "ZgFeRD pdf erstellen") = vbYes Then
            ActiveWorkbook.Save
            If Err.Number <> 0 Then
                MsgBox "Fehler beim Speichern: " & Err.Description, vbCritical
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 4. Überprüfe ob Excel2ZugFeRD installiert ist
    ' ─────────────────────────────────────────────────────────────────────────
    If Dir(EXCEL2ZUGFERD_EXE) = "" Then
        msg = "Die Anwendung Excel2ZugFeRD wurde nicht gefunden!" & vbCrLf & _
              "Pfad: " & EXCEL2ZUGFERD_EXE & vbCrLf & vbCrLf & _
              "Bitte installieren Sie Excel2ZugFeRD zuerst:" & vbCrLf & _
              "https://github.com/Lkammer/excel2zugferd/releases"
        MsgBox msg, vbCritical, "Fehler: Anwendung nicht installiert"
        Exit Sub
    End If
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 5. Kommando zusammenstellen
    ' ─────────────────────────────────────────────────────────────────────────
    excelFilePath = ActiveWorkbook.FullName
    command = """" & EXCEL2ZUGFERD_EXE & """ -" & BLATT_NR & " """ & excelFilePath & """"
    
    ' Debug: Optional - in Immediate Window anzeigen (Ctrl+G im Editor)
    ' Debug.Print "Command: " & command
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 6. Ausführen: Shell Objekt erstellen und Kommando ausführen
    ' ─────────────────────────────────────────────────────────────────────────
    Set shell = CreateObject("WScript.Shell")
    On Error Resume Next
    result = shell.Run(command, 0, True)  ' 0=Hidden, True=Warte auf Ende
    On Error GoTo ErrorHandler
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 7. Ergebnis anzeigen
    ' ─────────────────────────────────────────────────────────────────────────
    If result = 0 Then
        msg = "✓ ZUGFeRD PDF erfolgreich erstellt!" & vbCrLf & vbCrLf & _
              "Dateiname: " & ActiveWorkbook.Name & ".pdf" & vbCrLf & _
              "Speicherort: " & ActiveWorkbook.Path
        MsgBox msg, vbInformation, "ZgFeRD pdf erstellen - Erfolg"
    Else
        msg = "✗ Ein Fehler ist aufgetreten (Exit Code: " & result & ")" & vbCrLf & vbCrLf & _
              "Bitte überprüfen Sie:" & vbCrLf & _
              "1. Die Windows Ereignisanzeige (Event Viewer)" & vbCrLf & _
              "2. Dass alle Felder in der Excel-Datei korrekt ausgefüllt sind" & vbCrLf & _
              "3. Die Datei ist im korrekten Format"
        MsgBox msg, vbCritical, "ZgFeRD pdf erstellen - Fehler"
    End If
    
    ' ─────────────────────────────────────────────────────────────────────────
    ' 8. Cleanup
    ' ─────────────────────────────────────────────────────────────────────────
    Set shell = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical
    Set shell = Nothing
End Sub

' ────────────────────────────────────────────────────────────────────────────
' ALTERNATIVE: Mit Blatt-Auswahl
' ────────────────────────────────────────────────────────────────────────────
Sub CreateZugFeRDPDF_CustomSheet()
    ' Diese Funktion nutzt das aktuelle Blatt statt Indexnummer
    
    Const EXCEL2ZUGFERD_EXE As String = "C:\Users\Charis\Projekte\excel2zugferd\dist\excel2zugferd\excel2zugferd.exe"
    
    Dim shell As Object
    Dim sheetIndex As Integer
    Dim command As String
    
    If ActiveWorkbook Is Nothing Then Exit Sub
    If ActiveWorkbook.Saved = False Then ActiveWorkbook.Save
    
    ' Sheets sind 1-indexiert, aber Command-Zeile braucht 0-indexiert
    sheetIndex = ActiveSheet.Index - 1
    
    command = """" & EXCEL2ZUGFERD_EXE & """ -" & CStr(sheetIndex) & " """ & ActiveWorkbook.FullName & """"
    
    Set shell = CreateObject("WScript.Shell")
    shell.Run command, 0, False
    Set shell = Nothing
    
    MsgBox "ZUGFeRD PDF wird erstellt...", vbInformation
End Sub

' ────────────────────────────────────────────────────────────────────────────
' BUTTON AUF SHEET HINZUFÜGEN (im VBA Editor)
' ────────────────────────────────────────────────────────────────────────────
' 
' 1. Im Excel Sheet: Entwickler → Einfügungssteuerelemente → Schaltfläche
' 2. Rechteck zeichnen auf dem Sheet
' 3. Dialog erscheint → Macro auswählen → CreateZugFeRDPDF
' 4. OK
' 5. Button-Text bearbeiten: "ZgFeRD pdf erstellen" eingeben
' 6. Speichern als .xlam
'
' ────────────────────────────────────────────────────────────────────────────
