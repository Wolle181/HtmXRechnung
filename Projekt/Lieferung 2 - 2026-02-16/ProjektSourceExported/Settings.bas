Attribute VB_Name = "Settings"
Option Explicit

Public Const SETTINGS_PROTECTED_COLUMNS = "F:G"
Public Const SETTINGS_STARTZEILE_STDSATZLISTE = 9
Public Const SETTINGS_SPALTE_STDSATZLISTE = 6
Public Const SETTINGS_PWCELL = "G3"

Public Function GetPassword(ws As Worksheet) As String
    GetPassword = ws.Range(SETTINGS_PWCELL).Value
End Function

Private Function IsSettingsOpeningAllowed(ByVal pw As String) As Boolean
    Dim ws As Worksheet: ActivateAnySheet ws, "Settings", False
    
    IsSettingsOpeningAllowed = (pw = GetPassword(ws)) Or Utils.SheetExists("CrazyWolle19.12.") Or ENTWICKLERMODE
End Function

Public Sub SettingsSpaltenschutzAufheben()
    Dim ws As Worksheet: ActivateAnySheet ws, "Settings"
    Dim pw As String: pw = GetPassword(ws)

    Dim DoUnprotect As Boolean: DoUnprotect = False
    DoUnprotect = IsSettingsOpeningAllowed("")

    If Not DoUnprotect Then
        ' Passwort abfragen
        Dim eingabe As String
        eingabe = InputBox("Bitte Passwort eingeben zur Anzeige der erweiterten Einstellungen:")

        DoUnprotect = IsSettingsOpeningAllowed(eingabe)
    End If

    If DoUnprotect Then
        ws.Columns(SETTINGS_PROTECTED_COLUMNS).Hidden = False
        ws.Unprotect pw
    Else
        MsgBox "Falsches Passwort!", vbCritical
    End If
End Sub

Public Sub SpaltenSchutzStarten()
    Dim eingabe As String
    Dim ws As Worksheet: ActivateAnySheet ws, "Settings"
    
    ' Passwort auslesen
    Dim pw As String: pw = GetPassword(ws)
    If pw = "" Then
        MsgBox "Kein Passwort gesetzt - Kein Spaltenschutz möglich", vbCritical
        Exit Sub
    End If

    ws.Protect Password:=pw, UserInterfaceOnly:=True
    ws.Columns("F:G").Hidden = True
    ws.Range("A1").Select ' deselect
End Sub

Public Sub HideAllSettingItems()
    SheetVisibility False
    SpaltenSchutzStarten
End Sub

Public Sub SheetVisibility(visible As Boolean)
    Dim ws As Worksheet: ActivateAnySheet ws, "Settings"
    If visible Then
        ws.visible = xlSheetVisible
    Else
        ws.visible = xlSheetVeryHidden
    End If
End Sub

Public Function GetDialogStartsWithNameField() As Boolean
    GetDialogStartsWithNameField = UCase$(ThisWorkbook.Sheets("Settings").Range("B4").Value) = "J"
End Function

Public Function GetTimesheetNameTemplate() As String
    GetTimesheetNameTemplate = ThisWorkbook.Sheets("Settings").Range("B5").Value
End Function

Public Function GetTimesheetBasePath() As String
    Dim wsSettings As Worksheet
    Dim rawPath As String
    Dim resultPath As String
    Dim macroBasePath As String

    On Error Resume Next
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    On Error GoTo 0

    If wsSettings Is Nothing Then
        MsgBox "Settings-Sheet nicht gefunden." & vbCrLf & "Pfad für Zeiterfassungsdateien kann nicht ausgelesen werden.", vbExclamation
        Exit Function
    End If

    rawPath = Trim$(CStr(wsSettings.Range("B3").Value))
    If Len(rawPath) = 0 Then
        MsgBox "In 'Settings'!B3 ist kein Pfad hinterlegt.", vbExclamation
        Exit Function
    End If

    ' Basis-Pfad der Makro-Datei
    macroBasePath = ThisWorkbook.Path

    ' Prüfen ob absolut oder relativ
    ' Absolute Windows-Pfade:
    '   - "C:\..."
    '   - "\\Server\..."
    If Utils.IsAbsolutePath(rawPath) Then
        resultPath = rawPath
    Else
        ' relativer Pfad ? absolut auflösen
        If macroBasePath = "" Then
            MsgBox "RELATIVER Pfad angegeben, aber ThisWorkbook wurde noch nie gespeichert.", vbCritical
            Exit Function
        End If

        ' relativen Teil anhängen
        resultPath = macroBasePath & "\" & rawPath
    End If

    ' Doppelte Backslashes entfernen
    resultPath = Replace(resultPath, "\\", "\")

    ' Falls ein Slash am Ende steht ? entfernen
    If Right$(resultPath, 1) = "\" Then
        resultPath = Left$(resultPath, Len(resultPath) - 1)
    End If

    GetTimesheetBasePath = resultPath
End Function

