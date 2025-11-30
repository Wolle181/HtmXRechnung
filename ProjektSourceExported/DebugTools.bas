Attribute VB_Name = "DebugTools"
Option Explicit

Public Sub ExportAllVbaModules()
    Dim vbComp As Object
    Dim exportPath As String

    ' Zielordner festlegen
    exportPath = ThisWorkbook.Path & "\vba_export" & Format(Now, "_yyyymmdd_hhnnss") & "\"
    MkDir exportPath

    On Error GoTo Fehler

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 ' 1=Standardmodul, 2=Klassenmodul, 3=UserForm
                vbComp.Export exportPath & vbComp.name & "." & GetFileExtension(vbComp.Type)
        End Select
    Next vbComp

    MsgBox "Export abgeschlossen nach: " & vbCrLf & exportPath
    On Error GoTo 0
    Exit Sub
    
Fehler:
    MsgBox "Export nicht möglich. Wurde " & vbCrLf & "'Zugriff auf das VBA-Projektobjektmodell vertrauen'" & vbCrLf & "zugelassen?"
End Sub

Private Function GetFileExtension(typeId As Integer) As String
    Select Case typeId
        Case 1: GetFileExtension = "bas"
        Case 2: GetFileExtension = "cls"
        Case 3: GetFileExtension = "frm"
        Case Else: GetFileExtension = "txt"
    End Select
End Function

Public Sub SettingsUnprotect()
    Dim ws As Worksheet: ActivateAnySheet ws, "Settings"
    Dim pw As String: pw = Settings.GetPassword(ws)

    ws.visible = xlSheetVisible
    ws.Columns(SETTINGS_PROTECTED_COLUMNS).Hidden = False
    ws.Unprotect pw
End Sub

Public Sub ResetCursor()
    Application.Cursor = xlDefault
End Sub
