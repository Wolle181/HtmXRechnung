Attribute VB_Name = "Dialogfunktionen"
Option Explicit

Public FormExitCode As String

' Aufruf aus RibbonControl Menü (daher Parameter notwendig, auch wenn ungenutzt)
Public Sub StartAbrechnung(control As IRibbonControl)
    If False And ENTWICKLERMODE Then
        If Utils.DoesAbrSheetExist() Then
            Abrechnung.SchreibeAbgerechnetZurueck
            Exit Sub
        End If
    
        Call Abrechnung.SucheNichtAbgerechnetePositionen("Bender", True)
        Exit Sub
    End If

    FormExitCode = ""
    AbrechnungDialog_Show
    If FormExitCode = EXITCODE_WRITEBACK Then
        Abrechnung.SchreibeAbgerechnetZurueck
    End If
End Sub

Public Sub ShowSettings(control As IRibbonControl)
    Settings.SheetVisibility True
End Sub

Private Sub AbrechnungDialog_Show()
    On Error GoTo EH
    Wartebox.ShowToast "Einlesen der Mandanten aller Zeiterfassungen"
    
    Load AbrechnungDialog
    With AbrechnungDialog
        .Caption = "Suche Mandant anhand Nummer oder Name"

        Dim StartWithNameField As Boolean: StartWithNameField = Settings.GetDialogStartsWithNameField()
        .optMdId.Value = Not StartWithNameField
        .optName.Value = StartWithNameField
        
        ' Namen aus allen WORKSHEET_PREFIX_TO_COLLECT-Sheets, Spaltenüberschrift "MD", disjunkt laden
        Dim items As Collection: Set items = Helper.CollectUniqueMdNames()
        If Not items Is Nothing And items.Count > 0 Then
            Dim names() As String
            names = Utils.CollectionToStringArray(items)
            Utils.QuickSortStrings names, LBound(names), UBound(names)
        
            Dim arr() As Variant, i As Long
            ReDim arr(0 To UBound(names), 0 To 0)
            For i = LBound(names) To UBound(names)
                arr(i, 0) = names(i)
            Next
        
            Dim AbrSheetExists As Boolean: AbrSheetExists = Utils.DoesAbrSheetExist()
        
            .lstName.ColumnCount = 1
            .lstName.List = arr
            .btnSearch.Enabled = Not AbrSheetExists
            .btnWriteBack.Enabled = AbrSheetExists
            
            .lstName.Enabled = Not AbrSheetExists
            .txtMdId.Enabled = Not AbrSheetExists
            .optMdId.Enabled = Not AbrSheetExists
            .optName.Enabled = Not AbrSheetExists
        End If
    End With

    Wartebox.CloseToast
    AbrechnungDialog.Show
    Exit Sub
    
EH:
    Wartebox.CloseToast
    MsgBox "Fehler beim Öffnen des Dialogs: " & Err.Description, vbExclamation
End Sub


