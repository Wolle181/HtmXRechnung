Attribute VB_Name = "Modul1"
Public Sub SortByMdNr()
    Dim ws As Worksheet: Utils.ActivateAnySheet ws, "Abgleich", False
    Dim lastUsedRow As Long: lastUsedRow = Utils.FindLastUsedRow(ws)

    Dim rDest As Range
    Set rDest = ws.Range(ws.Cells(1, 1), ws.Cells(2, lastUsedRow))
    
    Dim rDestOhneHeader As Range
    Set rDestOhneHeader = ws.Range(ws.Cells(2, 1), ws.Cells(2, lastUsedRow))
    
    rDest.Select
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 key:=rDest, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange rDestOhneHeader
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

