Attribute VB_Name = "Utils"
Option Explicit

' --- GUID Typ ---
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' --- API-Deklarationen (32/64-bit) ---
#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID) As Long
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (ByRef rguid As GUID, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID) As Long
    Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByRef rguid As GUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long
#End If

Public Function SheetNameIsMA(ByVal sheetName As String) As Boolean
    Dim prefixToCompareWith As String: prefixToCompareWith = WORKSHEET_PREFIX_TO_COLLECT & "*"
    SheetNameIsMA = LCase$(sheetName) Like LCase$(prefixToCompareWith)
End Function

Public Function DoesAbrSheetExist() As Boolean
    ' Existenz eines ABR-Tempblatts prüfen (mit WORKSHEET_PREFIX_FOR_ABRECHNUNG beginnend)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase$(Left$(ws.name, 4)) = WORKSHEET_PREFIX_FOR_ABRECHNUNG Then
            DoesAbrSheetExist = True
            Exit Function
        End If
    Next
    DoesAbrSheetExist = False
End Function

Public Sub QuickSortStrings(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long, pivot As String, tmp As String
    i = first: j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(arr(j), pivot, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub

Public Function NewGuidString() As String
    ' Versuche echte GUID über Windows-API
    Dim g As GUID
    If CoCreateGuid(g) = 0 Then
        Dim buf As String
        buf = String$(39, vbNullChar) ' 38 Zeichen + Nullterminator
        If StringFromGUID2(g, StrPtr(buf), 39) > 0 Then
            NewGuidString = Mid$(buf, 2, 36) ' ohne { }
            Exit Function
        End If
    End If
    ' Fallback: pseudo-Zufalls-GUID v4
    NewGuidString = PseudoGuidV4()
End Function

Private Function Hex2(ByVal n As Byte) As String
    Hex2 = Right$("0" & Hex$(n), 2)
End Function

' --- Fallback: reine VBA v4-Pseudo-GUID ---
Private Function PseudoGuidV4() As String
    Dim b(0 To 15) As Byte, i As Long
    Randomize Timer
    For i = 0 To 15
        b(i) = Int(Rnd() * 256)
    Next

    ' RFC 4122 v4 Bits setzen
    b(6) = (b(6) And &HF) Or &H40    ' Version 4
    b(8) = (b(8) And &H3F) Or &H80   ' Variant 10xx

    PseudoGuidV4 = _
        Hex2(b(0)) & Hex2(b(1)) & Hex2(b(2)) & Hex2(b(3)) & "-" & _
        Hex2(b(4)) & Hex2(b(5)) & "-" & _
        Hex2(b(6)) & Hex2(b(7)) & "-" & _
        Hex2(b(8)) & Hex2(b(9)) & "-" & _
        Hex2(b(10)) & Hex2(b(11)) & Hex2(b(12)) & Hex2(b(13)) & Hex2(b(14)) & Hex2(b(15))
End Function

Public Function FindLastUsedRow(ws As Worksheet) As Long
    Dim r As Range
    Set r = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If r Is Nothing Then
        FindLastUsedRow = ws.Rows.Count
    Else
        FindLastUsedRow = r.row
    End If
End Function

Public Function FindLastUsedCol(ws As Worksheet, row As Long) As Long
    Dim lastCol As Long
    Dim col As Long
    
    lastCol = ws.Columns.Count   ' z.B. 16384 bei XLSX
    
    ' Von rechts nach links prüfen
    For col = 1 To lastCol
        If Len(Trim$(CStr(ws.Cells(row, col).Value))) = 0 Then
            FindLastUsedCol = col
            Exit Function
        End If
    Next col
    
    FindLastUsedCol = 0 ' Wenn keine einzige Zelle in der Zeile befüllt ist
End Function

' Suche nach einem Tabsheet, das mit WORKSHEET_PREFIX_FOR_ABRECHNUNG beginnt
Public Function GetAbrSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase$(Left$(ws.name, Len(WORKSHEET_PREFIX_FOR_ABRECHNUNG))) = WORKSHEET_PREFIX_FOR_ABRECHNUNG Then
            Set GetAbrSheet = ws
            Exit Function
        End If
    Next
    Set GetAbrSheet = Nothing
End Function

' Spalte ermitteln, in der der übergebene String der Titel (Header) ist.
Public Function FindHeaderCol(ws As Worksheet, hdrRow As Long, headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim cellValue As String
    
    ' Letzte genutzte Spalte im Blatt (inkl. ausgeblendeter Spalten)
    If ws.UsedRange Is Nothing Then
        FindHeaderCol = 0
        Exit Function
    End If
    
    lastCol = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    
    For c = 1 To lastCol
        cellValue = CStr(ws.Cells(hdrRow, c).Value)
        If StrComp(Trim$(cellValue), headerText, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next
    
    FindHeaderCol = 0
End Function

Public Function CollectionToStringArray(col As Collection) As String()
    Dim i As Long, arr() As String
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        arr(i - 1) = CStr(col(i))
    Next
    CollectionToStringArray = arr
End Function

' Suche einen Key (Typ string) in der übergebenen Collectoin und gib True zurück, wenn gefunden.
Public Function InCollection(col As Collection, ByVal key As String) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If StrComp(col(i), key, vbTextCompare) = 0 Then
            InCollection = True
            Exit Function
        End If
    Next
End Function

Public Sub ActivateAnySheet(ByRef ws As Worksheet, sheetName As String, Optional doActivate As Boolean = True)
    Set ws = ThisWorkbook.Sheets(sheetName)
    If doActivate Then ws.Activate
End Sub

Function FileExists(filePath As String) As Boolean
    FileExists = False
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Public Sub FormatHeader(wsDest As Worksheet, headerRange As String)
    ' Formatting des Headers (fett, Linien)
    wsDest.Range(headerRange).Select
    With Selection.Font
        .name = "Arial"
        .FontStyle = "Fett"
        .Size = 10
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .Weight = xlThin
    End With
End Sub

Public Function IsAbsolutePath(ByVal p As String) As Boolean
    p = Trim$(p)

    ' UNC: \\Server\Share
    If Left$(p, 2) = "\\" Then
        IsAbsolutePath = True
        Exit Function
    End If

    ' Drive-Pfad: C:\...
    If Len(p) >= 3 Then
        If Mid$(p, 2, 2) = ":\" Then
            IsAbsolutePath = True
            Exit Function
        End If
    End If

    IsAbsolutePath = False
End Function

Public Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Sub DeleteTabsheet(name As String)
    Application.DisplayAlerts = False ' Verhindert Rückfragen beim Löschen

    On Error Resume Next
    Sheets(name).Delete
    On Error GoTo 0
    
    Application.DisplayAlerts = True
End Sub
