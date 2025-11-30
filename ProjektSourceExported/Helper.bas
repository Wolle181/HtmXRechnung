Attribute VB_Name = "Helper"
Option Explicit

' Kapselt die Logik, ob der Mandant (Name oder Nummer) zum Suchkey passt.
Public Function IsMandantMatch(ByVal sMdNr As String, _
                                ByVal sMd As String, _
                                ByVal key As String, _
                                ByVal byName As Boolean) As Boolean
    If byName Then
        IsMandantMatch = (Len(sMd) >= Len(key) And StrComp(Left$(sMd, Len(key)), key, vbTextCompare) = 0)
    Else
        IsMandantMatch = (StrComp(sMdNr, key, vbTextCompare) = 0)
    End If
End Function

' Blendet die Spalte "Zeilen-ID" in einem Arbeitsblatt aus (falls vorhanden).
Public Sub HideZeilenIdColumn(ByVal ws As Worksheet, ByVal hdrRow As Long)
    Dim col As Long
    col = Utils.FindHeaderCol(ws, hdrRow, HEADER_ZEILEN_ID)
    If col > 0 Then
        ws.Columns(col).EntireColumn.Hidden = True
    End If
End Sub

Public Function GetStundensatz(key As String) As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Settings")

    ' Letzte belegte Zeile in Spalte 6 (MA-Kürzel / F)
    lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).row

    ' Zeile 4 bis lastRow durchlaufen
    For i = 4 To lastRow
        If Trim(CStr(ws.Cells(i, 6).Value)) = Trim$(key) Then
            ' Spalte 7 = Stundensatz
            GetStundensatz = CDbl(ws.Cells(i, 7).Value)
            Exit Function
        End If
    Next i

    ' Wenn nichts gefunden wird
    GetStundensatz = 0
End Function

' Kopiert die Kopfzeile aus einem WORKSHEET_PREFIX_TO_COLLECT-Sheet einmalig ins ABR-Sheet und ergänzt eine Spalte "abzurechnen".
Public Sub CopyAbrHeader(ByVal ws As Worksheet, _
                          ByVal wsDest As Worksheet, _
                          ByVal hdrRow As Long, _
                          ByRef destRow As Long, _
                          ByVal lastCol As Long, _
                          ByRef baseLastCol As Long, _
                          ByRef headerCopied As Boolean)
    baseLastCol = lastCol

    Dim srcHeader As Range, dstHeader As Range
    Set srcHeader = ws.Cells(hdrRow, 1).Resize(1, baseLastCol)
    Set dstHeader = wsDest.Cells(destRow, 1).Resize(1, baseLastCol)

    dstHeader.Value = srcHeader.Value
    wsDest.Rows(destRow).RowHeight = ws.Rows(hdrRow).RowHeight

    ' Zusätzliche Spalten
    wsDest.Cells(destRow, baseLastCol + 1).Value = HEADER_ABZURECHNEN
    wsDest.Cells(destRow, baseLastCol + 2).Value = HEADER_QUELLBLATT
    wsDest.Cells(destRow, baseLastCol + 3).Value = HEADER_STDSATZ

    Utils.FormatHeader wsDest, HEADER_RANGE

    headerCopied = True
    destRow = destRow + 1
End Sub

Public Function CollectUniqueMdNames() As Collection
    On Error GoTo EH

    Dim result As New Collection
    Dim ws As Worksheet
    Dim basePath As String

    ' 1) Lokale MA-Sheets im aktuellen Workbook
    For Each ws In ThisWorkbook.Worksheets
        If Utils.SheetNameIsMA(ws.name) Then
            CollectMdNamesFromMaSheet ws, result
        End If
    Next ws

    ' 2) Alle externen MA_*.xlsx-Dateien unterhalb des Pfads aus Settings!B3
    basePath = Settings.GetMaBasePathFromSettings()
    If Len(basePath) > 0 Then
        AppendMdNamesFromExternalMaFiles basePath, result
    End If

    Set CollectUniqueMdNames = result
    Exit Function

EH:
    Set CollectUniqueMdNames = Nothing
End Function

' Liest die MD-Namen aus einem einzelnen MA-Sheet in die übergebene Collection ein.
Private Sub CollectMdNamesFromMaSheet(ByVal ws As Worksheet, ByRef result As Collection)
    Dim hdrRow As Long
    Dim foundCol As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long
    Dim r As Long
    Dim v As String

    foundCol = 0

    ' Überschrift "MD" in den ersten 5 Zeilen suchen
    For hdrRow = 1 To 5
        lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
        If lastCol = 0 Then GoTo NextHeaderRow

        For c = 1 To lastCol
            If LCase$(Trim$(CStr(ws.Cells(hdrRow, c).Value))) = "md" Then
                foundCol = c
                Exit For
            End If
        Next c

        If foundCol > 0 Then Exit For

NextHeaderRow:
    Next hdrRow

    If foundCol = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, foundCol).End(xlUp).row
    If lastRow <= hdrRow Then Exit Sub

    For r = hdrRow + 1 To lastRow
        v = Trim$(CStr(ws.Cells(r, foundCol).Value))
        If Len(v) > 0 Then
            If Not Utils.InCollection(result, v) Then
                result.Add v
            End If
        End If
    Next r
End Sub

' Liest alle MA_*.xlsx-Dateien unterhalb basePath und sammelt deren MD-Namen.
Private Sub AppendMdNamesFromExternalMaFiles(ByVal basePath As String, _
                                             ByRef result As Collection)
    Dim fso As Object
    Dim rootFolder As Object
    Dim xlApp As Object

    ' Prüfen, ob der Pfad existiert
    If Len(Dir(basePath, vbDirectory)) = 0 Then
        MsgBox "Der in 'Settings'!B3 angegebene Pfad existiert nicht:" & vbCrLf & basePath, vbExclamation
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(basePath)

    ' zweite, unsichtbare Excel-Instanz
    Set xlApp = CreateObject("Excel.Application")
    xlApp.visible = False
    xlApp.DisplayAlerts = False

    On Error GoTo CleanUp

    ' rekursiv alle Unterordner / Dateien abarbeiten
    ProcessMaFilesInFolderForMdNames rootFolder, xlApp, result

CleanUp:
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        xlApp.Quit
    End If
    Set xlApp = Nothing
End Sub

' Rekursive Verarbeitung der Unterordner für MD-Namen.
Private Sub ProcessMaFilesInFolderForMdNames(ByVal folder As Object, _
                                             ByVal xlApp As Object, _
                                             ByRef result As Collection)
    Dim subFolder As Object
    Dim file As Object

    ' Dateien im aktuellen Ordner
    For Each file In folder.Files
        ' nur MA_*.xlsx berücksichtigen (keine .xlsm)
        If LCase$(file.name) Like "ma_*.xlsx" Then
            CollectMdNamesFromExternalWorkbook CStr(file.Path), xlApp, result
        End If
    Next file

    ' Unterordner rekursiv
    For Each subFolder In folder.SubFolders
        ProcessMaFilesInFolderForMdNames subFolder, xlApp, result
    Next subFolder
End Sub

' Öffnet eine einzelne externe MA-Datei und sammelt daraus die MD-Namen.
Private Sub CollectMdNamesFromExternalWorkbook(ByVal filePath As String, _
                                               ByVal xlApp As Object, _
                                               ByRef result As Collection)
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo LocalCleanUp

    Set wb = xlApp.Workbooks.Open(Filename:=filePath, ReadOnly:=True)

    ' In der externen Mappe alle "MA"-Sheets verarbeiten
    For Each ws In wb.Worksheets
        If Utils.SheetNameIsMA(ws.name) Then
            CollectMdNamesFromMaSheet ws, result
        End If
    Next ws

LocalCleanUp:
    On Error Resume Next
    If Not wb Is Nothing Then
        Application.DisplayAlerts = False
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    Set wb = Nothing
End Sub

Public Function IsUniqueMDName(ByVal byName As String) As Boolean
    Dim ws As Worksheet
    Dim hdrRow As Long: hdrRow = 1
    Dim colMd As Long, colMdNr As Long
    Dim lastCol As Long, lastRow As Long
    Dim findLast As Range
    Dim firstMdNr As String
    Dim found As Boolean
    Dim r As Long
    Dim mdName As String, mdNr As String
    
    byName = Trim$(byName)
    If Len(byName) = 0 Then
        IsUniqueMDName = False
        Exit Function
    End If

    For Each ws In ThisWorkbook.Worksheets
        If Utils.SheetNameIsMA(ws.name) Then
            lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
            If lastCol = 0 Then GoTo NextWs

            colMd = Utils.FindHeaderCol(ws, hdrRow, "MD")
            colMdNr = Utils.FindHeaderCol(ws, hdrRow, "MD-Nr")
            If colMd = 0 Or colMdNr = 0 Then GoTo NextWs

            Set findLast = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                         LookAt:=xlPart, SearchOrder:=xlByRows, _
                                         SearchDirection:=xlPrevious, MatchCase:=False)
            If findLast Is Nothing Then GoTo NextWs
            lastRow = findLast.row
            If lastRow <= hdrRow Then GoTo NextWs

            For r = hdrRow + 1 To lastRow
                mdName = Trim$(CStr(ws.Cells(r, colMd).Value))
                If Len(mdName) >= Len(byName) And StrComp(Left$(mdName, Len(byName)), byName, vbTextCompare) = 0 Then
                    mdNr = Trim$(CStr(ws.Cells(r, colMdNr).Value))
                    If Len(mdNr) > 0 Then
                        If Not found Then
                            firstMdNr = mdNr
                            found = True
                        ElseIf StrComp(mdNr, firstMdNr, vbTextCompare) <> 0 Then
                            IsUniqueMDName = False
                            Exit Function
                        End If
                    End If
                End If
            Next r
        End If
NextWs:
    Next ws

    IsUniqueMDName = found   ' True, wenn mind. eine MD-Nr gefunden und alle gleich; sonst False
End Function

Public Function SheetHasData(ws As Worksheet, hdrRow As Long) As Boolean
    SheetHasData = (Utils.FindLastUsedRow(ws) > hdrRow)
End Function

