Attribute VB_Name = "Abgleich"
Option Explicit
    
Private Const ABGLEICH_SHEET_NAME As String = "Abgleich"

Public Sub ErzeugeAbgleichSheet(control As IRibbonControl)
    On Error GoTo EH
    Application.ScreenUpdating = False

    Wartebox.ShowToast "Abgleich wird erstellt"

    Dim wsAbgleich As Worksheet
    Dim destRow As Long
    Dim headerWritten As Boolean
    Dim dictUnique As Object          ' Key: "MDNr||MD", Value: Zeilennummer im Abgleich-Sheet
    Dim dictSheetCols As Object       ' Key: "Workbook!Sheet", Value: Spaltennummer im Abgleich-Sheet
    Dim basePath As String

    ' Abgleich-Sheet holen oder anlegen
    Utils.DeleteTabsheet ABGLEICH_SHEET_NAME
    Set wsAbgleich = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsAbgleich.name = ABGLEICH_SHEET_NAME

    ' Dictionary für einzigartige Kombinationen MD-Nr + MD
    Set dictUnique = CreateObject("Scripting.Dictionary")
    dictUnique.CompareMode = vbTextCompare

    ' Dictionary für Spalten je Quelle (Workbook/Sheet)
    Set dictSheetCols = CreateObject("Scripting.Dictionary")
    dictSheetCols.CompareMode = vbTextCompare

    destRow = 1
    headerWritten = False

    ' 1) Haupt-Sheet "MA HA" im Makro-Workbook
    If Utils.SheetExists(WORKSHEET_HAMAIN) Then
        AppendMdRowsFromMaSheet ThisWorkbook.Worksheets(WORKSHEET_HAMAIN), _
                                wsAbgleich, destRow, headerWritten, dictUnique, dictSheetCols
    Else
        MsgBox "Hauptsheet '" & WORKSHEET_HAMAIN & "' wurde nicht gefunden.", vbExclamation
    End If

    ' 2) Alle externen TIMESHEETNAMETEMPLATE-Dateien unterhalb des Settings-Pfads
    basePath = Settings.GetTimesheetBasePath()
    If Len(basePath) > 0 Then
        AppendMdRowsFromExternalMaFiles basePath, wsAbgleich, destRow, headerWritten, dictUnique, dictSheetCols
    End If

    ' Leere Zeilen (ohne MD-Nr und MD) entfernen – zur Sicherheit noch einmal
    RemoveEmptyRowsInAbgleich wsAbgleich

    ' Header formatieren, wenn überhaupt Daten da sind
    Dim lastHeaderCol As Long: lastHeaderCol = Utils.FindLastUsedCol(wsAbgleich, HEADER_ROW)
    If headerWritten Then
        If lastHeaderCol < 2 Then lastHeaderCol = 2

        Utils.FormatHeader wsAbgleich, wsAbgleich.Range(wsAbgleich.Cells(HEADER_ROW, 1), wsAbgleich.Cells(HEADER_ROW, lastHeaderCol)).Address(False, False)
    End If

    ' Spalten/Zeilen anpassen, Hintergrundfarben zurücksetzen
    wsAbgleich.Columns.AutoFit
    wsAbgleich.Columns("A").ColumnWidth = 15
    wsAbgleich.Rows.AutoFit
    wsAbgleich.Cells.Interior.ColorIndex = xlColorIndexNone
    
    ' verbeitern für SORT-Buttons:
    wsAbgleich.Columns(lastHeaderCol).ColumnWidth = 20
    wsAbgleich.Columns(lastHeaderCol + 1).ColumnWidth = 20
    ' Sortierbuttons
    CreateSortButtons wsAbgleich, lastHeaderCol, HEADER_ROW

    wsAbgleich.Rows(HEADER_ROW).RowHeight = 25
    
    ' Oberste Zeile freezen:
    wsAbgleich.Range("A2").Select
    ActiveWindow.FreezePanes = True

CleanExit:
    Wartebox.CloseToast
    Application.ScreenUpdating = True
    Exit Sub

EH:
    MsgBox "Fehler beim Erzeugen des Abgleich-Sheets: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

' Liefert die Spalte im Abgleich-Sheet für ein bestimmtes Quell-Sheet.
' Legt die Spalte bei Bedarf neu an und schreibt den Spaltenkopf.
Private Function GetOrCreateSheetColumn( _
    ByVal wsAbgleich As Worksheet, _
    ByVal dictSheetCols As Object, _
    ByVal sourceWorkbookName As String, _
    ByVal sourceSheetName As String) As Long

    Dim key As String
    key = sourceWorkbookName & "!" & sourceSheetName

    If Not dictSheetCols.Exists(key) Then
        Dim nextCol As Long
        Dim hdrRow As Long
        hdrRow = HEADER_ROW

        nextCol = Utils.FindLastUsedCol(wsAbgleich, hdrRow)
        If nextCol < 2 Then nextCol = 2

        dictSheetCols.Add key, nextCol

        Dim caption As String: caption = sourceSheetName
        wsAbgleich.Cells(hdrRow, nextCol).Value = caption
    End If

    GetOrCreateSheetColumn = CLng(dictSheetCols(key))
End Function

Private Sub AppendMdRowsFromMaSheet(ByVal wsSource As Worksheet, _
                                    ByVal wsAbgleich As Worksheet, _
                                    ByRef destRow As Long, _
                                    ByRef headerWritten As Boolean, _
                                    ByVal dictUnique As Object, _
                                    ByVal dictSheetCols As Object)

    Dim hdrRow As Long: hdrRow = HEADER_ROW
    Dim lastCol As Long
    Dim lastRow As Long
    Dim colMdNr As Long
    Dim colMd As Long
    Dim r As Long
    Dim sMdNr As String
    Dim sMd As String
    Dim key As String
    Dim abgleichRow As Long
    Dim colSource As Long

    ' Letzte verwendete Spalte in der Header-Zeile
    lastCol = Utils.FindLastUsedCol(wsSource, hdrRow)
    If lastCol = 0 Then Exit Sub

    ' Spalten für MD-Nr und MD suchen
    colMdNr = Utils.FindHeaderCol(wsSource, hdrRow, HEADER_MDNR)
    colMd = Utils.FindHeaderCol(wsSource, hdrRow, HEADER_MD)

    If colMdNr = 0 Or colMd = 0 Then Exit Sub

    lastRow = Utils.FindLastUsedRow(wsSource)
    If lastRow <= hdrRow Then Exit Sub

    ' Header im Abgleich-Sheet einmalig schreiben
    If Not headerWritten Then
        wsAbgleich.Cells(HEADER_ROW, 1).Value = HEADER_MDNR
        wsAbgleich.Cells(HEADER_ROW, 2).Value = HEADER_MD
        headerWritten = True
        If destRow < HEADER_ROW + 1 Then destRow = HEADER_ROW + 1
    End If

    ' Spalte für dieses Quell-Sheet im Abgleich-Sheet
    colSource = GetOrCreateSheetColumn(wsAbgleich, dictSheetCols, wsSource.Parent.name, wsSource.name)

    ' Datenzeilen einsammeln
    For r = hdrRow + 1 To lastRow
        sMdNr = Trim$(CStr(wsSource.Cells(r, colMdNr).Value))
        sMd = Trim$(CStr(wsSource.Cells(r, colMd).Value))

        ' Nur Zeilen mit Inhalt in mindestens einem der beiden Felder
        If (Len(sMdNr) > 0) Or (Len(sMd) > 0) Then
            key = sMdNr & "||" & sMd

            If dictUnique.Exists(key) Then
                abgleichRow = CLng(dictUnique(key))
            Else
                abgleichRow = destRow
                dictUnique.Add key, abgleichRow
                wsAbgleich.Cells(abgleichRow, 1).Value = sMdNr
                wsAbgleich.Cells(abgleichRow, 2).Value = sMd
                destRow = destRow + 1
            End If

            ' Kennzeichnung, dass diese Kombination in diesem Sheet vorkommt
            wsAbgleich.Cells(abgleichRow, colSource).Value = "X"
        End If
    Next r
End Sub

' Erzeugen einer xeparatne XL-Instanz und Initialisieren der Dictionaries
Private Sub AppendMdRowsFromExternalMaFiles(ByVal basePath As String, _
                                            ByVal wsAbgleich As Worksheet, _
                                            ByRef destRow As Long, _
                                            ByRef headerWritten As Boolean, _
                                            ByVal dictUnique As Object, _
                                            ByVal dictSheetCols As Object)
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

    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False

    On Error GoTo CleanUp

    ' rekursiv alle Unterordner / Dateien abarbeiten
    ProcessMaFilesInFolderForAbgleich rootFolder, xlApp, wsAbgleich, destRow, headerWritten, dictUnique, dictSheetCols

CleanUp:
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        xlApp.Quit
    End If
    Set xlApp = Nothing
End Sub

Private Sub ProcessMaFilesInFolderForAbgleich(ByVal folder As Object, _
                                              ByVal xlApp As Object, _
                                              ByVal wsAbgleich As Worksheet, _
                                              ByRef destRow As Long, _
                                              ByRef headerWritten As Boolean, _
                                              ByVal dictUnique As Object, _
                                              ByVal dictSheetCols As Object)
    Dim subFolder As Object
    Dim file As Object

    Dim TimesheetNameTemplate As String: TimesheetNameTemplate = LCase$(Settings.GetTimesheetNameTemplate())

    ' Dateien im aktuellen Ordner
    For Each file In folder.Files
        ' nur TIMESHEETNAMETEMPLATE-Dateien berücksichtigen (keine .xlsm)
        If LCase$(file.name) Like TimesheetNameTemplate Then
            AppendMdRowsFromExternalWorkbookForAbgleich CStr(file.Path), xlApp, wsAbgleich, _
                                                        destRow, headerWritten, dictUnique, dictSheetCols
        End If
    Next file

    ' Unterordner rekursiv
    For Each subFolder In folder.SubFolders
        ProcessMaFilesInFolderForAbgleich subFolder, xlApp, wsAbgleich, destRow, headerWritten, dictUnique, dictSheetCols
    Next subFolder
End Sub

Private Sub AppendMdRowsFromExternalWorkbookForAbgleich(ByVal filePath As String, _
                                                        ByVal xlApp As Object, _
                                                        ByVal wsAbgleich As Worksheet, _
                                                        ByRef destRow As Long, _
                                                        ByRef headerWritten As Boolean, _
                                                        ByVal dictUnique As Object, _
                                                        ByVal dictSheetCols As Object)
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo LocalCleanUp

    Set wb = xlApp.Workbooks.Open(Filename:=filePath, ReadOnly:=True)

    ' In der externen Mappe alle MA-Sheets verarbeiten
    For Each ws In wb.Worksheets
        If Utils.SheetNameIsMA(ws.name) Then
            AppendMdRowsFromMaSheet ws, wsAbgleich, destRow, headerWritten, dictUnique, dictSheetCols
        End If
    Next ws

LocalCleanUp:
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    Set wb = Nothing
End Sub

Private Sub RemoveEmptyRowsInAbgleich(ByVal wsAbgleich As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    lastRow = Utils.FindLastUsedRow(wsAbgleich)
    If lastRow < 2 Then Exit Sub

    For r = lastRow To 2 Step -1
        If Trim$(CStr(wsAbgleich.Cells(r, 1).Value)) = "" _
           And Trim$(CStr(wsAbgleich.Cells(r, 2).Value)) = "" Then
            wsAbgleich.Rows(r).Delete
        End If
    Next r
End Sub

Private Sub CreateSortButtons(ws As Worksheet, ByVal spalte As Long, ByVal zeile As Long)
    Dim cell As Range: Set cell = ws.Cells(zeile, spalte)
    Dim topPos As Double, leftPos As Double, width As Double, height As Double
    topPos = cell.Top
    leftPos = cell.Left
    width = cell.width
    height = cell.height

    ' Buttons sind in der gefreezten Header-Zeile, damit sie nicht mitscrollen
    Dim btnSortByMdNr As Button: Set btnSortByMdNr = ws.Buttons.Add(leftPos, topPos, width, height)
    With btnSortByMdNr
        .caption = "Sortiere nach MD-Nr"
        .OnAction = "Abgleich.SortByMdNr"
    End With
    
    Dim btnSortByMd As Button: Set btnSortByMd = ws.Buttons.Add(leftPos + width + 3, topPos, width, height)
    With btnSortByMd
        .caption = "Sortiere nach MD"
        .OnAction = "Abgleich.SortByMd"
    End With
End Sub

Public Sub SortByMdNr()
    On Error GoTo EH

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(ABGLEICH_SHEET_NAME)
    Dim lastRow As Long: lastRow = Utils.FindLastUsedRow(ws)
    
    If lastRow <= 1 Then Exit Sub ' nichts zu sortieren

    Dim lastCol As Long: lastCol = Utils.FindLastUsedCol(ws, HEADER_ROW)

    ' Sortiere nach Spalte 1 (MD-Nr)
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With ws.Sort
'        .SetRange ws.Range("A1:B" & lastRow)
        .SetRange ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, lastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    HighlightDuplicateValuesInAbgleich ws, 1
    Exit Sub

EH:
    MsgBox "Fehler beim Sortieren nach MDNr: " & Err.Description, vbExclamation
End Sub

Public Sub SortByMd()
    On Error GoTo EH

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(ABGLEICH_SHEET_NAME)
    Dim lastRow As Long: lastRow = Utils.FindLastUsedRow(ws)
    
    If lastRow <= 1 Then Exit Sub ' nichts zu sortieren

    Dim lastCol As Long: lastCol = Utils.FindLastUsedCol(ws, HEADER_ROW)

    ' Sortiere nach Spalte 2 (MD)
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add key:=ws.Range("B2:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, lastCol))
        '.SetRange ws.Range("A1:B" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    
    HighlightDuplicateValuesInAbgleich ws, 2
    Exit Sub

EH:
    MsgBox "Fehler beim Sortieren nach MD: " & Err.Description, vbExclamation
End Sub

Private Sub HighlightDuplicateValuesInAbgleich(ws As Worksheet, ByVal colSpec As Variant)
    On Error GoTo EH

    Dim colIndex As Long
    Dim lastRow As Long
    Dim r As Long
    Dim prevValue As String
    Dim curValue As String
    Dim inRun As Boolean
    Dim prevRow As Long
    
    Dim color1 As Long
    Dim color2 As Long
    Dim runColor As Long
    Dim useFirstColor As Boolean

    ' Zwei verschiedene Farben definieren
    color1 = RGB(253, 233, 217) ' helles pastell-orange
    color2 = RGB(255, 255, 153) ' helles Gelb/Creme als zweite Farbe
    useFirstColor = True        ' erste Duplikatgruppe startet mit color1

    ' Alle Hintergrundfarben im gesamten Sheet zurücksetzen
    ws.Cells.Interior.ColorIndex = xlColorIndexNone

    ' Spaltenangabe auswerten: 1/2 oder "A"/"B"
    If IsNumeric(colSpec) Then
        colIndex = CLng(colSpec)
    Else
        colIndex = ws.Columns(CStr(colSpec)).Column
    End If

    If colIndex < 1 Then Exit Sub

    lastRow = Utils.FindLastUsedRow(ws)
    If lastRow <= 2 Then Exit Sub ' nur Header oder zu wenig Daten

    prevValue = ""
    inRun = False
    prevRow = 0

    ' Sicherheit: Farben in der relevanten Spalte zurücksetzen
    ws.Range(ws.Cells(2, colIndex), ws.Cells(lastRow, colIndex)).Interior.ColorIndex = xlColorIndexNone

    For r = 2 To lastRow
        curValue = Trim$(CStr(ws.Cells(r, colIndex).Value))

        If curValue <> "" Then
            If curValue = prevValue Then
                ' Beginn oder Fortsetzung einer Duplikatserie
                If Not inRun Then
                    ' Neue Duplikat-Gruppe startet -> Farbe für diese Gruppe wählen
                    If useFirstColor Then
                        runColor = color1
                    Else
                        runColor = color2
                    End If
                    useFirstColor = Not useFirstColor

                    ' Erstes Element der Gruppe ist die vorherige Zeile
                    If prevRow >= 2 Then
                        ws.Cells(prevRow, colIndex).Interior.Color = runColor
                    End If

                    inRun = True
                End If

                ' Aktuelle Zeile gehört zur laufenden Duplikat-Gruppe
                ws.Cells(r, colIndex).Interior.Color = runColor
            Else
                ' Neuer Wert -> aktuelle Duplikatgruppe endet
                inRun = False
            End If
        Else
            ' Leere Zellen beenden Duplikatgruppen
            inRun = False
        End If

        prevValue = curValue
        prevRow = r
    Next r

    Exit Sub

EH:
    MsgBox "Fehler in HighlightDuplicateValuesInAbgleich: " & Err.Description, vbExclamation
End Sub

