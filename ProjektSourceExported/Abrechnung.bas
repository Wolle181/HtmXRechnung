Attribute VB_Name = "Abrechnung"
Option Explicit

Public Sub SucheNichtAbgerechnetePositionen(ByVal key As String, ByVal byName As Boolean)
    On Error GoTo EH

    Dim wsDest As Worksheet
    Dim abrName As String
    Dim destRow As Long
    Dim baseLastCol As Long
    Dim headerCopied As Boolean

    ' 1) Existenz eines ABR-Tempblatts prüfen (mit WORKSHEET_PREFIX_FOR_ABRECHNUNG beginnend)
    If Utils.DoesAbrSheetExist() Then
        MsgBox "Vorherige Abrechnung beenden vor Erstellung einer neuen.", vbExclamation
        Application.Cursor = xlDefault
        Exit Sub
    End If

    Wartebox.ShowToast "Suche gewünschte Positionen in allen Zeiterfassungsdateien"

    ' 2) Neues ABR-Sheet anlegen
    abrName = WORKSHEET_PREFIX_FOR_ABRECHNUNG & Format(Now, "yyyymmdd_HhNnSs")
    Set wsDest = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsDest.name = abrName

    destRow = 1           ' nächste freie Zeile im ABR-Sheet
    baseLastCol = 0       ' letzte "Daten"-Spalte (ohne "abzurechnen")
    headerCopied = False

    ' 3a) Lokales MA_HA-Sheet aus der aktuellen Arbeitsmappe auswerten (falls vorhanden)
    Dim wsHa As Worksheet
    On Error Resume Next
    Set wsHa = ThisWorkbook.Worksheets(WORKSHEET_HAMAIN)
    On Error GoTo EH

    If Not wsHa Is Nothing Then
        AppendUnbilledPositionsFromMaSheet wsHa, wsDest, key, byName, destRow, baseLastCol, headerCopied
    End If

    ' 3b) Alle externen MA_*.xlsx-Dateien unterhalb des Pfads aus Settings!B3 einlesen
    Dim basePath As String
    basePath = Settings.GetMaBasePathFromSettings()
    If Len(basePath) > 0 Then
        AppendUnbilledPositionsFromExternalMaFiles basePath, wsDest, key, byName, destRow, baseLastCol, headerCopied
    End If

    Wartebox.CloseToast

    ' 4) Feedback / ggf. leeres ABR-Sheet wieder löschen
    If destRow <= 1 Then
        Application.DisplayAlerts = False
        wsDest.Delete
        Application.DisplayAlerts = True
        MsgBox "Keine passenden, nicht abgerechneten Positionen gefunden.", vbInformation
    End If
    
    Exit Sub

EH:
    Wartebox.CloseToast
    MsgBox "Fehler: " & Err.Description, vbExclamation
End Sub

' *** Hilfsfunktionen für 'SucheNichtAbgerechnetePositionen' ***

' Lädt aus einem einzelnen WORKSHEET_PREFIX_TO_COLLECT-Sheet alle nicht abgerechneten Zeilen
' zum gewünschten Mandanten ins ABR-Sheet.
Private Sub AppendUnbilledPositionsFromMaSheet(ByVal ws As Worksheet, _
                                               ByVal wsDest As Worksheet, _
                                               ByVal key As String, _
                                               ByVal byName As Boolean, _
                                               ByRef destRow As Long, _
                                               ByRef baseLastCol As Long, _
                                               ByRef headerCopied As Boolean)
    Dim hdrRow As Long: hdrRow = HEADER_ROW
    If Not Helper.SheetHasData(ws, hdrRow) Then Exit Sub

    ' relevante Spalten suchen
    Dim colZeilenId As Long
    Dim colAbgerechnet As Long
    Dim colMdNr As Long
    Dim colMd As Long

    colZeilenId = Utils.FindHeaderCol(ws, hdrRow, HEADER_ZEILEN_ID)
    colAbgerechnet = Utils.FindHeaderCol(ws, hdrRow, HEADER_ABGERECHNET)
    colMdNr = Utils.FindHeaderCol(ws, hdrRow, HEADER_MDNR)
    colMd = Utils.FindHeaderCol(ws, hdrRow, HEADER_MD)

    If colZeilenId = 0 Or colAbgerechnet = 0 Or colMdNr = 0 Or colMd = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = Utils.FindLastUsedRow(ws)
    If lastRow <= hdrRow Then Exit Sub

    ' 1) leere Zeilen-IDs auffüllen
    Dim lastCol As Long: lastCol = FindLastUsedCol(ws, 1)
    FillMissingZeilenIds ws, hdrRow, lastRow, lastCol, colZeilenId

    ' Zeilen-ID-Spalte im Quellblatt ausblenden
    Helper.HideZeilenIdColumn ws, hdrRow

    ' 2) Header einmalig ins ABR-Sheet kopieren
    If Not headerCopied Then
        Helper.CopyAbrHeader ws, wsDest, hdrRow, destRow, lastCol, baseLastCol, headerCopied
    End If

    ' 3) nicht abgerechnete Zeilen zum Mandanten kopieren
    CopyUnbilledMandantRowsToAbr ws, wsDest, key, byName, hdrRow, lastRow, colMdNr, colMd, colAbgerechnet, baseLastCol, destRow
    
    ' Spaltenbreite anpassen
    wsDest.Range(HEADER_RANGE).EntireColumn.AutoFit
   
    ' Zeilen-ID-Spalte im ABR-Sheet (wsDest) ausblenden
    Helper.HideZeilenIdColumn wsDest, hdrRow
End Sub

Private Sub AppendUnbilledPositionsFromExternalMaFiles(ByVal basePath As String, _
                                                       ByVal wsDest As Worksheet, _
                                                       ByVal key As String, _
                                                       ByVal byName As Boolean, _
                                                       ByRef destRow As Long, _
                                                       ByRef baseLastCol As Long, _
                                                       ByRef headerCopied As Boolean)
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
    ProcessMaFilesInFolder rootFolder, xlApp, wsDest, key, byName, destRow, baseLastCol, headerCopied

CleanUp:
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        xlApp.Quit
    End If
    Set xlApp = Nothing
End Sub

' Rekursive Verarbeitung der Unterordner:
Private Sub ProcessMaFilesInFolder(ByVal folder As Object, _
                                   ByVal xlApp As Object, _
                                   ByVal wsDest As Worksheet, _
                                   ByVal key As String, _
                                   ByVal byName As Boolean, _
                                   ByRef destRow As Long, _
                                   ByRef baseLastCol As Long, _
                                   ByRef headerCopied As Boolean)
    Dim subFolder As Object
    Dim file As Object

    ' Dateien im aktuellen Ordner
    For Each file In folder.Files
        ' nur MA_*.xlsx berücksichtigen (keine .xlsm)
        If LCase$(file.name) Like "ma_*.xlsx" Then
            ProcessSingleMaWorkbook CStr(file.Path), xlApp, wsDest, key, byName, destRow, baseLastCol, headerCopied
        End If
    Next file

    ' Unterordner rekursiv
    For Each subFolder In folder.SubFolders
        ProcessMaFilesInFolder subFolder, xlApp, wsDest, key, byName, destRow, baseLastCol, headerCopied
    Next subFolder
End Sub

' Öffnen einer einzelnen externen Datei:
Private Sub ProcessSingleMaWorkbook(ByVal filePath As String, _
                                    ByVal xlApp As Object, _
                                    ByVal wsDest As Worksheet, _
                                    ByVal key As String, _
                                    ByVal byName As Boolean, _
                                    ByRef destRow As Long, _
                                    ByRef baseLastCol As Long, _
                                    ByRef headerCopied As Boolean)
    Dim wb As Workbook
    Dim ws As Worksheet

    On Error GoTo LocalCleanUp

    Set wb = xlApp.Workbooks.Open(Filename:=filePath, ReadOnly:=False, Notify:=False)
    If wb.ReadOnly Then
        MsgBox "Die Datei" & vbCrLf & filePath & vbCrLf & "ist bereits schreibend geöffnet. Sie kann nicht gelesen werden, weil dann ggf. zurückgeschrieben werden muss.", vbExclamation
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' In der externen Mappe alle WORKSHEET_PREFIX_TO_COLLECT-Sheets verarbeiten
    For Each ws In wb.Worksheets
        If UCase$(Left$(ws.name, Len(WORKSHEET_PREFIX_TO_COLLECT))) = WORKSHEET_PREFIX_TO_COLLECT Then
            AppendUnbilledPositionsFromMaSheet ws, wsDest, key, byName, destRow, baseLastCol, headerCopied
        End If
    Next ws

LocalCleanUp:
    On Error Resume Next
    If Not wb Is Nothing Then
        Application.DisplayAlerts = False
        wb.Close SaveChanges:=True
        Application.DisplayAlerts = True
    End If
    
    Set wb = Nothing
End Sub

' Füllt fehlende Zeilen-IDs mit GUIDs, falls in der Zeile überhaupt Inhalte vorhanden sind.
Private Sub FillMissingZeilenIds(ByVal ws As Worksheet, _
                                 ByVal hdrRow As Long, _
                                 ByVal lastRow As Long, _
                                 ByVal lastCol As Long, _
                                 ByVal colZeilenId As Long)
    Dim r As Long, c As Long
    Dim hasContent As Boolean

    For r = hdrRow + 1 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, colZeilenId).Value))) = 0 Then
            hasContent = False
            For c = 1 To lastCol
                If Len(Trim$(ws.Cells(r, c).text)) > 0 Then
                    hasContent = True
                    Exit For
                End If
            Next c

            If hasContent Then
                ws.Cells(r, colZeilenId).Value = NewGuidString()
            End If
        End If
    Next r
End Sub

' Kopiert alle nicht abgerechneten Zeilen des passenden Mandanten aus einem WORKSHEET_PREFIX_TO_COLLECT-Sheet in das ABR-Sheet.
Private Sub CopyUnbilledMandantRowsToAbr(ByVal ws As Worksheet, _
                                         ByVal wsDest As Worksheet, _
                                         ByVal key As String, _
                                         ByVal byName As Boolean, _
                                         ByVal hdrRow As Long, _
                                         ByVal lastRow As Long, _
                                         ByVal colMdNr As Long, _
                                         ByVal colMd As Long, _
                                         ByVal colAbgerechnet As Long, _
                                         ByVal baseLastCol As Long, _
                                         ByRef destRow As Long)
    Dim r As Long
    For r = hdrRow + 1 To lastRow
        ' MD-Nr oder MD leer? -> Zeile ignorieren
        Dim sMdNr As String: sMdNr = Trim$(CStr(ws.Cells(r, colMdNr).Value))
        Dim sMd As String:   sMd = Trim$(CStr(ws.Cells(r, colMd).Value))
        If sMdNr = "" Or sMd = "" Then GoTo NextRow

        If Helper.IsMandantMatch(sMdNr, sMd, key, byName) Then
            Dim abg As String: abg = Trim$(CStr(ws.Cells(r, colAbgerechnet).Value))
            If abg = "" Then
                ' komplette Zeile kopieren
                wsDest.Cells(destRow, 1).Resize(1, baseLastCol).Value = ws.Cells(r, 1).Resize(1, baseLastCol).Value

                ' Zusatzspalte "abzurechnen" bleibt leer (User entscheidet)
                wsDest.Cells(destRow, baseLastCol + 1).Value = ""

                ' Name des Quellblatts (Mitarbeiter)
                wsDest.Cells(destRow, baseLastCol + 2).Value = ws.name

                ' Std-Satz für Mitarbeiter ausgeben
                wsDest.Cells(destRow, baseLastCol + 3).Value = Format(Helper.GetStundensatz(ws.name), "0.00")

                destRow = destRow + 1
            End If
        End If
NextRow:
    Next r
End Sub

' *** PUBLIC-FUNKTION ***

Public Sub SchreibeAbgerechnetZurueck()
    On Error GoTo EH

    Dim wsABR As Worksheet
    Set wsABR = Utils.GetAbrSheet()
    If wsABR Is Nothing Then
        MsgBox "Kein ABR-Temp-Sheet gefunden.", vbExclamation
        Exit Sub
    End If

    Dim hdrRow As Long: hdrRow = HEADER_ROW

    Dim colZeilenId As Long
    colZeilenId = Utils.FindHeaderCol(wsABR, hdrRow, HEADER_ZEILEN_ID)
    If colZeilenId = 0 Then
        MsgBox "Im ABR-Sheet fehlt die Spalte '" & HEADER_ZEILEN_ID & "'.", vbExclamation
        Exit Sub
    End If

    Dim colSelect As Long
    colSelect = Utils.FindHeaderCol(wsABR, hdrRow, HEADER_ABZURECHNEN)
    If colSelect = 0 Then
        MsgBox "Im ABR-Sheet fehlt die Spalte '" & HEADER_ABZURECHNEN & "'.", vbExclamation
        Exit Sub
    End If

    Wartebox.ShowToast "Schreibe markierte Positionen zurück in Zeiterfassungsdateien"

    ' 1) Markierte Zeilen-IDs im ABR-Sheet einsammeln
    Dim dictSelected As Object
    Set dictSelected = CollectMarkedAbrRowIds(wsABR, hdrRow, colZeilenId, colSelect)
    If dictSelected Is Nothing Then Exit Sub        ' Fehler bereits gemeldet
    If dictSelected.Count = 0 Then
        Wartebox.CloseToast
        MsgBox "Keine Zeilen zum Zurückschreiben markiert.", vbInformation
        Exit Sub
    End If

    ' 2) Zuordnung bauen: für jede ID genau eine Zielzelle finden
    Dim toUpdate As Object: Set toUpdate = CreateObject("Scripting.Dictionary")
    Dim ambiguous As Collection: Set ambiguous = New Collection
    Dim missing As Collection: Set missing = New Collection

    BuildUpdateMapForSelectedIds dictSelected, hdrRow, toUpdate, ambiguous, missing

    ' 3) Falls IDs fehlen oder mehrfach vorkommen, Dialog zeigen und ggf. abbrechen
    If Not ShowAmbiguousOrMissingIdsAndAbort(dictSelected, ambiguous, missing) Then
        Wartebox.CloseToast
        Exit Sub
    End If

    ' 4) Schreiben der ABGERECHNET_MARKER in die Spalte "abgerechnet"
    Dim countSet As Long
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ApplyAbgerechnetUpdates toUpdate, countSet

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Wartebox.CloseToast
    MsgBox "Zurückgeschrieben nach 'MA '-Sheets: " & countSet & " Zeile(n).", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Wartebox.CloseToast
    MsgBox "Fehler: " & Err.Description, vbExclamation
End Sub

' *** Hilfsfunktionen für 'SchreibeAbgerechnetZurueck' ***

' Liest alle im ABR-Sheet markierten Zeilen (Spalte "abzurechnen") und
' gibt ein Dictionary von Zeilen-ID -> True zurück.
Private Function CollectMarkedAbrRowIds(ByVal wsABR As Worksheet, _
                                        ByVal hdrRow As Long, _
                                        ByVal colZeilenId As Long, _
                                        ByVal colSelect As Long) As Object

    If Not Helper.SheetHasData(wsABR, hdrRow) Then
        MsgBox "ABR-Sheet enthält keine Datenzeilen.", vbInformation
        Set CollectMarkedAbrRowIds = Nothing
        Exit Function
    End If

    Dim dictSelected As Object
    Set dictSelected = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim id As String
    Dim mark As String

    Dim lastrowABR As Long: lastrowABR = Utils.FindLastUsedRow(wsABR)
    For r = hdrRow + 1 To lastrowABR
        mark = Trim$(CStr(wsABR.Cells(r, colSelect).Value))
        If Len(mark) > 0 Then
            id = Trim$(CStr(wsABR.Cells(r, colZeilenId).Value))

            If Len(id) = 0 Then
                MsgBox "Im ABR-Sheet ist eine markierte Zeile ohne '" & HEADER_ZEILEN_ID & "'.", vbExclamation
                Set dictSelected = Nothing
                Set CollectMarkedAbrRowIds = Nothing
                Exit Function
            End If

            If dictSelected.Exists(id) Then
                MsgBox "Zeilen-ID im ABR-Sheet mehrfach markiert: " & id, vbExclamation
                Set dictSelected = Nothing
                Set CollectMarkedAbrRowIds = Nothing
                Exit Function
            End If

            dictSelected.Add id, True
        End If
    Next

    Set CollectMarkedAbrRowIds = dictSelected
End Function

' Baut auf Basis der markierten IDs eine Map, welche MA-Zeile
' (Sheet, Zeile, Spalte "abgerechnet") später gesetzt werden soll.
Private Sub BuildUpdateMapForSelectedIds(ByVal dictSelected As Object, _
                                         ByVal hdrRow As Long, _
                                         ByRef toUpdate As Object, _
                                         ByRef ambiguous As Collection, _
                                         ByRef missing As Collection)
    Dim foundCount As Object
    Set foundCount = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    For Each k In dictSelected.Keys
        foundCount.Add CStr(k), 0
    Next k

    ' 1) In lokalen MA-Sheets (inkl. MA_HA) suchen
    BuildUpdateMapFromLocalMaSheets dictSelected, hdrRow, toUpdate, ambiguous, foundCount

    ' 2) In externen MA_*.xlsx-Dateien (Pfad aus Settings!B3) suchen
    Dim basePath As String
    basePath = Settings.GetMaBasePathFromSettings()   ' muss vorhanden sein (aus deiner vorherigen Anpassung)

    If Len(basePath) > 0 Then
        BuildUpdateMapFromExternalMaFiles dictSelected, hdrRow, toUpdate, ambiguous, foundCount, basePath
    End If

    ' 3) IDs, die gar nicht gefunden wurden, als "missing" markieren
    For Each k In dictSelected.Keys
        If CLng(foundCount(CStr(k))) = 0 Then
            missing.Add CStr(k)
        End If
    Next k
End Sub

Private Sub BuildUpdateMapFromExternalMaFiles(ByVal dictSelected As Object, _
                                              ByVal hdrRow As Long, _
                                              ByRef toUpdate As Object, _
                                              ByRef ambiguous As Collection, _
                                              ByRef foundCount As Object, _
                                              ByVal basePath As String)
    Dim fso As Object
    Dim rootFolder As Object
    Dim xlApp As Object

    If Len(Dir(basePath, vbDirectory)) = 0 Then
        MsgBox "Der in 'Settings'!B3 angegebene Pfad existiert nicht:" & vbCrLf & basePath, vbExclamation
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(basePath)

    Set xlApp = CreateObject("Excel.Application")
    xlApp.visible = False
    xlApp.DisplayAlerts = False

    On Error GoTo CleanUp

    ProcessExternalMaFilesForUpdate rootFolder, xlApp, dictSelected, hdrRow, toUpdate, ambiguous, foundCount

CleanUp:
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = True
        xlApp.Quit
    End If
    Set xlApp = Nothing
End Sub

Private Sub BuildUpdateMapFromLocalMaSheets(ByVal dictSelected As Object, _
                                            ByVal hdrRow As Long, _
                                            ByRef toUpdate As Object, _
                                            ByRef ambiguous As Collection, _
                                            ByRef foundCount As Object)
    Dim ws As Worksheet
    Dim colId As Long
    Dim colAbg As Long
    Dim lastRow As Long
    Dim r As Long
    Dim id As String
    Dim cnt As Long

    For Each ws In ThisWorkbook.Worksheets
        ' Explizit MA_HA und generisch alle "MA "-Sheets
        If ws.name = WORKSHEET_HAMAIN Or UCase$(Left$(ws.name, Len(WORKSHEET_PREFIX_TO_COLLECT))) = WORKSHEET_PREFIX_TO_COLLECT Then

            colId = Utils.FindHeaderCol(ws, hdrRow, HEADER_ZEILEN_ID)
            colAbg = Utils.FindHeaderCol(ws, hdrRow, HEADER_ABGERECHNET)

            If colId <> 0 And colAbg <> 0 Then
                lastRow = Utils.FindLastUsedRow(ws)

                For r = hdrRow + 1 To lastRow
                    id = Trim$(CStr(ws.Cells(r, colId).Value))
                    If Len(id) > 0 Then
                        If dictSelected.Exists(id) Then
                            If Not foundCount.Exists(id) Then foundCount.Add id, 0

                            cnt = CLng(foundCount(id)) + 1
                            foundCount(id) = cnt

                            If cnt = 1 Then
                                ' info: (0)=SourceType, (1)=SheetName, (2)=Row, (3)=ColAbgerechnet
                                toUpdate(id) = Array("LOCAL", ws.name, r, colAbg)
                            ElseIf cnt = 2 Then
                                ambiguous.Add id
                                If toUpdate.Exists(id) Then toUpdate.Remove id
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next ws
End Sub

' Zeigt ggf. eine Sammelmeldung zu nicht eindeutigen oder fehlenden IDs
' und liefert False, falls NICHT geschrieben werden soll.
Private Function ShowAmbiguousOrMissingIdsAndAbort(ByVal dictSelected As Object, _
                                                   ByVal ambiguous As Collection, _
                                                   ByVal missing As Collection) As Boolean

    If ambiguous.Count = 0 And missing.Count = 0 Then
        ShowAmbiguousOrMissingIdsAndAbort = True
        Exit Function
    End If

    Dim msg As String: msg = ""
    Dim k As Variant

    If ambiguous.Count > 0 Then
        msg = msg & "Nicht eindeutige Zeilen-ID(s):" & vbCrLf
        For Each k In ambiguous
            msg = msg & "  - " & k & vbCrLf
        Next
    End If

    If missing.Count > 0 Then
        msg = msg & IIf(Len(msg) > 0, vbCrLf, "") & "Nicht gefunden in 'MA '-Sheets:" & vbCrLf
        For Each k In missing
            msg = msg & "  - " & k & vbCrLf
        Next
    End If

    MsgBox msg, vbExclamation, "Abbruch – nichts zurückgeschrieben"
    ShowAmbiguousOrMissingIdsAndAbort = False
End Function

' Setzt für alle zu aktualisierenden Einträge den Wert ABGERECHNET_MARKER in die Spalte "abgerechnet".
Private Sub ApplyAbgerechnetUpdates(ByVal toUpdate As Object, ByRef countSet As Long)
    Dim k As Variant
    Dim info As Variant

    Dim updatesByFile As Object
    Set updatesByFile = CreateObject("Scripting.Dictionary")

    Dim wsLocal As Worksheet
    Dim filePath As String
    Dim items As Collection

    countSet = 0

    ' 1) Lokale Updates direkt schreiben, externe Updates zwischenspeichern
    For Each k In toUpdate.Keys
        info = toUpdate(k)

        Select Case UCase$(CStr(info(0)))
            Case "LOCAL"
                ' info: (0)=SourceType, (1)=SheetName, (2)=Row, (3)=ColAbgerechnet
                Set wsLocal = ThisWorkbook.Worksheets(CStr(info(1)))
                wsLocal.Cells(CLng(info(2)), CLng(info(3))).Value = ABGERECHNET_MARKER
                countSet = countSet + 1

            Case "EXTERNAL"
                ' info: (0)=SourceType, (1)=FilePath, (2)=SheetName, (3)=Row, (4)=ColAbgerechnet
                filePath = CStr(info(1))

                If Not updatesByFile.Exists(filePath) Then
                    Set items = New Collection
                    updatesByFile.Add filePath, items
                End If

                updatesByFile(filePath).Add info
        End Select
    Next k

    ' 2) Externe Dateien in einer zweiten Excel-Instanz aktualisieren
    If updatesByFile.Count > 0 Then
        Dim xlApp As Object
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim i As Long

        Set xlApp = CreateObject("Excel.Application")
        xlApp.visible = False
        xlApp.DisplayAlerts = False

        On Error GoTo ExternalCleanUp

        Dim filePathAsKey As Variant
        For Each filePathAsKey In updatesByFile.Keys
            Set wb = xlApp.Workbooks.Open(Filename:=filePathAsKey, ReadOnly:=False)

            Set items = updatesByFile(filePath)
            For i = 1 To items.Count
                info = items(i)
                Set ws = wb.Worksheets(CStr(info(2)))
                ws.Cells(CLng(info(3)), CLng(info(4))).Value = ABGERECHNET_MARKER
                countSet = countSet + 1
            Next i

            wb.Close SaveChanges:=True
            Set wb = Nothing
        Next filePathAsKey

ExternalCleanUp:
        On Error Resume Next
        If Not wb Is Nothing Then wb.Close SaveChanges:=True
        If Not xlApp Is Nothing Then
            xlApp.DisplayAlerts = True
            xlApp.Quit
        End If
        Set xlApp = Nothing
    End If
End Sub

Private Sub ProcessExternalMaFilesForUpdate(ByVal folder As Object, _
                                            ByVal xlApp As Object, _
                                            ByVal dictSelected As Object, _
                                            ByVal hdrRow As Long, _
                                            ByRef toUpdate As Object, _
                                            ByRef ambiguous As Collection, _
                                            ByRef foundCount As Object)
    Dim file As Object
    Dim subFolder As Object

    ' Alle Dateien im aktuellen Ordner
    For Each file In folder.Files
        If LCase$(file.name) Like "ma_*.xlsx" Then
            ProcessSingleMaWorkbookForUpdate CStr(file.Path), xlApp, dictSelected, hdrRow, toUpdate, ambiguous, foundCount
        End If
    Next file

    ' Unterordner rekursiv
    For Each subFolder In folder.SubFolders
        ProcessExternalMaFilesForUpdate subFolder, xlApp, dictSelected, hdrRow, toUpdate, ambiguous, foundCount
    Next subFolder
End Sub

Private Sub ProcessSingleMaWorkbookForUpdate(ByVal filePath As String, _
                                             ByVal xlApp As Object, _
                                             ByVal dictSelected As Object, _
                                             ByVal hdrRow As Long, _
                                             ByRef toUpdate As Object, _
                                             ByRef ambiguous As Collection, _
                                             ByRef foundCount As Object)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim colId As Long
    Dim colAbg As Long
    Dim lastRow As Long
    Dim r As Long
    Dim id As String
    Dim cnt As Long

    On Error GoTo LocalCleanUp

    Set wb = xlApp.Workbooks.Open(Filename:=filePath, ReadOnly:=False)

    For Each ws In wb.Worksheets
        If Left$(ws.name, Len(WORKSHEET_PREFIX_TO_COLLECT)) = WORKSHEET_PREFIX_TO_COLLECT Then
            colId = Utils.FindHeaderCol(ws, hdrRow, HEADER_ZEILEN_ID)
            colAbg = Utils.FindHeaderCol(ws, hdrRow, HEADER_ABGERECHNET)

            If colId <> 0 And colAbg <> 0 Then
                lastRow = Utils.FindLastUsedRow(ws)

                For r = hdrRow + 1 To lastRow
                    id = Trim$(CStr(ws.Cells(r, colId).Value))
                    If Len(id) > 0 Then
                        If dictSelected.Exists(id) Then
                            If Not foundCount.Exists(id) Then foundCount.Add id, 0

                            cnt = CLng(foundCount(id)) + 1
                            foundCount(id) = cnt

                            If cnt = 1 Then
                                ' info: (0)=SourceType, (1)=FilePath, (2)=SheetName, (3)=Row, (4)=ColAbgerechnet
                                toUpdate(id) = Array("EXTERNAL", filePath, ws.name, r, colAbg)
                            ElseIf cnt = 2 Then
                                ambiguous.Add id
                                If toUpdate.Exists(id) Then toUpdate.Remove id
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next ws

LocalCleanUp:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing
End Sub


