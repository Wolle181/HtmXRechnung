Attribute VB_Name = "Abgleich"
Option Explicit

Private Const ABGLEICH_SHEET_NAME As String = "Abgleich"

Public Sub ErzeugeAbgleichSheet(control As IRibbonControl)
    Dim ws As Worksheet
    Dim wsAbgleich As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim mdNrCol As Long
    Dim mdCol As Long
    Dim r As Long
    Dim last As Long
    
    ' Abgleich-Sheet holen oder anlegen
    On Error Resume Next
    Set wsAbgleich = ThisWorkbook.Worksheets(ABGLEICH_SHEET_NAME)
    On Error GoTo 0
    
    If wsAbgleich Is Nothing Then
        Set wsAbgleich = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsAbgleich.Name = ABGLEICH_SHEET_NAME
    End If
    
    ' Vorherigen Inhalt löschen
    wsAbgleich.Cells.Clear
    
    ' Header im Abgleich-Sheet setzen
    wsAbgleich.Cells(1, 1).Value = "MD-Nr"
    wsAbgleich.Cells(1, 2).Value = "MD"
    
    destRow = 2
    
    ' Alle MA-Sheets durchgehen
    For Each ws In ThisWorkbook.Worksheets
        ' Nur Sheets, die mit WORKSHEET_PREFIX_TO_COLLECT beginnen
        If Left$(ws.Name, 3) = WORKSHEET_PREFIX_TO_COLLECT Then
            
            lastRow = FindLastUsedRow(ws)
            If lastRow <= HEADER_ROW Then GoTo NextWs
            
            ' Spalten für MD-Nr und MD über Utils-Funktion finden
            mdNrCol = FindHeaderCol(ws, 1, "MD-Nr")
            mdCol = FindHeaderCol(ws, 1, "MD")
            
            ' Wenn eine der Spalten nicht gefunden wurde -> nächstes Sheet
            If mdNrCol <= 0 Or mdCol <= 0 Then GoTo NextWs
            
            ' Zeilen kopieren: nur MD-Nr und MD
            For r = HEADER_ROW + 1 To lastRow
                ' Nur Zeilen übernehmen, in denen mindestens eine Info steht
                If Trim(CStr(ws.Cells(r, mdNrCol).Value)) <> "" _
                   Or Trim(CStr(ws.Cells(r, mdCol).Value)) <> "" Then
                       
                    wsAbgleich.Cells(destRow, 1).Value = ws.Cells(r, mdNrCol).Value
                    wsAbgleich.Cells(destRow, 2).Value = ws.Cells(r, mdCol).Value
                    destRow = destRow + 1
                End If
            Next r
        End If
NextWs:
    Next ws
    
    ' Hintergrundfarben zurücksetzen
    wsAbgleich.Cells.Interior.ColorIndex = xlColorIndexNone
    
    ' Wenn es Daten gibt, Duplikate entfernen (Kombination MD-Nr + MD)
    If destRow > 2 Then
        With wsAbgleich.Range("A1").CurrentRegion
            .RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
        End With
    End If
    
    ' Leere Zeilen (weder MD-Nr noch MD) löschen
    last = wsAbgleich.Cells(wsAbgleich.Rows.Count, 1).End(xlUp).row
    For r = last To 2 Step -1
        If Trim(CStr(wsAbgleich.Cells(r, 1).Value)) = "" _
           And Trim(CStr(wsAbgleich.Cells(r, 2).Value)) = "" Then
            wsAbgleich.Rows(r).Delete
        End If
    Next r
    
    Utils.FormatHeader wsAbgleich, "A1:B1"
    
    ' Spaltenbreite anpassen
    wsAbgleich.Columns.AutoFit
    wsAbgleich.Rows.AutoFit
    
    ' *** HINTERGRUNDFARBEN ZURÜCKSETZEN ***
    wsAbgleich.Cells.Interior.ColorIndex = xlColorIndexNone
End Sub
