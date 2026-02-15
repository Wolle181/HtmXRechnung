Attribute VB_Name = "Kontext"
Option Explicit

Public Property Get ENTWICKLERMODE() As Boolean
    ENTWICKLERMODE = Utils.FileExists(Kontext.RootPath & "\Entwicklermode.info")
End Property

Public Property Get wbMain() As Workbook
    Set wbMain = ThisWorkbook
End Property

Public Property Get RootPath() As String
  RootPath = wbMain.Path
End Property

Public Sub ClearContext()
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Application.visible = True
End Sub

Public Property Get TimestampExtension() As String
    TimestampExtension = IIf(WRITEFILESWITHTIMESTAMPEXTENSION, Format(Now, "_yyyymmdd_hhnnss"), "")
End Property

