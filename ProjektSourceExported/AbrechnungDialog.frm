VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AbrechnungDialog 
   Caption         =   "Mandaten abrechnen"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11010
   OleObjectBlob   =   "AbrechnungDialog.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "AbrechnungDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IgnoreEvents As Boolean

Private Sub btnCheckUniqueName_Click()
    Dim key As String, byName As Boolean
    Dim nameIsUnique As Boolean

    If lstName.ListIndex >= 0 Then
        key = CStr(lstName.List(lstName.ListIndex, 0))
        byName = True
    End If

    If Len(key) = 0 Then
        MsgBox "Bitte einen Namen ausw‰hlen.", vbExclamation
        Exit Sub
    End If

    ' Start der Pr¸fung
    Me.Hide
    nameIsUnique = Helper.IsUniqueMDName(key)
    If Not nameIsUnique Then
        MsgBox "Der Name '" & key & "' ist nicht eindeutig einer MD-Nr. zugeordnet.", vbExclamation
    End If
    
    Me.Show
End Sub

Private Sub btnWriteBack_Click()
    FormExitCode = EXITCODE_WRITEBACK
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ToggleMode
End Sub

Private Sub optMdId_Click()
    If IgnoreEvents Then Exit Sub
    ToggleMode
    txtMdId.SetFocus
End Sub

Private Sub optName_Click()
    If IgnoreEvents Then Exit Sub
    ToggleMode
    lstName.SetFocus
End Sub

Private Sub lstName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnSearch_Click
End Sub

Private Sub btnSearch_Click()
    Dim key As String, byName As Boolean

    If optMdId.Value Then
        key = Trim$(txtMdId.text)
        byName = False
    Else
        If lstName.ListIndex >= 0 Then
            key = CStr(lstName.List(lstName.ListIndex, 0))
            byName = True
        End If
    End If

    If Len(key) = 0 Then
        MsgBox "Bitte eine MD-ID eingeben oder einen Namen ausw‰hlen.", vbExclamation
        Exit Sub
    End If

    ' Start der Verarbeitung
    Unload Me
    Call Abrechnung.SucheNichtAbgerechnetePositionen(key, byName)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub ToggleMode()
    ' Aktiv/Passiv schalten
    txtMdId.Enabled = optMdId.Value
    lstName.Enabled = optName.Value

    ' Farben setzen: aktiv = weiﬂ, inaktiv = hellgrau
    If txtMdId.Enabled Then
        txtMdId.BackColor = RGB(255, 255, 255)       ' weiﬂ
    Else
        txtMdId.BackColor = RGB(230, 230, 230)       ' hellgrau
    End If

    If lstName.Enabled Then
        lstName.BackColor = RGB(255, 255, 255)       ' weiﬂ
    Else
        lstName.BackColor = RGB(230, 230, 230)       ' hellgrau
    End If
    
    btnCheckUniqueName.Enabled = optName.Value
End Sub

