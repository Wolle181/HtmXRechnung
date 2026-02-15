VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Wartebox 
   Caption         =   "Bitte warten..."
   ClientHeight    =   1095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   OleObjectBlob   =   "Wartebox.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Wartebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowToast(msg As String)
    Wartebox.lblWaitmessage.caption = msg
    Application.StatusBar = msg
    
    If Not ENTWICKLERMODE() Then
        Application.Cursor = xlWait
    End If
    
    Wartebox.Show
End Sub

Public Sub CloseToast()
    Unload Me
    Application.StatusBar = False
    Application.Cursor = xlDefault
End Sub

