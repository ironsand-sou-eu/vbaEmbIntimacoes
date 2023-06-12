VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUsuarios 
   Caption         =   "Sísifo - Seleção de usuário"
   ClientHeight    =   5040
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4740
   OleObjectBlob   =   "frmUsuarios.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIr_Click()
    chbDeveGerar.Value = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
End Sub
