VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmComarcas 
   Caption         =   "Sísifo - Escolha as intimações a capturar"
   ClientHeight    =   5340
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4665
   OleObjectBlob   =   "frmComarcas.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmComarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bolTodos As Boolean

Private Sub lsComarcas_Change()
    Dim intCont As Integer
    
    If lsComarcas.Selected(0) <> bolTodos Then ' Se a opção "Todos" mudou (foi marcada ou desmarcada)...
        bolTodos = lsComarcas.Selected(0) '... atualiza o estado da variável que marca...
        For intCont = 0 To lsComarcas.ListCount - 1 '... e deixa todos no mesmo estado que a opção "Todos".
            lsComarcas.Selected(intCont) = bolTodos
        Next intCont
    End If
End Sub

Private Sub cmdIr_Click()
    chbDeveGerar.Value = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = 1
    chbDeveGerar.Value = False
    Me.Hide
End Sub
