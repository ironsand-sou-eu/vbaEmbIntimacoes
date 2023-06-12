VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdvs 
   Caption         =   "Sísifo - Escolha as intimações a capturar"
   ClientHeight    =   5865
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4740
   OleObjectBlob   =   "frmAdvs.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAdvs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colGerenciadoresDeEvento As Collection
Private bolTodos As Boolean

Private Sub txtDataInicial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If SisifoEmbasaFuncoes.ValidaNumeros(KeyAscii, Array("/", ":", " ")) = False Then KeyAscii = 0
End Sub

Private Sub txtDataInicial_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtDataInicial.text = SisifoEmbasaFuncoes.ValidaData(txtDataInicial.text)
End Sub

Private Sub txtDataFinal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If SisifoEmbasaFuncoes.ValidaNumeros(KeyAscii, Array("/", ":", " ")) = False Then KeyAscii = 0
End Sub

Private Sub txtDataFinal_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtDataFinal.text = SisifoEmbasaFuncoes.ValidaData(txtDataFinal.text)
End Sub

Private Sub lsAdvs_Change()
    Dim intCont As Integer
    
    If lsAdvs.Selected(0) <> bolTodos Then ' Se a opção "Todos" mudou (foi marcada ou desmarcada)...
        bolTodos = lsAdvs.Selected(0) '... atualiza o estado da variável que marca...
        
        For intCont = 0 To lsAdvs.ListCount - 1 '... e deixa todos no mesmo estado que a opção "Todos".
            lsAdvs.Selected(intCont) = bolTodos
        Next intCont
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim oControle As MSForms.Control
    Dim oGereEventoExitBotao As SisifoEmbasaFuncoes.GereEventoExitBotao
    Dim oGereEventoExitCxTexto As SisifoEmbasaFuncoes.GereEventoExitCxTexto
    Dim oGereEventoExitCombo As SisifoEmbasaFuncoes.GereEventoExitCombo
    
    Set colGerenciadoresDeEvento = New Collection
    txtDataInicial.SetFocus
    
    ' Criar gerenciadores de evento para o controle personalizado da barrinha colorida e para simular o evento Exit
    
    For Each oControle In Me.Controls
        If oControle.Visible = True Then
            Select Case TypeName(oControle)
            Case "TextBox"
                If oControle.TabStop = True Then
                    Set oGereEventoExitCxTexto = SisifoEmbasaFuncoes.New_GereEventoExitCxTexto
                    Set oGereEventoExitCxTexto.CxTexto = oControle
                    colGerenciadoresDeEvento.Add oGereEventoExitCxTexto
                End If
                
            Case "CommandButton"
                If oControle.TabStop = True Then
                    Set oGereEventoExitBotao = SisifoEmbasaFuncoes.New_GereEventoExitBotao
                    Set oGereEventoExitBotao.Botao = oControle
                    colGerenciadoresDeEvento.Add oGereEventoExitBotao
                End If
                
            Case "ComboBox"
                If oControle.TabStop = True Then
                    Set oGereEventoExitCombo = SisifoEmbasaFuncoes.New_GereEventoExitCombo
                    Set oGereEventoExitCombo.Combo = oControle
                    colGerenciadoresDeEvento.Add oGereEventoExitCombo
                End If
                
            End Select
        End If
    Next oControle
End Sub

Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)
    
    Select Case TypeName(Control)
    Case "TextBox"
        If Control.TabStop = True Then
            Dim oGereEventoExitCxTexto As SisifoEmbasaFuncoes.GereEventoExitCxTexto
            Set oGereEventoExitCxTexto = SisifoEmbasaFuncoes.New_GereEventoExitCxTexto
            Set oGereEventoExitCxTexto.CxTexto = Control
            colGerenciadoresDeEvento.Add oGereEventoExitCxTexto
        End If
        
    Case "CommandButton"
        If Control.TabStop = True Then
            Dim oGereEventoExitBotao As SisifoEmbasaFuncoes.GereEventoExitBotao
            Set oGereEventoExitBotao = SisifoEmbasaFuncoes.New_GereEventoExitBotao
            Set oGereEventoExitBotao.Botao = Control
            colGerenciadoresDeEvento.Add oGereEventoExitBotao
        End If
        
    Case "ComboBox"
        If Control.TabStop = True Then
            Dim oGereEventoExitCombo As SisifoEmbasaFuncoes.GereEventoExitCombo
            Set oGereEventoExitCombo = SisifoEmbasaFuncoes.New_GereEventoExitCombo
            Set oGereEventoExitCombo.Combo = Control
            colGerenciadoresDeEvento.Add oGereEventoExitCombo
        End If
        
    End Select
End Sub

Private Sub UserForm_Terminate()
    Set SisifoEmbasaFuncoes.oControleAtual = Nothing
    Set SisifoEmbasaFuncoes.oControleAnterior = Nothing
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
