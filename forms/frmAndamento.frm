VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAndamento 
   Caption         =   "Sísifo - Insira os dados do processo"
   ClientHeight    =   7860
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6420
   OleObjectBlob   =   "frmAndamento.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAndamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private colGerenciadoresDeEvento As Collection

Private Sub UserForm_Initialize()
    Dim oControle As MSForms.Control
    Dim oGereEventoExitBotao As SisifoEmbasaFuncoes.GereEventoExitBotao
    Dim oGereEventoExitCxTexto As SisifoEmbasaFuncoes.GereEventoExitCxTexto
    Dim oGereEventoExitCombo As SisifoEmbasaFuncoes.GereEventoExitCombo
    
    Set colGerenciadoresDeEvento = New Collection
    
    SisifoEmbasaFuncoes.AjustarLegendaSemTransicao LabelNumProc
    SisifoEmbasaFuncoes.AjustarLegendaSemTransicao LabelAndamento
    SisifoEmbasaFuncoes.AjustarLegendaSemTransicao LabelDataAndamento
    SisifoEmbasaFuncoes.AjustarLegendaSemTransicao LabelProvidencia
    SisifoEmbasaFuncoes.AjustarLegendaSemTransicao LabelJuizo
    SisifoEmbasaFuncoes.AjustarLegendaSemTransicao LabelObsAndamento
    
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

Private Sub txtDataAndamento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If SisifoEmbasaFuncoes.ValidaNumeros(KeyAscii, Array("/", ":", " ")) = False Then KeyAscii = 0
End Sub

Private Sub txtDataAndamento_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtDataAndamento.text = SisifoEmbasaFuncoes.ValidaData(txtDataAndamento.text)
End Sub
