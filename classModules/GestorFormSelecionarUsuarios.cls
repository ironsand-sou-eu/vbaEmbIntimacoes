VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestorFormSelecionarUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function SelecionarNomeDoNovoUsuario(Optional nomeUsuarioAtual As String = "") As String
    Dim usuariosDisponiveis As New Collection
    Dim form As frmUsuarios
    Dim nomeUsuarioSelecionado As String
    
    Set usuariosDisponiveis = PegarUsuariosComSenhaCadastradaNaRange(cfPerfisProjudi.Range("PerfisProjudi"))
    Set form = AjustarEMostrarFormAdvsParaUsuariosComSenha(usuariosDisponiveis, nomeUsuarioAtual)
    If form.chbDeveGerar.Value = False Then Exit Function
    nomeUsuarioSelecionado = form.lsUsuarios.List(form.lsUsuarios.ListIndex)
    SelecionarNomeDoNovoUsuario = nomeUsuarioSelecionado
End Function

Function PegarUsuariosComSenhaCadastradaNaRange(rng As Range) As Collection

    Dim qtdPerfis As Long, i As Long
    Dim nomeUsuario As String, senha As String
    Dim perfisComSenha As New Collection
    
    qtdPerfis = rng.Rows.Count
    For i = 1 To qtdPerfis
        nomeUsuario = rng.Cells(i, 2).text
        senha = rng.Cells(i, 3).text
        
        If nomeUsuario <> "" And senha <> "" Then
            perfisComSenha.Add rng.Cells(i, 1).text
        End If
    Next i
    
    Set PegarUsuariosComSenhaCadastradaNaRange = perfisComSenha
    
End Function

Function AjustarEMostrarFormAdvsParaUsuariosComSenha(usuariosDisponiveis As Collection, Optional nomeUsuarioAtual As String = "") As frmUsuarios
    Dim form As New frmUsuarios
    Dim i As Variant
    
    With form
        For Each i In usuariosDisponiveis
            .lsUsuarios.AddItem i
        Next i
        
        For i = 0 To UBound(.lsUsuarios.List) Step 1
            If .lsUsuarios.List(i) = nomeUsuarioAtual Then
                .lsUsuarios.ListIndex = i
                Exit For
            End If
        Next i
        
        .Show
    End With
    
    Set AjustarEMostrarFormAdvsParaUsuariosComSenha = form
End Function


