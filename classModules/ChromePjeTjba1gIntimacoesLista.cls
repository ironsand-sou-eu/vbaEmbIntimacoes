VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromePjeTjba1gIntimacoesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IControladorListadeIntimacoes

Dim chromePjeInt As ChromePjeTjba1gIntimacoes

Private Sub Class_Initialize()
    Set chromePjeInt = New ChromePjeTjba1gIntimacoes
End Sub

'Propriedade ChromeIntimacoes:
Public Property Get IControladorListadeIntimacoes_ChromeIntimacoes() As Variant
Set IControladorListadeIntimacoes_ChromeIntimacoes = chromePjeInt
End Property

Public Function IControladorListadeIntimacoes_PegarListaIntimacoes(Optional periodo As PeriodoDeTempo = Nothing, Optional Comarca As String = "") As Collection
    Dim colErro As Collection, colInt1g As Collection
    Dim resultado As String
    Dim ultimaPagina As Integer, paginaAtual As Integer
    
    chromePjeInt.AbrirAbaExpedientes
    Set colInt1g = chromePjeInt.ColetarIntimacoes("Pendentes de ciência ou de resposta", Comarca)
    
    Set IControladorListadeIntimacoes_PegarListaIntimacoes = colInt1g
End Function

Public Function IControladorListadeIntimacoes_PegarComarcasDisponiveis(nomeCaixaAColetar As String) As Collection
    Dim idAnchorCaixa As String, nomeComarca As String
    Dim i As Variant
    Dim colIdsTableComarcas As Collection, colComarcas As New Collection
    
    idAnchorCaixa = chromePjeInt.PegarIdAnchorCaixa(nomeCaixaAColetar)
    chromePjeInt.AbrirCaixa idAnchorCaixa
    Set colIdsTableComarcas = chromePjeInt.ColetarIdsDasTablesComarcas(idAnchorCaixa)
    
    For Each i In colIdsTableComarcas
        nomeComarca = chromePjeInt.PegarNomeComarcaPeloIdDaTabela(CStr(i))
        colComarcas.Add nomeComarca
    Next i
    Set IControladorListadeIntimacoes_PegarComarcasDisponiveis = colComarcas
End Function
