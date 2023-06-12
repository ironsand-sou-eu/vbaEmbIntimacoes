VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeProjudiIntimacoesLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IControladorListadeIntimacoes

Dim chromeProjudiInt As ChromeProjudiIntimacoes

Private Sub Class_Initialize()
    Set chromeProjudiInt = New ChromeProjudiIntimacoes
End Sub

'Propriedade ChromeIntimacoes:
Public Property Get IControladorListadeIntimacoes_ChromeIntimacoes() As Variant
Set IControladorListadeIntimacoes_ChromeIntimacoes = chromeProjudiInt
End Property

Public Function IControladorListadeIntimacoes_PegarListaIntimacoes(Optional periodo As PeriodoDeTempo = Nothing, Optional Comarca As String = "") As Collection
    Dim colErro As Collection, colInt1g As Collection, colInt2g As Collection
    Dim resultado As String
    Dim ultimaPagina As Integer, paginaAtual As Integer
    
    If periodo Is Nothing Then
        periodo.DataInicial = Date - 1
        periodo.Datafinal = Date - 1
    End If
    
    Set colInt1g = PegarListaDeIntimacoesDaInstancia(PrimeiroGrau, periodo)
    Set colInt2g = PegarListaDeIntimacoesDaInstancia(SegundoGrau, periodo)
    Set IControladorListadeIntimacoes_PegarListaIntimacoes = SisifoEmbasaFuncoes.SomarColecoes(colInt1g, colInt2g)
    
End Function

Private Function PegarListaDeIntimacoesDaInstancia(instancia As SisifoEmbasaFuncoes.sfInstancia, periodo As PeriodoDeTempo) As Collection
    Dim resposta As Collection
    Dim ultimaPagina As Integer
    
    chromeProjudiInt.AbrirAbaDeIntimacoesNovas instancia
    ultimaPagina = chromeProjudiInt.DescobrirUltimaPaginaDeIntimacoesNovas
    If ultimaPagina <> 0 Then
        Set resposta = ColetarListaIntimacoesPorPeriodo(ultimaPagina, periodo)
        If resposta.Count = 0 Then Set resposta = Nothing
    Else
        Set resposta = Nothing
    End If
    
    Set PegarListaDeIntimacoesDaInstancia = resposta
End Function

Private Function ColetarListaIntimacoesPorPeriodo(ultimaPagina As Integer, periodo As PeriodoDeTempo) As Collection
''
''  DEPENDENTE DE ORDENAMENTO DA TABELA DE INTIMAÇÕES (basta estar no ordenamento-padrão do Projudi, que é ascendente por data de expedição)
''
    Dim colIntimacoesUsuario As New Collection, colCont As New Collection
    Dim i As Integer
    Dim bolAcabou As Boolean
    
    For i = ultimaPagina To 1 Step -1
        If bolAcabou = True Then Exit For
        If chromeProjudiInt.DescobrirPaginaAtualDeIntimacoesNovas <> i Then chromeProjudiInt.AbrirPaginaDeIntimacoesPorNumero i
        If chromeProjudiInt.DescobrirDataIntimacaoMaisAntigaDaPaginaAtualDeIntimacoesNovas <= periodo.Datafinal Then
            Set colCont = ColetarListaIntimacoesDaPaginaAtual(periodo)
            bolAcabou = chromeProjudiInt.VerificarSePeriodoAcabouNaPaginaAtual(periodo)
            Set colIntimacoesUsuario = SisifoEmbasaFuncoes.SomarColecoes(colIntimacoesUsuario, colCont)
        End If
    Next i
    
    Set ColetarListaIntimacoesPorPeriodo = colIntimacoesUsuario
    
End Function

Private Function ColetarListaIntimacoesDaPaginaAtual(periodo As PeriodoDeTempo) As Collection
''
''  DEPENDENTE DE ORDENAMENTO DA TABELA DE INTIMAÇÕES (basta estar no ordenamento-padrão do Projudi, que é ascendente por data de expedição)
''
    Dim tabIntimacoes As Selenium.WebElements
    Dim colIntimacoes As New Collection
    Dim intimAtual As Intimacao
    Dim totalIntimacoes As Integer, i As Integer
    Dim dataExpedicao As Date
    Dim textoDataIntimacao As String
    
    Set tabIntimacoes = chromeProjudiInt.PegarTabIntimacoes
    totalIntimacoes = tabIntimacoes.Count
    
    For i = totalIntimacoes To 1 Step -1
        If chromeProjudiInt.VerificarSeLinhaEhCabecalhoDaTabelaDeIntimacoes(tabIntimacoes(i)) = False Then
            textoDataIntimacao = chromeProjudiInt.PegarDataDaLinhaIntimacao(tabIntimacoes, i)
            dataExpedicao = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(textoDataIntimacao)
            If dataExpedicao >= periodo.DataInicial And dataExpedicao <= periodo.Datafinal Then
                Set intimAtual = chromeProjudiInt.PegarIntimacaoDaLinha(tabIntimacoes, i, dataExpedicao)
                colIntimacoes.Add intimAtual
                chromeProjudiInt.SinalizarIntimacaoProcessada i
            End If
        End If
    Next i
    
    Set ColetarListaIntimacoesDaPaginaAtual = colIntimacoes
    
End Function

Public Function IControladorListadeIntimacoes_PegarComarcasDisponiveis(nomeCaixaAColetar As String) As Collection
    Set IControladorListadeIntimacoes_PegarComarcasDisponiveis = Nothing
End Function

