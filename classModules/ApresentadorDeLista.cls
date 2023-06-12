VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApresentadorDeLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IApresentadorDeLista

Public Sub IApresentadorDeLista_PassarListaIntimacoesParaView(grupoIntimacoes As ListaDeIntimacoes)
    Dim view As IVisualizadorDeLista
    Set view = New VisualizadorDeListaEmPlanilha
    
    With grupoIntimacoes
        Select Case .sistema
        Case SisifoEmbasaFuncoes.sfSistema.projudi
            .Origem = .usuario.nomeUsuario
        Case SisifoEmbasaFuncoes.sfSistema.pje1g
            .Origem = .Comarca
        End Select
    End With
    
    view.EntregarRespostaIntimacoes grupoIntimacoes
End Sub
