VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAcessadorDeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function PegarUsuariosDisponiveisParaLogin(rngOrigem As Range) As String()
    
End Function

Function PegarNomeUsuarioAtual() As String

End Function

Public Function PegarUsuarioPeloNome(nomeUsuario As String, sistema As SisifoEmbasaFuncoes.sfSistema) As UsuarioProjudiPje

End Function

Public Sub RegistrarDataMaisRecenteDeUsuario(data As Date, nomeUsuario As String, sistema As SisifoEmbasaFuncoes.sfSistema)

End Sub

Public Function PegarSistemaSelecionado() As SisifoEmbasaFuncoes.sfSistema

End Function

Public Function PegarTribunalSelecionado() As SisifoEmbasaFuncoes.sfTribunal

End Function

Public Sub SalvarCookieSessao(Cookies As Selenium.Cookies, nomeUsuario As String)

End Sub

Public Function PegarInformacoesCookieSessao() As Selenium.Dictionary

End Function

Public Sub AlterarUsuarioAtual(novoNomeUsuario As String)

End Sub

Public Sub LimparCookieSessao()

End Sub

Public Function PegarNomeEspaiderCorrespondenteAoNomeAndamentoDoSistemaJudicial(nomeAndamento As String) As String

End Function

Public Function PegarInformacoesEspaiderCorrespondentesAoNomeJuizoDoSistemaJudicial(nomeJuizo As String) As StructJuizo

End Function

Public Function PegarListaDeAndamentosEspaiderPossiveis() As String()

End Function

Public Function PegarListaDeProvidenciasEspaiderPossiveis() As String()

End Function

Public Function PegarListaDeJuizosEspaiderPossiveis() As String()

End Function

Public Function PegarNomeProvidenciaALancarNoAndamento(nomeAndamento As String) As String

End Function

Public Function PegarParametrosProvidenciaPeloNome(nomeProvidencia As String) As StructProvidencia

End Function

Public Function PegarDataFinalPadrao() As Date

End Function

Public Sub SalvarInformacoes(meuAndamento As andamento)

End Sub
