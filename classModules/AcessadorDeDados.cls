VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AcessadorDeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAcessadorDeDados

Public Function IAcessadorDeDados_PegarUsuariosDisponiveisParaLogin(rngOrigem As Range) As String()
    Dim qtdPerfis As Long, i As Long
    Dim perfis() As String
    
    qtdPerfis = rngOrigem.Rows.Count
    ReDim perfis(1 To qtdPerfis, 1 To 2)
    For i = 1 To qtdPerfis
        perfis(i, 1) = rngOrigem.Cells(i, 1)
        perfis(i, 2) = rngOrigem.Cells(i, 1).Offset(0, 3)
    Next i
    
    IAcessadorDeDados_PegarUsuariosDisponiveisParaLogin = perfis()
End Function

Public Function IAcessadorDeDados_PegarUsuarioPeloNome(nomeUsuario As String, sistema As SisifoEmbasaFuncoes.sfSistema) As UsuarioProjudiPje
    Dim usuario As New UsuarioProjudiPje
    Dim plan As Excel.Worksheet
    Dim celulaLogin As Excel.Range, celulaSenha As Excel.Range
    
    Select Case sistema
    Case SisifoEmbasaFuncoes.sfSistema.projudi
        Set plan = cfPerfisProjudi
    Case SisifoEmbasaFuncoes.sfSistema.pje1g, SisifoEmbasaFuncoes.sfSistema.pje2g
        Set plan = cfPerfisPje
    End Select
    
    usuario.nomeUsuario = nomeUsuario
    On Error Resume Next
    Set celulaLogin = plan.Cells().Find(nomeUsuario).Offset(0, 1)
    Set celulaSenha = plan.Cells().Find(nomeUsuario).Offset(0, 2)
    On Error GoTo 0
    
    If Not celulaLogin Is Nothing Then
        usuario.Login = celulaLogin.text
        usuario.senha = celulaSenha.text
    Else
        usuario.Login = ""
        usuario.senha = ""
    End If
    
    Set IAcessadorDeDados_PegarUsuarioPeloNome = usuario
End Function

Public Sub IAcessadorDeDados_RegistrarDataMaisRecenteDeUsuario(data As Date, nomeUsuario As String, sistema As SisifoEmbasaFuncoes.sfSistema)
    Dim plan As Excel.Worksheet
    
    Select Case sistema
    Case SisifoEmbasaFuncoes.sfSistema.projudi
        Set plan = cfPerfisProjudi
    Case SisifoEmbasaFuncoes.sfSistema.pje1g, SisifoEmbasaFuncoes.sfSistema.pje2g
        Set plan = cfPerfisPje
    End Select
    
    plan.Cells().Find(What:=nomeUsuario, LookAt:=xlWhole).Offset(0, 3).Formula = data
End Sub

Public Sub IAcessadorDeDados_SalvarCookieSessao(Cookies As Selenium.Cookies, nomeUsuario As String)
    Dim i As Selenium.Cookie
    
    For Each i In Cookies
        If i.Name = "JSESSIONID" Then
            With cfIntConfigurações.Cells()
                .Find(What:="CspUsuario", LookAt:=xlWhole).Offset(0, 1).Formula = nomeUsuario
                .Find(What:="CspName", LookAt:=xlWhole).Offset(0, 1).Formula = i.Name
                .Find(What:="CspValue", LookAt:=xlWhole).Offset(0, 1).Formula = i.Value
                .Find(What:="CspDomain", LookAt:=xlWhole).Offset(0, 1).Formula = i.domain
                .Find(What:="CspPath", LookAt:=xlWhole).Offset(0, 1).Formula = i.Path
                .Find(What:="CspSecure", LookAt:=xlWhole).Offset(0, 1).Formula = i.Secure
            End With
            Exit For
        End If
    Next i
End Sub

Public Function IAcessadorDeDados_PegarInformacoesCookieSessao() As Selenium.Dictionary
    Dim infoCookieSessao As Selenium.Dictionary
    
    Set infoCookieSessao = New Selenium.Dictionary
    With cfIntConfigurações.Cells()
        infoCookieSessao.Add "Name", .Find(What:="CspName", LookAt:=xlWhole).Offset(0, 1).Formula
        infoCookieSessao.Add "Value", .Find(What:="CspValue", LookAt:=xlWhole).Offset(0, 1).Formula
        infoCookieSessao.Add "domain", .Find(What:="CspDomain", LookAt:=xlWhole).Offset(0, 1).Formula
        infoCookieSessao.Add "Path", .Find(What:="CspPath", LookAt:=xlWhole).Offset(0, 1).Formula
        infoCookieSessao.Add "Secure", .Find(What:="CspSecure", LookAt:=xlWhole).Offset(0, 1).Formula
    End With
    
    If infoCookieSessao("Value") <> "" Then
        Set IAcessadorDeDados_PegarInformacoesCookieSessao = infoCookieSessao
    Else
        Set IAcessadorDeDados_PegarInformacoesCookieSessao = Nothing
    End If
End Function

Public Sub IAcessadorDeDados_LimparCookieSessao()
    With cfIntConfigurações.Cells()
        .Find(What:="CspName", LookAt:=xlWhole).Offset(0, 1).Formula = ""
        .Find(What:="CspValue", LookAt:=xlWhole).Offset(0, 1).Formula = ""
        .Find(What:="CspDomain", LookAt:=xlWhole).Offset(0, 1).Formula = ""
        .Find(What:="CspPath", LookAt:=xlWhole).Offset(0, 1).Formula = ""
        .Find(What:="CspSecure", LookAt:=xlWhole).Offset(0, 1).Formula = ""
    End With
End Sub

Public Function IAcessadorDeDados_PegarNomeUsuarioAtual() As String
    IAcessadorDeDados_PegarNomeUsuarioAtual = cfIntConfigurações.Cells().Find(What:="CspUsuario", LookAt:=xlWhole).Offset(0, 1).Formula
End Function

Public Sub IAcessadorDeDados_AlterarUsuarioAtual(novoNomeUsuario As String)
    Dim nomeUsuarioAtual As String
    
    nomeUsuarioAtual = IAcessadorDeDados_PegarNomeUsuarioAtual
    If nomeUsuarioAtual <> novoNomeUsuario Then
        cfIntConfigurações.Cells().Find(What:="CspUsuario", LookAt:=xlWhole).Offset(0, 1).Formula = novoNomeUsuario
        IAcessadorDeDados_LimparCookieSessao
    End If
End Sub

Public Function IAcessadorDeDados_PegarTribunalSelecionado() As SisifoEmbasaFuncoes.sfTribunal
    Dim valorEncontrado As String
    
    valorEncontrado = cfIntConfigurações.Cells().Find(What:="Tribunal selecionado", LookAt:=xlWhole).Offset(0, 1).Formula
    IAcessadorDeDados_PegarTribunalSelecionado = CLng(valorEncontrado)
End Function

Public Function IAcessadorDeDados_PegarSistemaSelecionado() As SisifoEmbasaFuncoes.sfSistema
    Dim valorEncontrado As String
    
    valorEncontrado = cfIntConfigurações.Cells().Find(What:="Sistema selecionado", LookAt:=xlWhole).Offset(0, 1).Formula
    IAcessadorDeDados_PegarSistemaSelecionado = CLng(valorEncontrado)
End Function

Public Function IAcessadorDeDados_PegarNomeEspaiderCorrespondenteAoNomeAndamentoDoSistemaJudicial(nomeAndamento As String) As String
    Dim rangeAndamento As Excel.Range
    Dim resposta As String
    
    On Error Resume Next
    Set rangeAndamento = cfAndamentos.Cells().Find(What:=nomeAndamento, LookAt:=xlWhole)
    On Error GoTo 0
    
    If Not rangeAndamento Is Nothing Then
        resposta = rangeAndamento.Offset(0, 1).Formula
    Else
        resposta = "Despacho de mero expediente"
    End If
    
    IAcessadorDeDados_PegarNomeEspaiderCorrespondenteAoNomeAndamentoDoSistemaJudicial = resposta
End Function

Public Function IAcessadorDeDados_PegarInformacoesEspaiderCorrespondentesAoNomeJuizoDoSistemaJudicial(nomeJuizo As String) As StructJuizo
    Dim resposta As StructJuizo
    
    With resposta
        If nomeJuizo = "" Then
            .nomeJuizo = ""
            .nomeComarca = ""
            .NomeOrgao = ""
        Else
            .nomeJuizo = cfJuizos.Cells().Find(What:=nomeJuizo, LookAt:=xlWhole).Offset(0, 1).Formula
            .nomeComarca = cfJuizos.Cells().Find(What:=nomeJuizo, LookAt:=xlWhole).Offset(0, 2).Formula
            .NomeOrgao = cfJuizos.Cells().Find(What:=nomeJuizo, LookAt:=xlWhole).Offset(0, 3).Formula
        End If
    End With
    
    IAcessadorDeDados_PegarInformacoesEspaiderCorrespondentesAoNomeJuizoDoSistemaJudicial = resposta
End Function

Public Function IAcessadorDeDados_PegarListaDeAndamentosEspaiderPossiveis() As String()
    IAcessadorDeDados_PegarListaDeAndamentosEspaiderPossiveis = SisifoEmbasaFuncoes.ConverterRangeColunaParaVetorString(cfIntConfigurações.Range("ListaAndamentos"))
End Function

Public Function IAcessadorDeDados_PegarListaDeProvidenciasEspaiderPossiveis() As String()
    IAcessadorDeDados_PegarListaDeProvidenciasEspaiderPossiveis = SisifoEmbasaFuncoes.ConverterRangeColunaParaVetorString(cfIntConfigurações.Range("ListaProvidencias"))
End Function

Public Function IAcessadorDeDados_PegarListaDeJuizosEspaiderPossiveis() As String()
    IAcessadorDeDados_PegarListaDeJuizosEspaiderPossiveis = SisifoEmbasaFuncoes.ConverterRangeColunaParaVetorString(cfJuizos.Range("ListaJuizos"))
End Function

Function IAcessadorDeDados_PegarNomeProvidenciaALancarNoAndamento(nomeAndamento As String) As String
    Dim resposta As String
    
    If Not cfAndamentosProvidencias.Cells().Find(What:=nomeAndamento, LookAt:=xlWhole) Is Nothing Then
        resposta = cfAndamentosProvidencias.Cells().Find(What:=nomeAndamento, LookAt:=xlWhole).Offset(0, 1).Formula
    Else
        resposta = ""
    End If
    IAcessadorDeDados_PegarNomeProvidenciaALancarNoAndamento = resposta
End Function

Function IAcessadorDeDados_PegarParametrosProvidenciaPeloNome(nomeProvidencia As String) As StructProvidencia
    Dim resposta As StructProvidencia
    
    With resposta
        .Nucleo = cfIntConfigurações.Cells().Find(What:=nomeProvidencia, LookAt:=xlWhole).Offset(0, 1).Formula
        .Responsavel = cfIntConfigurações.Cells().Find(What:=nomeProvidencia, LookAt:=xlWhole).Offset(0, 2).Formula
        .Observacao = cfIntConfigurações.Cells().Find(What:=nomeProvidencia, LookAt:=xlWhole).Offset(0, 3).Formula
        .GerarAndamento = IIf(cfIntConfigurações.Cells().Find(What:=nomeProvidencia, LookAt:=xlWhole).Offset(0, 4).Formula = "Sim", True, False)
        .andamento = cfIntConfigurações.Cells().Find(What:=nomeProvidencia, LookAt:=xlWhole).Offset(0, 6).Formula
        .TipoDataBaseAndamento = cfIntConfigurações.Cells().Find(What:=nomeProvidencia, LookAt:=xlWhole).Offset(0, 5).Formula
    End With
    
    IAcessadorDeDados_PegarParametrosProvidenciaPeloNome = resposta
End Function

Function IAcessadorDeDados_PegarDataFinalPadrao() As Date
    IAcessadorDeDados_PegarDataFinalPadrao = CDate(cfIntConfigurações.Cells().Find(What:="Criar providências para", LookAt:=xlWhole).Offset(0, 1).text)
End Function

Public Sub IAcessadorDeDados_SalvarInformacoes(meuAndamento As andamento)

    If NumeroProcessoJaEstaSalvo(meuAndamento) = True Then Exit Sub
    SalvarAndamento meuAndamento
    
    If Trim(meuAndamento.NovoJuizo) <> "" Then LancarJurisdicaoNaMemoria meuAndamento
    SalvarPastaDeTrabalho ThisWorkbook
    
End Sub

Private Function NumeroProcessoJaEstaSalvo(andamento As andamento) As Boolean
    Dim resposta As Boolean
    
    If Not sfCadAndamento.Cells().Find(andamento.numeroProcesso) Is Nothing Then
        resposta = True
        Do
        Loop Until MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", já existe um andamento deste processo na planilha! " & _
            "A inclusão foi cancelada. Descartamos os dados." & vbCrLf & _
            "Processo: " & andamento.numeroProcesso & vbCrLf & _
            "Andamento: " & andamento.nomeAndamento & " em " & andamento.DataAndamento & vbCrLf & _
            "Clique em 'Cancelar' e insira o próximo processo.", vbCritical + vbOKCancel, _
            "Sísifo - Processo repetido") = vbCancel
    Else
        resposta = False
    End If
    NumeroProcessoJaEstaSalvo = resposta
End Function

Private Sub SalvarAndamento(andamento As andamento)
    Dim i As Integer
    Dim linhaSeguinte As Long
    
    With sfCadAndamento
        linhaSeguinte = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        
        .Cells(linhaSeguinte, 1) = "'" & andamento.IdConversao 'Sequencial andamentos
        .Cells(linhaSeguinte, 2) = andamento.numeroProcesso 'Número do Processo
        .Cells(linhaSeguinte, 3) = andamento.numeroProcesso 'Número do desdobramento
        .Cells(linhaSeguinte, 4) = andamento.nomeAndamento 'Nome do andamento
        .Cells(linhaSeguinte, 5) = andamento.DataAndamento 'data
        .Cells(linhaSeguinte, 6) = andamento.obsAndamento 'Obs andamento

        If andamento.Providencias.Count > 0 Then
            For i = 0 To andamento.Providencias.Count - 1 Step 1
                SalvarProvidencia andamento.Providencias(i + 1)
            Next i
        End If
        
        .Cells.WrapText = False
    End With
End Sub

Private Sub SalvarProvidencia(providencia As providencia)
    Dim linhaSeguinte As Long
    
    With sfCadProvidencia
        linhaSeguinte = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        
        .Cells(linhaSeguinte, 1) = providencia.numeroProcesso 'Número do Processo
        .Cells(linhaSeguinte, 2) = providencia.numeroProcesso 'Número do desdobramento
        .Cells(linhaSeguinte, 3) = "'" & providencia.IdAndamento 'Sequencial do andamento
        .Cells(linhaSeguinte, 4) = providencia.Nome 'Providência
        .Cells(linhaSeguinte, 5) = "'" & providencia.IdConversao 'Sequencial da providência
        .Cells(linhaSeguinte, 6) = providencia.Datafinal 'data Final
        .Cells(linhaSeguinte, 7) = providencia.Nucleo 'Núcleo
        .Cells(linhaSeguinte, 8) = providencia.Responsavel 'Responsável
        .Cells(linhaSeguinte, 9) = providencia.PrazoDias 'Prazo
        .Cells(linhaSeguinte, 10) = providencia.Situacao 'Situação
        .Cells(linhaSeguinte, 11) = IIf(providencia.Alertar, "Sim", "Não") 'Enviar alertas
        .Cells(linhaSeguinte, 12) = providencia.PeriodicidadeAlerta 'Periodicidade
        .Cells(linhaSeguinte, 13) = providencia.Observacao 'Observacoes
        .Cells(linhaSeguinte, 14) = IIf(providencia.GerarAndamento, "Sim", "Não") 'Gerar Evento
        .Cells(linhaSeguinte, 15) = providencia.TipoDataBaseAndamento 'data Base
        .Cells(linhaSeguinte, 16) = providencia.AndamentoAGerar 'Evento
    End With
End Sub

Private Sub LancarJurisdicaoNaMemoria(meuAndamento As andamento)
    Dim proximaLinha As Long
    
    With sfCadJurisdicao
        proximaLinha = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(proximaLinha, 1) = meuAndamento.numeroProcesso
        .Cells(proximaLinha, 2) = meuAndamento.NovoOrgao
        .Cells(proximaLinha, 3) = meuAndamento.NovoOrgao
        .Cells(proximaLinha, 4) = meuAndamento.NovoJuizo
        .Cells(proximaLinha, 5) = meuAndamento.NovaComarca
        .UsedRange.Columns.AutoFit
        .Cells.WrapText = False
    End With
End Sub

Private Sub SalvarPastaDeTrabalho(pTrab As Excel.Workbook)
    Do
        Application.DisplayAlerts = False
        'ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, FileFormat:=xlOpenXMLAddIn
        pTrab.Save
        Application.DisplayAlerts = True
    Loop Until pTrab.Saved = True
End Sub
