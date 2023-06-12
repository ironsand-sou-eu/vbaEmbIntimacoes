VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeProjudiIntimacoesColeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IControladorCadastrarIntimacoes

Dim chromeProjudiInt As ChromeProjudiIntimacoes

Private Sub Class_Initialize()
    Set chromeProjudiInt = New ChromeProjudiIntimacoes
End Sub

Public Sub IControladorCadastrarIntimacoes_CarregarInformacoesWeb(ByRef andamento As andamento)
    Dim LinhaAndamento As Selenium.WebElement
    Dim resultado As String
    Dim TentarDeNovo As Boolean
    
    Do
        resultado = chromeProjudiInt.AbrirProcesso(andamento.numeroProcesso)
        If resultado = "Processo não encontrado" Then TentarDeNovo = Not TentarDeNovo
    Loop While TentarDeNovo = True
    
    If resultado <> "sucesso" Then
        andamento.mensagemErro = resultado
        ApresentarMensagensDeErro resultado
        Exit Sub
    End If
    
    Set LinhaAndamento = PegarElementoAndamentoCorreto
    andamento.numeroProcesso = chromeProjudiInt.PegarNumeroProcesso
    andamento.nomeAndamento = PegarNomeAndamentoAdaptadoAoSisifo(LinhaAndamento)
    andamento.DataAndamento = PegarDataAndamento(andamento.nomeAndamento, LinhaAndamento)
    andamento.obsAndamento = PegarObservacaoAndamento(LinhaAndamento)
    If HouveMudancaJurisdicaoAposEvento(LinhaAndamento, 6) = True Then andamento.NovoJuizo = chromeProjudiInt.PegarJuizo
End Sub

Private Function PegarDataAndamento(nomeAndamento As String, LinhaAndamento As Selenium.WebElement) As Date
    Dim resposta As Date
    
    If InStr(1, LCase(nomeAndamento), "audiência ") <> 0 And InStr(1, LCase(nomeAndamento), "cancelada") = 0 Then
        resposta = PegarDataAudiencia(LinhaAndamento)
    ElseIf nomeAndamento = "Pauta de Julgamento" Then
        resposta = PegarDataSessaoJulgamento(LinhaAndamento)
    ElseIf nomeAndamento = "Incluído em Sessão de Videoconferência" Then
        resposta = PegarDataSessaoJulgamentoIncluidoVideoconferencia(LinhaAndamento)
    Else
        resposta = Trim(LinhaAndamento.FindElementByTag("b").FindElementByXPath("./../..").FindElementsByTag("td")(3).text)
    End If
    
    PegarDataAndamento = resposta
End Function

Private Sub ApresentarMensagensDeErro(resposta As String)
    Select Case resposta
    Case "Processo não encontrado"
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", o processo não foi encontrado. Confira se a parte " & _
        "realmente tem o CNPJ da Embasa (13.504.675/0001-10); se não, escolha o login de advogado e tente novamente.", _
        vbOKOnly + vbCritical, "Sísifo - Processo não encontrado"
        
    Case "processo em segredo de justiça"
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", o processo está em segredo de justiça. Tente novamente " & _
        "com um usuário com acesso.", vbCritical + vbOKOnly, "Sísifo - Processo em segredo de justiça"
        
    Case "sem acesso ao processo"
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", eu não consigo acessar o processo através deste perfil. " & _
        "Isso pode ser por um problema momentâneo, ou por um sistema de bloqueio de automações. Tente novamente " & _
        "com outro usuário e, caso também não consiga, tente manualmente.", vbCritical + vbOKOnly, _
        "Sísifo - Sem acesso ao processo"
        
    End Select
End Sub

Private Function PegarElementoAndamentoCorreto() As Selenium.WebElement
    Dim linhasAndamentos As Selenium.WebElements
    Dim LinhaAndamento As Selenium.WebElement
    Dim tbTabelaAndamentos As HTMLTable
    
    Set linhasAndamentos = chromeProjudiInt.PegarLinhasAndamentosDoProcesso
    Set LinhaAndamento = chromeProjudiInt.PerguntarAndamentoCorreto(linhasAndamentos)
    If LinhaAndamento Is Nothing Then
        MsgBox "Mestre, você não selecionou um andamento válido! O programa será interrompido. Tente novamente!", vbCritical + vbOKOnly, "Sísifo - Erro na seleção de célula"
    End If
    
    Set PegarElementoAndamentoCorreto = LinhaAndamento
    
End Function

Private Function AdaptarNomeAndamentoANomesCompostosSisifo(LinhaAndamento As Selenium.WebElement, nomeAndamento As String) As String
    Dim conteudoCelulaCompleto As String
    
    conteudoCelulaCompleto = LinhaAndamento.FindElementByTag("b").FindElementByXPath("./..").text
    
    '' No Sísifo, os nomes das partes foram substituídos por "XXX".
    If Left(nomeAndamento, 22) = "Conhecido o recurso de" Then
        nomeAndamento = AdaptarNomeAndamentoAConhecidoORecurso(nomeAndamento)
    ElseIf Left(nomeAndamento, 31) = "Conhecido em parte o recurso de" Then
        nomeAndamento = AdaptarNomeAndamentoAConhecidoEmParte(nomeAndamento)
    ElseIf Left(nomeAndamento, 30) = "Não Concedida a Medida Liminar" Then
        nomeAndamento = "Não Concedida a Medida Liminar a XXX"
    ElseIf Left(nomeAndamento, 25) = "Não recebido o recurso de" Then
        nomeAndamento = "Não recebido o recurso de XXX"
    ElseIf Left(nomeAndamento, 22) = "Incluído em pauta para" Then
        nomeAndamento = "Incluído em pauta para XXX"
    ElseIf Left(nomeAndamento, 43) = "Rejeitada a exceção de pré-executividade de" Then
        nomeAndamento = "Rejeitada a exceção de pré-executividade de XXX"
    ElseIf InStr(1, conteudoCelulaCompleto, "Contadoria") <> 0 Then
        nomeAndamento = "Recebidos os autos da contadoria"
    End If
    
    AdaptarNomeAndamentoANomesCompostosSisifo = nomeAndamento
End Function

Private Function AdaptarNomeAndamentoAConhecidoORecurso(nomeAndamento As String) As String
    Dim partesTexto() As String
    
    partesTexto = Split(nomeAndamento, " ")
    Select Case partesTexto(UBound(partesTexto) - 1) & " " & partesTexto(UBound(partesTexto))
    Case "e não-provido"
        nomeAndamento = "Conhecido o recurso de XXX e não-provido"
    Case "e provido"
        nomeAndamento = "Conhecido o recurso de XXX e provido"
    Case "em parte"
        nomeAndamento = "Conhecido o recurso de XXX e provido em parte"
    End Select
    AdaptarNomeAndamentoAConhecidoORecurso = nomeAndamento
End Function

Private Function AdaptarNomeAndamentoAConhecidoEmParte(nomeAndamento As String) As String
    Dim partesTexto() As String
    Select Case partesTexto(UBound(partesTexto) - 1) & " " & partesTexto(UBound(partesTexto))
    Case "e não-provido"
        nomeAndamento = "Conhecido em parte o recurso de XXX e não-provido"
    Case "e provido"
        nomeAndamento = "Conhecido em parte o recurso de XXX e provido"
    Case "em parte"
        nomeAndamento = "Conhecido em parte o recurso de XXX e provido em parte"
    End Select
    AdaptarNomeAndamentoAConhecidoEmParte = nomeAndamento
End Function

Private Function PegarDataAudiencia(LinhaAndamento As Selenium.WebElement) As String
    Dim conteudoCelulaCompleto As String, textoAudiencia As String
    Dim inicioData As Integer
    
    conteudoCelulaCompleto = LinhaAndamento.FindElementByTag("b").FindElementByXPath("./..").text
    inicioData = InStr(1, conteudoCelulaCompleto, "(Agendada")
    textoAudiencia = Trim(Mid(conteudoCelulaCompleto, inicioData))
    textoAudiencia = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(textoAudiencia)
    PegarDataAudiencia = CStr(textoAudiencia)
End Function

Private Function PegarDataSessaoJulgamento(LinhaAndamento As Selenium.WebElement) As String
''
'' Se não encontrar data, alerta usuário e retorna "não há". Nos demais casos, trata e retorna a data da sessão.
''
    Dim textoSessao As String, partesTexto() As String
    Dim i As Integer
    
    textoSessao = chromeProjudiInt.PegarNomeAndamento(LinhaAndamento)
    textoSessao = Replace(textoSessao, "Incluído em pauta para ", "")
    partesTexto() = Split(textoSessao, " ")
    textoSessao = ""
    For i = 0 To UBound(partesTexto) - 3 Step 1 ' Retira as 3 últimas palavras (enésima turma recursal)
        textoSessao = textoSessao & partesTexto(i) & " "
    Next i
    textoSessao = Trim(textoSessao)
    textoSessao = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(textoSessao)
    PegarDataSessaoJulgamento = textoSessao
End Function

Private Function PegarDataSessaoJulgamentoIncluidoVideoconferencia(LinhaAndamento As Selenium.WebElement) As String
''
''Se não encontrar data, alerta usuário e retorna "não há". Nos demais casos, trata e retorna a data da sessão.
''
    Dim textoSessao As String
    Dim posicaoInicial As Integer, posicaoFinal As Integer
    
    textoSessao = PegarTextoCelulaDoNomeAndamento(LinhaAndamento)
    posicaoInicial = InStr(1, textoSessao, " do dia ") + 8
    posicaoFinal = InStr(1, textoSessao, " (Videoconferência)")
    textoSessao = Mid(textoSessao, posicaoInicial, posicaoFinal - posicaoInicial)
    
    textoSessao = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(textoSessao)
    PegarDataSessaoJulgamentoIncluidoVideoconferencia = textoSessao
End Function

Private Function PegarObservacaoAndamento(LinhaAndamento As Selenium.WebElement) As String
    Dim linkArquivo As Selenium.WebElement
    Dim textoAndamento As String
    
    On Error Resume Next
    Set linkArquivo = LinhaAndamento.FindElementByPartialLinkText("online.html", 0)
    On Error GoTo 0
    If Not linkArquivo Is Nothing Then 'And chromeProjudiInt.ChromeProjudi.PerfilLogado = "Representante" Then
        textoAndamento = PegarTextoAndamentoHtml(linkArquivo)
    End If
    textoAndamento = PegarTextoAndamentoObservacao(LinhaAndamento) & vbLf & textoAndamento
    PegarObservacaoAndamento = RetirarEspacosMultiplosEQuebrasDeLinhaVazias(textoAndamento)
End Function

Private Function PegarTextoAndamentoHtml(tagAArquivo As Selenium.WebElement) As String
    Dim urlAndamento As String, textoAndamento As String
    Dim outroChrome As Selenium.ChromeDriver
    Dim cookiesProjudi As Selenium.Cookies, i As Selenium.Cookie
    
    urlAndamento = tagAArquivo.Attribute("href") & "&vis=true"
    Set cookiesProjudi = chromeProjudiInt.PegarCookies
    Set outroChrome = New Selenium.ChromeDriver
    outroChrome.Get SisifoEmbasaFuncoes.sfUrlProjudiRecursoPequenoEstatico
    For Each i In cookiesProjudi
        outroChrome.Manage.AddCookie Name:=i.Name, Value:=i.Value, domain:=i.domain, Path:=i.Path, Secure:=i.Secure
    Next i
    outroChrome.Get urlAndamento
    textoAndamento = outroChrome.FindElementByTag("body").text
    PegarTextoAndamentoHtml = textoAndamento
End Function

Private Function PegarTextoAndamentoObservacao(LinhaAndamento As Selenium.WebElement) As String
    PegarTextoAndamentoObservacao = Trim(LinhaAndamento.FindElementByXPath("./td/span[1]").text)
End Function

Private Function PegarTextoCelulaDoNomeAndamento(LinhaAndamento As Selenium.WebElement) As String
    PegarTextoCelulaDoNomeAndamento = Trim(LinhaAndamento.FindElementByXPath("./td/table/tbody/tr/td[2]", 0).text)
End Function

Private Function RetirarEspacosMultiplosEQuebrasDeLinhaVazias(txt As String) As String
    Do Until InStr(1, txt, "  ") = 0
        txt = Replace(txt, "  ", " ")
    Loop
    
    Do Until InStr(1, txt, vbLf & " " & vbLf) = 0
        txt = Replace(txt, vbLf & " " & vbLf, vbLf)
    Loop
    
    Do Until InStr(1, txt, vbLf & vbLf) = 0
        txt = Replace(txt, vbLf & vbLf, vbLf)
    Loop
    
    RetirarEspacosMultiplosEQuebrasDeLinhaVazias = Trim(txt)
End Function

Private Function HouveMudancaJurisdicaoAposEvento(LinhaAndamento As Selenium.WebElement, qtdEventosAnterioresParaConferir As Integer) As Boolean
    Dim linha As Selenium.WebElement, todasAsLinhas As Selenium.WebElements
    Dim nomeAndamento As String
    Dim contador As Integer
    Dim encontradaMudancaJurisdicao As Boolean, contandoEventosAnteriores As Boolean
    
    Set todasAsLinhas = LinhaAndamento.FindElementByXPath("./..").FindElementsByClass("linhasisifo")
    
    For Each linha In todasAsLinhas
        If chromeProjudiInt.PegarNomeAndamento(linha) = "Redistribuído por Juiz Específico" Then
            encontradaMudancaJurisdicao = True
            Exit For
        ElseIf linha.Equals(LinhaAndamento) Then
            contandoEventosAnteriores = True
        ElseIf contandoEventosAnteriores = True Then
            contador = contador + 1
            If contador >= qtdEventosAnterioresParaConferir Then Exit For
        End If
    Next linha
    
    HouveMudancaJurisdicaoAposEvento = encontradaMudancaJurisdicao
    
End Function

Private Function PegarNomeAndamentoAdaptadoAoSisifo(LinhaAndamento As Selenium.WebElement) As String
    Dim nomeAndamento As String
    nomeAndamento = chromeProjudiInt.PegarNomeAndamento(LinhaAndamento)
    nomeAndamento = AdaptarNomeAndamentoANomesCompostosSisifo(LinhaAndamento, nomeAndamento)
    PegarNomeAndamentoAdaptadoAoSisifo = nomeAndamento
End Function
