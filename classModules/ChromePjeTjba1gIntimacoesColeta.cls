VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromePjeTjba1gIntimacoesColeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IControladorCadastrarIntimacoes

Dim chromePjeInt As ChromePjeTjba1gIntimacoes

Private Sub Class_Initialize()
    Set chromePjeInt = New ChromePjeTjba1gIntimacoes
End Sub

Public Sub IControladorCadastrarIntimacoes_CarregarInformacoesWeb(ByRef andamento As andamento)
    Dim LinhaAndamento As Selenium.WebElement
    Dim resultado As String
    Dim TentarDeNovo As Boolean
    
    Do
        resultado = chromePjeInt.ChromePje1g.AbrirProcesso(andamento.numeroProcesso)
        If resultado = "Processo n�o encontrado" Then TentarDeNovo = Not TentarDeNovo
    Loop While TentarDeNovo = True
    
    If resultado <> "sucesso" Then
        andamento.mensagemErro = resultado
        ApresentarMensagensDeErro resultado
        Exit Sub
    End If
    
    Set LinhaAndamento = PegarElementoAndamentoCorreto
    andamento.nomeAndamento = PegarNomeAndamentoAdaptadoAoSisifo(LinhaAndamento)
    andamento.DataAndamento = PegarDataAndamento(andamento.nomeAndamento, LinhaAndamento)
    andamento.obsAndamento = PegarObservacaoAndamento(LinhaAndamento)
    'If HouveMudancaJurisdicaoAposEvento(LinhaAndamento, 6) = True Then andamento.NovoJuizo = chromePjeInt.PegarJuizo
End Sub

Private Function PegarDataAndamento(nomeAndamento As String, LinhaAndamento As Selenium.WebElement) As Date
    Dim resposta As Date
    
    'If InStr(1, LCase(nomeAndamento), "audi�ncia ") <> 0 And InStr(1, LCase(nomeAndamento), "cancelada") = 0 Then
    '    resposta = PegarDataAudiencia(linhaAndamento)
    'ElseIf nomeAndamento = "Pauta de Julgamento" Then
    '    resposta = PegarDataSessaoJulgamento(LinhaAndamento)
    'ElseIf nomeAndamento = "Inclu�do em Sess�o de Videoconfer�ncia" Then
    '    resposta = PegarDataSessaoJulgamentoIncluidoVideoconferencia(LinhaAndamento)
    'Else
        resposta = PegarDataReferenteAAndamento(LinhaAndamento)
    'End If
    
    PegarDataAndamento = resposta
End Function

Private Function PegarDataReferenteAAndamento(LinhaAndamento As Selenium.WebElement) As Date
    Dim resposta As Date
    Dim divAndamentos As Selenium.WebElement
    Dim linhasAndamentos As Selenium.WebElements
    Dim i As Integer
    
    Set divAndamentos = chromePjeInt.PegarDivAndamentos
    Set linhasAndamentos = divAndamentos.FindElementsByXPath("./div")
    
    For i = 1 To linhasAndamentos.Count Step 1 'Descobrir o index de linhaAndamento
        If linhasAndamentos(i).Equals(LinhaAndamento) Then Exit For
    Next i
    
    For i = i - 1 To 1 Step -1 'Ir subindo at� a pr�xima div contendo data
        If InStr(1, linhasAndamentos(i).Attribute("class"), "data") <> 0 Then
            resposta = CDate(linhasAndamentos(i).text)
            Exit For
        End If
    Next i
    
    PegarDataReferenteAAndamento = resposta
End Function

Private Sub ApresentarMensagensDeErro(resposta As String)
    Select Case resposta
    Case "Processo n�o encontrado"
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", o processo n�o foi encontrado. Confira se a parte " & _
        "realmente tem o CNPJ da Embasa (13.504.675/0001-10); se n�o, escolha o login de advogado e tente novamente.", _
        vbOKOnly + vbCritical, "S�sifo - Processo n�o encontrado"
        
    Case "processo em segredo de justi�a"
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", o processo est� em segredo de justi�a. Tente novamente " & _
        "com um usu�rio com acesso.", vbCritical + vbOKOnly, "S�sifo - Processo em segredo de justi�a"
        
    Case "sem acesso ao processo"
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", eu n�o consigo acessar o processo atrav�s deste perfil. " & _
        "Isso pode ser por um problema moment�neo, ou por um sistema de bloqueio de automa��es. Tente novamente " & _
        "com outro usu�rio e, caso tamb�m n�o consiga, tente manualmente.", vbCritical + vbOKOnly, _
        "S�sifo - Sem acesso ao processo"
        
    End Select
End Sub

Private Function PegarElementoAndamentoCorreto() As Selenium.WebElement
    Dim divAndamentos As Selenium.WebElement
    Dim LinhaAndamento As Selenium.WebElement
    Dim tbTabelaAndamentos As HTMLTable
    
    Set divAndamentos = chromePjeInt.PegarDivAndamentos
    Set LinhaAndamento = chromePjeInt.PerguntarAndamentoCorreto(divAndamentos)
    If LinhaAndamento Is Nothing Then
        MsgBox "Mestre, voc� n�o selecionou um andamento v�lido! O programa ser� interrompido. Tente novamente!", vbCritical + vbOKOnly, "S�sifo - Erro na sele��o de c�lula"
    End If
    
    Set PegarElementoAndamentoCorreto = LinhaAndamento
    
End Function

Private Function AdaptarNomeAndamentoANomesCompostosSisifo(LinhaAndamento As Selenium.WebElement, nomeAndamento As String) As String
    Dim conteudoCelulaCompleto As String
    
    'conteudoCelulaCompleto = LinhaAndamento.FindElementByTag("b").FindElementByXPath("./..").text
    
    '' No S�sifo, os nomes das partes foram substitu�dos por "XXX".
    If Left(nomeAndamento, 22) = "Conhecido o recurso de" Then
        nomeAndamento = AdaptarNomeAndamentoAConhecidoORecurso(nomeAndamento)
    ElseIf Left(nomeAndamento, 31) = "Conhecido em parte o recurso de" Then
        nomeAndamento = AdaptarNomeAndamentoAConhecidoEmParte(nomeAndamento)
    ElseIf Left(nomeAndamento, 30) = "N�o Concedida a Medida Liminar" Then
        nomeAndamento = "N�o Concedida a Medida Liminar a XXX"
    ElseIf Left(nomeAndamento, 25) = "N�o recebido o recurso de" Then
        nomeAndamento = "N�o recebido o recurso de XXX"
    ElseIf Left(nomeAndamento, 22) = "Inclu�do em pauta para" Then
        nomeAndamento = "Inclu�do em pauta para XXX"
    ElseIf Left(nomeAndamento, 43) = "Rejeitada a exce��o de pr�-executividade de" Then
        nomeAndamento = "Rejeitada a exce��o de pr�-executividade de XXX"
    'ElseIf InStr(1, conteudoCelulaCompleto, "Contadoria") <> 0 Then
    '    nomeAndamento = "Recebidos os autos da contadoria"
    End If
    
    AdaptarNomeAndamentoANomesCompostosSisifo = nomeAndamento
End Function

Private Function AdaptarNomeAndamentoAConhecidoORecurso(nomeAndamento As String) As String
    Dim partesTexto() As String
    
    partesTexto = Split(nomeAndamento, " ")
    Select Case partesTexto(UBound(partesTexto) - 1) & " " & partesTexto(UBound(partesTexto))
    Case "e n�o-provido"
        nomeAndamento = "Conhecido o recurso de XXX e n�o-provido"
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
    Case "e n�o-provido"
        nomeAndamento = "Conhecido em parte o recurso de XXX e n�o-provido"
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
'' Se n�o encontrar data, alerta usu�rio e retorna "n�o h�". Nos demais casos, trata e retorna a data da sess�o.
''
    Dim textoSessao As String, partesTexto() As String
    Dim i As Integer
    
    'textoSessao = chromePjeInt.PegarNomeAndamento(LinhaAndamento)
    textoSessao = Replace(textoSessao, "Inclu�do em pauta para ", "")
    partesTexto() = Split(textoSessao, " ")
    textoSessao = ""
    For i = 0 To UBound(partesTexto) - 3 Step 1 ' Retira as 3 �ltimas palavras (en�sima turma recursal)
        textoSessao = textoSessao & partesTexto(i) & " "
    Next i
    textoSessao = Trim(textoSessao)
    textoSessao = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(textoSessao)
    PegarDataSessaoJulgamento = textoSessao
End Function

Private Function PegarDataSessaoJulgamentoIncluidoVideoconferencia(LinhaAndamento As Selenium.WebElement) As String
''
''Se n�o encontrar data, alerta usu�rio e retorna "n�o h�". Nos demais casos, trata e retorna a data da sess�o.
''
    Dim textoSessao As String
    Dim posicaoInicial As Integer, posicaoFinal As Integer
    
    textoSessao = PegarTextoCelulaDoNomeAndamento(LinhaAndamento)
    posicaoInicial = InStr(1, textoSessao, " do dia ") + 8
    posicaoFinal = InStr(1, textoSessao, " (Videoconfer�ncia)")
    textoSessao = Mid(textoSessao, posicaoInicial, posicaoFinal - posicaoInicial)
    
    textoSessao = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(textoSessao)
    PegarDataSessaoJulgamentoIncluidoVideoconferencia = textoSessao
End Function

Function PegarObservacaoAndamento(LinhaAndamento As Selenium.WebElement) As String
    Dim linkArquivo As Selenium.WebElement
    Dim resposta As String
    
    On Error Resume Next
    Set linkArquivo = LinhaAndamento.FindElementByClass("anexos", 0).FindElementByTag("a", 0)
    On Error GoTo 0
    
    If Not linkArquivo Is Nothing Then
        linkArquivo.Click
        resposta = PegarTextoDoAndamentoAberto
        resposta = RetirarEspacosMultiplosEQuebrasDeLinhaVazias(resposta)
    Else
        resposta = ""
    End If
    
    PegarObservacaoAndamento = resposta
End Function

Function PegarTextoDoAndamentoAberto() As String
    Dim resposta As String ', urlAndamento As String
    Dim containerTexto As Selenium.WebElement
    
    With chromePjeInt.ChromePje1g.Chrome.Chrome
        .SwitchToFrame 0
        If .FindElementsByTag("div").Count < 10 Then
            Set containerTexto = .FindElementByTag("body")
            resposta = containerTexto.text
        End If
        .SwitchToParentFrame
    End With
    
    PegarTextoDoAndamentoAberto = resposta
End Function

Function PegarTextoCelulaDoNomeAndamento(LinhaAndamento As Selenium.WebElement) As String
    PegarTextoCelulaDoNomeAndamento = Trim(LinhaAndamento.FindElementByXPath("./td/table/tbody/tr/td[2]", 0).text)
End Function

Function RetirarEspacosMultiplosEQuebrasDeLinhaVazias(txt As String) As String
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
        If chromePjeInt.PegarNomeAndamento(linha) = "Redistribu�do por Juiz Espec�fico" Then
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
    nomeAndamento = chromePjeInt.PegarNomeAndamento(LinhaAndamento)
    nomeAndamento = AdaptarNomeAndamentoANomesCompostosSisifo(LinhaAndamento, nomeAndamento)
    PegarNomeAndamentoAdaptadoAoSisifo = nomeAndamento
End Function


