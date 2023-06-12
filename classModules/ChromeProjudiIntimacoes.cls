VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeProjudiIntimacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private chromeProj As ChromeProjudi

Private Sub Class_Initialize()
    Set chromeProj = New ChromeProjudi
End Sub

'Propriedade ChromeProjudi:
Public Property Get ChromeProjudi() As ChromeProjudi
Set ChromeProjudi = chromeProj
End Property
Public Property Set ChromeProjudi(chProj As ChromeProjudi)
Set chromeProj = chProj
End Property

Sub AbrirAbaDeIntimacoesNovas(instancia As SisifoEmbasaFuncoes.sfInstancia)
    Dim por As New Selenium.By
    Dim urlIntimacoes As String
    Dim abaJaAberta As Boolean
    
    Select Case instancia
    Case SisifoEmbasaFuncoes.sfInstancia.PrimeiroGrau
        urlIntimacoes = SisifoEmbasaFuncoes.sfUrlProjudiIntimacoesNovas1gAdvOuRepresentante
    Case SisifoEmbasaFuncoes.sfInstancia.SegundoGrau
        urlIntimacoes = SisifoEmbasaFuncoes.sfUrlProjudiIntimacoesNovas2gAdvOuRepresentante
    End Select
    
    abaJaAberta = (chromeProj.Chrome.Chrome.url = urlIntimacoes) And chromeProj.Chrome.Chrome.IsElementPresent(por.Name("horarioInicio"))
    If Not abaJaAberta Then chromeProj.Chrome.Chrome.Get urlIntimacoes
End Sub

Function DescobrirUltimaPaginaDeIntimacoesNovas() As Integer
    Dim migalhaDePao As Selenium.WebElement
    Dim textoPagina As String, finalDoTexto As String, ultimaPagina As String
    
        On Error Resume Next
        Set migalhaDePao = chromeProj.Chrome.Chrome.FindElementByXPath("/html/body/div/table[1]/tbody/tr[1]/td", 0)
        On Error GoTo 0
        
        If migalhaDePao Is Nothing Then
            DescobrirUltimaPaginaDeIntimacoesNovas = 0
        Else
            textoPagina = migalhaDePao.text
            finalDoTexto = Right(textoPagina, Len(textoPagina) - InStr(1, textoPagina, " de "))
            ultimaPagina = Replace(finalDoTexto, "de ", "")
            DescobrirUltimaPaginaDeIntimacoesNovas = CInt(ultimaPagina)
        End If
        
End Function

Function DescobrirPaginaAtualDeIntimacoesNovas() As Integer

    Dim textoPagina As String, finalDoTexto As String, paginaAtual As String
    
'    Select Case LCase(strPerfilLogado)
'    Case "advogado", "representante"
        textoPagina = chromeProj.Chrome.Chrome.FindElementByXPath("/html/body/div/table[1]/tbody/tr[1]/td").text ' O texto é no formato "256 resultados encontrados - página 1 de 19"
        finalDoTexto = Right(textoPagina, Len(textoPagina) - InStr(1, textoPagina, " - página ") - 8)
        paginaAtual = Trim(Left(finalDoTexto, InStr(1, finalDoTexto, " de ")))
        DescobrirPaginaAtualDeIntimacoesNovas = CInt(paginaAtual)
        
'    Case Else
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve algum erro no login. Suplico que tente novamente, desta vez fazendo login " & _
'        "de representante da empresa ou de advogado (login de parte, aquele com o CNPJ da empresa, não serve).", vbOKOnly + vbCritical, _
'        "Sísifo - Erro no login"
'        Set chromeProj.Chrome.Chrome = Nothing
'        DescobrirPaginaAtualDeIntimacoesNovas = 0
'
'    End Select
    
End Function

Sub AbrirPaginaDeIntimacoesPorNumero(numeroPagina As Integer)
    chromeProj.Chrome.Chrome.ExecuteScript "goToPage(" & numeroPagina & ")"
End Sub

Function DescobrirDataIntimacaoMaisAntigaDaPaginaAtualDeIntimacoesNovas() As Date
    Dim tabIntimacoes As Selenium.WebElements
    Dim colIntimacoes As New Collection
    Dim arrNumData(1 To 2) As String
    Dim totalIntimacoes As Integer, i As Integer
    Dim dtCont As Date
    
    Set tabIntimacoes = PegarTabIntimacoes
    totalIntimacoes = tabIntimacoes.Count
    
    For i = 1 To totalIntimacoes Step 1
        If VerificarSeLinhaEhCabecalhoDaTabelaDeIntimacoes(tabIntimacoes(i)) = False Then
            DescobrirDataIntimacaoMaisAntigaDaPaginaAtualDeIntimacoesNovas = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(Trim(tabIntimacoes(i).FindElementsByTag("td")(4).text))
            Exit Function
        End If
    Next i
End Function

Function VerificarSeLinhaEhCabecalhoDaTabelaDeIntimacoes(linha As Selenium.WebElement) As Boolean
    If linha.Attribute("class") = "" And linha.FindElementsByTag("td").Count > 3 Then
        VerificarSeLinhaEhCabecalhoDaTabelaDeIntimacoes = False
    Else
        VerificarSeLinhaEhCabecalhoDaTabelaDeIntimacoes = True
    End If
End Function

Function VerificarSePeriodoAcabouNaPaginaAtual(periodo As PeriodoDeTempo) As Boolean
''
''  DEPENDENTE DE ORDENAMENTO DA TABELA DE INTIMAÇÕES (basta estar no ordenamento-padrão do Projudi, que é ascendente por data de expedição)
''
    Dim tabIntimacoes As Selenium.WebElements
    Dim totalIntimacoes As Integer, i As Integer
    Dim dtCont As Date
    
    Set tabIntimacoes = PegarTabIntimacoes
    totalIntimacoes = tabIntimacoes.Count
    
    For i = totalIntimacoes To 1 Step -1
        If VerificarSeLinhaEhCabecalhoDaTabelaDeIntimacoes(tabIntimacoes(i)) = False Then
            dtCont = SisifoEmbasaFuncoes.ConverteDataProjudiParaDate(Trim(tabIntimacoes(i).FindElementsByTag("td")(4).text))
            If dtCont < periodo.DataInicial Then
                VerificarSePeriodoAcabouNaPaginaAtual = True
                Exit For
            End If
        End If
    Next i
    
End Function

Function PegarTabIntimacoes() As Selenium.WebElements
    Set PegarTabIntimacoes = chromeProj.Chrome.Chrome.FindElementsByXPath("//*[@name='formIntimacoes']/table/tbody/tr")
End Function

Sub SinalizarIntimacaoProcessada(numero As Integer)
    chromeProj.Chrome.Chrome.ExecuteScript "sfCel = document.querySelector('body > div > form:nth-child(5) > table > tbody > tr:nth-child(" & numero & ") > td:nth-child(4)');"
    chromeProj.Chrome.Chrome.ExecuteScript "sfCel.removeAttribute('class');"
    chromeProj.Chrome.Chrome.ExecuteScript "sfCel.setAttribute('style','background-color: #a8ffd3;');"
End Sub

Function PegarLinkProcesso(ByVal numeroCnj As String) As String
    PegarLinkProcesso = chromeProj.PegarLinkProcesso(numeroCnj)
End Function

Function AbrirPaginaProcesso(url As String) As String
    AbrirPaginaProcesso = chromeProj.AbrirPaginaProcesso(url)
End Function

Function AbrirProcesso(numeroProcesso As String) As String
    AbrirProcesso = chromeProj.AbrirProcesso(numeroProcesso)
End Function

Sub ExpandirBotoesProcesso(Optional ByVal intQuantidadeAExpandir As Integer)
    chromeProj.ExpandirBotoesProcesso intQuantidadeAExpandir
End Sub

Function PegarLinhasAndamentosDoProcesso() As Selenium.WebElements
    Set PegarLinhasAndamentosDoProcesso = chromeProj.PegarLinhasAndamentosDoProcesso
End Function

Function PegarDivAndamentosDoProcesso() As Selenium.WebElement
    Set PegarDivAndamentosDoProcesso = chromeProj.PegarDivAndamentosDoProcesso
End Function

Public Function PerguntarAndamentoCorreto(ByRef linhasAndamentos As Selenium.WebElements) As Selenium.WebElement
    Dim indexIntimacao As Integer
    Dim LinhaAndamento As Selenium.WebElement, tbodyAndamentos As Selenium.WebElement
    Dim i As Integer
    
    indexIntimacao = ProcurarIndexIntimacaoMaisRecenteParaEmbasaOuSeusAdvs(linhasAndamentos)
    Set LinhaAndamento = DecifrarAndamentoReferenteAIntimacao(indexIntimacao, linhasAndamentos)
    Set tbodyAndamentos = LinhaAndamento.FindElementByXPath("./..")
    InserirCssJavascript linhasAndamentos, LinhaAndamento
    tbodyAndamentos.ScrollIntoView True
    Set LinhaAndamento = AguardarSelecaoDeLinha(tbodyAndamentos, 300000)
    Set PerguntarAndamentoCorreto = LinhaAndamento
End Function

Private Function ProcurarIndexIntimacaoMaisRecenteParaEmbasaOuSeusAdvs(linhasAndamentos As Selenium.WebElements) As Integer
    Dim i As Integer, primeiraLinhaAposCabecalho As Integer
    Dim nomeAndamento As String, obsAndamento As String
    
    primeiraLinhaAposCabecalho = 2
    For i = primeiraLinhaAposCabecalho To linhasAndamentos.Count Step 1
        nomeAndamento = linhasAndamentos(i).FindElementByTag("b").text
        obsAndamento = linhasAndamentos(i).FindElementByTag("tr").FindElementsByTag("td")(2).text
        obsAndamento = Trim(Replace(obsAndamento, nomeAndamento, ""))
        obsAndamento = Trim(Replace(obsAndamento, chr(10), ""))
        
        'Encontrada a intimação expedida para a Embasa, sai deste loop e vai para o próximo, procurar o andamento a que se refere.
        If nomeAndamento = "Intimação expedido(a)" And _
            (InStr(1, LCase(obsAndamento), "embasa") <> 0 Or InStr(1, LCase(obsAndamento), "empresa baiana de aguas e saneamento") <> 0) Then
            Exit For
        ElseIf Left(nomeAndamento, 22) = "Incluído em pauta para" Then
            Exit For
        End If
    Next i
    ProcurarIndexIntimacaoMaisRecenteParaEmbasaOuSeusAdvs = i
End Function

Private Function DecifrarAndamentoReferenteAIntimacao(indexIntimacao As Integer, linhasAndamentos As Selenium.WebElements) As Selenium.WebElement
    Dim i As Integer
    Dim nomeAndamento As String
    For i = indexIntimacao + 1 To linhasAndamentos.Count Step 1
        nomeAndamento = PegarNomeAndamento(linhasAndamentos(i))
        If Left(nomeAndamento, 9) <> "Intimação" _
            And nomeAndamento <> "Cumprimento Genérico expedido(a)" _
            And nomeAndamento <> "Juntada de Cumprimento Genérico" _
            And nomeAndamento <> "Expedição de Intimação" _
            And nomeAndamento <> "Juntada de Intimação" _
            And nomeAndamento <> "Citação lido(a)" _
            And nomeAndamento <> "Mandado devolvido Cumprido com finalidade atingida" _
            And nomeAndamento <> "Recebido o Mandado para Cumprimento" _
            And nomeAndamento <> "" _
            Then Exit For
    Next i
    Set DecifrarAndamentoReferenteAIntimacao = linhasAndamentos(i)
End Function

Public Function PegarNomeAndamento(LinhaAndamento As Selenium.WebElement) As String
    PegarNomeAndamento = chromeProj.PegarNomeAndamento(LinhaAndamento)
End Function

Private Sub InserirCssJavascript(linhasAndamentos As Selenium.WebElements, linhaEscolhida As Selenium.WebElement)
    Dim i As Integer, primeiraLinhaAposCabecalho As Integer, qtdElementosAManipular As Integer
    Dim teste As Selenium.WebElement
    
    AjustarLarguraTabela
    CriarTagStyle
    PrepararCssTransicoesLinhas
    InserirTagStyle
    DefinirVariavelJsTbodyAndamentos
    
    primeiraLinhaAposCabecalho = 2
    qtdElementosAManipular = linhasAndamentos.Count - 1
    For i = primeiraLinhaAposCabecalho - 1 To qtdElementosAManipular Step 1
        If i = 141 Then
            Debug.Print i
        End If
        MudarClasseDeLinhaParaLinhasisifo i
        AdicionarOnClickNaLinha i
    Next i
    
    linhaEscolhida.ExecuteScript "this.setAttribute('class', 'linhasisifoescolhida');"
    
End Sub

Private Sub CriarTagStyle()
    chromeProj.Chrome.Chrome.ExecuteScript "sfCss = document.createElement('style');"
    chromeProj.Chrome.Chrome.ExecuteScript "sfCss.type = 'text/css';"
End Sub

Private Sub PrepararCssTransicoesLinhas()
    chromeProj.Chrome.Chrome.ExecuteScript "sfCss.innerText = '.linhasisifo { opacity:0.6; transition:all 0.3s ease; } '"
    chromeProj.Chrome.Chrome.ExecuteScript "sfCss.innerText += '.linhasisifoescolhida, .linhasisifo:hover { cursor: pointer; opacity:1; box-shadow: 0 0 5px 0 #388aed; -webkit-transform: scale(1.15); transform: scale(1.15); } '"
End Sub

Private Sub InserirTagStyle()
    chromeProj.Chrome.Chrome.ExecuteScript "document.head.appendChild(sfCss)"
End Sub

Private Sub DefinirVariavelJsTbodyAndamentos()
    chromeProj.Chrome.Chrome.ExecuteScript "sfTbodyAndamentos = sfTabAndamentos.getElementsByTagName('tbody')[0]"
End Sub

Private Sub MudarClasseDeLinhaParaLinhasisifo(numeroLinha As Integer)
    chromeProj.Chrome.Chrome.ExecuteScript "sfTbodyAndamentos.children[" & numeroLinha & "].setAttribute('class', 'linhasisifo');"
End Sub

Private Sub AdicionarOnClickNaLinha(numeroLinha As Integer)
    chromeProj.Chrome.Chrome.ExecuteScript "sfTbodyAndamentos.children[" & numeroLinha & "].setAttribute(""onclick"", ""javascript: this.setAttribute('id', 'sfdefinitiva');"");"
End Sub

Private Sub AjustarLarguraTabela()
    chromeProj.Chrome.Chrome.ExecuteScript "sfTabAndamentos = document.getElementById('Arquivos');"
    chromeProj.Chrome.Chrome.ExecuteScript "sfTabAndamentos.children[0].setAttribute('width', '80%');"
End Sub

Private Function AguardarSelecaoDeLinha(tbodyAndamentos As Selenium.WebElement, timeoutMilissegundos As Long) As Selenium.WebElement
    Dim linhaEscolhida As Selenium.WebElement
    Dim por As Selenium.By
    Set por = New Selenium.By
    
    If tbodyAndamentos.IsElementPresent(por.ID("sfdefinitiva"), timeoutMilissegundos) = True Then
        Set linhaEscolhida = tbodyAndamentos.FindElementById("sfdefinitiva")
    End If
    Set AguardarSelecaoDeLinha = linhaEscolhida
End Function

Function PegarNumeroProcesso() As String
    PegarNumeroProcesso = chromeProj.PegarNumeroProcesso
End Function

Function PegarCookies() As Selenium.Cookies
    Set PegarCookies = chromeProj.PegarCookies
End Function

Function PegarJuizo() As String
    PegarJuizo = chromeProj.PegarJuizo
End Function

Public Function PegarDataDaLinhaIntimacao(linhasDeIntimacoes As Selenium.WebElements, numeroLinha As Integer) As String
    Dim textoData As String
    textoData = linhasDeIntimacoes(numeroLinha).FindElementsByTag("td")(4).text
    PegarDataDaLinhaIntimacao = Trim(textoData)
End Function

Public Function PegarIntimacaoDaLinha(linhasDeIntimacoes As Selenium.WebElements, numeroLinha As Integer, dataExpedicao As Date) As Intimacao
    Dim Intimacao As New Intimacao
    
    With Intimacao
        .numeroProcesso = Trim(linhasDeIntimacoes(numeroLinha).FindElementsByTag("td")(2).text)
        .Destinatario = Trim(linhasDeIntimacoes(numeroLinha).FindElementsByTag("td")(3).text)
        If InStr(1, LCase(.Destinatario), "embasa") = 0 And _
            InStr(1, LCase(.Destinatario), "saneamento") = 0 Then
            .SuspeitaDeNaoSerEmbasa = True
        End If
        .dataExpedicao = dataExpedicao
    End With
    
    Set PegarIntimacaoDaLinha = Intimacao
End Function
