VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromePjeTjba1gIntimacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ChromePje As ChromePjetjba1g
Dim usrUsuario As UsuarioProjudiPje

Private Sub Class_Initialize()
    Set ChromePje = New ChromePjetjba1g
End Sub

'Propriedade Usuario:
Public Property Get usuario() As UsuarioProjudiPje
usuario = usrUsuario
End Property
Public Property Set usuario(usr As UsuarioProjudiPje)
Set usrUsuario = usr
End Property

'Propriedade ChromePje:
Public Property Get ChromePje1g() As ChromePjetjba1g
Set ChromePje1g = ChromePje
End Property
Public Property Set ChromePje1g(chPje As ChromePjetjba1g)
Set ChromePje = chPje
End Property

Public Function PegarDivAndamentos() As Selenium.WebElement
    Dim resposta As Selenium.WebElement
    Dim i As Integer
    
    On Error Resume Next
    With ChromePje.Chrome.Chrome
        For i = 1 To .Windows.Count Step 1
            If IsNumeric(Left(.Windows(i).Title, 7)) Then
                Exit For
            Else
                .SwitchToNextWindow 0
            End If
        Next i
        Set resposta = .FindElementById("divTimeLine:eventosTimeLineElement")
    End With
    On Error GoTo 0
    
    Set PegarDivAndamentos = resposta
End Function

Public Sub AbrirAbaExpedientes()
    Dim divIntimacoes As Selenium.WebElement
    
    ChromePje.AbrirPaginaPainelUsuario
    If AbaExpedientesEstaAtiva = False Then
        Do
            ChromePje.Chrome.Chrome.FindElementById("tabExpedientes_cell").Click
            Set divIntimacoes = ChromePje.Chrome.Chrome.FindElementById("abaExpedientes", Raise:=False)
        Loop While divIntimacoes.IsDisplayed = False
    End If
End Sub

Public Function PerguntarAndamentoCorreto(ByRef divAndamentos As Selenium.WebElement) As Selenium.WebElement
    Dim linhasAndamentos As Selenium.WebElements
    Dim LinhaAndamento As Selenium.WebElement, elementoA As Selenium.WebElement
    Dim i As Integer, indexIntimacao As Integer
    
    indexIntimacao = 1
    Set linhasAndamentos = divAndamentos.FindElementsByXPath("./div")
    Set LinhaAndamento = DecifrarAndamentoReferenteAIntimacao(indexIntimacao, linhasAndamentos)
    On Error Resume Next
    Set elementoA = LinhaAndamento.FindElementByTag("a")
    On Error GoTo 0
    If Not elementoA Is Nothing Then elementoA.Click
    InserirCssJavascript linhasAndamentos, LinhaAndamento
    LinhaAndamento.ScrollIntoView True
    Set LinhaAndamento = AguardarSelecaoDeLinha(divAndamentos, 300000)
    Set PerguntarAndamentoCorreto = LinhaAndamento
End Function

Private Function DecifrarAndamentoReferenteAIntimacao(indexIntimacao As Integer, linhasAndamentos As Selenium.WebElements) As Selenium.WebElement
    Dim i As Integer
    Dim nomeAndamento As String
    For i = indexIntimacao To linhasAndamentos.Count Step 1
        If InStr(1, linhasAndamentos(i).Attribute("class"), "data") = 0 Then
            nomeAndamento = LCase(PegarNomeAndamento(linhasAndamentos(i)))
            If InStr(1, nomeAndamento, "expedição de") = 0 _
                And InStr(1, nomeAndamento, "dj eletrônico") = 0 _
                And InStr(1, nomeAndamento, "publicado intimação em") = 0 _
                And InStr(1, nomeAndamento, "parecer do") = 0 _
                And InStr(1, nomeAndamento, "juntada de petição") = 0 _
                Then Exit For
        End If
    Next i
    Set DecifrarAndamentoReferenteAIntimacao = linhasAndamentos(i)
End Function

Public Function PegarNomeAndamento(LinhaAndamento As Selenium.WebElement) As String
    Dim divMediaBodyBox As Selenium.WebElement, elementoResposta As Selenium.WebElement
    Dim resposta As String
    
    Set divMediaBodyBox = LinhaAndamento.FindElementByClass("box")
    
    On Error Resume Next
    Set elementoResposta = divMediaBodyBox.FindElementByClass("text-upper")
    If elementoResposta Is Nothing Then
        Set elementoResposta = divMediaBodyBox.FindElementByClass("anexos").FindElementByTag("a")
    End If
    On Error GoTo 0
    
    PegarNomeAndamento = elementoResposta.text
End Function

Private Sub InserirCssJavascript(linhasAndamentos As Selenium.WebElements, linhaEscolhida As Selenium.WebElement)
    Dim i As Integer, primeiraLinhaAposCabecalho As Integer, qtdElementosAManipular As Integer
    Dim teste As Selenium.WebElement
    
    CriarTagStyle
    PrepararCssTransicoesLinhas
    InserirTagStyle
    DefinirVariavelJsDivAndamentos
    
    primeiraLinhaAposCabecalho = 2
    qtdElementosAManipular = linhasAndamentos.Count - 1
    For i = primeiraLinhaAposCabecalho - 1 To qtdElementosAManipular Step 1
        If InStr(1, linhasAndamentos(i + 1).Attribute("class"), "data") = 0 Then
            MudarClasseDeLinhaParaLinhasisifo i
            AdicionarOnClickNaLinha i
        End If
    Next i
    
    ChromePje.Chrome.Chrome.Window.Maximize
    linhaEscolhida.ExecuteScript "this.children[1].children[0].setAttribute('class', 'linhasisifoescolhida');"
    
End Sub

Private Sub CriarTagStyle()
    ChromePje.Chrome.Chrome.ExecuteScript "sfCss = document.createElement('style');"
    ChromePje.Chrome.Chrome.ExecuteScript "sfCss.type = 'text/css';"
End Sub

Private Sub PrepararCssTransicoesLinhas()
    ChromePje.Chrome.Chrome.ExecuteScript "sfCss.innerText = '.linhasisifo { opacity:0.6; transition:all 0.3s ease; padding: 10px 15px 2px; line-height: 25px; margin-bottom: 5px; word-wrap: break-word; width: calc(100% - 36px); } '"
    ChromePje.Chrome.Chrome.ExecuteScript "sfCss.innerText += '.linhasisifoescolhida, .linhasisifo:hover { cursor: pointer; opacity: 1; box-shadow: 0 0 5px 0 #008000; background-color: #008000; font-weight: bold; color: white; } '"
End Sub

Private Sub InserirTagStyle()
    ChromePje.Chrome.Chrome.ExecuteScript "document.head.appendChild(sfCss)"
End Sub

Private Sub DefinirVariavelJsDivAndamentos()
    ChromePje.Chrome.Chrome.ExecuteScript "sfDivAndamentos = document.getElementById('divTimeLine:eventosTimeLineElement')"
End Sub

Private Sub MudarClasseDeLinhaParaLinhasisifo(numeroLinha As Integer)
    ChromePje.Chrome.Chrome.ExecuteScript "sfDivAndamentos.children[" & numeroLinha & "].children[1].children[0].setAttribute('class', 'linhasisifo');"
End Sub

Private Sub AdicionarOnClickNaLinha(numeroLinha As Integer)
    ChromePje.Chrome.Chrome.ExecuteScript "sfDivAndamentos.children[" & numeroLinha & "].children[1].children[0].setAttribute(""onclick"", ""javascript: this.setAttribute('id', 'sfdefinitiva');"");"
End Sub

Private Function AguardarSelecaoDeLinha(divAndamentos As Selenium.WebElement, timeoutMilissegundos As Long) As Selenium.WebElement
    Dim linhaEscolhida As Selenium.WebElement
    Dim por As New Selenium.By
    
    If divAndamentos.IsElementPresent(por.ID("sfdefinitiva"), timeoutMilissegundos) = True Then
        Set linhaEscolhida = divAndamentos.FindElementByXPath("//*[@id='sfdefinitiva']/../..")
    End If
    Set AguardarSelecaoDeLinha = linhaEscolhida
End Function

Public Function ColetarIntimacoes(nomeCaixaAColetar As String, nomeComarca As String) As Collection
    Dim i As Integer
    Dim idAnchorCaixa As String, idTableComarca As String, resultado As String
    Dim colIntimacoes As New Collection, colCont As New Collection
    
    idAnchorCaixa = PegarIdAnchorCaixa(nomeCaixaAColetar)
    AbrirCaixa idAnchorCaixa
    idTableComarca = ChromePje.PegarIdTableComarca(nomeComarca, idAnchorCaixa)
    AbrirComarcaECaixaEntrada idTableComarca, nomeComarca
    AssegurarExistenciaDaCaixaColetadas CStr(idTableComarca)
    Do
        Set colCont = PegarIntimacoesCaixaDeEntrada
        If colCont.Count > 0 Then
            MarcarIntimacoesCaixaDeEntrada colCont.Count
            MoverIntimacoesMarcadasParaCaixaColetadas
            Set colIntimacoes = SisifoEmbasaFuncoes.SomarColecoes(colIntimacoes, colCont)
        End If
    Loop Until CaixaAbertaEstaVazia = True
    
    Set ColetarIntimacoes = colIntimacoes
    
End Function

Private Function AbaExpedientesEstaAtiva() As Boolean
    Dim classesExpedientes As String
    
    classesExpedientes = ChromePje.Chrome.Chrome.FindElementById("tabExpedientes_cell").Attribute("class")
    If InStr(1, classesExpedientes, "rich-tabhdr-cell-active") <> 0 Then
        AbaExpedientesEstaAtiva = True
    Else
        AbaExpedientesEstaAtiva = False
    End If
End Function

Public Function PegarIdAnchorCaixa(nomeCaixa As String) As String
    Dim divMenu As Selenium.WebElement
    Dim i As Selenium.WebElement
    
    Set divMenu = ChromePje.Chrome.Chrome.FindElementById("formAbaExpediente:divMenuContexto")
    For Each i In divMenu.FindElementsByTag("a")
        If InStr(1, i.text, nomeCaixa) <> 0 Then
            PegarIdAnchorCaixa = i.Attribute("id")
            Exit For
        End If
    Next i
End Function

Public Sub AbrirCaixa(idAnchorCaixa As String)
    Dim por As New Selenium.By
    Dim idDivComarcas As String
    
    idDivComarcas = DerivarIdDivComarcasDoIdAnchor(idAnchorCaixa)
    If ChromePje1g.Chrome.Chrome.IsElementPresent(por.ID(idDivComarcas)) = False Then
        ChromePje1g.Chrome.Chrome.FindElementById(idAnchorCaixa).Click
        ChromePje.Chrome.EsperarElementoAparecer DerivarIdDivComarcasDoIdAnchor(idAnchorCaixa), "id"
    End If
End Sub

Public Function ColetarIdsDasTablesComarcas(idAnchorCaixa As String) As Collection
    Dim divComarcas As Selenium.WebElement, i As Selenium.WebElement
    Dim idDivComarcas As String, jId As String, idTableComarca As String
    Dim colIds As Collection
    
    Set colIds = New Collection
    idDivComarcas = DerivarIdDivComarcasDoIdAnchor(idAnchorCaixa)
    Set divComarcas = ChromePje.Chrome.Chrome.FindElementById(idDivComarcas)
    jId = PegarJidDaDiv(divComarcas)
    For Each i In divComarcas.FindElementsByTag("table")
        idTableComarca = i.Attribute("id")
        If InStr(1, idTableComarca, "::" & jId) <> 0 Then
            colIds.Add idTableComarca
        End If
    Next i
    Set ColetarIdsDasTablesComarcas = colIds
End Function

Private Function DerivarIdDivComarcasDoIdAnchor(idAnchor As String) As String
    Dim qtdLetrasManter As Integer
    Dim idBase As String
    
    qtdLetrasManter = InStr(1, idAnchor, ":j_id")
    idBase = Left(idAnchor, qtdLetrasManter)
    DerivarIdDivComarcasDoIdAnchor = idBase & "trPend:childs"
End Function

Private Function PegarJidDaDiv(divComTablesDentro As Selenium.WebElement) As String
    Dim idPrimeiraTable As String
    Dim numCaractereJid As Integer
    
    idPrimeiraTable = divComTablesDentro.FindElementByTag("table").Attribute("id")
    numCaractereJid = InStr(1, idPrimeiraTable, "::j_id")
    PegarJidDaDiv = Mid(idPrimeiraTable, InStr(1, idPrimeiraTable, "::j_id") + 2)
End Function

Public Function PegarNomeComarcaPeloIdDaTabela(idTableComarca As String) As String
    PegarNomeComarcaPeloIdDaTabela = ChromePje.Chrome.Chrome.FindElementById(idTableComarca).FindElementByXPath("./tbody/tr/td[3]/a/span[1]", 0).text
End Function

Private Sub AbrirComarcaECaixaEntrada(idTableComarca As String, nomeComarca As String)
    Dim idDivComarca As String, migalhasDePao As String
    Dim divComarca As Selenium.WebElement
    Dim tableCaixaEntrada As Selenium.WebElement
    
    idDivComarca = idTableComarca & ":childs"
    ChromePje.Chrome.Chrome.FindElementById(idTableComarca).Click
    ChromePje.Chrome.EsperarElementoAparecer idDivComarca, "id"
    migalhasDePao = ChromePje.Chrome.Chrome.FindElementById("formExpedientes:Filtros").text
    If InStr(1, migalhasDePao, "Caixa de entrada") = 0 Then
        Set divComarca = ChromePje.Chrome.Chrome.FindElementById(idDivComarca)
        Set tableCaixaEntrada = PegarTableCaixaEntradaDaDiv(divComarca)
        tableCaixaEntrada.Click
    End If
End Sub

Private Function PegarTableCaixaEntradaDaDiv(div As Selenium.WebElement) As Selenium.WebElement
    Dim i As Selenium.WebElement
    
    For Each i In div.FindElementsByTag("table")
        If InStr(1, i.Attribute("id"), ":-1:") <> 0 Then
            Set PegarTableCaixaEntradaDaDiv = i
            Exit For
        End If
    Next i
End Function

Private Sub AssegurarExistenciaDaCaixaColetadas(idTableComarca As String)
inicio:
    If CaixaColetadasExisteNaTabelaComarca(idTableComarca) = False Then
        CriarCaixaColetadas
        GoTo inicio
    End If
End Sub

Private Function CaixaColetadasExisteNaTabelaComarca(idTableComarca As String) As Boolean
    Dim idDivCaixasComarca As String, jIdTables As String, idTableCaixa As String
    Dim divCaixasComarca As Selenium.WebElement, i As Selenium.WebElement
    
    idDivCaixasComarca = idTableComarca & ":childs"
    Set divCaixasComarca = ChromePje.Chrome.Chrome.FindElementById(idDivCaixasComarca)
    jIdTables = PegarJidDaDiv(divCaixasComarca)
    
    'Dentro da div das caixas de comarca, há um table e um div, para cada linha de caixa dentro da comarca. A table contém a linha título.
    For Each i In divCaixasComarca.FindElementsByTag("table")
        idTableCaixa = i.Attribute("id")
        If InStr(1, idTableCaixa, "::" & jIdTables) <> 0 Then
            If PegarNomeDaCaixaAPartirDaTable(idTableCaixa) = "Coletadas" Then
                CaixaColetadasExisteNaTabelaComarca = True
                Exit Function
            End If
        End If
    Next i
    CaixaColetadasExisteNaTabelaComarca = False
End Function

Private Function PegarNomeDaCaixaAPartirDaTable(idTableCaixa As String) As String
    Dim idTdNomeCaixa As String
    
    idTdNomeCaixa = idTableCaixa & ":text"
    PegarNomeDaCaixaAPartirDaTable = ChromePje.Chrome.Chrome.FindElementById(idTdNomeCaixa).FindElementByClass("nomeTarefa").text
End Function

Private Sub CriarCaixaColetadas()
    Dim txtNomeCaixa As Selenium.WebElement
    
    ChromePje.Chrome.Chrome.ExecuteScript "showModal('#modalNovaCaixa')"
    ChromePje.Chrome.EsperarElementoAparecer "frmNovaCaixa:nmCx", "id"
    Set txtNomeCaixa = ChromePje.Chrome.Chrome.FindElementById("frmNovaCaixa:nmCx")
    SisifoEmbasaFuncoes.PreencherTextboxChromedriver txtNomeCaixa, "Coletadas"
    ChromePje.Chrome.Chrome.FindElementById("frmNovaCaixa:btNCx").Click
    EsperarAmpulheta
    ChromePje.FecharModal "modalNovaCaixa"
End Sub

Private Function PegarIntimacoesCaixaDeEntrada() As Collection
    Dim colIntimacoes As Collection
    Dim tbodyExpedientes As Selenium.WebElement
    Dim intimacaoAtual As Intimacao
    Dim i As Integer
    
    Set colIntimacoes = New Collection
    Set tbodyExpedientes = PegarTbodyExpedientes
    If Not tbodyExpedientes Is Nothing Then
        For i = 1 To tbodyExpedientes.FindElementsByTag("tr").Count Step 1
            Set intimacaoAtual = PegarIntimacaoNaLinha(i)
            colIntimacoes.Add intimacaoAtual
            SinalizarIntimacaoColetada i
        Next i
    End If
    Set PegarIntimacoesCaixaDeEntrada = colIntimacoes
End Function

Private Function CaixaAbertaEstaVazia() As Boolean
    Dim tbodyExpedientes As Selenium.WebElement
    
    Set tbodyExpedientes = PegarTbodyExpedientes
    If tbodyExpedientes Is Nothing Then CaixaAbertaEstaVazia = True
End Function

Private Function PegarTbodyExpedientes() As Selenium.WebElement
    Dim divExpedientes As Selenium.WebElement
    
    Set divExpedientes = PegarDivExpedientes
    If InStr(1, divExpedientes.text, "Não há itens aqui") <> 0 Then
        Set PegarTbodyExpedientes = Nothing
    Else
        Set PegarTbodyExpedientes = divExpedientes.FindElementByXPath("./table/tbody")
    End If
End Function

Private Function PegarDivExpedientes() As Selenium.WebElement
    Set PegarDivExpedientes = ChromePje.Chrome.Chrome.FindElementById("formExpedientes:conteudoPesquisaExpedientes")
End Function

Private Function PegarIntimacaoNaLinha(numeroTr As Integer) As Intimacao
    Dim tbodyExpedientes As Selenium.WebElement
    Dim textoTotal As String, ModoExpedicao As String, numProcesso As String, dataExpedicao As String
    Dim minhaInt As Intimacao
    
    Set tbodyExpedientes = PegarTbodyExpedientes
    textoTotal = tbodyExpedientes.FindElementByXPath("./tr[" & numeroTr & "]/td[2]/div/div[1]/div[3]").text
    dataExpedicao = tbodyExpedientes.FindElementByXPath("./tr[" & numeroTr & "]/td[2]/div/div[1]/div[3]/span/span").text
    ModoExpedicao = Replace(textoTotal, dataExpedicao, "")
    ModoExpedicao = Trim(Replace(ModoExpedicao, "()", ""))
    
    textoTotal = tbodyExpedientes.FindElementByXPath("./tr[" & numeroTr & "]/td[2]/div/div[2]/div/div[1]/a").text
    numProcesso = PegarPrimeiroNumeroProcessoEmStringViaRegex(textoTotal)
    
    Set minhaInt = New Intimacao
    minhaInt.numeroProcesso = numProcesso
    minhaInt.ModoExpedicao = ModoExpedicao
    minhaInt.dataExpedicao = dataExpedicao
    
    Set PegarIntimacaoNaLinha = minhaInt
End Function

Private Sub SinalizarIntimacaoColetada(numeroTr As Integer)
    Dim tbodyExpedientes As Selenium.WebElement, spanDataExpedicao As Selenium.WebElement
    
    Set tbodyExpedientes = PegarTbodyExpedientes
    Set spanDataExpedicao = tbodyExpedientes.FindElementByXPath("./tr[" & numeroTr & "]/td[2]/div/div[1]", 0)
    spanDataExpedicao.ExecuteScript "this.setAttribute('style','background-color: #a8ffd3;')"
End Sub

Private Function PegarPrimeiroNumeroProcessoEmStringViaRegex(texto As String) As String
    Dim regEx As RegExp
    Dim resultadoRegEx As Object
    
    Set regEx = New RegExp
    With regEx
        .Pattern = "\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}"
        .IgnoreCase = True
        Set resultadoRegEx = .Execute(texto)
        PegarPrimeiroNumeroProcessoEmStringViaRegex = resultadoRegEx(0)
    End With
End Function

Private Sub MarcarIntimacoesCaixaDeEntrada(qtdIntimacoes As Integer)
    If qtdIntimacoes = 1 Then
        MarcarUnicaIntimacaoCaixaDeEntrada
    Else
        MarcarTodasIntimacoesPaginaInicialCaixaDeEntrada
    End If
End Sub

Private Sub MarcarUnicaIntimacaoCaixaDeEntrada()
    Dim checkbox As Selenium.WebElement
    Dim ampulhetaVisivel As Boolean
    
    ChromePje.Chrome.EsperarElementoAparecer "./table/thead/tr/th[1]/div/div/span/input", "xpath"
    Set checkbox = PegarDivExpedientes.FindElementByXPath("./table/thead/tr/th[1]/div/div/span/input", Raise:=False)
    checkbox.Click
    EsperarAmpulheta
End Sub

Private Sub MarcarTodasIntimacoesPaginaInicialCaixaDeEntrada()
    Dim checkbox As Selenium.WebElement
    Dim ampulhetaVisivel As Boolean
    
    ChromePje.Chrome.EsperarElementoAparecer "./table/thead/tr/th[1]/div/div/span/input", "xpath"
    Set checkbox = PegarDivExpedientes.FindElementByXPath("./table/thead/tr/th[1]/div/div/span/input", Raise:=False)
    checkbox.Click
    EsperarAmpulheta
End Sub

Private Sub MoverIntimacoesMarcadasParaCaixaColetadas()
    Dim selectCaixa As Selenium.WebElement, chkProcessosRelacionados As Selenium.WebElement, modalResultado As Selenium.WebElement
    
    ChromePje.Chrome.Chrome.ExecuteScript "showModal('#modalMoverPara')"
    
    ChromePje.Chrome.EsperarElementoAparecer "frmMoverPara:cxDestino", "id"
    Set selectCaixa = ChromePje.Chrome.Chrome.FindElementById("frmMoverPara:cxDestino")
    Do
        selectCaixa.AsSelect.SelectByText "Coletadas"
    Loop Until selectCaixa.AsSelect.SelectedOption.text = "Coletadas"
    
    ChromePje.Chrome.EsperarElementoAparecer "frmMoverPara:forcarMoverRelacionados", "id"
    Set chkProcessosRelacionados = ChromePje.Chrome.Chrome.FindElementById("frmMoverPara:forcarMoverRelacionados")
    chkProcessosRelacionados.Click
    
    ChromePje.Chrome.Chrome.FindElementById("frmMoverPara:btMvPr").Click
    
    ChromePje.Chrome.EsperarElementoAparecer "modalMessage", "id"
    ChromePje.FecharModal "modalMessage"
    ChromePje.Chrome.EsperarElementoSumir "modalMessage", "id"
End Sub

Private Sub EsperarAmpulheta()
    ChromePje.Chrome.EsperarElementoAparecer "_viewRoot:status.start", "id"
    ChromePje.Chrome.EsperarElementoSumir "_viewRoot:status.start", "id"
End Sub
