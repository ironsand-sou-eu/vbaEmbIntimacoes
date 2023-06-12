VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeProjudi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim oChrome As Chrome
Dim strPerfilLogado As String

'Propriedade PerfilLogado:
Public Property Get PerfilLogado() As String
PerfilLogado = strPerfilLogado
End Property
Public Property Let PerfilLogado(str As String)
strPerfilLogado = str
End Property

'Propriedade Chrome:
Public Property Get Chrome() As Chrome
Set Chrome = oChrome
End Property
Public Property Set Chrome(ch As Chrome)
Set oChrome = ch
End Property

Private Sub Class_Initialize()
    Set oChrome = New Chrome
End Sub

Public Function AbrirProcesso(numeroProcesso As String) As String
    Dim Autenticador As New CasoUsoAutenticador
    Dim usuario As New UsuarioProjudiPje
    Dim resposta As String, urlProcesso As String
    
    Set usuario = Autenticador.PegarObjetoUsuarioAtual(projudi)
    Set Autenticador.Controlador = New ChromeProjudiAutenticacao
    Set oChrome.Chrome = Autenticador.CarregarSessaoValida(usuario)
    strPerfilLogado = Autenticador.PerfilLogado
    If oChrome.Chrome Is Nothing Then
        resposta = "não foi possível logar no sistema"
    Else
        AbrirPaginaDeBuscaProcessual strPerfilLogado
        urlProcesso = PegarLinkProcesso(numeroProcesso)
        If urlProcesso = "" Then
            resposta = "Processo não encontrado"
            Autenticador.LimparCookieSessao
            Exit Function
        End If
        resposta = AbrirPaginaProcesso(urlProcesso)
        If resposta = "sucesso" Then ExpandirBotoesProcesso 8
    End If
    AbrirProcesso = resposta
End Function

Public Sub AguardarCarregarHome()
    Dim frame As Selenium.WebElement
    Dim htmlAnchors As Selenium.WebElements
    
    On Error Resume Next
    Select Case LCase(strPerfilLogado)
    Case "advogado"
        
        
    Case "representante", ""
    
        oChrome.Chrome.SwitchToParentFrame
        oChrome.Chrome.SwitchToParentFrame
        
        Do
            Set frame = oChrome.Chrome.FindElementByName("mainFrame")
        Loop While frame Is Nothing
        
        oChrome.Chrome.SwitchToFrame 1
        Do
            Set frame = oChrome.Chrome.FindElementByName("userMainFrame")
        Loop While frame Is Nothing
        
        oChrome.Chrome.SwitchToFrame 0
        
        Do
            Set htmlAnchors = oChrome.Chrome.FindElementsByTag("a")
        Loop While htmlAnchors.Count = 0
        
    End Select
    On Error GoTo 0
End Sub

Public Sub AbrirPaginaDeBuscaProcessual(PerfilLogado As String)
    Dim por As New Selenium.By
    Dim jaEstaNaPagina As Boolean
    Dim paginaDeBusca As String
    
    strPerfilLogado = PerfilLogado
    Select Case LCase(strPerfilLogado)
    Case "advogado"
        jaEstaNaPagina = (oChrome.Chrome.url = SisifoEmbasaFuncoes.sfUrlProjudiBuscaAdvogado1g)
        paginaDeBusca = SisifoEmbasaFuncoes.sfUrlProjudiBuscaAdvogado1g
    Case "representante"
        jaEstaNaPagina = (oChrome.Chrome.url = SisifoEmbasaFuncoes.sfUrlProjudiBuscaParte1g)
        paginaDeBusca = SisifoEmbasaFuncoes.sfUrlProjudiBuscaParte1g
    End Select
    
    If jaEstaNaPagina = False Or _
    (jaEstaNaPagina = True And oChrome.Chrome.IsElementPresent(por.Name("busca")) = False) Then
        oChrome.Chrome.Get paginaDeBusca
    End If
    
End Sub

Public Function LidarComCaptcha() As String
    Dim token As Selenium.WebElement
    Dim textoAlerta As String
    
    On Error GoTo Erro
inicio:
    Do
        'Procurar token do recaptcha e aguardar ter o atributo. Enquanto não tiver, loop
        If oChrome.Chrome.FindElementsById("token").Count = 0 Then
            Set token = Nothing
            Exit Do
        Else
            Set token = oChrome.Chrome.FindElementsById("token")(1)
        End If
    Loop While token.Attribute("value") = ""
    On Error GoTo 0
    
    LidarComCaptcha = "sucesso"
    Exit Function
    
Erro:
    textoAlerta = PegarTextoAlerta
    If textoAlerta = "sem alerta" Then
        GoTo inicio ' Às vezes está vindo para o tratamento de erro sem código de erro (só encontrei a possibilidade de ser ao atribuir o objeto ("token")(1))
    Else
        LidarComCaptcha = textoAlerta
    End If
    
End Function

Private Function PegarTextoAlerta() As String
    
    Dim alerta As Selenium.alert
    
    On Error Resume Next
    Set alerta = oChrome.Chrome.SwitchToAlert(500)
    On Error GoTo 0
    
    If alerta Is Nothing Then
        PegarTextoAlerta = "sem alerta"
    Else
        PegarTextoAlerta = alerta.text
        alerta.accept
    End If
    
End Function

Function RealizarBusca(numeroCnj As String) As String
    Dim textboxNumeroProcesso As Selenium.WebElement
    Dim resultadoCaptcha As String
    
    On Error Resume Next
    Set textboxNumeroProcesso = oChrome.Chrome.FindElementById("numeroProcesso")
    On Error GoTo 0
    
    If Not textboxNumeroProcesso Is Nothing Then
        textboxNumeroProcesso.SendKeys numeroCnj
        oChrome.Chrome.FindElementByName("busca").submit
        resultadoCaptcha = LidarComCaptcha
        AguardarCarregarBusca
        RealizarBusca = "sucesso"
    Else
        RealizarBusca = "Textbox de número do processo não encontrada"
    End If
End Function

Sub AguardarCarregarBusca()
    Dim trsRespostaBusca As Selenium.WebElements, formsBusca As Selenium.WebElements
    Dim respostaBuscaContemItens As Boolean, retornouParaFormBusca As Boolean
    Dim nomeTabelaBusca As String
    
    Select Case LCase(strPerfilLogado)
    Case "advogado"
        nomeTabelaBusca = "form1"
    Case "representante", ""
        nomeTabelaBusca = "formProcessos"
    End Select
    
    Do
        Set trsRespostaBusca = oChrome.Chrome.FindElementsByXPath("//*[@name='" & nomeTabelaBusca & "']/table/tbody/tr")
        respostaBuscaContemItens = (trsRespostaBusca.Count >= 3)
        Set formsBusca = oChrome.Chrome.FindElementsByName("busca")
        retornouParaFormBusca = (formsBusca.Count > 0)
    Loop Until respostaBuscaContemItens Or retornouParaFormBusca
    
End Sub

Function PegarLinkProcesso(ByVal numeroCnj As String) As String
    Dim strContNumeroProcesso As String
    Dim formResultadosBusca As Selenium.WebElement, tabelaResultadosBusca As Selenium.WebElement
    Dim i As Variant
    
    Do
        AbrirPaginaDeBuscaProcessual strPerfilLogado
        i = RealizarBusca(numeroCnj)
    Loop Until i = "sucesso"
    
    Set formResultadosBusca = PegarFormResultadosBusca
    
    If formResultadosBusca Is Nothing Then
        PegarLinkProcesso = ""
    Else
        Set tabelaResultadosBusca = formResultadosBusca.FindElementByTag("tbody")
        For Each i In tabelaResultadosBusca.FindElementsByTag("a")
            If i.text = numeroCnj Then
                PegarLinkProcesso = i.Attribute("href")
                Exit For
            End If
        Next i
    End If
    
End Function

Function PegarFormResultadosBusca() As Selenium.WebElement
    Dim formsResultadosBusca As Selenium.WebElements
    Dim nomeFormResultadosBusca As String
    
    If LCase(strPerfilLogado) = "advogado" Then
        nomeFormResultadosBusca = "form1"
    Else
        nomeFormResultadosBusca = "formProcessos"
    End If
    
    Set formsResultadosBusca = oChrome.Chrome.FindElementsByName(nomeFormResultadosBusca, 0)
    
    If formsResultadosBusca.Count > 0 Then
        Set PegarFormResultadosBusca = formsResultadosBusca(1)
    Else
        Set PegarFormResultadosBusca = Nothing
    End If
End Function

Function AbrirPaginaProcesso(url As String) As String
    Dim resposta As String
    oChrome.Chrome.Get url
    
    'If strPerfilLogado = "Advogado" Then
        'AceitarAlerta 5000
        'If MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", rogo que resolva o recapctha que apareceu na janela " & _
        "do navegador e clique em Ok para continuar, ou em Cancelar caso queira parar minha execução.", _
        vbOKCancel + vbExclamation, "Sísifo - Recaptcha do Projudi") = vbCancel Then
        '    AbrirPaginaProcesso = "Execução cancelada pelo usuário"
        '    Exit Function
        'End If
    'End If
    
    resposta = VerificarSegredoDeJustica
    
    If resposta = "sem segredo" Then
        AbrirPaginaProcesso = "sucesso"
    Else
        AbrirPaginaProcesso = resposta
    End If
End Function

Private Sub AceitarAlerta(timeoutEmMs As Long)
    Dim alerta As Selenium.alert
    
    On Error Resume Next
    Set alerta = oChrome.Chrome.SwitchToAlert(timeoutEmMs)
    On Error GoTo 0
    
    If Not alerta Is Nothing Then alerta.accept
    
End Sub

Function VerificarSegredoDeJustica() As String
    Dim mensagem As String
    
    mensagem = oChrome.Chrome.FindElementByTag("p").text
    If mensagem = "Processo sob Segredo de Justiça" Then
        VerificarSegredoDeJustica = "processo em segredo de justiça"
    ElseIf InStr(1, mensagem, "Erro 403") <> 0 Then
        VerificarSegredoDeJustica = "sem acesso ao processo"
    Else
        VerificarSegredoDeJustica = "sem segredo"
    End If
End Function

Sub ExpandirBotoesProcesso(Optional ByVal intQuantidadeAExpandir As Integer)
    Dim imagem As Selenium.WebElement
    Dim i As Integer, iAbertos As Integer
    
    For i = 1 To oChrome.Chrome.FindElementsByTag("img").Count
        Set imagem = oChrome.Chrome.FindElementsByTag("img")(i)
        If InStr(1, imagem.Attribute("src"), "/observacao.png") <> 0 Or InStr(1, imagem.Attribute("src"), "/arquivos.png") <> 0 Then
            imagem.FindElementByXPath("./..").Click
            If Not IsMissing(intQuantidadeAExpandir) Then
                iAbertos = iAbertos + 1
                If iAbertos >= intQuantidadeAExpandir Then Exit For
            End If
        End If
    Next i
End Sub

Function PegarLinhasAndamentosDoProcesso() As Selenium.WebElements
    Set PegarLinhasAndamentosDoProcesso = oChrome.Chrome.FindElementsByXPath("//*[@id=""Arquivos""]/table/tbody/tr")
End Function

Function PegarDivAndamentosDoProcesso() As Selenium.WebElement
    Set PegarDivAndamentosDoProcesso = oChrome.Chrome.FindElementById("Arquivos")
End Function

Function PegarNomeAndamento(LinhaAndamento As Selenium.WebElement) As String
    Dim tagB As Selenium.WebElement
    
    On Error Resume Next
    Set tagB = LinhaAndamento.FindElementByTag("b", 0)
    On Error GoTo 0
    
    If Not tagB Is Nothing Then
        PegarNomeAndamento = tagB.text
    Else
        PegarNomeAndamento = ""
    End If
End Function

Function PegarNumeroProcesso() As String
    PegarNumeroProcesso = Trim(oChrome.Chrome.FindElementsByTag("a")(1).text)
End Function

Function PegarCookies() As Selenium.Cookies
    Set PegarCookies = oChrome.Chrome.Manage.Cookies
End Function

Function PegarJuizo() As String
    Dim juizo As String
    Dim posicaoFinal As Integer
    
    juizo = oChrome.Chrome.FindElementByXPath("//*[@id=""Partes""]/table/tbody/tr[7]/td[2]").text
    If InStr(1, juizo, "Juiz:") <> 0 Then
        posicaoFinal = InStr(1, juizo, "Juiz:") - 2
    Else
        posicaoFinal = InStr(1, juizo, "Histórico de Juízes") - 2
    End If
    PegarJuizo = Left(juizo, posicaoFinal)
End Function



