VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromePjetjba1g"
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
    Dim linhaProcesso As Selenium.WebElement
    
    Set usuario = Autenticador.PegarUsuarioPeloNome("Pje1g Representante Gestor Embasa", pje1g)
    Set Autenticador.Controlador = New ChromePjeTjba1gAutenticacao
    Set oChrome.Chrome = Autenticador.CarregarSessaoValida(usuario)
    strPerfilLogado = Autenticador.PerfilLogado
    If oChrome.Chrome Is Nothing Then
        resposta = "não foi possível logar no sistema"
    Else
        AbrirPaginaDeBuscaProcessual strPerfilLogado
        EfetuarPesquisa numeroProcesso
        Set linhaProcesso = PegarLinhaProcesso
    
        If linhaProcesso Is Nothing Then
            resposta = "Processo não encontrado"
            Autenticador.LimparCookieSessao
        Else
            linhaProcesso.FindElementByTag("td").FindElementByTag("a").Click
            AceitarAlerta 200
            'ExpandirAndamentosProcesso 8
            resposta = "sucesso"
        End If
    End If
    AbrirProcesso = resposta
End Function

Private Sub ExpandirAndamentosProcesso(qtdTelasAExpandir As Byte)
    
End Sub

Private Function PegarLinhaProcesso() As Selenium.WebElement
    Dim resposta As Selenium.WebElement
    
    On Error Resume Next
    Set resposta = oChrome.Chrome.FindElementByXPath("//*[@id='fPP:processosTable:tb']/tr[1]", 3000)
    On Error GoTo 0
    
    Set PegarLinhaProcesso = resposta
End Function

Private Sub EfetuarPesquisa(numeroProcesso As String)
    Dim numSequencial As String, numDigito As String, numAno As String, numOrgao As String
    
    numSequencial = Left(numeroProcesso, 7)
    numDigito = Mid(numeroProcesso, 9, 2)
    numAno = Mid(numeroProcesso, 12, 4)
    numOrgao = Right(numeroProcesso, 4)
    
    SisifoEmbasaFuncoes.PreencherTextboxChromedriver oChrome.Chrome.FindElementById("fPP:numeroProcesso:numeroSequencial"), numSequencial
    SisifoEmbasaFuncoes.PreencherTextboxChromedriver oChrome.Chrome.FindElementById("fPP:numeroProcesso:numeroDigitoVerificador"), numDigito
    SisifoEmbasaFuncoes.PreencherTextboxChromedriver oChrome.Chrome.FindElementById("fPP:numeroProcesso:Ano"), numAno
    SisifoEmbasaFuncoes.PreencherTextboxChromedriver oChrome.Chrome.FindElementById("fPP:numeroProcesso:NumeroOrgaoJustica"), numOrgao
    oChrome.Chrome.FindElementById("fPP:searchProcessos").Click
End Sub

Public Sub AbrirPaginaDeBuscaProcessual(PerfilLogado As String)
    Dim por As New Selenium.By
    Dim jaEstaNaPagina As Boolean
    Dim paginaDeBusca As String
    
    strPerfilLogado = PerfilLogado
    
    Do
        jaEstaNaPagina = (oChrome.Chrome.url = SisifoEmbasaFuncoes.sfUrlPJe1gTjbaBusca)
        paginaDeBusca = SisifoEmbasaFuncoes.sfUrlPJe1gTjbaBusca
        
        If jaEstaNaPagina = False Or _
        (jaEstaNaPagina = True And oChrome.Chrome.IsElementPresent(por.Name("fPP:numProcessoDiv")) = False) Then
            oChrome.Chrome.Get paginaDeBusca
        End If
    Loop Until oChrome.Chrome.IsElementPresent(por.ID("fPP:numeroProcesso:numeroSequencial"))
End Sub

Public Sub AbrirPaginaPainelUsuario()
    Dim urlPainelPje As String
    urlPainelPje = SisifoEmbasaFuncoes.sfUrlPJe1gTjbaPainel
    If oChrome.Chrome.url <> urlPainelPje Then oChrome.Chrome.Get urlPainelPje
End Sub

Public Function PegarTextoAlerta() As String
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

Private Sub AceitarAlerta(timeoutEmMs As Long)
    Dim alerta As Selenium.alert
    
    On Error Resume Next
    Set alerta = oChrome.Chrome.SwitchToAlert(timeoutEmMs)
    On Error GoTo 0
    
    If Not alerta Is Nothing Then alerta.accept
    
End Sub

Public Function PegarIdTableComarca(nomeComarca As String, idCaixa As String) As String
    Dim i As Selenium.WebElement
    Dim idDivComarcas As String
    
    idDivComarcas = Chrome.Chrome.FindElementById(idCaixa).FindElementByXPath("./../../..").Attribute("id")
    idDivComarcas = Replace(idDivComarcas, "linhaN1", "linhaN2")
    
    For Each i In Chrome.Chrome.FindElementById(idDivComarcas).FindElementsByTag("span")
        If i.text = nomeComarca Then
            PegarIdTableComarca = i.FindElementByXPath("./../../../../..").Attribute("id")
            Exit Function
        End If
    Next i
End Function

Public Sub FecharModal(idModal As String)
    Dim modal As Selenium.WebElement
    
    Set modal = oChrome.Chrome.FindElementById(idModal)
    modal.FindElementByTag("a").Click
End Sub
