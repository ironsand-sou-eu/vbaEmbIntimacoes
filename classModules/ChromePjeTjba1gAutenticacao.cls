VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromePjeTjba1gAutenticacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IChromeAutenticador

Private chPje As ChromePjetjba1g
Private navegador As ChromeDriver

'Propriedade Cookies
Public Property Get IChromeAutenticador_Cookies() As Cookies
Set IChromeAutenticador_Cookies = chPje.Chrome.Cookies
End Property

'Propriedade Navegador
Public Property Get IChromeAutenticador_Navegador() As ChromeDriver
Set IChromeAutenticador_Navegador = navegador
End Property

Private Sub Class_Initialize()
    Set chPje = New ChromePjetjba1g
    Set navegador = chPje.Chrome.Chrome
End Sub

Public Function IChromeAutenticador_SessaoNaMemoriaEstaAtiva(infoCookieSessao As Selenium.Dictionary, PerfilLogado As String) As Boolean
    Dim resposta As Boolean
    
    If infoCookieSessao Is Nothing Then
        resposta = False
    Else
        navegador.Get SisifoEmbasaFuncoes.sfUrlPJe1gTjbaRecursoPequenoEstatico
        navegador.Manage.AddCookie Name:=infoCookieSessao("Name"), Value:=infoCookieSessao("Value"), domain:=infoCookieSessao("domain"), Path:=infoCookieSessao("Path"), Secure:=CBool(infoCookieSessao("Secure"))
        chPje.PerfilLogado = PerfilLogado
        chPje.AbrirPaginaPainelUsuario
        If navegador.FindElementsById("barraSuperiorPrincipal").Count > 0 Then
            resposta = True
        Else
            resposta = False
        End If
    End If
    
    IChromeAutenticador_SessaoNaMemoriaEstaAtiva = resposta
End Function

Public Function IChromeAutenticador_FazerLogin(usuario As UsuarioProjudiPje) As ChromeDriver
    Dim resultadoCaptcha As String
    Dim divAssinatura As Selenium.WebElement
    
InicioLogin:
    navegador.Get SisifoEmbasaFuncoes.sfUrlPJe1gTjbaLogin
    navegador.Window.Maximize
    navegador.ExecuteScript ("fecharMpModoOperacaoSemEscolherOpcao();")
    SisifoEmbasaFuncoes.PreencherTextboxSimulandoDigitacaoChromedriver navegador.FindElementById("username"), usuario.Login
    SisifoEmbasaFuncoes.PreencherTextboxSimulandoDigitacaoChromedriver navegador.FindElementById("password"), usuario.senha
    'On Error GoTo InicioLogin
    'oChrome.FindElementById("btnEntrar").Click
    'On Error GoTo 0
    resultadoCaptcha = LidarComCaptchaLogin
    
    If resultadoCaptcha <> "sucesso" Then
        IChromeAutenticador_FazerLogin = resultadoCaptcha
        Exit Function
    End If
    
    MudarPerfilLogadoParaRepresentante
    
    Set IChromeAutenticador_FazerLogin = navegador
 End Function

Private Function LidarComCaptchaLogin() As String
    Dim textoAlerta As String
    Dim resposta As String
    
    On Error GoTo Erro
inicio:
    Do
    Loop Until InStr(1, navegador.url, "login.seam") = 0
    On Error GoTo 0
    
    LidarComCaptchaLogin = "sucesso"
    Exit Function
    
Erro:
    textoAlerta = chPje.PegarTextoAlerta
    If textoAlerta = "sem alerta" Then
        GoTo inicio
    Else
        LidarComCaptchaLogin = textoAlerta
    End If
    
End Function

Private Function MudarPerfilLogadoParaRepresentante()
    Dim selectPapel As Selenium.WebElement
    
    On Error Resume Next
    Do
        Do
            navegador.FindElementByClass("dropdown-toggle").Click
        Loop Until navegador.FindElementById("papeisUsuarioForm").IsDisplayed = True
        
        Set selectPapel = navegador.FindElementById("papeisUsuarioForm:usuarioLocalizacaoDecoration:usuarioLocalizacao")
        selectPapel.AsSelect.SelectByText "Procuradoria - REPRESENTAÇÃO EMBASA / Procurador/Gestor"
    Loop Until selectPapel.AsSelect.SelectedOption.text = "Procuradoria - REPRESENTAÇÃO EMBASA / Procurador/Gestor"
    On Error GoTo 0
    
End Function
