VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeProjudiAutenticacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IChromeAutenticador

Private chProjudi As ChromeProjudi
Private navegador As ChromeDriver

'Propriedade Cookies
Public Property Get IChromeAutenticador_Cookies() As Cookies
Set IChromeAutenticador_Cookies = chProjudi.Chrome.Cookies
End Property

'Propriedade Navegador
Public Property Get IChromeAutenticador_Navegador() As ChromeDriver
Set IChromeAutenticador_Navegador = navegador
End Property

Private Sub Class_Initialize()
    Set chProjudi = New ChromeProjudi
    Set navegador = chProjudi.Chrome.Chrome
End Sub

Public Function IChromeAutenticador_SessaoNaMemoriaEstaAtiva(infoCookieSessao As Selenium.Dictionary, PerfilLogado As String) As Boolean
    Dim resposta As Boolean
    
    If infoCookieSessao Is Nothing Then
        resposta = False
    Else
        navegador.Get SisifoEmbasaFuncoes.sfUrlProjudiRecursoPequenoEstatico
        navegador.Manage.AddCookie Name:=infoCookieSessao("Name"), Value:=infoCookieSessao("Value"), domain:=infoCookieSessao("domain"), Path:=infoCookieSessao("Path"), Secure:=CBool(infoCookieSessao("Secure"))
        chProjudi.PerfilLogado = PerfilLogado
        chProjudi.AbrirPaginaDeBuscaProcessual PerfilLogado
        If navegador.FindElementsByTag("a").Count > 0 Then
            resposta = True
        Else
            resposta = False
        End If
    End If
    
    IChromeAutenticador_SessaoNaMemoriaEstaAtiva = resposta
End Function

Public Function IChromeAutenticador_FazerLogin(usuario As UsuarioProjudiPje) As ChromeDriver
    Dim resultadoCaptcha As String
    
    navegador.Get SisifoEmbasaFuncoes.sfUrlProjudiLogin
    navegador.SwitchToFrame 1
    SisifoEmbasaFuncoes.PreencherTextboxSimulandoDigitacaoChromedriver navegador.FindElementById("login"), usuario.Login
    SisifoEmbasaFuncoes.PreencherTextboxSimulandoDigitacaoChromedriver navegador.FindElementById("senha"), usuario.senha
    navegador.FindElementByXPath("//*[@id=""formLogin""]/table/tbody/tr[6]/td[2]/a").Click
    resultadoCaptcha = chProjudi.LidarComCaptcha
    'TODO: Não se faz nada com o resultado do Captcha. Isso pode resultar em algum erro não sendo tratado.
    navegador.SwitchToParentFrame
    
    If InStr(1, usuario.nomeUsuario, "Representante") <> 0 Then ' Se for perfil de representante
        navegador.SwitchToFrame 1
        navegador.SwitchToFrame 0
        navegador.FindElementByPartialLinkText("EMBASA", 0).Click
        chProjudi.AguardarCarregarHome
    End If
    
    Set IChromeAutenticador_FazerLogin = navegador
    
 End Function
