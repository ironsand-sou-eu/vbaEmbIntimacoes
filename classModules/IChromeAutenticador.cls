VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IChromeAutenticador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Cookies() As Cookies

End Property

Public Property Get navegador() As ChromeDriver

End Property

Public Function SessaoNaMemoriaEstaAtiva(infoCookieSessao As Selenium.Dictionary, PerfilLogado As String) As Boolean

End Function

Public Function FazerLogin(usuario As UsuarioProjudiPje) As ChromeDriver
 
End Function
