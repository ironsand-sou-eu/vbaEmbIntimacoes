VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Chrome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oChr As Selenium.ChromeDriver

'Propriedade oChrome:
Public Property Get Chrome() As Selenium.ChromeDriver
Set Chrome = oChr
End Property
Public Property Set Chrome(chr As Selenium.ChromeDriver)
Set oChr = chr
End Property

'Propriedade Cookies:
Public Property Get Cookies() As Selenium.Cookies
Set Cookies = oChr.Manage.Cookies
End Property

Private Sub Class_Initialize()
    Set oChr = New Selenium.ChromeDriver
End Sub

Public Sub EsperarElementoAparecer(parametroBusca As String, tipoBusca As String)
    Dim elemento As Selenium.WebElement
    
    On Error Resume Next
    Do
        Select Case LCase(tipoBusca)
        Case "id"
            Set elemento = oChr.FindElementById(parametroBusca, Raise:=False)
        Case "name"
            Set elemento = oChr.FindElementByName(parametroBusca, Raise:=False)
        Case "xpath"
            Set elemento = oChr.FindElementByXPath(parametroBusca, Raise:=False)
        End Select
    Loop Until elemento.IsDisplayed = True
    On Error GoTo 0
End Sub

Public Sub EsperarElementoSumir(parametroBusca As String, tipoBusca As String)
    Dim elemento As Selenium.WebElement
    Do
        Select Case LCase(tipoBusca)
        Case "id"
            Set elemento = oChr.FindElementById(parametroBusca, Raise:=False)
        Case "name"
            Set elemento = oChr.FindElementByName(parametroBusca, Raise:=False)
        Case "xpath"
            Set elemento = oChr.FindElementByXPath(parametroBusca, Raise:=False)
        End Select
    Loop Until elemento.IsDisplayed = False
End Sub

