VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioProjudiPje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strNomeUsuario As String, strLogin As String, strSenha As String
Private lnQtdIntimacoes As Long

'Propriedade NomeUsuario:
Public Property Get nomeUsuario() As String
nomeUsuario = strNomeUsuario
End Property
Public Property Let nomeUsuario(str As String)
strNomeUsuario = str
End Property

'Propriedade Login:
Public Property Get Login() As String
Login = strLogin
End Property
Public Property Let Login(str As String)
strLogin = str
End Property

'Propriedade Senha:
Public Property Get senha() As String
senha = strSenha
End Property
Public Property Let senha(str As String)
strSenha = str
End Property

'Propriedade QtdIntimacoes:
Public Property Get qtdIntimacoes() As Long
qtdIntimacoes = lnQtdIntimacoes
End Property
Public Property Let qtdIntimacoes(ln As Long)
lnQtdIntimacoes = ln
End Property
