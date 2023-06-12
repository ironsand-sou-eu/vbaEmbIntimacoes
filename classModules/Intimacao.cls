VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Intimacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strNumeroProcesso As String, strDestinatario As String, strModoExpedicao As String
Private dtDataExpedicao As Date
Private bolSuspeitaDeNaoSerEmbasa As Boolean

'Propriedade NumeroProcesso:
Public Property Get numeroProcesso() As String
numeroProcesso = strNumeroProcesso
End Property
Public Property Let numeroProcesso(str As String)
strNumeroProcesso = str
End Property

'Propriedade Destinatario:
Public Property Get Destinatario() As String
Destinatario = strDestinatario
End Property
Public Property Let Destinatario(str As String)
strDestinatario = str
End Property

'Propriedade ModoExpedicao:
Public Property Get ModoExpedicao() As String
ModoExpedicao = strModoExpedicao
End Property
Public Property Let ModoExpedicao(str As String)
strModoExpedicao = str
End Property

'Propriedade DataExpedicao:
Public Property Get dataExpedicao() As String
dataExpedicao = CStr(Format(dtDataExpedicao, "dd/mm/yyyy"))
End Property
Public Property Let dataExpedicao(str As String)
dtDataExpedicao = CDate(str)
End Property

'Propriedade SuspeitaDeNaoSerEmbasa:
Public Property Get SuspeitaDeNaoSerEmbasa() As Boolean
SuspeitaDeNaoSerEmbasa = bolSuspeitaDeNaoSerEmbasa
End Property
Public Property Let SuspeitaDeNaoSerEmbasa(bln As Boolean)
bolSuspeitaDeNaoSerEmbasa = bln
End Property


