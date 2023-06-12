VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListaDeIntimacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sfUsuario As UsuarioProjudiPje
Private sfPeriodoDeExpedicao As PeriodoDeTempo
Private sfInstancia As SisifoEmbasaFuncoes.sfInstancia
Private sfSistema As SisifoEmbasaFuncoes.sfSistema
Private strMensagemErro As String, strComarca As String, strOrigem As String
Private colIntimacoes As Collection

'Propriedade Usuario:
Public Property Get usuario() As UsuarioProjudiPje
Set usuario = sfUsuario
End Property
Public Property Set usuario(usr As UsuarioProjudiPje)
Set sfUsuario = usr
End Property

'Propriedade PeriodoDeExpedicao:
Public Property Get PeriodoDeExpedicao() As PeriodoDeTempo
Set PeriodoDeExpedicao = sfPeriodoDeExpedicao
End Property
Public Property Set PeriodoDeExpedicao(per As PeriodoDeTempo)
Set sfPeriodoDeExpedicao = per
End Property

'Propriedade Intimacoes:
Public Property Get intimacoes() As Collection
Set intimacoes = colIntimacoes
End Property
Public Property Set intimacoes(col As Collection)
Set colIntimacoes = col
End Property

'Propriedade Instancia:
Public Property Get instancia() As SisifoEmbasaFuncoes.sfInstancia
instancia = sfInstancia
End Property
Public Property Let instancia(sf As SisifoEmbasaFuncoes.sfInstancia)
sfInstancia = sf
End Property

'Propriedade Sistema:
Public Property Get sistema() As SisifoEmbasaFuncoes.sfSistema
sistema = sfSistema
End Property
Public Property Let sistema(sf As SisifoEmbasaFuncoes.sfSistema)
sfSistema = sf
End Property

'Propriedade Origem:
Public Property Get Origem() As String
Origem = strOrigem
End Property
Public Property Let Origem(str As String)
strOrigem = str
End Property

'Propriedade Comarca:
Public Property Get Comarca() As String
Comarca = strComarca
End Property
Public Property Let Comarca(str As String)
strComarca = str
End Property

'Propriedade MensagemErro:
Public Property Get mensagemErro() As String
mensagemErro = strMensagemErro
End Property
Public Property Let mensagemErro(str As String)
strMensagemErro = str
End Property
