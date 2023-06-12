VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PeriodoDeTempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dtDataInicial As Date, dtDatafinal As Date

'Propriedade DataInicial:
Public Property Get DataInicial() As Date
DataInicial = dtDataInicial
End Property
Public Property Let DataInicial(dt As Date)
dtDataInicial = dt
End Property

'Propriedade Datafinal:
Public Property Get Datafinal() As Date
Datafinal = dtDatafinal
End Property
Public Property Let Datafinal(dt As Date)
dtDatafinal = dt
End Property
