VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControladorEntradaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get NomesUsuariosSelecionados() As Collection
    
End Property

Property Get NomesComarcasSelecionadas() As Collection

End Property

Property Get PeriodoSelecionado() As PeriodoDeTempo
    
End Property

Function MostrarFormConfirmandoSelecaoUsuarios(usuariosDisponiveis() As String) As Boolean

End Function

Public Function MostrarFormConfirmandoSelecaoComarcas(comarcas As Collection) As Boolean

End Function
