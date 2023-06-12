VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControladorEntradaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IControladorEntradaLista

Private NomesUsuariosSelecionados As Collection, NomesComarcasSelecionadas As Collection
Private PeriodoSelecionado As PeriodoDeTempo
Private formAdvogados As frmAdvs
Private formComarcas As frmComarcas

Property Get IControladorEntradaLista_NomesUsuariosSelecionados() As Collection
    CarregarNomesUsuariosAColetar
    Set IControladorEntradaLista_NomesUsuariosSelecionados = NomesUsuariosSelecionados
End Property

Property Get IControladorEntradaLista_NomesComarcasSelecionadas() As Collection
    CarregarNomesComarcasAColetar
    Set IControladorEntradaLista_NomesComarcasSelecionadas = NomesComarcasSelecionadas
End Property

Property Get IControladorEntradaLista_PeriodoSelecionado() As PeriodoDeTempo
    CarregarPeriodoAColetar
    Set IControladorEntradaLista_PeriodoSelecionado = PeriodoSelecionado
End Property

Public Function IControladorEntradaLista_MostrarFormConfirmandoSelecaoUsuarios(usuariosDisponiveis() As String) As Boolean
    Set formAdvogados = New frmAdvs
    
    With formAdvogados
        If Weekday(Date) = vbMonday Then
            .txtDataInicial.text = CStr(Date - 3)
        Else
            .txtDataInicial.text = CStr(Date - 1)
        End If
        SisifoEmbasaFuncoes.AjustarLegendaSemTransicao .LabelDataInicial
        .txtDataFinal.text = CStr(Date - 1)
        SisifoEmbasaFuncoes.AjustarLegendaSemTransicao .LabelDataFinal
        .lsAdvs.List = usuariosDisponiveis()
        .lsAdvs.AddItem "Todos", 0
        .lsAdvs.ListIndex = 0
        .lsAdvs.Selected(0) = True
        .Show
    
        IControladorEntradaLista_MostrarFormConfirmandoSelecaoUsuarios = .chbDeveGerar
    End With
    
End Function

Private Sub CarregarPeriodoAColetar()
    Set PeriodoSelecionado = New PeriodoDeTempo
    
    With formAdvogados
        'Se não estiverem, coloca as datas-parâmetro em ordem crescente
        If CDate(.txtDataFinal) >= CDate(.txtDataInicial) Then
            PeriodoSelecionado.DataInicial = .txtDataInicial
            PeriodoSelecionado.Datafinal = .txtDataFinal
        Else
            PeriodoSelecionado.DataInicial = .txtDataFinal
            PeriodoSelecionado.Datafinal = .txtDataInicial
        End If
    End With
    
End Sub

Private Sub CarregarNomesUsuariosAColetar()
    Dim i As Long
    
    Set NomesUsuariosSelecionados = New Collection
    For i = 1 To formAdvogados.lsAdvs.ListCount - 1 ' A partir do segundo item porque o primeiro (index 0) é o "Todos".
        If formAdvogados.lsAdvs.Selected(i) = True Then
            NomesUsuariosSelecionados.Add formAdvogados.lsAdvs.List(i, 0)
        End If
    Next i

End Sub

Public Function IControladorEntradaLista_MostrarFormConfirmandoSelecaoComarcas(comarcas As Collection) As Boolean
    Dim i As Integer
    
    Set formComarcas = New frmComarcas
    
    With formComarcas
        For i = 1 To comarcas.Count
            .lsComarcas.AddItem comarcas(i)
        Next i
        .lsComarcas.AddItem "Todos", 0
        .lsComarcas.ListIndex = 0
        .lsComarcas.Selected(0) = True
        .Show
        
        IControladorEntradaLista_MostrarFormConfirmandoSelecaoComarcas = .chbDeveGerar
    End With
    
End Function

Private Sub CarregarNomesComarcasAColetar()
    Dim i As Long
    
    Set NomesComarcasSelecionadas = New Collection
    For i = 1 To formComarcas.lsComarcas.ListCount - 1 ' A partir do segundo item porque o primeiro (index 0) é o "Todos".
        If formComarcas.lsComarcas.Selected(i) = True Then
            NomesComarcasSelecionadas.Add formComarcas.lsComarcas.List(i, 0)
        End If
    Next i

End Sub
