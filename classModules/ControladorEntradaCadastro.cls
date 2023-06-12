VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControladorEntradaCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IControladorEntradaCadastro

Public Function IControladorEntradaCadastro_PegarObjetoComNumeroProcesso() As Object
    Set IControladorEntradaCadastro_PegarObjetoComNumeroProcesso = ActiveCell
End Function

Public Function IControladorEntradaCadastro_PegarNumeroProcessoDoObjetoToken(objToken As Object) As String
    Dim resposta As String
    
    Select Case VarType(objToken)
    Case 8
        resposta = SisifoEmbasaFuncoes.PegaNumeroProcessoDeCelula(objToken)
    End Select
    
    IControladorEntradaCadastro_PegarNumeroProcessoDoObjetoToken = resposta
End Function

Public Sub IControladorEntradaCadastro_MarcarAndamentoCadastrado(tokenNumeroProcesso As Object)
    Select Case VarType(tokenNumeroProcesso)
    Case 8
        'Dim token As Excel.Range
        'Set token = tokenNumeroProcesso
        With tokenNumeroProcesso
            .Interior.ColorIndex = 24
            .Offset(0, 1).Formula = "Inserido no Sísifo"
            .Offset(1, 0).Select
        End With
    End Select
    
End Sub

Public Function IControladorEntradaCadastro_ConfirmarEPegarInformacoesFinais(andamento As andamento, infoListas As StructInformacoesParaInterface) As StructInformacoesAndamento
    Dim form As frmAndamento
    Dim resposta As StructInformacoesAndamento
    
    Set form = ExibirForm(andamento, infoListas)
    resposta = PegarInformacoesDoForm(form)
    IControladorEntradaCadastro_ConfirmarEPegarInformacoesFinais = resposta
End Function

Private Function ExibirForm(andamento As andamento, infoListas As StructInformacoesParaInterface) As frmAndamento
    Dim form As New frmAndamento
    
    With form
        .txtNumProc = andamento.numeroProcesso
        .cmbAndamento.List = infoListas.ListaDeAndamentos
        .cmbAndamento.text = andamento.nomeAndamento
        .txtDataAndamento.text = SisifoEmbasaFuncoes.ValidaData(CStr(andamento.DataAndamento))
        .cmbProvidencia.List = infoListas.ListaDeProvidencias
        .cmbProvidencia.text = infoListas.NomeProvidenciaRelacionada
        If andamento.NovoJuizo = "" Then
            EsconderControlesJuizo form
        Else
            .cmbJuizo.List = infoListas.ListaDeJuizos
            .cmbJuizo.text = andamento.NovoJuizo
        End If
        .txtObsAndamento.text = andamento.obsAndamento
        '.txtObsAndamento.WordWrap = False ' Para contornar o problema do texto que encolhe.
        .txtObsAndamento.WordWrap = True ' Para contornar o problema do texto que encolhe.
ReexibirForm:
        
        Application.SendKeys "%{TAB}"

        .Show
    
        If Len(.txtObsAndamento.text) >= 32000 Then
            MsgBox SisifoEmbasaFuncoes.determinartratamento & ", o texto do andamento é muito extenso! Para evitar um erro no Espaider, é necessário excluir " & _
            Len(.txtObsAndamento.text) - 31999 & " caracteres.", vbOKOnly, "Sísifo - Erro no cadastro de andamento"
            GoTo ReexibirForm
        End If
    End With
        
    Set ExibirForm = form
End Function

Private Sub EsconderControlesJuizo(form As frmAndamento)
    With form
        .cmbJuizo.Visible = False
        .LabelJuizo.Visible = False
        .LinhaJuizo.Visible = False
        .LinhaBaseJuizo.Visible = False
        .cmdIr.Top = .cmdIr.Top - 36
        .Height = .Height - 36
    End With
End Sub

Private Function PegarInformacoesDoForm(form As frmAndamento) As StructInformacoesAndamento
    Dim infoAndamento As StructInformacoesAndamento
    
    With infoAndamento
        .nomeAndamento = Trim(form.cmbAndamento.text)
        .DataAndamento = Trim(form.txtDataAndamento.text)
        .obsAndamento = Trim(form.txtObsAndamento.text)
        .nomeProvidencia = Trim(form.cmbProvidencia.text)
        .NomeNovoJuizo = Trim(form.cmbJuizo.text)
        .DeveGerar = form.chbDeveGerar.Value
    End With
    
    PegarInformacoesDoForm = infoAndamento
End Function
