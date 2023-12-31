VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Andamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oOrigemNumProcesso As Object
Private strNumProc As String, strNomeAndamento As String, strObsAndamento As String, strNovaComarca As String
Private strNovoOrgao As String, strNovoJuizo As String, strGerencia As String, strIdConversao As String, strMensagemErro As String
Private sfInstancia As SisifoEmbasaFuncoes.sfInstancia
Private ssftribunal As SisifoEmbasaFuncoes.sfTribunal
Private ssfsistema As SisifoEmbasaFuncoes.sfSistema
Private dtDataAndamento As Date
Private curValorCausa As Currency
Private listaProvidencias As Collection

Private Sub Class_Initialize()
    'Seta as cole��es
    Set listaProvidencias = New Collection
End Sub

'Propriedade Providencias:
Public Property Get Providencias() As Collection
Set Providencias = listaProvidencias
End Property

'Propriedade ObjetoTokenDoNumeroProcesso:
Public Property Get ObjetoTokenDoNumeroProcesso() As Object
Set ObjetoTokenDoNumeroProcesso = oOrigemNumProcesso
End Property
Public Property Set ObjetoTokenDoNumeroProcesso(obj As Object)
Set oOrigemNumProcesso = obj
End Property

'Propriedade NumeroProcesso:
Public Property Get numeroProcesso() As String
numeroProcesso = strNumProc
End Property
Public Property Let numeroProcesso(str As String)
strNumProc = str
End Property

'Propriedade ValorCausa:
Public Property Get ValorCausa() As Currency
ValorCausa = curValorCausa
End Property
Public Property Let ValorCausa(valor As Currency)
curValorCausa = valor
End Property

'Propriedade NomeAndamento:
Public Property Get nomeAndamento() As String
nomeAndamento = strNomeAndamento
End Property
Public Property Let nomeAndamento(str As String)
strNomeAndamento = str
End Property

'Propriedade DataAndamento:
Public Property Get DataAndamento() As Date
DataAndamento = dtDataAndamento
End Property
Public Property Let DataAndamento(dt As Date)
dtDataAndamento = dt
End Property

'Propriedade ObsAndamento:
Public Property Get obsAndamento() As String
obsAndamento = strObsAndamento
End Property
Public Property Let obsAndamento(str As String)
'strObsAndamento = Replace(str, ChrW(&H2028), "") ' Esse funcionava!
strObsAndamento = Replace(str, ChrW(2), "") ' Remove caracteres especiais Unicode que bugam o XML do Excel.
strObsAndamento = Replace(strObsAndamento, ChrW(11), "") ' Remove  caracteres especiais Unicode que bugam o XML do Excel.
strObsAndamento = Replace(strObsAndamento, ChrW(30), "-") ' Remove caracteres especiais Unicode que bugam o XML do Excel.
strObsAndamento = strObsAndamento
End Property

'Propriedade NovaComarca:
Public Property Get NovaComarca() As String
NovaComarca = strNovaComarca
End Property
Public Property Let NovaComarca(str As String)
strNovaComarca = str
End Property

'Propriedade NovoOrgao:
Public Property Get NovoOrgao() As String
NovoOrgao = strNovoOrgao
End Property
Public Property Let NovoOrgao(str As String)
strNovoOrgao = str
End Property

'Propriedade NovoJuizo:
Public Property Get NovoJuizo() As String
NovoJuizo = strNovoJuizo
End Property
Public Property Let NovoJuizo(str As String)
strNovoJuizo = str
End Property

'Propriedade Sistema
Public Property Get sistema() As SisifoEmbasaFuncoes.sfSistema
sistema = ssfsistema
End Property
Public Property Let sistema(sis As SisifoEmbasaFuncoes.sfSistema)
ssfsistema = sis
End Property

'Propriedade Tribunal
Public Property Get tribunal() As SisifoEmbasaFuncoes.sfTribunal
tribunal = ssftribunal
End Property
Public Property Let tribunal(trib As SisifoEmbasaFuncoes.sfTribunal)
ssftribunal = trib
End Property

'Propriedade Instancia:
Public Property Get instancia() As SisifoEmbasaFuncoes.sfInstancia
instancia = sfInstancia
End Property
Public Property Let instancia(inst As SisifoEmbasaFuncoes.sfInstancia)
sfInstancia = inst
End Property

'Propriedade Gerencia:
Public Property Get Gerencia() As String
Gerencia = strGerencia
End Property
Public Property Let Gerencia(str As String)
strGerencia = str
End Property

'Propriedade IdConversao:
Public Property Get IdConversao() As String
IdConversao = strIdConversao
End Property
Public Property Let IdConversao(str As String)
strIdConversao = str
End Property

'Propriedade MensagemErro:
Public Property Get mensagemErro() As String
mensagemErro = strMensagemErro
End Property
Public Property Let mensagemErro(str As String)
strMensagemErro = str
End Property

Public Sub InferirECarregarInstancia()
    Dim numOrgao As String
    Dim resposta As SisifoEmbasaFuncoes.sfInstancia
    
    numOrgao = Right(strNumProc, 4)
    If numOrgao = "9000" Then
        resposta = SisifoEmbasaFuncoes.sfInstancia.SegundoGrau
    Else
        resposta = SisifoEmbasaFuncoes.sfInstancia.PrimeiroGrau
    End If
    sfInstancia = resposta
End Sub

Function LidarComErro() As String
    If strMensagemErro = "" Then
        LidarComErro = ""
    Else
        LidarComErro = "erro"
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", ocorreu um erro no cadastro. O andamento n�o foi cadastrado. Erro: " & vbCrLf _
            & strMensagemErro, vbCritical + vbOKOnly, "S�sifo - Erro no cadastro"
    End If
End Function

Public Function PerguntarSeProssegueAposInconsistencias() As VbMsgBoxResult
    Dim numJustica As String, numDivisaoJustica As String
    Dim resposta As VbMsgBoxResult
    
    numJustica = Mid(strNumProc, 17, 1)
    numDivisaoJustica = Mid(strNumProc, 19, 2)
    
    If ssftribunal = sfTribunal.Erro Then
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", houve um erro ao buscar o tribunal selecionado. Suplico que " & _
        "aviseis ao administrador do sistema.", vbOKOnly, "S�sifo - Erro"
        resposta = vbNo
    
    ElseIf ssfsistema = sfSistema.Erro Then
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", houve um erro ao buscar o sistema selecionado. Suplico que " & _
        "aviseis ao administrador do sistema.", vbOKOnly, "S�sifo - Erro"
        resposta = vbNo
    
    ElseIf numJustica <> "8" And numJustica <> "5" Then ' Erros: Se n�o for TJ ou TRT
        resposta = MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", a julgar pelo n�mero, o processo n�o pertence a TJ " & _
        "nem a TRT, e ainda n�o � cadastrado pelo S�sifo. Desejais tentar prosseguir no tribunal informado?", vbYesNo, _
        "S�sifo - Confirma��o")
    
    ElseIf numJustica = "8" And numDivisaoJustica <> "05" Then ' Justi�a estadual de outros estados
        resposta = MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", a julgar pelo n�mero, o processo pertence a um TJ " & _
        "de outro estado que n�o a Bahia, e ainda n�o � cadastrado pelo S�sifo. Desejais tentar prosseguir no tribunal informado?", _
        vbYesNo, "S�sifo - Confirma��o")
    
    ElseIf numJustica = "5" And numDivisaoJustica <> "05" Then ' Justi�a do Trabalho de outros regionais
        resposta = MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", a julgar pelo n�mero, o processo pertence a um TRT " & _
        "de outra regi�o que n�o a 5�, e ainda n�o � cadastrado pelo S�sifo. Desejais tentar prosseguir no tribunal informado?", _
        vbYesNo, "S�sifo - Confirma��o")
        
    ElseIf numJustica = "8" And numDivisaoJustica = "05" Then ' TJ/BA
        If ssftribunal <> Tjba Then
            resposta = MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", embora tenhais selecionado um sistema de outro " & _
            "tribunal, o n�mero parece indicar que o processo pertence, em verdade, ao TJ/BA. Imploro que verifiqueis a " & _
            "inconsist�ncia. Desejais tentar prosseguir no tribunal informado?", vbYesNo, "S�sifo - Confirma��o")
        End If

    ElseIf numJustica = "5" And numDivisaoJustica = "05" Then ' TRT 5� Regi�o
        If ssftribunal <> trt5 Then
            resposta = MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", embora tenhais selecionado um sistema de outro " & _
            "tribunal, o n�mero parece indicar que o processo pertence, em verdade, ao TRT da 5� regi�o. Imploro que " & _
            "verifiqueis a inconsist�ncia. Desejais tentar prosseguir no tribunal informado?", vbYesNo, "S�sifo - Confirma��o")
        End If
    
    ElseIf ssftribunal = Tjba Then
        If Left(strNumProc, 1) <> "0" Or Left(strNumProc, 2) = "03" _
        Or Left(strNumProc, 2) = "05" Then   ' Se n�o come�ar com 0, � PJe. Se come�ar com 03 ou 05, � eSaj (o S�sifo vai tratar como PJe).
            If Not (ssfsistema = pje1g Or ssfsistema = pje2g) Then
                resposta = MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", embora tenhais selecionado outro sistema, " & _
                "o n�mero parece indicar que o processo pertence, em verdade, ao PJe. Imploro que " & _
                "verifiqueis a inconsist�ncia. Desejais tentar prosseguir no sistema informado?", vbYesNo, "S�sifo - Confirma��o")
            End If
        
        Else ' Nos demais casos, � Projudi
            If Not ssfsistema = projudi Then
                resposta = MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", embora tenhais selecionado outro sistema, " & _
                "o n�mero parece indicar que o processo pertence, em verdade, ao Projudi. Imploro que " & _
                "verifiqueis a inconsist�ncia. Desejais tentar prosseguir no sistema informado?", vbYesNo, "S�sifo - Confirma��o")
            End If
        End If
    End If
    
    PerguntarSeProssegueAposInconsistencias = resposta
End Function


