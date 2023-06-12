Attribute VB_Name = "sfIntRegNegFuncoesRibbon"
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
 
Sub FechaConfigIntimacoesVisivel(ByVal Controle As IRibbonControl, Optional ByRef returnedVal)
    SisifoEmbasaFuncoes.FechaConfigVisivel ThisWorkbook, cfIntConfigura��es, Controle, returnedVal
End Sub

Private Sub AoCarregarRibbonIntimacoes(Ribbon As IRibbonUI)
    ' Chama a fun��o geral AoCarregarRibbon com os par�metros corretos.
    SisifoEmbasaFuncoes.AoCarregarRibbon cfIntConfigura��es, Ribbon
End Sub

Sub LiberarEdicaoIntimacoes(ByVal Controle As IRibbonControl)
    ' Chama a fun��o geral LiberarEdicao
    SisifoEmbasaFuncoes.LiberarEdicao ThisWorkbook, cfIntConfigura��es
    
End Sub

Sub RestringirEdicaoRibbonIntimacoes(ByVal Controle As IRibbonControl)
    ' Chama a fun��o geral RestringirEdicaoRibbon
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook, cfIntConfigura��es, Controle
End Sub

Sub sfIntCmbSistemaMudou(ByVal Controle As IRibbonControl, ByVal text As String)
    Dim sistema As SisifoEmbasaFuncoes.sfSistema, tribunal As SisifoEmbasaFuncoes.sfTribunal
    Dim valorSelecionado As String
    
    valorSelecionado = LCase(Trim(text))
    If InStr(1, valorSelecionado, "projudi") <> 0 Then
        sistema = sfSistema.projudi
    ElseIf InStr(1, valorSelecionado, "pje1g") <> 0 Or InStr(1, valorSelecionado, "pje 1g") <> 0 Then
        sistema = sfSistema.pje1g
    ElseIf InStr(1, valorSelecionado, "pje2g") <> 0 Or InStr(1, valorSelecionado, "pje 2g") <> 0 Then
        sistema = sfSistema.pje2g
    Else
        sistema = sfSistema.Erro
    End If
    
    If InStr(1, valorSelecionado, "tj/ba") <> 0 Or InStr(1, valorSelecionado, "tjba") <> 0 Then
        tribunal = sfTribunal.Tjba
    ElseIf InStr(1, valorSelecionado, "trt5") <> 0 Then
        tribunal = sfTribunal.trt5
    Else
        tribunal = sfTribunal.Erro
    End If
    
    cfIntConfigura��es.Cells().Find(What:="Sistema selecionado", LookAt:=xlWhole).Offset(0, 1).Formula = sistema
    cfIntConfigura��es.Cells().Find(What:="Tribunal selecionado", LookAt:=xlWhole).Offset(0, 1).Formula = tribunal
    
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook, cfIntConfigura��es ' Salva as altera��es
End Sub

Sub sfIntCmbSistemaTexto(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim sistema As SisifoEmbasaFuncoes.sfSistema, tribunal As SisifoEmbasaFuncoes.sfTribunal
    Dim textoFinal As String
    
    sistema = CLng(cfIntConfigura��es.Cells().Find(What:="Sistema selecionado", LookAt:=xlWhole).Offset(0, 1).Formula)
    tribunal = CLng(cfIntConfigura��es.Cells().Find(What:="Tribunal selecionado", LookAt:=xlWhole).Offset(0, 1).Formula)
    
    Select Case sistema
    Case projudi
        textoFinal = "Projudi "
    Case pje1g
        textoFinal = "PJe1g "
    Case pje2g
         textoFinal = "PJe2g "
    Case Else
         textoFinal = "Erro"
    End Select
    
    Select Case tribunal
    Case Tjba
        textoFinal = textoFinal & "TJ/BA"
    Case trt5
        textoFinal = textoFinal & "TRT5"
    Case Else
        textoFinal = "Erro"
    End Select
    
    returnedVal = textoFinal
End Sub

Function PegaDataFinalProvidencia(Optional ByVal Controle As IRibbonControl, Optional ByRef returnedVal) As Date
''
'' Retorna a data final da provid�ncia prevista na planilha cfIntConfigura��es.
''
    
    If Controle Is Nothing Then
        PegaDataFinalProvidencia = CDate(cfIntConfigura��es.Cells().Find(What:="Criar provid�ncias para", LookAt:=xlWhole).Offset(0, 1).text)
    Else
        If Controle.ID = "edIntConfigDataFinal" Then
            returnedVal = CDate(cfIntConfigura��es.Cells().Find(What:="Criar provid�ncias para", LookAt:=xlWhole).Offset(0, 1).text)
        End If
    End If
End Function

Sub AjustaDataFinalProvidencia(ByVal Controle As IRibbonControl, ByRef text)
''
'' Atribui � data final de provid�ncias prevista na planilha cfIntConfigura��es o valor determinado pelo usu�rio.
''
    Dim rbSisifoUI As IRibbonUI
    Dim varDataFinalProv As Variant
    Dim lnDataFinalProv As Long
    
    Set rbSisifoUI = SisifoEmbasaFuncoes.RecuperarObjetoPorReferencia(ThisWorkbook, cfIntConfigura��es)
    
    varDataFinalProv = text
    varDataFinalProv = Replace(varDataFinalProv, " ", "")
    varDataFinalProv = Replace(varDataFinalProv, "/", "")
    
    'Primeiro erro: n�o ser composto por n�meros
    On Error Resume Next
    lnDataFinalProv = CLng(varDataFinalProv)
    On Error GoTo 0
    If varDataFinalProv <> lnDataFinalProv Then
        MsgBox "Munificente Mestre, o valor informado n�o parece corresponder a uma data. Favor tentar novamente, utilizando apenas n�meros no formato " & _
        "DD/MM/AAAA ou DD/MM/AA, podendo ou n�o usar as barras.", vbCritical + vbOKOnly, "S�sifo - Erro de data"
        
        rbSisifoUI.InvalidateControl (Controle.ID)
        Exit Sub
    End If
    
    ' Formatar conforme tamanho
    If Len(varDataFinalProv) = 5 Or Len(varDataFinalProv) = 6 Then 'Dia, m�s e ano com dois d�gitos
        varDataFinalProv = Format(varDataFinalProv, "00/00/00")
        varDataFinalProv = Left(varDataFinalProv, 6) & "20" & Mid(varDataFinalProv, 7)
    ElseIf Len(varDataFinalProv) = 7 Or Len(varDataFinalProv) = 8 Then
        varDataFinalProv = Format(varDataFinalProv, "00/00/0000")
    End If
    
    ' Segundo erro: data retroativa.
    If CDate(varDataFinalProv) <= Date Then
        MsgBox "Em�rito Mestre, a data informada � anterior ou igual � atual. N�o queremos enlouquecer nossos colegas com provid�ncas retroativas! Favor " & _
            "tentar novamente, usando datas a partir de amanh� (" & Format(Date + 1, "dd/mm/yyyy") & ").", vbCritical + vbOKOnly, "S�sifo - Erro de data"
        rbSisifoUI.InvalidateControl (Controle.ID)
        Exit Sub
    End If
    
    ' Sem erros, coloca na planilha.
    varDataFinalProv = "'" & CDate(varDataFinalProv)
    cfIntConfigura��es.Cells().Find(What:="Criar provid�ncias para", LookAt:=xlWhole).Offset(0, 1).Formula = CStr(varDataFinalProv)
    
    rbSisifoUI.InvalidateControl (Controle.ID)
    MsgBox "Inesgot�vel Mestre, a data final das provid�ncias foi alterada com sucesso para " & Replace(varDataFinalProv, "'", "") & ". As provid�ncias " & _
    "criadas a partir de agora ter�o essa data como data final.", vbInformation + vbOKOnly, "S�sifo - Data estabelecida com sucesso"
    
End Sub
