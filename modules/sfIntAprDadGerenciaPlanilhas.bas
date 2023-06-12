Attribute VB_Name = "sfIntAprDadGerenciaPlanilhas"
Sub RelatarProcessosArmazenadosIntimacoes(ByVal Controle As IRibbonControl)
''
'' Contar processos e registros existentes nas planilhas da mem�ria
''
    Dim arrPlans(1 To 3) As Excel.Worksheet
    
    Set arrPlans(1) = sfCadAndamento
    Set arrPlans(2) = sfCadProvidencia
    Set arrPlans(3) = sfCadJurisdicao
    
    
    SisifoEmbasaFuncoes.RelatarProcessosArmazenados arrPlans
    
End Sub

' Gerar planilha para Espaider (S� exclui ap�s usu�rio confimar que fez o upload no Espaider)
Sub ExportarPlanilhaAndamentoEspaider(ByVal Controle As IRibbonControl)
    Dim arq As Workbook
    Dim bolAndamentosVazia As Boolean, bolProvidenciasVazia As Boolean, bolNovosJuizosVazia As Boolean
    Dim lnUltimaLinhaAndamento As Long, lnUltimaLinhaProvidencias As Long, lnUltimaLinhaNovosJuizos As Long
    Dim strNome As String, strDesktop As String
    Dim contX As Byte
    
    bolAndamentosVazia = False
    bolProvidenciasVazia = False
    bolNovosJuizosVazia = False
    
    ' Caso a pasta de trabalho do S�sifo esteja sendo exibida (IsAddin), salva-a e oculta-a.
    If ThisWorkbook.IsAddin = False Then RestringirEdicaoRibbonIntimacoes Controle
    
    ' Pergunta se deseja continuar, para o caso de ser apertado sem querer (percebi que � f�cil errar, bot�es
    ' pr�ximos, e essa tarefa s� roda uma vez por dia, portanto a confirma��o n�o sobrecarrega o usu�rio.
    If MsgBox("Deseja gerar a planilha de exporta��o no formato do Espaider?", vbQuestion + vbYesNo, _
    "S�sifo - Exportar andamentos e DAJEs?") = vbNo Then Exit Sub
    
    ' Testa a planilha CadAndamento
    lnUltimaLinhaAndamento = sfCadAndamento.UsedRange.Rows(sfCadAndamento.UsedRange.Rows.Count).Row
    If lnUltimaLinhaAndamento = 4 Then bolAndamentosVazia = True
    
    ' Testa a planilha CadProvidencia
    lnUltimaLinhaProvidencias = sfCadProvidencia.UsedRange.Rows(sfCadProvidencia.UsedRange.Rows.Count).Row
    If lnUltimaLinhaProvidencias = 4 Then bolProvidenciasVazia = True
    
    ' Testa a planilha CadNovoJuizo
    lnUltimaLinhaNovosJuizos = sfCadJurisdicao.UsedRange.Rows(sfCadJurisdicao.UsedRange.Rows.Count).Row
    If lnUltimaLinhaNovosJuizos = 4 Then bolNovosJuizosVazia = True
    
    ' Se estiverem todas vazias, avisa e para o procedimento
    If bolAndamentosVazia = True And bolProvidenciasVazia = True And bolNovosJuizosVazia = True Then
        MsgBox "As planilhas de andamentos est�o vazias. N�o h� nada para exportar.", _
         vbInformation + vbOKOnly, "S�sifo - Planilhas vazias"
        Exit Sub
    End If

    ' Se as planilhas n�o estiverem vazias, exporta-as
    Set arq = Workbooks.Add
    
    ' Se houver mais de uma planilha na pasta de trabalho, exclui as demais
    Application.DisplayAlerts = False
    If arq.Sheets.Count > 1 Then
        For contX = arq.Sheets.Count To 2
            arq.Sheets(contX).Delete
        Next contX
    End If
    Application.DisplayAlerts = True
    
    If bolAndamentosVazia = False Then _
        sfCadAndamento.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolProvidenciasVazia = False Then _
        sfCadProvidencia.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolNovosJuizosVazia = False Then _
        sfCadJurisdicao.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    Application.DisplayAlerts = False
    arq.Sheets(1).Delete
    Application.DisplayAlerts = True

    arq.SaveAs SisifoEmbasaFuncoes.CaminhoDesktop & "Sisifo - Andamentos - " & Format(Year(Now), "0000") & "." & Format(Month(Now), "00") & "." & Format(Day(Now), "00") & " " & Format(Hour(Time), "00") & "." & Format(Minute(Time), "00") & ".xlsx"
    
    'Confirmar a inser��o no Espaider
    If arq.Saved = False Then
        MsgBox "N�o foi poss�vel salvar o arquivo para exporta��o dos andamentos. Ele ser� fechado." & chr(13) & _
        "Caso o arquivo para exporta��o n�o seja fechado automaticamente, descarte-o e tente exportar " & _
        "novamente, at� obter a confirma��o da exporta��o.", vbCritical + vbOKOnly, _
        "S�sifo - Erro ao salvar o arquivo"
        arq.Close False
    Else
        If MsgBox("Confira se a planilha de andamentos foi salva na �rea de trabalho. " & _
            "N�o esque�a de importar no Espaider. Caso n�o consiga fazer o upload no Espaider, " & _
            "tente novamente mais tarde.", vbExclamation + vbOKCancel + vbApplicationModal, _
            "S�sifo - Confirma exporta��o") = vbOK Then
            ' Usu�rio confirmou salvamento. Limpa a planilha de processos e salva.
            If bolAndamentosVazia = False Then sfCadAndamento.Rows("5:" & lnUltimaLinhaAndamento).Delete
            If bolProvidenciasVazia = False Then sfCadProvidencia.Rows("5:" & lnUltimaLinhaAndamento).Delete
            If bolNovosJuizosVazia = False Then sfCadJurisdicao.Rows("5:" & lnUltimaLinhaAndamento).Delete
            
            Application.DisplayAlerts = False
            ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, FileFormat:=xlOpenXMLAddIn
            Application.DisplayAlerts = True
        End If
    End If
End Sub
