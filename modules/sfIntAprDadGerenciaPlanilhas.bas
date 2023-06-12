Attribute VB_Name = "sfIntAprDadGerenciaPlanilhas"
Sub RelatarProcessosArmazenadosIntimacoes(ByVal Controle As IRibbonControl)
''
'' Contar processos e registros existentes nas planilhas da memória
''
    Dim arrPlans(1 To 3) As Excel.Worksheet
    
    Set arrPlans(1) = sfCadAndamento
    Set arrPlans(2) = sfCadProvidencia
    Set arrPlans(3) = sfCadJurisdicao
    
    
    SisifoEmbasaFuncoes.RelatarProcessosArmazenados arrPlans
    
End Sub

' Gerar planilha para Espaider (Só exclui após usuário confimar que fez o upload no Espaider)
Sub ExportarPlanilhaAndamentoEspaider(ByVal Controle As IRibbonControl)
    Dim arq As Workbook
    Dim bolAndamentosVazia As Boolean, bolProvidenciasVazia As Boolean, bolNovosJuizosVazia As Boolean
    Dim lnUltimaLinhaAndamento As Long, lnUltimaLinhaProvidencias As Long, lnUltimaLinhaNovosJuizos As Long
    Dim strNome As String, strDesktop As String
    Dim contX As Byte
    
    bolAndamentosVazia = False
    bolProvidenciasVazia = False
    bolNovosJuizosVazia = False
    
    ' Caso a pasta de trabalho do Sísifo esteja sendo exibida (IsAddin), salva-a e oculta-a.
    If ThisWorkbook.IsAddin = False Then RestringirEdicaoRibbonIntimacoes Controle
    
    ' Pergunta se deseja continuar, para o caso de ser apertado sem querer (percebi que é fácil errar, botões
    ' próximos, e essa tarefa só roda uma vez por dia, portanto a confirmação não sobrecarrega o usuário.
    If MsgBox("Deseja gerar a planilha de exportação no formato do Espaider?", vbQuestion + vbYesNo, _
    "Sísifo - Exportar andamentos e DAJEs?") = vbNo Then Exit Sub
    
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
        MsgBox "As planilhas de andamentos estão vazias. Não há nada para exportar.", _
         vbInformation + vbOKOnly, "Sísifo - Planilhas vazias"
        Exit Sub
    End If

    ' Se as planilhas não estiverem vazias, exporta-as
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
    
    'Confirmar a inserção no Espaider
    If arq.Saved = False Then
        MsgBox "Não foi possível salvar o arquivo para exportação dos andamentos. Ele será fechado." & chr(13) & _
        "Caso o arquivo para exportação não seja fechado automaticamente, descarte-o e tente exportar " & _
        "novamente, até obter a confirmação da exportação.", vbCritical + vbOKOnly, _
        "Sísifo - Erro ao salvar o arquivo"
        arq.Close False
    Else
        If MsgBox("Confira se a planilha de andamentos foi salva na área de trabalho. " & _
            "Não esqueça de importar no Espaider. Caso não consiga fazer o upload no Espaider, " & _
            "tente novamente mais tarde.", vbExclamation + vbOKCancel + vbApplicationModal, _
            "Sísifo - Confirma exportação") = vbOK Then
            ' Usuário confirmou salvamento. Limpa a planilha de processos e salva.
            If bolAndamentosVazia = False Then sfCadAndamento.Rows("5:" & lnUltimaLinhaAndamento).Delete
            If bolProvidenciasVazia = False Then sfCadProvidencia.Rows("5:" & lnUltimaLinhaAndamento).Delete
            If bolNovosJuizosVazia = False Then sfCadJurisdicao.Rows("5:" & lnUltimaLinhaAndamento).Delete
            
            Application.DisplayAlerts = False
            ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, FileFormat:=xlOpenXMLAddIn
            Application.DisplayAlerts = True
        End If
    End If
End Sub
