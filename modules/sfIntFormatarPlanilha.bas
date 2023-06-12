Attribute VB_Name = "sfIntFormatarPlanilha"
Sub FormatarIntimacoes()

    Dim arq As Workbook
    Dim plan As Worksheet
    Dim rngTudoSemCabecalho As Range
    Dim lngUltimaLinha As Long

    Set arq = ActiveWorkbook
    Set plan = ActiveSheet
    
    If arq Is Nothing Or plan Is Nothing Then
        MsgBox "Parece não haver nenhum arquivo ou planilha do Excel aberto. Abra uma planilha e tente novamente.", vbCritical + vbOKOnly, "Planilha não encontrada"
        End
    End If
    
    If arq.Sheets.Count <= 1 Then arq.Sheets.Add after:=arq.Sheets(1)
    
    plan.UsedRange.Range("$B:$G").Copy

    arq.Sheets(2).Cells().PasteSpecial Paste:=xlPasteValues, SkipBlanks:=True
    Application.DisplayAlerts = False
    arq.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    Set plan = ActiveSheet
    With plan
        .Range("$A:$C").ColumnWidth = 25
        .Range("$F:$F").ColumnWidth = 25
        .Range("$D:$E").Delete
        
        .Rows(1).Insert xlShiftDown, xlFormatFromRightOrBelow
        .Cells(1, 1).Formula = "Processo"
        .Cells(1, 2).Formula = "Réu"
        .Cells(1, 3).Formula = "Expedição"
        .Cells(1, 4).Formula = "Leitura"
        .Range(Cells(1, 1), Cells(1, 4)).Font.Bold = True
        
        .UsedRange.RemoveDuplicates Array(1, 3), xlYes
        
        lngUltimaLinha = .UsedRange.Rows.Count
        Set rngTudoSemCabecalho = Range("A2:D" & lngUltimaLinha)
        
        rngTudoSemCabecalho.Sort rngTudoSemCabecalho.Columns(3), xlDescending
        
    End With
        
End Sub
