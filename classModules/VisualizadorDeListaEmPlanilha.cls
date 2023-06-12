VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VisualizadorDeListaEmPlanilha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IVisualizadorDeLista

Dim plan As Excel.Worksheet
Dim conjuntoDeIntimacoes As ListaDeIntimacoes
    
Public Sub IVisualizadorDeLista_EntregarRespostaIntimacoes(grupoDeIntimacoes As ListaDeIntimacoes)
    Dim resultado As String
    Dim qtdDuplicatasExcluidas As Long
    
    Set conjuntoDeIntimacoes = grupoDeIntimacoes
    Set plan = PegarPlanilhaParaLancarListaDeIntimacoes()
    resultado = LancarGrupoDeIntimacoesNaPlanilha()
    qtdDuplicatasExcluidas = RemoverDuplicatasContarExcluidas
    resultado = FormatarStringLogResultado(resultado, qtdDuplicatasExcluidas)
    LancarLogNaPlanilha resultado
    
End Sub

Private Function PegarPlanilhaParaLancarListaDeIntimacoes() As Excel.Worksheet
    Dim nomeArquivo As String, nomePlanilha As String, nomeSistema As String
    Dim arq As Excel.Workbook
    Dim plan As Excel.Worksheet
    
    Select Case conjuntoDeIntimacoes.sistema
    Case SisifoEmbasaFuncoes.sfSistema.projudi
        nomeSistema = "Projudi"
    Case SisifoEmbasaFuncoes.sfSistema.pje1g
        nomeSistema = "PJe"
    Case SisifoEmbasaFuncoes.sfSistema.pje2g
        nomeSistema = "PJe"
    End Select
    
    nomeArquivo = "Sisifo - Lista intimacoes " & nomeSistema & " - " & Format(Date, "yyyy.mm.dd") & ".xlsx"
    nomePlanilha = "ListaIntim " & Format(Date, "yyyy.mm.dd")
    Set arq = PegarArquivoDentreOsAbertosNoExcel(nomeArquivo)
    If arq Is Nothing Then Set arq = PegarArquivoNaAreaDeTrabalho(nomeArquivo)
    If arq Is Nothing Then Set arq = CriarArquivoParaListaDeIntimacoesNovas(nomeArquivo)
    Set plan = PegarPlanilhaDeIntimacoesNovasNoArquivo(nomePlanilha, arq)
    If plan Is Nothing Then Set plan = CriarPlanilhaParaIntimacoesNovasNoArquivo(nomePlanilha, arq)
    Set PegarPlanilhaParaLancarListaDeIntimacoes = plan
End Function

Private Function PegarArquivoDentreOsAbertosNoExcel(nomeArquivo As String) As Excel.Workbook
    Dim i As Excel.Workbook
    
    For Each i In Excel.Workbooks
        If i.Name = nomeArquivo Then
            Set PegarArquivoDentreOsAbertosNoExcel = i
            Exit Function
        End If
    Next i
    Set PegarArquivoDentreOsAbertosNoExcel = Nothing
End Function

Private Function PegarArquivoNaAreaDeTrabalho(nomeArquivo As String) As Excel.Workbook
    Dim caminhoCompletoArquivo As String
    
    caminhoCompletoArquivo = SisifoEmbasaFuncoes.CaminhoDesktop & nomeArquivo
    If Dir(caminhoCompletoArquivo) = nomeArquivo Then
        Set PegarArquivoNaAreaDeTrabalho = Excel.Workbooks.Open(SisifoEmbasaFuncoes.CaminhoDesktop & nomeArquivo)
    Else
        Set PegarArquivoNaAreaDeTrabalho = Nothing
    End If
End Function

Private Function CriarArquivoParaListaDeIntimacoesNovas(nomeArquivo As String) As Excel.Workbook
    Dim arq As Excel.Workbook
    
    Set arq = Workbooks.Add
    Application.DisplayAlerts = False
    arq.SaveAs SisifoEmbasaFuncoes.CaminhoDesktop & nomeArquivo
    Application.DisplayAlerts = True
    Set CriarArquivoParaListaDeIntimacoesNovas = arq
End Function

Private Function PegarPlanilhaDeIntimacoesNovasNoArquivo(nomePlanilha As String, arquivo As Excel.Workbook) As Excel.Worksheet
    Dim plan As Excel.Worksheet
    
    For Each plan In arquivo.Sheets
        If plan.Name = nomePlanilha Then
            Set PegarPlanilhaDeIntimacoesNovasNoArquivo = plan
            Exit Function
        End If
    Next plan
End Function

Private Function CriarPlanilhaParaIntimacoesNovasNoArquivo(nomePlanilha As String, arquivo As Excel.Workbook) As Excel.Worksheet
    Dim plan As Excel.Worksheet
    
    Set plan = arquivo.Sheets(1)
    With plan
        .Name = nomePlanilha
        .Cells(1, 1).Formula = "Processo"
        .Columns(1).ColumnWidth = 25
        .Cells(1, 2).Formula = "Destinatário"
        .Columns(2).ColumnWidth = 20
        .Cells(1, 3).Formula = "Data de expedição"
        .Columns(3).ColumnWidth = 11
        .Cells(1, 4).Formula = "Origem da intimação/Modo de expedição"
        .Columns(4).ColumnWidth = 30
        .Cells(1, 5).Formula = "Log de intimações coletadas"
        .Cells().Range("$A$1:$E$1").Font.Bold = True
    End With
    Set CriarPlanilhaParaIntimacoesNovasNoArquivo = plan
End Function

Private Function LancarGrupoDeIntimacoesNaPlanilha() As String
    Dim celulaAlvo As Excel.Range
    Dim ultimaLinha As Long
    Dim i As Intimacao
    
    If conjuntoDeIntimacoes.intimacoes Is Nothing Then
        LancarGrupoDeIntimacoesNaPlanilha = "Não há intimações no período"
        Exit Function
    ElseIf conjuntoDeIntimacoes.intimacoes.Count = 0 Then
        LancarGrupoDeIntimacoesNaPlanilha = "Não há intimações no período"
        Exit Function
    ElseIf conjuntoDeIntimacoes.usuario.qtdIntimacoes = 1 And TypeOf conjuntoDeIntimacoes.intimacoes(1) Is SisifoEmbasaIntimacoes.UsuarioProjudiPje Then
        LancarGrupoDeIntimacoesNaPlanilha = CStr(conjuntoDeIntimacoes.intimacoes(1))
        Exit Function
    End If
    
    On Error Resume Next
    
    ultimaLinha = plan.Cells.SpecialCells(xlLastCell).Row
    Set celulaAlvo = plan.Cells(ultimaLinha + 1, 1)
    
    For Each i In conjuntoDeIntimacoes.intimacoes
        celulaAlvo.Formula = i.numeroProcesso
        celulaAlvo.Offset(0, 1).Formula = "'" & i.Destinatario
        celulaAlvo.Offset(0, 2).Formula = "'" & Format(i.dataExpedicao, "dd/mm/yyyy")
        If conjuntoDeIntimacoes.sistema = SisifoEmbasaFuncoes.sfSistema.projudi Then
            celulaAlvo.Offset(0, 3).Formula = conjuntoDeIntimacoes.usuario.nomeUsuario
        ElseIf conjuntoDeIntimacoes.sistema = SisifoEmbasaFuncoes.sfPJe1g Then
            celulaAlvo.Offset(0, 3).Formula = i.ModoExpedicao
        End If
        If i.SuspeitaDeNaoSerEmbasa = True Then
            celulaAlvo.Interior.ColorIndex = 44
            celulaAlvo.Offset(0, 1).Interior.ColorIndex = 44
            celulaAlvo.Offset(0, 2).Interior.ColorIndex = 44
            celulaAlvo.Offset(0, 3).Interior.ColorIndex = 44
        End If
        Set celulaAlvo = celulaAlvo.Offset(1, 0)
    Next i
    
    On Error GoTo 0
    
    If Err.Number = 0 Then
        LancarGrupoDeIntimacoesNaPlanilha = "sucesso"
    Else
        LancarGrupoDeIntimacoesNaPlanilha = Err.Description
    End If
    
End Function

Private Function RemoverDuplicatasContarExcluidas() As Long
    Dim qtdLinhasAntes As Long, qtdLinhasDepois As Long
    
    qtdLinhasAntes = plan.UsedRange.Rows.Count
    plan.Range("A:D").RemoveDuplicates 1, xlYes
    qtdLinhasDepois = plan.UsedRange.Rows.Count
    
    RemoverDuplicatasContarExcluidas = qtdLinhasAntes - qtdLinhasDepois
End Function

Private Function FormatarStringLogResultado(resultado As String, qtdDuplicatasExcluidas As Long) As String
    Dim log As String
    Dim qtdExclusivas As Long
    
    If resultado <> "sucesso" Then
        log = conjuntoDeIntimacoes.Origem & ": " & resultado & ";"
        
    Else
        log = conjuntoDeIntimacoes.Origem & ": " & conjuntoDeIntimacoes.usuario.qtdIntimacoes & " intimações capturadas"
        
        If qtdDuplicatasExcluidas = 0 Then
            log = log & ", nenhuma repetida"
        ElseIf qtdDuplicatasExcluidas > 0 Then
            qtdExclusivas = conjuntoDeIntimacoes.usuario.qtdIntimacoes - qtdDuplicatasExcluidas
            log = log & ", " & qtdExclusivas & " únicas e " & qtdDuplicatasExcluidas & " duplicadas e excluídas"
        End If
        
        If Not conjuntoDeIntimacoes.PeriodoDeExpedicao Is Nothing Then
            log = log & ", no período de " & conjuntoDeIntimacoes.PeriodoDeExpedicao.DataInicial & " até " & conjuntoDeIntimacoes.PeriodoDeExpedicao.Datafinal & ";"
        End If
        
        log = log & ";"
        
    End If
        
    FormatarStringLogResultado = log
    
End Function

Private Sub LancarLogNaPlanilha(log As String)
    Dim ultimaLinha As Long
    Dim arq As Excel.Workbook
    
    ultimaLinha = plan.Rows.Count
    ultimaLinha = plan.Cells(ultimaLinha, 5).End(xlUp).Row
    plan.Cells(ultimaLinha + 1, 5) = log
    Set arq = plan.Parent
    arq.Save
    
End Sub

