Attribute VB_Name = "sfIntRegNegPrincipal"
Option Explicit

Sub PegarListaIntimacoes(ByVal Controle As IRibbonControl)
    Dim casoUso As ICasoUsoListarIntimacoes
    Dim acDados As IAcessadorDeDados
    Dim sistemaRibbon As SisifoEmbasaFuncoes.sfSistema
    Dim tribunalRibbon As SisifoEmbasaFuncoes.sfTribunal
    
    Set acDados = New AcessadorDeDados
    sistemaRibbon = acDados.PegarSistemaSelecionado
    tribunalRibbon = acDados.PegarTribunalSelecionado
    
    Select Case sistemaRibbon
    Case SisifoEmbasaFuncoes.sfSistema.projudi
        Set casoUso = New CasoUsoListarIntimProjudi
        casoUso.PegarListaIntimacoesPorData
        
    Case SisifoEmbasaFuncoes.sfSistema.pje1g
        Select Case tribunalRibbon
        Case SisifoEmbasaFuncoes.sfTribunal.Tjba
            Set casoUso = New CasoUsoListarIntimPje1gTjba
            casoUso.PegarListaIntimacoesPorData
            
        Case SisifoEmbasaFuncoes.sfTribunal.trt5
            GoTo NaoFaz
        End Select
        
'    Case sisifoembasafuncoes.sfSistema.PJe2g
'        Select Case tribunalRibbon
'        Case sisifoembasafuncoes.sfTribunal.TJBA
'            PegarListaIntimacoesPjeTjba sisifoembasafuncoes.sfPJe1g
'        Case sisifoembasafuncoes.sfTribunal.TRT5
'            GoTo NaoFaz
'        End Select
        
    Case Else
NaoFaz:
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", eu ainda não sei buscar processos no sistema e tribunal " & _
        "selecionados.", vbExclamation + vbOKOnly, "Sísifo - Sistema ainda não abrangido"
        
    End Select
End Sub

Sub CadastrarAndamentoIndividual(ByVal Controle As IRibbonControl)
    Dim casoUso As ICasoUsoCadastrarIntimacoes
    Dim acDados As IAcessadorDeDados
    Dim sistemaRibbon As SisifoEmbasaFuncoes.sfSistema
    Dim tribunalRibbon As SisifoEmbasaFuncoes.sfTribunal
    
    Set acDados = New AcessadorDeDados
    sistemaRibbon = acDados.PegarSistemaSelecionado
    tribunalRibbon = acDados.PegarTribunalSelecionado

    Select Case sistemaRibbon
    Case SisifoEmbasaFuncoes.sfSistema.projudi
        Set casoUso = New CasoUsoCadastrarIntimProjudi
        casoUso.CadastrarAndamentoIndividual
        
    Case SisifoEmbasaFuncoes.sfSistema.pje1g
        Select Case tribunalRibbon
        Case SisifoEmbasaFuncoes.sfTribunal.Tjba
            Set casoUso = New CasoUsoCadastrarIntimPje1gTjba
            casoUso.CadastrarAndamentoIndividual
            
        Case SisifoEmbasaFuncoes.sfTribunal.trt5
            GoTo NaoFaz
        End Select
        
'    Case sisifoembasafuncoes.sfSistema.PJe2g
'        Select Case tribunalRibbon
'        Case sisifoembasafuncoes.sfTribunal.TJBA
'            PegarListaIntimacoesPjeTjba sisifoembasafuncoes.sfPJe1g
'        Case sisifoembasafuncoes.sfTribunal.TRT5
'            GoTo NaoFaz
'        End Select
        
    Case Else
NaoFaz:
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", eu ainda não sei buscar processos no sistema e tribunal " & _
        "selecionados.", vbExclamation + vbOKOnly, "Sísifo - Sistema ainda não abrangido"
        
    End Select


End Sub

Sub AlterarUsuarioPerfil(ByVal Controle As IRibbonControl)
    Dim acDados As IAcessadorDeDados
    Dim sistemaSelecionado As sfSistema
    Dim tribunalSelecionado As sfTribunal
    
    Set acDados = New AcessadorDeDados
    sistemaSelecionado = acDados.PegarSistemaSelecionado
    tribunalSelecionado = acDados.PegarTribunalSelecionado
    
    Select Case sistemaSelecionado
    Case sfSistema.projudi
        AlterarUsuarioPerfilProjudi
    
    Case sfSistema.pje1g
        Select Case tribunalSelecionado
        Case SisifoEmbasaFuncoes.sfTribunal.Tjba
            
        Case SisifoEmbasaFuncoes.sfTribunal.trt5
            GoTo NaoFaz
        End Select
        
'    Case sisifoembasafuncoes.sfSistema.PJe2g
'        Select Case tribunalRibbon
'        Case sisifoembasafuncoes.sfTribunal.TJBA
'            PegarListaIntimacoesPjeTjba sisifoembasafuncoes.sfPJe1g
'        Case sisifoembasafuncoes.sfTribunal.TRT5
'            GoTo NaoFaz
'        End Select
        
    Case Else
NaoFaz:
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", eu ainda não conheço usuários do sistema e tribunal " & _
        "selecionados.", vbExclamation + vbOKOnly, "Sísifo - Sistema e/ou tribunal ainda não abrangido"
        
    End Select
    
End Sub

Sub AlterarUsuarioPerfilProjudi()
    Dim gestorForm As New GestorFormSelecionarUsuarios
    Dim acDados As IAcessadorDeDados
    Dim nomeUsuario As String
    
    Set acDados = New AcessadorDeDados
    nomeUsuario = acDados.PegarNomeUsuarioAtual
    nomeUsuario = gestorForm.SelecionarNomeDoNovoUsuario(nomeUsuario)
    acDados.AlterarUsuarioAtual nomeUsuario
    RestringirEdicaoRibbon ThisWorkbook, cfIntConfigurações
End Sub
