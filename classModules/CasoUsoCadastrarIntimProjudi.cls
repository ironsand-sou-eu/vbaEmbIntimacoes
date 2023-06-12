VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CasoUsoCadastrarIntimProjudi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICasoUsoCadastrarIntimacoes

Private oAndamento As andamento

Public Sub ICasoUsoCadastrarIntimacoes_CadastrarAndamentoIndividual()
    Dim iu As IControladorEntradaCadastro
    Dim acDados As IAcessadorDeDados
    Dim formInput As frmAndamento
    Dim coletadorIntimacoes As IControladorCadastrarIntimacoes
    Dim infoJuizo As StructJuizo, infoListasInterface As StructInformacoesParaInterface, informacoesFinaisAndamento As StructInformacoesAndamento
    Dim usuarioConfirmouProsseguimentoNoForm As Boolean
    
    Set oAndamento = New andamento
    Set iu = New ControladorEntradaCadastro
    Set acDados = New AcessadorDeDados
    Set oAndamento.ObjetoTokenDoNumeroProcesso = iu.PegarObjetoComNumeroProcesso
    oAndamento.numeroProcesso = iu.PegarNumeroProcessoDoObjetoToken(oAndamento.ObjetoTokenDoNumeroProcesso)
    
    If oAndamento.numeroProcesso = "Número não é CNJ" Then
        If MsgBox(SisifoEmbasaFuncoes.determinartratamento & ", o número de processo """ & _
        oAndamento.numeroProcesso & """ não parece ser um número no padrão CNJ. " & _
        "Deseja continuar assim mesmo?", vbQuestion + vbYesNo, "Sísifo - Número não CNJ") = vbNo Then Exit Sub
    End If
    
    oAndamento.tribunal = acDados.PegarTribunalSelecionado
    oAndamento.sistema = acDados.PegarSistemaSelecionado
    oAndamento.InferirECarregarInstancia
    If oAndamento.PerguntarSeProssegueAposInconsistencias = vbNo Then Exit Sub
    
    Set coletadorIntimacoes = New ChromeProjudiIntimacoesColeta
    coletadorIntimacoes.CarregarInformacoesWeb oAndamento
    If Not (oAndamento.mensagemErro = "sucesso" Or oAndamento.mensagemErro = "") Then Exit Sub
    
    oAndamento.nomeAndamento = acDados.PegarNomeEspaiderCorrespondenteAoNomeAndamentoDoSistemaJudicial(oAndamento.nomeAndamento)
    infoJuizo = acDados.PegarInformacoesEspaiderCorrespondentesAoNomeJuizoDoSistemaJudicial(oAndamento.NovoJuizo)
    oAndamento.NovoJuizo = infoJuizo.nomeJuizo
    
    With infoListasInterface
        .ListaDeAndamentos = acDados.PegarListaDeAndamentosEspaiderPossiveis
        .ListaDeProvidencias = acDados.PegarListaDeProvidenciasEspaiderPossiveis
        .ListaDeJuizos = acDados.PegarListaDeJuizosEspaiderPossiveis
        .NomeProvidenciaRelacionada = acDados.PegarNomeProvidenciaALancarNoAndamento(oAndamento.nomeAndamento)
    End With
    
    informacoesFinaisAndamento = iu.ConfirmarEPegarInformacoesFinais(oAndamento, infoListasInterface)
    If informacoesFinaisAndamento.DeveGerar = False Then Exit Sub
    If oAndamento.NovoJuizo <> infoJuizo.nomeJuizo Then infoJuizo = acDados.PegarInformacoesEspaiderCorrespondentesAoNomeJuizoDoSistemaJudicial(oAndamento.NovoJuizo)
    MontarAndamento informacoesFinaisAndamento, infoJuizo
    'If oAndamento.LidarComErro <> "" Then Exit Sub
    
    acDados.SalvarInformacoes oAndamento
    iu.MarcarAndamentoCadastrado oAndamento.ObjetoTokenDoNumeroProcesso
    MsgBox "Mestre, o Andamento foi registrado com sucesso!" & vbCrLf & _
        "Processo: " & oAndamento.numeroProcesso & vbCrLf & _
        "Andamento: " & oAndamento.nomeAndamento & " em " & oAndamento.DataAndamento & vbCrLf & _
        "Lembre-se de inserir no Espaider no final do dia.", vbInformation + vbOKOnly, "Sísifo - Processo salvo para inclusão"
    
    Set oAndamento = Nothing
End Sub

Private Sub MontarAndamento(informacoes As StructInformacoesAndamento, infoJuizo As StructJuizo)
    With oAndamento
        .nomeAndamento = informacoes.nomeAndamento
        .IdConversao = SisifoEmbasaFuncoes.PegarSequencialAndamento(oAndamento.numeroProcesso, informacoes.DataAndamento)
        .DataAndamento = informacoes.DataAndamento
        .obsAndamento = informacoes.obsAndamento
        .NovoJuizo = infoJuizo.nomeJuizo
        .NovoOrgao = infoJuizo.NomeOrgao
        .NovaComarca = infoJuizo.nomeComarca
        If informacoes.nomeProvidencia <> "" Then CarregarObjetoProvidencia (informacoes.nomeProvidencia)
    End With
End Sub

Private Sub CarregarObjetoProvidencia(nomeProvidencia As String)
    Dim acDados As IAcessadorDeDados
    Dim providencia As providencia
    Dim infoProvidencia As StructProvidencia
    
    Set providencia = SisifoEmbasaFuncoes.New_Providencia
    Set acDados = New AcessadorDeDados
    infoProvidencia = acDados.PegarParametrosProvidenciaPeloNome(nomeProvidencia)
    
    With providencia
        .Nome = nomeProvidencia
        .numeroProcesso = oAndamento.numeroProcesso
        .Situacao = SisifoEmbasaFuncoes.sfstatusprovidencia.aCumprir
        .Datafinal = acDados.PegarDataFinalPadrao
        .Nucleo = infoProvidencia.Nucleo
        .Responsavel = infoProvidencia.Responsavel
        .Observacao = infoProvidencia.Observacao
        
        If .Nome = "Contestar - Remarcação de audiência" Then .Observacao = oAndamento.nomeAndamento & " - " & oAndamento.DataAndamento & vbCrLf & _
            "Sísifo Intimações - remarcação." & vbCr & _
            "1) Verificar se já tem defesa." & vbCr & _
            "2) Caso não tenha, distribuir ao responsável." & vbCr

        .Alertar = False
        .DataAlerta = 0
        .DiasAntecedenciaAlerta = 0
        .PeriodicidadeAlerta = 0
        .PrazoDias = 0
        .GerarAndamento = infoProvidencia.GerarAndamento
        .TipoDataBaseAndamento = infoProvidencia.TipoDataBaseAndamento
        .AndamentoAGerar = infoProvidencia.andamento
        .IdConversao = SisifoEmbasaFuncoes.PegarSequencialProvidencia(providencia, 1)
        .IdAndamento = oAndamento.IdConversao
    End With
    
    oAndamento.Providencias.Add providencia
End Sub
