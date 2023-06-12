VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CasoUsoListarIntimProjudi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICasoUsoListarIntimacoes

Dim chromeIntimacoesLista As IControladorListadeIntimacoes
Dim acDados As IAcessadorDeDados

Private Sub Class_Initialize()
    Set chromeIntimacoesLista = New ChromeProjudiIntimacoesLista
    Set acDados = New AcessadorDeDados
End Sub

Public Sub ICasoUsoListarIntimacoes_PegarListaIntimacoesPorData()
    Dim ui As IControladorEntradaLista
    Dim devolveDados As IApresentadorDeLista
    Dim Autenticador As New CasoUsoAutenticador
    Dim periodo As PeriodoDeTempo
    Dim usuario As UsuarioProjudiPje
    Dim grupoDeIntimacoes As ListaDeIntimacoes
    Dim usuariosSelecionados As Collection, intimacoes As Collection
    Dim usuariosDisponiveis() As String, resultado As String
    Dim i As Variant
    Dim usuarioConfirmouProsseguimentoNoForm As Boolean
    
    Set ui = New ControladorEntradaLista
    usuariosDisponiveis = acDados.PegarUsuariosDisponiveisParaLogin(cfPerfisProjudi.Range("PerfisProjudi"))
    usuarioConfirmouProsseguimentoNoForm = ui.MostrarFormConfirmandoSelecaoUsuarios(usuariosDisponiveis)
    If usuarioConfirmouProsseguimentoNoForm = False Then Exit Sub
    Set periodo = ui.PeriodoSelecionado
    Set usuariosSelecionados = ui.NomesUsuariosSelecionados
    Set devolveDados = New ApresentadorDeLista
    Set Autenticador.Controlador = New ChromeProjudiAutenticacao
    
    For Each i In usuariosSelecionados
        Set usuario = acDados.PegarUsuarioPeloNome(CStr(i), SisifoEmbasaFuncoes.sfSistema.projudi)
        If usuario.Login = "" Or usuario.senha = "" Then
            MsgBox SisifoEmbasaFuncoes.determinartratamento & ", eu não sei o login e/ou a senha do usuário """ & _
                usuario.nomeUsuario & """ para o sistema Projudi. Este usuário será desconsiderado.", vbCritical + vbOKOnly, _
                "Sísifo - Usuário desconhecido"
        Else
            Set chromeIntimacoesLista.ChromeIntimacoes.ChromeProjudi.Chrome.Chrome = Autenticador.CarregarSessaoValida(usuario)
            If chromeIntimacoesLista.ChromeIntimacoes.ChromeProjudi.Chrome.Chrome Is Nothing Then Exit Sub
            Set intimacoes = chromeIntimacoesLista.PegarListaIntimacoes(periodo)
            usuario.qtdIntimacoes = IIf(intimacoes Is Nothing, 0, intimacoes.Count)
            Set grupoDeIntimacoes = New ListaDeIntimacoes
            With grupoDeIntimacoes
                Set .usuario = usuario
                Set .intimacoes = intimacoes
                Set .PeriodoDeExpedicao = periodo
                '.Instancia =
                .sistema = projudi
            End With
            acDados.RegistrarDataMaisRecenteDeUsuario periodo.Datafinal, usuario.nomeUsuario, SisifoEmbasaFuncoes.sfSistema.projudi
            devolveDados.PassarListaIntimacoesParaView grupoDeIntimacoes
        End If
    Next i
End Sub
