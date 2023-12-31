VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CasoUsoListarIntimPje1gTjba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICasoUsoListarIntimacoes

Dim chromeIntimacoesLista As IControladorListadeIntimacoes
Dim acDados As IAcessadorDeDados

Private Sub Class_Initialize()
    Set chromeIntimacoesLista = New ChromePjeTjba1gIntimacoesLista
    Set acDados = New AcessadorDeDados
End Sub

Public Sub ICasoUsoListarIntimacoes_PegarListaIntimacoesPorData()
    Dim ui As IControladorEntradaLista
    Dim devolveDados As IApresentadorDeLista
    Dim Autenticador As New CasoUsoAutenticador
    Dim usuario As UsuarioProjudiPje
    Dim grupoDeIntimacoes As ListaDeIntimacoes
    Dim comarcasDisponiveis As Collection, comarcasSelecionadas As Collection, intimacoes As Collection
    Dim resultado As String
    Dim i As Variant
    Dim usuarioConfirmouProsseguimentoNoForm As Boolean
    
    Set ui = New ControladorEntradaLista
    Set usuario = acDados.PegarUsuarioPeloNome("Pje1g Representante Gestor Embasa", SisifoEmbasaFuncoes.sfSistema.pje1g)
    If usuario.Login = "" Or usuario.senha = "" Then
        MsgBox SisifoEmbasaFuncoes.determinartratamento & ", eu n�o sei o login e/ou a senha do usu�rio """ & _
            usuario.nomeUsuario & """ para o sistema PJe1g do TJ/BA. Este usu�rio ser� desconsiderado.", vbCritical + vbOKOnly, _
            "S�sifo - Usu�rio desconhecido"
    Else
        Set Autenticador.Controlador = New ChromePjeTjba1gAutenticacao
        Set chromeIntimacoesLista.ChromeIntimacoes.ChromePje1g.Chrome.Chrome = Autenticador.CarregarSessaoValida(usuario)
        If chromeIntimacoesLista.ChromeIntimacoes.ChromePje1g.Chrome.Chrome Is Nothing Then Exit Sub
        
        Set devolveDados = New ApresentadorDeLista
        Set comarcasDisponiveis = chromeIntimacoesLista.PegarComarcasDisponiveis("Pendentes de ci�ncia ou de resposta")
        usuarioConfirmouProsseguimentoNoForm = ui.MostrarFormConfirmandoSelecaoComarcas(comarcasDisponiveis)
        If usuarioConfirmouProsseguimentoNoForm = False Then Exit Sub
        Set comarcasSelecionadas = ui.NomesComarcasSelecionadas
        
        For Each i In comarcasSelecionadas
            Set intimacoes = chromeIntimacoesLista.PegarListaIntimacoes(Comarca:=CStr(i))
            'acDados.RegistrarDataMaisRecenteDeUsuario Date, usuario.nomeUsuario, sisifoembasafuncoes.sfSistema.Projudi
            usuario.qtdIntimacoes = IIf(intimacoes Is Nothing, 0, intimacoes.Count)
            Set grupoDeIntimacoes = New ListaDeIntimacoes
            With grupoDeIntimacoes
                Set .usuario = usuario
                .Comarca = CStr(i)
                Set .intimacoes = intimacoes
                'Set .PeriodoDeExpedicao = periodo
                .instancia = PrimeiroGrau
                .sistema = pje1g
            End With
            devolveDados.PassarListaIntimacoesParaView grupoDeIntimacoes
        Next i
    End If
End Sub
