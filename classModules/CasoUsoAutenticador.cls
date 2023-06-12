VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CasoUsoAutenticador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim acDados As IAcessadorDeDados
Dim chControlador As IChromeAutenticador
Dim strPerfilLogado As String

Public Property Set Controlador(cont As IChromeAutenticador)
Set chControlador = cont
End Property

Public Property Get PerfilLogado() As String
PerfilLogado = strPerfilLogado
End Property

Private Sub Class_Initialize()
    Set acDados = New AcessadorDeDados
End Sub

Public Function CarregarSessaoValida(usuario As UsuarioProjudiPje) As ChromeDriver
    Dim resposta As ChromeDriver
    Dim mensagemErro As String
    Dim sessaoAtualEstaAtiva As Boolean, sessaoNaMemoriaEhDoUsuario As Boolean
    
    acDados.AlterarUsuarioAtual usuario.nomeUsuario
    CarregarPerfilLogado "sucesso", usuario.nomeUsuario
    
    If chControlador.SessaoNaMemoriaEstaAtiva(PegarInformacoesCookieSessao, strPerfilLogado) = True Then
        Set resposta = chControlador.navegador
    Else
        Set resposta = chControlador.FazerLogin(usuario)
    End If
    
    CarregarPerfilLogado "sucesso", usuario.nomeUsuario
    'mensagemErro = LidarComErroLogin(strPerfilLogado)
    'If mensagemErro = "" Then
        If sessaoAtualEstaAtiva = False Then
            acDados.SalvarCookieSessao chControlador.Cookies, usuario.nomeUsuario
        End If
    'Else
    '    resultado = mensagemErro
    'End If
    
    resposta.Window.Maximize
    Set CarregarSessaoValida = resposta

End Function

Private Sub CarregarPerfilLogado(resultadoLogin As String, nomeUsuario As String)
    If resultadoLogin = "sucesso" Then
        strPerfilLogado = DescobrirPerfilLogadoPeloNome(nomeUsuario)
    Else
        strPerfilLogado = resultadoLogin
    End If
End Sub

Private Function DescobrirPerfilLogadoPeloNome(nomeUsuario As String) As String
    If InStr(1, nomeUsuario, "Representante") <> 0 Then
        DescobrirPerfilLogadoPeloNome = "Representante"
    Else
        DescobrirPerfilLogadoPeloNome = "Advogado"
    End If
End Function

Private Function LidarComErroLogin(PerfilLogado) As String
    Dim resposta As String
    
    If LCase(PerfilLogado) = "usuário ou senha inválidos." Then
        resposta = "Erro: " & PerfilLogado
    ElseIf LCase(PerfilLogado) <> "advogado" And LCase(PerfilLogado) <> "representante" Then
        resposta = "Erro: " & PerfilLogado
    End If
    
    LidarComErroLogin = resposta
End Function

Public Function PegarInformacoesCookieSessao() As Selenium.Dictionary
    Set PegarInformacoesCookieSessao = acDados.PegarInformacoesCookieSessao
End Function

Public Sub LimparCookieSessao()
    acDados.LimparCookieSessao
End Sub

Public Function PegarObjetoUsuarioAtual(sistema As sfSistema) As UsuarioProjudiPje
    Dim usuario As New UsuarioProjudiPje
    Dim nomeUsuario As String
    
    nomeUsuario = acDados.PegarNomeUsuarioAtual
    If nomeUsuario <> "" Then
        Set PegarObjetoUsuarioAtual = acDados.PegarUsuarioPeloNome(nomeUsuario, sistema)
    Else
        Set PegarObjetoUsuarioAtual = Nothing
    End If
End Function

Public Function PegarUsuarioPeloNome(nomeUsuario As String, sistema As SisifoEmbasaFuncoes.sfSistema) As UsuarioProjudiPje
    Set PegarUsuarioPeloNome = acDados.PegarUsuarioPeloNome(nomeUsuario, sistema)
End Function
