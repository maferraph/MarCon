VERSION 5.00
Begin VB.UserControl MarcadorPecas 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   ControlContainer=   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   4110
   Begin VB.CommandButton BT_Posicao 
      Caption         =   "Configuração de Posição"
      Height          =   1335
      Left            =   2760
      Picture         =   "MarcadorPecas.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BT_MapaCaracteres 
      Caption         =   "Mapa de Caracteres"
      Height          =   1335
      Left            =   1440
      Picture         =   "MarcadorPecas.ctx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BT_Emergencia 
      Caption         =   "EMERGÊNCIA"
      Height          =   1335
      Left            =   120
      Picture         =   "MarcadorPecas.ctx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "MarcadorPecas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub BT_Emergencia_Click()
     'MarcarChap01 "CORPO", "PREME", "CUNHA", "ANEIS", "HASTE", "BUCHA", "GAXETA", "JUNTA", "PPCORPO", "PPPREME", "BITOLA", "CLASSE", "EXTREMIDADE", "OM", "DATA", "CAPACIDADE"
     MarcarChap02 "CORPO", "PREME", "CUNHA", "ANEIS", "HASTE", "BUCHA", "GAXETA", "JUNTA", "PPCORPO", "PPPREME", "BITOLA", "CLASSE", "EXTREMIDADE", "OM", "DATA", "CAPACIDADE"
     'MarcarChap03 "CORPO", "PREME", "CONTRASEDE", "SEDE", "HASTE", "BUCHA", "GAXETA", "JUNTA", "PPCORPO", "PPPREME", "BITOLA", "CLASSE", "EXTREMIDADE", "OM", "DATA", "CAPACIDADE"
     'MarcarChap04 "CORPO", "PREME", "CONTRASEDE", "SEDE", "HASTE", "BUCHA", "GAXETA", "JUNTA", "PPCORPO", "PPPREME", "BITOLA", "CLASSE", "EXTREMIDADE", "OM", "DATA", "CAPACIDADE"
     'MarcarChap05 "CORPO", "PISTAO", "SEDE", "MOLA", "PPCORPO", "JUNTA", "BITOLA", "CLASSE", "EXTREMIDADE", "OM", "DATA", "CAPACIDADE"
     'MarcarChap06 "CORPO", "PENDULO", "ANEL", "DISCO", "PPCORPO", "JUNTA", "BITOLA", "CLASSE", "EXTREMIDADE", "OM", "DATA", "CAPACIDADE"
     'Tela_MarcadorPecasPlano.Show
End Sub
Private Sub BT_MapaCaracteres_Click()
    Tela_Caracter_7x7.Show vbModal
End Sub
Private Sub BT_Posicao_Click()
    Tela_Posicao.Show vbModal
End Sub
Private Sub UserControl_Initialize()
    'arquivos de configuraçào geral
    'SVAR_CAMINHO_ARQUIVOS = App.Path
    SVAR_CAMINHO_ARQUIVOS = "\\Servidor1\intranet\intranet\maquinas\marcon"
    '************ TIMERS ****************
    SetaTimers
    '************ DADOS SOBRE A PORTA PARALELA ****************
    ZeraPinos
    'escreve valor inicial de byte nas portas
    ZeraPortas
    'zera passos
    IVAR_PASSO_X = 0
    IVAR_PASSO_Y = 0
    'carrega mapa de caracteres
    CarregaVetorCaracter
End Sub


'***************************************************************************
'              FUNCOES ESPECIFICAS PARA O MARCADOR DE CHAPINHAS
'***************************************************************************
Private Sub CarregaVetorCaracter()
    Dim SVAR_CARACTERES, VALOR As String
    Dim I, J, K As Integer
    SVAR_ARQUIVO = SVAR_CAMINHO_ARQUIVOS & "\7x7.car"
    VETOR_CARACTER = Array()
    SVAR_CARACTERES = "ABCDEFGHIJKLMNOPQRSTUVWYXZ0123456789.,;:/\()-*+@#" & Chr(34)
    For I = 1 To Len(SVAR_CARACTERES)
        VALOR = ""
        For J = 1 To 7 'colunas dos 7 caracteres
            For K = 1 To 7 'linhas dos 7 caracteres
                VALOR = VALOR & LeINI(SVAR_ARQUIVO, Mid(SVAR_CARACTERES, I, 1), "L" & J & "C" & K)
            Next K
        Next J
        ReDim Preserve VETOR_CARACTER(UBound(VETOR_CARACTER) + 1)
        VETOR_CARACTER(UBound(VETOR_CARACTER)) = Array(Mid(SVAR_CARACTERES, I, 1), VALOR)
    Next I
End Sub

'Marcador de Gaveta de 1/2, 3/4 e 1 - Modelo 01
Public Sub MarcarChap01(CORPO As String, PREME As String, CUNHA As String, ANEIS As String, HASTE As String, BUCHA As String, GAXETA As String, JUNTA As String, PPCORPO As String, PPPREME As String, BITOLA As String, CLASSE As String, EXTREMIDADE As String, OM As String, DATA As String, CAPACIDADE As String)
    'muda mouse
    Screen.MousePointer = vbHourglass
    'posiciona maquina no ZERO-MAQUINHA(HOME)
    'PosicionamentoPlano_Home
    'nome do arquivo
    SVAR_ARQUIVO_POSICAO = SVAR_CAMINHO_ARQUIVOS & "\chap01.pos"
    SVAR_MARCACAO_ATUAL = "CHAP01"
    MarcadorChapinha Array(Array(CORPO, LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "Y")), _
                            Array(PREME, LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "Y")), _
                            Array(ANEIS, LeINI(SVAR_ARQUIVO_POSICAO, "ANEIS", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "ANEIS", "Y")), _
                            Array(CUNHA, LeINI(SVAR_ARQUIVO_POSICAO, "CUNHA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CUNHA", "Y")), _
                            Array(HASTE, LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "Y")), _
                            Array(BUCHA, LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "Y")), _
                            Array(JUNTA, LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "Y")), _
                            Array(GAXETA, LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "Y")), _
                            Array(PPCORPO, LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y")), _
                            Array(PPPREME, LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "Y")), _
                            Array(CLASSE, LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "Y")), _
                            Array(BITOLA, LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "Y")), _
                            Array(EXTREMIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y")), _
                            Array(OM, LeINI(SVAR_ARQUIVO_POSICAO, "OM", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "OM", "Y")), _
                            Array(CAPACIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y")), _
                            Array(DATA, LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "Y")) _
                            )
    Screen.MousePointer = vbDefault
    'exibe tela
    Tela_Chapinha_D43.Show vbModal
End Sub
'Marcador de Gaveta de 1.1/2 e 2 - Modelo 02
Public Sub MarcarChap02(CORPO As String, PREME As String, CUNHA As String, ANEIS As String, HASTE As String, BUCHA As String, GAXETA As String, JUNTA As String, PPCORPO As String, PPPREME As String, BITOLA As String, CLASSE As String, EXTREMIDADE As String, OM As String, DATA As String, CAPACIDADE As String)
    'muda mouse
    Screen.MousePointer = vbHourglass
    'posiciona maquina no ZERO-MAQUINHA(HOME)
    'PosicionamentoPlano_Home
    'nome do arquivo
    SVAR_ARQUIVO_POSICAO = SVAR_CAMINHO_ARQUIVOS & "\chap02.pos"
    SVAR_MARCACAO_ATUAL = "CHAP02"
    MarcadorChapinha Array(Array(CORPO, LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "Y")), _
                            Array(PREME, LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "Y")), _
                            Array(ANEIS, LeINI(SVAR_ARQUIVO_POSICAO, "ANEIS", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "ANEIS", "Y")), _
                            Array(CUNHA, LeINI(SVAR_ARQUIVO_POSICAO, "CUNHA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CUNHA", "Y")), _
                            Array(HASTE, LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "Y")), _
                            Array(BUCHA, LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "Y")), _
                            Array(JUNTA, LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "Y")), _
                            Array(GAXETA, LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "Y")), _
                            Array(PPCORPO, LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y")), _
                            Array(PPPREME, LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "Y")), _
                            Array(CLASSE, LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "Y")), _
                            Array(BITOLA, LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "Y")), _
                            Array(EXTREMIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y")), _
                            Array(OM, LeINI(SVAR_ARQUIVO_POSICAO, "OM", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "OM", "Y")), _
                            Array(CAPACIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y")), _
                            Array(DATA, LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "Y")) _
                            )
    Screen.MousePointer = vbDefault
    'exibe tela
    Tela_Chapinha_D57.Show vbModal
End Sub
'Marcador de Globo de 1/2, 3/4 e 1 - Modelo 03
Public Sub MarcarChap03(CORPO As String, PREME As String, CONTRASEDE As String, SEDE As String, HASTE As String, BUCHA As String, GAXETA As String, JUNTA As String, PPCORPO As String, PPPREME As String, BITOLA As String, CLASSE As String, EXTREMIDADE As String, OM As String, DATA As String, CAPACIDADE As String)
    'muda mouse
    Screen.MousePointer = vbHourglass
    'posiciona maquina no ZERO-MAQUINHA(HOME)
    'PosicionamentoPlano_Home
    'nome do arquivo
    SVAR_ARQUIVO_POSICAO = SVAR_CAMINHO_ARQUIVOS & "\chap03.pos"
    SVAR_MARCACAO_ATUAL = "CHAP03"
    MarcadorChapinha Array(Array(CORPO, LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "Y")), _
                            Array(PREME, LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "Y")), _
                            Array(SEDE, LeINI(SVAR_ARQUIVO_POSICAO, "SEDE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "SEDE", "Y")), _
                            Array(CONTRASEDE, LeINI(SVAR_ARQUIVO_POSICAO, "CONTRASEDE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CONTRASEDE", "Y")), _
                            Array(HASTE, LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "Y")), _
                            Array(BUCHA, LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "Y")), _
                            Array(JUNTA, LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "Y")), _
                            Array(GAXETA, LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "Y")), _
                            Array(PPCORPO, LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y")), _
                            Array(PPPREME, LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "Y")), _
                            Array(CLASSE, LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "Y")), _
                            Array(BITOLA, LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "Y")), _
                            Array(EXTREMIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y")), _
                            Array(OM, LeINI(SVAR_ARQUIVO_POSICAO, "OM", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "OM", "Y")), _
                            Array(CAPACIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y")), _
                            Array(DATA, LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "Y")) _
                            )
    Screen.MousePointer = vbDefault
    'exibe tela
    Tela_Chapinha_D43.Show vbModal
End Sub
'Marcador de Globo de 1.1/2 e 2 - Modelo 04
Public Sub MarcarChap04(CORPO As String, PREME As String, CONTRASEDE As String, SEDE As String, HASTE As String, BUCHA As String, GAXETA As String, JUNTA As String, PPCORPO As String, PPPREME As String, BITOLA As String, CLASSE As String, EXTREMIDADE As String, OM As String, DATA As String, CAPACIDADE As String)
    'muda mouse
    Screen.MousePointer = vbHourglass
    'posiciona maquina no ZERO-MAQUINHA(HOME)
    'PosicionamentoPlano_Home
    'nome do arquivo
    SVAR_ARQUIVO_POSICAO = SVAR_CAMINHO_ARQUIVOS & "\chap04.pos"
    SVAR_MARCACAO_ATUAL = "CHAP04"
    MarcadorChapinha Array(Array(CORPO, LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "Y")), _
                            Array(PREME, LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "Y")), _
                            Array(SEDE, LeINI(SVAR_ARQUIVO_POSICAO, "SEDE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "SEDE", "Y")), _
                            Array(CONTRASEDE, LeINI(SVAR_ARQUIVO_POSICAO, "CONTRASEDE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CONTRASEDE", "Y")), _
                            Array(HASTE, LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "Y")), _
                            Array(BUCHA, LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "Y")), _
                            Array(JUNTA, LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "Y")), _
                            Array(GAXETA, LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "Y")), _
                            Array(PPCORPO, LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y")), _
                            Array(PPPREME, LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "Y")), _
                            Array(CLASSE, LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "Y")), _
                            Array(BITOLA, LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "Y")), _
                            Array(EXTREMIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y")), _
                            Array(OM, LeINI(SVAR_ARQUIVO_POSICAO, "OM", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "OM", "Y")), _
                            Array(CAPACIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y")), _
                            Array(DATA, LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "Y")) _
                            )
    Screen.MousePointer = vbDefault
    'exibe tela
    Tela_Chapinha_D57.Show vbModal
End Sub
'Marcador de Retenção Pistão (todas) - Modelo 05
Public Sub MarcarChap05(CORPO As String, PISTAO As String, SEDE As String, MOLA As String, PPCORPO As String, JUNTA As String, BITOLA As String, CLASSE As String, EXTREMIDADE As String, OM As String, DATA As String, CAPACIDADE As String)
    'muda mouse
    Screen.MousePointer = vbHourglass
    'posiciona maquina no ZERO-MAQUINHA(HOME)
    'PosicionamentoPlano_Home
    'nome do arquivo
    SVAR_ARQUIVO_POSICAO = SVAR_CAMINHO_ARQUIVOS & "\chap05.pos"
    SVAR_MARCACAO_ATUAL = "CHAP05"
    MarcadorChapinha Array(Array(CORPO, LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/TAMPA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/TAMPA", "Y")), _
                            Array(PISTAO, LeINI(SVAR_ARQUIVO_POSICAO, "PISTAO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PISTAO", "Y")), _
                            Array(MOLA, LeINI(SVAR_ARQUIVO_POSICAO, "MOLA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "MOLA", "Y")), _
                            Array(SEDE, LeINI(SVAR_ARQUIVO_POSICAO, "SEDE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "SEDE", "Y")), _
                            Array(PPCORPO, LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y")), _
                            Array(JUNTA, LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "Y")), _
                            Array(CLASSE, LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "Y")), _
                            Array(BITOLA, LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "Y")), _
                            Array(EXTREMIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y")), _
                            Array(OM, LeINI(SVAR_ARQUIVO_POSICAO, "OM", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "OM", "Y")), _
                            Array(CAPACIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y")), _
                            Array(DATA, LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "Y")) _
                            )
    Screen.MousePointer = vbDefault
    'exibe tela
    Tela_Chapinha_D43.Show vbModal
End Sub
'Marcador de Retenção Portinhola (todas) - Modelo 06
Public Sub MarcarChap06(CORPO As String, PENDULO As String, ANEL As String, DISCO As String, PPCORPO As String, JUNTA As String, BITOLA As String, CLASSE As String, EXTREMIDADE As String, OM As String, DATA As String, CAPACIDADE As String)
    'muda mouse
    Screen.MousePointer = vbHourglass
    'posiciona maquina no ZERO-MAQUINHA(HOME)
    'PosicionamentoPlano_Home
    'nome do arquivo
    SVAR_ARQUIVO_POSICAO = SVAR_CAMINHO_ARQUIVOS & "\chap06.pos"
    SVAR_MARCACAO_ATUAL = "CHAP06"
    MarcadorChapinha Array(Array(CORPO, LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/TAMPA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/TAMPA", "Y")), _
                            Array(PENDULO, LeINI(SVAR_ARQUIVO_POSICAO, "PENDULO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PENDULO", "Y")), _
                            Array(DISCO, LeINI(SVAR_ARQUIVO_POSICAO, "DISCO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "DISCO", "Y")), _
                            Array(ANEL, LeINI(SVAR_ARQUIVO_POSICAO, "ANEL", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "ANEL", "Y")), _
                            Array(PPCORPO, LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y")), _
                            Array(JUNTA, LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "Y")), _
                            Array(CLASSE, LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "Y")), _
                            Array(BITOLA, LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "Y")), _
                            Array(EXTREMIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y")), _
                            Array(OM, LeINI(SVAR_ARQUIVO_POSICAO, "OM", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "OM", "Y")), _
                            Array(CAPACIDADE, LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y")), _
                            Array(DATA, LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "X"), LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "Y")) _
                            )
    Screen.MousePointer = vbDefault
    'exibe tela
    Tela_Chapinha_D43.Show vbModal
End Sub
