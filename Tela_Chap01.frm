VERSION 5.00
Begin VB.Form Tela_Chap01 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chapinha Modelo 1: Gaveta de 1/2"", 3/4"" e 1"""
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Posicao 
      Caption         =   "Posicionamento dos textos na chapinha:"
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4215
      Begin VB.TextBox TXT_Y_CAPACIDADE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   56
         Text            =   "000,0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_DATA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         Text            =   "000,0"
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_OM 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   54
         Text            =   "000,0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_EXTREMIDADE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   53
         Text            =   "000,0"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_CLASSE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   52
         Text            =   "000,0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_BITOLA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   51
         Text            =   "000,0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_PPPREME 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   50
         Text            =   "000,0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_PPCORPO 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   49
         Text            =   "000,0"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_JUNTA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   48
         Text            =   "000,0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_GAXETA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   47
         Text            =   "000,0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_BUCHA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   46
         Text            =   "000,0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_HASTE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   45
         Text            =   "000,0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_ANEIS 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   44
         Text            =   "000,0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_CUNHA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   43
         Text            =   "000,0"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_PREME 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   42
         Text            =   "000,0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TXT_Y_CORPO 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3360
         TabIndex        =   41
         Text            =   "000,0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox TXT_X_CAPACIDADE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   40
         Text            =   "000,0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox TXT_X_DATA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   39
         Text            =   "000,0"
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox TXT_X_OM 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   38
         Text            =   "000,0"
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox TXT_X_EXTREMIDADE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   37
         Text            =   "000,0"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox TXT_X_CLASSE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   36
         Text            =   "000,0"
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox TXT_X_BITOLA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   35
         Text            =   "000,0"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox TXT_X_PPPREME 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Text            =   "000,0"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox TXT_X_PPCORPO 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   33
         Text            =   "000,0"
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox TXT_X_JUNTA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Text            =   "000,0"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox TXT_X_GAXETA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Text            =   "000,0"
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox TXT_X_BUCHA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Text            =   "000,0"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox TXT_X_HASTE 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   29
         Text            =   "000,0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox TXT_X_ANEIS 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   28
         Text            =   "000,0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox TXT_X_CUNHA 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Text            =   "000,0"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TXT_X_PREME 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Text            =   "000,0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox TXT_X_CORPO 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Text            =   "000,0"
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton BT_Tras 
         Height          =   615
         Left            =   3360
         Picture         =   "Tela_Chap01.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton BT_Frente 
         Height          =   615
         Left            =   2280
         Picture         =   "Tela_Chap01.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton BT_Desce 
         Height          =   615
         Left            =   1200
         Picture         =   "Tela_Chap01.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton BT_Sobe 
         Height          =   615
         Left            =   120
         Picture         =   "Tela_Chap01.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4440
         Width           =   615
      End
      Begin VB.OptionButton RB_CAPACIDADE 
         Caption         =   "CAPACIDADE"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   1455
      End
      Begin VB.OptionButton RB_DATA 
         Caption         =   "DATA"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   855
      End
      Begin VB.OptionButton RB_OM 
         Caption         =   "OM / N�"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   975
      End
      Begin VB.OptionButton RB_EXTREMIDADE 
         Caption         =   "EXTREMIDADE"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   1575
      End
      Begin VB.OptionButton RB_CLASSE 
         Caption         =   "CLASSE"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1095
      End
      Begin VB.OptionButton RB_BITOLA 
         Caption         =   "BITOLA"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   975
      End
      Begin VB.OptionButton RB_PPPREME 
         Caption         =   "P/P PREME"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   1335
      End
      Begin VB.OptionButton RB_PPCORPO 
         Caption         =   "P/P CORPO"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1335
      End
      Begin VB.OptionButton RB_JUNTA 
         Caption         =   "JUNTA"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton RB_GAXETA 
         Caption         =   "GAXETA"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton RB_BUCHA 
         Caption         =   "BUCHA"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton RB_HASTE 
         Caption         =   "HASTE"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton RB_CUNHA 
         Caption         =   "CUNHA"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton RB_PREME 
         Caption         =   "PREME"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton RB_CORPO 
         Caption         =   "CORPO/CASTELO"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton RB_ANEIS 
         Caption         =   "AN�IS"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Campo:"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   540
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Y (Passo):"
         Height          =   195
         Index           =   22
         Left            =   3360
         TabIndex        =   20
         Top             =   240
         Width           =   720
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "X (Passo):"
         Height          =   195
         Index           =   21
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   1335
      Left            =   2760
      Picture         =   "Tela_Chap01.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton BT_Fechar 
      Caption         =   "&Fechar"
      Height          =   1335
      Left            =   600
      Picture         =   "Tela_Chap01.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "Tela_Chap01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VETOR As Variant
Private Sub BT_Desce_Click()
    MoveUmPassoEixoY_Antihorario
    AtualizaCamposPosicao 'atualiza campos
End Sub
Private Sub BT_Fechar_Click()
    Unload Tela_Chap01
End Sub
Private Sub BT_Frente_Click()
    MoveUmPassoEixoX_Horario
    AtualizaCamposPosicao 'atualiza campos
End Sub
Private Sub BT_Salvar_Click()
    'salvo todos dados de posicao
    EscreveINI SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "X", TXT_X_CORPO.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "Y", TXT_Y_CORPO.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "PREME", "X", TXT_X_PREME.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "PREME", "Y", TXT_Y_PREME.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "CUNHA", "X", TXT_X_CUNHA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "CUNHA", "Y", TXT_Y_CUNHA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "ANEIS", "X", TXT_X_ANEIS.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "ANEIS", "Y", TXT_Y_ANEIS.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "HASTE", "X", TXT_X_HASTE.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "HASTE", "Y", TXT_Y_HASTE.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "BUCHA", "X", TXT_X_BUCHA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "BUCHA", "Y", TXT_Y_BUCHA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "JUNTA", "X", TXT_X_JUNTA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "JUNTA", "Y", TXT_Y_JUNTA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "GAXETA", "X", TXT_X_GAXETA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "GAXETA", "Y", TXT_Y_GAXETA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "PPCORPO", "X", TXT_X_PPCORPO.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y", TXT_Y_PPCORPO.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "PPPREME", "X", TXT_X_PPPREME.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "PPPREME", "Y", TXT_Y_PPPREME.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "BITOLA", "X", TXT_X_BITOLA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "BITOLA", "Y", TXT_Y_BITOLA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "CLASSE", "X", TXT_X_CLASSE.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "CLASSE", "Y", TXT_Y_CLASSE.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X", TXT_X_EXTREMIDADE.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y", TXT_Y_EXTREMIDADE.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "OM", "X", TXT_X_OM.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "OM", "Y", TXT_Y_OM.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "DATA", "X", TXT_X_DATA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "DATA", "Y", TXT_Y_DATA.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X", TXT_X_CAPACIDADE.Text
    EscreveINI SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y", TXT_Y_CAPACIDADE.Text
End Sub
Private Sub BT_Sobe_Click()
    MoveUmPassoEixoY_Horario
    AtualizaCamposPosicao 'atualiza campos
End Sub
Private Sub BT_Tras_Click()
    MoveUmPassoEixoX_Antihorario
    AtualizaCamposPosicao 'atualiza campos
End Sub
Private Sub Form_Load()
    'define nomes dos arquivos desta tela
    SVAR_ARQUIVO_POSICAO = SVAR_CAMINHO_ARQUIVOS & "\chap01.pos"
    'carrega valores das posi��es pr�-configuradas
    TXT_X_CORPO.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "X")
    TXT_Y_CORPO.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CORPO/CASTELO", "Y")
    TXT_X_PREME.Text = LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "X")
    TXT_Y_PREME.Text = LeINI(SVAR_ARQUIVO_POSICAO, "PREME", "Y")
    TXT_X_CUNHA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CUNHA", "X")
    TXT_Y_CUNHA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CUNHA", "Y")
    TXT_X_ANEIS.Text = LeINI(SVAR_ARQUIVO_POSICAO, "ANEIS", "X")
    TXT_Y_ANEIS.Text = LeINI(SVAR_ARQUIVO_POSICAO, "ANEIS", "Y")
    TXT_X_HASTE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "X")
    TXT_Y_HASTE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "HASTE", "Y")
    TXT_X_BUCHA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "X")
    TXT_Y_BUCHA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "BUCHA", "Y")
    TXT_X_JUNTA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "X")
    TXT_Y_JUNTA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "JUNTA", "Y")
    TXT_X_GAXETA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "X")
    TXT_Y_GAXETA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "GAXETA", "Y")
    TXT_X_PPCORPO.Text = LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "X")
    TXT_Y_PPCORPO.Text = LeINI(SVAR_ARQUIVO_POSICAO, "PPCORPO", "Y")
    TXT_X_PPPREME.Text = LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "X")
    TXT_Y_PPPREME.Text = LeINI(SVAR_ARQUIVO_POSICAO, "PPPREME", "Y")
    TXT_X_BITOLA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "X")
    TXT_Y_BITOLA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "BITOLA", "Y")
    TXT_X_CLASSE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "X")
    TXT_Y_CLASSE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CLASSE", "Y")
    TXT_X_EXTREMIDADE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "X")
    TXT_Y_EXTREMIDADE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "EXTREMIDADE", "Y")
    TXT_X_OM.Text = LeINI(SVAR_ARQUIVO_POSICAO, "OM", "X")
    TXT_Y_OM.Text = LeINI(SVAR_ARQUIVO_POSICAO, "OM", "Y")
    TXT_X_DATA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "X")
    TXT_Y_DATA.Text = LeINI(SVAR_ARQUIVO_POSICAO, "DATA", "Y")
    TXT_X_CAPACIDADE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "X")
    TXT_Y_CAPACIDADE.Text = LeINI(SVAR_ARQUIVO_POSICAO, "CAPACIDADE", "Y")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Tela_Posicao.Show vbModal
End Sub

'***************************************************************************
'                           FUN�OES DESTE C�DIGO
'***************************************************************************

Private Sub AtualizaCamposPosicao()
    If RB_CORPO.Value = True Then
        TXT_X_CORPO.Text = IVAR_PASSO_X
        TXT_Y_CORPO.Text = IVAR_PASSO_Y
    ElseIf RB_PREME.Value = True Then
        TXT_X_PREME.Text = IVAR_PASSO_X
        TXT_Y_PREME.Text = IVAR_PASSO_Y
    ElseIf RB_CUNHA.Value = True Then
        TXT_X_CUNHA.Text = IVAR_PASSO_X
        TXT_Y_CUNHA.Text = IVAR_PASSO_Y
    ElseIf RB_ANEIS.Value = True Then
        TXT_X_ANEIS.Text = IVAR_PASSO_X
        TXT_Y_ANEIS.Text = IVAR_PASSO_Y
    ElseIf RB_HASTE.Value = True Then
        TXT_X_HASTE.Text = IVAR_PASSO_X
        TXT_Y_HASTE.Text = IVAR_PASSO_Y
    ElseIf RB_BUCHA.Value = True Then
        TXT_X_BUCHA.Text = IVAR_PASSO_X
        TXT_Y_BUCHA.Text = IVAR_PASSO_Y
    ElseIf RB_GAXETA.Value = True Then
        TXT_X_GAXETA.Text = IVAR_PASSO_X
        TXT_Y_GAXETA.Text = IVAR_PASSO_Y
    ElseIf RB_JUNTA.Value = True Then
        TXT_X_JUNTA.Text = IVAR_PASSO_X
        TXT_Y_JUNTA.Text = IVAR_PASSO_Y
    ElseIf RB_PPCORPO.Value = True Then
        TXT_X_PPCORPO.Text = IVAR_PASSO_X
        TXT_Y_PPCORPO.Text = IVAR_PASSO_Y
    ElseIf RB_PPPREME.Value = True Then
        TXT_X_PPPREME.Text = IVAR_PASSO_X
        TXT_Y_PPPREME.Text = IVAR_PASSO_Y
    ElseIf RB_BITOLA.Value = True Then
        TXT_X_BITOLA.Text = IVAR_PASSO_X
        TXT_Y_BITOLA.Text = IVAR_PASSO_Y
    ElseIf RB_CLASSE.Value = True Then
        TXT_X_CLASSE.Text = IVAR_PASSO_X
        TXT_Y_CLASSE.Text = IVAR_PASSO_Y
    ElseIf RB_EXTREMIDADE.Value = True Then
        TXT_X_CLASSE.Text = IVAR_PASSO_X
        TXT_Y_CLASSE.Text = IVAR_PASSO_Y
    ElseIf RB_OM.Value = True Then
        TXT_X_OM.Text = IVAR_PASSO_X
        TXT_Y_OM.Text = IVAR_PASSO_Y
    ElseIf RB_DATA.Value = True Then
        TXT_X_DATA.Text = IVAR_PASSO_X
        TXT_Y_DATA.Text = IVAR_PASSO_Y
    ElseIf RB_CAPACIDADE.Value = True Then
        TXT_X_CAPACIDADE.Text = IVAR_PASSO_X
        TXT_Y_CAPACIDADE.Text = IVAR_PASSO_Y
    End If
End Sub

