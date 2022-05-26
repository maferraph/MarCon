VERSION 5.00
Begin VB.Form Tela_Chapinha_D43 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Marcando a chapinha..."
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Emergencia 
      Height          =   615
      Left            =   5760
      Picture         =   "Tela_Chapinha_D43.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "EMERGÊNCIA - Desliga máquina"
      Top             =   6600
      Width           =   615
   End
   Begin VB.Timer TIMER_POSICIONAMENTO_EIXOS 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox PIC_CHAP 
      Height          =   6495
      Left            =   0
      ScaleHeight     =   429
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label LB_Pontos 
      AutoSize        =   -1  'True
      Caption         =   "0 de 0"
      Height          =   195
      Left            =   4320
      TabIndex        =   6
      Top             =   6960
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pontos Marcados:"
      Height          =   195
      Left            =   4320
      TabIndex        =   5
      Top             =   6720
      Width           =   1290
   End
   Begin VB.Label LB_PosicaoX 
      AutoSize        =   -1  'True
      Caption         =   "Posição Eixo X:"
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   6720
      Width           =   1110
   End
   Begin VB.Label LB_PassoX 
      AutoSize        =   -1  'True
      Caption         =   "Passo Eixo X:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label LB_PosicaoY 
      AutoSize        =   -1  'True
      Caption         =   "Posição Eixo Y:"
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   6960
      Width           =   1110
   End
   Begin VB.Label LB_PassoY 
      AutoSize        =   -1  'True
      Caption         =   "Passo Eixo Y:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   975
   End
End
Attribute VB_Name = "Tela_Chapinha_D43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VALOR_VETOR As Long
Private Const ICONST_PASSO_CANTOSUPERIORESQUERDO_X As Integer = 0 'em passos
Private Const ICONST_PASSO_CANTOSUPERIORESQUERDO_Y As Integer = 0 'em passos
Private Sub BT_Emergencia_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    VALOR_VETOR = 0
    'apaga PIC
    PIC_CHAP.Cls
    If SVAR_MARCACAO_ATUAL = "CHAP01" Then
        PIC_CHAP.Picture = LoadPicture(SVAR_CAMINHO_ARQUIVOS & "\chap01.fig")
    ElseIf SVAR_MARCACAO_ATUAL = "CHAP03" Then
        PIC_CHAP.Picture = LoadPicture(SVAR_CAMINHO_ARQUIVOS & "\chap03.fig")
    ElseIf SVAR_MARCACAO_ATUAL = "CHAP05" Then
        PIC_CHAP.Picture = LoadPicture(SVAR_CAMINHO_ARQUIVOS & "\chap05.fig")
    ElseIf SVAR_MARCACAO_ATUAL = "CHAP06" Then
        PIC_CHAP.Picture = LoadPicture(SVAR_CAMINHO_ARQUIVOS & "\chap06.fig")
    End If
    'começa marcação
    TIMER_POSICIONAMENTO_EIXOS.Interval = ICONST_TEMPO_ESPERA_PASSO_MOTOR
    TIMER_POSICIONAMENTO_EIXOS.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    TIMER_POSICIONAMENTO_EIXOS.Enabled = False
End Sub
Private Sub TIMER_POSICIONAMENTO_EIXOS_Timer()
    If IVAR_PASSODESTINO_EIXOX <> VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(0) And IVAR_PASSODESTINO_EIXOY <> VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(1) Then
        PosicionaMarcadorChapinha VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(0), VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(1)
    End If
    LB_Pontos.Caption = VALOR_VETOR & " de " & UBound(VETOR_POSICIONAMENTO_MARCACAO)
    'quando chegar no ponto
    If IVAR_PASSO_X = VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(0) And IVAR_PASSO_Y = VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(1) Then
        'marca o ponto
        AtuaPistao_MarcadorChapinha
        'desenha no PICTUREBOX
        PIC_CHAP.PSet ((VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(0) - ICONST_PASSO_CANTOSUPERIORESQUERDO_X) * 10 * DCONST_MOVIMENTO_POR_PASSO, (VETOR_POSICIONAMENTO_MARCACAO(VALOR_VETOR)(1) - ICONST_PASSO_CANTOSUPERIORESQUERDO_Y) * 10 * DCONST_MOVIMENTO_POR_PASSO), QBColor(0)
        'muda campo do vetor
        VALOR_VETOR = VALOR_VETOR + 1
    End If
    'atualiza ponto na tela
    LB_PassoX.Caption = "Passo Eixo X: " & Int(IVAR_PASSO_X)
    LB_PosicaoX.Caption = "Posição Eixo X: " & SGVAR_POSICAO_X & " mm"
    LB_PassoY.Caption = "Passo Eixo Y: " & Int(IVAR_PASSO_Y)
    LB_PosicaoY.Caption = "Posição Eixo Y: " & SGVAR_POSICAO_Y & " mm"
    'verifica se acabou de marcar o vetor
    If VALOR_VETOR = UBound(VETOR_POSICIONAMENTO_MARCACAO) Then
        TIMER_POSICIONAMENTO_EIXOS.Enabled = False
        Unload Tela_Chapinha_D43
    End If
End Sub
