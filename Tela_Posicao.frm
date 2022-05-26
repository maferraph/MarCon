VERSION 5.00
Begin VB.Form Tela_Posicao 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações de Posição"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TIMER_CHAPINHA_POSICIONAMENTO_HOME 
      Enabled         =   0   'False
      Left            =   1920
      Top             =   4320
   End
   Begin VB.Timer TIMER_CHAPINHA_POSICIONAMENTO_EIXOS 
      Enabled         =   0   'False
      Left            =   1440
      Top             =   4320
   End
   Begin VB.CommandButton BT_Fechar 
      Caption         =   "&Fechar"
      Height          =   1335
      Left            =   120
      Picture         =   "Tela_Posicao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chapinhas de Válvula:"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton BT_CHAP_M6 
         Caption         =   "Modelo 6: Retenção Portinhola (todas)"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   3015
      End
      Begin VB.CommandButton BT_CHAP_M5 
         Caption         =   "Modelo 5: Retenção Pistão (todas)"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton BT_CHAP_M4 
         Caption         =   "Modelo 4: Globo de 1.1/2"" e 2"""
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton BT_CHAP_M3 
         Caption         =   "Modelo 3: Globo de 1/2"" , 3/4"" e 1"""
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton BT_CHAP_M2 
         Caption         =   "Modelo 2: Gavetas de 1.1/2"" e 2"""
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton BT_CHAP_M1 
         Caption         =   "Modelo 1: Gavetas de 1/2"" , 3/4"" e 1"""
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Tela_Posicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_CHAP_M1_Click()
    Tela_Posicao.Hide
    Tela_Chap01.Show vbModal
End Sub
Private Sub BT_CHAP_M2_Click()
    Tela_Posicao.Hide
    Tela_Chap02.Show vbModal
End Sub
Private Sub BT_CHAP_M3_Click()
    Tela_Posicao.Hide
    Tela_Chap03.Show vbModal
End Sub
Private Sub BT_CHAP_M4_Click()
    Tela_Posicao.Hide
    Tela_Chap04.Show vbModal
End Sub
Private Sub BT_CHAP_M5_Click()
    Tela_Posicao.Hide
    Tela_Chap05.Show vbModal
End Sub
Private Sub BT_CHAP_M6_Click()
    Tela_Posicao.Hide
    Tela_Chap06.Show vbModal
End Sub
Private Sub BT_Fechar_Click()
    Unload Tela_Posicao
End Sub
Private Sub TIMER_CHAPINHA_POSICIONAMENTO_EIXOS_Timer()
    MoveEixos_MarcadorChapinha
End Sub
Private Sub TIMER_CHAPINHA_POSICIONAMENTO_HOME_Timer()
    'move eixo até X0 e Y0
    MoveEixos_MarcadorChapinha
    'verifica se chegou em home
    LePortaStatus
    If BVAR_LPT1_P10 = 1 And BVAR_LPT1_P11 = 1 Then 'porta lógica invertida - P10=zero-eixoX e P11=zero-eixoY
        'zera contadores
        IVAR_PASSO_X = 0
        IVAR_PASSO_Y = 0
        SGVAR_POSICAO_X = 0
        SGVAR_POSICAO_Y = 0
        TIMER_CHAPINHA_POSICIONAMENTO_HOME.Enabled = False
    End If
End Sub
