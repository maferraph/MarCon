VERSION 5.00
Begin VB.Form Tela_Chapinhas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escolha o modelo da Chapinha"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10860
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TIMER_POSICIONAMENTO_EIXO_X_ANTIHORARIO 
      Enabled         =   0   'False
      Left            =   2160
      Top             =   7680
   End
   Begin VB.Timer TIMER_POSICIONAMENTO_EIXO_X_HORARIO 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   7680
   End
   Begin VB.Timer TIMER_POSICIONAMENTO_HOME 
      Enabled         =   0   'False
      Left            =   240
      Top             =   7680
   End
   Begin VB.CommandButton BT_M6 
      Caption         =   "Modelo 6: Retenção Portinhola (todas)"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   4335
   End
   Begin VB.CommandButton BT_M5 
      Caption         =   "Modelo 5: Retenção Pistão (todas)"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   4335
   End
   Begin VB.CommandButton BT_M4 
      Caption         =   "Modelo 4: Globo de 1.1/2"" e 2"""
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4335
   End
   Begin VB.CommandButton BT_M3 
      Caption         =   "Modelo 3: Globo de 1/2"" , 3/4"" e 1"""
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton BT_M2 
      Caption         =   "Modelo 2: Gavetas de 1.1/2"" e 2"""
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton BT_M1 
      Caption         =   "Modelo 1: Gavetas de 1/2"" , 3/4"" e 1"""
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Tela_Chapinhas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BT_M1_Click()
    Tela_Chapinhas.Hide
    Tela_Chap01.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Tela_Principal.Show
End Sub
Private Sub TIMER_POSICIONAMENTO_EIXO_X_ANTIHORARIO_Timer()
    MoveUmPassoEixoX_Antihorario
End Sub
Private Sub TIMER_POSICIONAMENTO_EIXO_X_HORARIO_Timer()
    MoveUmPassoEixoX_Horario
End Sub

Private Sub TIMER_POSICIONAMENTO_HOME_Timer()
    'neste timer está o código que irá rodar os motores X e Y até chegarem em home
    
    Tela_Chapinhas.TIMER_POSICIONAMENTO_HOME.Enabled = False 'apagar depois esta linha
End Sub
