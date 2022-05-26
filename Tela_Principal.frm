VERSION 5.00
Begin VB.Form Tela_Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MarCon - Marcador de Peças da Conesteel"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   5040
      Top             =   1680
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Solenoide"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Para"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Antihorário"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Horário"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Branco"
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Marrom"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vermelho"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Amarelo"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton BT_Redondas 
      Caption         =   "Peças Redondas"
      Height          =   1335
      Left            =   3720
      Picture         =   "Tela_Principal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BT_Forjados 
      Caption         =   "Forjados && Fundidos"
      Height          =   1335
      Left            =   2520
      Picture         =   "Tela_Principal.frx":49E2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BT_Valvulas 
      Caption         =   "Válvulas"
      Height          =   1335
      Left            =   1320
      Picture         =   "Tela_Principal.frx":A1D0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BT_Chapinhas 
      Caption         =   "Chapinhas"
      Height          =   1335
      Left            =   120
      Picture         =   "Tela_Principal.frx":A4DA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Menu Menu_Sair 
      Caption         =   "&Sair"
   End
   Begin VB.Menu Menu_Configuracoes 
      Caption         =   "&Configurações"
      Begin VB.Menu Menu_Configuracoes_Caracter35 
         Caption         =   "Caracter de 3,5mm"
      End
   End
End
Attribute VB_Name = "Tela_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VALOR As Integer
Private Sub BT_Chapinhas_Click()
    Tela_Principal.Hide
    Tela_Chapinhas.Show
End Sub

Private Sub BT_Forjados_Click()
    EscrevePorta &H378, Text1.Text
    Text1.Text = Text1.Text + 1
    If Text1.Text > 64 Then
        Text1.Text = 0
    End If

End Sub

Private Sub BT_Valvulas_Click()
    EscrevePorta &H378, Text1.Text
    Text1.Text = Text1.Text - 1
    If Text1.Text < 0 Then
        Text1.Text = 0
    End If
End Sub

Private Sub Command1_Click()
    EscrevePorta &H378, 1
End Sub

Private Sub Command2_Click()
    EscrevePorta &H378, 2
End Sub

Private Sub Command3_Click()
    EscrevePorta &H378, 4
End Sub

Private Sub Command4_Click()
    EscrevePorta &H378, 8
End Sub

Private Sub Command5_Click()

Tela_Chapinhas.TIMER_POSICIONAMENTO_EIXO_X_HORARIO.Enabled = True
Tela_Chapinhas.TIMER_POSICIONAMENTO_EIXO_X_ANTIHORARIO.Enabled = False
'    Timer1.Enabled = True
'    Timer2.Enabled = False

End Sub

Private Sub Command6_Click()
Tela_Chapinhas.TIMER_POSICIONAMENTO_EIXO_X_HORARIO.Enabled = False
Tela_Chapinhas.TIMER_POSICIONAMENTO_EIXO_X_ANTIHORARIO.Enabled = True

'    Timer1.Enabled = False
'    Timer2.Enabled = True

End Sub

Private Sub Command7_Click()
Tela_Chapinhas.TIMER_POSICIONAMENTO_EIXO_X_HORARIO.Enabled = False
Tela_Chapinhas.TIMER_POSICIONAMENTO_EIXO_X_ANTIHORARIO.Enabled = False
'    Timer1.Enabled = False
        
        
        BVAR_LPT_P2 = 0
        BVAR_LPT_P3 = 0
        BVAR_LPT_P4 = 0
        BVAR_LPT_P5 = 0
    
BVAR_LPT_P14 = 1

    

End Sub

Private Sub Command8_Click()
    If Timer3.Enabled = True Then
        Timer3.Enabled = False
    Else
        Timer3.Enabled = True
        'Timer1.Interval = Text1.Text
    End If
End Sub

Private Sub Form_Load()
    VALOR = 0
    LVAR_PASSO_X = 0
    SGVAR_POSICAO_X = 0
    Text1.Text = 100
    Timer3.Enabled = False
End Sub

Private Sub Menu_Configuracoes_Caracter35_Click()
    Tela_Principal.Hide
    Tela_Caracter_7x7.Show
End Sub

Private Sub Menu_Sair_Click()
    End
End Sub

Private Sub Text1_Change()
    If IsNumeric(Text1.Text) Then
    Timer3.Interval = Text1.Text
    End If
End Sub


Private Sub Timer3_Timer()
        
    If VALOR = 1 Then
       EscrevePorta &H378, 255
       VALOR = 0
     Else
        EscrevePorta &H378, 0
        VALOR = 1
    End If
    
    


End Sub
