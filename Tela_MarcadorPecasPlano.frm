VERSION 5.00
Begin VB.Form Tela_MarcadorPecasPlano 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações do Marcador de Peças Plano"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BT_Fechar 
      Caption         =   "&Fechar"
      Height          =   1335
      Left            =   3000
      Picture         =   "Tela_MarcadorPecasPlano.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Frame FR_Textos 
      Caption         =   "Textos que serão impressos nesta peça:"
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   7095
      Begin VB.CommandButton BT_Textos_RemoverLista 
         Caption         =   "Remover da Lista"
         Height          =   495
         Left            =   1920
         TabIndex        =   32
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton BT_Textos_AdicionarLista 
         Caption         =   "Adicionar na Lista"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TXT_NomeCampo 
         Height          =   285
         Left            =   3720
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CommandButton BT_Textos_Maquina_Esquerda 
         Height          =   615
         Left            =   3360
         Picture         =   "Tela_MarcadorPecasPlano.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_Textos_Maquina_Direita 
         Height          =   615
         Left            =   4080
         Picture         =   "Tela_MarcadorPecasPlano.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_Textos_Maquina_Desce 
         Height          =   615
         Left            =   2640
         Picture         =   "Tela_MarcadorPecasPlano.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_Textos_Maquina_Sobe 
         Height          =   615
         Left            =   1920
         Picture         =   "Tela_MarcadorPecasPlano.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   600
         Width           =   615
      End
      Begin VB.ListBox LT_Textos 
         Height          =   1035
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton BT_Textos_Salvar 
         Caption         =   "Salvar dados"
         Height          =   375
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton BT_Textos_Maquina_Home 
         Caption         =   "HOME"
         Height          =   495
         Left            =   5640
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton BT_Textos_Maquina_Goto 
         Caption         =   "GOTO"
         Height          =   495
         Left            =   6360
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Campo:"
         Height          =   195
         Left            =   3720
         TabIndex        =   29
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label LB_Textos_X 
         AutoSize        =   -1  'True
         Caption         =   "Posição X:"
         Height          =   195
         Left            =   1920
         TabIndex        =   28
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Textos_Y 
         AutoSize        =   -1  'True
         Caption         =   "Posição Y:"
         Height          =   195
         Left            =   3720
         TabIndex        =   27
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame FR_Posicao 
      Caption         =   "Configuração de Posição de Colunas X Linhas:"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   7095
      Begin VB.CommandButton BT_Posicoes_Maquina_Goto 
         Caption         =   "GOTO"
         Height          =   495
         Left            =   6360
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton BT_Posicoes_Maquina_Home 
         Caption         =   "HOME"
         Height          =   495
         Left            =   5640
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton BT_Posicoes_Salvar 
         Caption         =   "Salvar dados"
         Height          =   375
         Left            =   5640
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.ListBox LT_Posicoes 
         Height          =   1035
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton BT_Posicoes_Maquina_Sobe 
         Height          =   615
         Left            =   1920
         Picture         =   "Tela_MarcadorPecasPlano.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_Posicoes_Maquina_Desce 
         Height          =   615
         Left            =   2640
         Picture         =   "Tela_MarcadorPecasPlano.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_Posicoes_Maquina_Direita 
         Height          =   615
         Left            =   4080
         Picture         =   "Tela_MarcadorPecasPlano.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_Posicoes_Maquina_Esquerda 
         Height          =   615
         Left            =   3360
         Picture         =   "Tela_MarcadorPecasPlano.frx":1850
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
      Begin VB.Label LB_Posicoes_Y 
         AutoSize        =   -1  'True
         Caption         =   "Posição Y:"
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Posicoes_X 
         AutoSize        =   -1  'True
         Caption         =   "Posição X:"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame FR_Peca 
      Caption         =   "Escolha a peça que você deseja configurar:"
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton BT_Pecas_Salvar 
         Caption         =   "Salvar dados"
         Height          =   375
         Left            =   5640
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TXT_Linhas 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TXT_Colunas 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox CB_Pecas 
         Height          =   315
         ItemData        =   "Tela_MarcadorPecasPlano.frx":1B5A
         Left            =   120
         List            =   "Tela_MarcadorPecasPlano.frx":1BB5
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peças - número de linhas:"
         Height          =   195
         Left            =   2400
         TabIndex        =   12
         Top             =   600
         Width           =   1830
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Peças - número de colunas:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1980
      End
   End
End
Attribute VB_Name = "Tela_MarcadorPecasPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Fechar_Click()
    Unload Me
End Sub
Private Sub BT_Pecas_Salvar_Click()
    'salva textos de linhas e colunas
    EscreveINI SVAR_ARQUIVO, CB_Pecas.Text, "COLUNAS", TXT_Colunas.Text
    EscreveINI SVAR_ARQUIVO, CB_Pecas.Text, "LINHAS", TXT_Linhas.Text
    BT_Pecas_Salvar.Enabled = False
    'carrega lista de posicoes
    For I = 0 To Int(TXT_Linhas.Text)
        For J = 0 To Int(TXT_Colunas.Text)
            LT_Posicoes.AddItem "L" & Str(I) & "C" & Str(J)
        Next J
    Next I
    LT_Posicoes.Enabled = True
End Sub
Private Sub CB_Pecas_Change()
    CB_Pecas_Click
End Sub
Private Sub CB_Pecas_Click()
    'carrega textos de linhas e colunas
    If CB_Pecas.ListIndex >= 0 Then 'escolheu algum item da lista
        TXT_Colunas.Enabled = False
        TXT_Linhas.Enabled = False
    End If
    TXT_Colunas.Text = LeINI(SVAR_ARQUIVO, CB_Pecas.Text, "COLUNAS")
    TXT_Linhas.Text = LeINI(SVAR_ARQUIVO, CB_Pecas.Text, "LINHAS")
    BT_Pecas_Salvar.Enabled = False
    
    

End Sub
Private Sub Form_Load()
    'caminho do arquivo de configuracoes
    SVAR_ARQUIVO = SVAR_CAMINHO_ARQUIVOS & "\pplana.cfg"
    'limpa tela
    CarregaTela
End Sub




Private Sub CarregaTela()
    TXT_Colunas.Text = ""
    TXT_Linhas.Text = ""
    TXT_Colunas.Enabled = False
    TXT_Linhas.Enabled = False
    BT_Pecas_Salvar.Enabled = False
    LT_Posicoes.Clear
    LT_Posicoes.Enabled = False
    LB_Posicoes_X.Caption = "Posição X: "
    
End Sub
Private Sub TXT_Colunas_Change()
    BT_Pecas_Salvar.Enabled = True
End Sub
Private Sub TXT_Linhas_Change()
    BT_Pecas_Salvar.Enabled = True
End Sub
