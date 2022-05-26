VERSION 5.00
Begin VB.Form Tela_Caracter_7x7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuração dos caracteres 7x7 posições (3,5mm)"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "Salvar Caracter"
      Height          =   375
      Left            =   2160
      TabIndex        =   90
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton BT_Fechar 
      Caption         =   "Fechar Configuração"
      Height          =   375
      Left            =   120
      TabIndex        =   89
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Frame FR_Caracter 
      Caption         =   "Escolha o caracter que deseja configurar:"
      Height          =   3735
      Left            =   120
      TabIndex        =   52
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   """"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   39
         Left            =   1200
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   38
         Left            =   840
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   37
         Left            =   480
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   120
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox PIC_LEG 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   2
         Left            =   120
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   97
         Top             =   3480
         Width           =   135
      End
      Begin VB.PictureBox PIC_LEG 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   1
         Left            =   120
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   96
         Top             =   3240
         Width           =   135
      End
      Begin VB.PictureBox PIC_LEG 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   135
         Index           =   0
         Left            =   120
         ScaleHeight     =   105
         ScaleWidth      =   105
         TabIndex        =   92
         Top             =   3000
         Width           =   135
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   3360
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   34
         Left            =   3000
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   2640
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   2280
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   1920
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   1560
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   1200
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   840
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   480
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   120
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   1920
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   1560
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   1200
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   840
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   480
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   120
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   3360
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   3000
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   2640
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   2280
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   1920
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   1560
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   1200
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   840
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   480
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3360
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3000
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2640
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2280
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1920
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1560
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1200
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   840
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H00FFFFFF&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton BT_CARACTER 
         BackColor       =   &H000000FF&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Caracter já foi configurado"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   95
         Top             =   3240
         Width           =   1860
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Caracter que está sendo configurado agora"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   94
         Top             =   3480
         Width           =   3075
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Caracter ainda não foi configurado"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   93
         Top             =   3000
         Width           =   2445
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Legenda de cores dos caracteres:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   91
         Top             =   2760
         Width           =   2430
      End
   End
   Begin VB.CommandButton BT_Preto 
      Caption         =   "Todos Pretos"
      Height          =   375
      Left            =   6000
      TabIndex        =   51
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton BT_Branco 
      Caption         =   "Todos Brancos"
      Height          =   375
      Left            =   4080
      TabIndex        =   50
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame FR 
      Height          =   3735
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   48
         Left            =   3000
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   49
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   47
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   48
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   46
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   47
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   45
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   46
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   44
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   45
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   43
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   44
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   42
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   43
         Top             =   3120
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   41
         Left            =   3000
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   42
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   41
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   40
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   38
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   39
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   37
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   38
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   36
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   37
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   35
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   36
         Top             =   2640
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   3000
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   35
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   33
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   34
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   32
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   33
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   31
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   32
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   31
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   30
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   29
         Top             =   2160
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   3000
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   26
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   27
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   26
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   24
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   25
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   24
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   22
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   23
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   21
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   22
         Top             =   1680
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   3000
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   21
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   20
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   18
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   19
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   17
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   18
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   16
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   17
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   14
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   15
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   3000
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   10
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   9
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   9
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   3000
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   2520
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   2040
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   1560
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1080
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   600
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox PIC_CAR 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   120
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Tela_Caracter_7x7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub BT_Branco_Click()
    Dim I As Integer
    For I = 0 To 48
        PIC_CAR(I).BackColor = &HFFFFFF
    Next I
End Sub
Private Sub BT_CARACTER_Click(Index As Integer)
    BT_Branco_Click
    CarregaBotaoCaracter Index
    FR.Caption = BT_CARACTER(Index).Caption
    HabilitaPIC True
    If EscrevePIC_CAR(BT_CARACTER(Index).Caption) = True Then
    
    End If
End Sub
Private Sub BT_Fechar_Click()
    Unload Tela_Caracter_7x7
End Sub
Private Sub BT_Preto_Click()
    Dim I As Integer
    For I = 0 To 48
        PIC_CAR(I).BackColor = &H0&
    Next I
End Sub
Private Sub BT_Salvar_Click()
    Dim I, J As Integer
    VAR_INTEMP = 0
    For I = 1 To 7
        For J = 1 To 7
            EscreveINI SVAR_ARQUIVO, FR.Caption, "L" & I & "C" & J, LePIC_CAR(VAR_INTEMP)
            VAR_INTEMP = VAR_INTEMP + 1
        Next J
    Next I
    BVAR_SALVAR = False
    BT_Salvar.Enabled = False
    BT_Branco_Click
    CarregaBotaoCaracter -1
    FR.Caption = ""
    HabilitaPIC False
End Sub
Private Sub Form_Load()
    'define arquivos de caracteres
    SVAR_ARQUIVO = SVAR_CAMINHO_ARQUIVOS & "\7x7.car"
    'limpa todos os quadrados
    BT_Branco_Click
    'carrega botoes dos caracteres
    CarregaBotaoCaracter -1
    'botao salvar desabilitado
    BVAR_SALVAR = False
    BT_Salvar.Enabled = False
    'desabilita clicar nos PICS
    HabilitaPIC False
End Sub
Private Sub PIC_CAR_Click(Index As Integer)
    If PIC_CAR(Index).BackColor = &HFFFFFF Then
        PIC_CAR(Index).BackColor = &H0&
    Else
        PIC_CAR(Index).BackColor = &HFFFFFF
    End If
    BVAR_SALVAR = True
    BT_Salvar.Enabled = True
End Sub

'***************************************************************************
'                       FUNCOES ESPECIFICAS DESTE CÓDIGO
'***************************************************************************

Private Function LePIC_CAR(ByVal Posicao As Integer) As Integer
    If PIC_CAR(Posicao).BackColor = &HFFFFFF Then 'branco = 0
        LePIC_CAR = 0
    Else 'preto = 1
        LePIC_CAR = 1
    End If
End Function
Private Function EscrevePIC_CAR(ByVal Caracter As String) As Boolean
    Dim I, J As Integer
    VAR_INTEMP = 0
    For I = 1 To 7
        For J = 1 To 7
            If LeINI(SVAR_ARQUIVO, Caracter, "L" & I & "C" & J) = "0" Then 'branco = 0
                PIC_CAR(VAR_INTEMP).BackColor = &HFFFFFF
            ElseIf LeINI(SVAR_ARQUIVO, Caracter, "L" & I & "C" & J) = "1" Then 'preto = 1
                PIC_CAR(VAR_INTEMP).BackColor = &H0&
            Else 'nao existe a chave
                EscrevePIC_CAR = False
                Exit Function
            End If
            VAR_INTEMP = VAR_INTEMP + 1
        Next J
    Next I
    EscrevePIC_CAR = True
End Function
Private Sub HabilitaPIC(VALOR As Boolean)
    Dim I As Integer
    For I = 0 To 48
        PIC_CAR(I).Enabled = VALOR
    Next I
    BT_Branco.Enabled = VALOR
    BT_Preto.Enabled = VALOR
End Sub
Private Sub CarregaBotaoCaracter(ByVal CARACTERESCOLHIDO As Integer)
    Dim I As Integer
    For I = 0 To 35 'numero de botoes de caracteres
        If LeINI(SVAR_ARQUIVO, BT_CARACTER(I).Caption, "L1C1") = "NAOEXISTE" Then 'caracter ainda não configurado
            BT_CARACTER(I).BackColor = &HFFFFFF
        Else 'caracter ja foi configurado
            BT_CARACTER(I).BackColor = &HFFFF&
        End If
    Next I
    'verifica se tem algum caracter escolhido
    If CARACTERESCOLHIDO >= 0 Then
        BT_CARACTER(CARACTERESCOLHIDO).BackColor = &HFF&
    End If
End Sub
