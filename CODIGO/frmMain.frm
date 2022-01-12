VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Balanceo"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   8895
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8400
         TabIndex        =   87
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   8400
         TabIndex        =   86
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   8400
         TabIndex        =   85
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8400
         TabIndex        =   84
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8400
         TabIndex        =   83
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8400
         TabIndex        =   82
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8040
         TabIndex        =   81
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7680
         TabIndex        =   80
         Top             =   5280
         Width           =   375
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   8040
         TabIndex        =   79
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7680
         TabIndex        =   78
         Top             =   4560
         Width           =   375
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   8040
         TabIndex        =   77
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7680
         TabIndex        =   76
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   8040
         TabIndex        =   75
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7680
         TabIndex        =   74
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8040
         TabIndex        =   73
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7680
         TabIndex        =   72
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   8040
         TabIndex        =   70
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtModDañoWre 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   69
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7680
         TabIndex        =   67
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtModAtaqueArmas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   66
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtModAtaqueWre 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   64
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtModDañoArmas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   62
         Top             =   4560
         Width           =   975
      End
      Begin VB.TextBox txtMODESCUDO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   60
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtModEvasion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         MaxLength       =   5
         TabIndex        =   58
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtApuñalarAtacante 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   55
         Top             =   4440
         Width           =   615
      End
      Begin VB.ComboBox lstDefensaExtraOponente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   3480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   5760
         Width           =   2775
      End
      Begin VB.ComboBox lstDañoExtraAtacante 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   5400
         Width           =   2775
      End
      Begin VB.TextBox txtAgilidadAtacante 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   47
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtCombate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   46
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox txtWrestling 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   45
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtFuerzaAtacante 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   44
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtAgilidadOponente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   42
         Top             =   5160
         Width           =   615
      End
      Begin VB.TextBox txtDefensa 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   40
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtTacticas 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   39
         Top             =   4800
         Width           =   615
      End
      Begin VB.ComboBox lstClaseOponente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   3480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1920
         Width           =   2835
      End
      Begin VB.ComboBox lstRazaOponente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   3480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1200
         Width           =   2835
      End
      Begin VB.CommandButton cmdClean 
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6885
         TabIndex        =   32
         Top             =   6360
         Width           =   1575
      End
      Begin VB.TextBox txtNivelOponente 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   29
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox txtNivelAtacante 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         MaxLength       =   3
         TabIndex        =   27
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton cmdAttack 
         Caption         =   "¡ATACAR!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   5760
         Width           =   2775
      End
      Begin VB.ComboBox lstRazaAtacante 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   2835
      End
      Begin VB.ComboBox lstClaseAtacante 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1920
         Width           =   2835
      End
      Begin VB.ComboBox ListaArmadurasOponente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   3480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3360
         Width           =   2835
      End
      Begin VB.ComboBox lstCascoOponente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   3480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2640
         Width           =   2835
      End
      Begin VB.ComboBox lstEscudoOponente 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   3480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4080
         Width           =   2835
      End
      Begin VB.ComboBox lstArmasAtacante 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IntegralHeight  =   0   'False
         Left            =   360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2640
         Width           =   2835
      End
      Begin RichTextLib.RichTextBox RecTxt 
         Height          =   1095
         Left            =   600
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Mensajes del servidor"
         Top             =   6720
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   1931
         _Version        =   393217
         BackColor       =   0
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmMain.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "¡¡¡SEPARAR CON "",""!!!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   6600
         TabIndex        =   88
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "G= GUARDAR, D= DEFAULT, A=APLICAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   71
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daño nudi/puño (ATACANTE)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   68
         Top             =   2800
         Width           =   2055
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prob acierto armas (ATACANTE)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   65
         Top             =   4850
         Width           =   2055
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prob daño nudi/puño (ATACANTE)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   63
         Top             =   3550
         Width           =   2055
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daño de armas (ATACANTE)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   61
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escudos (VICTIMA)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   59
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Evasion (VICTIMA)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   57
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   56
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Apuñalar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   54
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Defensa de montura/barca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   53
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daño de montura/barca"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   51
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Combate sin armas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   49
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Agilidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   48
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fuerza:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   43
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agilidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tacticas y defensa escu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   38
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Combate con armas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   35
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Raza"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   34
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Consola..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   6360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Victima:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Atacante:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Min Golpe / Max Golpe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Raza"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Armaduras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cascos y gorros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escudos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Armas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   2400
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de objeto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5280
      TabIndex        =   1
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox PicView 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   1560
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblDaño 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Daño: -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblDef 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Defensa: -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblNivel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel minimo: -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre: -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3495
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ataque/Defensa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label lblcargando 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private nLineas As Long
Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub cmdAplicar_Click(Index As Integer)

    Dim ClaseOponente As Byte
    
    ClaseOponente = ObtenerClase(frmMain.lstClaseOponente.List(frmMain.lstClaseOponente.ListIndex))
    
    Dim ClaseAtacante As Byte
    
    ClaseAtacante = ObtenerClase(frmMain.lstClaseAtacante.List(frmMain.lstClaseAtacante.ListIndex))
 
    Select Case Index

        Case 0 'Evasion

            If ClaseOponente = 0 Then
                MsgBox "Debes asignar las clases para ejecutar esta acción."
                Exit Sub
            End If

            ModClase(ClaseOponente).Evasion = CDbl(frmMain.txtModEvasion.Text)

        Case 1 'Escudos

            If ClaseOponente = 0 Then
                MsgBox "Debes asignar las clases para ejecutar esta acción."
                Exit Sub
            End If
            
            ModClase(ClaseOponente).Escudo = CDbl(frmMain.txtMODESCUDO.Text)

        Case 2

            If ClaseAtacante = 0 Then
                MsgBox "Debes asignar las clases para ejecutar esta acción."
                Exit Sub
            End If
            
            ModClase(ClaseAtacante).DañoWrestling = CDbl(frmMain.txtModDañoWre.Text)

        Case 3

            If ClaseAtacante = 0 Then
                MsgBox "Debes asignar las clases para ejecutar esta acción."
                Exit Sub
            End If
            
            ModClase(ClaseAtacante).AtaqueWrestling = CDbl(frmMain.txtModAtaqueWre.Text)

        Case 4

            If ClaseAtacante = 0 Then
                MsgBox "Debes asignar las clases para ejecutar esta acción."
                Exit Sub
            End If
            
            ModClase(ClaseAtacante).DañoArmas = CDbl(frmMain.txtModDañoArmas.Text)

        Case 5

            If ClaseAtacante = 0 Then
                MsgBox "Debes asignar las clases para ejecutar esta acción."
                Exit Sub
            End If
 
            ModClase(ClaseAtacante).AtaqueArmas = CDbl(frmMain.txtModAtaqueArmas.Text)
            
    End Select
    
    lblcargando.Caption = "Aplicado con éxito. (" & Time & ")"
    
End Sub

Private Sub cmdClean_Click()
Me.RecTxt.Text = ""
End Sub
 Public Sub RefreshPicView()
    PicView.Picture = Nothing
    Me.lblNombre.Caption = "Nombre: - "
    Me.lblNivel.Caption = "Nivel minimo: - "
    Me.lblDef.Caption = "Defensa: - "
    Me.lblDaño.Caption = "Daño: - "
End Sub




Private Sub cmdDefault_Click(Index As Integer)

    Dim ClaseOponente As Byte
    
    ClaseOponente = ObtenerClase(frmMain.lstClaseOponente.List(frmMain.lstClaseOponente.ListIndex))
    
    Dim ClaseAtacante As Byte
    
    ClaseAtacante = ObtenerClase(frmMain.lstClaseAtacante.List(frmMain.lstClaseAtacante.ListIndex))

    If MsgBox("¿Está seguro de cargar en default este item?", vbYesNo) = vbYes Then
        
        Select Case Index

            Case 0 'Evasion
    
                If ClaseOponente = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
    
                ModClase(ClaseOponente).Evasion = ModClase(ClaseOponente).EvasionBackUp
                Me.txtModEvasion.Text = CStr(ModClase(ClaseOponente).Evasion)
        
            Case 1 'Escudos
    
                If ClaseOponente = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
                
                ModClase(ClaseOponente).Escudo = ModClase(ClaseOponente).EscudoBackUp
                Me.txtMODESCUDO.Text = CStr(ModClase(ClaseOponente).Escudo)
            
            Case 2
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
    
                ModClase(ClaseAtacante).DañoWrestling = ModClase(ClaseAtacante).DañoWrestlingBackUp
                Me.txtModDañoWre.Text = CStr(ModClase(ClaseAtacante).DañoWrestling)

            Case 3
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
    
                ModClase(ClaseAtacante).AtaqueWrestling = ModClase(ClaseAtacante).AtaqueWrestlingBackUp
                Me.txtModAtaqueWre.Text = CStr(ModClase(ClaseAtacante).AtaqueWrestling)
                
            Case 4
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
    
                ModClase(ClaseAtacante).DañoArmas = ModClase(ClaseAtacante).DañoArmasBackUp
                Me.txtModDañoArmas.Text = CStr(ModClase(ClaseAtacante).DañoArmas)
                
            Case 5
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
                
                ModClase(ClaseAtacante).AtaqueArmas = ModClase(ClaseAtacante).AtaqueArmasBackUp
                Me.txtModAtaqueArmas.Text = CStr(ModClase(ClaseAtacante).AtaqueArmas)
                
        End Select
        
    End If
    
    lblcargando.Caption = "Restaurado con éxito. (" & Time & ")"
    
End Sub

Private Sub cmdSave_Click(Index As Integer)
    
    Dim tmp As String
    
    Dim PathAux As String
    
    PathAux = App.Path & "\Datos\Balance.dat"
    
    Dim ClaseOponente As Byte

    ClaseOponente = ObtenerClase(frmMain.lstClaseOponente.List(frmMain.lstClaseOponente.ListIndex))

    Dim ClaseAtacante As Byte

    ClaseAtacante = ObtenerClase(frmMain.lstClaseAtacante.List(frmMain.lstClaseAtacante.ListIndex))

    If MsgBox("¿Está seguro de guardar este item?", vbYesNo) = vbYes Then
        
        Select Case Index
        
            Case 0 'Evasion
    
                If ClaseOponente = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
                                
                Call cmdAplicar_Click(Index)
                 
                tmp = CStr(ModClase(ClaseOponente).Evasion)
                tmp = Replace(tmp, ",", ".")
                Call WriteVar(PathAux, "MODEVASION", ListaClases(ClaseOponente), tmp)
                
            Case 1 'Escudos
    
                If ClaseOponente = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
                
                Call cmdAplicar_Click(Index)
                tmp = CStr(ModClase(ClaseOponente).Escudo)
                tmp = Replace(tmp, ",", ".")
                Call WriteVar(PathAux, "MODEVASION", ListaClases(ClaseOponente), tmp)
                
            Case 2
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
                
                Call cmdAplicar_Click(Index)
                tmp = CStr(ModClase(ClaseOponente).DañoWrestling)
                tmp = Replace(tmp, ",", ".")
                Call WriteVar(PathAux, "MODEVASION", ListaClases(ClaseOponente), tmp)
                
            Case 3
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
                
                Call cmdAplicar_Click(Index)
                tmp = CStr(ModClase(ClaseOponente).AtaqueWrestling)
                tmp = Replace(tmp, ",", ".")
                Call WriteVar(PathAux, "MODEVASION", ListaClases(ClaseOponente), tmp)
                
            Case 4
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
            
                Call cmdAplicar_Click(Index)
                tmp = CStr(ModClase(ClaseOponente).DañoArmas)
                tmp = Replace(tmp, ",", ".")
                Call WriteVar(PathAux, "MODEVASION", ListaClases(ClaseOponente), tmp)
                
            Case 5
    
                If ClaseAtacante = 0 Then
                    MsgBox "Debes asignar las clases para ejecutar esta acción."
                    Exit Sub
                End If
                
                Call cmdAplicar_Click(Index)
                tmp = CStr(ModClase(ClaseOponente).AtaqueArmas)
                tmp = Replace(tmp, ",", ".")
                Call WriteVar(PathAux, "MODEVASION", ListaClases(ClaseOponente), tmp)
                
        End Select
        
    End If
    
    lblcargando.Caption = "Guardado con éxito. (" & Time & ")"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    prgRun = False
    
End Sub


Private Sub lstDañoExtraAtacante_Click()

    Dim NumObj As Integer
    
    NumObj = Val(lstDañoExtraAtacante.List(lstDañoExtraAtacante.ListIndex))
    
    If NumObj <= 1 Then
        RefreshPicView
        Exit Sub
    End If
    
    PicView.Cls
    
    Call DrawGrhtoHdc(PicView.hDC, ObjData(NumObj).Grhindex, 0, 0)
    
    PicView.Refresh
    
    Me.lblNombre.Caption = ObjData(NumObj).Name
    Me.lblNivel.Caption = "Nivel minimo: " & ObjData(NumObj).Nivel
    Me.lblDef.Caption = "Defensa: " & ObjData(NumObj).MinDef & "/" & ObjData(NumObj).MaxDef
    Me.lblDaño.Caption = "Daño:" & ObjData(NumObj).MinHit & "/" & ObjData(NumObj).MaxHit

End Sub

Private Sub lstDefensaExtraOponente_Click()
 
    Dim NumObj As Integer
    
    NumObj = Val(lstDefensaExtraOponente.List(lstDefensaExtraOponente.ListIndex))
    
    If NumObj <= 1 Then
        RefreshPicView
        Exit Sub
    End If
    
    PicView.Cls
    
    Call DrawGrhtoHdc(PicView.hDC, ObjData(NumObj).Grhindex, 0, 0)
    
    PicView.Refresh
    
    Me.lblNombre.Caption = ObjData(NumObj).Name
    Me.lblNivel.Caption = "Nivel minimo: " & ObjData(NumObj).Nivel
    Me.lblDef.Caption = "Defensa: " & ObjData(NumObj).MinDef & "/" & ObjData(NumObj).MaxDef
    Me.lblDaño.Caption = "Daño:" & ObjData(NumObj).MinHit & "/" & ObjData(NumObj).MaxHit


End Sub

Private Sub RecTxt_Change()
nLineas = SendMessageLong(frmMain.RecTxt.hwnd, &HBA, 0&, 0&)

If nLineas = 100 Then Call cmdClean_Click

End Sub

Public Sub cmdAttack_Click()

If Me.lstRazaAtacante.List(lstRazaAtacante.ListIndex) = "" Or lstRazaOponente.List(lstRazaOponente.ListIndex) = "" Then
    MsgBox "La raza del contrincante o del atacante son inválidos."
    Exit Sub
End If

If lstClaseAtacante.List(lstClaseAtacante.ListIndex) = "" Or lstClaseOponente.List(lstClaseOponente.ListIndex) = "" Then
    MsgBox "La clase del contrincante o del atacante son inválidos."
    Exit Sub
End If

If Me.txtNivelAtacante.Text = "" Or txtNivelOponente.Text = "" Then
    MsgBox "El nivel del contrincante o del atacante son inválidos."
    Exit Sub
End If

'Skill
If Me.txtWrestling.Text = "" Or Me.txtCombate.Text = "" Or Me.txtDefensa.Text = "" Or Me.txtTacticas.Text = "" Or Me.txtApuñalarAtacante = "" Then
    MsgBox "Los skills a asignar son inválidos."
    Exit Sub
End If

If Me.txtFuerzaAtacante.Text = "" Or Me.txtAgilidadAtacante.Text = "" Or Me.txtAgilidadOponente.Text = "" Then
    MsgBox "Los atributos son inválidos"
    Exit Sub
End If


'Agregamos los flags

Dim tmpDaño As String

Atacante.agilidad = Val(txtAgilidadAtacante.Text)
Atacante.fuerza = Val(Me.txtFuerzaAtacante.Text)
Atacante.Arma = Val(lstArmasAtacante.List(lstArmasAtacante.ListIndex))
Atacante.Nivel = Val(Me.txtNivelAtacante.Text)
Atacante.SkillCombateConArmas = Val(Me.txtCombate.Text)
Atacante.SkillWrestling = Val(Me.txtWrestling.Text)
Atacante.SkillApuñalar = Val(Me.txtApuñalarAtacante.Text)
Atacante.clase = Val(ObtenerClase(Me.lstClaseAtacante.List(Me.lstClaseAtacante.ListIndex)))
tmpDaño = ObtenerHit(Atacante.Nivel, lstRazaAtacante.List(lstRazaAtacante.ListIndex), lstClaseAtacante.List(lstClaseAtacante.ListIndex))
Atacante.MinHit = Val(ReadField(1, tmpDaño, Asc("/")))
Atacante.MaxHit = Val(ReadField(2, tmpDaño, Asc("/")))
Atacante.Raza = Val(ObtenerRaza(Me.lstRazaAtacante.List(Me.lstRazaAtacante.ListIndex)))
Atacante.dañoextra = Val(Me.lstDañoExtraAtacante.List(lstDañoExtraAtacante.ListIndex))

'Oponente
Oponente.agilidad = Val(Me.txtAgilidadOponente.Text)
Oponente.Armadura = Val(Me.ListaArmadurasOponente.List(ListaArmadurasOponente.ListIndex))
Oponente.Casco = Val(Me.lstCascoOponente.List(lstCascoOponente.ListIndex))
Oponente.clase = Val(ObtenerClase(Me.lstClaseOponente.List(Me.lstClaseOponente.ListIndex)))
Oponente.Escudo = Val(Me.lstEscudoOponente.List(lstEscudoOponente.ListIndex))
Oponente.Nivel = Val(Me.txtNivelOponente.Text)
Oponente.Raza = Val(ObtenerRaza(Me.lstRazaOponente.List(Me.lstRazaOponente.ListIndex)))
Oponente.SkillDefensaConEscudos = Val(Me.txtDefensa.Text)
Oponente.SkillTacticas = Val(Me.txtTacticas.Text)
Oponente.DefensaExtra = Val(Me.lstDefensaExtraOponente.List(lstDefensaExtraOponente.ListIndex))

Call UsuarioAtacaUsuario

End Sub
Private Sub lstRazaAtacante_Click()

If lstRazaAtacante.ListIndex <= 0 Or lstRazaAtacante.ListIndex > 6 Then Exit Sub

Dim tmpValor As Integer

tmpValor = Val(Me.txtNivelAtacante.Text)

If Len(lstRazaAtacante.List(lstRazaAtacante.ListIndex)) > 2 And Len(lstClaseAtacante.List(lstClaseAtacante.ListIndex)) > 2 Then
    Label9.Caption = ObtenerHit(tmpValor, lstRazaAtacante.List(lstRazaAtacante.ListIndex), lstClaseAtacante.List(lstClaseAtacante.ListIndex))
End If

End Sub

Private Sub lstClaseAtacante_Click()

If lstClaseAtacante.ListIndex <= 0 Or lstClaseAtacante.ListIndex > 18 Then Exit Sub

Dim tmpValor As Integer

tmpValor = Val(Me.txtNivelAtacante.Text)

If Len(lstRazaAtacante.List(lstRazaAtacante.ListIndex)) > 2 And Len(lstClaseAtacante.List(lstClaseAtacante.ListIndex)) > 2 Then
    Label9.Caption = ObtenerHit(tmpValor, lstRazaAtacante.List(lstRazaAtacante.ListIndex), lstClaseAtacante.List(lstClaseAtacante.ListIndex))
End If

Call RefreshListaBalanceAtacante

End Sub


Private Sub LstArmasAtacante_Click()
    
    Dim NumObj As Integer
    
    NumObj = Val(lstArmasAtacante.List(lstArmasAtacante.ListIndex))
    
    If NumObj <= 1 Then
        RefreshPicView
        Exit Sub
    End If
    
    PicView.Cls
    
    Call DrawGrhtoHdc(PicView.hDC, ObjData(NumObj).Grhindex, 0, 0)
    
    PicView.Refresh
    
    Me.lblNombre.Caption = ObjData(NumObj).Name
    Me.lblNivel.Caption = "Nivel minimo: " & ObjData(NumObj).Nivel
    Me.lblDef.Caption = "Defensa: " & ObjData(NumObj).MinDef & "/" & ObjData(NumObj).MaxDef
    Me.lblDaño.Caption = "Daño:" & ObjData(NumObj).MinHit & "/" & ObjData(NumObj).MaxHit
    
End Sub


Private Sub txtNivelAtacante_Change()

Dim tmpValor As Integer

tmpValor = Val(Me.txtNivelAtacante.Text)


If tmpValor > 50 Then
    tmpValor = 50
ElseIf tmpValor < 1 Then
    tmpValor = 1
End If

txtNivelAtacante.Text = CStr(tmpValor)

If Len(lstRazaAtacante.List(lstRazaAtacante.ListIndex)) > 2 And Len(lstClaseAtacante.List(lstClaseAtacante.ListIndex)) > 2 Then
    Label9.Caption = ObtenerHit(tmpValor, lstRazaAtacante.List(lstRazaAtacante.ListIndex), lstClaseAtacante.List(lstClaseAtacante.ListIndex))
End If

End Sub
Private Sub txtNivelAtacante_KeyPress(KeyAscii As Integer)
 
Dim tmpValor As Integer

tmpValor = Val(Me.txtNivelAtacante.Text)

If KeyAscii = 43 Then
    tmpValor = tmpValor + 1
ElseIf KeyAscii = 45 Then
    tmpValor = tmpValor - 1
End If

If tmpValor > 50 Then
    tmpValor = 50
ElseIf tmpValor < 1 Then
    tmpValor = 1
End If

If (KeyAscii <> 8) Then

    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
    
End If
 
txtNivelAtacante.Text = CStr(tmpValor)

End Sub

Private Sub txtCombate_Change()
If Val(txtCombate.Text) > 100 Then
    txtCombate.Text = 100
ElseIf Val(txtCombate.Text) < 1 Then
    txtCombate.Text = 1
End If

End Sub
Private Sub txtWrestling_Change()
If Val(txtWrestling.Text) > 100 Then
    txtWrestling.Text = 100
ElseIf Val(txtWrestling.Text) < 1 Then
    txtWrestling.Text = 1
End If

End Sub
Private Sub txtApuñalarAtacante_Change()
If Val(txtApuñalarAtacante.Text) > 100 Then
    txtApuñalarAtacante.Text = 100
ElseIf Val(txtApuñalarAtacante.Text) < 1 Then
    txtApuñalarAtacante.Text = 1
End If

End Sub
Private Sub txtAgilidadAtacante_change()

Dim agilidad As Byte

agilidad = Val(Me.txtAgilidadAtacante.Text)
If agilidad > 35 Then
    agilidad = 35
ElseIf agilidad < 1 Then
    agilidad = 1
End If

txtAgilidadAtacante.Text = CStr(agilidad)

End Sub
Private Sub txtAgilidadAtacante_KeyPress(KeyAscii As Integer)
 
Dim agilidad As Byte
agilidad = Val(txtAgilidadAtacante.Text)

If KeyAscii = 43 Then
    agilidad = agilidad + 1
ElseIf KeyAscii = 45 Then
    agilidad = agilidad - 1
End If

If agilidad > 35 Then
    agilidad = 35
ElseIf agilidad < 1 Then
    agilidad = 1
End If

If (KeyAscii <> 8) Then

    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
    
End If
 
txtAgilidadAtacante.Text = CStr(agilidad)

End Sub
Private Sub txtFuerzaAtacante_change()

Dim fuerza As Byte

fuerza = Val(Me.txtFuerzaAtacante.Text)

If fuerza > 35 Then
    fuerza = 35
ElseIf fuerza < 1 Then
    fuerza = 1
End If
 
txtFuerzaAtacante.Text = CStr(fuerza)

End Sub
Private Sub txtFuerzaAtacante_KeyPress(KeyAscii As Integer)
 
Dim fuerza As Byte

fuerza = Val(txtFuerzaAtacante.Text)

If KeyAscii = 43 Then
    fuerza = fuerza + 1
ElseIf KeyAscii = 45 Then
    fuerza = fuerza - 1
End If

If fuerza > 35 Then
    fuerza = 35
ElseIf fuerza < 1 Then
    fuerza = 1
End If
 
If (KeyAscii <> 8) Then

    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
    
End If

txtFuerzaAtacante.Text = CStr(fuerza)

End Sub
Private Sub lstRazaOponente_Click()

If lstRazaOponente.ListIndex <= 0 Or lstRazaOponente.ListIndex > 6 Then Exit Sub

End Sub

Private Sub lstClaseOponente_Click()

If lstClaseOponente.ListIndex <= 0 Or lstClaseOponente.ListIndex > 18 Then Exit Sub

Call RefreshListaBalanceOponente

End Sub
Private Sub lstCascoOponente_Click()
 
    Dim NumObj As Integer
    
    NumObj = Val(lstCascoOponente.List(lstCascoOponente.ListIndex))
    
    If NumObj <= 1 Then
        RefreshPicView
        Exit Sub
    End If
    
    PicView.Cls
    
    Call DrawGrhtoHdc(PicView.hDC, ObjData(NumObj).Grhindex, 0, 0)
    PicView.Refresh
    
    Me.lblNombre.Caption = ObjData(NumObj).Name
    Me.lblNivel.Caption = "Nivel minimo: " & ObjData(NumObj).Nivel
    Me.lblDef.Caption = "Defensa: " & ObjData(NumObj).MinDef & "/" & ObjData(NumObj).MaxDef
    Me.lblDaño.Caption = "Daño:" & ObjData(NumObj).MinHit & "/" & ObjData(NumObj).MaxHit
    
End Sub

Private Sub listaArmadurasOponente_Click()

    Dim NumObj As Integer
    
    NumObj = Val(ListaArmadurasOponente.List(ListaArmadurasOponente.ListIndex))
    
    If NumObj <= 1 Then
        RefreshPicView
        Exit Sub
    End If
    
    PicView.Cls
    
    Call DrawGrhtoHdc(PicView.hDC, ObjData(NumObj).Grhindex, 0, 0)
    PicView.Refresh
    
    Me.lblNombre.Caption = ObjData(NumObj).Name
    Me.lblNivel.Caption = "Nivel minimo: " & ObjData(NumObj).Nivel
    Me.lblDef.Caption = "Defensa: " & ObjData(NumObj).MinDef & "/" & ObjData(NumObj).MaxDef
    Me.lblDaño.Caption = "Daño:" & ObjData(NumObj).MinHit & "/" & ObjData(NumObj).MaxHit
    
End Sub
Private Sub lstEscudoOponente_Click()
 
    Dim NumObj As Integer
    
    NumObj = Val(lstEscudoOponente.List(lstEscudoOponente.ListIndex))
    
    If NumObj <= 1 Then
        RefreshPicView
        Exit Sub
    End If
    
    PicView.Cls
    
    Call DrawGrhtoHdc(PicView.hDC, ObjData(NumObj).Grhindex, 0, 0)
    PicView.Refresh
    
    Me.lblNombre.Caption = ObjData(NumObj).Name
    Me.lblNivel.Caption = "Nivel minimo: " & ObjData(NumObj).Nivel
    Me.lblDef.Caption = "Defensa: " & ObjData(NumObj).MinDef & "/" & ObjData(NumObj).MaxDef
    Me.lblDaño.Caption = "Daño:" & ObjData(NumObj).MinHit & "/" & ObjData(NumObj).MaxHit
    
End Sub
Private Sub txtNivelOponente_Change()

Dim tmpValor As Integer

tmpValor = Val(Me.txtNivelOponente.Text)

If tmpValor > 50 Then
    tmpValor = 50
ElseIf tmpValor < 1 Then
    tmpValor = 1
End If

txtNivelOponente.Text = CStr(tmpValor)

End Sub
Private Sub txtNivelOponente_KeyPress(KeyAscii As Integer)
 
Dim tmpValor As Integer

tmpValor = Val(Me.txtNivelOponente.Text)

If KeyAscii = 43 Then
    tmpValor = tmpValor + 1
ElseIf KeyAscii = 45 Then
    tmpValor = tmpValor - 1
End If

If tmpValor > 50 Then
    tmpValor = 50
ElseIf tmpValor < 1 Then
    tmpValor = 1
End If

If (KeyAscii <> 8) Then

    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
    
End If
 
txtNivelOponente.Text = CStr(tmpValor)

End Sub
Private Sub txtAgilidadOponente_change()

Dim agilidad As Byte

agilidad = Val(Me.txtAgilidadOponente.Text)

If agilidad > 35 Then
    agilidad = 35
ElseIf agilidad < 1 Then
    agilidad = 1
End If

txtAgilidadOponente.Text = CStr(agilidad)

End Sub
Private Sub txtAgilidadOponente_KeyPress(KeyAscii As Integer)
 
Dim agilidad As Byte
agilidad = Val(txtAgilidadOponente.Text)

If KeyAscii = 43 Then
    agilidad = agilidad + 1
ElseIf KeyAscii = 45 Then
    agilidad = agilidad - 1
End If

If agilidad > 35 Then
    agilidad = 35
ElseIf agilidad < 1 Then
    agilidad = 1
End If

If (KeyAscii <> 8) Then

    If (KeyAscii <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
    
End If
 
txtAgilidadOponente.Text = CStr(agilidad)

End Sub

Private Sub txtTacticas_Change()
If Val(txtTacticas.Text) > 100 Then
    txtTacticas.Text = 100
ElseIf Val(txtTacticas.Text) < 1 Then
    txtTacticas.Text = 1
End If

End Sub
Private Sub txtDefensa_Change()
If Val(txtDefensa.Text) > 100 Then
    txtDefensa.Text = 100
ElseIf Val(txtDefensa.Text) < 1 Then
    txtDefensa.Text = 1
End If

End Sub
 
