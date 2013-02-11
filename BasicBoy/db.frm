VERSION 5.00
Begin VB.Form frmRomInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ROM Information"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "db.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   345
      Left            =   2280
      TabIndex        =   22
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label GbColor 
      Height          =   195
      Left            =   1680
      TabIndex        =   21
      Top             =   1290
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "ROM Size:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lRas 
      Height          =   390
      Left            =   1680
      TabIndex        =   20
      Top             =   510
      Width           =   3855
   End
   Begin VB.Label lRos 
      Height          =   390
      Left            =   1680
      TabIndex        =   19
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label iName 
      Height          =   195
      Left            =   1680
      TabIndex        =   18
      Top             =   1095
      Width           =   3855
   End
   Begin VB.Label Cart 
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   900
      Width           =   3855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "GameBoy Color:"
      Height          =   195
      Left            =   60
      TabIndex        =   16
      Top             =   1290
      Width           =   1515
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Internal Name:"
      Height          =   195
      Left            =   15
      TabIndex        =   15
      Top             =   1095
      Width           =   1560
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Cart Type:"
      Height          =   195
      Left            =   75
      TabIndex        =   14
      Top             =   900
      Width           =   1500
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "RAM Banks:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   705
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "ROM Banks:"
      Height          =   195
      Left            =   105
      TabIndex        =   12
      Top             =   315
      Width           =   1470
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "RAM Size:"
      Height          =   195
      Left            =   15
      TabIndex        =   11
      Top             =   510
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "E :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "F :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "H :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "L :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "SP :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "PC :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "D :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "B :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "C :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmRomInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    Call InitCommonControls
    Cart.Caption = Ct(rominfo.Ctype)
    iName.Caption = rominfo.Title
    lRos.Caption = Ros(rominfo.romsize) & vbNewLine & Rosn(rominfo.romsize)
    lRas.Caption = Ras(rominfo.ramsize) & vbNewLine & Rasn(rominfo.ramsize)
    GbColor.Caption = GBM
End Sub

