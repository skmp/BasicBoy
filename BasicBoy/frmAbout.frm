VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About BasicBoy..."
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   345
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   2055
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1320
      Width           =   5415
   End
   Begin VB.Label Label5 
      Caption         =   "If you like this Program then VOTE for it ;)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   960
      MouseIcon       =   "frmAbout.frx":058A
      TabIndex        =   6
      Top             =   3360
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "drk||Raziel:Most of the emu/Project leader  Xeon:Gui,Savestages,beta testing"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Thanks to:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Coders:"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "BasicBoy v[version]"
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmAbout.frx":295C
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    Call InitCommonControls
    Label1.Caption = "BasicBoy v" & App.Major & "." & App.Minor & "." & App.Revision
    Text1.Text = " • Many thanks to xeon (from xeons.net), He ""forced"" me to write the new sound system ,he made the new GUI/Logo and also made the new website" & vbNewLine & _
                 " • To Christopher for command line rom loading and the some other ideas" & vbNewLine & _
                 " • No$gbc (For the pan docs)" & vbNewLine & _
                 " • The Author of VisBoy (This begun as a 'Mod' of his emulator)" & vbNewLine & _
                 " • The writers of GameBoy Cpu Manual" & vbNewLine & _
                 " • The no Doevents tutorial writer" & vbNewLine & _
                 " • Frenzied-Panda for his Common Dialog Class" & vbNewLine & _
                 " • Black Tornado for his Trainer Maker kit (Part of it used for the link emulation)" & vbNewLine & _
                 " • Roja for his sugestions and his amazing code fixer" & vbNewLine & _
                 " • Emuhost.com for hosting the new site" & vbNewLine & _
                 " • www.upx.com for their exelent exe packer(60% smaller exe now)"
End Sub

Private Sub Label5_Click()
    Call Shell(pschome, vbMaximizedFocus)
End Sub
