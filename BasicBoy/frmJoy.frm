VERSION 5.00
Begin VB.Form frmJoy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure keys"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "frmJoy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdj 
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
      Index           =   9
      Left            =   3120
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Index           =   8
      Left            =   3120
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdj 
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
      Index           =   7
      Left            =   3120
      TabIndex        =   15
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Index           =   6
      Left            =   3120
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Left            =   1080
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdj 
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
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Fast Disable"
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
      Index           =   9
      Left            =   2160
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fast Enable"
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
      Index           =   8
      Left            =   2160
      TabIndex        =   19
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Select"
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
      Index           =   7
      Left            =   2160
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
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
      Index           =   6
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "B button"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "A button"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Right"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Left"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Down"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Up:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmJoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim keys As clsDik2, done As Long, wk As Long
Sub confkeys()
Me.Show
done = 0
wk = -1
Set keys = New clsDik2
keys.Startup frmMain.dih, Me.hwnd
cmdj(0).Caption = keys.KeyName(Up)
cmdj(1).Caption = keys.KeyName(Dn)
cmdj(2).Caption = keys.KeyName(Lf)
cmdj(3).Caption = keys.KeyName(Rg)
cmdj(4).Caption = keys.KeyName(ABut)
cmdj(5).Caption = keys.KeyName(BBut)
cmdj(6).Caption = keys.KeyName(St1)
cmdj(7).Caption = keys.KeyName(Sl1)
cmdj(8).Caption = keys.KeyName(SpeedKeyD)
cmdj(9).Caption = keys.KeyName(SpeedKeyU)
Do
DoEvents
keys.Check_Keyboard
Loop While done = 0
Me.Hide
Set keys = Nothing
End Sub
Public Sub ckeydown(key As Long)
If wk > -1 Then
Select Case wk
Case 0 'up
Up = key
Case 1 'dn
Dn = key
Case 2 'lf
Lf = key
Case 3 'rg
Rg = key
Case 4 'abut
ABut = key
Case 5 'bbut
BBut = key
Case 6 'st1
St1 = key
Case 7 'sl1
Sl1 = key
Case 8 'speedkeyd
SpeedKeyD = key
Case 9 'speedkeyu
SpeedKeyU = key
End Select
SaveSetting "BasicBoy", "Joy", "up", Up
SaveSetting "BasicBoy", "Joy", "dn", Dn
SaveSetting "BasicBoy", "Joy", "lf", Lf
SaveSetting "BasicBoy", "Joy", "rg", Rg
SaveSetting "BasicBoy", "Joy", "ab", ABut
SaveSetting "BasicBoy", "Joy", "bb", BBut
SaveSetting "BasicBoy", "Joy", "st1", St1: SaveSetting "BasicBoy", "Joy", "st2", St1: SaveSetting "BasicBoy", "Joy", "st3", St1
SaveSetting "BasicBoy", "Joy", "sl1", Sl1: SaveSetting "BasicBoy", "Joy", "sl2", Sl1
SaveSetting "BasicBoy", "Joy", "spdd", SpeedKeyD: SaveSetting "BasicBoy", "Joy", "spdu", SpeedKeyU
cmdj(0).Caption = keys.KeyName(Up)
cmdj(1).Caption = keys.KeyName(Dn)
cmdj(2).Caption = keys.KeyName(Lf)
cmdj(3).Caption = keys.KeyName(Rg)
cmdj(4).Caption = keys.KeyName(ABut)
cmdj(5).Caption = keys.KeyName(BBut)
cmdj(6).Caption = keys.KeyName(St1)
cmdj(7).Caption = keys.KeyName(Sl1)
cmdj(8).Caption = keys.KeyName(SpeedKeyD)
cmdj(9).Caption = keys.KeyName(SpeedKeyU)
wk = -1
End If
End Sub

Private Sub cmdj_Click(Index As Integer)
wk = Index
cmdj(Index).Caption = "..."
End Sub

Private Sub Command1_Click()
done = 1
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub

Private Sub Form_Load()
Call InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
done = 1
End Sub
