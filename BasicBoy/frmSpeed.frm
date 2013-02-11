VERSION 5.00
Begin VB.Form frmSpeed 
   Caption         =   "Frame Delay Configuration"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmCpu 
      Caption         =   "Cpu Over/Under-clock"
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4335
      Begin VB.HScrollBar sSpeed 
         Height          =   255
         Left            =   120
         Max             =   200
         Min             =   1
         TabIndex        =   10
         Top             =   360
         Value           =   100
         Width           =   4095
      End
      Begin VB.Label Label7 
         Caption         =   $"frmSpeed.frx":0000
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   3975
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   2160
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Label Label6 
         Caption         =   "100%"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Speed : "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.HScrollBar fde 
      Height          =   255
      LargeChange     =   10
      Left            =   240
      Max             =   6400
      Min             =   1
      TabIndex        =   4
      Top             =   600
      Value           =   1600
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "restore"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4200
      Width           =   4335
   End
   Begin VB.CheckBox chkTgt 
      Caption         =   "Do not use QueryPerformaceCounter(if you have speed problems, try it)"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Use GetTickCount Insted of QueryPerformaceCounter.Try it if you have speed problems."
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Frame frmDG 
      Caption         =   "Frame Delay"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.Label Label1 
         Caption         =   "Normal value : 16.00,Current :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "15.00 ->"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   " <- 16.00"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "16.00"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Line Line4 
         X1              =   1245
         X2              =   1245
         Y1              =   360
         Y2              =   840
      End
      Begin VB.Line Line3 
         X1              =   1320
         X2              =   1320
         Y1              =   360
         Y2              =   840
      End
   End
   Begin VB.Line Line2 
      X1              =   1480
      X2              =   1470
      Y1              =   240
      Y2              =   720
   End
End
Attribute VB_Name = "frmSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'To configure the frame limiting mode/delay
'Added on v2.0.1
Option Explicit

Private Sub chkTgt_Click()
SaveSetting "BasicBoy", "misc", "timermode", Abs(chkTgt.value)
gtc = Abs(chkTgt.value)
If gtc Then
Label1.Caption = "Normal value : 15.00,Current :"
Else
Label1.Caption = "Normal value : 16.00,Current :"
End If
End Sub

Private Sub Command1_Click()
'since gettickcount has lower resolution(15 ms)
'will round 16 to 30 ...so use 15
If gtc Then
fde.value = 1500
Else
fde.value = 1600
End If
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub fde_Change()
framedelay = fde.value / 100
Label2.Caption = format$(framedelay, "00.00")
End Sub

Private Sub fde_Scroll()
fde_Change
End Sub

Private Sub Form_Load()
If gtc Then
Label1.Caption = "Normal value : 15.00,Current :"
Else
Label1.Caption = "Normal value : 16.00,Current :"
End If
fde.value = framedelay * 100
chkTgt.value = gtc
sSpeed.value = (1 / Cpu_Speed) * 100
End Sub

Private Sub sSpeed_Change()
Cpu_Speed = 1 / (sSpeed.value / 100)
Label6.Caption = format$(1 / Cpu_Speed, "000%")
InitCPU
End Sub

Private Sub sSpeed_Scroll()
sSpeed_Change
End Sub
