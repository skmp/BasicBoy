VERSION 5.00
Begin VB.Form frmDebugger 
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCh 
      Height          =   1035
      Left            =   120
      TabIndex        =   33
      Top             =   5040
      Width           =   2415
   End
   Begin VB.ListBox lstEh 
      Height          =   2205
      Left            =   120
      TabIndex        =   30
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame fraMisk 
      Caption         =   "Misk Variables"
      Height          =   6135
      Left            =   8640
      TabIndex        =   28
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame fraInf 
      Caption         =   "GameBoy/Cart Info"
      Height          =   4695
      Left            =   2640
      TabIndex        =   23
      Top             =   1560
      Width           =   1815
      Begin VB.Label lblZ80S 
         Caption         =   "8 MHz"
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblcpus 
         Caption         =   "CPU speed :"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblGBS 
         Caption         =   "GBC"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblsys 
         Caption         =   "System :"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame fraMMR 
      Caption         =   "Memory Mapped Registers"
      Height          =   6135
      Left            =   4560
      TabIndex        =   22
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame fraRegs 
      Caption         =   "CPU Registers"
      Height          =   1335
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1815
      Begin VB.Label lblCarry 
         Caption         =   "C"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblHalfCarry 
         Caption         =   "H"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblPrevOp 
         Caption         =   "N"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblZero 
         Caption         =   "Z"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblFlags 
         Caption         =   "Flags:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblde 
         Caption         =   "FFFF"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblpc 
         Caption         =   "FFFF"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblhl 
         BackStyle       =   0  'Transparent
         Caption         =   "FFFF"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblsp 
         BackStyle       =   0  'Transparent
         Caption         =   "FFFF"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblaf 
         Caption         =   "FFFF"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label SP 
         Caption         =   "SP :"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.Label PC 
         Caption         =   "PC :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label BC 
         Caption         =   "BC :"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label DE 
         Caption         =   "DE :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label HL 
         Caption         =   "HL :"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.Label af 
         Caption         =   "AF :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblbc 
         BackStyle       =   0  'Transparent
         Caption         =   "FFFF"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9840
      TabIndex        =   3
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run/Stop"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdStep 
      Caption         =   "Step"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   6360
      Width           =   975
   End
   Begin VB.ListBox lstDiss 
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Call History"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label lstst3 
      Caption         =   "Execute History"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lblst2 
      Caption         =   "Disassembly :"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public w As Long, den As Long
Sub execcommand()
If den = 0 Then
Me.lblaf = hex2(A * 256 + F)
Me.lblbc = hex2(b * 256 + C)
Me.lblde = hex2(D * 256 + E)
Me.lblhl = hex2(H * 256 + L)
Me.lblpc = hex2(z80.PC)
Me.lblsp = hex2(z80.SP)
Me.lblZero = IIf(getZ, "Z", "-")
Me.lblPrevOp = IIf(GetN, "N", "-")
Me.lblHalfCarry = IIf(getH, "H", "-")
Me.lblCarry = IIf(getC, "C", "-")
Do
DoEvents
Loop While w = 1
If w = 2 Then w = 1
End If
End Sub

Function hex2(ByVal val As Long) As String
hex2 = String(4 - Len(Hex(val)), "0") & Hex(val)
End Function

Private Sub cmdExit_Click()
den = 12
Unload Me
End Sub

Private Sub cmdRun_Click()
If w = 0 Then: w = 2: Else w = 0
End Sub

Private Sub cmdStep_Click()
If w Then w = 2
End Sub

Private Sub Form_Load()
den = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
den = 12
Unload Me
End Sub
