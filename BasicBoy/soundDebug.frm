VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2400
      Top             =   3600
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Dim i As Long
clsc
prt "GB sound regs:"
For i = 65296 To 65318
prt Hex$(i) & " : " & "value = " & RAM(i, 0)
Next i

prt CStr(wave3.MCount)
prt CStr(((1 / (65536 / (2048 - (RAM(65309, 0) + (RAM(65310, 0) And 7) * 256)))) * 44100) / 32)  '32 wave phases

End Sub
Sub prt(str As String)
Label1.Caption = Label1.Caption & str & vbNewLine
End Sub
Sub clsc()
Label1.Caption = ""
End Sub


