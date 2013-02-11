VERSION 5.00
Begin VB.Form frmCheat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ram Cheats"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9735
   Icon            =   "frmCheat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Rescmd 
      Caption         =   "Resume"
      Height          =   345
      Left            =   7560
      TabIndex        =   11
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   345
      Left            =   9000
      TabIndex        =   10
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdCo 
      Caption         =   "Change offset"
      Height          =   345
      Left            =   7560
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CheckBox chkck 
      Caption         =   "Use cheats"
      Height          =   255
      Left            =   7680
      TabIndex        =   8
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ListBox lsttm 
      Height          =   2010
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
   End
   Begin VB.ListBox lstAdr 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   7215
   End
   Begin VB.CommandButton cmdLod 
      Caption         =   "Load List"
      Height          =   345
      Left            =   7560
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmdSav 
      Caption         =   "Save List"
      Height          =   345
      Left            =   7560
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Restart"
      Height          =   345
      Left            =   7680
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find the value"
      Height          =   345
      Left            =   7680
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtVal 
      Height          =   285
      Left            =   7680
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox cmbsiz 
      Height          =   315
      ItemData        =   "frmCheat.frx":058A
      Left            =   7680
      List            =   "frmCheat.frx":0597
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmCheat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type cheat
Adr As Long
Siz As Byte
val(3) As Byte
Frz As Boolean
rb As Long
End Type
Dim fs As Boolean
Dim cheats() As cheat
Dim tcheats() As cheat
Dim uch As Boolean

Private Sub chkck_Click()
uch = chkck.value
End Sub

Private Sub cmdCo_Click()
On Error GoTo to1
cheats(lstAdr.ListIndex).Adr = val(InputBox("Give a new address (old was " & cheats(lstAdr.ListIndex).Adr & " ) ", "Cheats"))
UpdateList lstAdr, cheats
Exit Sub
to1:
End Sub

Private Sub cmdCopy_Click()
On Error GoTo to1
ReDim Preserve cheats(UBound(cheats) + 1)
cheats(UBound(cheats)) = cheats(lstAdr.ListIndex)
UpdateList lstAdr, cheats
Exit Sub
to1:
End Sub

Private Sub cmdFind_Click()
Dim i As Long, csiz As Byte, wval(3) As Byte, ti As Long, rb As Long
On Error Resume Next
ReDim tcheats(99999)
csiz = cmbsiz.ListIndex
If csiz = 1 Then
wval(1) = val(txtVal.Text) \ 256: wval(0) = val(txtVal.Text) And 255
Else
wval(0) = val(txtVal.Text) And 255
End If
If csiz = 0 Then
For rb = 0 To 7
For i = LBound(RAM) To UBound(RAM)
If readHackM(i, rb) = wval(0) Then
tcheats(ti).Adr = i
tcheats(ti).rb = rb
tcheats(ti).val(0) = readHackM(i, rb)
tcheats(ti).Siz = csiz
ti = ti + 1
End If
Next i
Next rb
ElseIf csiz = 1 Then
For rb = 0 To 7
For i = LBound(RAM) To UBound(RAM) Step 2
If readHackM(i, rb) = wval(1) And readHackM(i + 1, rb) = wval(0) Then
tcheats(ti).Adr = i
tcheats(ti).val(0) = readHackM(i, rb): tcheats(ti).val(1) = readHackM(i + 1, rb)
tcheats(ti).Siz = csiz
tcheats(ti).rb = rb
ti = ti + 1
End If
Next i
Next rb
End If
If ti = 0 Then ti = 1
ReDim Preserve tcheats(ti - 1)
UpdateList lsttm, tcheats
End Sub

Private Sub cmdLod_Click()
Dim tmp As Long, mt As String
mt = CD.ShowOpen(Me.hwnd, "Select a Cheat List to load", CD.FileName, "Cheat Files (*.clf)|*.clf")
If Len(mt) < 1 Then Exit Sub
Open mt For Binary As #1
Get #1, , tmp
ReDim cheats(tmp)
Get #1, , cheats
Close #1
UpdateList lstAdr, cheats
End Sub

Private Sub cmdNew_Click()
fs = True
lstAdr.Clear
End Sub

Private Sub cmdSav_Click()
Dim tmp As Long, mt As String
On Error Resume Next

mt = CD.ShowSave(Me.hwnd, "Select a name to save the Cheat List", CD.FileName, "Cheat Files (*.clf)|*.clf", "")
If Len(mt) < 1 Then Exit Sub
Open mt$ For Binary As #1
tmp = UBound(cheats)
Put #1, , tmp
Put #1, , cheats
Close #1
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub

Private Sub Form_Load()
Call InitCommonControls
cmbsiz.ListIndex = 0
End Sub

Private Sub UpdateList(list As ListBox, cheats() As cheat)
Dim i As Long
list.Clear
For i = 0 To UBound(cheats)
If cheats(i).Siz = 0 Then list.AddItem cheats(i).Adr & "," & cheats(i).rb & " :" & cheats(i).val(0), i Else list.AddItem cheats(i).Adr & "," & cheats(i).rb & ":" & CLng(cheats(i).val(0)) * 256 + cheats(i).val(1), i
list.Selected(i) = cheats(i).Frz
Next i
End Sub

Private Sub lstAdr_DblClick()
Dim tmp As Long
tmp = InputBox("Give a value")
If cheats(lstAdr.ListIndex).Siz = 2 Then
cheats(lstAdr.ListIndex).val(0) = tmp \ 256: cheats(lstAdr.ListIndex).val(1) = tmp And 255
Else
cheats(lstAdr.ListIndex).val(0) = tmp And 255
End If
UpdateList lstAdr, cheats
End Sub

Private Sub lsttm_DblClick()
On Error GoTo to1
ReDim Preserve cheats(UBound(cheats) + 1)
res1:
cheats(UBound(cheats)) = tcheats(lsttm.ListIndex)
UpdateList lstAdr, cheats
Exit Sub
to1:
ReDim Preserve cheats(0)
GoTo res1
End Sub
Public Sub ChkCheats()
Dim i As Long
On Error GoTo st0
If uch Then
For i = 0 To UBound(cheats)
If cheats(i).Siz = 1 Then
wHackM cheats(i).Adr, cheats(i).rb, cheats(i).val(0)
wHackM cheats(i).Adr + 1, cheats(i).rb, cheats(i).val(1)
Else
wHackM cheats(i).Adr, cheats(i).rb, cheats(i).val(0)
End If
Next i
End If
st0:
End Sub

Private Sub Rescmd_Click()
Dim tch() As cheat, tch2() As cheat, i As Long, i2 As Long, i3 As Long
tch = tcheats
tch2 = tcheats
cmdFind_Click
For i = 0 To UBound(tch)
For i2 = 0 To UBound(tcheats)
If tch(i).Adr = tcheats(i2).Adr Then
ReDim Preserve tch2(i3)
tch2(i3) = tch(i)
i3 = i3 + 1
End If
Next
Next
tcheats = tch2
UpdateList lsttm, tcheats
End Sub
