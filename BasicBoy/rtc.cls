VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rtc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Well yes this is bad writen
'But i'm getting booooored with Gameboy..
'It's a 5 mins work :)

Option Explicit
Public act As Long
Dim Secs As Long, halt As Long, dccb As Long
Dim dat As Date, tim As Date

Public Property Get S() As Long
update
S = Secs Mod 60
End Property
Public Property Let S(ByVal value As Long)
Secs = Secs - S + value
End Property

Public Property Get M() As Long
update
M = (Secs \ 60) Mod 60
End Property
Public Property Let M(ByVal value As Long)
Secs = Secs - M * 60 + value * 60
End Property

Public Property Get H() As Long
update
H = (Secs \ 3600) Mod 24
End Property
Public Property Let H(ByVal value As Long)
Secs = Secs - H * 3600 + value * 3600
End Property

Public Property Get DL() As Long
update
DL = (Secs / 86400) And 255
End Property
Public Property Let DL(value As Long)
Secs = Secs - DL * 86400 + value * 86400
End Property

Public Property Get DH() As Long
update
DH = (Secs / 22118400) And 1 + ((Secs / 22118400) And 2) * 32 + halt * 32
End Property
Public Property Let DH(value As Long)
Secs = Secs - ((S / 22118400) And 1) * 22118400 + (value And 1) * 22118400
If (value And 64) = 0 Then Secs = Secs - ((S / 22118400) And 2) * 22118400
End Property

Public Sub save(flname As String)
Dim ssv As Long
ssv = FreeFile
Open flname & ".rtc" For Binary As ssv
Put ssv, , Secs
Put ssv, , halt
Put ssv, , dat
Put ssv, , tim
Close ssv
End Sub
Public Sub load(flname As String)
Dim ssv As Long
ssv = FreeFile
Open flname & ".rtc" For Binary As ssv
If LOF(ssv) = 0 Then Exit Sub
Get ssv, , Secs
Get ssv, , halt
Get ssv, , dat
Get ssv, , tim
Close ssv
update
End Sub

Private Sub Class_Initialize()
dat = Date
tim = time
End Sub
Private Sub update()
Dim ssc As Long
ssc = (Date - dat) * 86400
ssc = ssc + (time - tim) * 86400
Secs = Secs + ssc
dat = Date
tim = time
End Sub
Public Function readReg() As Long
Select Case act
    Case &H8
    readReg = S
    Case &H9
    readReg = M
    Case &HA
    readReg = H
    Case &HB
    readReg = DL
    Case &HC
    readReg = DH
End Select
End Function

Public Sub writeReg(val As Long)
Select Case act
    Case &H8
    S = val
    Case &H9
    M = val
    Case &HA
    H = val
    Case &HB
    DL = val
    Case &HC
    DH = val
End Select
End Sub
