Attribute VB_Name = "bbAC"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)



'v2.0.1
'Link emulation ...
'Almost full emulation (no speed limitation)
'comments added

'Sory for my bad english ...

Option Explicit
Public LinkState As Long
Public TdataB(1) As Byte, tmp As Long, id As String, t2(1) As Byte, Bs As Long, sent As Boolean, tmpdat(255) As Byte
Sub Con() ' connect
Dim tset As Long
tset = GetSetting("BasicBoy", "link", "COP", "0")
If tset And 1 Then 'we are at slot 1
If tset = 3 Then Exit Sub
tset = 3: SaveSetting "BasicBoy", "link", "LID2", frmMain.hwnd
          SaveSetting "BasicBoy", "link", "ptr2", VarPtr(TdataB(0))
LinkState = 3
ElseIf tset And 2 Then 'we are at slot 2
tset = 3: SaveSetting "BasicBoy", "link", "LID1", frmMain.hwnd
          SaveSetting "BasicBoy", "link", "ptr1", VarPtr(TdataB(0))
LinkState = 2
Else 'well, we can chose slot 1 or 2
tset = 1: SaveSetting "BasicBoy", "link", "LID1", frmMain.hwnd
          SaveSetting "BasicBoy", "link", "ptr1", VarPtr(TdataB(0))
LinkState = 2
End If
SaveSetting "BasicBoy", "link", "COP", tset
End Sub
Sub check_link_connection()
If GetSetting("BasicBoy", "link", "COP", "0") = 3 And LinkState > 1 Then
Select Case LinkState
Case 2
init GetSetting("BasicBoy", "link", "LID2"), 2, GetSetting("BasicBoy", "link", "ptr2")
frmMain.Caption = frmMain.Caption & "*Conected (BB1)*"
LinkState = 1
Case 3
init GetSetting("BasicBoy", "link", "LID1"), 2, GetSetting("BasicBoy", "link", "ptr1")
frmMain.Caption = frmMain.Caption & "*Conected (BB2)*"
LinkState = 1
End Select
End If
End Sub
Sub link_kill()
SaveSetting "BasicBoy", "link", "LID1", 0
SaveSetting "BasicBoy", "link", "ptr1", 0
SaveSetting "BasicBoy", "link", "LID2", 0
SaveSetting "BasicBoy", "link", "ptr2", 0
SaveSetting "BasicBoy", "link", "COP", 0
End Sub

Sub Check() 'check the link state
If TdataB(0) = 1 Then ' is this a recieve msg??
    'send the data
    frmMain.SendData2 Chr$(9) & Chr$(TdataB(1))
    'reset the data
    TdataB(0) = 0
    TdataB(1) = 0
End If
End Sub



