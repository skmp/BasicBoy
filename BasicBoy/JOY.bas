Attribute VB_Name = "modJOY"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)



'v3.0.1
'Joypad emulation ...
'Full emulation
'Using DierctInput now
'Moved from the form
'You can configure the keys now
'comments added

'Sory for my bad english ...
Option Explicit

Public Sub KeyDown(KeyCode As Byte)
Dim temp As Long, old As Long
    Select Case KeyCode
        Case Lf  'Left
            joyval1 = joyval1 Or 2
        Case Up    'Up
            joyval1 = joyval1 Or 4
        Case Rg    'Right
            joyval1 = joyval1 Or 1
        Case Dn    'Down
            joyval1 = joyval1 Or 8
        Case ABut     'Z - A Button
            joyval2 = joyval2 Or 1
        Case BBut     'X - B button
            joyval2 = joyval2 Or 2
        Case St1, St2, St3  ' <Enter> - Start
            joyval2 = joyval2 Or 8
        Case Sl1, Sl2    ' <Space> - Select
            joyval2 = joyval2 Or 4
        Case 66
        SrceenShot 'take a screenshot
        'Case -1
        'SS 'Save Stage-not working
        'Case -2
        'LS 'Load Stage-not working
        Case SpeedKeyD
        slfp = lfp
        lfp = False 'Set fullspeed
        If stpsnd Then
        ssnd = snd
        snd = 0
        initWave
        End If
        If stpsk Then
        ofskip = fskip
        ofmode = fmode
        fskip = 20
        fmode = 0
        End If
    End Select
If old <> joyval1 * 16 + joyval2 Then RAM(65295, 0) = RAM(65295, 0) Or 16 'update joy reg
End Sub
Public Sub KeyUp(KeyCode As Byte)
Dim temp As Long, old As Long
old = joyval1 * 16 + joyval2
    Select Case KeyCode
        Case Lf  'Left
            joyval1 = joyval1 And 253
        Case Up    'Up
            joyval1 = joyval1 And 251
        Case Rg   'Right
            joyval1 = joyval1 And 254
        Case Dn   'Down
            joyval1 = joyval1 And 247
        Case ABut    'Z - A Button
            joyval2 = joyval2 And 254
        Case BBut     'X - B button
            joyval2 = joyval2 And 253
        Case St1, St2, St3 ' <Enter> - Start
            joyval2 = joyval2 And 247
        Case Sl1, Sl2  ' <Space> - Select
            joyval2 = joyval2 And 251
        Case SpeedKeyU 'set normal speed
        lfp = slfp
        If stpsnd Then
        snd = ssnd
        initWave
        End If
        If stpsk Then
        If ofskip Then
        fskip = ofskip
        fmode = ofmode
        End If
        End If
    End Select
    If old <> joyval1 * 16 + joyval2 Then RAM(65295, 0) = RAM(65295, 0) Or 16 'update joy reg
End Sub
