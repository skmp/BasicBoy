Attribute VB_Name = "modSoundChip"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)

'Sound Chip:Sound generator and mixer
'The interface with the emulator is in the modeSoundInterface.bas file
'I'm sure that waveform generation is not optimal and that there are
'many bugs..But it works :)

'v1.3.1
'Fisrt Implementation of the idea
'Sound Generation,Command procesing


Option Explicit
Option Base 0
Global ssound As Long
Dim inited As Byte
Dim SoundS As clsSStream
Dim i As Long, i_max As Long, t_chan As Long, i2 As Long
Dim low_i2 As Long
Public Sound_Sync_Pos As Long
Dim buf() As Long, buf_() As Byte, buflenhalf As Long
Dim com() As SoundCommand, cmd_index As Long, cmd_index_max As Long, cmd_index_t As Long
Dim cmd_lo_idx As Long
Public Const Sound_Sync As Double = 95.1089342403628 '4mega/44100
Dim tmp_1 As Long
'***Generation Variables***
Public wave1 As SoundCD12
Public wave2 As SoundCD12
Public wave3 As SoundCD3
Public wave4 As SoundCD4

Dim tmp_vard As Long

'Init everything
Sub Init_Sound(ByVal buflen As Long)
If inited Then
inited = 0
SoundS.ch.Stop
Set SoundS = Nothing
Init_Sound buflen
Else
inited = 1
If buflen = 0 Then buflen = 8
buflen = buflen * 441 '* 2
ReDim buf(buflen)
ReDim buf_(buflen)
ReDim com(0 To buflen * 10)
cmd_index_max = UBound(com)
buflenhalf = buflen / 2
Set SoundS = New clsSStream
SoundS.init buflen
generate 0
generate 1
SoundS.ch.Play DSBPLAY_LOOPING
End If
End Sub
'**Send a command to the sound chip**
Sub send_command(chan As Sound_Chans, cmd As Sound_CMD, param As Sound_Pars)
If snd Then
    If cmd_index > UBound(com) Then ReDim Preserve com(cmd_index + 1000): cmd_index_max = cmd_index + 1000
    com(cmd_index).chan = chan
    com(cmd_index).cmd = cmd
    com(cmd_index).param = param
    com(cmd_index).en = 1
    com(cmd_index).pos = Sound_Sync_Pos / Sound_Sync
    cmd_index = cmd_index + 1
End If
End Sub
'***Generate some samples***
Sub generate(half As Long)
If (snd And gb_snd) = 1 Then
cmd_lo_idx = 0
chan1play buflenhalf - 1, buf: cmd_lo_idx = 0
chan2play buflenhalf - 1, buf: cmd_lo_idx = 0
chan3play buflenhalf - 1, buf: cmd_lo_idx = 0
chan4play buflenhalf - 1, buf
Sound_Sync_Pos = 0
cmd_index = 0
Else
        For cmd_index_t = 0 To cmd_index - 1
        If com(cmd_index_t).en = 1 Then  'command exec
        com(cmd_index_t).en = 0 'command was executed
            If com(cmd_index_t).chan = 1 Then proc_cmd1 com(cmd_index_t)
            If com(cmd_index_t).chan = 2 Then proc_cmd2 com(cmd_index_t)
            If com(cmd_index_t).chan = 3 Then proc_cmd3 com(cmd_index_t)
            If com(cmd_index_t).chan = 4 Then proc_cmd4 com(cmd_index_t)
        End If
        Next cmd_index_t
        cmd_index = 0
End If
mix32d buf, 4, buf_, buflenhalf - 1

If half = 0 Then '0-half
SoundS.ch.writebuffer 0, buflenhalf, buf_(0), DSBLOCK_DEFAULT
Else 'half-end
SoundS.ch.writebuffer buflenhalf, buflenhalf, buf_(0), DSBLOCK_DEFAULT
End If
End Sub
Sub chan1play(Siz As Long, buf() As Long)
    For i = 0 To Siz
        'generate sound chanel 1
        With wave1
        If .Play Then
        buf(i) = buf(i) + .Current * .Volume
        .Count = .Count + 1
        If .Count > CLng(.MCount) Then
        .Count = .Count - .MCount
        .Index = (.Index + 1) Mod 8
        .Current = sqrW(.Index, .Duty)
        End If
        End If
        End With
        
        'Execute any command for this pos
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).pos = i Then
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 1 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd1 com(cmd_index_t)
        End If
        ElseIf com(cmd_index_t).pos > i Then
        GoTo ext:
        End If
        Next cmd_index_t
ext:
        cmd_lo_idx = cmd_index_t
    Next i
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 1 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd1 com(cmd_index_t)
        End If
        Next cmd_index_t

End Sub
Sub chan2play(Siz As Long, buf() As Long)
    For i = 0 To Siz
        'generate sound chanel 2
        With wave2
        If .Play Then
        buf(i) = buf(i) + .Current * .Volume
        .Count = .Count + 1
        If .Count > CLng(.MCount) Then
        .Count = .Count - .MCount
        .Index = (.Index + 1) Mod 8
        .Current = sqrW(.Index, .Duty)
        End If
        End If
        End With
        
        'Exec any command for this sound  pos
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).pos = i Then
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 2 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd2 com(cmd_index_t)
        End If
        ElseIf com(cmd_index_t).pos > i Then
        GoTo ext:
        End If
        Next cmd_index_t
ext:
        cmd_lo_idx = cmd_index_t
    Next i
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 2 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd2 com(cmd_index_t)
        End If
        Next cmd_index_t

        
End Sub
Sub chan3play(Siz As Long, buf() As Long)
   
   For i = 0 To Siz
        'generate sound chanel 3
        With wave3
        If .Play Then
        buf(i) = buf(i) + .Current * .Volume
        .Count = .Count + 1
        If .Count > CLng(wave3.MCount) Then
        .Count = .Count - .MCount
        .Index = (.Index + 1) Mod 32
        .Current = swm(.Waveform(wave3.Index))
        End If
        End If
        End With
                   'Exec any command for this sound  pos
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).pos = i Then
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 3 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd3 com(cmd_index_t)
        End If
        ElseIf com(cmd_index_t).pos > i Then
        GoTo ext:
        End If
        Next cmd_index_t
ext:
        cmd_lo_idx = cmd_index_t
    Next i
    'exec any odly positioned commands
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 3 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd3 com(cmd_index_t)
        End If
        Next cmd_index_t

        
End Sub
Sub chan4play(Siz As Long, buf() As Long)
    For i = 0 To Siz
        'generate sound chanel 4
        If wave4.Play Then
        buf(i) = buf(i) + wave4.Current * wave4.Volume
        wave4.Count = wave4.Count + 1
        If wave4.Count > CLng(wave4.MCount) Then
        wave4.Count = wave4.Count - wave4.MCount
        If wave4.bits = 1 Then
        wave4.Index = (wave4.Index + 1) Mod 128
        wave4.Current = noise7(wave4.Index)
        Else
        wave4.Index = (wave4.Index + 1) Mod 32768
        wave4.Current = noise15(wave4.Index)
        End If
        End If
        End If
        
        'Execute any command for this pos
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).pos = i Then
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 4 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd4 com(cmd_index_t)
        End If
        ElseIf com(cmd_index_t).pos > i Then
        GoTo ext:
        End If
        Next cmd_index_t
ext:
        cmd_lo_idx = cmd_index_t
    Next i
        For cmd_index_t = cmd_lo_idx To cmd_index - 1
        If com(cmd_index_t).en = 1 And com(cmd_index_t).chan = 4 Then 'command exec
                com(cmd_index_t).en = 0 'command was executed
                proc_cmd4 com(cmd_index_t)
        End If
        Next cmd_index_t

End Sub


Sub proc_cmd1(ByRef command As SoundCommand) 'cmd proc for chanel1
Select Case command.cmd
Case 1 ' freq set
command.param = 2048 - command.param
wave1.MCount = command.param * 4.20570373535156E-02  '(0.042057037353515625) '8 wave phases
Case 2 ' Pattern wave duty set
wave1.Duty = command.param
Case 3 ' volume set
wave1.Volume = command.param / 15
Case 13 ' play
wave1.Play = 1
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
Case 14 ' stop
wave1.Play = 0
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
End Select
End Sub

Sub proc_cmd2(ByRef command As SoundCommand)
Select Case command.cmd
Case 1 ' freq set
command.param = 2048 - command.param
wave2.MCount = command.param * 4.20570373535156E-02 '(0.042057037353515625) '8 wave phases
Case 2 ' Pattern wave duty set
wave2.Duty = command.param
Case 3 ' volume set
wave2.Volume = command.param / 15
Case 13 ' play
wave2.Play = 1
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
Case 14 ' stop
wave2.Play = 0
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
End Select
End Sub

Sub proc_cmd3(ByRef command As SoundCommand)
Select Case command.cmd
Case 1 ' freq set
command.param = 2048 - command.param
wave3.MCount = ((1 / (65536 / command.param)) * 44100) / 32 '32 wave phases
Case 5 ' volume set
wave3.Volume = command.param / 256
Case 13 ' play
wave3.Play = 1
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
Case 14 ' stop
wave3.Play = 0
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
Case 15 'write wave
tmp_vard = (command.param \ 512)
wave3.Waveform((command.param And 31) * 2) = tmp_vard
tmp_vard = (command.param \ 32) And 15
wave3.Waveform((command.param And 31) * 2 + 1) = tmp_vard
End Select
End Sub

Sub proc_cmd4(ByRef command As SoundCommand)
Select Case command.cmd
Case 1 ' freq set
If (command.param And 7) = 0 Then
wave4.MCount = 1 / (4194304 * 1 / 2 ^ 3 * 2 * 1 / (2 ^ ((command.param \ 16) + 1))) * 44100
Else
wave4.MCount = 1 / (4194304 * 1 / 2 ^ 3 * 1 / (command.param And 7) * 1 / (2 ^ ((command.param \ 16) + 1))) * 44100
End If
Case 2 ' bits selection
wave4.bits = command.param
Case 3 ' volume set
wave4.Volume = command.param / 15
Case 13 ' play
wave4.Play = 1
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
Case 14 ' stop
wave4.Play = 0
'updateNR52
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
End Select
End Sub

'Helper functions for the software mixer
Sub mix32d(data() As Long, chans As Long, tar() As Byte, upto As Long) '32 bits
Dim temp As Long
For i = 0 To upto
temp = (data(i) \ chans)
If temp > 126 Then
    temp = 126
ElseIf temp < -126 Then
    temp = -126
End If
tar(i) = temp + 127
data(i) = 0
Next i
End Sub

