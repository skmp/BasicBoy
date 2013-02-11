Attribute VB_Name = "modSoundInterface"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)

'v3.0.1
'(almost)Complete rewrite of the Sound interface/emulation
'Emulated Sound chanels:Ch1,Ch2,Ch3,Ch4
'Still missing stereo,flags and volume(nr5x)
'Coments will be added with the next releases

'Sory for my bad english ...
Option Explicit
Global swm(15) As Long
Dim tmp_1 As Long, freq As Double

'Public Whan As Long
Public cl1(3) As Long
Public en1(3) As Byte
Public cl2(3) As Long
Public en2(3) As Byte
Public cl3(3) As Long
Public en3(3) As Byte
Public cl4(3) As Long
Public en4(3) As Byte
Public ClCleft As Long
Public buf1(31) As Byte
Public buf2(31) As Byte
Public buf3(31) As Byte
Public buf4(31) As Byte
Dim tones(4) As Double, valumes(4) As Long, VOLS(7) As Byte
Dim tmp2(7) As Byte, tlol As Byte, Tsnd As Long, tmp3(31) As Byte, i As Long
'Gb=2048-(131072/Hz)
'Hz=131072/(2048-Gb)
Sub update_sound(clc As Long) 'update sound
If snd Then 'sound is enabled
Sound_Sync_Pos = Sound_Sync_Pos + clc
'****Chanel 1****
If en1(0) Then cl1(0) = cl1(0) - clc 'ok

If cl1(0) < 1 And en1(0) Then 'stop
send_command Chanel1, Sound_Stop, No_Par
wave1p = 0
en1(0) = 0
End If


If en1(1) Then cl1(1) = cl1(1) - clc

If cl1(1) < 1 And en1(1) Then 'Volume envelop
Tsnd = (RAM(65298, 0) And 240) \ 16
If RAM(65298, 0) And 8 Then
Tsnd = Tsnd + 1
If Tsnd > 15 Then en1(1) = 0: Tsnd = 15
Else
Tsnd = Tsnd - 1
If Tsnd < 0 Then en1(1) = 0: Tsnd = 0
End If
RAM(65298, 0) = (RAM(65298, 0) And 15) + Tsnd * 16
send_command Chanel1, Wave_Volume_Set, (RAM(65298, 0) And 224) \ 16
cl1(1) = (en1(1) / 64) * 4194304
End If



If en1(2) Then cl1(2) = cl1(2) - clc

If cl1(2) < 1 And en1(2) Then 'Sweep envelop
Tsnd = (RAM(65300, 0) And 7) * 256 + RAM(65299, 0)
If RAM(65296, 0) And 8 Then
Tsnd = Tsnd - Tsnd / 2 ^ (RAM(65296, 0) And 7)
Else
Tsnd = Tsnd + Tsnd / 2 ^ (RAM(65296, 0) And 7)
End If
RAM(65300, 0) = (RAM(65300, 0) And 248) Or (Tsnd \ 256)
RAM(65299, 0) = Tsnd And 255
send_command Chanel1, Wave_Frequency_Set, Tsnd
cl1(2) = ((en1(2) \ 16) / 128) * 4194304
End If



'****Channel 2****
If en2(0) Then cl2(0) = cl2(0) - clc

If cl2(0) < 1 And en2(0) Then 'stop
send_command Chanel2, Sound_Stop, No_Par
wave2p = 0
en2(0) = 0
End If


If en2(1) Then cl2(1) = cl2(1) - clc

If cl2(1) < 1 And en2(1) Then 'Volume envelop
Tsnd = (RAM(65303, 0) And 240) \ 16
If RAM(65303, 0) And 8 Then
Tsnd = Tsnd + 1
If Tsnd > 15 Then Tsnd = 15
Else
Tsnd = Tsnd - 1
If Tsnd < 0 Then Tsnd = 0
End If
RAM(65303, 0) = (RAM(65303, 0) And 15) + Tsnd * 16
send_command Chanel2, Wave_Volume_Set, (RAM(65303, 0) And 224) \ 16
cl2(1) = (en2(1) / 64) * 4194304
End If

'****Channel 3****
If en3(0) Then cl3(0) = cl3(0) - clc

If cl3(0) < 1 And en3(0) Then 'stop
send_command Chanel3, Sound_Stop, No_Par
wave3p = 0
en3(0) = 0
End If
End If

'****Chanel 4****
If en4(0) Then cl4(0) = cl4(0) - clc

If cl4(0) < 1 And en4(0) Then 'stop
send_command Chanel4, Sound_Stop, No_Par
wave4p = 0
en4(0) = 0
End If


If en4(1) Then cl4(1) = cl4(1) - clc

If cl4(1) < 1 And en4(1) Then 'Volume envelop
Tsnd = (RAM(65313, 0) And 240) \ 16
If RAM(65313, 0) And 8 Then
Tsnd = Tsnd + 1
If Tsnd > 15 Then Tsnd = 15
Else
Tsnd = Tsnd - 1
If Tsnd < 0 Then Tsnd = 0
End If
RAM(65313, 0) = (RAM(65313, 0) And 15) + Tsnd * 16
send_command Chanel4, Wave_Volume_Set, (RAM(65313, 0) And 240) \ 16
cl4(1) = (en4(1) / 64) * 4194304
End If

End Sub
'Register Writes
Sub setNR10(val As Long) 'Frequency Sweep,time,mode,sweep shift
en1(2) = val And 112
If en1(2) Then
cl1(2) = ((en1(2) \ 16) / 128) * 4194304
End If
End Sub
Sub setNR11(val As Long) 'wpd,len
send_command Chanel1, Wave_Pattern_Duty_Set, val \ 64 'bits 7-6
cl1(0) = (64 - (val And 63)) / 256
en1(0) = RAM(65300, 0) And 64
End Sub
Sub setNR12(val As Long) 'evelope reg
en1(1) = val And 7
cl1(1) = en1(1) * (1 / 64) * 4194304
send_command Chanel1, Wave_Volume_Set, (RAM(65298, 0) And 224) \ 16
End Sub
Sub setNR13(val As Long) 'freq 8 low bits
send_command Chanel1, Wave_Frequency_Set, (RAM(65300, 0) And 7) * 256 + RAM(65299, 0)
End Sub
Sub setNR14(val As Long) 'freq 3 hi bits, intial,counter
'if intial is set then play and set freq
send_command Chanel1, Wave_Frequency_Set, (RAM(65300, 0) And 7) * 256 + RAM(65299, 0)
If val And 128 Then send_command Chanel1, sound_play, No_Par: wave1p = 1

en1(0) = val And 64
cl1(0) = (64 - (RAM(65297, 0) And 63)) / 256 * 4194304
End Sub

Sub setNR21(val As Long)
send_command Chanel2, Wave_Pattern_Duty_Set, val \ 64 'bits 7-6
cl2(0) = (64 - (val And 63)) / 256
en2(0) = RAM(65305, 0) And 64
End Sub
Sub setNR22(val As Long) 'evelope reg
en2(1) = val And 7
cl2(1) = en2(1) * (1 / 64) * 4194304
send_command Chanel2, Wave_Volume_Set, (RAM(65303, 0) And 224) \ 16
End Sub
Sub setNR23(val As Long) 'freq 8 bit low
send_command Chanel2, Wave_Frequency_Set, (RAM(65305, 0) And 7) * 256 + RAM(65304, 0)
End Sub
Sub setNR24(val As Long)
send_command Chanel2, Wave_Frequency_Set, (RAM(65305, 0) And 7) * 256 + RAM(65304, 0)
If val And 128 Then send_command Chanel2, sound_play, No_Par: wave2p = 1
en2(0) = val And 64
cl2(0) = (64 - (RAM(65302, 0) And 63)) / 256 * 4194304
End Sub

Sub setNR30(val As Long)
If val And 128 Then
    send_command Chanel3, sound_play, No_Par
    wave3p = 1
Else
 send_command Chanel3, Sound_Stop, No_Par
 wave3p = 0
End If
End Sub
Sub setNR31(val As Long) 'len 0 - 255

End Sub
Sub setNR32(ByVal val As Long)
val = (val And 96) \ 32
If val = 1 Then val = 256 '1
If val = 2 Then val = 128 '1/2
If val = 3 Then val = 64 '1/4
send_command Chanel3, Wave_Pattern_Volume_Set, val
End Sub
Sub setNR33(val As Long)
send_command Chanel3, Wave_Frequency_Set, (RAM(65310, 0) And 7) * 256 + RAM(65309, 0)
End Sub
Sub setNR34(val As Long)
send_command Chanel3, Wave_Frequency_Set, (RAM(65310, 0) And 7) * 256 + RAM(65309, 0)
If val And 128 Then
send_command Chanel3, sound_play, No_Par
wave3p = 1
End If
en3(0) = val And 64
cl3(0) = (256 - (RAM(65307, 0))) / 256 * 4194304
End Sub
Sub setNR41(val As Long) 'sound len : 5-0
cl4(0) = (64 - (val And 63)) / 256 * 4194304
End Sub
Sub setNR42(val As Long) 'Evelope : 7-4 = envelope,3= envelope up/dn,2-0 = env. sweep
en4(1) = -((val And 7) > 0)
cl4(1) = (1 / 64) * 4194304 * (val And 7)
send_command Chanel4, Wave_Volume_Set, (val And 240) \ 16
End Sub
Sub setNR43(val As Long) 'freq
'Bit 7-4(m) - Selection of the shift clock
'Bit 3 - Selection of the polynomial bits
'Bit 2-0(n) - Selection of the dividing ratio
'if n=0 then n=0.5
freq = val '4194304 * 1 / 2 ^ 3 * 1 / (val And 7) * 1 / (2 ^ ((val \ 16) + 1))
send_command Chanel4, Wave_Frequency_Set, CLng(freq)
send_command Chanel4, 2, (val \ 8) And 1
End Sub
Sub setNR44(val As Long) 'intial/counter : 7 = intial , 6 = counter
'if intial is set then play and set freq
If val And 128 Then send_command Chanel4, sound_play, No_Par: _
send_command Chanel4, Wave_Frequency_Set, CLng(freq): wave4p = 1
en4(0) = (val And 64) \ 64
cl4(0) = (64 - (RAM(65312, 0) And 63)) / 256 * 4194304
en4(1) = -((RAM(65313, 0) And 7) > 0)
cl4(1) = (1 / 64) * 4194304 * (RAM(65313, 0) And 7)
send_command Chanel4, Wave_Volume_Set, (RAM(65313, 0) And 240) \ 16
End Sub


Sub setNR50(val As Long)

End Sub
Sub setNR51(val As Long)

End Sub
Sub setNR52(val As Long)
gb_snd = val \ 128
RAM(65318, 0) = gb_snd * 128 + _
            wave4p * 8 + _
            wave3p * 4 + _
            wave2p * 2 + _
            wave1p * 1
End Sub

Sub initWave() 'init square and noise waves patterns
Dim i As Long

'noise 7 bits
Randomize 10
For i = 0 To 127
noise7(i) = ((255 * Rnd) - 127)
Next i

'noise 15 bits
Randomize 11
For i = 0 To 32767
noise15(i) = ((255 * Rnd) - 127)
Next i

'I have some new info on this.. now waveforms are correct
'8*0.125%=1 : 00  : 12.5%  ____=___
sqrW(0, 0) = -127: sqrW(1, 0) = -127: sqrW(2, 0) = -127: sqrW(3, 0) = -127: sqrW(4, 0) = 128: sqrW(5, 0) = -127: sqrW(6, 0) = -127: sqrW(7, 0) = -127
'8*0.25=2   : 01  : 25%    ____==__
sqrW(0, 1) = -127: sqrW(1, 1) = -127: sqrW(2, 1) = -127: sqrW(3, 1) = -127: sqrW(4, 1) = 128: sqrW(5, 1) = 128: sqrW(6, 1) = -127: sqrW(7, 1) = -127
'8*0.50=4   : 10  : 50%    __====__
sqrW(0, 2) = -127: sqrW(1, 2) = -127: sqrW(2, 2) = 128: sqrW(3, 2) = 128: sqrW(4, 2) = 128: sqrW(5, 2) = 128: sqrW(6, 2) = -127: sqrW(7, 2) = -127
'8*0.75=6   : 11  : 75%    ====__==
sqrW(0, 3) = 128: sqrW(1, 3) = 128: sqrW(2, 3) = 128: sqrW(3, 3) = 128: sqrW(4, 3) = -127: sqrW(5, 3) = -127: sqrW(6, 3) = 128: sqrW(7, 3) = 128
'chanel 3 wave values
For i = 0 To 15
swm(i) = (i * 16) - 127
Next i
End Sub
