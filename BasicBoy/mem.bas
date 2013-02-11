Attribute VB_Name = "modMem"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latestt version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)

'v2.1.5
'Ram I/O functions...
'Base taken from VisBoy (Uptaded,Corected,Recoded)
'Currenlty no optimizations
'Coments added
'Rewrote mbc code
'RTC support
'Sory for my bad english ...

Option Explicit
Option Base 0

'Index/Temp internal vars
Dim i As Long, j As Long
Dim memptr2 As Long
Dim bgtmp As Long, bgtmp2 As Long, bgtmp3 As Long
Dim tmpcolor As Byte, tmpcolor2 As Byte
Public Function readM(memptr As Long) As Long
If GBM = 0 Then
    If memptr < 16384 Then
            readM = ROM(memptr, 0)      ' Read from ROM
    ElseIf memptr < 32768 Then
            readM = ROM(memptr - 16384, CurROMBank)      ' Read from ROM
    Else
            If memptr > 40959 And memptr < 49152 Then
            If mbcrtcE = 0 Then
                readM = bRam(memptr - 40960, CurRAMBank)    ' Read from sRAM
            Else
                readM = mbc3rtc.readReg 'read RTC registers
            End If
            Else
                readM = RAM(memptr, 0)      ' Read from RAM
            End If
    End If
Else
    If memptr < 16384 Then
            readM = ROM(memptr, 0)      ' Read from ROM
    ElseIf memptr < 32768 Then
            readM = ROM(memptr - 16384, CurROMBank)      ' Read from ROM
    ElseIf memptr < 40960 Then 'read Vram
        readM = RAM(memptr, vRamB)
    ElseIf memptr < 49152 Then 'read sRam
    If mbcrtcE = 0 Then
        readM = bRam(memptr - 40960, CurRAMBank)
    Else
        readM = mbc3rtc.readReg
    End If
    ElseIf memptr < 53248 Then 'read wRam(0)
        readM = RAM(memptr, 0)
    ElseIf memptr < 57344 Then 'read wRam(1-7)
        readM = RAM(memptr, wRamB)
    Else 'read ram
        readM = RAM(memptr, 0)      ' Read from RAM
    End If
End If
End Function
Public Sub WriteM(memptr As Long, ByVal value As Long)
    If memptr > 32767 Then 'ram/mmio
    'ram
    If GBM = 0 Then 'Old gameboy
                If memptr > 40959 And memptr < 49152 Then    ' write to sRAM
                If mbcrtcE = 0 Then
                    bRam(memptr - 40960, CurRAMBank) = value
                    Exit Sub
                Else
                    mbc3rtc.writeReg value
                End If
                Else
                RAM(memptr, 0) = value    ' write to RAM
                If memptr > &HE000 And memptr < &HFE00 Then ' echo
                RAM(memptr - 8192, 0) = value
                ElseIf memptr > &HC000 And memptr < &HDE00 Then ' echo
                RAM(memptr + 8192, 0) = value
                End If
                End If
    Else
    Select Case memptr 'GameBoy color
            Case Is < 40960 'write Vram
            RAM(memptr, vRamB) = value
            Exit Sub
            Case Is < 49152 'write sRam
            If mbcrtcE = 0 Then
                bRam(memptr - 40960, CurRAMBank) = value
            Else
                mbc3rtc.writeReg value
            End If
            Exit Sub
            Case Is < 53248 'write wRam(0)
            RAM(memptr, 0) = value
            Exit Sub
            Case Is < 57344 'write wRam(1-7)
            RAM(memptr, wRamB) = value
            Exit Sub
            Case Else 'write ram
            RAM(memptr, 0) = value
    End Select
End If

    'Memory Maped Registers
    If memptr > 65279 Then
    If memptr > 65327 And memptr < 65344 Then 'Wave data for chanel 3
    send_command Chanel3, wave_waveform_write, value * 32 + (memptr - 65328)
    Exit Sub
    End If
    Select Case memptr
        Case Is = 65280     ' Joypad
            If (value And 32) = 32 Then         'Directional
                RAM(65280, 0) = 223 And (255 - joyval1)
            ElseIf (value And 16) = 16 Then     ' Buttons
                RAM(65280, 0) = 239 And (255 - joyval2)
            Else
                RAM(65280, 0) = 255
            End If
        Case Is = 65350     ' DMA Xfer
            RAM(65350, 0) = value
            j = value * 256
            For i = 65024 To 65183
                RAM(i, 0) = readM(j)
                j = j + 1
            Next i
        Case 65351, 65352, 65353 'Old gameboy palets
            ccolid2 value, memptr - 65351
        Case 65287 'Timer
            Select Case value And 3
            Case 0
                tvm = 1024
            Case 1
                tvm = 65536
            Case 2
                tvm = 16384
            Case 3
                tvm = 4096
            End Select
            
        'Sound regs
        Case 65296 ' NR10
        If snd Then setNR10 value
        Case 65297 ' NR11
        If snd Then setNR11 value
        Case 65298 ' NR12
        If snd Then setNR12 value
        Case 65299 ' NR13
        If snd Then setNR13 value
        Case 65300 ' NR14
        If snd Then setNR14 value
        Case 65302 ' NR21
        If snd Then setNR21 value
        Case 65303 ' NR22
        If snd Then setNR22 value
        Case 65304 ' NR23
        If snd Then setNR23 value
        Case 65305 ' NR24
        If snd Then setNR24 value
        Case 65306 ' NR30
        If snd Then setNR30 value
        Case 65307 ' NR31
        If snd Then setNR31 value
        Case 65308 ' NR32
        If snd Then setNR32 value
        Case 65309 ' NR33
        If snd Then setNR33 value
        Case 65310 ' NR34
        If snd Then setNR34 value
        Case 65312 ' NR41
        If snd Then setNR41 value
        Case 65313 ' NR42
        If snd Then setNR42 value
        Case 65314 ' NR43
        If snd Then setNR43 value
        Case 65315 ' NR44
        If snd Then setNR44 value
        Case 65316 ' NR50
        If snd Then setNR50 value
        Case 65317 ' NR51
        If snd Then setNR51 value
        Case 65318 ' NR52
        If snd Then
        setNR52 value
        Else
        RAM(65318, 0) = 0
        End If
    End Select
    'Gameboy Color Olny
        If GBM = 1 Then
            Select Case memptr
            Case 65357  'Speed SW
                smp = value And 1
                RAM(65357, 0) = CpuS * 128 + smp
            Case 65359  'Vram Bank
                vRamB = value And 1
                RAM(65359, 0) = vRamB
            Case 65361  'HDMA1 sh
                hdmaS = (hdmaS And 255) + value * 256
            Case 65362 'HDMA2 sl
                hdmaS = (hdmaS And 65280) + value
            Case 65363  'HDMA3 dh
                hdmaD = (hdmaD And 255) + value * 256
            Case 65364 'HDMA4 dl
                hdmaD = (hdmaD And 65280) + value
            Case 65365 'HDMA5 lms
                hdmaD = (hdmaD And 8176) + 32768
                hdmaS = hdmaS And 65520
                If Hdma = True Then If (value And 128) = 0 Then Hdma = False: RAM(65365, 0) = 128 + 70: Exit Sub Else Exit Sub
                If value And 128 Then Hdma = True: Hdmal = value And 127: tHdmal = value And 127: RAM(65365, 0) = Hdmal: Exit Sub
                j = hdmaD
                For i = hdmaS To hdmaS + (value And 127) * 16 + 15
                RAM(j, vRamB) = readM(i)
                j = j + 1
                Next i
                RAM(65365, 0) = 255
            Case 65366  'Rp
            'InfraRed
            
            Case 65384  'BG pal indx
                bgpi = value And 63
                bgai = value And 128
                If bgpi Mod 2 Then RAM(65385, 0) = bgp(bgpi \ 8, (bgpi \ 2) Mod 4) \ 256 Else RAM(65385, 0) = bgp(bgpi \ 8, (bgpi \ 2) Mod 4) And 255
        
            Case 65385 'BG Pal Val
                i = bgpi Mod 2
                bgtmp = bgpi \ 8
                bgtmp2 = (bgpi \ 2) Mod 4
                If i = 0 Then ' 1st byte
                bgp(bgtmp, bgtmp2) = (bgp(bgtmp, bgtmp2) And 65280) + value
                bgtmp3 = bgp(bgtmp, bgtmp2)
                bgpCC(bgtmp, bgtmp2) = (bgtmp3 And 31744) \ 1024 + (bgtmp3 And 992) + (bgtmp3 And 31) * 1024
                Else '2nd byte
                bgp(bgtmp, bgtmp2) = ((bgp(bgtmp, bgtmp2) And 255) + value * 256) And 32767
                bgtmp3 = bgp(bgtmp, bgtmp2)
                bgpCC(bgtmp, bgtmp2) = (bgtmp3 And 31744) \ 1024 + (bgtmp3 And 992) + (bgtmp3 And 31) * 1024
                End If
                If bgai Then bgpi = bgpi + 1
                WriteM 65384, (RAM(65384, 0) And 128) Or (bgpi And 63)
        
            Case 65386  'OBJ pal indx
                objpi = value And 63
                objai = value And 128
                If objpi Mod 2 Then RAM(65387, 0) = objp(objpi \ 8, (objpi \ 2) Mod 4) \ 256 Else RAM(65387, 0) = objp(objpi \ 8, (objpi \ 2) Mod 4) And 255
        
            Case 65387 'OBJ Pal Val
        
                i = objpi Mod 2
                bgtmp = objpi \ 8
                bgtmp2 = (objpi \ 2) Mod 4
                If i = 0 Then ' 1st byte
                objp(bgtmp, bgtmp2) = (objp(bgtmp, bgtmp2) And 65280) + value
                bgtmp3 = objp(bgtmp, bgtmp2)
                objpCC(bgtmp, bgtmp2) = (bgtmp3 And 31744) \ 1024 + (bgtmp3 And 992) + (bgtmp3 And 31) * 1024
                Else '2nd byte
                objp(bgtmp, bgtmp2) = ((objp(bgtmp, bgtmp2) And 255) + value * 256) And 32767
                bgtmp3 = objp(bgtmp, bgtmp2)
                objpCC(bgtmp, bgtmp2) = (bgtmp3 And 31744) \ 1024 + (bgtmp3 And 992) + (bgtmp3 And 31) * 1024
                End If
        
        
                If objai Then objpi = objpi + 1
                WriteM 65386, (RAM(65386, 0) And 128) Or (objpi And 63)
        
                Case 65392  'SVBK
                wRamB = value And 7
                If wRamB < 1 Then wRamB = 1
                RAM(65392, 0) = wRamB
        End Select
        End If
        
    End If
    Else    'MBCs Control registers
            Select Case rominfo.Ctype
            
            Case 1, 2, 3 ' mbc1
            If memptr > &H1FFF And memptr < &H4000 Then 'std rom banks
            value = value And 31 'XXXBBBBB->00011111->31
            If value = 0 Then value = 1
            CurROMBank = value
            ElseIf memptr > &H3FFF And memptr < &H6000 Then 'ram/extended rom
            value = value And 2
            If mbc1mode = 1 Then 'ram
            CurRAMBank = value
            Else 'rom
            CurROMBank = value * 32 + (CurROMBank And 31) 'value << 5 + last 5 from crb
            End If
            ElseIf memptr > &H5FFF And memptr < &H8000 Then 'Model selection
            mbc1mode = value And 1
            End If
            
            Case &HF, &H10, &H11, &H12, &H13 'mbc3
            If memptr > &H1FFF And memptr < &H4000 Then 'rom
            value = value And 127 'olny 7 lower
            If value = 0 Then value = 1
            CurROMBank = value
            ElseIf memptr > &H3FFF And memptr < &H6000 Then 'ram
            If value < rominfo.ramsize Then
            CurRAMBank = value: mbcrtcE = 0
            Else
            If mbcrtc = 1 Then mbc3rtc.act = value: mbcrtcE = 1
            End If
            End If
            
            Case &H19, &H1A, &H1B, &HC, &H1D, &H1E 'mbc5
            If memptr > &H1FFF And memptr < &H3000 Then
            CurROMBank = value + (CurROMBank And 256)
            ElseIf memptr < &H4000 Then
            value = value And 1
            CurROMBank = (CurROMBank And 255) + value * 256
            ElseIf memptr < &H6000 Then
            CurRAMBank = value And 16
            End If
            

            
        End Select
    End If
End Sub

Public Sub initCI() 'Init mem interface
Dim i As Long, i2 As Long
For i2 = 0 To 7
For i = 32768 To 65535
RAM(i, i2) = 0
Next i
Next i2
For i = 0 To 7
bgp(i, 0) = 32767: objp(i, 0) = 32767
bgp(i, 1) = 32767: objp(i, 1) = 32767
bgp(i, 2) = 32767: objp(i, 2) = 32767
bgp(i, 3) = 32767: objp(i, 3) = 32767
Next i
CurROMBank = 1
Ct(0) = "Rom Only": Ct(&H12) = "Rom+MBC3+Ram"
Ct(1) = "Rom+MBC1": Ct(&H13) = "Rom+MBC3+Ram+Batt"
Ct(2) = "Rom+MBC1+Ram": Ct(&H19) = "Rom+MBC5"
Ct(3) = "Rom+MBC1+Ram+Batt": Ct(&H1A) = "Rom+MBC5+Ram"
Ct(5) = "Rom+MBC2": Ct(&H1B) = "Rom+MBC5+Ram+Batt"
Ct(6) = "Rom+MBC2+Batt": Ct(&H1C) = "Rom+MBC5+Rumble"
Ct(8) = "Rom+Ram": Ct(&H1D) = "Rom+MBC5+Rumble+Sram"
Ct(9) = "Rom+Ram+Batt": Ct(&H1E) = "Rom+MBC5+Rumble+Sram+Batt"
Ct(&HB) = "Rom+MMO1": Ct(&H1F) = "Pocet Camera"
Ct(&HC) = "Rom+MMO1+Sram": Ct(&HFD) = "Bandai TAMA5"
Ct(&HD) = "Rom+MMO1+Sram+Batt": Ct(&HFE) = "Hudson HuC-3"
Ct(&HF) = "Rom+MBC3+Timer+Batt": Ct(&HFF) = "Hudson HuC-1"
Ct(&H10) = "Rom+MBC3+Timer+Ram+Batt"
Ct(&H11) = "Rom+MBC3"
rominfo.Ctype = ROM(&H147, 0)
For i = &H134 To &H142
    rominfo.titleB(i - &H134) = ROM(i, 0)
Next i
If rominfo.Ctype = &HF Or rominfo.Ctype = &H10 Then
mbcrtc = 1
Else
mbcrtc = 0
End If
rominfo.Title = StrConv(rominfo.titleB, vbUnicode)
rominfo.romsize = ROM(&H148, 0)
rominfo.ramsize = ROM(&H149, 0)
Ros(0) = "32 Kbyte": Rosn(0) = 2
Ros(1) = "64 Kbyte": Rosn(1) = 4
Ros(2) = "128 Kbyte": Rosn(2) = 8
Ros(3) = "256 Kbyte": Rosn(3) = 16
Ros(4) = "512 Kbyte": Rosn(4) = 32
Ros(5) = "1 Mbyte": Rosn(5) = 64
Ros(6) = "2 Mbyte": Rosn(6) = 128
Ros(52) = "1.1 Mbyte": Rosn(52) = 72
Ros(53) = "1.2 Mbyte": Rosn(53) = 80
Ros(54) = "1.5 Mbyte": Rosn(54) = 96

Ras(0) = "None": Rasn(0) = 0
Ras(1) = "2 Kbyte": Rasn(1) = 1
Ras(2) = "8 Kbyte": Rasn(2) = 1
Ras(3) = "32 Kbyte": Rasn(3) = 4
Ras(4) = "128 Kbyte": Rasn(4) = 16
End Sub

Sub wrRam() 'Write Sram to disk
Dim tRam() As Byte
On Error GoTo enf
If Len(ro) > 0 Then
If mbcrtc Then mbc3rtc.save ro
ReDim tRam(Rasn(rominfo.ramsize) * 8192 - 1)
CopyMemory tRam(0), bRam(0, 0), UBound(tRam) + 1
Open ro For Binary As #1
Put #1, , tRam
Close #1
enf:
Close #1
ro = ""
End If
End Sub
Sub rdRam() 'Read Sram from disk
Dim tRam() As Byte
On Error GoTo enf
If Len(ro) > 0 Then wrRam
ro = CD.FileName & ".sav"
If mbcrtc Then mbc3rtc.load ro
ReDim tRam(Rasn(rominfo.ramsize) * 8192 - 1)
Open ro For Binary As #1
Get #1, , tRam
Close #1
CopyMemory bRam(0, 0), tRam(0), UBound(tRam) + 1
enf:
Close #1
End Sub
Public Function pb() As Long 'read a byte at pc ,increase pc
If GBM = 0 Then
    Select Case PC
        Case Is < 16384
            pb = ROM(PC, 0)      ' Read from ROM
        Case Is < 32768
            pb = ROM(PC - 16384, CurROMBank)      ' Read from ROM
        Case Else
            If PC > 40959 And PC < 49152 Then
            pb = bRam(PC - 40960, CurRAMBank)    ' Read from sRAM
            Else
            pb = RAM(PC, 0)      ' Read from RAM
            End If
    End Select
Else
    Select Case PC
        Case Is < 16384
            pb = ROM(PC, 0)      ' Read from ROM
        Case Is < 32768
            pb = ROM(PC - 16384, CurROMBank)      ' Read from ROM
        Case Is < 40960 'read Vram
        pb = RAM(PC, vRamB)
        Case Is < 49152 'read sRam
        pb = bRam(PC - 40960, CurRAMBank)
        Case Is < 53248 'read wRam(0)
        pb = RAM(PC, 0)
        Case Is < 57344 'read wRam(1-7)
        pb = RAM(PC, wRamB)
        Case Else 'read ram
        pb = RAM(PC, 0) ' Read from RAM
    End Select
End If
PC = PC + 1
End Function
Public Function pw() As Long 'read a word at pc ,increase pc
If GBM = 0 Then
    Select Case PC
        Case Is < 16384
            pw = ROM(PC, 0)      ' Read from ROM
        Case Is < 32768
            pw = ROM(PC - 16384, CurROMBank)      ' Read from ROM
        Case Else
            If PC > 40959 And PC < 49152 Then
            pw = bRam(PC - 40960, CurRAMBank)    ' Read from sRAM
            Else
            pw = RAM(PC, 0)      ' Read from RAM
            End If
    End Select
    PC = PC + 1
        Select Case PC
        Case Is < 16384
            pw = pw + ROM(PC, 0) * 256  ' Read from ROM
        Case Is < 32768
            pw = pw + ROM(PC - 16384, CurROMBank) * 256  ' Read from ROM
        Case Else
            If PC > 40959 And PC < 49152 Then
            pw = pw + bRam(PC - 40960, CurRAMBank) * 256 ' Read from sRAM
            Else
            pw = pw + RAM(PC, 0) * 256  ' Read from RAM
            End If
    End Select
    PC = PC + 1
Else
    Select Case PC
        Case Is < 16384
            pw = ROM(PC, 0)      ' Read from ROM
        Case Is < 32768
            pw = ROM(PC - 16384, CurROMBank)      ' Read from ROM
        Case Is < 40960 'read Vram
        pw = RAM(PC, vRamB)
        Case Is < 49152 'read sRam
        pw = bRam(PC - 40960, CurRAMBank)
        Case Is < 53248 'read wRam(0)
        pw = RAM(PC, 0)
        Case Is < 57344 'read wRam(1-7)
        pw = RAM(PC, wRamB)
        Case Else 'read ram
        pw = RAM(PC, 0)      ' Read from RAM
    End Select
    PC = PC + 1
    Select Case PC
        Case Is < 16384
        pw = pw + ROM(PC, 0) * 256  ' Read from ROM
        Case Is < 32768
        pw = pw + ROM(PC - 16384, CurROMBank) * 256  ' Read from ROM
        Case Is < 40960 'read Vram
        pw = pw + RAM(PC, vRamB) * 256
        Case Is < 49152 'read sRam
        pw = pw + bRam(PC - 40960, CurRAMBank) * 256
        Case Is < 53248 'read wRam(0)
        pw = pw + RAM(PC, 0) * 256
        Case Is < 57344 'read wRam(1-7)
        pw = pw + RAM(PC, wRamB) * 256
        Case Else 'read ram
        pw = pw + RAM(PC, 0) * 256  ' Read from RAM
    End Select
    PC = PC + 1
End If


End Function
Public Function readHackM(ByVal memptr As Long, ByVal rb As Long) As Long
If GBM = 0 Then
    If memptr < 16384 Then
    ElseIf memptr < 32768 Then
    Else
            If memptr > 40959 And memptr < 49152 Then
            readHackM = bRam(memptr - 40960, rb)    ' Read from sRAM
            Else
            readHackM = RAM(memptr, 0)      ' Read from RAM
            End If
    End If
Else
    If memptr < 16384 Then
    ElseIf memptr < 32768 Then
    ElseIf memptr < 40960 Then 'read Vram
        readHackM = RAM(memptr, rb)
    ElseIf memptr < 49152 Then 'read sRam
        readHackM = bRam(memptr - 40960, rb)
    ElseIf memptr < 53248 Then 'read wRam(0)
        readHackM = RAM(memptr, 0)
    ElseIf memptr < 57344 Then 'read wRam(1-7)
        readHackM = RAM(memptr, rb)
    Else 'read ram
        readHackM = RAM(memptr, 0)      ' Read from RAM
    End If
End If
End Function
Public Sub wHackM(ByVal memptr As Long, ByVal rb As Long, ByVal val As Long)
If GBM = 0 Then
    If memptr < 16384 Then
    ElseIf memptr < 32768 Then
    Else
            If memptr > 40959 And memptr < 49152 Then
            bRam(memptr - 40960, rb) = val  ' Read from sRAM
            Else
            RAM(memptr, 0) = val   ' Read from RAM
            End If
    End If
Else
    If memptr < 16384 Then
    ElseIf memptr < 32768 Then
    ElseIf memptr < 40960 Then 'read Vram
        RAM(memptr, rb) = val
    ElseIf memptr < 49152 Then 'read sRam
        bRam(memptr - 40960, rb) = val
    ElseIf memptr < 53248 Then 'read wRam(0)
        RAM(memptr, 0) = val
    ElseIf memptr < 57344 Then 'read wRam(1-7)
        RAM(memptr, rb) = val
    Else 'read ram
        RAM(memptr, 0) = val    ' Read from RAM
    End If
End If
End Sub

Public Function loadrom(FileName As String)
Dim ROMBank  As Long, i As Long
loadrom = False
Dim tmp(16383) As Byte
'filename = CD.ShowOpen(Me.hwnd, "", CD.filename, "GameBoy Roms (*.gb;*.gbc;*.cgb)|*.gb;*.gbc;*.cgb")
If Len(FileName) > 0 Then
If Right(FileName, 1) = Chr$(0) Then FileName = Left(FileName, Len(FileName) - 1)
CD.FileName = FileName
If FileLen(FileName) = 0 Then
MsgBox "Rom file " & FileName & " was not found/has size 0"
Exit Function
End If
Open FileName For Binary As #1
'ReDim ROM(16383, (LOF(1) / 16384))
ROMBank = 0
While Not EOF(1)
    Get #1, , tmp
    For i = 0 To 16383
    ROM(i, ROMBank) = tmp(i)
    Next i
    ROMBank = ROMBank + 1
Wend
Close #1
loadrom = True
End If
End Function
