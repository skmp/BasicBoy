Attribute VB_Name = "modZ80cmd"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)


'v1.2.0
'This module contains the emulation of some z80 instructions
'Comments added
'Fixed some bugs
'Optimized a bit (for v2.0.0)
'Fixed some bugs from the optimization (2.0.2)

'Sory for my bad english ...

Option Explicit
Dim temp_var As Long, temp_var2 As Long 'temporary vars
Public Sub rl(ByRef reg8 As Long) 'Rotate left thru carry
    temp_var = reg8 \ 128
    reg8 = ((reg8 * 2) Or cf) And 255
    setZ reg8 = 0
    cf = temp_var
    hf = 0
    nf = 0
End Sub
Public Sub rla() 'Rotate left thru carry register a
    temp_var = A \ 128
    A = ((A * 2) Or cf) And 255
    setZ A = 0
    cf = temp_var
    hf = 0
    nf = 0
End Sub

Public Sub rlc(ByRef reg8 As Long) 'rotate left
    cf = reg8 \ 128
    reg8 = (reg8 * 2) And 255 Or cf
    setZ reg8 = 0
    hf = 0
    nf = 0
End Sub

Public Sub rlca() 'rotate left register A
    cf = A \ 128
    A = (A * 2) And 255 Or cf
    setZ A = 0
    hf = 0
    nf = 0
End Sub

Public Sub rr(ByRef reg8 As Long) 'Rotate right thru carry
    temp_var = reg8 And 1
    reg8 = (reg8 \ 2) Or (128 * cf)
    setZ reg8 = 0
    cf = temp_var
    hf = 0
    nf = 0
End Sub

Public Sub rra() 'Rotate right thru carry Register A
    temp_var = A And 1
    A = (A \ 2) Or (128 * cf)
    setZ A = 0
    cf = temp_var
    hf = 0
    nf = 0
End Sub

Public Sub rrc(ByRef reg8 As Long) 'Rotate right
    cf = reg8 And 1
    reg8 = (reg8 \ 2) Or (128 * cf)
    setZ reg8 = 0
    hf = 0
    nf = 0
End Sub
Public Sub rrca() 'Rotate right register A
    cf = A And 1
    A = (A \ 2) Or (128 * cf)
    setZ A = 0
    hf = 0
    nf = 0
End Sub
Public Sub sla(ByRef reg8 As Long) 'Shift Left
    cf = reg8 \ 128
    reg8 = (reg8 * 2) And 255
    setZ reg8 = 0
    hf = 0
    nf = 0
End Sub
Public Sub sra(ByRef reg8 As Long) 'Shift Right arithmetic
    cf = reg8 And 1
    reg8 = (reg8 \ 2) Or (reg8 And 128)
    setZ reg8 = 0
    hf = 0
    nf = 0
End Sub
Public Sub srl(ByRef reg8 As Long) 'Shift Right logical
    cf = reg8 And 1
    reg8 = reg8 \ 2
    setZ reg8 = 0
    hf = 0
    nf = 0
End Sub
Public Sub zsub(ByRef reg8 As Long) 'Substract from a
    temp_var = A - reg8
    temp_var2 = temp_var And 255
    setZ temp_var2 = 0
    setC A < temp_var2
    setH (A And 15) < (temp_var2 And 15)
    nf = 1
    A = temp_var2
End Sub
Public Sub sbc(ByRef reg8 As Long) 'Substract from a - carry
    temp_var = (A - reg8 - cf) And 255
    setZ temp_var = 0
    setC A < temp_var
    setH (A And &HF&) < (temp_var And 15)
    nf = 1
    A = temp_var
End Sub
Public Sub add(ByRef reg8 As Long) 'add to a
    temp_var = (A + reg8) And 255
    setZ temp_var = 0
    cf = (A + reg8) \ 256
    hf = ((A And 15) + (reg8 And 15)) \ 16
    nf = 0
    A = temp_var
End Sub
Public Sub adc(ByRef reg8 As Long) 'add to a + carry
    temp_var = (A + reg8 + cf) And 255
    setZ temp_var = 0
    cf = (A + reg8 + cf) \ 256
    hf = ((A And 15) + (reg8 And 15) + cf) \ 16
    nf = 0
    A = temp_var
End Sub
Sub push(ByRef reg8 As Long) 'Push to the stack
SP = SP - 1
WriteM SP, reg8
End Sub

Sub pop(ByRef reg8 As Long) 'pop from the stack
reg8 = readM(SP)
SP = SP + 1
End Sub

Sub zand(ByRef val As Long) 'Logocal and
A = A And val
setZ A = 0
nf = 0
cf = 0
hf = 1
End Sub
Sub zor(ByRef val As Long) 'Logocal or
A = A Or val
setZ A = 0
nf = 0
hf = 0
cf = 0
End Sub
Sub zxor(ByRef val As Long) 'Logocal xor
A = A Xor val
setZ A = 0
nf = 0
hf = 0
cf = 0
End Sub
Sub cp(ByRef val As Long) 'Compare with A
setC A < val
setH (A And 15) < (val And 15)
setZ A = val
nf = 1
End Sub
Sub dec(ByRef reg8 As Long) 'decrease
temp_var = reg8
reg8 = reg8 - 1
reg8 = reg8 And 255
setH (reg8 And 15) < (temp_var And 15)
setZ reg8 = 0
nf = 1
End Sub
Sub dec16(ByRef reg81 As Long, ByRef reg82 As Long) 'same but for 16bit
reg82 = reg82 - 1
reg81 = reg81 + (reg82 < 0)
reg81 = reg81 And 255
reg82 = reg82 And 255
End Sub
Sub inc(ByRef reg8 As Long) 'increase
reg8 = reg8 + 1
reg8 = reg8 And 255
setH ((reg8 And 15) = 0)
setZ reg8 = 0
nf = 0
End Sub
Sub inc16(ByRef reg81 As Long, ByRef reg82 As Long) 'same but for 16bit
reg82 = reg82 + 1
reg81 = reg81 + (reg82 \ 256)
reg81 = reg81 And 255: reg82 = reg82 And 255
End Sub
Sub addHL(ByRef r1h As Long, ByRef r1l As Long) 'add to hl
setC ((H * 256 Or L) + (r1h * 256 Or r1l)) > 65535
setH (((H * 256 Or L) And 4095) + ((r1h * 256 Or r1l) And 4095)) > 4095
L = L + r1l
H = H + r1h
If L > 255 Then H = H + 1
L = L And 255: H = H And 255
nf = 0
End Sub
Sub addSP(ByVal value As Long) 'add to sp(StackPointer)
If value > 127 Then value = value - 256
temp_var = SP + value
temp_var = temp_var And 65535
        If value > 0 Then
          setC SP > temp_var
          setH ((SP Xor value Xor temp_var) And 4096) > 0
          SP = temp_var
        Else
          setC SP < temp_var
          setH ((SP Xor value Xor temp_var) And 4096) > 0
          SP = temp_var
        End If
        zf = 0
        nf = 0
End Sub
Sub swap(ByRef reg8 As Long) 'Swap nibles
reg8 = (reg8 \ 16) Or ((reg8 And 15) * 16)
setZ reg8 = 0
nf = 0
hf = 0
cf = 0
End Sub
Sub daa() 'Demical adjust register A
        If hf Then
           If ((A And 15) >= 10 Or hf) Then A = A - 6
           If ((A And 240) >= 160 Or cf) Then A = A - 96: cf = 1
        Else
           If ((A And 15) >= 10 Or hf) Then A = A + 6
           If ((A And 240) >= 160 Or cf) Then A = A + 96: cf = 1
        End If
        A = A And 255
setZ A = 0
hf = 0
End Sub
Sub cpl() 'logical not
A = 255 - A
hf = 1
nf = 1
End Sub
Sub halt() 'wait interupt
   If IME = False Then Exit Sub
    temp_var = RAM(65535, 0) And RAM(65295, 0)    ' AND IE, IF
    If temp_var = 0 Then PC = PC - 1: Exit Sub                  'If no Interrupt occured exit
    'Process Interrput
    'Push pc
    SP = SP - 1
    WriteM SP, PC \ 256
    SP = SP - 1
    WriteM SP, PC And 255
    IME = False
    If (temp_var And 1) = 1 Then        'V-Blank ?
        PC = 64
        RAM(65295, 0) = RAM(65295, 0) And 254
    ElseIf (temp_var And 2) = 2 Then    'LCDC ?
        PC = 72
        RAM(65295, 0) = RAM(65295, 0) And 253
    ElseIf (temp_var And 4) = 4 Then    'Timer ?
        PC = 80
        RAM(65295, 0) = RAM(65295, 0) And 251
    ElseIf (temp_var And 8) = 8 Then    'Serial ?
        PC = 88
        RAM(65295, 0) = RAM(65295, 0) And 247
    ElseIf (temp_var And 16) = 16 Then  'Joypad ?
        PC = 96
        RAM(65295, 0) = RAM(65295, 0) And 239
    End If
End Sub
Sub bit(ByRef reg8 As Long, b As Long) 'test bit
setZ ((reg8 And b) = 0)
nf = 0
hf = 1
End Sub
Sub zset(ByRef reg8 As Long, b As Long) ' set bit
reg8 = reg8 Or b
End Sub
Sub res(ByRef reg8 As Long, b As Long) 'reset bit
reg8 = reg8 And b
End Sub
Sub jp(Adr As Long, Optional cc As Byte = 1) 'jump to
If cc Then PC = Adr
End Sub
Sub jr(ByVal val As Long, Optional cc As Byte = 1) 'jump local to
If val > 127 Then val = val - 256
If cc Then PC = PC + val
End Sub
Sub zcall(Adr As Long, Optional cc As Byte = 1) 'call subroutine
If cc Then
SP = SP - 1
WriteM SP, PC \ 256
SP = SP - 1
WriteM SP, PC And 255
PC = Adr
End If
End Sub
Sub rst(value As Long) 'restart at
SP = SP - 1
WriteM SP, PC \ 256
SP = SP - 1
WriteM SP, PC And 255
PC = value
End Sub
Sub ret(Optional cc As Byte = 1) 'return from subroutine
If cc Then PC = readM(SP) Or readM(SP + 1) * 256: SP = SP + 2
End Sub
Sub reti() 'return from subroutine ,enable interups
PC = readM(SP) Or readM(SP + 1) * 256
SP = SP + 2
IME = True
End Sub

