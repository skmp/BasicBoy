Attribute VB_Name = "modVars"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)

'v1.1.0
'All global / public vars moved here
'Comments Added
'minor timing fixes

'Sory for my bad english ...
Option Explicit
Option Base 0
'****Gui/General vars****
Global CD As New clsDialog 'Common dialog class
Global Const pschome As String = "Explorer " & """" & "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=54373&lngWId=1" & """"
Global gtc As Long 'use gettickcount instead of queryperformancecounter
'****Cpu vars****
Global Wait_Data As Long
Global zf As Byte, nf As Byte, hf As Byte, cf As Byte 'bits 7-4
Global f_lowbits As Byte 'bits 3-0
Global A  As Long, f As Long  'Registers
Global b As Long, c As Long 'Registers
Global D As Long, E As Long 'Registers
Global H As Long, L As Long 'Registers
Global PC As Long 'PC Register
Global SP As Long 'SP Register
Global IME As Boolean 'Interupt Master Enable Register
Global timerC As Long 'Timer interupt Cycle counter
'Global brkAddr As Long'CPU break Address (not used anymore)
Global Clcount As Long 'CPU sync cycle counter
Global ime_stat As Long 'IME delay state
Global SETT(0 To 56) As Byte, BITT(0 To 56) As Byte 'Precalculated for speed
Global bCpuRun As Boolean 'Cpu run :P
'Global tval As Long
Global tvm As Long 'Timer Iterupt Speed
Global cpc(&HFF) As Long 'Intruction Cylces table
'Global lw As Long
Global mm As Long, zm As Long, cm As Long 'GRFX mode,GRFX Size,CPU mode(not used currently)
Global GBM As Long '0 = gb, 1 = gbc
Global message As msg 'replacent for the doevents
Global TGBC As Boolean 'Try to emulate GBC(if rom is color)
Global smp As Long 'Prepare Speed Change reg
Global Clm0 As Long, clm3 As Long, cllc As Long, cldr As Long 'CPU Cycle Sync
Global CpuS As Long 'Cpu speed(1,0)
Global lfp As Byte 'Limit FPS
Global hline As Long, Mhz As Long 'Cur Hblank line,Cpu speed
Global Mips As Long 'Cpu mips
Global Cpu_Speed As Single 'cpu over/under-clock

'****Memory Vars****
Global rominfo As CartIinfo 'Cart info All together :)
Global Ct(255) As String 'cart type name
Global Ros(255) As String 'Rom size name
Global Ras(255) As String 'Ram size name
Global Rosn(255) As Long 'Rom size number
Global Rasn(255) As Long 'Ram size Number
Global ROM(16383, 128) As Long 'Memory to hold the rom/Ram:Staticaly dimed to max to
Global RAM(32768 To 65535, 7) As Long 'Max and as long in order to increase speed(
Global bRam(8191, 15) As Byte '8megs wow gb has max 2...)
Global CurROMBank As Long, CurRAMBank As Long ' No comment :P
Global mbcrtc As Long 'Cart has RTC??
Global mbcrtcE As Long  'Is it Emulated??
Global mbc1mode As Long  '16/8 or 4/32 mode for mbc1
Global mbc3rtc As New rtc  'Real Time Clock class
Global wRamB As Long, vRamB As Long 'work ram and video ram banks (GBC)
Global objpi As Long, bgpi As Long 'Color Paletes indexes
Global bgai As Boolean, objai As Boolean 'Color Paletes indexes
Global hdmaS As Long, hdmaD As Long 'hDMA source / Dest
Global Hdma As Byte, Hdmal As Long, tHdmal As Long 'hDMA len/Eneble disable
Global joyval1 As Long, joyval2 As Long 'Joystick Values
Global ro As String 'Rom FileName

'****Sound vars****
Global snd As Long 'Sound enabled (1,0)
Global sqrW(7, 3) As Integer 'Square Wave Waveform
Global noise7(127) As Long, noise15(32767) As Long, npointer As Long 'Noise tables/pointer
Global ssnd As Long, gb_snd As Long
Global wave1p As Long
Global wave2p As Long
Global wave3p As Long
Global wave4p As Long

'****Joystic Vars****
Global Up As Long 'Button keycodes
Global Dn As Long 'Button keycodes
Global Lf As Long 'Button keycodes
Global Rg As Long 'Button keycodes
Global ABut As Long 'Button keycodes
Global BBut As Long 'Button keycodes
Global St1 As Long, St2 As Long, St3 As Long 'Button keycodes
Global Sl1 As Long, Sl2 As Long 'Button keycodes
Global SpeedKeyD As Long, SpeedKeyU As Long 'Button keycodes
Global sTes As Long, ofmode As Long, ofskip As Long 'Button keycodes
Global stpsnd As Boolean, stpsk As Boolean, slfp As Boolean 'Button keycodes


Function GetTickCount2() As Double
If gtc = 1 Then
GetTickCount2 = GetTickCount
Exit Function
End If
Dim curC As Currency
If curFreq = 0 Then
    QueryPerformanceFrequency curFreq   'Get the timer frequency
    curFreq = curFreq * 10 'in ms
End If
QueryPerformanceCounter curC
GetTickCount2 = (curC * 10000) / curFreq
End Function


