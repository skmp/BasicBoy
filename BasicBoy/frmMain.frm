VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BasicBoy - v[version]"
   ClientHeight    =   5025
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4080
      Left            =   0
      Picture         =   "frmMain.frx":058A
      ScaleHeight     =   272
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4545
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4080
         Top             =   240
      End
   End
   Begin VB.PictureBox sbp 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4965
      TabIndex        =   1
      Top             =   4770
      Width           =   4965
      Begin VB.Label sb 
         Caption         =   "Runing @ xxx.xx"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.Image Image1 
      Height          =   3930
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4260
   End
   Begin VB.Menu fl 
      Caption         =   "File"
      Begin VB.Menu starte 
         Caption         =   "Load ROM..."
      End
      Begin VB.Menu res 
         Caption         =   "Reset ROM..."
      End
      Begin VB.Menu ris 
         Caption         =   "ROM Information"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu lstate 
         Caption         =   "Load State"
         Begin VB.Menu loadstateslot 
            Caption         =   "Slot #1"
            Index           =   0
         End
      End
      Begin VB.Menu sstate 
         Caption         =   "Save State"
         Begin VB.Menu savestateslot 
            Caption         =   "Slot #1"
            Index           =   0
         End
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu ebdf 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu bn 
      Caption         =   "Options"
      Begin VB.Menu mnuEmulatedHardware 
         Caption         =   "Emulated Hardware"
         Begin VB.Menu EGBC 
            Caption         =   "GameBoy Color"
         End
         Begin VB.Menu ESE 
            Caption         =   "GameBoy Sound"
         End
         Begin VB.Menu el 
            Caption         =   "Gameboy Link"
         End
      End
      Begin VB.Menu fs 
         Caption         =   "Frame Skip"
         Begin VB.Menu fs0 
            Caption         =   "No Skip"
         End
         Begin VB.Menu fs6m1 
            Caption         =   "set 1.20x"
         End
         Begin VB.Menu fs5m1 
            Caption         =   "set 1.25x"
         End
         Begin VB.Menu fs4m1 
            Caption         =   "set 1.3x"
         End
         Begin VB.Menu fs3m1 
            Caption         =   "set 1.5x"
         End
         Begin VB.Menu fs1 
            Caption         =   "set x2"
         End
         Begin VB.Menu fs2 
            Caption         =   "set x3"
         End
         Begin VB.Menu fs3 
            Caption         =   "set x4"
         End
         Begin VB.Menu fs4 
            Caption         =   "set x5"
         End
         Begin VB.Menu fs9 
            Caption         =   "set x10"
         End
      End
      Begin VB.Menu md 
         Caption         =   "Render Method"
         Begin VB.Menu WA 
            Caption         =   "WinApi"
         End
         Begin VB.Menu DD7 
            Caption         =   "DirectDraw"
         End
      End
      Begin VB.Menu rz 
         Caption         =   "Resolution"
         Begin VB.Menu z1 
            Caption         =   "1x"
            Checked         =   -1  'True
         End
         Begin VB.Menu z2 
            Caption         =   "2x"
            Checked         =   -1  'True
         End
         Begin VB.Menu z3 
            Caption         =   "3x"
         End
         Begin VB.Menu z4 
            Caption         =   "4x"
         End
         Begin VB.Menu sep2 
            Caption         =   "-"
         End
         Begin VB.Menu full 
            Caption         =   "Fullscreen"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu sh 
         Caption         =   "Layers"
         Begin VB.Menu sbg 
            Caption         =   "BG"
            Checked         =   -1  'True
         End
         Begin VB.Menu swn 
            Caption         =   "Window"
            Checked         =   -1  'True
         End
         Begin VB.Menu sobj 
            Caption         =   "OBJs"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSoundBuffer 
         Caption         =   "Sound Buffer"
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "2 Milliseconds"
            Index           =   0
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "4 Milliseconds"
            Index           =   1
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "6 Milliseconds"
            Index           =   2
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "8 Milliseconds"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "10 Milliseconds"
            Index           =   4
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "12 Milliseconds"
            Index           =   5
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "14 Milliseconds"
            Index           =   6
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "16 Milliseconds"
            Index           =   7
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "32 Milliseconds"
            Index           =   8
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "64 Milliseconds"
            Index           =   9
         End
         Begin VB.Menu mnuSndBufferSize 
            Caption         =   "128 Milliseconds"
            Index           =   10
         End
      End
      Begin VB.Menu mnuSoundRenderMethod 
         Caption         =   "Sound Render Method"
         Begin VB.Menu mnuSoundRender 
            Caption         =   "Hardware"
            Index           =   0
         End
         Begin VB.Menu mnuSoundRender 
            Caption         =   "Software (Recommended)"
            Index           =   1
         End
      End
      Begin VB.Menu cc 
         Caption         =   "Cheat Codes..."
      End
      Begin VB.Menu cjoy 
         Caption         =   "Modify Controls..."
      End
      Begin VB.Menu fsno 
         Caption         =   "Fast Sound off"
      End
      Begin VB.Menu fsko 
         Caption         =   "Fast skip on"
      End
      Begin VB.Menu lfps 
         Caption         =   "Limit FPS"
      End
      Begin VB.Menu fd 
         Caption         =   "Speed Options"
      End
      Begin VB.Menu cr 
         Caption         =   "CPU Core"
         Visible         =   0   'False
         Begin VB.Menu cr1 
            Caption         =   "Mode 1"
            Checked         =   -1  'True
         End
         Begin VB.Menu cr2 
            Caption         =   "Mode 2"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu dbv 
         Caption         =   "View Debug Window"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu Abt 
         Caption         =   "About..."
      End
      Begin VB.Menu mnuHelpDoc 
         Caption         =   "Help Documentation..."
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "BasicBoy Website..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dih As clsDirectInput8
Public di As clsDIKeyboard8
Option Explicit
#Const msg = 1

Private Sub Abt_Click()
    frmAbout.Show
End Sub

Private Sub cc_Click()
frmCheat.Show
End Sub

Private Sub cjoy_Click()
frmJoy.confkeys
End Sub

Private Sub cr1_Click()
cr1.Checked = True
cr2.Checked = False
cm = 1
SaveSetting "BasicBoy", "CPU", "CM", 1
End Sub

Private Sub cr2_Click()
cr1.Checked = False
cr2.Checked = True
cm = 2
SaveSetting "BasicBoy", "CPU", "CM", 2
End Sub

Private Sub DD7_Click()
full.Visible = True
WA.Checked = False
DD7.Checked = True
mm = 1
SaveSetting "BasicBoy", "GFX", "MM", 2
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
End Sub



Private Sub ebdf_Click()
#If 0 Then
Open App.Path & "\lol.b" For Binary As #1
Put #1, , ic
Close #1
#End If
wrRam
On Error Resume Next
#If 0 Then
DDraw.Shutdown
#End If
di.Shutdown
link_kill
Me.Hide
Unload Me
End
End Sub

Private Sub EGBC_Click()
TGBC = Not TGBC
EGBC.Checked = TGBC
SaveSetting "BasicBoy", "GBC", "EMU", TGBC
End Sub


Private Sub el_Click()
Con
End Sub

Private Sub ESE_Click()
    If snd = 1 Then
        snd = 0
        ESE.Checked = False
    Else
        snd = 1
        ESE.Checked = True
        initWave
    End If
    SaveSetting "BasicBoy", "Snd", "en", snd
End Sub

Private Sub fd_Click()
frmSpeed.Show
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub

Private Sub Form_Load()
Dim tmpcur As Currency
Call InitCommonControls
Me.Caption = "BasicBoy - v" & App.Major & "." & App.Minor & "." & App.Revision
Dim StateSlots As Integer
For StateSlots = 1 To 7
    load loadstateslot(StateSlots)
    loadstateslot(StateSlots).Caption = "Slot #" & (StateSlots + 1)
    load savestateslot(StateSlots)
    savestateslot(StateSlots).Caption = "Slot #" & (StateSlots + 1)
Next StateSlots
Cpu_Speed = 1
gtc = GetSetting("BasicBoy", "misc", "timermode", "0")
If gtc = 1 Then framedelay = 15
If QueryPerformanceFrequency(tmpcur) = 0 Then
gtc = 1 'not supported
framedelay = 15
End If

If App.LogMode = 0 Then
    MsgBox "Compile to EXE for GREATLY increased speed and sound emulation", vbCritical, "Alert"
End If
If App.LogMode = 1 Then
If GetSetting("BasicBoy", "snd", "sofwaresnd", 1) = 1 Then
    ssound = 1
    mnuSoundRender(1).Checked = True
Else
    ssound = 0
    mnuSoundRender(0).Checked = True
End If

mnuSndBufferSize_Click GetSetting("BasicBoy", "snd", "soundbuf", 3)
If mnuSndBufferSize.Item(0).Checked = True Then Init_Sound 2
If mnuSndBufferSize.Item(1).Checked = True Then Init_Sound 4
If mnuSndBufferSize.Item(2).Checked = True Then Init_Sound 6
If mnuSndBufferSize.Item(3).Checked = True Then Init_Sound 8
If mnuSndBufferSize.Item(4).Checked = True Then Init_Sound 10
End If

#If dver Then
Open "db.txt" For Binary As #90
#End If
ns:
#If msg Then
If GetSetting("BasicBoy", "misc", "fistrun", 0) = 0 Then
    If MsgBox("This project is the result of many hours of work " & vbNewLine & _
              "Please vote", vbOKCancel) = vbOK Then
              Shell pschome
    End If
End If
#End If

If App.LogMode = 0 Then snd = False Else snd = True
InitDirectX Me.hwnd, Me.Picture1.hwnd
'On Error Resume Next
If err.Number <> 0 Then snd = False: err.Clear
Me.Width = 4950
Me.Height = 5415
Picture1.Width = Me.Width - 110
Picture1.Height = Me.Height - 800
bgv = True
wv = True
objv = True
smp = 0
Dim i As Long
DD7_Click
cm = GetSetting("BasicBoy", "CPU", "CM", 1)
ESE.Checked = GetSetting("BasicBoy", "Snd", "en", "1")
stpsk = GetSetting("BasicBoy", "FSC", "sfk", "1")
stpsnd = GetSetting("BasicBoy", "FSC", "snd", "1")
fsko_Click
fsko_Click
fsno_Click
fsno_Click
snd = GetSetting("BasicBoy", "Snd", "en", "1")
If App.LogMode = 0 Then snd = False
If cm = 1 Then: cr1.Checked = True: cr2.Checked = False: Else cr1.Checked = False: cr2.Checked = True
If GetSetting("BasicBoy", "GFX", "ZM", "2") = 1 Then z1_Click
If GetSetting("BasicBoy", "GFX", "ZM", "2") = 2 Then z2_Click
If GetSetting("BasicBoy", "GFX", "ZM", "2") = 3 Then z3_Click
If GetSetting("BasicBoy", "GFX", "ZM", "2") = 4 Then z4_Click
If GetSetting("BasicBoy", "GFX", "ZM", "2") = 5 Then full_Click
lfp = GetSetting("BasicBoy", "CPU", "LFPS", "1")
fskip = GetSetting("BasicBoy", "GFX", "FS", "1")
fmode = GetSetting("BasicBoy", "GFX", "FM", "0")
TGBC = GetSetting("BasicBoy", "GBC", "EMU", "True"): EGBC_Click: EGBC_Click
fs0.Checked = fskip = 1 And fmode = 0
fs6m1.Checked = fskip = 6 And fmode = 1
fs5m1.Checked = fskip = 5 And fmode = 1
fs4m1.Checked = fskip = 4 And fmode = 1
fs3m1.Checked = fskip = 3 And fmode = 1
fs1.Checked = fskip = 2 And fmode = 0
fs2.Checked = fskip = 3 And fmode = 0
fs3.Checked = fskip = 4 And fmode = 0
fs4.Checked = fskip = 5 And fmode = 0
fs9.Checked = fskip = 10 And fmode = 0
Up = GetSetting("BasicBoy", "Joy", "up", "200")
Dn = GetSetting("BasicBoy", "Joy", "dn", "208")
Lf = GetSetting("BasicBoy", "Joy", "lf", "203")
Rg = GetSetting("BasicBoy", "Joy", "rg", "205")
ABut = GetSetting("BasicBoy", "Joy", "ab", "44")
BBut = GetSetting("BasicBoy", "Joy", "bb", "45")
St1 = GetSetting("BasicBoy", "Joy", "st1", "28"): St2 = GetSetting("BasicBoy", "Joy", "st2", "42"): St3 = GetSetting("BasicBoy", "Joy", "st3", "54")
Sl1 = GetSetting("BasicBoy", "Joy", "sl1", "32"): Sl2 = GetSetting("BasicBoy", "Joy", "sl2", "2")
SpeedKeyD = GetSetting("BasicBoy", "Joy", "spdd", "15"): SpeedKeyU = GetSetting("BasicBoy", "Joy", "spdu", "15")
lfps.Checked = lfp
'gxmode2
Set dih = New clsDirectInput8
Set di = New clsDIKeyboard8
   dih.Startup Me.hwnd
   di.Startup dih, Me.hwnd
   InitCPU
   For i = 0 To 7 ' For Set,Bit and Res
   BITT(i * 8) = 2 ^ i
   SETT(i * 8) = 255 - 2 ^ i
   Next i
   Me.Show
   On Error Resume Next
   initWave
slfp = lfp
SaveSetting "BasicBoy", "misc", "fistrun", 1
End Sub



Private Sub Form_Unload(Cancel As Integer)
wrRam
#If dver Then
Close #99
#End If
On Error Resume Next
#If 0 Then
DDraw.Shutdown
#End If
di.Shutdown
link_kill
End
End Sub


Sub resize()
Me.Width = Me.Picture1.Width + 105
Me.Height = Me.Picture1.Height + 1095
Me.Image1.Width = 15 * 160 * zm
Me.Image1.Height = 15 * 144 * zm
End Sub

'frame skip mode 1(act skip(x1(1),x2(2),x3(3),x4(4),x5(5),x6(6))
'frame skip mode 2(act skip(x1.20(6),x1.25(5),x1,3(4),x1.5(3))
Private Sub fs0_Click()
fs0.Checked = True
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = False
fs3.Checked = False
fs4.Checked = False
fs9.Checked = False
fskip = 1
fmode = 0
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs1_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = True
fs2.Checked = False
fs3.Checked = False
fs4.Checked = False
fs9.Checked = False
fskip = 2
fmode = 0
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs2_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = True
fs3.Checked = False
fs4.Checked = False
fs9.Checked = False
fskip = 3
fmode = 0
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs3_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = False
fs3.Checked = True
fs4.Checked = False
fs9.Checked = False
fskip = 4
fmode = 0
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs4_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = False
fs3.Checked = False
fs4.Checked = True
fs9.Checked = False
fskip = 5
fmode = 0
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs9_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = False
fs3.Checked = False
fs4.Checked = False
fs9.Checked = True
fskip = 10
fmode = 0
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs6m1_Click()
fs0.Checked = False
fs6m1.Checked = True
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = False
fs3.Checked = False
fs4.Checked = False
fs9.Checked = False
fskip = 6
fmode = 1
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs5m1_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = True
fs4m1.Checked = False
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = False
fs3.Checked = False
fs4.Checked = False
fs9.Checked = False
fskip = 5
fmode = 1
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs4m1_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = True
fs3m1.Checked = False
fs1.Checked = False
fs2.Checked = False
fs3.Checked = False
fs4.Checked = False
fs9.Checked = False
fskip = 4
fmode = 1
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub
Private Sub fs3m1_Click()
fs0.Checked = False
fs6m1.Checked = False
fs5m1.Checked = False
fs4m1.Checked = False
fs3m1.Checked = True
fs1.Checked = False
fs2.Checked = False
fs3.Checked = False
fs4.Checked = False
fs9.Checked = False
fskip = 3
fmode = 1
SaveSetting "BasicBoy", "GFX", "FS", fskip
SaveSetting "BasicBoy", "GFX", "FM", fmode
End Sub

Private Sub fsko_Click()
stpsk = Not stpsk
SaveSetting "BasicBoy", "FSC", "sfk", stpsk
fsko.Checked = stpsk
End Sub

Private Sub fsno_Click()
stpsnd = Not stpsnd
SaveSetting "BasicBoy", "FSC", "snd", stpsnd
fsno.Checked = stpsnd
End Sub

Private Sub lfps_Click()
lfp = Not lfp
lfps.Checked = lfp
SaveSetting "BasicBoy", "CPU", "LFPS", lfp
End Sub

Private Sub loadstateslot_Click(Index As Integer)
    loadState Index
End Sub

Private Sub mnuHelpDoc_Click()
    Call Shell("explorer " & App.Path & "\help.html", vbNormalFocus)
End Sub

Private Sub mnuSndBufferSize_Click(Index As Integer)
    Dim i As Integer
    SaveSetting "BasicBoy", "snd", "soundbuf", Index
    If Index = 0 Then Init_Sound 2
    If Index = 1 Then Init_Sound 4
    If Index = 2 Then Init_Sound 6
    If Index = 3 Then Init_Sound 8
    If Index = 4 Then Init_Sound 10
    If Index = 5 Then Init_Sound 12
    If Index = 6 Then Init_Sound 14
    If Index = 7 Then Init_Sound 16
    If Index = 8 Then Init_Sound 32
    If Index = 9 Then Init_Sound 64
    If Index = 10 Then Init_Sound 128
    For i = 0 To mnuSndBufferSize.UBound
        mnuSndBufferSize(i).Checked = False
    Next i
    mnuSndBufferSize(Index).Checked = True
End Sub

Private Sub mnuSoundRender_Click(Index As Integer)
    Dim i As Integer
    SaveSetting "BasicBoy", "snd", "sofwaresnd", Index
    For i = 0 To mnuSoundRender.UBound
        mnuSoundRender(i).Checked = False
    Next i
    mnuSoundRender(Index).Checked = True

If GetSetting("BasicBoy", "snd", "sofwaresnd", 1) = 1 Then
    ssound = 1
    mnuSoundRender(1).Checked = True
Else
    ssound = 0
    mnuSoundRender(0).Checked = True
End If
If mnuSndBufferSize.Item(0).Checked = True Then Init_Sound 2
If mnuSndBufferSize.Item(1).Checked = True Then Init_Sound 4
If mnuSndBufferSize.Item(2).Checked = True Then Init_Sound 6
If mnuSndBufferSize.Item(3).Checked = True Then Init_Sound 8
If mnuSndBufferSize.Item(4).Checked = True Then Init_Sound 10

End Sub

Private Sub mnuWebsite_Click()
    Call Shell("explorer http://basicboy.emuhost.com/", vbNormalFocus)
End Sub



Private Sub res_Click()
reset
End Sub

Public Sub rp_Click()
RunCpu
End Sub


Private Sub ris_Click()
frmRomInfo.Show
End Sub

Private Sub savestateslot_Click(Index As Integer)
    saveState Index
End Sub


Private Sub sbg_Click()
bgv = Not bgv
sbg.Checked = bgv
End Sub



Private Sub sbp_Resize()
sb.Width = sbp.Width
sb.Height = sbp.Height
End Sub

Private Sub sobj_Click()
objv = Not objv
sobj.Checked = objv
End Sub

Private Sub starte_Click()
Dim strtemp As String, bol As Boolean, tls As String, i As Long
If loadrom(CD.ShowOpen(Me.hwnd, "", CD.FileName, "GameBoy Roms (*.gb;*.gbc;*.cgb)|*.gb;*.gbc;*.cgb")) Then
initCI
rdRam
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
If TGBC Then
If ROM(&H143, 0) = 192 Then strtemp = "(GBC) ": GBM = 1 Else If ROM(&H143, 0) <> 0 Then strtemp = "(GB/GBC) ": GBM = 1 Else strtemp = "(GB) ": GBM = 0
Else
If ROM(&H143, 0) = 192 Then strtemp = "(GBC) " Else If ROM(&H143, 0) <> 0 Then strtemp = "(GB/GBC) " Else strtemp = "(GB) "
GBM = 0
End If
resize
reset
tls = ""
For i = 0 To 15
If rominfo.titleB(i) = 0 Then GoTo tiend
tls = tls & Chr(rominfo.titleB(i))
Next i
tiend:
rominfo.Title = tls
frmMain.Caption = "BasicBoy - " & tls & " " & strtemp
initWave
rp_Click
End If
err.Clear
End Sub

Private Sub swn_Click()
wv = Not wv
swn.Checked = wv
End Sub


Private Sub Timer1_Timer()
Dim mestr$, tmp As Single, tmp2 As Single, mipsd As Single, frmS As Single
If fpsT = 0 Then fpsT = GetTickCount2: Exit Sub
If bCpuRun Then mestr = "on" Else mestr = "off"
tmp = (GetTickCount2 - fpsT) / 1000
tmp = CLng(modGrfx.FPS / tmp)
tmp2 = (GetTickCount2 - fpsT) / 1000
mipsd = Mips / tmp2
tmp2 = (Mhz * 70224 + Clcount) / tmp2
frmS = tmp2 / 70224 / (CpuS + 1)
tmp2 = format$(tmp2 / 1024 / 1024, "000.000")
sb.Caption = "CPU Running @ " & tmp2 & " Mhz(" & format$(mipsd / 1024 / 1024, "0.000") & " MIPS), FPS " & format$(frmS, "00.00") & " (" & format(frmS / 59.7, "00.0%") & ")"
modGrfx.FPS = 0
fpsT = GetTickCount2
Mhz = 0
Mips = 0
check_link_connection
'Timer1.Enabled = False
End Sub

Private Sub WA_Click()
WA.Checked = True
DD7.Checked = False
full.Visible = False
mm = 2
SaveSetting "BasicBoy", "GFX", "MM", 1
If zm = 5 Then zm = 4: SaveSetting "BasicBoy", "GFX", "ZM", 4: z4_Click
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
End Sub

Private Sub z1_Click()
z1.Checked = True
z2.Checked = False
z3.Checked = False
z4.Checked = False
full.Checked = False
SaveSetting "BasicBoy", "GFX", "ZM", 1
zm = 1
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
resize
End Sub

Private Sub z2_Click()
z1.Checked = False
z2.Checked = True
z3.Checked = False
z4.Checked = False
full.Checked = False
SaveSetting "BasicBoy", "GFX", "ZM", 2
zm = 2
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
resize
End Sub
Private Sub full_Click()
z1.Checked = False
z2.Checked = False
z3.Checked = False
z4.Checked = False
full.Checked = True
SaveSetting "BasicBoy", "GFX", "ZM", 5
zm = 5
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
resize
End Sub

Private Sub z3_Click()
z1.Checked = False
z2.Checked = False
z3.Checked = True
z4.Checked = False
full.Checked = False
SaveSetting "BasicBoy", "GFX", "ZM", 3
zm = 3
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
resize
End Sub

Private Sub z4_Click()
z1.Checked = False
z2.Checked = False
z3.Checked = False
z4.Checked = True
full.Checked = False
SaveSetting "BasicBoy", "GFX", "ZM", 4
zm = 4
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, full.Checked
End If
resize
End Sub
