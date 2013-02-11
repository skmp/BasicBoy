VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDik2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'local references
Dim dinput As clsDirectInput8

'class variables
Dim diDev_Keyboard As DirectInputDevice8
Dim diState_Keyboard As DxVBLibA.DIKEYBOARDSTATE
Dim diOState_Keyboard As DxVBLibA.DIKEYBOARDSTATE
Dim sKeyNames(255) As String

Public Sub Startup(ByRef di As clsDirectInput8, ByVal hWindowHandle As Long)
    'create reference to DirectInput object
    Set dinput = di
    
    'create an interface to the keyboard
    Set diDev_Keyboard = dinput.DIObj.CreateDevice("GUID_SysKeyboard")
    
    diDev_Keyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
    diDev_Keyboard.SetCooperativeLevel hWindowHandle, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    diDev_Keyboard.Acquire

    'initialize the key name array
    InitKeyNames
End Sub

Public Sub Check_Keyboard()
    Dim n As Long
    If GetActiveWindow = frmJoy.hWnd Then
    'make sure the keyboard handler is running
    If diDev_Keyboard Is Nothing Then Exit Sub
    
    'get the list of pressed keys
    diDev_Keyboard.GetDeviceStateKeyboard diState_Keyboard
    
    'scan the entire list for pressed keys
    For n = 0 To 255
        If diState_Keyboard.key(n) <> diOState_Keyboard.key(n) Then
            If diState_Keyboard.key(n) <> 0 Then
            If diOState_Keyboard.key(n) = 0 Then
            frmJoy.ckeydown n
            End If
            End If
        End If
    Next
    End If
    diOState_Keyboard = diState_Keyboard
End Sub

Public Sub Shutdown()
    On Local Error Resume Next

    'release the keyboard
    If Not (diDev_Keyboard Is Nothing) Then
        diDev_Keyboard.Unacquire
        Set diDev_Keyboard = Nothing
    End If
End Sub

Public Property Get KeyName(ByVal lKey As Long) As String
    KeyName = sKeyNames(lKey)
End Property

Private Sub InitKeyNames()
    sKeyNames(1) = "ESC"
    sKeyNames(2) = "1"
    sKeyNames(3) = "2"
    sKeyNames(4) = "3"
    sKeyNames(5) = "4"
    sKeyNames(6) = "5"
    sKeyNames(7) = "6"
    sKeyNames(8) = "7"
    sKeyNames(9) = "8"
    sKeyNames(10) = "9"
    sKeyNames(11) = "0"
    sKeyNames(12) = "-"
    sKeyNames(13) = "="
    sKeyNames(14) = "BACKSPACE"
    sKeyNames(15) = "TAB"
    sKeyNames(16) = "Q"
    sKeyNames(17) = "W"
    sKeyNames(18) = "E"
    sKeyNames(19) = "R"
    sKeyNames(20) = "T"
    sKeyNames(21) = "Y"
    sKeyNames(22) = "U"
    sKeyNames(23) = "I"
    sKeyNames(24) = "O"
    sKeyNames(25) = "P"
    sKeyNames(26) = "["
    sKeyNames(27) = " ]"
    sKeyNames(28) = "ENTER"
    sKeyNames(29) = "LCTRL"
    sKeyNames(30) = "A"
    sKeyNames(31) = "S"
    sKeyNames(32) = "D"
    sKeyNames(33) = "F"
    sKeyNames(34) = "G"
    sKeyNames(35) = "H"
    sKeyNames(36) = "J"
    sKeyNames(37) = "K"
    sKeyNames(38) = "L"
    sKeyNames(39) = ";"
    sKeyNames(40) = "'"
    sKeyNames(41) = "`"
    sKeyNames(42) = "LSHIFT"
    sKeyNames(43) = "\"
    sKeyNames(44) = "Z"
    sKeyNames(45) = "X"
    sKeyNames(46) = "C"
    sKeyNames(47) = "V"
    sKeyNames(48) = "B"
    sKeyNames(49) = "N"
    sKeyNames(50) = "M"
    sKeyNames(51) = ","
    sKeyNames(52) = "."
    sKeyNames(53) = "/"
    sKeyNames(54) = "RSHIFT"
    sKeyNames(55) = "NUMPAD*"
    sKeyNames(56) = "LALT"
    sKeyNames(57) = "SPACE"
    sKeyNames(58) = "CAPSLOCK"
    sKeyNames(59) = "F1"
    sKeyNames(60) = "F2"
    sKeyNames(61) = "F3"
    sKeyNames(62) = "F4"
    sKeyNames(63) = "F5"
    sKeyNames(64) = "F6"
    sKeyNames(65) = "F7"
    sKeyNames(66) = "F8"
    sKeyNames(67) = "F9"
    sKeyNames(68) = "F10"
    sKeyNames(69) = "NUMLOCK"
    sKeyNames(70) = "SCRLLOCK"
    sKeyNames(71) = "NUMPAD7"
    sKeyNames(72) = "NUMPAD8"
    sKeyNames(73) = "NUMPAD9"
    sKeyNames(74) = "NUMPAD-"
    sKeyNames(75) = "NUMPAD4"
    sKeyNames(76) = "NUMPAD5"
    sKeyNames(77) = "NUMPAD6"
    sKeyNames(78) = "NUMPAD+"
    sKeyNames(79) = "NUMPAD1"
    sKeyNames(80) = "NUMPAD2"
    sKeyNames(81) = "NUMPAD3"
    sKeyNames(82) = "NUMPAD0"
    sKeyNames(83) = "NUMPAD."
    sKeyNames(87) = "F11"
    sKeyNames(88) = "F12"
    sKeyNames(86) = "F13"
    sKeyNames(84) = "F14"
    sKeyNames(85) = "F15"
    sKeyNames(156) = "NUMPADENTER"
    sKeyNames(157) = "RCONTROL"
    sKeyNames(91) = "NUMPAD,"
    sKeyNames(181) = "NUMPAD/"
    sKeyNames(183) = "SYSRQ"
    sKeyNames(184) = "RALT"
    sKeyNames(199) = "HOME"
    sKeyNames(200) = "UP"
    sKeyNames(201) = "PGUP"
    sKeyNames(203) = "LEFT"
    sKeyNames(205) = "RIGHT"
    sKeyNames(207) = "END"
    sKeyNames(208) = "DOWN"
    sKeyNames(209) = "PGDN"
    sKeyNames(210) = "INSERT"
    sKeyNames(211) = "DELETE"
    sKeyNames(219) = "LWIN"
    sKeyNames(220) = "RWIN"
    sKeyNames(221) = "APPS"
    sKeyNames(116) = "PAUSE"
End Sub




