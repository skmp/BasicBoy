Attribute VB_Name = "Declares"
'Genral Declares/Const/Types
'(mainly WinApi Declares/Const/Types)
'You can use this in any program
'as long as you do not take credit for it

'This module contains many things that are
'not used in basicboy.I use this Module in many
'projects so it will not be "cleaned up" (:P)


Option Explicit
Option Base 0


'Windows API constants
Public Const WM_CLOSE As Long = &H10
Public Const BITSPIXEL As Long = 12
Public Const SYSTEM_FONT  As Long = 13
Public Const LR_LOADFROMFILE  As Long = &H10
Public Const CAPS1  As Long = 94
Public Const C1_TRANSPARENT  As Long = &H1
Public Const NEWTRANSPARENT As Long = 3
Public Const PM_NOREMOVE As Long = &H0
Public Const PM_REMOVE  As Long = &H1
Public Const WM_QUIT As Long = &H12
Public Const OFN_ALLOWMULTISELECT   As Long = &H200
Public Const OFN_CREATEPROMPT       As Long = &H2000
Public Const OFN_EXPLORER           As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST      As Long = &H1000
Public Const OFN_HIDEREADONLY       As Long = &H4
Public Const OFN_LONGNAMES          As Long = &H200000
Public Const OFN_NOCHANGEDIR        As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_OVERWRITEPROMPT    As Long = &H2
Public Const OFN_PATHMUSTEXIST      As Long = &H800
Public Const OFN_READONLY           As Long = &H1

Public Const CALLBACK_WINDOW As Long = &H10000
Public Const MMIO_READ  As Long = &H0
Public Const MMIO_FINDCHUNK  As Long = &H10
Public Const MMIO_FINDRIFF  As Long = &H20
Public Const MM_WOM_DONE  As Long = &H3BD
Public Const MMSYSERR_NOERROR  As Long = 0
Public Const SEEK_CUR As Long = 1
Public Const SEEK_END As Long = 2
Public Const SEEK_SETv = 0
Public Const TIME_BYsnd  As Long = &H4
Public Const WHDR_DONE As Long = &H1
Public Const GWL_WNDPROC As Long = -4

'Windows API structures
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type


Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Type mmioinfo
        dwFlags As Long
        fccIOProc As Long
        pIOProc As Long
        wErrorRet As Long
        htask As Long
        cchBuffer As Long
        pchBuffer As String
        pchNext As String
        pchEndRead As String
        pchEndWrite As String
        lBufOffset As Long
        lDiskOffset As Long
        adwInfo(4) As Long
        dwReserved1 As Long
        dwReserved2 As Long
        hmmio As Long
End Type

Type WAVEHDR
        lpData As Long
        dwBufferLength As Long
        dwBytesRecorded As Long
        dwUser As Long
        dwFlags As Long
        dwLoops As Long
        lpNext As Long
        Reserved As Long
End Type

Type WAVEINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * 32
        dwFormats As Long
        wChannels As Integer
End Type

Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wBitsPerSample As Integer
        cbSize As Integer
End Type

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Type MMTIME
        wType As Long
        u As Long
        X As Long
End Type

Public Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

Public Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Public Type BITMAPINFO
  Header As BITMAPINFOHEADER
  bits() As Byte
End Type

Public Type BITMAP_STRUCT
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type RECT_API
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type point
    X As Long
    Y As Long
End Type

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    point As point
End Type

Public Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    alpha As Byte
End Type
'Memory Emulation
Public Type CartIinfo
    NGraph(47) As Byte
    Title As String
    titleB(15) As Byte
    GBC As Byte
    lcode(1) As Byte
    isSGB As Byte
    Ctype As Byte
    romsize As Byte
    ramsize As Byte
    DestCode As Byte
    lcodeold As Byte
    MaskRomV As Byte
    CCheck As Byte
    Checksum(1) As Byte
End Type

'***Sound chip***
'**Enums**
Public Enum Sound_Pars
    '*Enabled*
    Enabled = 1
    Disabled = 0
    '*Wave pattern duty*
    Wave_Pattern_1_8 = 0
    Wave_Pattern_2_8 = 1
    Wave_Pattern_4_8 = 2
    Wave_Pattern_6_8 = 3
    '*volume*
    Wave_Pattern_Volume_Mute = 0
    Wave_Pattern_Volume_No_Change = 1
    Wave_Pattern_Volume_Half = 2
    Wave_Pattern_Volume_Quarter = 3
    '*Sound left/right output*
    Sound_Chanel_None = 0
    Sound_Chanel_Left = 1
    Sound_Chanel_Right = 2
    
    No_Par = 0
End Enum

'*Commands*
Public Enum Sound_CMD
    'Standard cmds,GB
    Wave_Frequency_Set = 1 'Set frequency
    Wave_Pattern_Duty_Set = 2 '0-3 /Chanel 1,2
    Wave_Volume_Set = 3 '0-15 volume
    Wave_Reset = 4 'no param , reset
    Wave_Pattern_Volume_Set = 5 '0-3/chanel 3
    Wave_SO1_Level_Set = 6
    Wave_SO2_Level_Set = 7
    Wave_Output_To_SOx_Set = 8 'none,left,right
    Wave_Chanel_Enabled_Set = 9
    'Chanel 4 cmds
    Wave_Polynomial_shift_clock_Frequency = 10
    Wave_Polynomial_Step = 11
    Wave_Polynomial_Dividing_Ratio = 12
    'Internal
    sound_play = 13 'No param
    Sound_Stop = 14 'No param
    'chanel 3
    wave_waveform_write = 15
End Enum

Public Enum Sound_Chans
    Chanel1 = 1
    Chanel2 = 2
    Chanel3 = 3
    Chanel4 = 4
End Enum

'**Types**
Public Type SoundCommand
    en As Long 'enabled
    chan As Long 'chanel to aply
    cmd As Long 'command code
    param As Long 'value to set
    pos As Long 'position to aply
End Type

Public Type SoundCommands
    sc() As SoundCommand
    ind_m As Long
End Type

Public Type SoundCD12
    Duty As Long
    Volume As Single
    Count As Single
    MCount As Single
    Current As Long
    Index As Long
    Enabled As Byte 'not used
    Left As Byte 'not used - waiting for stereo emulation
    Right As Byte 'not used - waiting for stereo emulation
    Play As Byte
End Type

Public Type SoundCD3
    Shift As Long
    Volume As Single
    Count As Single
    MCount As Single
    Current As Long
    Index As Long
    Enabled As Byte 'not used
    Left As Byte 'not used - waiting for stereo emulation
    Right As Byte 'not used - waiting for stereo emulation
    Waveform(31) As Byte
    Play As Byte
End Type

Public Type SoundCD4
    bits As Long
    Volume As Single
    Count As Single
    MCount As Single
    Current As Long
    Index As Long
    Enabled As Byte 'not used
    Left As Byte 'not used - waiting for stereo emulation
    Right As Byte 'not used - waiting for stereo emulation
    Play As Byte
End Type

'Windows API functions
'wave_out/ not used
Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Declare Function waveOutOpen Lib "winmm.dll" (hWaveOut As Long, ByVal uDeviceID As Long, ByVal format As String, ByVal dwCallback As Long, ByRef fPlaying As Boolean, ByVal dwFlags As Long) As Long
Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
'mmio / not used
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadString Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Declare Sub CopyStructFromString Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal source As String, ByVal cb As Long)
Declare Function PostWavMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef hdr As WAVEHDR) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByRef lParam As WAVEHDR) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Declare Function BitBlt _
    Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long _
) As Long

Public Declare Function CloseWindow _
    Lib "user32" ( _
    ByVal hwnd As Long _
) As Long

Public Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long _
) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessID As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Declare Function SetDIBitsToDevice _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal DX As Long, _
    ByVal dy As Long, _
    ByVal SrcX As Long, _
    ByVal SrcY As Long, _
    ByVal Scan As Long, _
    ByVal NumScans As Long, _
    lpBits As Any, _
    lpBI As BITMAPINFO, _
    ByVal wUsage As Long _
) As Long

Declare Function PeekMessage _
    Lib "user32" Alias "PeekMessageA" ( _
    lpMsg As msg, _
    ByVal hwnd As Long, _
    ByVal wMsgFilterMin As Long, _
    ByVal wMsgFilterMax As Long, _
    ByVal wRemoveMsg As Long _
) As Long

Declare Function GetMessage _
    Lib "user32" Alias "GetMessageA" ( _
    lpMsg As msg, ByVal hwnd As Long, _
    ByVal wMsgFilterMin As Long, _
    ByVal wMsgFilterMax As Long _
) As Long

Declare Function TranslateMessage _
    Lib "user32" ( _
    lpMsg As msg _
) As Long

Declare Function DispatchMessage _
    Lib "user32" Alias "DispatchMessageA" ( _
    lpMsg As msg _
) As Long

Public Declare Sub CopyMemory _
    Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, _
    lpvSource As Any, _
    ByVal cbCopy As Long _
)

Public Declare Function QueryPerformanceFrequency _
    Lib "kernel32" ( _
    lpFrequency As Currency _
) As Long

Public Declare Function QueryPerformanceCounter _
    Lib "kernel32" ( _
    lpPerformanceCount As Currency _
) As Long


Public Declare Function CreateCompatibleBitmap _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long _
) As Long

Public Declare Function CreateCompatibleDC _
    Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Public Declare Function CreatePen _
    Lib "gdi32" ( _
    ByVal nPenStyle As Long, _
    ByVal nWidth As Long, _
    ByVal crColor As Long _
) As Long

Public Declare Function DeleteDC _
    Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Public Declare Function DeleteObject _
    Lib "gdi32" ( _
    ByVal hObject As Long _
) As Long

Public Declare Function Ellipse _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long _
) As Long

Public Declare Function GetBitmapBits _
    Lib "gdi32" ( _
    ByVal hBitmap As Long, _
    ByVal dwCount As Long, _
    lpBits As Any _
) As Long

Public Declare Sub Sleep _
    Lib "kernel32" ( _
    ByVal dwMilliseconds As Long _
)

Public Declare Function GetClientRect _
    Lib "user32" ( _
    ByVal hwnd As Long, _
    lpRect As RECT_API _
) As Long

Public Declare Function GetDC _
    Lib "user32" ( _
    ByVal hwnd As Long _
) As Long

Public Declare Function GetDesktopWindow _
    Lib "user32" ( _
) As Long

Public Declare Function GetDeviceCaps _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nIndex As Long _
) As Long

Public Declare Function GetObjectA _
    Lib "gdi32" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
) As Long

Declare Function PtrArr _
    Lib "msvbvm60.dll" _
    Alias "VarPtr" ( _
    ptr() As Any _
) As Long

Public Declare Function GetObjectW _
    Lib "gdi32" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
) As Long

Public Declare Function GetStockObject _
    Lib "gdi32" ( _
    ByVal nIndex As Long _
) As Long

Public Declare Function GetPixel _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long _
) As Long

Public Declare Function GetTickCount _
    Lib "kernel32" ( _
) As Long

Public Declare Function GetVersionEx _
    Lib "kernel32" Alias "GetVersionExA" ( _
    lpVersionInformation As OSVERSIONINFO _
) As Long

Public Declare Function IntersectRect _
    Lib "user32" ( _
    lpDestRect As RECT_API, _
    lpSrc1Rect As RECT_API, _
    lpSrc2Rect As RECT_API _
) As Long

Public Declare Function LineTo _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long _
) As Long

Public Declare Function LoadImage _
    Lib "user32" Alias "LoadImageA" ( _
    ByVal hInst As Long, _
    ByVal FileName As String, _
    ByVal un1 As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal opmode As Long _
) As Long

Public Declare Function MoveTo _
    Lib "gdi32" Alias "MoveToEx" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    lpPoint As point _
) As Long

Public Declare Function Polyline _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    lpPoint As point, _
    ByVal nCount As Long _
) As Long

Public Declare Function SelectObject _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal hObject As Long _
) As Long

Public Declare Function SetBkColor _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal crColor As Long _
) As Long

Public Declare Function SetBkMode _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nBkMode As Long _
) As Long

Public Declare Function SetBitmapBits _
    Lib "gdi32" ( _
    ByVal hBitmap As Long, _
    ByVal dwCount As Long, _
    lpBits As Any _
) As Long

Public Declare Function SetPixel _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal crColor As Long _
) As Long

Public Declare Function SetTextColor _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal crColor As Long _
) As Long

Public Declare Function TextOutA _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal lpString As String, _
    ByVal nCount As Long _
) As Long

Public Declare Function TextOutW _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal lpString As String, _
    ByVal nCount As Long _
) As Long

Public Declare Function ValidateRect _
    Lib "user32" ( _
    ByVal hwnd As Long, _
    lpRect As RECT_API _
) As Long


Public Declare Function GetDIBits _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    lpBits As Any, _
    lpBI As BITMAPINFO, _
    ByVal wUsage As Long _
) As Long

Public Declare Function SetDIBits _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    lpBits As Any, _
    lpBI As BITMAPINFO, _
    ByVal wUsage As Long _
) As Long

Public Declare Function StretchDIBits _
    Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal DX As Long, _
    ByVal dy As Long, _
    ByVal SrcX As Long, _
    ByVal SrcY As Long, _
    ByVal wSrcWidth As Long, _
    ByVal wSrcHeight As Long, _
    lpBits As Any, _
    lpBitsInfo As BITMAPINFO, _
    ByVal wUsage As Long, _
    ByVal dwRop As Long _
) As Long

 
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
