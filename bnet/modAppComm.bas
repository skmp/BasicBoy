Attribute VB_Name = "modAppComm"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)



'v1.1.1
'App comm function (based on Black Tornado's Trainer Maker Kit)
'I'm using direct memory writes
'Comments added
'hmm,This can be done with subclassing too

'Sory for my bad english ...
Option Explicit
Public BBhWnd As Long, MemAddr As Long, wlen As Integer
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessID As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Sub init(InitData As String, Length As Integer, ptr As Long) 'Init the MemIo system
BBhWnd = InitData
MemAddr = ptr
wlen = Length
End Sub

Public Function Send(value() As Byte) As Boolean 'Send a value
Dim ProcessID As Long
Dim ProcessHandle As Long
If BBhWnd = False Then Send = False: Exit Function
GetWindowThreadProcessId BBhWnd, ProcessID
ProcessHandle = OpenProcess(2035711, False, ProcessID)
Call WriteProcessMemory(ProcessHandle, MemAddr, value(0), wlen, 0&)
CloseHandle ProcessHandle
Send = True
End Function
