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
