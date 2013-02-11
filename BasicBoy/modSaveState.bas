Attribute VB_Name = "modSaveState"
Option Explicit

'## This is a part of the BasicBoy emulator
'## You are not allowed to release modified(or unmodified) versions
'## without asking me (Raziel).
'## For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'## To download the latestt version/source goto basicboy.emuhost.com
'##
'## Save state module by Xeon. (http://www.xeons.net)
'## Version 1.0

Public Sub saveState(Index As Integer)
    On Error GoTo BadSave:
    If rominfo.Title = vbNullString Then
        MsgBox "No rom loaded!", vbCritical, "Invalid Slot"
        Exit Sub
    End If
    Dim iFreeFile As Integer
    iFreeFile = FreeFile
    Open App.Path & "\gbsavestate.dat" For Binary As #iFreeFile
        Put #iFreeFile, , hline
        Put #iFreeFile, , IME
        Put #iFreeFile, , GBM
        Put #iFreeFile, , f_lowbits
        Put #iFreeFile, , zf
        Put #iFreeFile, , nf
        Put #iFreeFile, , hf
        Put #iFreeFile, , cf
        Put #iFreeFile, , b
        Put #iFreeFile, , c
        Put #iFreeFile, , D
        Put #iFreeFile, , E
        Put #iFreeFile, , H
        Put #iFreeFile, , L
        Put #iFreeFile, , PC
        Put #iFreeFile, , SP
        Put #iFreeFile, , cldr
        Put #iFreeFile, , Clm0
        Put #iFreeFile, , clm3
        Put #iFreeFile, , cllc
        Put #iFreeFile, , CpuS
        Put #iFreeFile, , RAM
        Put #iFreeFile, , bRam
        Put #iFreeFile, , Vram
        Put #iFreeFile, , mir
        Put #iFreeFile, , mv1
        Put #iFreeFile, , mv2
        Put #iFreeFile, , vf
        Put #iFreeFile, , cid2
        Put #iFreeFile, , tiletmp
        Put #iFreeFile, , xs
        Put #iFreeFile, , ys
        Put #iFreeFile, , xt
        Put #iFreeFile, , tobj
        Put #iFreeFile, , thdc
        Put #iFreeFile, , DH
        Put #iFreeFile, , Dw
        Put #iFreeFile, , xr
        Put #iFreeFile, , yr
        Put #iFreeFile, , fskip
        Put #iFreeFile, , fmode
        Put #iFreeFile, , objp
        Put #iFreeFile, , bgp
        Put #iFreeFile, , vrm
        Put #iFreeFile, , ccp
        Put #iFreeFile, , ccid
        Put #iFreeFile, , tm2
        Put #iFreeFile, , tm1
        Put #iFreeFile, , bgat
        Put #iFreeFile, , wv
        Put #iFreeFile, , bgv
        Put #iFreeFile, , objv
        Put #iFreeFile, , xflip
        Put #iFreeFile, , yflip
        Put #iFreeFile, , lastline
        Put #iFreeFile, , curline
        Put #iFreeFile, , tcls
        Put #iFreeFile, , curFreq
        Put #iFreeFile, , curStart
        'Put #iFreeFile, , CurEnd
        Put #iFreeFile, , dblResult
        Put #iFreeFile, , Skipf
        Put #iFreeFile, , bgpCC
        Put #iFreeFile, , objpCC
        Put #iFreeFile, , CurRAMBank
        Put #iFreeFile, , CurROMBank
    Close #iFreeFile
    'Excellent Fast Compression System from BasicNES ;-)
    RLECompress App.Path & "\gbsavestate.dat", App.Path & "\" & LCase(rominfo.Title) & ".st" & CStr(Index)
    'Delete temp
    Kill App.Path & "\gbsavestate.dat"
    Exit Sub
BadSave:
    MsgBox "Save State Error" & vbCrLf & _
    "--------------------------------------------" & vbCrLf & _
    err.Description & vbCrLf & _
    "--------------------------------------------", vbCritical, "Save State Error"
End Sub

Public Sub loadState(Index As Integer)
    On Error GoTo BadLoad:
    If rominfo.Title = vbNullString Then
        MsgBox "No rom loaded!", vbCritical, "Invalid Slot"
        Exit Sub
    End If
    If FileExist(App.Path & "\" & LCase(rominfo.Title) & ".st" & CStr(Index)) = False Then
        MsgBox "Invalid Load Slot", vbCritical, "Invalid Slot"
        Exit Sub
    End If
    Dim iFreeFile As Integer
    iFreeFile = FreeFile
    RLEDecompress App.Path & "\" & LCase(rominfo.Title) & ".st" & CStr(Index), App.Path & "\gbsavestate.dat"
    Open App.Path & "\gbsavestate.dat" For Binary As #iFreeFile
        Get #iFreeFile, , hline
        Get #iFreeFile, , IME
        Get #iFreeFile, , GBM
        Get #iFreeFile, , f_lowbits
        Get #iFreeFile, , zf
        Get #iFreeFile, , nf
        Get #iFreeFile, , hf
        Get #iFreeFile, , cf
        Get #iFreeFile, , b
        Get #iFreeFile, , c
        Get #iFreeFile, , D
        Get #iFreeFile, , E
        Get #iFreeFile, , H
        Get #iFreeFile, , L
        Get #iFreeFile, , PC
        Get #iFreeFile, , SP
        Get #iFreeFile, , cldr
        Get #iFreeFile, , Clm0
        Get #iFreeFile, , clm3
        Get #iFreeFile, , cllc
        Get #iFreeFile, , CpuS
        Get #iFreeFile, , RAM
        Get #iFreeFile, , bRam
        Get #iFreeFile, , Vram '??
        Get #iFreeFile, , mir
        Get #iFreeFile, , mv1
        Get #iFreeFile, , mv2
        Get #iFreeFile, , vf
        Get #iFreeFile, , cid2
        Get #iFreeFile, , tiletmp
        Get #iFreeFile, , xs
        Get #iFreeFile, , ys
        Get #iFreeFile, , xt
        Get #iFreeFile, , tobj
        Get #iFreeFile, , thdc
        Get #iFreeFile, , DH
        Get #iFreeFile, , Dw
        Get #iFreeFile, , xr
        Get #iFreeFile, , yr
        Get #iFreeFile, , fskip
        Get #iFreeFile, , fmode
        Get #iFreeFile, , objp
        Get #iFreeFile, , bgp
        Get #iFreeFile, , vrm
        Get #iFreeFile, , ccp
        Get #iFreeFile, , ccid
        Get #iFreeFile, , tm2
        Get #iFreeFile, , tm1
        Get #iFreeFile, , bgat
        Get #iFreeFile, , wv
        Get #iFreeFile, , bgv
        Get #iFreeFile, , objv
        Get #iFreeFile, , xflip
        Get #iFreeFile, , yflip
        Get #iFreeFile, , lastline
        Get #iFreeFile, , curline
        Get #iFreeFile, , tcls
        Get #iFreeFile, , curFreq
        Get #iFreeFile, , curStart
        'Get #iFreeFile, , CurEnd
        Get #iFreeFile, , dblResult
        Get #iFreeFile, , Skipf
        Get #iFreeFile, , bgpCC
        Get #iFreeFile, , objpCC
        Get #iFreeFile, , CurRAMBank
        Get #iFreeFile, , CurROMBank
    Close #iFreeFile
    Kill App.Path & "\gbsavestate.dat"
    Exit Sub
BadLoad:
    MsgBox "Load State Error" & vbCrLf & _
    "--------------------------------------------" & vbCrLf & _
    err.Description & vbCrLf & _
    "--------------------------------------------", vbCritical, "Load State Error"
End Sub

Public Function FileExist(strFileName As String) As Boolean
    On Error Resume Next
    If strFileName = "" Then
        FileExist = False
        Exit Function
    End If
    FileExist = (Dir(strFileName) <> "")
End Function
