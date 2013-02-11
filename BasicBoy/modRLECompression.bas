Attribute VB_Name = "modRLECompression"
'***Added by Xeon (www.xeons.net)***
'Just has compression code for the saved states, and a simple routine to count the length of a file in lines.

'mostly contains simple rle compression routines, but they work great on the save states

'Buffered IO. faster.
Private readBuffer(4095) As Byte
Private writebuffer(4095) As Byte
Private readptr As Long
Private writeptr As Long
Private readsize As Long

Private runlength As Byte
Private runchar As Byte
Private nextchar As Byte
Private temp As Byte
Public gif As IPictureDisp
Private ind As Long

Public palName As String
Private Function readchar() As Byte
    If readptr = 0 Then
        If readsize >= 4096 Then
            Get #1, , readBuffer
        Else
            Dim b() As Byte, i As Long
            ReDim b(readsize - 1)
            Get #1, , b
            For i = 0 To readsize - 1
                readBuffer(i) = b(i)
            Next i
        End If
    End If
    readchar = readBuffer(readptr)
    readptr = (readptr + 1) And 4095
    readsize = readsize - 1
End Function

Private Sub writechar(c As Byte)
    writebuffer(writeptr) = c
    writeptr = (writeptr + 1) And 4095
    If writeptr = 0 Then Put #2, , writebuffer
End Sub

Private Sub preclose()
    If writeptr > 0 Then
        Dim b() As Byte, i As Long
        ReDim b(writeptr - 1)
        For i = 0 To writeptr - 1: b(i) = writebuffer(i): Next i
        Put #2, , b
    End If
End Sub

Public Sub delete(f As String)
    On Error Resume Next
    Kill f
End Sub

Private Sub scanrun()
runchar = nextchar
runlength = 0
Do
    runlength = runlength + 1
    nextchar = readchar
Loop Until nextchar <> runchar Or runlength = 255 Or readsize = 0
End Sub

Private Sub writerun()
Dim i As Long
For i = 1 To runlength
writechar runchar
Next i
End Sub

Private Sub encoderun()
Dim i As Long
If runlength > 3 Then
    temp = 207
    writechar temp
    writechar runlength
    writechar runchar
Else
    For i = 1 To runlength
        writechar runchar
        If runchar = 207 Then
            temp = 0
            writechar temp
        End If
    Next i
End If
End Sub

Private Sub decoderun()
runchar = readchar
If runchar = 207 Then
    runlength = readchar
    If runlength > 0 Then
        runchar = readchar
    Else
        runlength = 1
    End If
Else
    runlength = 1
End If
End Sub

'very simple RLE file compression
Public Sub RLECompress(infile As String, outfile As String)
    delete outfile
    Open infile For Binary As #1
    Open outfile For Binary As #2
    readsize = LOF(1)
    readptr = 0
    writeptr = 0
    
    Get #1, , nextchar
    While readsize > 0
        scanrun
        If readsize = 0 Then
            If nextchar = runchar And runlength < 255 Then
                runlength = runlength + 1
                encoderun
            Else
                encoderun
                runlength = 1
                runchar = nextchar
                encoderun
            End If
        Else
            encoderun
        End If
    Wend
    preclose
    Close
End Sub

Public Sub RLEDecompress(infile As String, outfile As String)
    delete outfile
    Open infile For Binary As #1
    Open outfile For Binary As #2
    readsize = LOF(1)
    readptr = 0
    writeptr = 0
    
    While readsize > 0
        decoderun
        writerun
    Wend
    preclose
    Close
End Sub


