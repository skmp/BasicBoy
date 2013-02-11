Attribute VB_Name = "modComLine"
'Command Line boot (exename romname)
'Added by Christopher
Option Explicit
Sub Main()
Dim strtemp As String, bol As Boolean, tls As String, i As Long
Dim f As String
framedelay = 16
'Form1.Show
frmMain.Show
frmSplash.Show

On Local Error GoTo ErrorHandler
If InStr(command$, ".g") Or InStr(command$, ".c") Then

f = Replace(command$, """", "")

If loadrom(f) Then
initCI
rdRam
If mm = 2 Then
initGxMode2 frmMain.Picture1, zm
Else
initGxMode1 zm, frmMain.full.Checked
End If
If TGBC Then
If ROM(&H143, 0) = 192 Then strtemp = "(GBC) ": GBM = 1 Else If ROM(&H143, 0) <> 0 Then strtemp = "(GB/GBC) ": GBM = 1 Else strtemp = "(GB) ": GBM = 0
Else
If ROM(&H143, 0) = 192 Then strtemp = "(GBC) " Else If ROM(&H143, 0) <> 0 Then strtemp = "(GB/GBC) " Else strtemp = "(GB) "
GBM = 0
End If
frmMain.resize
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
frmMain.rp_Click
End If
End If





err.Clear
ErrorHandler:
If err.Number <> 0 Then
MsgBox "An Error Occurred trying to load your ROM:" & vbCrLf & vbCrLf & err.Description, vbCritical, "Error " & err.Number
End If
End Sub




