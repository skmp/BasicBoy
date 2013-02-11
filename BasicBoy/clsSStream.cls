VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'implements a Streamed sound buffer with dynamicaly generated data
Option Explicit
Option Base 0
Dim bufflen As Long
Dim ch As DirectSoundSecondaryBuffer8
Dim bd As DSBUFFERDESC
Public ev0 As Long, ev1 As Long
Dim i As Long, si As Double, step As Double
Dim wave() As Byte, buff() As Byte, wlen As Long
Public freq As Long, R As Long, L As Long
Public Volume As Double, Enabled As Long, Play As Long
Dim pn(1) As DSBPOSITIONNOTIFY
Implements DirectXEvent8
Sub init(sbl As Double)
bd.fxFormat.nFormatTag = WAVE_FORMAT_PCM
bd.fxFormat.nChannels = 2
bd.fxFormat.lSamplesPerSec = 44000
bd.fxFormat.nBitsPerSample = 8
bd.fxFormat.nBlockAlign = 1
bd.fxFormat.lAvgBytesPerSec = bd.fxFormat.lSamplesPerSec * bd.fxFormat.nBlockAlign * bd.fxFormat.nChannels
bd.lFlags = DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_CTRLPOSITIONNOTIFY
bd.lBufferBytes = bd.fxFormat.lAvgBytesPerSec * sbl
bufflen = bd.lBufferBytes
Set ch = ds.CreateSoundBuffer(bd)
pn(0).hEventNotify = ev0
pn(0).lOffset = bd.lBufferBytes / 2
pn(1).hEventNotify = ev1
pn(1).lOffset = 0
ch.SetNotificationPositions 2, pn
ReDim buff(bufflen)
End Sub
Sub setWave(waveform() As Byte)
wave = waveform
wlen = UBound(wave)
step = (wlen * freq) / 44000
End Sub
Sub createWave(siz As Long)
For i = 0 To (siz / 2) - 1
If L And Enabled And Play Then buff(i * 2) = wave(si) Else buff(i * 2) = 128
If R And Enabled And Play Then buff(i * 2 + 1) = wave(si) Else buff(i * 2) = 128
si = si + step
If si > wlen Then si = si - wlen
Next i
End Sub
Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
Select Case eventid
Case ev0 'write from 0 to midle
    createWave bufflen / 2
    ch.WriteBuffer 0, bufflen / 2, buff(0), DSBLOCK_DEFAULT
Case ev1 'write from midle to end
    createWave bufflen / 2
    ch.WriteBuffer bufflen / 2, bufflen / 2, buff(0), DSBLOCK_DEFAULT
End Select
End Sub

