VERSION 5.00
Begin VB.Form clsSStream 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "clsSStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'implements a Streamed sound buffer with dynamicaly generated data
Option Explicit
Option Base 0
Public bufflen As Long, halfbuf As Long
Public ev0 As Long, ev1 As Long
Public ch As DirectSoundSecondaryBuffer8
Dim bd As DSBUFFERDESC
Dim i As Long
Dim pn(1) As DSBPOSITIONNOTIFY
Implements DirectXEvent8
Sub init(buflen As Long)
bufflen = buflen
Set dx8 = New DirectX8
ev0 = dx8.CreateEvent(Me)
ev1 = dx8.CreateEvent(Me)
Set dsound = dx8.DirectSoundCreate("")
dsound.SetCooperativeLevel frmMain.hwnd, DSSCL_NORMAL

bd.fxFormat.nFormatTag = WAVE_FORMAT_PCM
bd.fxFormat.nChannels = 1
bd.fxFormat.lSamplesPerSec = 44100
bd.fxFormat.nBitsPerSample = 8
bd.fxFormat.nBlockAlign = 1
bd.fxFormat.lAvgBytesPerSec = bd.fxFormat.lSamplesPerSec * bd.fxFormat.nBlockAlign
bd.lFlags = DSBCAPS_GETCURRENTPOSITION2 Or DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_STATIC Or (DSBCAPS_LOCSOFTWARE * ssound)
'**********************If sound is bad try increasing the buffer here***********************
bd.lBufferBytes = bufflen ' x ms buffer , x\2 ms delay
Set ch = dsound.CreateSoundBuffer(bd)
pn(0).hEventNotify = ev0
pn(0).lOffset = bd.lBufferBytes / 2 + 1
pn(1).hEventNotify = ev1
pn(1).lOffset = 1
ch.SetNotificationPositions 2, pn
End Sub
Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
Select Case eventid
Case ev0 'play >half write 0 (0-half-1)
generate 0
Case ev1 'Play <half write 1 (half-end)
generate 1
End Select
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub

Private Sub Form_Load()
Call InitCommonControls
Me.Hide
End Sub
