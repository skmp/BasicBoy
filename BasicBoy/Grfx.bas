Attribute VB_Name = "modGrfx"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)


'v1.2.0
'Graphics Engine
'Currenlty a bit optimized
'Color Gameboy Part is 95% finished
'Coments will be added with the next release
'New modes : DDraw DMA ,DDraw DMA Unsafe
'Fidex DrawLine and DrawLin4 for spirte/window priority
'Sory for my bad english ...

#Const bltDD_DMA = 0 'set to 1 is a bit faster but may be unstable
#Const Unsafe = 1 And bltDD_DMA 'set to 1 with bltdd_dma is faster but even more unstable
Global framedelay As Double
Dim bmh  As String
Option Base 0
Option Explicit
Dim tmpdesc As DDSURFACEDESC2
Public bmbH(53) As Byte
Public FPS As Long, fpsT As Long
Dim colid(128, 128) As Long
Dim colidx(128, 128) As Byte
Dim colid2(2, 128, 128) As Long
Dim re2 As dxvblib.RECT
Public FULLSCREEN As Boolean
Public TCol As Long, mode1 As Long
Dim i As Long, i2 As Long, X As Long, Y As Long, tilemap As Long, tileloc As Long, tileptr As Long
Dim xoffset As Long, yoffset As Long, TileData As Long, tileend As Long
Dim LByte As Long, HByte As Long, SpriteY As Long
Dim dy As Long, DX As Long, spat As Long
Dim tmp1 As Long, tmp2 As Long, memptr As Long, tmp3 As Long, rx As Long, ry As Long, lolm As Long, tms As Long
Global BB As BITMAPINFO, desthdc As Long, destW As Long, destH As Long, desthimg As Long
Global Vram(0 To 159, 0 To 143) As Integer, mir(255) As Byte, mv1 As Byte, mv2 As Byte, vf As Boolean, cid2 As Long, tiletmp As Long, xs As Long, ys As Long, xt As Long
Global tobj As Long, thdc As Long, DH As Long, Dw As Long, xr As Long, yr As Long
Global fskip As Long, fmode As Long, objp(7, 3) As Long, bgp(7, 3) As Long
Global vrm As Long, ccp As Long, ccid(128, 128) As Long, tm2 As Long, tm1 As Long, bgat As Byte ', gbcP(32767) As Integer
Global wv As Boolean, bgv As Boolean, objv As Boolean, xflip As Long, yflip As Long, lastline As Long
Global curline As Long, tcls(19) As Integer, curFreq As Currency, curStart As Long, dblResult As Double
Global Skipf As Boolean
Global bgpCC(7, 3) As Long, objpCC(7, 3) As Long
#If Unsafe Then
Dim sptr As Long
#End If

Public Sub DrawScreen() 'Using SetBits/StrechBits and directdraw
    FPS = FPS + 1
    
    If fmode = 0 Then 'frame skip mode 1(act skip(x1(1),x2(2),x3(3),x4(4),x5(5),x6(6))
    If FPS Mod fskip > 0 Then
    Skipf = True
    If lfp Then
    Do
    dblResult = GetTickCount2 - curStart
    'DoEvents
    If PeekMessage(message, 0&, 0&, 0&, PM_REMOVE) Then
        Call TranslateMessage(message)
        Call DispatchMessage(message)
    End If
    Loop While dblResult < framedelay
    End If
    Exit Sub
    End If
    Else 'frame skip mode 2(act skip(x1.20(6),x1.25(5),x1,3(4),x1.5(3))
    If FPS Mod fskip = 0 Then
    Skipf = True
    If lfp Then
    Do
    dblResult = GetTickCount2 - curStart
    'DoEvents
    If PeekMessage(message, 0&, 0&, 0&, PM_REMOVE) Then
        Call TranslateMessage(message)
        Call DispatchMessage(message)
    End If
    Loop While dblResult < framedelay
    End If
    Exit Sub
    End If
    End If
    
   If mode1 Then 'WinApi + DX
   #If bltDD_DMA = 0 Then
   desthdc = backbuffer.GetDC
   StretchDIBits desthdc, 0, 0, 160, 144, 0, 0, 160, 144, Vram(0, 0), BB, 0, vbSrcCopy
   backbuffer.ReleaseDC desthdc
   #Else
   #If Unsafe Then
   CopyMemory ByVal sptr, Vram(0, 0), 46080
   #Else
   backbuffer.Lock re2, tmpdesc, DDLOCK_DONOTWAIT, dHandle
   CopyMemory ByVal tmpdesc.lpSurface, Vram(0, 0), 46080
   backbuffer.Unlock re2
   #End If
   #End If
   dx7.GetWindowRect dHandle, surfaceRect
   If zm = 5 Then
   primary.Blt surfaceRect, backbuffer, re2, DDBLT_ASYNC
   Else
   primary.Blt surfaceRect, backbuffer, re2, DDBLT_ASYNC Or DDBLT_DONOTWAIT
   End If
   Else 'WinApi
   StretchDIBits desthdc, 0, 0, Dw, DH, 0, 0, 160, 144, Vram(0, 0), BB, 0, vbSrcCopy
   frmMain.Picture1.Refresh
   End If
   Skipf = False
   If lfp Then
   Do
   dblResult = GetTickCount2 - curStart
   'DoEvents
   If PeekMessage(message, 0&, 0&, 0&, PM_REMOVE) Then
        Call TranslateMessage(message)
        Call DispatchMessage(message)
   End If
   Loop While dblResult < framedelay
   End If
   End Sub
   
Public Sub initGxMode1(Siz As Long, Optional FULLSCREEN As Boolean = False)
Dim lWindowHandle As Long
Dim tm1 As Long, tm2 As Long, i As Long, i2 As Long
If Siz = 5 And FULLSCREEN = True Then Siz = 2
initGxMode2 frmMain.Picture1, Siz
mode1 = True
re2.Top = 0
re2.Left = 0
re2.Right = 159
re2.Bottom = 143
'setup the form
frmMain.Picture1.AutoRedraw = False
'init the gfx colors
backbuffer.BltColorFill re2, colid(0, 0)
frmMain.Picture1.Visible = (modVars.ro <> "")
primary.BltColorFill re2, RGB(255, 255, 255)
If FULLSCREEN And (modVars.ro <> "") Then
modDXEngine.fullSc
End If
#If Unsafe Then
   backbuffer.Lock re2, tmpdesc, DDLOCK_DONOTWAIT, dHandle
   sptr = tmpdesc.lpSurface
   backbuffer.Unlock re2
#End If
dx7.GetWindowRect dHandle, surfaceRect
primary.BltColorFill surfaceRect, 0
Exit Sub
sxe:
MsgBox "Error in drect draw init"
mode1 = 0 'fallback to WinApi
End Sub
Sub ccolid2(ByVal col As Byte, target As Long)
Dim tm1 As Long, tm2 As Long
    colid2(target, 0, 0) = colid(col And 2, col And 1)
    For tm2 = 1 To 128 Step 0
    colid2(target, 0, tm2) = colid(col And 8, col And 4)
    tm2 = tm2 * 2
    Next tm2
    For tm1 = 1 To 128 Step 0
    colid2(target, tm1, 0) = colid(col And 32, col And 16)
    tm1 = tm1 * 2
    Next tm1
    For tm1 = 1 To 128 Step 0
    For tm2 = 1 To 128 Step 0
    colid2(target, tm1, tm2) = colid(col And 128, col And 64)
    tm2 = tm2 * 2
    Next tm2
    tm1 = tm1 * 2
    Next tm1
End Sub
Sub initGxMode2(dest As PictureBox, Siz As Long)
mode1 = 0
frmMain.Picture1.AutoRedraw = True
frmMain.Picture1.BorderStyle = 0
frmMain.Picture1.ClipControls = False
frmMain.Picture1.ScaleMode = 3
frmMain.Picture1.BackColor = RGB(255, 255, 255)
frmMain.Picture1.Width = 15 * 160 * Siz
frmMain.Picture1.Height = 15 * 144 * Siz
frmMain.Picture1.Visible = True
With BB.Header
    .biSize = 40
    .biWidth = 160
    .biHeight = -144
    .biPlanes = 1
    .biBitCount = 16
    .biSizeImage = 46080
End With
destW = dest.ScaleWidth
destH = dest.ScaleHeight
desthdc = dest.hdc
desthimg = dest.Image.Handle
DH = 144 * Siz
Dw = 160 * Siz
frmMain.resize
StretchDIBits desthdc, 0, 0, 160, 144, 0, 0, 160, 144, Vram(0, 0), BB, 0, vbSrcCopy
End Sub
Sub initCol()
For i = 0 To 7
objpCC(i, 0) = rgb15(31, 31, 31)
bgpCC(i, 0) = rgb15(31, 31, 31)
objpCC(i, 1) = rgb15(21, 21, 21)
bgpCC(i, 1) = rgb15(21, 21, 21)
objpCC(i, 2) = rgb15(10, 10, 10)
bgpCC(i, 2) = rgb15(10, 10, 10)
objpCC(i, 3) = rgb15(0, 0, 0)
bgpCC(i, 3) = rgb15(0, 0, 0)
Next i
    initMir
    Dim tm2 As Long, tm1 As Long
    If TGBC Then
    colid(0, 0) = rgb15(31, 31, 31)
    For tm2 = 1 To 128
    colid(0, tm2) = rgb15(21, 21, 21)
    Next tm2
    
    For tm1 = 1 To 128
    colid(tm1, 0) = rgb15(10, 10, 10)
    Next tm1
    
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid(tm1, tm2) = rgb15(0, 0, 0)
    Next tm2
    Next tm1
    Else
    colid(0, 0) = rgb15(31, 31, 31)
    For tm2 = 1 To 128
    colid(0, tm2) = rgb15(21, 21, 21)
    Next tm2
    
    For tm1 = 1 To 128
    colid(tm1, 0) = rgb15(10, 10, 10)
    Next tm1
    
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid(tm1, tm2) = rgb15(0, 0, 0)
    Next tm2
    Next tm1
    End If
    colidx(0, 0) = 0
    
    For tm2 = 1 To 128
    colidx(0, tm2) = 2
    Next tm2
    
    For tm1 = 1 To 128
    colidx(tm1, 0) = 1
    Next tm1
    
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colidx(tm1, tm2) = 3
    Next tm2
    Next tm1
    
    colid2(0, 0, 0) = colid(0, 0)
    For tm2 = 1 To 128
    colid2(0, 0, tm2) = colid(0, 1)
    Next tm2
    For tm1 = 1 To 128
    colid2(0, tm1, 0) = colid(1, 0)
    Next tm1
    For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid2(0, tm1, tm2) = colid(1, 1)
    Next tm2
    Next tm1
    colid2(1, 0, 0) = colid(0, 0)
    For tm2 = 1 To 128
    colid2(1, 0, tm2) = colid(0, 1)
    Next tm2: For tm1 = 1 To 128
    colid2(1, tm1, 0) = colid(1, 0)
    Next tm1: For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid2(1, tm1, tm2) = colid(1, 1)
    Next tm2
    Next tm1
    colid2(2, 0, 0) = colid(0, 0)
    For tm2 = 1 To 128
    colid2(2, 0, tm2) = colid(0, 1)
    Next tm2
    For tm1 = 1 To 128
    colid2(2, tm1, 0) = colid(1, 0)
    Next tm1: For tm1 = 1 To 128
    For tm2 = 1 To 128
    colid2(2, tm1, tm2) = colid(1, 1)
    Next tm2: Next tm1
End Sub
Public Function initMir()
For i = 0 To 255
mir(i) = (i And 128) \ 128 + (i And 64) \ 32 + (i And 32) \ 8 + (i And 16) \ 2 + _
         (i And 8) * 2 + (i And 4) * 8 + (i And 2) * 32 + (i And 1) * 128
Next i
End Function

Function rgb15(Red As Byte, Green As Byte, Blue As Byte) As Integer
rgb15 = Red * 1024 + Green * 32 + Blue
End Function
Sub SrceenShot()
Dim flname As String, ni As Long
bmh = Chr(&H42) & Chr(&H4D) & Chr(&H36) & Chr(&HE) & Chr(&H1) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H36) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H28) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&HA0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H70) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&H1) & Chr(&H0) & Chr(&H10) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&HB4) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0)
ret:
flname = App.Path & "\" & Replace(rominfo.Title, "/", "") & " - " & ni & ".bmp"
If Not FileExist(flname) Then

If mode1 Then
Dim lng As Long
Open flname For Binary As #1
Put #1, , bmh
Put #1, , Vram
Close #1

Else
SavePicture frmMain.Picture1.Image, flname
End If
Else
ni = ni + 1
GoTo ret
End If
End Sub

Function FileExist(file As String) As Boolean
On Error Resume Next
If FileLen(file) = 0 Then FileExist = False Else FileExist = True
End Function
Sub Drawline4() 'Using Api and SetBits/StrechBits (Old Gameboy)
curline = RAM(65348, 0)
If curline = lastline Then Exit Sub
lastline = curline
    ' Draw Background
    ' Get BG & window Tile Pattern Data Address
    If RAM(65344, 0) And 16 Then
        TileData = 32768
    Else
        TileData = 34816
    End If
    If bgv Then
    ' Get BG Tile Table Address
    If RAM(65344, 0) And 8 Then
        tilemap = 39936
        tms = 39936
    Else
        tilemap = 38912
        tms = 38912
    End If
    tileend = tilemap + 1023
    xoffset = RAM(65347, 0)
    yoffset = RAM(65346, 0) + curline
    xs = xoffset \ 8
    ys = yoffset \ 8
    yoffset = yoffset And 7
    xoffset = -(xoffset And 7)
    For X = xoffset To 159 Step 8
        tiletmp = tilemap + ys * 32 + xs
        If tiletmp > tileend Then tiletmp = tiletmp - 1024
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tiletmp, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tiletmp, 0) Xor 128) * 16
        End If
        
        xs = (xs + 1) Mod 32
        
        memptr = TileData + tileptr + (yoffset And 7) * 2
        mv1 = RAM(memptr + 1, 0): mv2 = RAM(memptr, 0)
        If X > -1 And X < 153 Then
            Vram(X + 7, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 6, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 5, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 4, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 3, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 2, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 1, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X, curline) = colid2(0, mv1 And 1, mv2 And 1)
        Else
            If X < 153 Then Vram(X + 7, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -7 And X < 154 Then Vram(X + 6, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -6 And X < 155 Then Vram(X + 5, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -5 And X < 156 Then Vram(X + 4, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -4 And X < 157 Then Vram(X + 3, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -3 And X < 158 Then Vram(X + 2, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -2 And X < 159 Then Vram(X + 1, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -1 And X < 160 Then Vram(X, curline) = colid2(0, mv1 And 1, mv2 And 1)
        End If
    Next X
    End If
    
    
        'Draw Sprites
If (RAM(65344, 0) And 2) And objv Then
    If (RAM(65344, 0) And 4) = 0 Then
        SpriteY = 7
    Else
        SpriteY = 15
    End If
    For tilemap = 65180 To 65024 Step -4
        Y = RAM(tilemap, 0) - 16
        X = RAM(tilemap + 1, 0) - 8
        If Y <= curline And Y + SpriteY >= curline And X > -8 And X < 160 Then
        
        If SpriteY = 7 Then tileptr = readM(tilemap + 2) * 16 Else tileptr = (RAM(tilemap + 2, 0) And 254) * 16 'Get pointer to tile
        TCol = -((RAM(tilemap + 3, 0) And 128) > 0)
        spat = -((readM(tilemap + 3) And 16) > 0) + 1
        vf = RAM(tilemap + 3, 0) And 32
        'init palete
        memptr = 32768 + tileptr
        
        If (RAM(tilemap + 3, 0) And 64) Then memptr = memptr + SpriteY * 2 - (curline - Y) * 2 Else memptr = memptr + (curline - Y) * 2
        If vf Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        
        If TCol = 0 Then
            
            If X < 153 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 7, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -7 And X < 154 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 6, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -6 And X < 155 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 5, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -5 And X < 156 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 4, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -4 And X < 157 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 3, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -3 And X < 158 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 2, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -2 And X < 159 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 1, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -1 And X < 160 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            
        Else
            cid2 = colid2(0, 0, 0)
            If X < 153 Then If Vram(X + 7, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 7, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -7 And X < 154 Then If Vram(X + 6, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 6, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -6 And X < 155 Then If Vram(X + 5, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 5, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -5 And X < 156 Then If Vram(X + 4, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 4, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -4 And X < 157 Then If Vram(X + 3, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 3, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -3 And X < 158 Then If Vram(X + 2, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 2, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -2 And X < 159 Then If Vram(X + 1, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 1, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -1 And X < 160 Then If Vram(X, curline) = cid2 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X, curline) = colid2(spat, mv1 And 1, mv2 And 1)
            
        End If
    End If
    Next tilemap
   End If
   
        'Draw Window
    
If ((RAM(65344, 0) And 32) = 32) And wv And (curline >= RAM(65354, 0)) And (RAM(65355, 0) < 167) Then
    ' Get window Tile Table Address
    If RAM(65344, 0) And 64 Then
        tilemap = 39936
    Else
        tilemap = 38912
    End If
    yoffset = curline - RAM(65354, 0)
    tilemap = tilemap + (yoffset \ 8) * 32
    yoffset = yoffset And 7
    For X = RAM(65355, 0) - 7 To 159 Step 8
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tilemap, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tilemap, 0) Xor 128) * 16
        End If
        memptr = TileData + tileptr + (yoffset And 7) * 2
        mv1 = RAM(memptr + 1, 0): mv2 = RAM(memptr, 0)
        
        If X > -1 And X < 153 Then
            Vram(X + 7, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 6, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 5, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 4, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 3, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 2, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X + 1, curline) = colid2(0, mv1 And 1, mv2 And 1): mv1 = mv1 \ 2: mv2 = mv2 \ 2
            Vram(X, curline) = colid2(0, mv1 And 1, mv2 And 1)
            tilemap = tilemap + 1
        Else
            If X < 153 Then Vram(X + 7, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -7 And X < 154 Then Vram(X + 6, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -6 And X < 155 Then Vram(X + 5, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -5 And X < 156 Then Vram(X + 4, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -4 And X < 157 Then Vram(X + 3, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -3 And X < 158 Then Vram(X + 2, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -2 And X < 159 Then Vram(X + 1, curline) = colid2(0, mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -1 And X < 160 Then Vram(X, curline) = colid2(0, mv1 And 1, mv2 And 1)
            tilemap = tilemap + 1
        End If
        
    Next X
End If
    
End Sub
Sub Drawline() 'Using Api and SetBits/StrechBits
curline = RAM(65348, 0)
If curline = lastline Then Exit Sub
lastline = curline
    ' Draw Background
    ' Get BG & window Tile Pattern Data Address
    If RAM(65344, 0) And 16 Then
        TileData = 32768
    Else
        TileData = 34816
    End If
If bgv Then
    ' Get BG Tile Table Address
    If RAM(65344, 0) And 8 Then
        tilemap = 39936
        tms = 39936
    Else
        tilemap = 38912
        tms = 38912
    End If
    tileend = tilemap + 1023
    xoffset = RAM(65347, 0)
    yoffset = RAM(65346, 0) + curline
    xs = xoffset \ 8
    ys = yoffset \ 8
    yoffset = yoffset And 7
    xoffset = -(xoffset And 7)
    
    For X = xoffset To 159 Step 8
        tiletmp = tilemap + ys * 32 + xs
        If tiletmp > tileend Then tiletmp = tiletmp - 1024
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tiletmp, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tiletmp, 0) Xor 128) * 16
        End If
        bgat = RAM(tiletmp, 1)
        ccp = bgat And 7
        tcls(X \ 8) = bgpCC(ccp, 0)
        vrm = (bgat And 8) \ 8
        xflip = (bgat And 32) \ 32: yflip = (bgat And 64) \ 64
        ccid(0, 0) = bgpCC(ccp, 0)
        ccid(0, 1) = bgpCC(ccp, 1)
        ccid(1, 0) = bgpCC(ccp, 2)
        ccid(1, 1) = bgpCC(ccp, 3)
        If yflip Then memptr = TileData + tileptr + 14 - (yoffset And 7) * 2 Else memptr = TileData + tileptr + (yoffset And 7) * 2
        If xflip Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        xs = (xs + 1) Mod 32
        
        If X > -1 And X < 153 Then
        Vram(X + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X, curline) = ccid(mv1 And 1, mv2 And 1)
        Else
        If X < 153 Then Vram(X + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -7 And X < 154 Then Vram(X + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -6 And X < 155 Then Vram(X + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -5 And X < 156 Then Vram(X + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -4 And X < 157 Then Vram(X + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -3 And X < 158 Then Vram(X + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -2 And X < 159 Then Vram(X + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -1 And X < 160 Then Vram(X, curline) = ccid(mv1 And 1, mv2 And 1)
        End If
        
    Next X
    End If
    
    'Draw Sprites
If (RAM(65344, 0) And 2) And objv Then
    If (RAM(65344, 0) And 4) = 0 Then
        SpriteY = 7
    Else
        SpriteY = 15
    End If
    For tilemap = 65180 To 65024 Step -4
        Y = RAM(tilemap, 0) - 16
        X = RAM(tilemap + 1, 0) - 8
        If Y <= curline And Y + SpriteY >= curline And X > -8 And X < 160 Then
        
        If SpriteY = 7 Then tileptr = readM(tilemap + 2) * 16 Else tileptr = (RAM(tilemap + 2, 0) And 254) * 16 'Get pointer to tile
        TCol = -((RAM(tilemap + 3, 0) And 128) > 0)
        vf = RAM(tilemap + 3, 0) And 32
        vrm = (RAM(tilemap + 3, 0) And 8) \ 8
        ccp = RAM(tilemap + 3, 0) And 7
        'init palete
        ccid(0, 0) = objpCC(ccp, 0)
        ccid(0, 1) = objpCC(ccp, 1)
        ccid(1, 0) = objpCC(ccp, 2)
        ccid(1, 1) = objpCC(ccp, 3)
        memptr = 32768 + tileptr
        
        If (RAM(tilemap + 3, 0) And 64) Then memptr = memptr + SpriteY * 2 - (curline - Y) * 2 Else memptr = memptr + (curline - Y) * 2
        If vf Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        
        If TCol = 0 Then
            If X < 153 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 7, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -7 And X < 154 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 6, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -6 And X < 155 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 5, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -5 And X < 156 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 4, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -4 And X < 157 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 3, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -3 And X < 158 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 2, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -2 And X < 159 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X + 1, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -1 Then If colidx(mv1 And 1, mv2 And 1) Then Vram(X, curline) = ccid(mv1 And 1, mv2 And 1)
            
        Else
            If X < 153 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X + 7, curline) = tcls((X + 7) \ 8) Then Vram(X + 7, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -7 And X < 154 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X + 6, curline) = tcls((X + 6) \ 8) Then Vram(X + 6, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -6 And X < 155 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X + 5, curline) = tcls((X + 5) \ 8) Then Vram(X + 5, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -5 And X < 156 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X + 4, curline) = tcls((X + 4) \ 8) Then Vram(X + 4, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -4 And X < 157 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X + 3, curline) = tcls((X + 3) \ 8) Then Vram(X + 3, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -3 And X < 158 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X + 2, curline) = tcls((X + 2) \ 8) Then Vram(X + 2, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -2 And X < 159 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X + 1, curline) = tcls((X + 1) \ 8) Then Vram(X + 1, curline) = ccid(mv1 And 1, mv2 And 1)
            mv1 = mv1 \ 2: mv2 = mv2 \ 2
            If X > -1 Then If colidx(mv1 And 1, mv2 And 1) Then If Vram(X, curline) = tcls(X \ 8) Then Vram(X, curline) = ccid(mv1 And 1, mv2 And 1)
            
        End If
    End If
    Next tilemap
   End If
   
           'Draw Window
    
If ((RAM(65344, 0) And 32) = 32) And wv And (curline >= RAM(65354, 0)) And (RAM(65355, 0) < 167) Then
    ' Get window Tile Table Address
    If RAM(65344, 0) And 64 Then
        tilemap = 39936
    Else
        tilemap = 38912
    End If
    yoffset = curline - RAM(65354, 0)
    tilemap = tilemap + (yoffset \ 8) * 32
    yoffset = yoffset And 7
    For X = RAM(65355, 0) - 7 To 159 Step 8
        If TileData = 32768 Then           ' Tile Data @ &H8800-&h97FF is 128ed
            tileptr = RAM(tilemap, 0) * 16             'Get pointer to tile
        Else
            tileptr = (RAM(tilemap, 0) Xor 128) * 16
        End If
        bgat = RAM(tilemap, 1)
        ccp = bgat And 7
        vrm = (bgat And 8) \ 8
        xflip = (bgat And 32) \ 32: yflip = (bgat And 64) \ 64
        ccid(0, 0) = bgpCC(ccp, 0)
        ccid(0, 1) = bgpCC(ccp, 1)
        ccid(1, 0) = bgpCC(ccp, 2)
        ccid(1, 1) = bgpCC(ccp, 3)
        If yflip Then memptr = TileData + tileptr + 14 - yoffset * 2 Else memptr = TileData + tileptr + yoffset * 2
        If xflip Then mv1 = mir(RAM(memptr + 1, vrm)): mv2 = mir(RAM(memptr, vrm)) Else mv1 = RAM(memptr + 1, vrm): mv2 = RAM(memptr, vrm)
        
        If X > -1 And X < 153 Then
        Vram(X + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        Vram(X, curline) = ccid(mv1 And 1, mv2 And 1)
        Else
        If X < 153 Then Vram(X + 7, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -7 And X < 154 Then Vram(X + 6, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -6 And X < 155 Then Vram(X + 5, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -5 And X < 156 Then Vram(X + 4, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -4 And X < 157 Then Vram(X + 3, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -3 And X < 158 Then Vram(X + 2, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -2 And X < 159 Then Vram(X + 1, curline) = ccid(mv1 And 1, mv2 And 1)
        mv1 = mv1 \ 2: mv2 = mv2 \ 2
        If X > -1 And X < 160 Then Vram(X, curline) = ccid(mv1 And 1, mv2 And 1)
        End If
        
        tilemap = tilemap + 1
    Next X
End If
End Sub
