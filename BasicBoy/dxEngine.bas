Attribute VB_Name = "modDXEngine"
'This is a part of the BasicBoy emulator
'You are not allowed to release modified(or unmodified) versions
'without asking me (Raziel).
'For Suggestions ect please e-mail at :stef_mp@yahoo.gr
'To download the latest version/source goto basicboy.emuhost.com
'(I know the emulator is NOT OPTIMIZED AT ALL)

'This is a complete rewrite of the DX interface
'I liked more the vb.net way so i decided to
'rewrite the dx interface in VB6 too :)
'Rewrite of the sound interface
'v 1.3.0
Option Explicit
'DirectX Stuff
'DX7 vars
Public dx7 As DirectX7
Public DDraw As DirectDraw7
Public clipper As DirectDrawClipper
Public primary   As DirectDrawSurface7
Public backbuffer As DirectDrawSurface7
Public surfaceRect As dxvblib.RECT
'DX8 vars
Public dx8 As DirectX8
Public dsound As DirectSound8
'Everything else
Public dHandle As Long, wHandle As Long
Sub InitDirectX(ByVal whwnd As Long, ByVal dhwnd As Long)
dHandle = dhwnd
wHandle = whwnd
Set dx7 = New DirectX7
Set dx8 = New DirectX8
initDirectDraw
End Sub
Sub initDirectDraw()
Dim ddsd As DDSURFACEDESC2
On Error GoTo create_error
    Set DDraw = dx7.DirectDrawCreate("")
    DDraw.SetCooperativeLevel dHandle, DDSCL_NORMAL
    'create the primary display surface
    ddsd.lFlags = DDSD_CAPS 'Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    'create the  surface
    Set primary = DDraw.CreateSurface(ddsd)
    'Craete a clipper
    Set clipper = DDraw.CreateClipper(0)
    'assoiciate the window handle with the clipper
    clipper.SetHWnd dHandle
    'clip blitting routines to the window
    primary.SetClipper clipper
    'create a normal surface
    ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'set surface resolution
    ddsd.lWidth = 160
    ddsd.lHeight = 144
    ddsd.ddpfPixelFormat.lFlags = DDPF_RGB
    ddsd.ddpfPixelFormat.lRGBBitCount = 16
    ddsd.ddpfPixelFormat.lRBitMask = 31744
    ddsd.ddpfPixelFormat.lGBitMask = 992
    ddsd.ddpfPixelFormat.lBBitMask = 31
    Set backbuffer = DDraw.CreateSurface(ddsd)
    'surface rectangle
    surfaceRect.Bottom = ddsd.lHeight
    surfaceRect.Right = ddsd.lWidth
    Exit Sub
create_error:
    MsgBox "Direct Draw Error : " & err.Description & " <No:" & err.Number & ">"
    mode1 = 0
End Sub
Sub fullSc()
Dim ddsd As DDSURFACEDESC2
On Error GoTo create_error
    dHandle = frmRender.hwnd
    Set primary = Nothing
    Set backbuffer = Nothing
    Set DDraw = Nothing
    Set DDraw = dx7.DirectDrawCreate("")
    DDraw.SetCooperativeLevel dHandle, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    DDraw.SetDisplayMode 800, 600, 32, 75, DDSDM_DEFAULT
    'create the primary display surface
    ddsd.lFlags = DDSD_CAPS 'Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    'create the  surface
    Set primary = DDraw.CreateSurface(ddsd)
    'Craete a clipper
    Set clipper = DDraw.CreateClipper(0)
    'assoiciate the window handle with the clipper
    clipper.SetHWnd dHandle
    'clip blitting routines to the window
    primary.SetClipper clipper
    'create a normal surface
    ddsd.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    'set surface resolution
    ddsd.lWidth = 160
    ddsd.lHeight = 144
    Set backbuffer = DDraw.CreateSurface(ddsd)
    'surface rectangle
    surfaceRect.Bottom = ddsd.lHeight
    surfaceRect.Right = ddsd.lWidth
    Exit Sub
create_error:
    MsgBox "Direct Draw Error : " & err.Description & " <No:" & err.Number & ">"
    mode1 = 0
End Sub
