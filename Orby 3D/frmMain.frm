VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Orby 3D"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrBGM 
      Interval        =   1000
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox viewport 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3600
      Left            =   2400
      ScaleHeight     =   240
      ScaleMode       =   0  'User
      ScaleWidth      =   320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   4800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Orby 3D RayCaster
' Version 0.1.4 (v4a)
'   version format: (release.demo.alpha)
'
' By: Hou Xiong
'
' Features:
' - Recursive ray casting
' - multi-layers
' - alpha layers
' - water animation
' - basic sprites
' - sprite z-sorting
' - custom surface colors

' About this version:
' - uses a direct pointer to the backbuffer to draw the surfaces
' - uses an rgb array to hold the map data instead of using GetPixel
' - supports alpha blending for water
' - supports window glass
' - supports calculation for number of frames per second
' - supports fully-working map display
' - bulging fixed
' - supports adjustable player size
' - does not support sprite clipping


' NOTE: if the program ends unexpectedly or is forced to stop, VB6 might
' crash because the custom pointers to the buffers will not have been
' cleared and VB will try to clear it.  Avoid the Stop button and End command.
' Remember to comment out the ShowMouse commands when debugging.
'
' There seems to be lots of code here because i decided to put
' practically everything in this class form, which I shouldn't if it
' was to be a real game.  However, about 50% of it goes to just
' loading everything.
'
' Enjoy!!


' Soundtracks:
' 1 - Soul Blazer
' 2 - Super Mario Kart
' 3 - Resident Evil 2
' 4 \ RPGM2K (ASCII/Enterbrain)
' 5 /
' 6 \
' 7  > Doom
' 8 /


'API functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Function VarPtrArray Lib "MSVBVM60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Type BITMAPINFOHEADER
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
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Const BLACKNESS = &H42
Private Const PS_SOLID = 0
Private Const WHITENESS = &HFF0062
Private Const TEXT_TRANSPARENT = 1
Private Const TEXT_OPAQUE = 2

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type StaticSprite
    X As Long
    Y As Long
    Z As Long
    Visible As Byte
    xScreen As Long
    Y1 As Long
    Y2 As Long
    clipLeft As Long
    clipRight As Long
    layer As Long
End Type

Private Const IMAGE_BITMAP = 0
Private Const LR_LOADFROMFILE = 16

'Trig constants
Private Const PI = 3.141592653
Private Const DEGREES = 180 / PI
Private Const RADIANS = PI / 180

' ** Map Legend **
' This legend is different in the older verions.
'
' White:    empty space
' Black:    solid wall
' Magenta:  door way
' Red:      window
' Yellow:   ledge
' Blue:     blue carpet
' Cyan:     water
' Green:    orb

Private Const mNone = 16777215     'white  : BGR(255, 255, 255)
Private Const mWall = 0            'black  : BGR(  0,   0,   0)
Private Const mDoorWay = 16711935  'magenta: BGR(  255, 0, 255)
Private Const mWindow = 16711680   'red    : BGR(255,   0,   0)
Private Const mLedge = 16776960    'yellow : BGR(255, 255,   0)
Private Const mCarpet = 255        'blue   : BGR(  0,   0, 255)
Private Const mWater = 65535       'cyan   : BGR(  0, 255, 255)
Private Const mOrb = 65280        'green  : BGR(  0, 255,   0)
Private Const mPlayerStart = 32768 'dkgreen: BGR(  0, 128,   0)
Private Const mPlayerDir = 128     'navy   : BGR(  0,   0, 128)

'Screen attributes
Dim angle As Single
Dim hFov As Long
Dim vFov As Long
Dim fsc As Single
Dim fog As Long
Dim zDetail As Single

Dim key() As Boolean

'look up tables
Private Cosine() As Single
Private Sine() As Single


'Player attributes
Dim speed As Single
Dim pHeight As Single
Dim pSize As Long

'Graphic memory pointers
Dim bgDc As Long                'background
Dim bgBmp As Long
Dim mapDc As Long               'map display
Dim mapBmp As Long
Dim maskDc As Long              'map display mask
Dim maskBmp As Long
Dim mapSize As POINTAPI
Dim mapBits() As Long
Dim bufferDC As Long            'backbuffer
Dim buffBmp As Long
Dim bBuffer() As RGBQUAD
Dim lBuffer() As Long
Dim bDib As Long
Dim waterDc As Long             'water buffer
Dim waterBmp As Long
Dim waterDib As Long
Dim waterBBuffer() As RGBQUAD
Dim waterLBuffer() As Long
Dim glassDc As Long             'glass buffer
Dim glassBmp As Long
Dim glassDib As Long
Dim glassBBuffer() As RGBQUAD
Dim glassLBuffer() As Long

'pens
Dim pBlack As Long
Dim pRed As Long
Dim pGreen As Long
Dim pDkGreen As Long
Dim pWhite As Long

'Orb data
Dim orbBits() As Long
Private Const vOrbsMax = 50
Dim vOrbs() As StaticSprite   'array of orbs to draw
Dim vOrbsCount As Long                   'orbs count
Dim orbSeen() As Byte
Dim orbCount As Long

'Map attributes                 loaded from map
Dim WallColor As RGBQUAD        'Map(0,0)
Dim CarpetColor As RGBQUAD      'Map(1,0)
Dim FloorColor As RGBQUAD       'Map(2,0)
Dim CeilingColor As RGBQUAD     'Map(3,0)
Dim GlassColor As Long          'Map(4,0)
Dim WaterColor As RGBQUAD       'Map(5,0)
Dim WaterFloorColor As RGBQUAD  'Map(6,0)

'color tables
Dim wallS() As Long
Dim floorS() As Long
Dim waterS() As Long
Dim waterFloorS() As Long

'water animation variables
Dim wa2 As Single, wa1 As Single

Dim ot As Long, fps As Long, oFps As Long

Dim showMap As Boolean
Dim cTrack As Long
Dim endLevel As Boolean

Dim cm As POINTAPI  ' current mouse pos
Dim mc As POINTAPI  ' mouse center pos

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    key(KeyCode) = True
    Select Case KeyCode
        Case vbKeyShift
            speed = 2
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    key(KeyCode) = False
    
    Select Case KeyCode
        Case vbKeyEscape
            timeToEnd = True
        Case vbKeyM
            showMap = Not showMap
        Case vbKeyF
            showFPS = Not showFPS
        Case vbKeyShift
            speed = 1
    End Select
End Sub

'center screen
Private Sub Form_Resize()
    viewport.Left = (Me.ScaleWidth - viewport.Width) / 2
    viewport.Top = (Me.ScaleHeight - viewport.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timeToEnd = True
End Sub

'Using a timer simulates a new thread to start the main loop instead of starting the loop
'on the form Loading thread, which resolves some Loading/Unloading issues.
Private Sub Timer1_Timer()
    Dim i As Long, X As Long, Y As Long, Z As Single
    
    Randomize
    
    'Disable the timer.
    'The timer was only used to start the main continuous loop.
    Timer1.Enabled = False
    
    'create pens
    pBlack = CreatePen(PS_SOLID, 1, vbBlack)
    pRed = CreatePen(PS_SOLID, 1, vbRed)
    pGreen = CreatePen(PS_SOLID, 1, vbGreen)
    pDkGreen = CreatePen(PS_SOLID, 1, RGB(0, 128, 0))
    pWhite = CreatePen(PS_SOLID, 1, vbWhite)
    
    ' Map transparent masking
    'mapBmp = CreateCompatibleBitmap(Me.hDC, mapSize.X, mapSize.Y)
    mapBmp = LoadImage(0, App.Path & "\Maps\" & LevelList(curLevel - 1), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    mapDc = CreateCompatibleDC(Me.hdc)
    SelectObject mapDc, mapBmp
    
    Dim bmp As BITMAP
    GetObject mapBmp, LenB(bmp), bmp
    mapSize.X = bmp.bmWidth
    mapSize.Y = bmp.bmHeight
    
        ' get map dibits
        Dim bi As BITMAPINFO
        With bi.bmiHeader
            .biSize = LenB(bi)
            .biWidth = mapSize.X
            .biHeight = -mapSize.Y
            .biPlanes = 1
            .biBitCount = 32
            .biSizeImage = ((((.biWidth * 32) + 31) \ 32) * 4) * Abs(.biHeight)
        End With
        
        ReDim mapBits(mapSize.X - 1, mapSize.Y - 1)
        GetDIBits mapDc, mapBmp, 0, mapSize.Y, mapBits(0, 0), bi, 0
        ' end get bits
    
    'create monochrome bitmap for mask
    maskBmp = CreateBitmap(mapSize.X, mapSize.Y, 1, 1, ByVal 0)
    maskDc = CreateCompatibleDC(Me.hdc)
    SelectObject maskDc, maskBmp
    
    BitBlt maskDc, 0, 0, mapSize.X, mapSize.Y, mapDc, 0, 0, vbSrcCopy
    BitBlt mapDc, 0, 0, mapSize.X, mapSize.Y, maskDc, 0, 0, vbSrcInvert
    ' end masking
    
    ' create dib section for backbuffer
    bufferDC = CreateCompatibleDC(Me.hdc)
    
    Dim bbi As BITMAPINFO
    With bbi.bmiHeader
        .biSize = LenB(bbi)
        .biWidth = 320
        .biHeight = -240
        .biPlanes = 1
        .biBitCount = 32
        .biSizeImage = ((((.biWidth * 32) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    
    buffBmp = CreateDIBSection(Me.hdc, bbi, 0, bDib, 0, 0)
    SelectObject bufferDC, buffBmp
    
    Dim sa As SAFEARRAY2D
    With sa
        .cbElements = 4
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Abs(bbi.bmiHeader.biHeight)
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bbi.bmiHeader.biWidth
        .pvData = bDib
    End With
    
    ' make bBuffer and lBuffer point to dibsection,
    ' this will gain direct access to bitmap color bits
    ' unlike GetDibits above where they are only copied.
    '
    ' The reason for two different pointers is to allow one to be
    ' in the Long format used for fast comparisons, and the other one
    ' to be in the RGB type format used for alhpa blending.
    ' Ultimately, this means there is no need to convert
    ' one format to the other which improves speed.  They both point
    ' to the same chunk of data, so editing either one will affect the bitmap.
    CopyMemory ByVal VarPtrArray(bBuffer()), VarPtr(sa), 4  'byte bgr format
    CopyMemory ByVal VarPtrArray(lBuffer()), VarPtr(sa), 4  'long bgr format
    
    SetBkMode bufferDC, TEXT_TRANSPARENT
    SetBkColor bufferDC, vbWhite
    ' end backbuffer dib creation
    
    'Load the sprites
    ReDim vOrbs(vOrbsMax - 1)
    Dim orbBmp As Long
    Dim orbBI As BITMAPINFO
    orbBmp = LoadImage(0, App.Path & "\Sprites\Orb.bmp", IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    With orbBI.bmiHeader
        .biSize = LenB(orbBI)
        .biWidth = 64
        .biHeight = -64
        .biPlanes = 1
        .biBitCount = 32
        .biSizeImage = ((((.biWidth * 32) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    ReDim orbBits(0 To 63, 0 To 63)
    GetDIBits Me.hdc, orbBmp, 0, 64, orbBits(0, 0), orbBI, 0
    DeleteObject orbBmp
    'end load sprites
    
    ' allocate array for visible orbs to be drawn
    ReDim orbSeen(mapSize.X - 1, mapSize.Y - 1)
    
    ' create dib section for water layer
    waterDc = CreateCompatibleDC(Me.hdc)
    waterBmp = CreateDIBSection(Me.hdc, bbi, 0, waterDib, 0, 0)
    SelectObject waterDc, waterBmp
    Dim sa2 As SAFEARRAY2D
    With sa2
        .cbElements = 4
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Abs(bbi.bmiHeader.biHeight)
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bbi.bmiHeader.biWidth
        .pvData = waterDib
    End With
    CopyMemory ByVal VarPtrArray(waterBBuffer()), VarPtr(sa2), 4  'byte bgr format
    CopyMemory ByVal VarPtrArray(waterLBuffer()), VarPtr(sa2), 4  'byte bgr format
    ' end dib creation
    
    ' create dib section for glass window layer
    Dim gsa As SAFEARRAY2D
    
    glassDc = CreateCompatibleDC(Me.hdc)
    glassBmp = CreateDIBSection(Me.hdc, bbi, 0, glassDib, 0, 0)
    SelectObject glassDc, glassBmp
    With gsa
        .cbElements = 4
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Abs(bbi.bmiHeader.biHeight)
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bbi.bmiHeader.biWidth
        .pvData = glassDib
    End With
    CopyMemory ByVal VarPtrArray(glassBBuffer()), VarPtr(gsa), 4  'byte bgr format
    CopyMemory ByVal VarPtrArray(glassLBuffer()), VarPtr(gsa), 4  'byte bgr format
    ' end dib creation
    
    'game attributes
    hFov = 60    ' horizontal field of view (degrees)
    vFov = 60    ' use same as horizontal for correct proportions (fixes bulging)
    fog = 120    ' visible distance (pixels)
    fsc = 80     ' scale for 3d -> 2d coordinate conversion, greater = taller doors/walls
    pHeight = 50 ' player height (pixels)
    pSize = 3    ' the radius of the area the player covers (pixels) used for collision
    speed = 1    ' in pixels
    
    'load map colors
    CopyMemory WallColor, mapBits(0, 0), 4
    CopyMemory CarpetColor, mapBits(1, 0), 4
    CopyMemory FloorColor, mapBits(2, 0), 4
    CopyMemory CeilingColor, mapBits(3, 0), 4
    GlassColor = mapBits(4, 0)
    CopyMemory WaterColor, mapBits(5, 0), 4
    CopyMemory WaterFloorColor, mapBits(6, 0), 4
    For i = 0 To 6
        SetPixel mapDc, i, 0, vbBlack
        SetPixel maskDc, i, 0, vbWhite
        mapBits(i, 0) = vbWhite
    Next
    cTrack = 255
    For i = 7 To 7 + 9
        If mapBits(i, 0) = vbBlack Then
            cTrack = i - 7
            SetPixel mapDc, i, 0, vbBlack
            SetPixel maskDc, i, 0, vbWhite
            mapBits(i, 0) = vbWhite
            Exit For
        End If
    Next
    
    'create fading color table for walls and blue floor (enhances depth)
    ReDim wallS(fog)
    ReDim floorS(fog)
    ReDim glassS(fog)
    ReDim waterS(fog)
    ReDim waterFloorS(fog)
    For i = 0 To fog
        wallS(i) = RGB((fog - i) / fog * WallColor.rgbBlue, (fog - i) / fog * WallColor.rgbGreen, (fog - i) / fog * WallColor.rgbRed)
        floorS(i) = RGB((fog - i) / fog * CarpetColor.rgbBlue, (fog - i) / fog * CarpetColor.rgbGreen, (fog - i) / fog * CarpetColor.rgbRed)
        waterS(i) = RGB((fog - i) / fog * WaterColor.rgbBlue, (fog - i) / fog * WaterColor.rgbGreen, (fog - i) / fog * WaterColor.rgbRed)
        waterFloorS(i) = RGB((fog - i) / fog * WaterFloorColor.rgbBlue, (fog - i) / fog * WaterFloorColor.rgbGreen, (fog - i) / fog * WaterFloorColor.rgbRed)
    Next
    
    Dim dirPos As POINTAPI
    orbCount = 0
    dirPos.X = 0
    dirPos.Y = 0
    For X = 0 To mapSize.X - 1
        For Y = 0 To mapSize.Y - 1
            If mapBits(X, Y) = mOrb Then
                'count the orbs in the level
                orbCount = orbCount + 1
            ElseIf (mapBits(X, Y) = mPlayerStart) And (pos.X = 0) And (pos.Y = 0) Then
                'load player start location
                pos.X = X
                pos.Y = Y
                mapBits(X, Y) = vbWhite
                SetPixel mapDc, X, Y, vbBlack
                SetPixel maskDc, X, Y, vbWhite
            ElseIf mapBits(X, Y) = mPlayerDir Then
                'load player direction
                dirPos.X = X
                dirPos.Y = Y
                mapBits(X, Y) = vbWhite
                SetPixel mapDc, X, Y, vbBlack
                SetPixel maskDc, X, Y, vbWhite
            End If
        Next
    Next
    
    angle = 0
    If (dirPos.X <> 0) And (dirPos.Y <> 0) And (pos.X <> 0) And (pos.Y <> 0) Then
        'find the starting angle
        If pos.X = dirPos.X Then
            If dirPos.Y >= pos.Y Then
                angle = 90
            Else
                angle = 270
            End If
        ElseIf pos.X > dirPos.X Then
            angle = Atn((dirPos.Y - pos.Y) / (dirPos.X - pos.X)) * DEGREES + 180
        Else
            angle = Atn((dirPos.Y - pos.Y) / (dirPos.X - pos.X)) * DEGREES
        End If
    End If
    'if no default start location then center the player in the map
    If pos.X = 0 Then pos.X = mapSize.X / 2
    If pos.Y = 0 Then pos.Y = mapSize.Y / 2
    
    
    'create lookup tables
    ReDim Cosine(-hFov / 2 To 360 + hFov / 2)
    ReDim Sine(-hFov / 2 To 360 + hFov / 2)
    For i = -hFov / 2 To 360 + hFov / 2
        Cosine(i) = Cos(i * RADIANS)
        Sine(i) = Sin(i * RADIANS)
    Next
    
    'buffer for background (floor and ceiling/sky),
    'this is basically a still image of two different gradients
    bgDc = CreateCompatibleDC(Me.hdc)
    bgBmp = CreateCompatibleBitmap(Me.hdc, 320, 240)
    SelectObject bgDc, bgBmp
    
    'draw the floor and the ceiling (enhances depth)
    For Z = 1 To fog Step 0.05
        Y = fsc * (pHeight - 90) / Z + 120
        SetPixel bgDc, 0, Y, RGB((fog - Z) / fog * CeilingColor.rgbRed, (fog - Z) / fog * CeilingColor.rgbGreen, (fog - Z) / fog * CeilingColor.rgbBlue)
        Y = fsc * pHeight / Z + 120
        SetPixel bgDc, 0, Y, RGB((fog - Z) / fog * FloorColor.rgbRed, (fog - Z) / fog * FloorColor.rgbGreen, (fog - Z) / fog * FloorColor.rgbBlue)
    Next
    StretchBlt bgDc, 1, 0, 319, 241, bgDc, 0, 1, 1, 240, vbSrcCopy
    
    'center mouse
    mc.X = Screen.Width / Screen.TwipsPerPixelX / 2
    mc.Y = Screen.Height / Screen.TwipsPerPixelY / 2
    SetCursorPos mc.X, mc.Y
    ShowCursor 0   'hide cursor
    
    'initialize level variables
    wa2 = 0
    endLevel = False
    timeToEnd = False
    tmrBGM.Enabled = True
    ReDim key(255)
    
    'play music
    If cTrack > 8 Then
        cTrack = Int(Rnd * 8) + 1
    End If
    PlayBGM cTrack
    
    FadeIn
    
    ot = GetTickCount
    
    'Main game loop
    Do While Not timeToEnd
        DoEvents
        
        RenderScene
        FlipBackBuffer
        Movement
    Loop
    
    FadeOut
    
    'clean up
    CopyMemory ByVal VarPtrArray(bBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(lBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(waterBBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(waterLBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassBBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassLBuffer()), 0&, 4

    DeleteObject buffBmp
    DeleteDC bufferDC
    DeleteObject bgBmp
    DeleteDC bgDc
    DeleteObject mapBmp
    DeleteDC mapDc
    DeleteObject maskBmp
    DeleteDC maskDc
    DeleteObject waterBmp
    DeleteDC waterDc
    DeleteObject pBlack
    DeleteObject pRed
    DeleteObject pGreen
    DeleteObject pDkGreen
    
    ShowCursor 1
    
    StopBGM cTrack
    
    If endLevel Then
        curLevel = curLevel + 1
        If curLevel - 1 < numLevels Then
            pos.X = 0
            pos.Y = 0
            tmrBGM.Enabled = False
            Timer1.Interval = 1000
            Timer1.Enabled = True
            Exit Sub
        End If
    End If
    
    Unload Me
End Sub

Private Sub Movement()
    'update mouse rotation
    GetCursorPos cm
    angle = angle + (cm.X - mc.X) / 4
    If angle > 360 Then
        angle = angle - 360
    ElseIf angle < 0 Then
        angle = angle + 360
    End If
    SetCursorPos mc.X, mc.Y
    
    'detect keys and adjust movement direction
    Dim C As Long, a As Single, i As Long, j As Long
    
    If key(vbKeyW) And Not key(vbKeyD) And Not key(vbKeyS) And key(vbKeyA) Then         'up-left
        a = angle - 45
    ElseIf Not key(vbKeyW) And Not key(vbKeyD) And key(vbKeyS) And key(vbKeyA) Then     'down-left
        a = angle - 90 - 45
    ElseIf key(vbKeyW) And key(vbKeyD) And Not key(vbKeyS) And Not key(vbKeyA) Then     'up-right
        a = angle + 45
    ElseIf Not key(vbKeyW) And key(vbKeyD) And key(vbKeyS) And Not key(vbKeyA) Then     'down-right
        a = angle + 90 + 45
    ElseIf Not key(vbKeyW) And key(vbKeyD) And Not key(vbKeyS) And Not key(vbKeyA) Then 'right
        a = angle + 90
    ElseIf key(vbKeyW) And Not key(vbKeyD) Or Not key(vbKeyS) And Not key(vbKeyA) Then  'up
        a = angle
    ElseIf Not key(vbKeyW) And Not key(vbKeyD) And key(vbKeyS) And Not key(vbKeyA) Then 'down
        a = angle - 180
    ElseIf Not key(vbKeyW) And Not key(vbKeyD) And Not key(vbKeyS) And key(vbKeyA) Then 'left
        a = angle - 90
    End If
    
    If a > 360 Then
        a = a - 360
    ElseIf a < 0 Then
        a = a + 360
    End If
    
    Dim a1 As Single
    a1 = a
    
    'move then handle collision
    On Error Resume Next
    If key(vbKeyW) Or key(vbKeyD) Or key(vbKeyS) Or key(vbKeyA) Then
        'move along the X axis
        pos.X = pos.X + Cosine(a1) * speed
        For i = -pSize To pSize
            For j = -pSize To pSize
                C = mapBits(pos.X + i, pos.Y + j)
                If (C = mOrb) Then
                    'collision with orb
                    mapBits(pos.X + i, pos.Y + j) = vbWhite
                    SetPixel mapDc, pos.X + i, pos.Y + j, vbBlack
                    SetPixel maskDc, pos.X + i, pos.Y + j, vbWhite
                    PlaySE
                    orbCount = orbCount - 1
                    If orbCount = 0 Then
                        endLevel = True
                        timeToEnd = True
                    End If
                ElseIf (C <> mNone) And (C <> mDoorWay) And (C <> mCarpet) And (C <> mWater) Then
                    'collision with wall, move back
                    pos.X = pos.X - Cosine(a1) * speed
                    j = pSize
                    Exit For
                End If
            Next
        Next
        'move along the Y axis
        pos.Y = pos.Y + Sine(a1) * speed
        For i = -pSize To pSize
            For j = -pSize To pSize
                C = mapBits(pos.X + i, pos.Y + j)
                If (C = mOrb) Then
                    'collision with orb
                    mapBits(pos.X + i, pos.Y + j) = vbWhite
                    SetPixel mapDc, pos.X + i, pos.Y + j, vbBlack
                    SetPixel maskDc, pos.X + i, pos.Y + j, vbWhite
                    PlaySE
                    orbCount = orbCount - 1
                    If orbCount = 0 Then
                        endLevel = True
                        timeToEnd = True
                    End If
                ElseIf (C <> mNone) And (C <> mDoorWay) And (C <> mCarpet) And (C <> mWater) Then
                    'collision with wall, move back
                    pos.Y = pos.Y - Sine(a1) * speed
                    j = pSize
                    Exit For
                End If
            Next
        Next
    End If
End Sub

Private Sub RenderScene()
    'Clear buffers
    BitBlt bufferDC, 0, 0, 320, 240, bgDc, 0, 0, vbSrcCopy
    ZeroMemory ByVal waterDib, 307200  '320 * 240 * 4 (width * height * 4 bytes)
    ZeroMemory ByVal glassDib, 307200
    ZeroMemory orbSeen(0, 0), mapSize.X * mapSize.Y
    
    Dim i As Long, a As Single, a2 As Single, va As Single
    
    vOrbsCount = 0
    wa1 = wa1 + 10  'wave z component, increase this value for more waves going into the screen
    If wa1 > 360 Then wa1 = wa1 - 360
    wa2 = wa1
    zDetail = 1 'ray casting step, the smaller the smoother but slower
    'scan through the screen (width of 320)
    For i = 0 To 319
        a = ((i + 1) / 320 * hFov) - hFov / 2   'current ray angle
        va = ((i + 1) / 320 * vFov) - vFov / 2  'current vertical ray angle
        wa2 = wa2 + 2                 'wave x component, increase this value for more waves going across screen
        a2 = (a + angle) * RADIANS
        'scan along the rays outwards
        '                    |-> this cancels out the Z bulging caused by using only 1 field of view
        '                    |---------------------|
        rayRecur i, 1, fsc * (2 - Cos(va * RADIANS)), Cos(a2), Sin(a2), 0, 240, False
    Next
    
    'Blend/add in the water, glass, and sprites layers
    BlendLayers
    
    Dim strT As String, lenT As Long
    'display orb count
    If orbCount > 0 Then
        strT = orbCount
        lenT = Len(strT)
        SetTextColor bufferDC, vbBlack
        TextOut bufferDC, 251, 6, "Orbs: " & strT, 6 + lenT
        
        SetTextColor bufferDC, vbWhite
        TextOut bufferDC, 250, 5, "Orbs: " & strT, 6 + lenT
    End If
    
    '** Possibly implement an auto frame skip/limit algorithm **
    If showFPS Then
        fps = fps + 1
        If GetTickCount - ot >= 500 Then
            oFps = fps * 2
            fps = 0
            ot = GetTickCount
        End If
        TextOut bufferDC, 250, 220, "FPS: " & oFps, 5 + Len(Str$(oFps))
    End If
    
    'draw the transparent map and then draw player location.
    Dim fovLeft As POINTAPI, fovRight As POINTAPI
    Dim fovLeft2 As POINTAPI, fovRight2 As POINTAPI
    If showMap Then
        SetTextColor bufferDC, vbBlack
        BitBlt bufferDC, 0, 0, fog * 2, fog * 2, maskDc, pos.X - fog, pos.Y - fog, vbSrcAnd
        BitBlt bufferDC, 0, 0, fog * 2, fog * 2, mapDc, pos.X - fog, pos.Y - fog, vbSrcPaint
        
        fovLeft.X = fog + Cos((angle + hFov / 2) * RADIANS) * fog
        fovLeft.Y = fog + Sin((angle + hFov / 2) * RADIANS) * fog
        fovRight.X = fog + Cos((angle - hFov / 2) * RADIANS) * fog
        fovRight.Y = fog + Sin((angle - hFov / 2) * RADIANS) * fog
        fovLeft2.X = fog + Cos((angle + hFov / 2) * RADIANS) * 10
        fovLeft2.Y = fog + Sin((angle + hFov / 2) * RADIANS) * 10
        fovRight2.X = fog + Cos((angle - hFov / 2) * RADIANS) * 10
        fovRight2.Y = fog + Sin((angle - hFov / 2) * RADIANS) * 10

        SelectObject bufferDC, pGreen
        MoveToEx bufferDC, fovLeft.X, fovLeft.Y, 0&
        LineTo bufferDC, fog, fog
        LineTo bufferDC, fovRight.X, fovRight.Y
        SelectObject bufferDC, pDkGreen
        LineTo bufferDC, fovLeft.X, fovLeft.Y
        SelectObject bufferDC, pRed
        MoveToEx bufferDC, fovLeft2.X, fovLeft2.Y, 0&
        LineTo bufferDC, fovRight2.X, fovRight2.Y
        Ellipse bufferDC, -pSize + fog, -pSize + fog, pSize + fog, pSize + fog
    End If
End Sub

Private Sub FlipBackBuffer()
    'copy the back buffer to front buffer (in this case, the viewport)
    BitBlt viewport.hdc, 0, 0, 320, 240, bufferDC, 0, 0, vbSrcCopy
    'StretchBlt viewport.hDC, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, 320, 240, vbSrcCopy
End Sub

Private Sub FadeOut()
    Dim X As Long, Y As Long, i As Long
    Dim tempBuffer() As RGBQUAD
    
    RenderScene
    
    ReDim tempBuffer(319, 239)
    
    CopyMemory tempBuffer(0, 0), bBuffer(0, 0), 307200 '320 * 240 * 4
    
    For i = 50 To 0 Step -1
        DoEvents
        For X = 0 To 319
            For Y = 0 To 239
                bBuffer(X, Y).rgbBlue = (tempBuffer(X, Y).rgbBlue * (i / 50))
                bBuffer(X, Y).rgbGreen = (tempBuffer(X, Y).rgbGreen * (i / 50))
                bBuffer(X, Y).rgbRed = (tempBuffer(X, Y).rgbRed * (i / 50))
            Next
        Next
        
        'copy the back buffer to front buffer (in this case, the viewport)
        BitBlt viewport.hdc, 0, 0, 320, 240, bufferDC, 0, 0, vbSrcCopy
        'StretchBlt viewport.hDC, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, 320, 240, vbSrcCopy
    Next
End Sub

Private Sub FadeIn()
    Dim X As Long, Y As Long, i As Long
    Dim tempBuffer() As RGBQUAD
    
    RenderScene
    
    ReDim tempBuffer(319, 239)
    
    CopyMemory tempBuffer(0, 0), bBuffer(0, 0), 307200 '320 * 240 * 4
    
    For i = 0 To 50
        DoEvents
        For X = 0 To 319
            For Y = 0 To 239
                bBuffer(X, Y).rgbBlue = (tempBuffer(X, Y).rgbBlue * (i / 50))
                bBuffer(X, Y).rgbGreen = (tempBuffer(X, Y).rgbGreen * (i / 50))
                bBuffer(X, Y).rgbRed = (tempBuffer(X, Y).rgbRed * (i / 50))
            Next
        Next
        
        'copy the back buffer to front buffer (in this case, the viewport)
        BitBlt viewport.hdc, 0, 0, 320, 240, bufferDC, 0, 0, vbSrcCopy
        'StretchBlt viewport.hDC, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, 320, 240, vbSrcCopy
    Next
End Sub

'** Possibly optimizable? **
Private Sub rayRecur(ByVal i As Long, ByVal j As Single, ByVal vScl As Single, ByVal xp As Single, ByVal yp As Single, ByVal wTop As Long, ByVal wBottom As Long, ByVal glass As Boolean)
    Dim C As Long, j2 As Long, w1 As Single, X As Long, Y As Long
    Dim Y1 As Long, Y2 As Long, c2 As Long, cGlass As Boolean
    
    'start or continue casting ray where parent function left off
    cGlass = glass
    Do While j < fog
        j = j + zDetail
        'zDetail = zDetail + 0.1    'attempted detail algorithm
            'calculate current point of ray
            X = pos.X + xp * j
            Y = pos.Y + yp * j
            C = -1
            If (X >= 0) And (X < mapSize.X) And (Y >= 0) And (Y < mapSize.Y) Then
                C = mapBits(X, Y)
            End If
            If C = -1 Then
                'no further raycasting is needed, exit and move to adjacent ray
                Exit Sub
            ElseIf C = mWall Then
                DrawVLine i, (vScl * (pHeight - 90) / j) + 120, (vScl * pHeight / j) + 120, wallS(j), wTop, wBottom
                'no further raycasting is needed, exit and move to adjacent ray
                Exit Sub
                
            '** Calculate a window region to constrain the drawing behind this surface, **
            '** these bounds will be passed into the next recursive call.               **
            '
            '** Possibly implement lookup tables for these y values instead?            **
            ElseIf C = mDoorWay Then
                Y1 = (vScl * (pHeight - 80) / j) + 120
                Y2 = (vScl * (pHeight) / j) + 120
                Exit Do
            ElseIf C = mLedge Then
                Y1 = (vScl * (pHeight - 90) / j) + 120
                Y2 = (vScl * (pHeight - 20) / j) + 120
                Exit Do
            ElseIf C = mWindow Then
                Y1 = (vScl * (pHeight - 70) / j) + 120
                Y2 = (vScl * (pHeight - 35) / j) + 120
                glass = True
                Exit Do
            ElseIf C = mCarpet Then
                Y1 = (vScl * (pHeight - 90) / j) + 120
                Y2 = (vScl * (pHeight - 2) / j) + 120
                Exit Do
            ElseIf (C <> mNone) And (C <> -1) Then
                Y1 = wTop
                Y2 = wBottom
                'Another surface with an opening ("window") has been detected, exit this loop and enter
                'a new loop in the recursion to find surfaces behind this one.
                Exit Do
            End If
    Loop
    
    'update the min and max of the drawing window region
    If Y1 < wTop Then
        Y1 = wTop
    ElseIf Y1 < 0 Then
        Y1 = 0
    End If
    If Y2 > wBottom Then
        Y2 = wBottom
    ElseIf Y2 >= 320 Then
        Y2 = 239
    End If
    'call this function again
    If j < fog Then rayRecur i, j, vScl, xp, yp, Y1, Y2, glass
    
    '** Check and draw only the visible part seen through the  **
    '** window region of the surface in front of this surface. **
    '** The bounds for this calculation will be passed in from **
    '** the parent function, which is the surface in front of  **
    '** this one.                                              **
    '** This allows drawing water on a separate                **
    '** buffer, then alpha blended to backbuffer in the end.   **
    '** Similarly, glass windows and semi-transparent surfaces **
    '** can be processed this way on a separate buffer.        **
    '** Keep in mind, the more layers, the more complex        **
    '** the Z-sorting is.  See the attempted version (3.5a)    **
    '** of multi-layered glass.                                **
    
    Select Case C
        Case mDoorWay
            DrawVLine i, (vScl * (pHeight - 90) / j) + 120, (vScl * (pHeight - 80) / j) + 120, wallS(j), wTop, wBottom
        Case mLedge
            DrawVLine i, (vScl * (pHeight - 20) / j) + 120, (vScl * pHeight / j) + 120, wallS(j), wTop, wBottom
        Case mWindow
            DrawVLine i, (vScl * (pHeight - 35) / j) + 120, (vScl * pHeight / j) + 120, wallS(j), wTop, wBottom
            DrawVLine i, (vScl * (pHeight - 90) / j) + 120, (vScl * (pHeight - 70) / j) + 120, wallS(j), wTop, wBottom
            'Skip if glass has already been drawn, improves speed?
            If Not cGlass Then
                DrawGLine i, Y1, Y2, GlassColor, Y1, Y2
            End If
        Case mCarpet
            DrawVLine i, (vScl * (pHeight - 2) / j) + 120, (vScl * pHeight / j) + 120, floorS(j), wTop, wBottom
        Case mWater
            w1 = Cos(wa2 * RADIANS)
            j2 = j + w1 * 20
            If j2 < 0 Then
                j2 = 0
            ElseIf j2 > fog Then
                j2 = fog
            End If
            DrawWLine i, (vScl * (pHeight - 16 + w1) / j) + 120, (vScl * pHeight / j) + 120, waterS(j2), wTop, wBottom
            'custom colored floor instead of default color floor, may be commented out for possible speed gain
            DrawVLine i, (vScl * (pHeight - 2) / j) + 120, (vScl * pHeight / j) + 120, waterFloorS(j), wTop, wBottom
        Case mOrb
            'Prevents the recording of the same orb more than once.
            'This can occur because as we scan across the screen,
            'a different ray will hit the same pixel a few times.
            'This may be inefficient, but results in smoother rendering.
            If orbSeen(X, Y) = 0 Then
                'the orb has not been detected, so record it for later drawing
                With vOrbs(vOrbsCount)
                    .xScreen = i
                    .Z = j
                    w1 = Cos(wa2 * RADIANS)
                    .Y1 = (vScl * (pHeight - 45 + w1) / j) + 120
                    .Y2 = (vScl * (pHeight - 40 + w1) / j) + 120
                    If cGlass Then
                        .layer = 2
                    Else
                        .layer = 1
                    End If
                End With
                vOrbsCount = vOrbsCount + 1
                If vOrbsCount >= vOrbsMax Then vOrbsCount = vOrbsMax
                orbSeen(X, Y) = 1
            End If
    End Select
End Sub

'The next 3 functions are practically identical except the buffer to draw to.
'I decided to use separate functions for less confusion.

'Draws the walls (a vertical line)
'Note: y1 must be less than y2 because the line will be drawn top to bottom
Private Sub DrawVLine(ByVal X As Long, ByVal Y1 As Long, ByVal Y2 As Long, ByVal color As Long, ByVal wTop As Long, ByVal wBottom As Long)
    'is the surface even in the visible screen
    If Y1 >= wBottom Then
        Exit Sub
    'is the surface in the visible window
    ElseIf Y1 < wTop Then
        Y1 = wTop
    End If
    'is the surface in the visible screen
    If Y2 <= wTop Then
        Exit Sub
    'is the surface in the visible window
    ElseIf Y2 > wBottom Then
        Y2 = wBottom
    End If
    'the surface is visible so draw it
    For Y1 = Y1 To Y2 - 1
        lBuffer(X, Y1) = color
    Next
End Sub
'Draws the water
Private Sub DrawWLine(ByVal X As Long, ByVal Y1 As Long, ByVal Y2 As Long, ByVal color As Long, ByVal wTop As Long, ByVal wBottom As Long)
    If Y1 >= wBottom Then
        Exit Sub
    ElseIf Y1 < wTop Then
        Y1 = wTop
    End If
    If Y2 <= wTop Then
        Exit Sub
    ElseIf Y2 > wBottom Then
        Y2 = wBottom
    End If
    For Y1 = Y1 To Y2 - 1
        waterLBuffer(X, Y1) = color
    Next
End Sub
'Draws the glass
Private Sub DrawGLine(ByVal X As Long, ByVal Y1 As Long, ByVal Y2 As Long, ByVal color As Long, ByVal wTop As Long, ByVal wBottom As Long)
    If Y1 >= wBottom Then
        Exit Sub
    ElseIf Y1 < wTop Then
        Y1 = wTop
    End If
    If Y2 <= wTop Then
        Exit Sub
    ElseIf Y2 > wBottom Then
        Y2 = wBottom
    End If
    For Y1 = Y1 To Y2 - 1
        glassLBuffer(X, Y1) = color
    Next
End Sub

'draws the orbs
Private Sub DrawOrb(ByVal X As Long, ByVal Y1 As Long, ByVal Y2 As Long)
    Dim xx As Long, yy As Long, u As Single, v As Single, du As Single, dv As Single
    Dim Width As Long
    
    If Y1 >= 240 Then Exit Sub
    Width = Y2 - Y1
    X = X - Width / 2
    
    du = (64 - 1) / Width
    dv = (64 - 1) / Width
    
    'simple texture mapping/interpolation
    v = 0
    For yy = Y1 To Y2
        u = 0
        For xx = X To X + Width
            If orbBits(u, v) <> vbBlack Then
                If (xx >= 0) And (xx < 320) And (yy >= 0) And (yy <= 239) Then
                    lBuffer(xx, yy) = orbBits(u, v)
                End If
            End If
            u = u + du
        Next
        v = v + dv
    Next
End Sub

'quick sort the sprites
Private Sub QuickZSort(ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, mv As Long
    
    If last <= first Then Exit Sub
    
    low = first
    high = last
    mv = vOrbs((first + last) \ 2).Z
    
    Do
        While vOrbs(low).Z < mv
            low = low + 1
        Wend
        While vOrbs(high).Z > mv
            high = high - 1
        Wend
        If low <= high Then
            Swap vOrbs(low), vOrbs(high)
            low = low + 1
            high = high - 1
        End If
    Loop While low <= high
    
    If first < high Then QuickZSort first, high
    If low < last Then QuickZSort low, last
End Sub

Private Sub Swap(ByRef a As StaticSprite, ByRef b As StaticSprite)
    Dim sp As StaticSprite, spLen As Long
    
    spLen = LenB(sp)
    
    CopyMemory sp, a, spLen
    CopyMemory a, b, spLen
    CopyMemory b, sp, spLen
End Sub


'Blend the water, glass, and sprites layers onto the backbuffer
Private Sub BlendLayers()
    Dim X As Long, Y As Long, C As Single, i As Long, j As Long
    
    'sort sprites
    QuickZSort 0, vOrbsCount - 1
    
    For Y = 0 To 239
        For X = 0 To 319
            'alpha blend the water
            If waterLBuffer(X, Y) <> vbBlack Then
                C = (CSng(waterBBuffer(X, Y).rgbBlue) - bBuffer(X, Y).rgbBlue) * 0.3 + bBuffer(X, Y).rgbBlue
                bBuffer(X, Y).rgbBlue = C
                C = (CSng(waterBBuffer(X, Y).rgbGreen) - bBuffer(X, Y).rgbGreen) * 0.3 + bBuffer(X, Y).rgbGreen
                bBuffer(X, Y).rgbGreen = C
                C = (CSng(waterBBuffer(X, Y).rgbRed) - bBuffer(X, Y).rgbRed) * 0.3 + bBuffer(X, Y).rgbRed
                bBuffer(X, Y).rgbRed = C
            End If
        Next
    Next
    
    'draw the first layer of sprites before glass but after water
    For i = vOrbsCount - 1 To 0 Step -1
        If vOrbs(i).layer = 2 Then DrawOrb vOrbs(i).xScreen, vOrbs(i).Y1, vOrbs(i).Y2
    Next
    
    For Y = 0 To 239
        For X = 0 To 319
            'alpha blend the glass
            If glassLBuffer(X, Y) <> vbBlack Then
                C = (CSng(glassBBuffer(X, Y).rgbBlue) - bBuffer(X, Y).rgbBlue) * 0.4 + bBuffer(X, Y).rgbBlue
                bBuffer(X, Y).rgbBlue = C
                C = (CSng(glassBBuffer(X, Y).rgbGreen) - bBuffer(X, Y).rgbGreen) * 0.4 + bBuffer(X, Y).rgbGreen
                bBuffer(X, Y).rgbGreen = C
                C = (CSng(glassBBuffer(X, Y).rgbRed) - bBuffer(X, Y).rgbRed) * 0.4 + bBuffer(X, Y).rgbRed
                bBuffer(X, Y).rgbRed = C
            End If
        Next
    Next
    
    'draw the rest of the sprites
    For i = vOrbsCount - 1 To 0 Step -1
        If vOrbs(i).layer = 1 Then DrawOrb vOrbs(i).xScreen, vOrbs(i).Y1, vOrbs(i).Y2
    Next
End Sub

Private Sub tmrBGM_Timer()
    RepeatBGM cTrack
End Sub

