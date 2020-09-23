VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Play - Recursive Ray Tracer"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3
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
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Version 3.5 Alpha

' About this version:
' - uses a direct pointer to the backbuffer to draw the surfaces
' - uses an rgb array to hold the map data instead of using GetPixel
' - supports alpha blending for water
' - supports window glass
' - bulging fixed
' - **attempted to add alpha multi-layered glass windows
' - supports semi-working map display

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function GetTickCount Lib "KERNEL32" () As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Sub FillMemory Lib "KERNEL32" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
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

Private Const BLACKNESS = &H42
Private Const PS_SOLID = 0
Private Const WHITENESS = &HFF0062

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

Private Const PI = 3.141592653
Private Const DEGREES = 180 / PI
Private Const RADIANS = PI / 180

' ** Map Legend **
'
' White:    empty space
' Black:    solid wall
' Green:    door way
' Red:      window
' Yellow:   ledge
' Blue:     blue carpet
' Cyan:     water

Private Const mNone = 16777215     'white  : BGR(255, 255, 255)
Private Const mWall = 0            'black  : BGR(  0,   0,   0)
Private Const mDoorWay = 65280     'green  : BGR(  0, 255,   0)
Private Const mWindow = 16711680   'red    : BGR(255,   0,   0)
Private Const mLedge = 16776960    'yellow : BGR(255, 255,   0)
Private Const mCarpet = 255        'blue   : BGR(  0,   0, 255)
Private Const mWater = 65535       'cyan   : BGR(  0, 255, 255)

Dim angle As Single
Dim hFov As Long
Dim vFov As Long
Dim fsc As Single
Dim fog As Long
Dim key(255) As Boolean
Dim zDetail As Single

Dim speed As Single
Dim pHeight As Single

Dim bgDc As Long
Dim bgBmp As Long
Dim mapDc As Long
Dim mapBmp As Long
Dim maskDc As Long
Dim maskBmp As Long
Dim mapSize As POINTAPI
Dim mapBits() As Long
Dim bufferDC As Long
Dim buffBmp As Long
Dim bBuffer() As RGBQUAD
Dim lBuffer() As Long
Dim bDib As Long
Dim waterDc As Long
Dim waterBmp As Long
Dim waterDib As Long
Dim waterBBuffer() As RGBQUAD
Dim waterLBuffer() As Long
Dim glassDc(2) As Long
Dim glassBmp(2) As Long
Dim glassDib(2) As Long
Dim glassBBuffer1() As RGBQUAD
Dim glassLBuffer1() As Long
Dim glassBBuffer2() As RGBQUAD
Dim glassLBuffer2() As Long
Dim glassBBuffer3() As RGBQUAD
Dim glassLBuffer3() As Long
Dim cGLayer As Long

Dim wallS() As Long
Dim floorS() As Long
Dim glassColor As Long

Dim wa2 As Single, wa1 As Single

Dim showMap As Boolean

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
        Case vbKeyShift
            speed = 1
    End Select
End Sub

Private Sub Form_Resize()
    viewport.Left = (Me.ScaleWidth - viewport.Width) / 2
    viewport.Top = (Me.ScaleHeight - viewport.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timeToEnd = True
End Sub

Private Sub Timer1_Timer()
    Dim i As Long
    
    Timer1.Enabled = False
    frmLoad.picMaze.Refresh
    mapSize.X = frmLoad.picMaze.ScaleWidth
    mapSize.Y = frmLoad.picMaze.ScaleHeight
    
    ' Map transparent masking
    
    maskBmp = CreateBitmap(mapSize.X, mapSize.Y, 1, 1, ByVal 0)
    mapBmp = CreateCompatibleBitmap(Me.hDC, mapSize.X, mapSize.Y)
    
    mapDc = CreateCompatibleDC(Me.hDC)
    maskDc = CreateCompatibleDC(Me.hDC)
    
    SelectObject mapDc, mapBmp
    SelectObject maskDc, maskBmp
    SetBkColor mapDc, vbWhite
    
    BitBlt mapDc, 0, 0, mapSize.X, mapSize.Y, frmLoad.picMaze.hDC, 0, 0, vbSrcCopy
    
        ' get dibits
        
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
    
    BitBlt maskDc, 0, 0, mapSize.X, mapSize.Y, mapDc, 0, 0, vbSrcCopy
    'BitBlt mapDc, 0, 0, mapSize.X, mapSize.Y, maskDc, 0, 0, vbSrcInvert
    
    ' end masking
    
    
    ' create dib section for backbuffer
    
    bufferDC = CreateCompatibleDC(Me.hDC)
    
    Dim bbi As BITMAPINFO
    With bbi.bmiHeader
        .biSize = LenB(bbi)
        .biWidth = 320
        .biHeight = -240
        .biPlanes = 1
        .biBitCount = 32
        .biSizeImage = ((((.biWidth * 32) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    
    buffBmp = CreateDIBSection(Me.hDC, bbi, 0, bDib, 0, 0)
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
    ' this will gain direct access to color bits,
    ' unlike the map above where they are only copied.
    CopyMemory ByVal VarPtrArray(bBuffer()), VarPtr(sa), 4  'byte bgr format
    CopyMemory ByVal VarPtrArray(lBuffer()), VarPtr(sa), 4  'long bgr format
    ' end buffer dib creation
    
    
    ' create dib section for water buffer
    waterDc = CreateCompatibleDC(Me.hDC)
    waterBmp = CreateDIBSection(Me.hDC, bbi, 0, waterDib, 0, 0)
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
    
    ' create dib section for glass window buffer
    Dim gsa(2) As SAFEARRAY2D
    
    For i = 0 To 2
        glassDc(i) = CreateCompatibleDC(Me.hDC)
        glassBmp(i) = CreateDIBSection(Me.hDC, bbi, 0, glassDib(i), 0, 0)
        SelectObject glassDc(i), glassBmp(i)
        With gsa(i)
            .cbElements = 4
            .cDims = 2
            .Bounds(0).lLbound = 0
            .Bounds(0).cElements = Abs(bbi.bmiHeader.biHeight)
            .Bounds(1).lLbound = 0
            .Bounds(1).cElements = bbi.bmiHeader.biWidth
            .pvData = glassDib(i)
        End With
        Select Case i
            Case 0
                CopyMemory ByVal VarPtrArray(glassBBuffer1()), VarPtr(gsa(i)), 4  'byte bgr format
                CopyMemory ByVal VarPtrArray(glassLBuffer1()), VarPtr(gsa(i)), 4  'byte bgr format
            Case 1
                CopyMemory ByVal VarPtrArray(glassBBuffer2()), VarPtr(gsa(i)), 4  'byte bgr format
                CopyMemory ByVal VarPtrArray(glassLBuffer2()), VarPtr(gsa(i)), 4  'byte bgr format
            Case 2
                CopyMemory ByVal VarPtrArray(glassBBuffer3()), VarPtr(gsa(i)), 4  'byte bgr format
                CopyMemory ByVal VarPtrArray(glassLBuffer3()), VarPtr(gsa(i)), 4  'byte bgr format
        End Select
    Next
    ' end dib creation
    
    
    hFov = 60    ' horizontal field of view (degrees)
    vFov = 60    ' same as horizontal for correct proportions
    fog = 120    ' visible distance (pixels)
    fsc = 80     ' scale for 3d -> 2d coordinate conversion, greater = taller doors/walls
    pHeight = 60 ' player height (pixels)
    
    speed = 1   ' in pixels
    
    'create fading color table for walls and carpet (enhances depth)
    ReDim wallS(fog)
    ReDim floorS(fog)
    ReDim glassS(fog)
    For i = 0 To fog
        wallS(i) = RGB((fog - i) / fog * 200, (fog - i) / fog * 240, (fog - i) / fog * 240)
        floorS(i) = RGB((fog - i) / fog * 255, (fog - i) / fog * 10, (fog - i) / fog * 25)
    Next
    glassColor = RGB(50, 0, 0)
    
    'buffer for background (floor and ceiling/sky)
    'this is basically a still image of two different gradients
    bgDc = CreateCompatibleDC(Me.hDC)
    bgBmp = CreateCompatibleBitmap(Me.hDC, 320, 240)
    SelectObject bgDc, bgBmp
    
    'draw the floor and the ceiling (enhances depth)
    Dim Y As Long, z As Single
    For z = 1 To fog Step 0.05
        Y = fsc * (pHeight - 90) / z + 120
        SetPixel bgDc, 0, Y, RGB((fog - z) / fog * 200, (fog - z) / fog * 200, (fog - z) / fog * 180)
        Y = fsc * pHeight / z + 120
        SetPixel bgDc, 0, Y, RGB((fog - z) / fog * 200, (fog - z) / fog * 10, (fog - z) / fog * 25)
    Next
    StretchBlt bgDc, 1, 0, 319, 241, bgDc, 0, 1, 1, 240, vbSrcCopy
    
    'center mouse
    mc.X = Screen.Width / Screen.TwipsPerPixelX / 2
    mc.Y = Screen.Height / Screen.TwipsPerPixelY / 2
    SetCursorPos mc.X, mc.Y
    ShowCursor 0   'hide cursor
    
    Dim ot As Long, fps As Long, oFps As Long
    ot = GetTickCount
    
    Do While Not timeToEnd
        'If timeToEnd Then Exit Sub
        DoEvents
        'Do While GetTickCount - ot < 2
        '    If timeToEnd Then Exit Sub
        '    DoEvents
        '    ot = GetTickCount
        'Loop
        RenderScene
        Movement
        '** Implement an auto frame skip/limit algorithm **
        '** Idea: count milliseconds for last frame
        '**   - use a loop to limit if too fast
        '**   - skip the next render if too slow
        If showFPS Then
            fps = fps + 1
            If GetTickCount - ot >= 500 Then
                oFps = fps * 2
                fps = 0
                ot = GetTickCount
            End If
            viewport.CurrentX = viewport.ScaleWidth - 70
            viewport.CurrentY = viewport.ScaleHeight - 20
            viewport.Print "FPS: " & oFps
        End If
    Loop
    
    'clean up
    CopyMemory ByVal VarPtrArray(bBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(lBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(waterBBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(waterLBuffer()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassBBuffer1()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassLBuffer1()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassBBuffer2()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassLBuffer2()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassBBuffer3()), 0&, 4
    CopyMemory ByVal VarPtrArray(glassLBuffer3()), 0&, 4

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
    ShowCursor 1
    frmLoad.Visible = True
    
    Unload Me
End Sub

Private Sub Movement()
    'update mouse rotation
    GetCursorPos cm
    angle = angle + (cm.X - mc.X) / 4
    SetCursorPos mc.X, mc.Y
    
    'detect keys and adjust movement direction
    Dim c As Long, a As Single
    
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
    
    Dim a1 As Single
    a1 = a * RADIANS
    
    'move then handle collision
    On Error Resume Next
    If key(vbKeyW) Or key(vbKeyD) Or key(vbKeyS) Or key(vbKeyA) Then
        'move along the X axis
        pos.X = pos.X + Cos(a1) * speed
        c = mapBits(pos.X, pos.Y)
        If (c <> mNone) And (c <> mDoorWay) And (c <> mCarpet) And (c <> mWater) Then
            'collision, move back
            pos.X = pos.X - Cos(a1) * speed
        End If
        'move along the Y axis
        pos.Y = pos.Y + Sin(a1) * speed
        c = mapBits(pos.X, pos.Y)
        If (c <> mNone) And (c <> mDoorWay) And (c <> mCarpet) And (c <> mWater) Then
            'collision, move back
            pos.Y = pos.Y - Sin(a1) * speed
        End If
    End If
End Sub

Private Sub RenderScene()
    'Clear buffers
    BitBlt bufferDC, 0, 0, 320, 240, bgDc, 0, 0, vbSrcCopy
    ZeroMemory ByVal waterDib, 307200  '320 * 240 * 4 (width * height * 4 bytes)
    ZeroMemory ByVal glassDib(0), 307200
    ZeroMemory ByVal glassDib(1), 307200
    ZeroMemory ByVal glassDib(2), 307200
    
    Dim i As Long, a As Single, a2 As Single, va As Single
    
    'wa2 = 0
    'scan through the screen (width of 320)
    For i = 0 To 319
        a = ((i + 1) / 320 * hFov) - hFov / 2
        va = ((i + 1) / 320 * vFov) - vFov / 2
        cGLayer = 0
        wa1 = 0
        wa2 = wa2 + 1 + Rnd / 10
        zDetail = 1 'ray tracing step, the smaller the smoother but slower
        a2 = (a + angle) * RADIANS
        'scan along the rays outwards
        '                    > this cancels out the horizontal bulging caused by using only 1 field of view
        '                    |---------------------|
        rayRecur i, 1, fsc * (2 - Cos(va * RADIANS)), Cos(a2), Sin(a2), 0, 240, -1, 90
    Next
    
    BlendAlpha
    
    'copy the back buffer to front buffer (in this case, the viewport)
    'BitBlt viewport.hdc, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, vbSrcCopy
    StretchBlt viewport.hDC, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, 320, 240, vbSrcCopy
    
    'draw the transparent map and then draw player location
    '!!! transparency not working properly !!!
    If showMap Then
        'BitBlt viewport.hdc, 0, 0, mapSize.X, mapSize.Y, maskDc, 0, 0, vbSrcAnd
        'BitBlt viewport.hdc, 0, 0, mapSize.X, mapSize.Y, mapDc, 0, 0, vbSrcPaint
        BitBlt viewport.hDC, 0, 0, mapSize.X, mapSize.Y, mapDc, 0, 0, vbSrcAnd      'temporary
        
        viewport.ForeColor = vbGreen
        Pie viewport.hDC, pos.X - fog, pos.Y - fog, pos.X + fog, pos.Y + fog, pos.X + Cos((angle + hFov / 2) * RADIANS) * fog, pos.Y + Sin((angle + hFov / 2) * RADIANS) * fog, pos.X + Cos((angle - hFov / 2) * RADIANS) * fog, pos.Y + Sin((angle - hFov / 2) * RADIANS) * fog
        viewport.ForeColor = vbRed
        Ellipse viewport.hDC, pos.X - 1, pos.Y - 1, pos.X + 1, pos.Y + 1
    End If
End Sub

'** Possibly optimizable **
Private Sub rayRecur(ByVal i As Long, ByVal rStart As Single, ByVal vScl As Single, ByVal xp As Single, ByVal yp As Single, ByVal wTop As Long, ByVal wBottom As Long, ByVal pType As Long, ByVal dist As Long)
    Dim j As Single, c As Long, j2 As Long, w1 As Single, X As Single, Y As Single
    Dim y1 As Long, y2 As Long, c2 As Long, cLayer As Long
    
    'start or continue tracing ray where parent function left off
    cLayer = cGLayer
    j = rStart
    Do While j < fog
        j = j + zDetail
        'zDetail = zDetail + 0.1
        wa1 = wa1 + 1
            'calculate current point of ray
            X = pos.X + xp * j
            Y = pos.Y + yp * j
            On Error Resume Next    '\  * Uncomment this if debugging *
            c = -1                  ' > I think this is faster than manually checking for map bounds
            c = mapBits(X, Y)       '/
            If c = -1 Then
                'fill in the skipped gap of the water wave,
                'this prevents broken water animation
                wa1 = wa1 + fog - j
                'no further raytracing is needed, exit and move to adjacent ray
                Exit Sub
            ElseIf c = mWall Then
                DrawVLine i, (vScl * (pHeight - 90) / j) + 120, (vScl * pHeight / j) + 120, wallS(j), wTop, wBottom
                'fill in the skipped gap of the water wave,
                'this prevents broken water animation
                wa1 = wa1 + fog - j
                'no further raytracing is needed, exit and move to adjacent ray
                Exit Sub
            '** Calculate a window region to constrain the drawing behind this surface, **
            '** these bounds will be passed into the next recursive call.               **
            ElseIf c = mDoorWay Then
                y1 = (vScl * (pHeight - 80) / j) + 120
                y2 = (vScl * (pHeight) / j) + 120
                Exit Do
            ElseIf c = mLedge Then
                y1 = (vScl * (pHeight - 90) / j) + 120
                y2 = (vScl * (pHeight - 20) / j) + 120
                Exit Do
            ElseIf c = mWindow Then
                y1 = (vScl * (pHeight - 70) / j) + 120
                y2 = (vScl * (pHeight - 35) / j) + 120
                If ((j - dist) > 3) Then
                    cGLayer = cGLayer + 1
                End If
                dist = j
                Exit Do
            ElseIf c = mCarpet Then
                y1 = (vScl * (pHeight - 90) / j) + 120
                y2 = (vScl * (pHeight - 2) / j) + 120
                Exit Do
            ElseIf (c <> mNone) And (c <> -1) Then
                y1 = wTop
                y2 = wBottom
                'A surface with an opening ("window") has been detected, exit this loop and enter
                'a new loop in the recursive call to find surfaces behind this one.
                Exit Do
            End If
    Loop
    
    If y1 < wTop Then
        y1 = wTop
    ElseIf y1 < 0 Then
        y1 = 0
    End If
    If y2 > wBottom Then
        y2 = wBottom
    ElseIf y2 >= 320 Then
        y2 = 239
    End If
    'call this function again
    If j < fog Then rayRecur i, j, vScl, xp, yp, y1, y2, c, dist
    
    '** Check and draw only the visible part seen through the  **
    '** window region of the surface in front of this surface. **
    '** The bounds for this calculation will be passed in from **
    '** the parent function, which is the surface in front of  **
    '** this one.                                              **
    '** This can also be used to draw water on a separate      **
    '** buffer, then alpha blended to backbuffer in the end.   **
    '** Similarly, glass windows and semi-transparent surfaces **
    '** can be processed this way on a separate buffer.        **
    
    Select Case c
        Case mDoorWay
            DrawVLine i, (vScl * (pHeight - 90) / j) + 120, (vScl * (pHeight - 80) / j) + 120, wallS(j), wTop, wBottom
        Case mLedge
            DrawVLine i, (vScl * (pHeight - 20) / j) + 120, (vScl * pHeight / j) + 120, wallS(j), wTop, wBottom
        Case mWindow
            DrawVLine i, (vScl * (pHeight - 35) / j) + 120, (vScl * pHeight / j) + 120, wallS(j), wTop, wBottom
            DrawVLine i, (vScl * (pHeight - 90) / j) + 120, (vScl * (pHeight - 70) / j) + 120, wallS(j), wTop, wBottom
            DrawGLine i, y1, y2, glassColor, y1, y2, cLayer
        Case mCarpet
            DrawVLine i, (vScl * (pHeight - 2) / j) + 120, (vScl * pHeight / j) + 120, floorS(j), wTop, wBottom
        Case mWater
            w1 = Cos((wa1 + wa2) * RADIANS)
            j2 = j + w1 * 20
            If j2 < 0 Then
                j2 = 0
            ElseIf j2 > fog Then
                j2 = fog
            End If
            DrawWLine i, (vScl * (pHeight - 16 + w1 / 3) / j) + 120, (vScl * pHeight / j) + 120, floorS(j2), wTop, wBottom
    End Select
End Sub

'Note: y1 must be less than y2 because the line will be drawn top to bottom
Private Sub DrawVLine(ByVal X As Long, ByVal y1 As Long, ByVal y2 As Long, ByVal color As Long, ByVal wTop As Long, ByVal wBottom As Long)
    If y1 >= wBottom Then
        Exit Sub
    ElseIf y1 < wTop Then
        y1 = wTop
    End If
    If y2 <= wTop Then
        Exit Sub
    ElseIf y2 > wBottom Then
        y2 = wBottom
    End If
    For y1 = y1 To y2 - 1
        lBuffer(X, y1) = color
    Next
End Sub
Private Sub DrawWLine(ByVal X As Long, ByVal y1 As Long, ByVal y2 As Long, ByVal color As Long, ByVal wTop As Long, ByVal wBottom As Long)
    If y1 >= wBottom Then
        Exit Sub
    ElseIf y1 < wTop Then
        y1 = wTop
    End If
    If y2 <= wTop Then
        Exit Sub
    ElseIf y2 > wBottom Then
        y2 = wBottom
    End If
    For y1 = y1 To y2 - 1
        waterLBuffer(X, y1) = color
    Next
End Sub
Private Sub DrawGLine(ByVal X As Long, ByVal y1 As Long, ByVal y2 As Long, ByVal color As Long, ByVal wTop As Long, ByVal wBottom As Long, ByVal glassLayer As Long)
    If y1 >= wBottom Then
        Exit Sub
    ElseIf y1 < wTop Then
        y1 = wTop
    End If
    If y2 <= wTop Then
        Exit Sub
    ElseIf y2 > wBottom Then
        y2 = wBottom
    End If
    Select Case glassLayer
        Case 0
            For y1 = y1 To y2 - 1
                glassLBuffer1(X, y1) = color
            Next
        Case 1
            For y1 = y1 To y2 - 1
                glassLBuffer2(X, y1) = color
            Next
        Case 2
            For y1 = y1 To y2 - 1
                glassLBuffer3(X, y1) = color
            Next
    End Select
End Sub

'mostly will not be used
Private Sub DrawVLineAlpha(ByVal X As Long, ByVal y1 As Long, ByVal y2 As Long, color As RGBQUAD, ByVal Opacity As Single)
    Dim t As Long
    If y1 < 0 Then y1 = 0
    If y2 >= 240 Then y2 = 239
    For y1 = y1 To y2
        t = (CSng(bBuffer(X, y1).rgbBlue) - color.rgbBlue) * Opacity + color.rgbBlue
        bBuffer(X, y1).rgbBlue = t
        t = (CSng(bBuffer(X, y1).rgbGreen) - color.rgbGreen) * Opacity + color.rgbGreen
        bBuffer(X, y1).rgbGreen = t
        t = (CSng(bBuffer(X, y1).rgbRed) - color.rgbRed) * Opacity + color.rgbRed
        bBuffer(X, y1).rgbRed = t
    Next
End Sub

Private Sub DrawPixel(ByVal X As Long, ByVal Y As Long, ByVal color As Long)
    If Y >= 0 And Y < 230 Then
        lBuffer(X, Y) = color
    End If
End Sub

Private Sub BlendAlpha()
    Dim X As Long, Y As Long, c As Single
    For Y = 0 To 239
        For X = 0 To 319
            If waterLBuffer(X, Y) <> vbBlack Then
                c = (CLng(waterBBuffer(X, Y).rgbBlue) - bBuffer(X, Y).rgbBlue) * 0.75 + bBuffer(X, Y).rgbBlue
                bBuffer(X, Y).rgbBlue = c
                c = (CLng(waterBBuffer(X, Y).rgbGreen) - bBuffer(X, Y).rgbGreen) * 0.75 + bBuffer(X, Y).rgbGreen
                bBuffer(X, Y).rgbGreen = c
                c = (CLng(waterBBuffer(X, Y).rgbRed) - bBuffer(X, Y).rgbRed) * 0.75 + bBuffer(X, Y).rgbRed
                bBuffer(X, Y).rgbRed = c
            End If
            If glassLBuffer1(X, Y) <> vbBlack Then
                c = (CLng(glassBBuffer1(X, Y).rgbBlue) - bBuffer(X, Y).rgbBlue) * 0.25 + bBuffer(X, Y).rgbBlue
                bBuffer(X, Y).rgbBlue = c
                c = (CLng(glassBBuffer1(X, Y).rgbGreen) - bBuffer(X, Y).rgbGreen) * 0.25 + bBuffer(X, Y).rgbGreen
                bBuffer(X, Y).rgbGreen = c
                c = (CLng(glassBBuffer1(X, Y).rgbRed) - bBuffer(X, Y).rgbRed) * 0.25 + bBuffer(X, Y).rgbRed
                bBuffer(X, Y).rgbRed = c
            End If
            If glassLBuffer2(X, Y) <> vbBlack Then
                c = (CLng(glassBBuffer2(X, Y).rgbBlue) - bBuffer(X, Y).rgbBlue) * 0.25 + bBuffer(X, Y).rgbBlue
                bBuffer(X, Y).rgbBlue = c
                c = (CLng(glassBBuffer2(X, Y).rgbGreen) - bBuffer(X, Y).rgbGreen) * 0.25 + bBuffer(X, Y).rgbGreen
                bBuffer(X, Y).rgbGreen = c
                c = (CLng(glassBBuffer2(X, Y).rgbRed) - bBuffer(X, Y).rgbRed) * 0.25 + bBuffer(X, Y).rgbRed
                bBuffer(X, Y).rgbRed = c
            End If
            If glassLBuffer3(X, Y) <> vbBlack Then
                c = (CLng(glassBBuffer3(X, Y).rgbBlue) - bBuffer(X, Y).rgbBlue) * 0.25 + bBuffer(X, Y).rgbBlue
                bBuffer(X, Y).rgbBlue = c
                c = (CLng(glassBBuffer3(X, Y).rgbGreen) - bBuffer(X, Y).rgbGreen) * 0.25 + bBuffer(X, Y).rgbGreen
                bBuffer(X, Y).rgbGreen = c
                c = (CLng(glassBBuffer3(X, Y).rgbRed) - bBuffer(X, Y).rgbRed) * 0.25 + bBuffer(X, Y).rgbRed
                bBuffer(X, Y).rgbRed = c
            End If
        Next
    Next
End Sub

Private Sub viewport_Resize()
    viewport.Left = (Me.ScaleWidth - viewport.Width) / 2
    viewport.Top = (Me.ScaleHeight - viewport.Height) / 2
End Sub
