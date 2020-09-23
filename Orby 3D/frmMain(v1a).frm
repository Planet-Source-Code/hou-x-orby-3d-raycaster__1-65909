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

' Version 1 Alpha

' About this version:
' - uses GDI MoveToEx and LineTo to draw the surfaces.
' - uses GetPixel to get map data
' - does not support alpha blending for water
' - does not support window glass
' - does not fix bulging (fish eye) side effect
' - supports semi-working map display

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

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
Dim bufferDC As Long
Dim buffBmp As Long
Dim fov As Long
Dim fsc As Single
Dim fog As Long
Dim pens() As Long
Dim fPens() As Long
Dim key(255) As Boolean
Dim zDetail As Single

Dim speed As Single

Dim bgDc As Long
Dim bgBmp As Long
Dim mapDc As Long
Dim mapBmp As Long
Dim maskDc As Long
Dim maskBmp As Long
Dim mapSize As POINTAPI
Dim mapBits() As Long

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
    Timer1.Enabled = False
    frmLoad.picMaze.Refresh
    mapSize.X = frmLoad.picMaze.ScaleWidth
    mapSize.Y = frmLoad.picMaze.ScaleHeight
    
    ' Map transparent masking
    
    maskBmp = CreateBitmap(mapSize.X, mapSize.Y, 1, 1, ByVal 0)
    mapBmp = CreateCompatibleBitmap(Me.hdc, mapSize.X, mapSize.Y)
    
    mapDc = CreateCompatibleDC(Me.hdc)
    maskDc = CreateCompatibleDC(Me.hdc)
    
    SelectObject mapDc, mapBmp
    SelectObject maskDc, maskBmp
    SetBkColor mapDc, vbWhite
    
    BitBlt mapDc, 0, 0, mapSize.X, mapSize.Y, frmLoad.picMaze.hdc, 0, 0, vbSrcCopy
    
        ' get dibits
        
        Dim bi As BITMAPINFO
        bi.bmiHeader.biSize = LenB(bi)
        bi.bmiHeader.biWidth = mapSize.X
        bi.bmiHeader.biHeight = -mapSize.Y
        bi.bmiHeader.biPlanes = 1
        bi.bmiHeader.biBitCount = 32
        bi.bmiHeader.biSizeImage = ((((mapSize.X * 32) + 31) \ 32) * 4) * mapSize.Y
        
        ReDim mapBits(mapSize.X - 1, mapSize.Y - 1)
        GetDIBits mapDc, mapBmp, 0, mapSize.Y, mapBits(0, 0), bi, 0
        ' end get bits
    
    BitBlt maskDc, 0, 0, mapSize.X, mapSize.Y, mapDc, 0, 0, vbSrcCopy
    'BitBlt mapDc, 0, 0, mapSize.X, mapSize.Y, maskDc, 0, 0, vbSrcInvert
    
    ' end masking
    
    bufferDC = CreateCompatibleDC(Me.hdc)
    buffBmp = CreateCompatibleBitmap(Me.hdc, 320, 240)
    SelectObject bufferDC, buffBmp
    
    fov = 45    ' horizontal field of view (degrees)
    fog = 120   ' visible distance (pixels)
    fsc = 60    ' scale for 3d -> 2d coordinate conversion
    
    speed = 1   ' in pixels
    
    'create fading pens for walls and carpet (enhances depth)
    Dim i As Long
    ReDim pens(fog)
    ReDim fPens(fog)
    For i = 0 To fog
        pens(i) = CreatePen(PS_SOLID, 1, RGB((fog - i) / fog * 240, (fog - i) / fog * 240, (fog - i) / fog * 200))
        fPens(i) = CreatePen(PS_SOLID, 1, RGB((fog - i) / fog * 25, (fog - i) / fog * 10, (fog - i) / fog * 255))
    Next
    
    'buffer for background (floor and ceiling/sky)
    'this is basically a still image of two different gradients
    bgDc = CreateCompatibleDC(Me.hdc)
    bgBmp = CreateCompatibleBitmap(Me.hdc, 320, 240)
    SelectObject bgDc, bgBmp
    
    'draw the floor and the ceiling (enhances depth)
    Dim Y As Long, z As Single
    For z = 1 To fog Step 0.1
        Y = fsc * -60 / z + 120
        SetPixel bgDc, 0, Y, RGB((fog - z) / fog * 200, (fog - z) / fog * 200, (fog - z) / fog * 180)
        Y = fsc * 30 / z + 120
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
    DeleteObject buffBmp
    DeleteDC bufferDC
    DeleteObject bgBmp
    DeleteDC bgDc
    DeleteObject mapBmp
    DeleteDC mapDc
    DeleteObject maskBmp
    DeleteDC maskDc
    For i = 0 To fog
        DeleteObject pens(i)
        DeleteObject fPens(i)
    Next
    ShowCursor 1
    frmLoad.Visible = True
    
    Unload Me
End Sub

Private Sub Movement()
    'update mouse rotation
    GetCursorPos cm
    angle = angle + (cm.X - mc.X) / 4
    SetCursorPos mc.X, mc.Y
    
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
    BitBlt bufferDC, 0, 0, 320, 240, bgDc, 0, 0, vbSrcCopy
    
    Dim i As Long, a As Single, a2 As Single
    
    'wa2 = 0
    'scan across the actual screen (width of 320)
    For i = 0 To 319
        a = (i / 240 * fov) - fov / 2
        wa1 = 0
        wa2 = wa2 + 1 + Rnd / 10
        zDetail = 1
        a2 = (a + angle) * RADIANS
        'scan along the rays outwards
        rayRecur i, 1, Cos(a2), Sin(a2)
    Next
    
    'copy the back buffer to front buffer (in this case, the viewport)
    'BitBlt viewport.hdc, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, vbSrcCopy
    StretchBlt viewport.hdc, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, 320, 240, vbSrcCopy
    
    'draw the transparent map and then draw player location
    '!!! transparency not working properly !!!
    If showMap Then
        'BitBlt viewport.hdc, 0, 0, mapSize.X, mapSize.Y, maskDc, 0, 0, vbSrcAnd
        'BitBlt viewport.hdc, 0, 0, mapSize.X, mapSize.Y, mapDc, 0, 0, vbSrcPaint
        BitBlt viewport.hdc, 0, 0, mapSize.X, mapSize.Y, mapDc, 0, 0, vbSrcAnd      'temporary
        
        viewport.ForeColor = vbGreen
        Pie viewport.hdc, pos.X - fog, pos.Y - fog, pos.X + fog, pos.Y + fog, pos.X + Cos((angle + fov / 2) * RADIANS) * fog, pos.Y + Sin((angle + fov / 2) * RADIANS) * fog, pos.X + Cos((angle - fov / 2) * RADIANS) * fog, pos.Y + Sin((angle - fov / 2) * RADIANS) * fog
        viewport.ForeColor = vbRed
        Ellipse viewport.hdc, pos.X - 1, pos.Y - 1, pos.X + 1, pos.Y + 1
    End If
End Sub

'** Possibly optimizable **
Private Sub rayRecur(ByVal i As Long, ByVal rStart As Single, ByVal xp As Single, ByVal yp As Single)
    Dim j As Single, c As Long, j2 As Long, w1 As Single, X As Single, Y As Single
    
    'start or continue tracing ray where parent function left off
    j = rStart
    Do While j < fog
        j = j + zDetail
        If j > fog Then j = fog
        'zDetail = zDetail + 0.1
        wa1 = wa1 + 1
            'calculate current point of ray
            X = pos.X + xp * j
            Y = pos.Y + yp * j
            On Error Resume Next    '\
            c = -1                  ' > I think this is faster than manually checking for map bounds
            c = mapBits(X, Y)       '/
            If c = -1 Then
                'fill in the skipped gap of the water wave,
                'this prevents broken water animation
                wa1 = wa1 + fog - j
                'no further raytracing is needed, exit and move to adjacent ray
                Exit Sub
            ElseIf c = mWall Then
                MoveToEx bufferDC, i, (fsc * -60 / j) + 120, 0
                SelectObject bufferDC, pens(j)
                LineTo bufferDC, i, (fsc * 30 / j) + 120
                'fill in the skipped gap of the water wave,
                'this prevents broken water animation
                wa1 = wa1 + fog - j
                'no further raytracing is needed, exit and move to adjacent ray
                Exit Sub
            ElseIf (c <> mNone) And (c <> -1) Then
                '** Calculate a window region to constrain the drawing behind this surface, **
                '** these bounds will be passed into the next recursive call.               **
            
                'A surface with an opening ("window") has been detected, exit this loop and enter
                'a new loop in the recursive call to find surfaces behind this one.
                Exit Do
            End If
    Loop
    
    'call this function again
    If j < fog Then rayRecur i, j, xp, yp
    
    '** Check and draw only the visible part seen through the  **
    '** window region of the surface in front of this surface. **
    '** The bounds for this calculation will be passed in from **
    '** the parent function, which is the surface in front of  **
    '** this one.
    
    '** Instead of using LoneTo, maybe create a dib section    **
    '** and use loops or possibly memory functions, such as    **
    '** a fill memory function if it exists.
    
    Select Case c
        Case mDoorWay
            MoveToEx bufferDC, i, (fsc * -60 / j) + 120, 0
            SelectObject bufferDC, pens(j)
            LineTo bufferDC, i, (fsc * -40 / j) + 120
        Case mLedge
            MoveToEx bufferDC, i, (fsc * 10 / j) + 120, 0
            SelectObject bufferDC, pens(j)
            LineTo bufferDC, i, (fsc * 30 / j) + 120
        Case mWindow
            SelectObject bufferDC, pens(j)
            MoveToEx bufferDC, i, (fsc * 5 / j) + 120, 0
            LineTo bufferDC, i, (fsc * 30 / j) + 120
            MoveToEx bufferDC, i, (fsc * -60 / j) + 120, 0
            LineTo bufferDC, i, (fsc * -30 / j) + 120
        Case mCarpet
            MoveToEx bufferDC, i, (fsc * 28 / j) + 120, 0
            SelectObject bufferDC, fPens(j)
            LineTo bufferDC, i, (fsc * 30 / j) + 120
        Case mWater
            w1 = Cos((wa1 + wa2) * RADIANS)
            MoveToEx bufferDC, i, (fsc * (14 + w1 / 3) / j) + 120, 0
            j2 = j + w1 * 20
            If j2 < 0 Then
                j2 = 0
            ElseIf j2 > fog Then
                j2 = fog
            End If
            SelectObject bufferDC, fPens(j2)
            LineTo bufferDC, i, (fsc * 30 / j) + 120
    End Select
End Sub

Private Sub viewport_Resize()
    viewport.Left = (Me.ScaleWidth - viewport.Width) / 2
    viewport.Top = (Me.ScaleHeight - viewport.Height) / 2
End Sub
