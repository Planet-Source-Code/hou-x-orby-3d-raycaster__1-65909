VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ray Caster Maze"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   621
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   907
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   7680
   End
   Begin VB.PictureBox viewport 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   120
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
   Begin VB.PictureBox picMaze 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   2340
      Left            =   9840
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   156
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   1
      Top             =   1560
      Width           =   2940
   End
   Begin VB.Label Label2 
      Caption         =   "Controls:  W,D,S,A, and Mouse, Esc to quit."
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "To start, click on the map to position the starting point."
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long

Private Const BLACKNESS = &H42
Private Const PS_SOLID = 0

Private Type POINTAPI
        X As Long
        y As Long
End Type
Private Type POINTAPIF
        X As Single
        y As Single
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

Dim pos As POINTAPIF
Dim angle As Single
Dim bufferDC As Long
Dim buffBmp As Long
Dim fov As Long
Dim fog As Long
Dim mazeDC As Long
Dim mazeBmp As Long
Dim pens() As Long
Dim key(255) As Boolean

Dim bgDc As Long
Dim bgBmp As Long

Dim cm As POINTAPI
Dim mc As POINTAPI

Dim timeToEnd As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    key(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    key(KeyCode) = False
End Sub

Private Sub Form_Load()
    bufferDC = CreateCompatibleDC(Me.hdc)
    buffBmp = CreateCompatibleBitmap(Me.hdc, 320, 240)
    SelectObject bufferDC, buffBmp
    
    pos.X = 50
    pos.y = 50
    fov = 60
    fog = 60
    
    Dim i As Long
    ReDim pens(fog)
    For i = 0 To fog
        pens(i) = CreatePen(PS_SOLID, 1, RGB((fog - i) / fog * 240, (fog - i) / fog * 240, (fog - i) / fog * 200))
    Next
    
    picMaze.Refresh
    mazeDC = CreateCompatibleDC(Me.hdc)
    mazeBmp = CreateCompatibleBitmap(Me.hdc, picMaze.ScaleWidth, picMaze.ScaleHeight)
    SelectObject mazeDC, mazeBmp
    BitBlt mazeDC, 0, 0, picMaze.ScaleWidth, picMaze.ScaleHeight, picMaze.hdc, 0, 0, vbSrcCopy
    
    bgDc = CreateCompatibleDC(Me.hdc)
    bgBmp = CreateCompatibleBitmap(Me.hdc, 320, 240)
    SelectObject bgDc, bgBmp
    
    Dim y As Long, z As Single
    For z = 1 To fog Step 0.1
        y = 60 * -60 / z + 120
        SetPixel bgDc, 0, y, RGB((fog - z) / fog * 200, (fog - z) / fog * 200, (fog - z) / fog * 180)
        y = 60 * 30 / z + 120
        SetPixel bgDc, 0, y, RGB((fog - z) / fog * 200, (fog - z) / fog * 10, (fog - z) / fog * 25)
    Next
    StretchBlt bgDc, 1, 0, 319, 240, bgDc, 0, 1, 1, 240, vbSrcCopy
    
    'picMaze.AutoRedraw = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    timeToEnd = True
    DeleteObject buffBmp
    DeleteDC bufferDC
    DeleteObject mazeBmp
    DeleteDC mazeDC
    DeleteObject bgBmp
    DeleteDC bgDc
    Dim i As Long
    For i = 0 To fog
        DeleteObject pens(i)
    Next
    ShowCursor 1
End Sub

Private Sub picMaze_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    pos.X = X
    pos.y = y
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    
    mc.X = Me.Left / Screen.TwipsPerPixelX + viewport.Left + viewport.ScaleWidth / 2
    mc.y = Me.Top / Screen.TwipsPerPixelY + viewport.Top + viewport.ScaleHeight / 2
    SetCursorPos mc.X, mc.y
    ShowCursor 0
    
    Dim ot As Long
    ot = GetTickCount
    
    Do While Not timeToEnd
        If timeToEnd Then Exit Sub
        DoEvents
        'Do While GetTickCount - ot < 2
        '    If timeToEnd Then Exit Sub
        '    DoEvents
        '    ot = GetTickCount
        'Loop
        RenderScene
        Movement
        'ot = GetTickCount
    Loop
    
    Unload Me
End Sub

Private Sub Movement()
    If key(vbKeyEscape) Then
        Unload Me
    End If
    
    GetCursorPos cm
    angle = angle + (cm.X - mc.X) / 4
    SetCursorPos mc.X, mc.y
    
    Dim c As Long, a As Single
    
    If False Then
        If key(vbKeyUp) And Not key(vbKeyRight) And Not key(vbKeyDown) And key(vbKeyLeft) Then
            a = angle - 45
        ElseIf Not key(vbKeyUp) And Not key(vbKeyRight) And key(vbKeyDown) And key(vbKeyLeft) Then
            a = angle - 90 - 45
        ElseIf key(vbKeyUp) And key(vbKeyRight) And Not key(vbKeyDown) And Not key(vbKeyLeft) Then
            a = angle + 45
        ElseIf Not key(vbKeyUp) And key(vbKeyRight) And key(vbKeyDown) And Not key(vbKeyLeft) Then
            a = angle + 90 + 45
        ElseIf Not key(vbKeyUp) And key(vbKeyRight) And Not key(vbKeyDown) And Not key(vbKeyLeft) Then
            a = angle + 90
        ElseIf key(vbKeyUp) And Not key(vbKeyRight) Or Not key(vbKeyDown) And Not key(vbKeyLeft) Then
            a = angle
        ElseIf Not key(vbKeyUp) And Not key(vbKeyRight) And key(vbKeyDown) And Not key(vbKeyLeft) Then
            a = angle - 180
        ElseIf Not key(vbKeyUp) And Not key(vbKeyRight) And Not key(vbKeyDown) And key(vbKeyLeft) Then
            a = angle - 90
        End If
        
        If key(vbKeyUp) Or key(vbKeyRight) Or key(vbKeyDown) Or key(vbKeyLeft) Then
            pos.X = pos.X + Cos(a * RADIANS)
            c = GetPixel(mazeDC, pos.X, pos.y)
            If c <> vbWhite Then
                pos.X = pos.X - Cos(a * RADIANS)
            End If
            pos.y = pos.y + Sin(a * RADIANS)
            c = GetPixel(mazeDC, pos.X, pos.y)
            If c <> vbWhite Then
                pos.y = pos.y - Sin(a * RADIANS)
            End If
        End If
    End If
    
    If key(vbKeyW) And Not key(vbKeyD) And Not key(vbKeyS) And key(vbKeyA) Then
        a = angle - 45
    ElseIf Not key(vbKeyW) And Not key(vbKeyD) And key(vbKeyS) And key(vbKeyA) Then
        a = angle - 90 - 45
    ElseIf key(vbKeyW) And key(vbKeyD) And Not key(vbKeyS) And Not key(vbKeyA) Then
        a = angle + 45
    ElseIf Not key(vbKeyW) And key(vbKeyD) And key(vbKeyS) And Not key(vbKeyA) Then
        a = angle + 90 + 45
    ElseIf Not key(vbKeyW) And key(vbKeyD) And Not key(vbKeyS) And Not key(vbKeyA) Then
        a = angle + 90
    ElseIf key(vbKeyW) And Not key(vbKeyD) Or Not key(vbKeyS) And Not key(vbKeyA) Then
        a = angle
    ElseIf Not key(vbKeyW) And Not key(vbKeyD) And key(vbKeyS) And Not key(vbKeyA) Then
        a = angle - 180
    ElseIf Not key(vbKeyW) And Not key(vbKeyD) And Not key(vbKeyS) And key(vbKeyA) Then
        a = angle - 90
    End If
    
    If key(vbKeyW) Or key(vbKeyD) Or key(vbKeyS) Or key(vbKeyA) Then
        pos.X = pos.X + Cos(a * RADIANS)
        c = GetPixel(mazeDC, pos.X, pos.y)
        If c <> vbWhite Then
            pos.X = pos.X - Cos(a * RADIANS)
        End If
        pos.y = pos.y + Sin(a * RADIANS)
        c = GetPixel(mazeDC, pos.X, pos.y)
        If c <> vbWhite Then
            pos.y = pos.y - Sin(a * RADIANS)
        End If
    End If
End Sub

Private Sub RenderScene()
    BitBlt bufferDC, 0, 0, 320, 240, bgDc, 0, 0, vbSrcCopy
    
    Dim i As Long, j As Single, a As Single, c As Long, X As Single, y As Single
    
    picMaze.Cls
    
    picMaze.ForeColor = vbGreen
    Pie picMaze.hdc, pos.X - fog, pos.y - fog, pos.X + fog, pos.y + fog, pos.X + Cos((angle + fov / 2) * RADIANS) * fog, pos.y + Sin((angle + fov / 2) * RADIANS) * fog, pos.X + Cos((angle - fov / 2) * RADIANS) * fog, pos.y + Sin((angle - fov / 2) * RADIANS) * fog
    picMaze.ForeColor = vbRed
    Ellipse picMaze.hdc, pos.X - 1, pos.y - 1, pos.X + 1, pos.y + 1
    
    picMaze.Refresh
    
    For i = 0 To 320
        a = (i / 240 * fov) - (fov / 2)
        For j = 1 To fog
            X = pos.X + Cos((a + angle) * RADIANS) * j
            y = pos.y + Sin((a + angle) * RADIANS) * j
            c = GetPixel(mazeDC, X, y)
            Select Case c
                Case vbBlack
                    MoveToEx bufferDC, i, (60 * -60 / j) + 120, 0
                    SelectObject bufferDC, pens(j)
                    LineTo bufferDC, i, (60 * 30 / j) + 120
                    Exit For
            End Select
        Next
    Next
    
    'BitBlt viewport.hdc, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, vbSrcCopy
    StretchBlt viewport.hdc, 0, 0, viewport.ScaleWidth, viewport.ScaleHeight, bufferDC, 0, 0, 320, 240, vbSrcCopy
End Sub
