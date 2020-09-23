VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load Level - Orby 3D - Created by Hou Xiong"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRes 
      Height          =   315
      Left            =   5880
      TabIndex        =   14
      Top             =   4710
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Refresh"
      Height          =   420
      Left            =   2880
      TabIndex        =   13
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   420
      Left            =   240
      TabIndex        =   12
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   420
      Left            =   240
      TabIndex        =   11
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   420
      Left            =   1440
      TabIndex        =   10
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   420
      Left            =   1440
      TabIndex        =   8
      Top             =   5280
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2535
   End
   Begin VB.HScrollBar HScroll1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   3360
      Width           =   4335
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   4335
      Left            =   7440
      TabIndex        =   0
      Top             =   -120
      Width           =   255
   End
   Begin VB.PictureBox picCont 
      BackColor       =   &H8000000C&
      Height          =   3735
      Left            =   2880
      ScaleHeight     =   245
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   325
      TabIndex        =   2
      Top             =   480
      Width           =   4935
      Begin VB.PictureBox picMaze 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   0
         ScaleHeight     =   153
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Resolution:"
      Height          =   375
      Left            =   4920
      TabIndex        =   15
      Top             =   4770
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   8
      X2              =   536
      Y1              =   392
      Y2              =   392
   End
   Begin VB.Label Label4 
      Caption         =   "Created by Hou Xiong"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Choose a map:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "To start, click on ""Play"" or on the Map."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Controls:  W,D,S,A, and Mouse to move.  Shift to run.  Esc to return.  M to toggle map.  F to toggle FPS."
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   5280
      Width           =   4935
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEdit_Click()
    Shell "c:\windows\system32\mspaint.exe """ & File1.Path & "\" & File1.List(File1.ListIndex) & """", vbNormalFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPlay_Click()
    pos.X = 0
    pos.Y = 0
    timeToEnd = False
    ReDim LevelList(0)
    LevelList(0) = File1.List(File1.ListIndex)
    numLevels = 1
    curLevel = 1
    Me.Visible = False
    
    Dim iWidth As Long, iHeight As Long
    
    iWidth = Screen.Width / Screen.TwipsPerPixelX
    iHeight = Screen.Height / Screen.TwipsPerPixelY
    
    Dim noChange As Boolean
    Select Case cmbRes.ListIndex
        Case 1
            ChangeRes 320, 240
        Case 2
            ChangeRes 512, 384
        Case 3
            ChangeRes 640, 480
        Case 4
            ChangeRes 800, 600
        Case Else
            noChange = True
    End Select
    
    frmMain.Show vbModal, Me
    
    If Not noChange Then ChangeRes iWidth, iHeight
    
    Me.Visible = True
End Sub

Private Sub Command1_Click()
    Dim Name As String
    Name = InputBox("Enter a name:", "New Map", "new")
    If Name = "" Then Exit Sub
    If Dir(File1.Path & "\" & Name & ".bmp") <> "" Then GoTo fExists
    On Error GoTo fExists
    FileSystem.FileCopy File1.Path & "\Template.tmp", File1.Path & "\" & Name & ".bmp"
    Shell "c:\windows\system32\mspaint.exe """ & File1.Path & "\" & Name & ".bmp""", vbNormalFocus
    File1.Refresh
    Exit Sub
    
fExists:
    MsgBox "The filename already exists.", vbInformation
End Sub

Private Sub Command2_Click()
    Dim i As Long
    i = File1.ListIndex
    File1.Refresh
    
    If i >= 0 Then
        File1.ListIndex = i
        File1_Click
    End If
End Sub

Private Sub File1_Click()
    loadPic File1.List(File1.ListIndex)
    cmdEdit.Enabled = True
    cmdPlay.Enabled = True
End Sub

Private Sub File1_DblClick()
    cmdPlay_Click
End Sub

Private Sub Form_Load()
    
    If MsgBox("For a significant speed boost, compile the program and run the executable. Continue?", vbYesNo) = vbNo Then End

    HScroll1.Width = picCont.Width
    HScroll1.Left = picCont.Left
    HScroll1.Top = picCont.Top + picCont.Height - 2
    
    VScroll1.Height = picCont.Height
    VScroll1.Top = picCont.Top
    VScroll1.Left = picCont.Left + picCont.Width - 2
    
    File1.Path = App.Path & "\Maps"
    File1.FileName = "*.bmp"
    
    cmbRes.AddItem "Don't Change"
    cmbRes.AddItem "320 x 240"
    cmbRes.AddItem "512 x 384"
    cmbRes.AddItem "640 x 480"
    cmbRes.AddItem "800 x 600"
    cmbRes.ListIndex = 0
    
    diagLoading.Show
    DoEvents
    InitAudio
    Unload diagLoading
    Me.Show
End Sub


Private Sub loadPic(map As String)
    'load the pic here
    picMaze.Picture = LoadPicture(File1.Path & "\" & map)
    
    picMaze.Left = 0
    picMaze.Top = 0
    picMaze.Visible = True
    
    HScroll1.Value = 0
    HScroll1.Max = picCont.ScaleWidth - picMaze.Width
    HScroll1.Min = 0
    HScroll1.LargeChange = 100
    
    VScroll1.Value = 0
    VScroll1.Max = picCont.ScaleHeight - picMaze.Height
    VScroll1.Min = 0
    VScroll1.LargeChange = 100
    
    If picMaze.Width <= picCont.ScaleWidth Then
        HScroll1.Enabled = False
    Else
        HScroll1.Enabled = True
    End If
    If picMaze.Height <= picCont.ScaleHeight Then
        VScroll1.Enabled = False
    Else
        VScroll1.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeinitAudio
End Sub

Private Sub HScroll1_Change()
    picMaze.Left = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    picMaze.Left = HScroll1.Value
End Sub

Private Sub picMaze_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pos.X = X
    pos.Y = Y
    timeToEnd = False
    ReDim LevelList(0)
    LevelList(0) = File1.List(File1.ListIndex)
    numLevels = 1
    curLevel = 1
    
    Me.Visible = False
    
    Dim iWidth As Long, iHeight As Long
    
    iWidth = Screen.Width / Screen.TwipsPerPixelX
    iHeight = Screen.Height / Screen.TwipsPerPixelY
    
    Dim noChange As Boolean
    Select Case cmbRes.ListIndex
        Case 1
            ChangeRes 320, 240
        Case 2
            ChangeRes 512, 384
        Case 3
            ChangeRes 640, 480
        Case 4
            ChangeRes 800, 600
        Case Else
            noChange = True
    End Select
    
    frmMain.Show vbModal, Me
    
    If Not noChange Then ChangeRes iWidth, iHeight
    
    Me.Visible = True
End Sub

Private Sub VScroll1_Change()
    picMaze.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    picMaze.Top = VScroll1.Value
End Sub

