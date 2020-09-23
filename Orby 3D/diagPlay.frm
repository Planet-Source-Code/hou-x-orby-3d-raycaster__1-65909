VERSION 5.00
Begin VB.Form diagPlay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Orby 3D"
   ClientHeight    =   5010
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "diagPlay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRes 
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Resolution:"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Orby 3D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "diagPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If MsgBox("For a significant speed boost, compile the program and run the executable. Continue?", vbYesNo) = vbNo Then End
    
    Label2.Caption = "Created by: Hou Xiong"
    Label2.Caption = Label2.Caption & vbCrLf & vbCrLf & "Welcome to Orby 3D RayCaster!!!  " & _
        "This little game demonstrates the recursive ray casting " & _
        "engine that I have developed in pure VB and API.  Look through the code for more description.  Enjoy!!" & _
        vbCrLf & vbCrLf & "Objective:" & vbCrLf & _
        "   Collect all the orbs." & _
        vbCrLf & vbCrLf & "Controls:" & vbCrLf & _
        "   Up - W" & vbCrLf & _
        "   Down - S" & vbCrLf & _
        "   Left - A" & vbCrLf & _
        "   Right - D" & vbCrLf & _
        "   Look - Mouse" & vbCrLf & _
        "   Map - M" & vbCrLf & _
        "   Run - Shift" & vbCrLf & _
        "   Quit - Esc" & vbCrLf & _
        "   Show FPS - F" & vbCrLf
        
        cmbRes.AddItem "Don't Change"
        cmbRes.AddItem "320 x 240"
        cmbRes.AddItem "512 x 384"
        cmbRes.AddItem "640 x 480"
        cmbRes.AddItem "800 x 600"
        cmbRes.ListIndex = 1
End Sub

Private Sub OKButton_Click()
    Dim f As Long
    
    f = FreeFile
    Open App.Path & "\Maps\Playlist.txt" For Input As #f
    
    numLevels = 0
    Do While Not EOF(f)
        numLevels = numLevels + 1
        ReDim Preserve LevelList(numLevels - 1)
        Input #f, LevelList(numLevels - 1)
    Loop
    
    Close f
    
    If numLevels = 0 Then
        MsgBox "There is no playlist!", vbInformation
        Exit Sub
    End If
    
    curLevel = 1
    Me.Visible = False
    
    diagLoading.Show
    DoEvents
    InitAudio
    Unload diagLoading
    
    Dim iWidth As Long, iHeight As Long
    
    iWidth = Screen.width / Screen.TwipsPerPixelX
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
    
    pos.X = 0
    pos.Y = 0
    timeToEnd = False
    frmMain.Show vbModal, Me
    DeinitAudio
    Unload Me
    
    If Not noChange Then ChangeRes iWidth, iHeight
    
    MsgBox "That's it! I hope you enjoyed it.  Thank you for playing.", , "Orby 3D"
End Sub
