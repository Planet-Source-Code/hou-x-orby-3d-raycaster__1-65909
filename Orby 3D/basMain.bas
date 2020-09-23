Attribute VB_Name = "basMain"
Option Explicit

Public Type POINTAPIF
        X As Single
        Y As Single
End Type

Public pos As POINTAPIF

Public timeToEnd As Boolean
Public showFPS As Boolean

Public LevelList() As String
Public numLevels As Long
Public curLevel As Long

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
    
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'Audio
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public BGMLen(1 To 8) As Long
Private SEBuffLen As Long
Private curSE As Long

Public Sub InitAudio()
    Dim i As Long
    Dim Path As String
    Dim mssg As String * 255
    
    Path = App.Path
    i = GetShortPathName(Path, Path, Len(Path))
    Path = Left$(Path, i)
    
    For i = 1 To 8
        mciSendString "open " & Path & "\Music\" & i & ".mid type sequencer alias BGM" & i, 0&, 0, 0
        mciSendString "status BGM" & i & " length", mssg, 255, 0
        BGMLen(i) = Val(mssg)
    Next
    
    SEBuffLen = 3
    ReDim SE(1 To SEBuffLen)
    For i = 1 To SEBuffLen
        mciSendString "open " & Path & "\SE\1.wav type waveaudio alias SE" & i, 0&, 0, 0
    Next
End Sub

Public Sub DeinitAudio()
    Dim i As Long
    
    For i = 1 To 8
        mciSendString "stop BGM" & i, 0&, 0, 0
        mciSendString "close BGM" & i, 0&, 0, 0
    Next
    
    For i = 1 To SEBuffLen
        mciSendString "stop SE" & i, 0&, 0, 0
        mciSendString "close SE" & i, 0&, 0, 0
    Next
End Sub

Public Sub PlaySE()
    curSE = curSE + 1
    mciSendString "play SE" & curSE & " from 0", 0&, 0, 0
    If curSE = SEBuffLen Then curSE = 0
End Sub

Public Sub PlayBGM(ByVal Track As Long)
    If Track = 0 Then Exit Sub
    mciSendString "play BGM" & Track & " from 0", 0&, 0, 0
End Sub

Public Sub StopBGM(ByVal Track As Long)
    If Track = 0 Then Exit Sub
    mciSendString "stop BGM" & Track, 0&, 0, 0
End Sub

Public Sub RepeatBGM(ByVal Track As Long)
    Dim pos As Long
    Dim mssg As String * 255
    Dim i As Long
    
    If Track = 0 Then Exit Sub
    
    mciSendString "status BGM" & Track & " position", mssg, 255, 0
    pos = Val(mssg)
    
    If pos = BGMLen(Track) Then
        mciSendString "play BGM" & Track & " from 0", 0&, 0, 0
    End If
End Sub

Public Sub ChangeRes(ByVal iWidth As Long, ByVal iHeight As Long)
    Dim blnWorked As Boolean
    Dim i As Long
    Dim DevM As DEVMODE
    
    i = 0
    
    Do
        blnWorked = EnumDisplaySettings(0&, i, DevM)
        i = i + 1
    Loop Until (blnWorked = False)
        
    With DevM
        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        .dmPelsWidth = iWidth
        .dmPelsHeight = iHeight
    End With
    Call ChangeDisplaySettings(DevM, 0)
End Sub

