VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "CDPlayer"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TTimer 
      Interval        =   10
      Left            =   4200
      Top             =   840
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmMain.frx":0442
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label lblTiming 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "[00]00:00 - 00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   2010
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "No CD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   1965
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.Image imgNext 
      Height          =   255
      Left            =   3960
      ToolTipText     =   "::- Forward Track -::"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgFNext 
      Height          =   255
      Left            =   3480
      ToolTipText     =   "::- Forward -::"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgFBack 
      Height          =   255
      Left            =   3000
      ToolTipText     =   "::- Rewind -::"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image imgBack 
      Height          =   255
      Left            =   2640
      ToolTipText     =   "::- Previous Track -::"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgMin 
      Height          =   255
      Left            =   4200
      ToolTipText     =   "::- Minimize -::"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image imgEject 
      Height          =   255
      Left            =   2040
      ToolTipText     =   "::- Eject -::"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPause 
      Height          =   255
      Left            =   1440
      ToolTipText     =   "::- Pause -::"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgEnd 
      Height          =   255
      Left            =   4440
      ToolTipText     =   "::- Exit -::"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgStop 
      Height          =   255
      Left            =   1080
      ToolTipText     =   "::- Stop -::"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgPlay 
      Height          =   255
      Left            =   600
      ToolTipText     =   "::- Play -::"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image imgMain 
      Height          =   1695
      Left            =   0
      Picture         =   "frmMain.frx":0884
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pos As Integer
Dim curY As Integer
Dim curPos As Integer
Dim toY As Integer
Dim slider As Boolean

Private Sub Form_Load()
    Static rc As Long
    
    TTimer.Enabled = False
    fForwardSpeed = 5
    fCDLoaded = False
    ' if already running then quit
    If (App.PrevInstance = True) Then
        End
    End If
    
    Dim re As String
    Dim temp As String * 40
    Dim wDir As String
    re = GetWindowsDirectory(temp, Len(temp))
    wDir = Left$(temp, re)
    re = WritePrivateProfileString("MCI", "MPEGVideo", "mciqtz.drv", wDir & "\" & "system.ini")
    
    'if cd is in used then quit
    rc = mciSendString("open mpegvideo type MPEGVideo alias cd", 0, 0, 0) ' wait shareable", 0, 0, Me.hwnd)
    If Not rc = 0 Then
        MsgBox "Your CD-Player is in used", vbExclamation
        End
    End If
    
    mciSendString "set cd time format tmsf wait", 0, 0, 0
    TTimer.Enabled = True
    UpdateTimer
    If startplay Then
        imgPlay_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mciSendString "stop cd", 0, 0, 0
    mciSendString "close all", 0, 0, 0
End Sub

Private Sub imgBack_Click()
    Dim from As String
    If (min = 0 And sec = 0) Then
        If (track > 1) Then
            from = CStr(track - 1)
        Else
            from = CStr(numTracks)
        End If
    Else
        from = CStr(track)
    End If
    
    If (fPlaying) Then
        cmd = "play cd from " & from
        mciSendString cmd, 0, 0, 0
    Else
        cmd = "seek cd to " & from
        mciSendString cmd, 0, 0, 0
    End If
    Call UpdateTimer
End Sub

Private Sub imgEject_Click()
    mciSendString "set cd door open", 0, 0, 0
    Call UpdateTimer
End Sub

Private Sub imgEnd_Click()
    Unload Me
End Sub

Private Sub imgFBack_Click()
    Dim s As String * 40
    mciSendString "set cd time format milliseconds", 0, 0, 0
    mciSendString "status cd position wait", s, Len(s), 0
    If (fPlaying) Then
        cmd = "play cd from " & CStr(CLng(s) - fForwardSpeed * 1000)
    Else
        cmd = "seek cd to " & CStr(CLng(s) - fForwardSpeed * 1000)
    End If
    
    mciSendString cmd, 0, 0, 0
    mciSendString "set cd time format tmsf", 0, 0, 0
    Call UpdateTimer
End Sub

Private Sub imgFNext_Click()
    Dim s As String * 40
    mciSendString "set cd time format milliseconds", 0, 0, 0
    mciSendString "status cd position wait", s, Len(s), 0
    If (fPlaying) Then
        cmd = "play cd from " & CStr(CLng(s) + fForwardSpeed * 1000)
    Else
        cmd = "seek cd to " & CStr(CLng(s) + fForwardSpeed * 1000)
    End If
    
    mciSendString cmd, 0, 0, 0
    mciSendString "set cd time format tmsf", 0, 0, 0
    Call UpdateTimer
End Sub

Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragFrm(Me)
    End If
End Sub

Private Sub imgMin_Click()
    Me.WindowState = 1
End Sub

Private Sub imgNext_Click()
    If (track < numTracks) Then
        If (fPlaying) Then
            cmd = "play cd from " & track + 1
            mciSendString cmd, 0, 0, 0
        Else
            cmd = "seek cd to " & track + 1
            mciSendString cmd, 0, 0, 0
        End If
    Else
        mciSendString "seek to cd 1", 0, 0, 0
    End If
    Call UpdateTimer
End Sub

Private Sub imgPause_Click()
    mciSendString "pause cd", 0, 0, 0
    fPlaying = False
    Call UpdateTimer
End Sub

Private Sub imgPlay_Click()
    mciSendString "play cd", 0, 0, 0
    fPlaying = True
End Sub


Private Sub imgStop_Click()
    mciSendString "stop cd wait", 0, 0, 0
    cmd = "seek cd to " & track
    mciSendString cmd, 0, 0, 0
    fPlaying = False
    Call UpdateTimer
End Sub

Private Sub TTimer_Timer()
    Call UpdateTimer
End Sub
