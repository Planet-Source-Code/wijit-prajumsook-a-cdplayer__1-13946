Attribute VB_Name = "Module1"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public fForwardSpeed As Long
Public fPlaying As Boolean
Public fCDLoaded As Boolean
Public numTracks As Integer
Public track As Integer
Public trackLength() As String
Public min As Integer
Public sic As Integer
Public cmd As String
Public startplay As Boolean
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Public Sub DragFrm(frm As Form)
    Call ReleaseCapture
    Call SendMessage(frm.hwnd, &HA1, 2, 0)
End Sub

Public Sub UpdateTimer()
    Static cdmedia As String * 30
    
    mciSendString "status cd media present", cdmedia, Len(cdmedia), 0
    
        If (fCDLoaded = False) Then
            mciSendString "status cd number of tracks wait", cdmedia, Len(cdmedia), 0
            
            If (numTracks = 1) Then
                Exit Sub
            End If
            
            mciSendString "status cd length wait", cdmedia, Len(cdmedia), 0
            frmMain.lblTime.Caption = " Tracks: " & numTracks & " Time: " & cdmedia
            Dim i As Integer
            For i = 1 To numTracks
                cmd = "status cd length track " & i
                mciSendString cmd, cdmedia, Len(cdmedia), 0
                trackLength(i) = cdmedia
            Next
            fCDLoaded = True
            mciSendString "seek cd to 1", 0, 0, 0
            startplay = True
        End If
        
        mciSendString "status cd position", cdmedia, Len(cdmedia), 0
        
        mciSendString "status cd mode", cdmedia, Len(cdmedia), 0
        startplay = True
        If (fCDLoaded = True) Then
            fCDLoaded = False
            fPlaying = False
            frmMain.lblTime.Caption = "No CD"
            frmMain.lblTiming.Caption = "[00]00:00 - 00:00:00"
        End If
End Sub
