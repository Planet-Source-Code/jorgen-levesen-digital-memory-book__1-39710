Attribute VB_Name = "ModuleMPEG"
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Const SND_LOOP = &H8
Public Const SND_ASYNC = &H1
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
    (ByVal lpszName As String, ByVal hModule As Long, _
    ByVal dwFlags As Long) As Long
Public strSoundPath As String

Public Total As String
Public Windir As String
Public glo_hWnd As Long
Public glo_from As Long
Public glo_to As Long

Public Function OpenMPEG(hWnd As Long, filename As String, typeAviOrMpeg As String) As String
Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFie As String
    On Error Resume Next
    lenShort = GetShortPathName(filename, tmp, 255)
    ShortPathAndFie = Left$(tmp, lenShort)
        
    glo_hWnd = hWnd
    cmdToDo = "open " & ShortPathAndFie & " type " & typeAviOrMpeg & " Alias mpeg parent " & hWnd & " Style 1073741824"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        OpenMPEG = ret: Exit Function
    End If
        
    OpenMPEG = "Success"
End Function

Public Function PlayMPEG(from_where As String, to_where As String) As String
    On Error Resume Next
    If from_where = vbNullString And to_where = vbNullString Then
        glo_from = 1
        glo_to = GetTotalframes
    ElseIf Not from_where = vbNullString And Not to_where = vbNullString Then
        glo_from = from_where
        glo_to = to_where
    ElseIf Not from_where = vbNullString And to_where = vbNullString Then
        glo_from = from_where
        glo_to = GetTotalframes
    ElseIf from_where = vbNullString And Not to_where = vbNullString Then
        glo_from = 1
        glo_to = to_where
    End If

Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128

    cmdToDo = "play mpeg from " & glo_from & " to " & glo_to
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        PlayMPEG = ret
        Exit Function
    End If
    
    PlayMPEG = "Success"
End Function

Public Function CloseMPEG() As String
Dim dwReturn As Long
Dim ret As String * 128
    On Error Resume Next
    dwReturn = mciSendString("Close mpeg", 0&, 0&, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        CloseMPEG = ret
        Exit Function
    End If
    
    CloseMPEG = "Success"
End Function

Public Function PauseMPEG() As String
Dim dwReturn As Long
Dim ret As String * 128
    dwReturn = mciSendString("Pause mpeg", 0&, 0&, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        PauseMPEG = ret
        Exit Function
    End If
        
    PauseMPEG = "Success"
End Function

Public Function StopMPEG() As String
Dim dwReturn As Long
Dim ret As String * 128
    dwReturn = mciSendString("Stop mpeg", 0&, 0&, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        StopMPEG = ret
        Exit Function
    End If
    
    StopMPEG = "Success"
End Function

Public Function ResumeMPEG() As String
'calling ResumeMPEG will Resume the multimedia file
Dim dwReturn As Long
Dim ret As String * 128
    dwReturn = mciSendString("Resume mpeg", 0&, 0&, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        ResumeMPEG = ret
        Exit Function
    End If
    
    ResumeMPEG = "Success"
End Function

Public Function GetStatusMPEG() As String
Dim dwReturn As Long
Dim status As String * 255
Dim ret As String * 255

    dwReturn = mciSendString("status mpeg mode", status, 255, 0&)
    
    If Not dwReturn = 0 Then  'not success
        GetStatusMPEG = "ERROR"
        Exit Function
    End If

Dim i As Integer
Dim CharA As String
Dim RChar As String
    RChar = Right$(status, 1)
    For i = 1 To Len(status)
        CharA = Mid(status, i, 1)
        If CharA = RChar Then Exit For
        GetStatusMPEG = GetStatusMPEG + CharA
    Next i
End Function

Public Function GetTotalframes() As Long
Dim dwReturn As Long
Dim Total As String * 255
    
    dwReturn = mciSendString("set mpeg time format frames", Total, 255, 0&)
    dwReturn = mciSendString("status mpeg length", Total, 255, 0&)
    
    If Not dwReturn = 0 Then  'not success
        GetTotalframes = -1
        Exit Function
    End If
    
    GetTotalframes = Val(Total)
End Function

Public Function GetTotalTimeByMS() As Long
Dim dwReturn As Long
Dim TotalTime As String * 255

    dwReturn = mciSendString("set mpeg time format ms", Total, 255, 0&)
    dwReturn = mciSendString("status mpeg length", TotalTime, 255, 0&)
    
    mciSendString "set mpeg time format frames", Total, 255, 0& ' return focus to frames not to time
    
    If Not dwReturn = 0 Then  'not success
        GetTotalTimeByMS = -1
        Exit Function
    End If
    
    GetTotalTimeByMS = Val(TotalTime)
End Function

Public Function MoveMPEG(to_where As Long) As String
Dim dwReturn As Long
Dim ret As String * 255
    
    dwReturn = mciSendString("seek mpeg to " & to_where, 0&, 0&, 0&)
    mciSendString "Play mpeg", 0&, 0&, 0&
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        MoveMPEG = ret
        Exit Function
    End If
    MoveMPEG = "Success"
End Function

Public Function GetCurrentMPEGPos() As Long
Dim dwReturn As Long
Dim pos As String * 255
    
    dwReturn = mciSendString("status mpeg position", pos, 255, 0&)
    
    If Not dwReturn = 0 Then  'not success
        GetCurrentMPEGPos = -1
        Exit Function
    End If
    
    GetCurrentMPEGPos = Val(pos)
End Function

Public Function PutMPEG(Left As Long, Top As Long, Width As Long, Height As Long) As String
Dim dwReturn As Long
Dim ret As String * 255
    If Width = 0 Or Height = 0 Then
        Dim rec As RECT
        Call GetWindowRect(glo_hWnd, rec)
        Width = rec.Right - rec.Left
        Height = rec.Bottom - rec.Top
    End If
    
    dwReturn = mciSendString("put mpeg window at " & Left & " " & Top & " " & Width & " " & Height, 0&, 0&, 0&)
    
    If Not dwReturn = 0 Then  'not success
        mciGetErrorString dwReturn, ret, 128
        PutMPEG = ret
        Exit Function
    End If
    
    PutMPEG = "Success"
End Function
Public Function GetPercent() As Long
On Error Resume Next
Dim TotalFrames As Long
Dim CurrFrame As Long
    TotalFrames = GetTotalframes
    CurrFrame = GetCurrentMPEGPos
    
    If TotalFrames = -1 Or CurrFrame = -1 Then
    GetPercent = -1
    Exit Function
    End If
    
    GetPercent = CurrFrame * 100 / TotalFrames
End Function
Public Function GetFramesPerSecond() As Long
Dim TotalFrames As Long
Dim TotalTime As Long
    TotalTime = GetTotalTimeByMS
    TotalFrames = GetTotalframes
    If TotalFrames = -1 Or TotalTime = -1 Then
        GetFramesPerSecond = -1
        Exit Function
    End If
    GetFramesPerSecond = TotalFrames / (TotalTime / 1000)
End Function
Public Function AreMPEGAtEnd() As Boolean
Dim currpos As Long
    currpos = Val(GetCurrentMPEGPos)
    If glo_to = currpos Or (glo_to - 1) < currpos Then
        AreMPEGAtEnd = True
    Else
        AreMPEGAtEnd = False
    End If
End Function
Public Sub SetAutoRepeat(autoTrueOrFalse As Boolean)
'This cool sub if you want to make the multimedia file auto repeat by it self or remove the auto repeat
'if the parameter is true will make the function Auto repeat or it else will remove the auto repeat
    If autoTrueOrFalse = True Then
        Call SetTimer(glo_hWnd, 500, 100, AddressOf TimerFunction)
    Else
        Call KillTimer(glo_hWnd, 500)
    End If
End Sub

Sub TimerFunction()
'Important for auto repeat
Dim currpos As Long
    currpos = Val(GetCurrentMPEGPos)
    If glo_to = currpos Or (glo_to - 1) < currpos Then PlayMPEG Str(glo_from), Str(glo_to)
End Sub

Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
Dim Res As String
Dim tmp As String * 255
    Res = GetWindowsDirectory(tmp, 255)
    Windir = Left$(tmp, Res)
    Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
Dim tmp As String * 255
Dim Res As String
    Res = GetWindowsDirectory(tmp, 255)
    Windir = Left$(tmp, Res)
    Res = GetPrivateProfileString("MCI", typeDevice, "None", tmp, 255, Windir & "\" & "system.ini")
    GetDefaultDevice = Left$(tmp, Res)
End Function
