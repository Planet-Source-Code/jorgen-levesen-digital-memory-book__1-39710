Attribute VB_Name = "MasterKid"
Option Explicit
Public Declare Function GetLogicalDrives Lib "kernel32" () As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As _
    Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_PASTE = &H302

Declare Function GetModuleHandle Lib _
    "Kernel" (ByVal lpModuleName As String) As Integer
    
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

Public Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wid           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String
    cch           As Long
End Type

'Menu item constants.
Public Const SC_CLOSE       As Long = &HF060&
Public Const xSC_CLOSE   As Long = -10
'SetMenuItemInfo fMask constants.
Public Const MENU_STATE     As Long = &H1&
Public Const MENU_ID        As Long = &H2&

'SetMenuItemInfo fState constants.
Public Const MFS_GRAYED     As Long = &H3&
Public Const MFS_CHECKED    As Long = &H8&

'SendMessage constants.
Public Const WM_NCACTIVATE  As Long = &H86

Public Declare Function MoveWindow Lib "user32" _
                       (ByVal hWnd As Long, _
                        ByVal x As Long, ByVal y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, _
                        ByVal bRepaint As Long) As Long
                        
Public Declare Function ShellExceCute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, _
    ByVal bInvert As Long) As Long


Public Declare Function SetWindowPos Lib "user32" _
   (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cX As Long, _
    ByVal cY As Long, ByVal wFlags As Long) As Long

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
                        (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

' Rtns True (non zero) on succes, False on failure
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
                        (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

' Rtns True (non zero) on succes, False on failure
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Const MaxLFNPath = 260

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MaxLFNPath
        cShortFileName As String * 14
End Type

Public Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp&, ByVal wPixTypes&) As Long
Public Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long

Public Const LB_INITSTORAGE = &H1A8

'FindFirstFile failure rtn value
Public Const INVALID_HANDLE_VALUE = -1
Public Const LB_ADDSTRING = &H180
Public Const WM_SETREDRAW = &HB
Public Const WM_VSCROLL = &H115
Public Const SB_BOTTOM = 7

Global bProgramNotAccesible As Boolean
Global boolNewRecord As Boolean
Global bErrorLog As Boolean
Global Password As Boolean
Global sErrorFile As String
Global sFileName As String
Global iWhichForm As Long
Global FileExt As String
Global retText As String
Global i As Integer
Global n As Integer
Global iTab As Integer
Global glChildNo As Long
Global gsChildName As String
Global WClone As Recordset
Global sRegNo As String
Global Const SW_SHOWNORMAL = 1

Public cPrint As clsMultiPgPreview
Public QuitCommand As Boolean

Global dbKids As Database
Global dbKidsTxt As String
Global dbKidPic As Database
Global dbKidPicTxt As String
Global dbKidLang As Database
Global dbKidLangTxt As String

Global iSelBold As Boolean
Global iSelUlin As Boolean
Global iSelItal As Boolean
Global iSelLeft As Boolean
Global iSelMid As Boolean
Global iSelRight As Boolean

Global LineCounter As Integer
Global MaxLines As Integer
Global PageCounter As Integer
Global LeftMargin As Integer
Global TopMargin As Integer
Global BottomMargin As Integer

Public Const LB_GETITEMHEIGHT = &H1A1

'colors
Global lRed As Long
Global lGreen As Long
Global lBlue As Long

' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message
Global Const MB_DEFBUTTON2 = 256       ' Second button is default

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IdNo = 7                  ' No button pressed

' Rich text box values
Global Const rtfLeft = 0
Global Const rtfRight = 1
Global Const rtfCenter = 2

' Colors
Global Const BLACK = &H0&
Global Const Red = &HFF&
Global Const Green = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const Blue = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
Global Const GRAY = &HC0C0C0

'center child forms
'--------------------------------------------------------------------------------------------------------------------------------
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, _
    ByVal uParam As Long, _
    lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

'Type RECT
    'Left As Long
   'Top As Long
   'Right As Long
   'Bottom As Long
'End Type

Public Const SPI_GETWORKAREA = 48
'--------------------------------------------------------------------------------------------------------------------------------
Private Declare Function RegOpenKey Lib _
   "advapi32" Alias "RegOpenKeyA" (ByVal hKey _
   As Long, ByVal lpSubKey As String, _
   phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
   Lib "advapi32" Alias "RegQueryValueExA" _
   (ByVal hKey As Long, ByVal lpValueName As _
   String, lpReserved As Long, lptype As _
   Long, lpData As Any, lpcbData As Long) _
   As Long

Private Declare Function RegCloseKey& Lib "advapi32" (ByVal hKey&)

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Sub CheckSpelling(cBox As RichTextBox)
Dim oSpellChecker As Object
Dim stText As String
Dim stNew_Text As String
Dim iPosition As Integer

    On Error GoTo OpenError
    
    'If you want to spell check a group of textboxes in a
    'control array.. You can Spell check all of these with a
    'do until loop.. There are multiple ways to do this but
    'this is a simple one.

    Set oSpellChecker = CreateObject("Word.Basic")
    
    oSpellChecker.FileNew
    oSpellChecker.Insert cBox.Text
    oSpellChecker.ToolsSpelling
    oSpellChecker.EditSelectAll
    stText = oSpellChecker.Selection()
    oSpellChecker.FileExit 2

    If Right$(stText, 1) = vbCr Then _
      stText = Left$(stText, Len(stText) - 1)
    stNew_Text = ""
    iPosition = InStr(stText, vbCr)
    Do While iPosition > 0
      stNew_Text = stNew_Text & Left$(stText, iPosition - 1) & vbCrLf
      stText = Right$(stText, Len(stText) - iPosition)
      iPosition = InStr(stText, vbCr)
    Loop
    
    stNew_Text = stNew_Text & stText
    cBox.Text = stNew_Text
    MsgBox "Spell Check Complete"
    Exit Sub

OpenError:
    MsgBox "Error" & Str$(Error.Number) & " opening Word." & vbCrLf & Error.Description
End Sub

Public Sub CenterMe(frm As Form)
    frm.Move ((MDIMasterKid.Width - MDIMasterKid.Width) - frm.Width) / 2, 0
        '((MDIMasterKid.Height - MDIMasterKid.Picture2.Height - MDIMasterKid.StatusBar1.Height) - frm.Height) / 2
End Sub


Public Sub Dither(vForm As Form)
Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
      vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub
Public Sub btnClipboardClick(Index As Integer, cBox As RichTextBox)
  Select Case Index
  Case 0  'Copy text to clipboard
        Clipboard.Clear
        Clipboard.SetText cBox.SelText
    Case 1  'Paste text/picture from clipboard
        If Clipboard.GetFormat(vbCFText) Then
            cBox.SelText = Clipboard.GetText()
        ElseIf Clipboard.GetFormat(vbCFDIB) Then
            ' Paste the picture into the RichTextBox.
            SendMessage cBox.hWnd, WM_PASTE, 0, 0
        End If
    Case 2  'cut
        Clipboard.Clear
        Clipboard.SetText cBox.SelText
        cBox.SelText = ""
    Case Else
    End Select
End Sub


Public Sub btnJustifyClick(Index As Integer, cBox As RichTextBox)
    Select Case Index
    Case 0  'left justified text
        If iSelLeft = False Then
            cBox.SelAlignment = 0
            iSelLeft = True
        ElseIf iSelLeft = True Then
            cBox.SelAlignment = 0
            iSelLeft = False
        End If
    Case 1  'mid justified text
        If iSelMid = False Then
            cBox.SelAlignment = 2
            iSelMid = True
        ElseIf iSelMid = True Then
            cBox.SelAlignment = 0
            iSelMid = True
        End If
    Case 2  'Right justified text
        If iSelRight = False Then
            cBox.SelAlignment = 1
            iSelRight = True
        ElseIf iSelRight = True Then
            cBox.SelAlignment = 0
            iSelRight = False
        End If
    Case Else
    End Select
End Sub
Public Sub btnLetterClick(Index As Integer, cBox As RichTextBox)
    Select Case Index
    Case 0  'font bold
        If iSelBold = False Then
            cBox.SelBold = True
            iSelBold = True
        ElseIf iSelBold = True Then
            cBox.SelBold = False
            iSelBold = False
        End If
    Case 1  'italic text
        If iSelItal = False Then
            cBox.SelItalic = True
            iSelItal = True
        ElseIf iSelItal = True Then
            cBox.SelItalic = False
            iSelItal = False
        End If
    Case 2  'underlined text
        If iSelUlin = False Then
            cBox.SelUnderline = True
            iSelUlin = True
        ElseIf iSelUlin = True Then
            cBox.SelUnderline = False
            iSelUlin = False
        End If
    Case Else
    End Select
End Sub

Public Sub FontPopUp(FontName As String, cBox As RichTextBox)
        On Error Resume Next
        cBox.SelFontName = FontName
End Sub
Public Sub FontSizePopUp(iNum As Integer, cBox As RichTextBox)
    On Error Resume Next
    cBox.SelFontSize = iNum
End Sub
Public Sub formatColor(cBox As RichTextBox)
     cBox.SelColor = RGB(lRed, lGreen, lBlue)
End Sub

Public Sub HideAllButtons()
    On Error Resume Next
    With MDIMasterKid.Toolbar1
        For i = 3 To 10
            .Buttons(i).Enabled = False
        Next
    End With
End Sub

Public Sub HideKids()
    On Error Resume Next
    With MDIMasterKid
        .Label1.Enabled = False
        .cmbChildren.Enabled = False
    End With
End Sub

Public Function IsMicrosoftMailRunning()
    On Error GoTo IsMicrosoftMailRunning_Err
    IsMicrosoftMailRunning = GetModuleHandle("MSMail")
    
IsMicrosoftMailRunning_Err:
    If Err Then
        MsgBox "Please start MS Outlook"
    End If
End Function

Public Sub ShowAllButtons()
    On Error Resume Next
    With MDIMasterKid.Toolbar1
        For i = 3 To 10
            .Buttons(i).Enabled = True
        Next
    End With
End Sub

Public Sub ShowKids()
    On Error Resume Next
    With MDIMasterKid
        .Label1.Enabled = True
        .cmbChildren.Enabled = True
    End With
End Sub

Public Sub StrikeLine(rch As RichTextBox)
Dim txt As String
Dim TxtLen As Integer
Dim crlen As Integer
Dim start_pos As Integer
Dim end_pos As Integer

    ' Find the previous carriage return.
    txt = rch.Text
    crlen = Len(vbCrLf)
    For start_pos = rch.SelStart To 1 Step -1
        If Mid$(txt, start_pos, crlen) = vbCrLf _
            Then Exit For
    Next start_pos
    If start_pos < 1 Then
        start_pos = 1
    Else
        start_pos = start_pos + crlen
    End If
    
    ' Find the next carriage return.
    end_pos = InStr(rch.SelStart, txt, vbCrLf)
    If end_pos = 0 Then end_pos = Len(txt)

    rch.SelStart = start_pos
    rch.SelLength = end_pos - start_pos
    rch.SelStrikeThru = Not rch.SelStrikeThru
End Sub
Public Function compareDates(Date1 As Date, Date2 As Date) As Integer
' If Date1 is greater the function returns the number 1
' If Date2 is greater the function returns the number 2
' If Date1 equals Date2 the function returns the number 0
    Select Case Date1
        Case Is > Date2
            compareDates = 1
        Case Is < Date2
            compareDates = 2
        Case Else
            compareDates = 0
        End Select
End Function

Public Sub CheckAlarm()
Dim iDif As Integer
Dim rsAlarm As Recordset
    On Error Resume Next
    Set rsAlarm = dbKids.OpenRecordset("Alarm")
   
    With rsAlarm
        .MoveFirst
        Do While Not .EOF
            If compareDates(Format(CDate(.Fields("AlarmDate")), "dd.mm.yyyy"), Format(CDate(Now), "dd.mm.yyyy")) = 2 Then
                frmNotice.Text1.Text = Format(CDate(.Fields("RealAlarmDate")), "dd.mm.yyyy")
                frmNotice.Text2.Text = Format(CDate(.Fields("RealAlarmTime")), "hh:mm")
                If Not IsNull(.Fields("AlarmNote")) Then
                    frmNotice.Text3.Text = .Fields("AlarmNote")
                Else
                    frmNotice.Text3.Text = "??????????"
                End If
                frmNotice.Show 1
                .Delete
            ElseIf compareDates(Format(CDate(.Fields("AlarmDate")), "dd.mm.yyyy"), Format(CDate(Now), "dd.mm.yyyy")) = 0 Then
                iDif = DateDiff("n", Format(CDate(.Fields("AlarmTimeSet")), "hh:mm"), Format(CDate(Now), "hh:mm"))
                If iDif >= 0 Then
                    frmNotice.Text1.Text = Format(CDate(.Fields("RealAlarmDate")), "dd.mm.yyyy")
                    frmNotice.Text2.Text = Format(CDate(.Fields("RealAlarmTime")), "hh:mm")
                    If Not IsNull(.Fields("AlarmNote")) Then
                        frmNotice.Text3.Text = .Fields("AlarmNote")
                    Else
                        frmNotice.Text3.Text = "??????????"
                    End If
                    frmNotice.Show 1
                    .Delete
                End If
            End If
        .MoveNext
        Loop
    End With
    
    rsAlarm.Close
End Sub


Public Sub RichTextSelChange(cBox As RichTextBox)
        On Error Resume Next
        If cBox.SelBold = True Then
            iSelBold = True
        ElseIf cBox.SelUnderline = True Then
            iSelUlin = True
        ElseIf cBox.SelItalic = True Then
            iSelItal = True
        ElseIf cBox.SelAlignment = 0 Then
            iSelLeft = True
        ElseIf cBox.SelAlignment = 2 Then
            iSelMid = True
        ElseIf cBox.SelAlignment = 1 Then
            iSelRight = True
        End If
End Sub

Public Function StopMidiFile(FilePath As String) As Boolean
Dim iset As Long
    On Error Resume Next
    iset = mciSendString("stop midi", "", 0, 0)
    iset = mciSendString("close midi", "", 0, 0)
End Function
Public Function PlayMidiFile(FilePath As String) As Boolean
Dim iset As Long
    On Error Resume Next
    iset = mciSendString("open sequencer!" & FilePath & " alias midi", "", 0, 0)
    iset = mciSendString("play midi", "", 0, 0)
    PlayMidiFile = (iset = 0)
End Function

Public Sub WriteErrorFile(ErrorString As String, sForm As String)
    On Error Resume Next
    If bErrorLog Then
        Open sErrorFile For Append As #1
        Write #1, ErrorString, sForm
        Close #1
    Else
        sErrorFile = "C:\DayPlanerError.txt"
        Open sErrorFile For Output As #1
        Write #1, ErrorString, sForm
        Close #1
        bErrorLog = True
    End If
End Sub
Public Sub SendErrorMail()
Dim oLapp As Object, rsUser As Recordset
Dim oItem As Object
Dim ErrorString As String, sForm As String, sError As String

    If bErrorLog Then
         On Error GoTo errorHandler2
         Set rsUser = dbKids.OpenRecordset("User")
         
         Set oLapp = CreateObject("Outlook.application")
         Set oItem = oLapp.CreateItem(0)
         With oItem
            Open sErrorFile For Input As #1
            Do While Not EOF(1) ' Loop until end of file.
                Input #1, ErrorString, sForm
                sError = sError & sForm & ":  " & ErrorString & vbCrLf
            Loop
                .Body = sError
                .Subject = "Error in system"
                .To = rsUser.Fields("E-MailSysRespons")
                .Send
         End With
         bErrorLog = False
         Close #1
         Kill sErrorFile
    End If
    
errorHandler2:
    Set oLapp = Nothing
    Set oItem = Nothing
    Exit Sub
End Sub

Public Sub SendOutlookMail(Subject As String, Recipient As String, Message As String)
    On Error GoTo errorHandler
    Dim oLapp As Object
    Dim oItem As Object
    
    'IsMicrosoftMailRunning
    
    Set oLapp = CreateObject("Outlook.application")
    Set oItem = oLapp.CreateItem(0)
   
    With oItem
       .Subject = Subject
       .To = Recipient
       .Body = Message
       .Send
    End With
    
errorHandler:
    Set oLapp = Nothing
    Set oItem = Nothing
    Exit Sub
End Sub

Public Sub onGotFocus()
    On Error Resume Next
    If TypeOf Screen.ActiveControl Is TextBox Then
        With Screen.ActiveControl
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
End Sub

