VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEmail 
   BackColor       =   &H00000000&
   Caption         =   "E-mail"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   ControlBox      =   0   'False
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSpell 
      Caption         =   "&Spell check"
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton btnAttachment 
      Caption         =   "&Attachment"
      Height          =   495
      Left            =   1440
      TabIndex        =   14
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send this Mail"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000080&
      Caption         =   "Attachments"
      ForeColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   6360
      TabIndex        =   7
      Top             =   0
      Width           =   2175
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Dbl Click to remove"
         Top             =   4560
         Width           =   1935
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   1665
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   5895
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Message:"
         ForeColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   5895
         Begin RichTextLib.RichTextBox RichText1 
            Height          =   3615
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   6376
            _Version        =   393217
            BackColor       =   16777152
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmEmail.frx":0442
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton btnCarbonCopy 
         Caption         =   "&Cc.."
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton btnEmailGroup 
         Caption         =   "&To..."
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "E-mail Group"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   0
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBookmarkLang() As Variant, boolGroup As Boolean, boolFirst As Boolean
Dim rsLanguage As Recordset
Public Sub Set_Form_To_Stay_On_Top(handle As Long, set_on_top As Boolean)
' SetWindowPos Flags
  Const SWP_NOSIZE = &H1
  Const SWP_NOMOVE = &H2
  Const SWP_NOACTIVATE = &H10
  Const SWP_SHOWWINDOW = &H40
' SetWindowPos() hwndInsertAfter values
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2
  Dim window_flags As Variant
  Dim position_flag As Variant
  Dim return_code As Long
' Set the configuration to apply to the window.
  window_flags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
  Select Case set_on_top
    Case True
      position_flag = HWND_TOPMOST
    Case False
      position_flag = HWND_NOTOPMOST
  End Select
' Apply the configuration.
    return_code = SetWindowPos(handle, position_flag, 0, 0, 0, 0, window_flags)
End Sub
Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("Label1")) Then
                    .Fields("Label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("Label1")
                End If
                If IsNull(.Fields("Frame2")) Then
                    .Fields("Frame2") = Frame2.Caption
                Else
                    Frame2.Caption = .Fields("Frame2")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
                If IsNull(.Fields("btnSend")) Then
                    .Fields("btnSend") = btnSend.ToolTipText
                Else
                    btnSend.ToolTipText = .Fields("btnSend")
                End If
                If IsNull(.Fields("btnSpell")) Then
                    .Fields("btnSpell") = btnSpell.ToolTipText
                Else
                    btnSpell.ToolTipText = .Fields("btnSpell")
                End If
                If IsNull(.Fields("btnEmailGroup1")) Then
                    .Fields("btnEmailGroup1") = btnEmailGroup.Caption
                Else
                    btnEmailGroup.Caption = .Fields("btnEmailGroup1")
                End If
                If IsNull(.Fields("btnCarbonCopy")) Then
                    .Fields("btnCarbonCopy") = btnCarbonCopy.Caption
                Else
                    btnCarbonCopy.Caption = .Fields("btnCarbonCopy")
                End If
                If IsNull(.Fields("btnCancel")) Then
                    .Fields("btnCancel") = btnCancel.Caption
                Else
                    btnCancel.Caption = .Fields("btnCancel")
                End If
                If IsNull(.Fields("btnAttachment")) Then
                    .Fields("btnAttachment") = btnAttachment.ToolTipText
                Else
                    btnAttachment.ToolTipText = .Fields("btnAttachment")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
            
            'this language was not found, make it. Find the English text first
            strHelp = " "
            .MoveFirst
            Do While Not .EOF
                If .Fields("Language") = "ENG" Then
                    If Not IsNull(.Fields("Help")) Then
                        strHelp = .Fields("Help")
                        Exit Do
                    End If
                End If
            .MoveNext
            Loop
            
        .MoveFirst
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Label1") = Label1.Caption
        .Fields("Frame2") = Frame2.Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("btnSend") = btnSend.ToolTipText
        .Fields("btnSpell") = btnSpell.ToolTipText
        .Fields("btnEmailGroup1") = btnEmailGroup.Caption
        .Fields("btnCarbonCopy") = btnCarbonCopy.Caption
        .Fields("btnCancel") = btnCancel.Caption
        .Fields("btnAttachment") = btnAttachment.ToolTipText
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub btnAttachment_Click()
    List1.Clear
    Me.Width = Me.Width + 2300
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCarbonCopy_Click()
    'Set_Form_To_Stay_On_Top Me.hwnd, False
    'frmSelectAddress.Show 1
    'Set_Form_To_Stay_On_Top Me.hwnd, True
End Sub

Private Sub btnEmailGroup_Click()
    'Set_Form_To_Stay_On_Top Me.hwnd, False
    'frmSelectAddress.Show 1
    'Set_Form_To_Stay_On_Top Me.hwnd, True
End Sub

Private Sub btnSend_Click()
Dim clsOutlook As cOutlookSendMail
    On Error GoTo errorHandler
    Set_Form_To_Stay_On_Top Me.hWnd, False
    
    If Len(Text1(0).Text) <> 0 Then
        Set clsOutlook = New cOutlookSendMail
        With clsOutlook
            .StartOutlook
            .CreateNewMail
           .Recipient_TO = Text1(0).Text
           If Len(Text2.Text) <> 0 Then
            .Recipient_CC = Text2.Text
           End If
           .Subject = Text1(1).Text
           .Body = RichText1.Text
           If List1.ListCount <> 0 Then
                For i = 0 To List1.ListCount - 1
                    .Attachment List1.List(i)
                Next
           End If
           .SendMail
           .CloseOutlook
        End With
    Else
        Exit Sub
    End If

errorHandler:
    Set clsOutlook = Nothing    ' free memory
    Unload Me
End Sub

Private Sub btnSpell_Click()
    On Error Resume Next
    Call CheckSpelling(frmEmail.RichText1)
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    List1.AddItem Dir1.List(Dir1.ListIndex) & "\" & File1.List(File1.ListIndex)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If boolFirst Then Exit Sub
    ShowText
    Set_Form_To_Stay_On_Top Me.hWnd, True
    boolFirst = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmEmail")
    'iWhichForm = 22
    boolFirst = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    'iWhichForm = 0
    Set_Form_To_Stay_On_Top Me.hWnd, False
    Set frmEmail = Nothing
End Sub
Private Sub List1_DblClick()
    List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim ItemHeight As Long
   Dim NewIndex As Long
   Static OldIndex As Long

   With List1
      ItemHeight = SendMessage(.hWnd, LB_GETITEMHEIGHT, 0, ByVal 0&)
      ItemHeight = .Parent.ScaleY(ItemHeight, vbPixels, vbTwips)
      NewIndex = .TopIndex + (y \ ItemHeight)
      If NewIndex <> OldIndex Then
         If NewIndex < .ListCount Then
            .ToolTipText = .List(NewIndex)
         Else
            .ToolTipText = vbNullString
         End If
         OldIndex = NewIndex
     End If
   End With
End Sub

Private Sub RichText1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichText1.SelText = vbTab
        KeyCode = 0
    End If
End Sub
