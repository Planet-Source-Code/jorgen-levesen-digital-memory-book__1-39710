VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLiveUpdate 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Update"
   ClientHeight    =   4095
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7650
   ControlBox      =   0   'False
   Icon            =   "frmLiveUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton btnBeginUpdate 
      Caption         =   "&Start Program Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   3
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7938
            MinWidth        =   7938
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Message"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   5535
      End
      Begin VB.Timer Timer2 
         Left            =   5160
         Top             =   720
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5160
         Top             =   1200
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   6000
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmLiveUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strMyVer As String
Dim strUpdateVer As String
Dim strUpdateDate As String
Dim status As String
Dim lngUpdateTime As Long
Dim boolTransferSuccess As Boolean
Private Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
' Written by:  Blake B. Pell
'              blakepell@hotmail.com
'              bpell@indiana.edu
'              http://www.blakepell.com
'              December 7, 2000
    On Error GoTo errGetInternetFile
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)

    For x = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, x), 1) = "/" Then RealFile$ = Right$(myURL, x - 1)
    Next x
    
    myFile$ = DestDIR + "\" + RealFile$
    Open myFile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    GetInternetFile = True
    Exit Function

errGetInternetFile:
    ' error handler
    MsgBox "An error has occured in the file transfer or write.  Please try again later.", vbInformation
    GetInternetFile = False
    Err.Clear
End Function

Private Sub ReadText()
    On Error Resume Next    'this is only text
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("btnBeginUpdate")) Then
                    .Fields("btnBeginUpdate") = btnBeginUpdate.Caption
                Else
                    btnBeginUpdate.Caption = .Fields("btnBeginUpdate")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Form") = Me.Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("btnBeginUpdate") = btnBeginUpdate.Caption
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Private Sub btnBeginUpdate_Click()

    lngUpdateTime = 0
    Timer2.Interval = 1000
    btnBeginUpdate.Enabled = False
    ProgressBar1.Value = 1
    
    On Error GoTo errUpdate
    status = rsLanguage.Fields("Msg4")
    boolTransferSuccess = GetInternetFile(Inet1, "http://www.levesen.com/downloads/MasterKidInfo.inf", App.Path)

    If boolTransferSuccess = False Then
        ProgressBar1.Value = 3
        Timer2.Interval = 0
        Exit Sub
    End If
       
    ProgressBar1.Value = 2
    status = rsLanguage.Fields("Msg5")
    
    Open App.Path & "\MasterKidInfo.inf" For Input As #1
        Input #1, strUpdateVer
        Input #1, strUpdateDate
    Close #1
      
    If strUpdateVer > strMyVer Then
        Text1.Text = Text1.Text & vbCrLf & cvCrLf & _
                    rsLanguage.Fields("Msg1") & strUpdateVer & vbCrLf & _
                    rsLanguage.Fields("Msg6") & " " & strUpdateDate
    Else
        Text1.Text = Text1.Text & vbCrLf & cvCrLf & _
                    rsLanguage.Fields("Msg2") & vbCrLf & _
                    rsLanguage.Fields("Msg6") & " " & strUpdateDate
        ProgressBar1.Value = 3
        btnBeginUpdate.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If

    status = "Getting updated file."
    boolTransferSuccess = GetInternetFile(Inet1, "http://www.levesen.com/downloads/OpdateMasterKid.zip", App.Path)

    If boolTransferSuccess = False Then
        ProgressBar1.Value = 3
        btnBeginUpdate.Enabled = True
        Timer2.Interval = 0
        Exit Sub
    End If
    
    ProgressBar1.Value = 3
    Timer2.Interval = 0
    
    Beep
    Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
                rsLanguage.Fields("Text1")
    Exit Sub
    
errUpdate:
    Beep
    Text1.Text = Text1.Text & vbCrLf & vbCeLf & Err.Description
    WriteErrorFile Err.Description, "frmLiveUpdate: Internet Update"
    Err.Clear
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
    Text1.Text = rsLanguage.Fields("Msg7") & " " & strMyVer & vbCrLf & _
                rsLanguage.Fields("Msg3")
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsLanguage = dbKidLang.OpenRecordset("frmLiveUpdate")
    strMyVer = App.Major & "." & App.Minor & "." & App.Revision
    status = "Idle"
    lngUpdateTime = 0
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbExclamation, "Load Form"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    Set frmLiveUpdate = Nothing
End Sub


Private Sub Timer1_Timer()
    If Inet1.StillExecuting = False Then
        StatusBar1.Panels(1).Text = "Status: Idle"
    Else
        StatusBar1.Panels(1).Text = "Status: " & status
    End If
End Sub

Private Sub Timer2_Timer()
    lngUpdateTime = lngUpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(lngUpdateTime) & " Seconds"
End Sub
