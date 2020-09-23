VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVideo 
   BackColor       =   &H0000C0C0&
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   10050
   Begin TabDlg.SSTab Tab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Birth"
      TabPicture(0)   =   "frmVideo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "List2(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Picture1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Babtism"
      TabPicture(1)   =   "frmVideo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(1)"
      Tab(1).Control(1)=   "Label2(1)"
      Tab(1).Control(2)=   "Picture1(1)"
      Tab(1).Control(3)=   "Text1(1)"
      Tab(1).Control(4)=   "List2(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Baby (0 -1 Year)"
      TabPicture(2)   =   "frmVideo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List2(2)"
      Tab(2).Control(1)=   "Text1(2)"
      Tab(2).Control(2)=   "Picture1(2)"
      Tab(2).Control(3)=   "Label2(2)"
      Tab(2).Control(4)=   "Label1(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Childhood"
      TabPicture(3)   =   "frmVideo.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(3)"
      Tab(3).Control(1)=   "Label2(3)"
      Tab(3).Control(2)=   "Picture1(3)"
      Tab(3).Control(3)=   "Text1(3)"
      Tab(3).Control(4)=   "List2(3)"
      Tab(3).ControlCount=   5
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4905
         Index           =   3
         Left            =   -74760
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "VideoNote"
         DataSource      =   "rsVideoChild"
         Height          =   1095
         Index           =   3
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   5880
         Width           =   7905
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   5055
         Index           =   3
         Left            =   -73200
         ScaleHeight     =   5025
         ScaleWidth      =   6345
         TabIndex        =   18
         Top             =   600
         Width           =   6375
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4710
         Index           =   2
         Left            =   -74760
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "VideoNote"
         DataSource      =   "rsVideoBaby"
         Height          =   1095
         Index           =   2
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   5880
         Width           =   7875
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   5055
         Index           =   2
         Left            =   -73320
         ScaleHeight     =   5025
         ScaleWidth      =   6465
         TabIndex        =   13
         Top             =   600
         Width           =   6495
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4710
         Index           =   1
         Left            =   -74760
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "VideoNote"
         DataSource      =   "rsVideoBaptism"
         Height          =   1095
         Index           =   1
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   5880
         Width           =   8025
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   5175
         Index           =   1
         Left            =   -73320
         ScaleHeight     =   5145
         ScaleWidth      =   6465
         TabIndex        =   8
         Top             =   600
         Width           =   6495
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   5145
         Index           =   0
         Left            =   1680
         ScaleHeight     =   5115
         ScaleWidth      =   6555
         TabIndex        =   5
         Top             =   600
         Width           =   6585
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "VideoNote"
         DataSource      =   "rsVideoBirth"
         Height          =   1080
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   5880
         Width           =   8040
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4710
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "Video Note:"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   22
         Top             =   5640
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Video No."
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Video Note:"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   17
         Top             =   5640
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Video No."
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Video Note:"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   12
         Top             =   5640
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Video No."
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Video No."
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Video Note:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   5640
         Width           =   1740
      End
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "&Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      Picture         =   "frmVideo.frx":0070
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Play the video"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "&Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      Picture         =   "frmVideo.frx":037A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Stop the Video"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Data rsVideoBirth 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VideoBirth"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Data rsVideoBaptism 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VideoBaptism"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Data rsVideoBaby 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VideoBaby"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Data rsVideoChild 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VideoChild"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   8880
      Top             =   3840
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8760
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vi0RecordBookmark() As Variant
Dim vi1RecordBookmark() As Variant
Dim vi2RecordBookmark() As Variant
Dim vi3RecordBookmark() As Variant
Dim result As String
Dim rsLanguage As Recordset
Public Sub NewChildVideo()
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        SelectVideoBirth
        FillList20
    Case 1
        SelectVideoBaptism
        FillList21
    Case 2
        SelectVideoBaby
        FillList22
    Case 3
        SelectVideoChild
        FillList23
    Case Else
    End Select
End Sub

Public Sub NewVideo()
    On Error Resume Next
    With CommonDialog1
        .Filter = "mpeg,mpg,mov,avi,mpe,mpv,m1v,vbs,dat|*.mpeg;*.mov;*.avi;*.mpg;*.mpe;*.mpv;*.m1v;*.vbs;*.dat"
        .ShowOpen
        sFileName = .filename
    End With
        boolNewRecord = True
    Select Case frmVideo.Tab1.Tab
    Case 0
        rsVideoBirth.Recordset.AddNew
        Text1(0).SetFocus
    Case 1
        rsVideoBaptism.Recordset.AddNew
        Text1(1).SetFocus
    Case 2
        rsVideoBaby.Recordset.AddNew
        Text1(2).SetFocus
    Case 3
        rsVideoChild.Recordset.AddNew
        Text1(3).SetFocus
    Case Else
    End Select
End Sub

Public Sub DeleteVideo()
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        rsVideoBirth.Recordset.Delete
        FillList20
    Case 1
        rsVideoBaptism.Recordset.Delete
        FillList21
    Case 2
        rsVideoBaby.Recordset.Delete
        FillList22
    Case 3
        rsVideoChild.Recordset.Delete
        FillList23
    Case Else
    End Select
End Sub

Public Sub FillList20()
    On Error Resume Next
    List2(0).Clear
    With rsVideoBirth.Recordset
        .MoveLast
        .MoveFirst
        ReDim vi0RecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(0).AddItem .Fields("AutoLine")
            List2(0).ItemData(List2(0).NewIndex) = List2(0).ListCount - 1
            vi0RecordBookmark(List2(0).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Public Sub FillList21()
    On Error Resume Next
    List2(1).Clear
    With rsVideoBaptism.Recordset
        .MoveLast
        .MoveFirst
        ReDim vi1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(1).AddItem .Fields("AutoLine")
            List2(1).ItemData(List2(1).NewIndex) = List2(1).ListCount - 1
            vi1RecordBookmark(List2(1).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub FillList23()
    On Error Resume Next
    List2(3).Clear
    With rsVideoChild.Recordset
        .MoveLast
        .MoveFirst
        ReDim vi3RecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(3).AddItem .Fields("AutoLine")
            List2(3).ItemData(List2(3).NewIndex) = List2(3).ListCount - 1
            vi3RecordBookmark(List2(3).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub FillList22()
    On Error Resume Next
    List2(2).Clear
    With rsVideoBaby.Recordset
        .MoveLast
        .MoveFirst
        ReDim vi2RecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(2).AddItem .Fields("AutoLine")
            List2(2).ItemData(List2(2).NewIndex) = List2(2).ListCount - 1
            vi2RecordBookmark(List2(2).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub SelectTab()
    On Error Resume Next
    With Me
        Select Case iTab
        Case 0
            .Tab1.Tab = 0
            .BackColor = &H8000&
            .Tab1.BackColor = &H8000&
        Case 1
            .Tab1.Tab = 1
            .BackColor = &H400040
            .Tab1.BackColor = &H400040
        Case 2
            .Tab1.Tab = 2
            .BackColor = &HC0C0&
            .Tab1.BackColor = &HC0C0&
        Case 3
            .Tab1.Tab = 3
            .BackColor = &H4040&
            .Tab1.BackColor = &H4040&
        Case Else
        End Select
    End With
End Sub

Public Sub SelectVideoBaby()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM VideoBaby WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsVideoBaby.RecordSource = Sql
    rsVideoBaby.Refresh
End Sub
Public Sub SelectVideoBaptism()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM VideoBaptism WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsVideoBaptism.RecordSource = Sql
    rsVideoBaptism.Refresh
End Sub
Public Sub SelectVideoBirth()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM VideoBirth WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsVideoBirth.RecordSource = Sql
    rsVideoBirth.Refresh
End Sub
Public Sub SelectVideoChild()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM VideoChild WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsVideoChild.RecordSource = Sql
    rsVideoChild.Refresh
End Sub
Function StopPlay()
Dim result As String
    On Error Resume Next
    result = CloseMPEG
End Function
Private Sub PlayVideo()
Dim result As String
Dim filename As String
Dim typeDevice As String

    result = CloseMPEG()
    Timer1.Enabled = False 'Disable Timer1
    
    On Error GoTo errPlayVideo

    If Right(sFileName, 4) = ".avi" Then 'if the movie is avi then select type
        typeDevice = "AviVideo"
    ElseIf Right(sFileName, 4) = ".rmi" Or Right(sFileName, 4) = ".mid" Then
        typeDevice = "sequencer" ' select this type for midi and rmi files
    Else 'else this mean it mpg,mp3,mp2,mp1,wav,,,etc then we will choose "MpegVideo" type
        typeDevice = "MPEGVideo"
    End If

    Select Case Tab1.Tab
    Case 0
        result = OpenMPEG(Picture1(0).hWnd, sFileName, typeDevice) 'call now function openMPEG
    Case 1
        result = OpenMPEG(Picture1(1).hWnd, sFileName, typeDevice) 'call now function openMPEG
    Case 2
        result = OpenMPEG(Picture1(2).hWnd, sFileName, typeDevice) 'call now function openMPEG
    Case 3
        result = OpenMPEG(Picture1(3).hWnd, sFileName, typeDevice) 'call now function openMPEG
    Case Else
    End Select

    result = PlayMPEG("", "")
    
    result = PutMPEG(0, 0, 0, 0)
    Timer1.Enabled = True
    Exit Sub
    
errPlayVideo:
    MsgBox Err.Description, vbCritical, "Play Video"
    WriteErrorFile Err.Description, "frmVideo:  Play Video"
    Resume errPlayVideo2
errPlayVideo2:
End Sub

Private Sub ReadText()
Dim strMemo As String
    'find YOUR rsLanguage text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1(0).Caption
                Else
                    For i = 0 To 3
                        Label1(i).Caption = .Fields("label1")
                    Next
                End If
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2(0).Caption
                Else
                    For i = 0 To 3
                        Label2(i).Caption = .Fields("label2")
                    Next
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab12")
                End If
                Tab1.Tab = 3
                If IsNull(.Fields("Tab13")) Then
                    .Fields("Tab13") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab13")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("btnPlay")) Then
                    .Fields("btnPlay") = btnPlay.Caption
                Else
                        btnPlay.Caption = .Fields("btnPlay")
                End If
                If IsNull(.Fields("btnStop")) Then
                    .Fields("btnStop") = btnStop.Caption
                Else
                        btnStop.Caption = .Fields("btnStop")
                End If
                .Update
                DBEngine.Idle dbFreeLocks
                Me.MousePointer = Default
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = "ENG" Then
                If Not IsNull(.Fields("Help")) Then
                    strMemo = .Fields("Help")
                Else
                    strMemo = " "
                End If
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = FileExt
        .Fields("label1") = Label1(0).Caption
        .Fields("label2") = Label2(0).Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("btnPlay") = btnPlay.Caption
        .Fields("btnStop") = btnStop.Caption
        .Fields("Help") = strMemo
        .Update
    End With
    DBEngine.Idle dbFreeLocks
End Sub
Private Sub btnPlay_Click()
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        sFileName = rsVideoBirth.Recordset.Fields("FilePath")
    Case 1
        sFileName = rsVideoBaptism.Recordset.Fields("FilePath")
    Case 2
        sFileName = rsVideoBaby.Recordset.Fields("FilePath")
    Case 3
        sFileName = rsVideoChild.Recordset.Fields("FilePath")
    Case Else
    End Select
    PlayVideo
End Sub

Private Sub btnStop_Click()
    On Error Resume Next
    StopPlay
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsVideoBirth.Refresh
    rsVideoBaptism.Refresh
    rsVideoBaby.Refresh
    rsVideoChild.Refresh
    ReadText
    ShowAllButtons
    ShowKids
    SelectTab
    Select Case Tab1.Tab
    Case 0
        SelectVideoBirth
        FillList20
        rsVideoBirth.Recordset.Bookmark = vi0RecordBookmark(List2(0).ItemData(List2(0).ListIndex = 0))
    Case 1
        SelectVideoBaptism
        FillList21
        rsVideoBaptism.Recordset.Bookmark = vi1RecordBookmark(List2(1).ItemData(List2(1).ListIndex = 0))
    Case 2
        SelectVideoBaby
        FillList22
        rsVideoBaby.Recordset.Bookmark = vi2RecordBookmark(List2(2).ItemData(List2(2).ListIndex = 0))
    Case 3
        SelectVideoChild
        FillList23
        rsVideoChild.Recordset.Bookmark = vi3RecordBookmark(List2(3).ItemData(List2(3).ListIndex = 0))
    Case Else
    End Select
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsVideoBirth.DatabaseName = dbKidsTxt
    rsVideoBaptism.DatabaseName = dbKidsTxt
    rsVideoBaby.DatabaseName = dbKidsTxt
    rsVideoChild.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmVideo")
    iWhichForm = 30
    
    If Not GetDefaultDevice("MPEGVideo") = "mciqtz.drv" Then
        SetDefaultDevice "MPEGVideo", "mciqtz.drv"
    End If
    If Not GetDefaultDevice("sequencer") = "mciseq.drv" Then
        SetDefaultDevice "sequencer", "mciseq.drv"
    End If
    If Not GetDefaultDevice("avivideo") = "mciavi.drv" Then
        SetDefaultDevice "avivideo", "mciavi.drv"
    End If
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmVideo: Load Form"
    Err.Clear
    Unload Me
End Sub
Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsVideoBirth.Recordset.Close
    rsVideoBaptism.Recordset.Close
    rsVideoBaby.Recordset.Close
    rsVideoChild.Recordset.Close
    rsLanguage.Close
    CloseMPEG
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmVideo = Nothing
End Sub
Private Sub List2_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0  'birth
        Picture1(0).Picture = LoadPicture()
        rsVideoBirth.Recordset.Bookmark = vi0RecordBookmark(List2(0).ItemData(List2(0).ListIndex))
    Case 1  'baptism
        Picture1(1).Picture = LoadPicture()
        rsVideoBaptism.Recordset.Bookmark = vi1RecordBookmark(List2(1).ItemData(List2(1).ListIndex))
    Case 2  'baby
        Picture1(2).Picture = LoadPicture()
        rsVideoBaby.Recordset.Bookmark = vi2RecordBookmark(List2(2).ItemData(List2(2).ListIndex))
    Case 3  'childhood
        Picture1(3).Picture = LoadPicture()
        rsVideoChild.Recordset.Bookmark = vi3RecordBookmark(List2(3).ItemData(List2(3).ListIndex))
    Case Else
    End Select
End Sub
Private Sub Tab1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    On Error Resume Next
    StopPlay
    Select Case NewTab
    Case 0
        SelectVideoBirth
        FillList20
        Me.BackColor = &H8000&
        Tab1.BackColor = &H8000&
    Case 1
        SelectVideoBaptism
        FillList21
        Me.BackColor = &H400040
        Tab1.BackColor = &H400040
    Case 2
        SelectVideoBaby
        FillList22
        Me.BackColor = &HC0C0&
        Tab1.BackColor = &HC0C0&
    Case 3
        SelectVideoChild
        FillList23
        Me.BackColor = &H4040&
        Tab1.BackColor = &H4040&
    Case Else
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error Resume Next
    If boolNewRecord Then
        Select Case Tab1.Tab
        Case 0
            With rsVideoBirth.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("FilePath") = sFileName
                .Fields("VideoNote") = Trim(Text1(Index).Text)
                .Update
                FillList20
                .Bookmark = .LastModified
            End With
        Case 1
            With rsVideoBaptism.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("FilePath") = sFileName
                .Fields("VideoNote") = Trim(Text1(Index).Text)
                .Update
                FillList21
                .Bookmark = .LastModified
            End With
        Case 2
            With rsVideoBaby.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("FilePath") = sFileName
                .Fields("VideoNote") = Trim(Text1(Index).Text)
                .Update
                FillList22
                .Bookmark = .LastModified
            End With
        Case 3
            With rsVideoChild.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("FilePath") = sFileName
                .Fields("VideoNote") = Trim(Text1(Index).Text)
                .Update
                FillList23
                .Bookmark = .LastModified
            End With
        Case Else
        End Select
        boolNewRecord = False
    End If
End Sub
