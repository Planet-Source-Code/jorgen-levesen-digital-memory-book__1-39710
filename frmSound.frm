VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSound 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sound Tracks"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   5775
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   10186
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   32768
      TabCaption(0)   =   "Birth"
      TabPicture(0)   =   "frmSound.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Infant (0-1 year)"
      TabPicture(1)   =   "frmSound.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "List2(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Childhood"
      TabPicture(2)   =   "frmSound.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).Control(1)=   "List2(2)"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   5175
         Index           =   2
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   3735
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SoundDate"
            DataSource      =   "rsSoundChild"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SoundNote"
            DataSource      =   "rsSoundChild"
            Height          =   3855
            Index           =   2
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   1200
            Width           =   3495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sound Note:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sound from date:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5100
         Index           =   2
         Left            =   -71040
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   5175
         Index           =   1
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   3735
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SoundDate"
            DataSource      =   "rsSoundBaby"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SoundNote"
            DataSource      =   "rsSoundBaby"
            Height          =   3855
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1200
            Width           =   3495
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sound Note:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sound from date:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5100
         Index           =   1
         Left            =   -71040
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4905
         Index           =   0
         Left            =   3960
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   5055
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3735
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SoundDate"
            DataSource      =   "rsSoundBirth"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   2280
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SoundNote"
            DataSource      =   "rsSoundBirth"
            Height          =   3615
            Index           =   0
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1200
            Width           =   3255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Sound Note:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sound from date:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
      End
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "&Stop Track"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton btnPlay 
      Caption         =   "&Play Track"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Data rsSoundChild 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SoundChild"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsSoundBaby 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SoundBaby"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsSoundBirth 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SoundBirth"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1920
      Picture         =   "frmSound.frx":0054
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String
Dim vi0RecordBookmark() As Variant
Dim vi1RecordBookmark() As Variant
Dim vi2RecordBookmark() As Variant
Dim rsLanguage As Recordset
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
            .BackColor = &HC0C0&
            .Tab1.BackColor = &HC0C0&
        Case 2
            .Tab1.Tab = 2
            .BackColor = &H4040&
            Tab1.BackColor = &H4040&
        Case Else
        End Select
    End With
End Sub

Public Sub DeleteSound()
    On Error Resume Next
    Select Case frmSound.Tab1.Tab
    Case 0
        rsSoundBirth.Recordset.Delete
        If SelectSoundBirth Then
            FillList20
            List2(0).ListIndex = 0
        End If
    Case 1
        rsSoundBaby.Recordset.Delete
        If SelectSoundBaby Then
            FillList21
            List2(1).ListIndex = 0
        End If
    Case 2
        rsSoundChild.Recordset.Delete
        If SelectSoundChild Then
            FillList22
            List2(2).ListIndex = 0
        End If
    Case Else
    End Select
End Sub

Public Sub NewSound()
    On Error Resume Next
    With CommonDialog1
        .Filter = "wav;mid|*.wav;*.mid"
        .ShowOpen
        sFileName = .filename
    End With
        boolNewRecord = True
    Select Case frmSound.Tab1.Tab
    Case 0
        rsSoundBirth.Recordset.Move 0
        rsSoundBirth.Recordset.AddNew
        Date1(0).SetFocus
    Case 1
        rsSoundBaby.Recordset.Move 0
        rsSoundBaby.Recordset.AddNew
        Date1(1).SetFocus
    Case 2
        rsSoundChild.Recordset.Move 0
        rsSoundChild.Recordset.AddNew
        Date1(2).SetFocus
    Case Else
    End Select
End Sub

Private Sub ReadText()
Dim strMemo As String
    'find YOUR rsLanguage text
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
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1(0).Caption
                Else
                    For i = 0 To 2
                        Label1(i).Caption = .Fields("label1")
                    Next
                End If
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2(0).Caption
                Else
                    For i = 0 To 2
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
                    .Fields("Tab11") = Tab1.Caption
                Else
                        Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("Tab12") = Tab1.Caption
                Else
                        Tab1.Caption = .Fields("Tab12")
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
        .Fields("Form") = Me.Caption
        .Fields("label1") = Label1(0).Caption
        .Fields("label2") = Label2(0).Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("btnPlay") = btnPlay.Caption
        .Fields("btnStop") = btnStop.Caption
        .Fields("Help") = strMemo
        .Update
    End With
    DBEngine.Idle dbFreeLocks
End Sub

Public Function SelectSoundChild() As Boolean
    On Error GoTo errSelectSoundChild
    Sql = "SELECT * FROM SoundChild WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsSoundChild.RecordSource = Sql
    rsSoundChild.Refresh
    rsSoundChild.Recordset.MoveFirst
    SelectSoundChild = True
    Exit Function
    
errSelectSoundChild:
    SelectSoundChild = False
    Err.Clear
End Function
Public Function SelectSoundBirth() As Boolean
    On Error GoTo errSelectSoundBirth
    Sql = "SELECT * FROM SoundBirth WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsSoundBirth.RecordSource = Sql
    rsSoundBirth.Refresh
    rsSoundBirth.Recordset.MoveFirst
    SelectSoundBirth = True
    Exit Function
    
errSelectSoundBirth:
    SelectSoundBirth = False
    Err.Clear
End Function
Public Function SelectSoundBaby() As Boolean
    On Error GoTo errSelectSoundBaby
    Sql = "SELECT * FROM SoundBaby WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsSoundBaby.RecordSource = Sql
    rsSoundBaby.Refresh
    rsSoundBaby.Recordset.MoveFirst
    SelectSoundBaby = True
    Exit Function
    
errSelectSoundBaby:
    SelectSoundBaby = False
    Err.Clear
End Function

Public Sub FillList22()
    On Error Resume Next
    List2(2).Clear
    With rsSoundChild.Recordset
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

Public Sub FillList21()
    On Error Resume Next
    List2(1).Clear
    With rsSoundBaby.Recordset
        .MoveLast
        .MoveFirst
        ReDim vi1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(1).AddItem .Fields("AutoLine")
            List2(1).ItemData(List2(1).NewIndex) = List2(1).ListCount - 1
            vi2RecordBookmark(List2(1).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub FillList20()
    On Error Resume Next
    List2(0).Clear
    With rsSoundBirth.Recordset
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
Private Sub btnPlay_Click()
Dim x As Long
    'On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        If Right(rsSoundBirth.Recordset.Fields("FilePath"), 4) = ".wav" Then
            x = PlaySound(rsSoundBirth.Recordset.Fields("FilePath"), 0, SND_ASYNC)
        ElseIf Right(rsSoundBirth.Recordset.Fields("FilePath"), 4) = ".mid" Then
            PlayMidiFile (rsSoundBirth.Recordset.Fields("FilePath"))
        End If
    Case 1
        If Right(rsSoundBaby.Recordset.Fields("FilePath"), 4) = ".wav" Then
            x = PlaySound(rsSoundBaby.Recordset.Fields("FilePath"), 0, SND_ASYNC)
        ElseIf Right(rsSoundBaby.Recordset.Fields("FilePath"), 4) = ".mid" Then
            PlayMidiFile (rsSoundBaby.Recordset.Fields("FilePath"))
        End If
    Case 2
        If Right(rsSoundChild.Recordset.Fields("FilePath"), 4) = ".wav" Then
            x = PlaySound(rsSoundChild.Recordset.Fields("FilePath"), 0, SND_ASYNC)
        ElseIf Right(rsSoundChild.Recordset.Fields("FilePath"), 4) = ".mid" Then
            PlayMidiFile (rsSoundChild.Recordset.Fields("FilePath"))
        End If
    Case Else
    End Select
End Sub

Private Sub btnStop_Click()
Dim x As Long
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        If Right(rsSoundBirth.Recordset.Fields("FilePath"), 4) = ".wav" Then
            'x = PlaySound(rsSoundBirth.Recordset.Fields("FilePath"), 0, SND_ASYNC)
        ElseIf Right(rsSoundBirth.Recordset.Fields("FilePath"), 4) = ".mid" Then
            StopMidiFile (rsSoundBirth.Recordset.Fields("FilePath"))
        End If
    Case 1
        If Right(rsSoundBaby.Recordset.Fields("FilePath"), 4) = ".wav" Then
            'x = PlaySound(rsSoundBaby.Recordset.Fields("FilePath"), 0, SND_ASYNC)
        ElseIf Right(rsSoundBaby.Recordset.Fields("FilePath"), 4) = ".mid" Then
            StopMidiFile (rsSoundBaby.Recordset.Fields("FilePath"))
        End If
    Case 2
        If Right(rsSoundChild.Recordset.Fields("FilePath"), 4) = ".wav" Then
            'x = PlaySound(rsSoundChild.Recordset.Fields("FilePath"), 0, SND_ASYNC)
        ElseIf Right(rsSoundChild.Recordset.Fields("FilePath"), 4) = ".mid" Then
            StopMidiFile (rsSoundChild.Recordset.Fields("FilePath"))
        End If
    Case Else
    End Select
End Sub

Private Sub Date1_Click(Index As Integer)
Dim UserDate As Date
    If IsDate(Date1(Index).Text) Then
        UserDate = CVDate(Date1(Index).Text)
    Else
        UserDate = Format(Now, "dd.mm.yyyy")
    End If
    If frmCalendar.GetDate(UserDate) Then
        Date1(Index).Text = UserDate
    End If
End Sub

Private Sub Date1_LostFocus(Index As Integer)
    On Error Resume Next
    If boolNewRecord Then
        Select Case Tab1.Tab
        Case 0
            With rsSoundBirth.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("FilePath") = sFileName
                .Fields("SoundDate") = CDate(Format(Date1(Index).Text, "dd.mm.yyyy"))
                .Update
                FillList20
                .Bookmark = .LastModified
            End With
        Case 1
            With rsSoundBaby.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("FilePath") = sFileName
                .Fields("SoundDate") = CDate(Format(Date1(Index).Text, "dd.mm.yyyy"))
                .Update
                FillList21
                .Bookmark = .LastModified
            End With
        Case 2
            With rsSoundChild.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("FilePath") = sFileName
                .Fields("SoundDate") = CDate(Format(Date1(Index).Text, "dd.mm.yyyy"))
                .Update
                FillList22
                .Bookmark = .LastModified
            End With
        Case Else
        End Select
        boolNewRecord = False
    End If
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsSoundBirth.Refresh
    rsSoundBaby.Refresh
    rsSoundChild.Refresh
    ReadText
    ShowAllButtons
    ShowKids
    SelectTab
    Select Case Tab1.Tab
    Case 0
        If SelectSoundBirth Then
            FillList20
            List2(0).ListIndex = 0
        End If
    Case 1
        If SelectSoundBaby Then
            FillList21
            List2(1).ListIndex = 0
        End If
    Case 2
        If SelectSoundChild Then
            FillList22
            List2(2).ListIndex = 0
        End If
    Case Else
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsSoundBirth.DatabaseName = dbKidsTxt
    rsSoundBaby.DatabaseName = dbKidsTxt
    rsSoundChild.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmSound")
    iWhichForm = 33
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmSound: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSoundBirth.Recordset.Close
    rsSoundBaby.Recordset.Close
    rsSoundChild.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmSound = Nothing
End Sub
Private Sub List2_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0  'birth
        rsSoundBirth.Recordset.Bookmark = vi0RecordBookmark(List2(0).ItemData(List2(0).ListIndex))
        Frame1(Index).Caption = rsSoundBirth.Recordset.Fields("FilePath")
    Case 1  'baby
        rsSoundBaby.Recordset.Bookmark = vi1RecordBookmark(List2(1).ItemData(List2(1).ListIndex))
        Frame1(Index).Caption = rsSoundBaby.Recordset.Fields("FilePath")
    Case 2  'childhood
        rsSoundChild.Recordset.Bookmark = vi2RecordBookmark(List2(2).ItemData(List2(2).ListIndex))
        Frame1(Index).Caption = rsSoundChild.Recordset.Fields("FilePath")
    Case Else
    End Select
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        If SelectSoundBirth Then
            FillList20
            List2(0).ListIndex = 0
        End If
        Me.BackColor = &H8000&
        Tab1.BackColor = &H8000&
    Case 1
        If SelectSoundBaby Then
            FillList21
            List2(1).ListIndex = 0
        End If
        Me.BackColor = &HC0C0&
        Tab1.BackColor = &HC0C0&
    Case 2
        If SelectSoundChild Then
            FillList22
            List2(2).ListIndex = 0
        End If
        Me.BackColor = &H4040&
        Tab1.BackColor = &H4040&
    Case Else
    End Select
End Sub

