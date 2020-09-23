VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHospital 
   BackColor       =   &H00008000&
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   7575
   Begin TabDlg.SSTab Tab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   32768
      TabCaption(0)   =   "Notes"
      TabPicture(0)   =   "frmHospital.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Acquaintance"
      TabPicture(1)   =   "frmHospital.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Cmd1"
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Leaving"
      TabPicture(2)   =   "frmHospital.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(3)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Home"
      TabPicture(3)   =   "frmHospital.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(2)"
      Tab(3).Control(1)=   "rsHospitalHome"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   6855
         Index           =   3
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   6975
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "LeftHospitalHour"
            DataSource      =   "rsLeaving"
            Height          =   375
            Index           =   5
            Left            =   4680
            TabIndex        =   34
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "LeftHospitalDate"
            DataSource      =   "rsLeaving"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4680
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.Data rsLeaving 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "HospitalLeaving"
            Top             =   120
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "OurCar"
            DataSource      =   "rsLeaving"
            Height          =   2175
            Index           =   3
            Left            =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   4560
            Width           =   2490
         End
         Begin VB.CommandButton btnPaste 
            Height          =   495
            Left            =   6210
            Picture         =   "frmHospital.frx":0070
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Paste picture"
            Top             =   4320
            Width           =   585
         End
         Begin VB.CommandButton btnCopy 
            Height          =   495
            Left            =   6210
            Picture         =   "frmHospital.frx":0732
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Copy picture"
            Top             =   5280
            Width           =   585
         End
         Begin VB.CommandButton btnFromFile 
            Height          =   495
            Left            =   6210
            Picture         =   "frmHospital.frx":0DF4
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Read picture from file"
            Top             =   4800
            Width           =   585
         End
         Begin VB.CommandButton btnDelete 
            Height          =   495
            Left            =   6210
            Picture         =   "frmHospital.frx":14B6
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Delete picture"
            Top             =   6240
            Width           =   585
         End
         Begin VB.CommandButton btnScan 
            Height          =   495
            Index           =   0
            Left            =   6210
            Picture         =   "frmHospital.frx":1600
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Scan a picture"
            Top             =   5760
            Width           =   585
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1665
            ScaleHeight     =   345
            ScaleWidth      =   315
            TabIndex        =   23
            Top             =   960
            Visible         =   0   'False
            Width           =   345
         End
         Begin MSMask.MaskEdBox DTPicker1 
            Height          =   375
            Left            =   4680
            TabIndex        =   1
            Top             =   840
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "DrivenBy"
            DataSource      =   "rsLeaving"
            Height          =   2535
            Index           =   2
            Left            =   255
            TabIndex        =   2
            Top             =   1560
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   4471
            _Version        =   393217
            BackColor       =   16777152
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmHospital.frx":174A
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date left Hospital:"
            Height          =   330
            Index           =   1
            Left            =   2550
            TabIndex        =   33
            Top             =   360
            Width           =   1965
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Time left Hospital:"
            Height          =   330
            Index           =   2
            Left            =   2550
            TabIndex        =   32
            Top             =   840
            Width           =   1965
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Who drove us:"
            Height          =   330
            Index           =   3
            Left            =   255
            TabIndex        =   31
            Top             =   1320
            Width           =   1965
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Our car was:"
            Height          =   285
            Index           =   4
            Left            =   255
            TabIndex        =   30
            Top             =   4200
            Width           =   1965
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "OurCarPic"
            DataSource      =   "rsLeaving"
            Height          =   2175
            Left            =   3165
            Stretch         =   -1  'True
            Top             =   4560
            Width           =   2925
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Picture of our car:"
            Height          =   285
            Index           =   5
            Left            =   3165
            TabIndex        =   29
            Top             =   4200
            Width           =   1950
         End
         Begin VB.Image Image3 
            Height          =   855
            Left            =   255
            Picture         =   "frmHospital.frx":181F
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1770
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Acquaintance Names, Adress and notes"
         Height          =   6615
         Index           =   1
         Left            =   -74880
         TabIndex        =   20
         Top             =   720
         Width           =   6975
         Begin VB.Data rsAcquaintance 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   3825
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "HospitalAcquaintance"
            Top             =   0
            Visible         =   0   'False
            Width           =   1350
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "Note"
            DataSource      =   "rsAcquaintance"
            Height          =   6135
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   10821
            _Version        =   393217
            BackColor       =   16777152
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmHospital.frx":37C2
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Notes"
         Height          =   6735
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   6975
         Begin VB.Data rsHospitalNotes 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   2040
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "BirthDiary"
            Top             =   120
            Visible         =   0   'False
            Width           =   1410
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "DiaryNotes"
            DataSource      =   "rsHospitalNotes"
            Height          =   6135
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   10821
            _Version        =   393217
            BackColor       =   16777152
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmHospital.frx":3897
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "At home from the Hospital"
         Height          =   6735
         Index           =   2
         Left            =   -74880
         TabIndex        =   11
         Top             =   600
         Width           =   6975
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstNightSleptAt"
            DataSource      =   "rsHospitalHome"
            Height          =   375
            Index           =   7
            Left            =   2520
            TabIndex        =   36
            Top             =   3720
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CameHomeTime"
            DataSource      =   "rsHospitalHome"
            Height          =   375
            Index           =   6
            Left            =   2520
            TabIndex        =   35
            Top             =   1800
            Width           =   735
         End
         Begin MSMask.MaskEdBox DTPicker3 
            Height          =   375
            Left            =   2520
            TabIndex        =   7
            Top             =   3720
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox DTPicker2 
            Height          =   375
            Left            =   2520
            TabIndex        =   5
            Top             =   1800
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "OurAddress"
            DataSource      =   "rsHospitalHome"
            Height          =   1245
            Index           =   0
            Left            =   2520
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WasMetBy"
            DataSource      =   "rsHospitalHome"
            Height          =   1230
            Index           =   1
            Left            =   2505
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   2325
            Width           =   4215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WokeNoOfTimes"
            DataSource      =   "rsHospitalHome"
            Height          =   300
            Index           =   2
            Left            =   6105
            TabIndex        =   8
            Top             =   3780
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstToVisit"
            DataSource      =   "rsHospitalHome"
            Height          =   2460
            Index           =   4
            Left            =   2505
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   4155
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Our Address:"
            Height          =   300
            Index           =   0
            Left            =   360
            TabIndex        =   17
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Came home at:"
            Height          =   285
            Index           =   6
            Left            =   360
            TabIndex        =   16
            Top             =   1830
            Width           =   1995
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Was met by:"
            Height          =   285
            Index           =   7
            Left            =   360
            TabIndex        =   15
            Top             =   2445
            Width           =   1995
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Slept first night at:"
            Height          =   300
            Index           =   8
            Left            =   360
            TabIndex        =   14
            Top             =   3720
            Width           =   1995
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Woke no. of times:"
            Height          =   420
            Index           =   9
            Left            =   3735
            TabIndex        =   13
            Top             =   3660
            Width           =   2220
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "First persons to visit me at home:"
            Height          =   1140
            Index           =   10
            Left            =   360
            TabIndex        =   12
            Top             =   4275
            Width           =   1995
         End
      End
      Begin VB.Data rsHospitalHome 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "HospitalHome"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSComDlg.CommonDialog Cmd1 
         Left            =   -74880
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Public Sub DeleteHospital()
    On Error Resume Next
    SelectHospitalNotes
    rsHospitalNotes.Recordset.Delete
    SelectHospitalAcquaintance
    rsAcquaintance.Recordset.Delete
    SelectHospitalLeaving
    rsLeaving.Recordset.Delete
    SelectHospitalHome
    rsHospitalHome.Recordset.Delete
    MDIMasterKid.Toolbar1.Buttons(6).Enabled = True
End Sub
Public Sub NewHospital()
    On Error Resume Next
    rsHospitalNotes.Recordset.Move 0
    rsHospitalNotes.Recordset.AddNew
    RichTextBox1(0).SetFocus
    boolNewRecord = True
End Sub
Public Sub SelectHospitalAcquaintance()
Dim Sql As String
    Sql = "SELECT * FROM HospitalAcquaintance WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsAcquaintance.RecordSource = Sql
    rsAcquaintance.Refresh
End Sub


Public Sub SelectHospitalHome()
Dim Sql As String
    Sql = "SELECT * FROM HospitalHome WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsHospitalHome.RecordSource = Sql
    rsHospitalHome.Refresh
End Sub

Public Sub SelectHospitalLeaving()
Dim Sql As String
    Sql = "SELECT * FROM HospitalLeaving WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsLeaving.RecordSource = Sql
    rsLeaving.Refresh
End Sub

Public Function SelectHospitalNotes() As Boolean
Dim Sql As String
    On Error GoTo errSelectHospitalNotes
    Sql = "SELECT * FROM BirthDiary WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsHospitalNotes.RecordSource = Sql
    rsHospitalNotes.Refresh
    rsHospitalNotes.Recordset.MoveFirst
    SelectHospitalNotes = True
    Exit Function
    
errSelectHospitalNotes:
    SelectHospitalNotes = False
    Err.Clear
End Function


Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                For i = 0 To 10
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
                If IsNull(.Fields("Frame1(0)")) Then
                    .Fields("Frame1(0)") = Frame1(0).Caption
                Else
                    Frame1(0).Caption = .Fields("Frame1(0)")
                End If
                If IsNull(.Fields("Frame1(1)")) Then
                    .Fields("Frame1(1)") = Frame1(1).Caption
                Else
                    Frame1(1).Caption = .Fields("Frame1(1)")
                End If
                If IsNull(.Fields("Frame1(2)")) Then
                    .Fields("Frame1(2)") = Frame1(2).Caption
                Else
                    Frame1(2).Caption = .Fields("Frame1(2)")
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
                Tab1.Tab = 3
                If IsNull(.Fields("Tab13")) Then
                    .Fields("Tab13") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab13")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("btnPaste")) Then
                    .Fields("btnPaste") = btnPaste.ToolTipText
                Else
                    btnPaste.ToolTipText = .Fields("btnPaste")
                End If
                If IsNull(.Fields("btnFromFile")) Then
                    .Fields("btnFromFile") = btnFromFile.ToolTipText
                Else
                    btnFromFile.ToolTipText = .Fields("btnFromFile")
                End If
                If IsNull(.Fields("btnCopy")) Then
                    .Fields("btnCopy") = btnCopy.ToolTipText
                Else
                    btnCopy.ToolTipText = .Fields("btnCopy")
                End If
                If IsNull(.Fields("btnScan")) Then
                    .Fields("btnScan") = btnScan(0).ToolTipText
                Else
                    btnScan(0).ToolTipText = .Fields("btnScan")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete.ToolTipText
                Else
                    btnDelete.ToolTipText = .Fields("btnDelete")
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
            
        .AddNew
        .Fields("Language") = FileExt
        For i = 0 To 10
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("Frame1(0)") = Frame1(0).Caption
        .Fields("Frame1(1)") = Frame1(1).Caption
        .Fields("Frame1(2)") = Frame1(2).Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("btnPaste") = btnPaste.ToolTipText
        .Fields("btnFromFile") = btnFromFile.ToolTipText
        .Fields("btnCopy") = btnCopy.ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        .Fields("btnDelete") = btnDelete.ToolTipText
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Public Sub WriteHospital()
    'On Error Resume Next
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    sHeader = rsLanguage.Fields("FormName1")
    cPrint.pStartDoc
    Call PrintFront
    
    'hospital notes
    cPrint.FontBold = True
    cPrint.pPrint Frame1(0).Caption, 1
    cPrint.FontBold = False
    If Len(RichTextBox1(0).Text) <> 0 Then
        cPrint.pMultiline RichTextBox1(0).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    
    'hospital Acquaintance
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Frame1(1).Caption, 1
    cPrint.FontBold = False
    If Len(RichTextBox1(1).Text) <> 0 Then
        cPrint.pMultiline RichTextBox1(1).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    
    'leaving hospital
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint
    cPrint.pPrint Label1(1).Caption, 1, True   'date left hospital
    If IsDate(CDate(Date1.Text)) Then
        cPrint.pPrint Format(CDate(Date1.Text), "dd.mm.yyyy"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(2).Caption, 1, True   'time left hospital
    If Len(Text1(5).Text) <> 0 Then
        cPrint.pPrint Format(Text1(5).Text, "hh:mm"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(3).Caption, 1, True   'who drove us home
    If Len(RichTextBox1(2).Text) <> 0 Then
        cPrint.pMultiline RichTextBox1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(4).Caption, 1, True   'our car was a..
    If Len(Text1(3).Text) <> 0 Then
        cPrint.pMultiline Text1(3).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    If Not IsNull(rsLeaving.Recordset.Fields("OurCarPic")) Then
        cPrint.pPrintPicture Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, Image1.Picture.Height, False, True
        cPrint.CurrentY = cPrint.CurrentY + Image1.Picture.Height
    End If
    
    'home from hospital
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint
    cPrint.pPrint Label1(6).Caption, 1, True   'time we came home
    If Len(Text1(6).Text) <> 0 Then
        cPrint.pPrint Format(Text1(6).Text, "hh:mm"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(0).Caption, 1, True   'our address
    If Len(Text1(0).Text) <> 0 Then
        cPrint.pMultiline Text1(0).Text, 3.5, , cPrint.GetPaperWidth - 1.2, False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Label1(7).Caption, 1   'was met at home by...
    cPrint.FontBold = False
    If Len(Text1(1).Text) <> 0 Then
        cPrint.pMultiline Text1(1).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    cPrint.pPrint Label1(8).Caption, 1, True   'time slept first night home
    If Len(Text1(7).Text) <> 0 Then
        cPrint.pPrint Text1(7).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(9).Caption, 1, True   'woke no.of times
    If Len(Text1(2).Text) <> 0 Then
        cPrint.pPrint Text1(2).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Label1(10).Caption, 1  'first person to visit me at home
    cPrint.FontBold = False
    If Len(Text1(4).Text) Then
        cPrint.pMultiline Text1(4).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint , 3.5
    End If
    
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub
Public Sub WriteHospitalWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName1"))
    With wdApp
        .Selection.Tables(1).AllowAutoFit = False
        'hospital notes
        .Selection.TypeText Text:=Frame1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox1(0).TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        'hospital Acquaintance
        .Selection.TypeText Text:=Frame1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox1(1).TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        'leaving hospital
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("FormName3")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(Date1.Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(DTPicker1.Text), "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox1(2).TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(3).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(5).Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData Image1.Picture, vbCFBitmap
        .Selection.Paste
        'home from hospital
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("FormName4")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(DTPicker2.Text), "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(8).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(DTPicker3.Text), "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(9).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(10).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(4).Text
    End With
    Set wdApp = Nothing
End Sub

Private Sub btnCopy_Click()
    On Error Resume Next
    Clipboard.SetData Image1.Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click()
    On Error Resume Next
    Set Image1.Picture = LoadPicture()
End Sub

Private Sub btnFromFile_Click()
    On Error Resume Next
        With Cmd1
            .filename = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "Pictures (*.bmp; *.pcx;*.jpg;*.jpeg;*.gif)|*.bmp;*.pcx;*.jpg;*.jpeg;*.gif"
            .FilterIndex = 1
            .Action = 1
        End With
        Set Image1.Picture = LoadPicture(Cmd1.filename)
End Sub

Private Sub btnPaste_Click()
        On Error Resume Next
        Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub btnScan_Click(Index As Integer)
Dim ret As Long, t As Single
    On Error Resume Next
    ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Date1_Click()
Dim UserDate As Date
    If IsDate(Date1.Text) Then
        UserDate = CVDate(Date1.Text)
    Else
        UserDate = Format(Now, "dd.mm.yyyy")
    End If
    If frmCalendar.GetDate(UserDate) Then
        Date1.Text = UserDate
    End If
End Sub
Private Sub DTPicker1_LostFocus()
    Text1(5).Text = Format(DTPicker1.Text, "hh:mm")
    DTPicker1.Visible = False
    Text1(5).Visible = True
End Sub

Private Sub DTPicker2_LostFocus()
    Text1(6).Text = Format(DTPicker2.Text, "hh:mm")
    DTPicker2.Visible = False
    Text1(6).Visible = True
End Sub
Private Sub DTPicker3_LostFocus()
    Text1(7).Text = Format(DTPicker3.Text, "hh:mm")
    DTPicker3.Visible = False
    Text1(7).Visible = True
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsHospitalNotes.Refresh
    rsAcquaintance.Refresh
    rsLeaving.Refresh
    rsHospitalHome.Refresh
    ShowText
    ShowAllButtons
    ShowKids
    If SelectHospitalNotes Then
        MDIMasterKid.Toolbar1.Buttons(6).Enabled = False
    Else
        MDIMasterKid.Toolbar1.Buttons(6).Enabled = True
    End If
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsHospitalNotes.DatabaseName = dbKidsTxt
    rsAcquaintance.DatabaseName = dbKidsTxt
    rsLeaving.DatabaseName = dbKidsTxt
    rsHospitalHome.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmHospital")
    iWhichForm = 18
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmHospital:  Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsHospitalNotes.Recordset.Close
    rsAcquaintance.Recordset.Close
    rsLeaving.Recordset.Close
    rsHospitalHome.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmHospital = Nothing
End Sub

Private Sub RichTextBox1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    On Error Resume Next
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1(Index).SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox1_LostFocus(Index As Integer)
    'On Error Resume Next
    If boolNewRecord Then
        Select Case Index
        Case 0
            'make all 4 recordset
            With rsHospitalNotes.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("DiaryNotes") = Trim(RichTextBox1(0).TextRTF)
                .Update
                .Bookmark = .LastModified
            End With
            With rsAcquaintance.Recordset
                .AddNew
                .Fields("ChildNo") = glChildNo
                .Update
                .Bookmark = .LastModified
            End With
            With rsLeaving.Recordset
                .AddNew
                .Fields("ChildNo") = glChildNo
                .Update
                .Bookmark = .LastModified
            End With
            With rsHospitalHome.Recordset
                .AddNew
                .Fields("ChildNo") = glChildNo
                .Update
                .Bookmark = .LastModified
            End With
        Case Else
        End Select
        boolNewRecord = False
    End If
End Sub

Private Sub RichTextBox1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox1_SelChange(Index As Integer)
    On Error Resume Next
    Call RichTextSelChange(frmHospital.RichTextBox1(Index))
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab1.Tab
        Case 0
            SelectHospitalNotes
        Case 1
            SelectHospitalAcquaintance
        Case 2
            SelectHospitalLeaving
        Case 3
            SelectHospitalHome
    Case Else
    End Select
End Sub
Private Sub Text1_Click(Index As Integer)
    Select Case Index
    Case 5
        Text1(5).Visible = False
        DTPicker1.Visible = True
    Case 6
        Text1(6).Visible = False
        DTPicker2.Visible = True
    Case 7
        Text1(7).Visible = False
        DTPicker3.Visible = True
    Case Else
    End Select
End Sub
