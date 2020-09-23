VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFirstTimes 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First Time ...."
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   11
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   49344
      TabCaption(0)   =   "Sleep"
      TabPicture(0)   =   "frmFirstTimes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sat"
      TabPicture(1)   =   "frmFirstTimes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Crawl"
      TabPicture(2)   =   "frmFirstTimes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Stood"
      TabPicture(3)   =   "frmFirstTimes.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Clap"
      TabPicture(4)   =   "frmFirstTimes.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Vave"
      TabPicture(5)   =   "frmFirstTimes.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Speek"
      TabPicture(6)   =   "frmFirstTimes.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame1(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Walk"
      TabPicture(7)   =   "frmFirstTimes.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame1(7)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Travel"
      TabPicture(8)   =   "frmFirstTimes.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame1(8)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Haircut"
      TabPicture(9)   =   "frmFirstTimes.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame1(9)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Christmas"
      TabPicture(10)  =   "frmFirstTimes.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame1(10)"
      Tab(10).ControlCount=   1
      Begin VB.Frame Frame1 
         Caption         =   "The first Christmas"
         Height          =   5175
         Index           =   10
         Left            =   -74760
         TabIndex        =   41
         Top             =   840
         Width           =   6375
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstChristmasYear"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   10
            Left            =   4920
            TabIndex        =   55
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ChristmasNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   10
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   1080
            Width           =   6105
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   10
            Left            =   3600
            TabIndex        =   44
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The first time the child had a haircut "
         Height          =   5055
         Index           =   9
         Left            =   -74760
         TabIndex        =   37
         Top             =   840
         Width           =   6375
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeHaircut"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   9
            Left            =   5040
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "HaircutNote"
            DataSource      =   "rsFirstTime"
            Height          =   4095
            Index           =   9
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   840
            Width           =   6135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   9
            Left            =   3720
            TabIndex        =   40
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The first time the child travled"
         Height          =   5175
         Index           =   8
         Left            =   -74880
         TabIndex        =   33
         Top             =   840
         Width           =   6615
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstimeTravel"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   8
            Left            =   5160
            TabIndex        =   53
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "TravelNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   8
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   1080
            Width           =   6375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   8
            Left            =   3840
            TabIndex        =   36
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The very firsttime the child took its first steps"
         Height          =   5055
         Index           =   7
         Left            =   -74760
         TabIndex        =   29
         Top             =   840
         Width           =   6375
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeWalk"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   7
            Left            =   5040
            TabIndex        =   52
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WalkNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   7
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   960
            Width           =   6135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   7
            Left            =   3720
            TabIndex        =   32
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The very first time the child spoked"
         Height          =   5055
         Index           =   6
         Left            =   -74760
         TabIndex        =   25
         Top             =   840
         Width           =   6375
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstimeSpeak"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   6
            Left            =   5040
            TabIndex        =   51
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SpeakNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   6
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   960
            Width           =   6135
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   6
            Left            =   3720
            TabIndex        =   28
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The first time the child vaved to/at you"
         Height          =   5175
         Index           =   5
         Left            =   -74880
         TabIndex        =   21
         Top             =   840
         Width           =   6615
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeVave"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   5
            Left            =   5160
            TabIndex        =   50
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "VaveNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   5
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1080
            Width           =   6375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   5
            Left            =   3960
            TabIndex        =   24
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The first time thechild claped its hands"
         Height          =   5175
         Index           =   4
         Left            =   -74880
         TabIndex        =   17
         Top             =   840
         Width           =   6615
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeClap"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   4
            Left            =   5160
            TabIndex        =   49
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ClapNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   4
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   1080
            Width           =   6375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   4
            Left            =   3960
            TabIndex        =   20
            Top             =   360
            Width           =   1110
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The very first time the child stood by itself"
         Height          =   5175
         Index           =   3
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   6615
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeStood"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   3
            Left            =   5160
            TabIndex        =   48
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "StoodNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   3
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1080
            Width           =   6375
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   3
            Left            =   3840
            TabIndex        =   15
            Top             =   360
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The first time the child crawled "
         Height          =   5175
         Index           =   2
         Left            =   -74880
         TabIndex        =   9
         Top             =   720
         Width           =   6495
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeCrawl"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   2
            Left            =   5040
            TabIndex        =   47
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CrawlNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   2
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1080
            Width           =   6225
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   11
            Top             =   360
            Width           =   1140
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "The first time the child sat up by itself"
         Height          =   5055
         Index           =   1
         Left            =   -74880
         TabIndex        =   5
         Top             =   840
         Width           =   6495
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeSat"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   1
            Left            =   5040
            TabIndex        =   46
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SatNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   960
            Width           =   6135
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "First time the child slept through the night"
         Height          =   5175
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   6375
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FirstTimeSleep"
            DataSource      =   "rsFirstTime"
            Height          =   315
            Index           =   0
            Left            =   5040
            TabIndex        =   45
            Top             =   360
            Width           =   1215
         End
         Begin VB.Data rsFirstTime 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Jørgen Programmer\MasterKid\Source\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   3000
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "FirstTime"
            Top             =   720
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SleepNote"
            DataSource      =   "rsFirstTime"
            Height          =   3975
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1080
            Width           =   6135
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Note:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmFirstTimes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Public Sub NewFirstTime()
    On Error Resume Next
    rsFirstTime.Recordset.Move 0
    rsFirstTime.Recordset.AddNew
    Date1(0).BackColor = &HC0E0FF
    Date1(0).SetFocus
    boolNewRecord = True
End Sub
Public Sub WriteFirstTimesWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        'sleep
        For i = 0 To 10
            .Selection.TypeText Text:=Frame1(i).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(i).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(CDate(Date1(i).Text), "dd.mm.yyyy")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label2(i).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Text1(i).Text
            .Selection.MoveRight Unit:=wdCell
        Next
    End With
    Set wdApp = Nothing
End Sub

Public Sub WriteFirstTimes()
    On Error Resume Next
    If Len(MDIMasterKid.cmbChildren.Text) = 0 Then Exit Sub
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    sHeader = rsLanguage.Fields("FormName")
    cPrint.pStartDoc
    Call PrintFront
    
    For i = 0 To 10
        If IsDate(Date1(i).Text) Then
            cPrint.FontBold = True
            cPrint.pPrint rsLanguage.Fields(i + 4), 1
            cPrint.FontBold = False
            cPrint.pPrint Label1(i).Caption & "  " & Format(CDate(Date1(i).Text), "dd.mm.yyyy"), 1
            cPrint.pPrint Label2(i).Caption, 1, True
            If Len(Text1(i).Text) <> 0 Then
                cPrint.pMultiline Text1(i).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
            Else
                cPrint.pPrint " ", 3.5
            End If
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                cPrint.pFooter
                cPrint.pNewPage
                Call PrintFront
            End If
        End If
    Next
    
    cPrint.pFooter
    Screen.MousePointer = vbDefault
    
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub

Public Function ShowFirstTime() As Boolean
Dim Sql As String
    On Error GoTo errShowFirstTime
    rsFirstTime.UpdateRecord
    Sql = "SELECT * FROM FirstTime WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsFirstTime.RecordSource = Sql
    rsFirstTime.Refresh
    rsFirstTime.Recordset.MoveFirst
    ShowFirstTime = True
    Label3.Caption = MDIMasterKid.cmbChildren.Text
    Exit Function
    
errShowFirstTime:
    ShowFirstTime = False
    Err.Clear
    Label3.Caption = " "
End Function


Private Sub ShowText()
Dim strHelp As String
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
                If IsNull(.Fields("Label1")) Then
                    .Fields("Label1") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("Label1")
                    Label1(1).Caption = .Fields("Label1")
                    Label1(2).Caption = .Fields("Label1")
                    Label1(3).Caption = .Fields("Label1")
                    Label1(4).Caption = .Fields("Label1")
                    Label1(5).Caption = .Fields("Label1")
                    Label1(6).Caption = .Fields("Label1")
                    Label1(7).Caption = .Fields("Label1")
                    Label1(8).Caption = .Fields("Label1")
                    Label1(9).Caption = .Fields("Label1")
                    Label1(10).Caption = .Fields("Label1")
                End If
                If IsNull(.Fields("Label2")) Then
                    .Fields("Label2") = Label2(0).Caption
                Else
                    Label2(0).Caption = .Fields("Label2")
                    Label2(1).Caption = .Fields("Label2")
                    Label2(2).Caption = .Fields("Label2")
                    Label2(3).Caption = .Fields("Label2")
                    Label2(4).Caption = .Fields("Label2")
                    Label2(5).Caption = .Fields("Label2")
                    Label2(6).Caption = .Fields("Label2")
                    Label2(7).Caption = .Fields("Label2")
                    Label2(8).Caption = .Fields("Label2")
                    Label2(9).Caption = .Fields("Label2")
                    Label2(10).Caption = .Fields("Label2")
                End If
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
                If IsNull(.Fields("Frame1(3)")) Then
                    .Fields("Frame1(3)") = Frame1(3).Caption
                Else
                    Frame1(3).Caption = .Fields("Frame1(3)")
                End If
                If IsNull(.Fields("Frame1(4)")) Then
                    .Fields("Frame1(4)") = Frame1(4).Caption
                Else
                    Frame1(4).Caption = .Fields("Frame1(4)")
                End If
                If IsNull(.Fields("Frame1(5)")) Then
                    .Fields("Frame1(5)") = Frame1(5).Caption
                Else
                    Frame1(5).Caption = .Fields("Frame1(5)")
                End If
                If IsNull(.Fields("Frame1(6)")) Then
                    .Fields("Frame1(6)") = Frame1(6).Caption
                Else
                    Frame1(6).Caption = .Fields("Frame1(6)")
                End If
                If IsNull(.Fields("Frame1(7)")) Then
                    .Fields("Frame1(7)") = Frame1(7).Caption
                Else
                    Frame1(7).Caption = .Fields("Frame1(7)")
                End If
                If IsNull(.Fields("Frame1(8)")) Then
                    .Fields("Frame1(8)") = Frame1(8).Caption
                Else
                    Frame1(8).Caption = .Fields("Frame1(8)")
                End If
                If IsNull(.Fields("Frame1(9)")) Then
                    .Fields("Frame1(9)") = Frame1(9).Caption
                Else
                    Frame1(9).Caption = .Fields("Frame1(9)")
                End If
                If IsNull(.Fields("Frame1(10)")) Then
                    .Fields("Frame1(10)") = Frame1(10).Caption
                Else
                    Frame1(10).Caption = .Fields("Frame1(10)")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("IndexTab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("IndexTab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("IndexTab12") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab12")
                End If
                Tab1.Tab = 3
                If IsNull(.Fields("Tab13")) Then
                    .Fields("IndexTab13") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab13")
                End If
                Tab1.Tab = 4
                If IsNull(.Fields("Tab14")) Then
                    .Fields("IndexTab14") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab14")
                End If
                Tab1.Tab = 5
                If IsNull(.Fields("Tab15")) Then
                    .Fields("IndexTab15") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab15")
                End If
                Tab1.Tab = 6
                If IsNull(.Fields("Tab16")) Then
                    .Fields("IndexTab16") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab16")
                End If
                Tab1.Tab = 7
                If IsNull(.Fields("Tab17")) Then
                    .Fields("IndexTab17") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab17")
                End If
                Tab1.Tab = 8
                If IsNull(.Fields("Tab18")) Then
                    .Fields("IndexTab18") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab18")
                End If
                Tab1.Tab = 9
                If IsNull(.Fields("Tab19")) Then
                    .Fields("IndexTab19") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab19")
                End If
                Tab1.Tab = 10
                If IsNull(.Fields("Tab110")) Then
                    .Fields("IndexTab110") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab110")
                End If
                Tab1.Tab = 0
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
        .Fields("Form") = Me.Caption
        .Fields("Label1(0)") = Label1(0).Caption
        .Fields("Label2(1)") = Label1(1).Caption
        .Fields("Frame1(0)") = Frame1(0).Caption
        .Fields("Frame1(1)") = Frame1(1).Caption
        .Fields("Frame1(2)") = Frame1(2).Caption
        .Fields("Frame1(3)") = Frame1(3).Caption
        .Fields("Frame1(4)") = Frame1(4).Caption
        .Fields("Frame1(5)") = Frame1(5).Caption
        .Fields("Frame1(6)") = Frame1(6).Caption
        .Fields("Frame1(7)") = Frame1(7).Caption
        .Fields("Frame1(8)") = Frame1(8).Caption
        .Fields("Frame1(9)") = Frame1(9).Caption
        .Fields("Frame1(10)") = Frame1(10).Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 4
        .Fields("Tab14") = Tab1.Caption
        Tab1.Tab = 5
        .Fields("Tab15") = Tab1.Caption
        Tab1.Tab = 6
        .Fields("Tab16") = Tab1.Caption
        Tab1.Tab = 7
        .Fields("Tab17") = Tab1.Caption
        Tab1.Tab = 8
        .Fields("Tab18") = Tab1.Caption
        Tab1.Tab = 9
        .Fields("Tab19") = Tab1.Caption
        Tab1.Tab = 10
        .Fields("Tab110") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("FormName") = Me.Caption
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
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
Private Sub Form_Activate()
    On Error Resume Next
    rsFirstTime.Refresh
    ShowText
    ShowAllButtons
    ShowKids
    ShowFirstTime
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsFirstTime.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmFirstTimes")
    iWhichForm = 19
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmFirstTime:  Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsFirstTime.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmFirstTimes = Nothing
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    Select Case Index
    Case 0
   On Error Resume Next
    Select Case Index
    Case 0
        If boolNewRecord Then
            If IsDate(Date1(0).Text) Then
                With rsFirstTime.Recordset
                    .Fields("ChildNo") = glChildNo
                    .Fields("FirstTimeSleep") = CDate(Format(Date1(0).Text, "dd.mm.yyyy"))
                    .Update
                    .Bookmark = .LastModified
                End With
            End If
            Date1(0).BackColor = &HFFFFC0
            boolNewRecord = False
        End If
    Case Else
    End Select
    Case Else
    End Select
End Sub
