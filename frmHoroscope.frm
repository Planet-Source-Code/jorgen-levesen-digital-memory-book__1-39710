VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHoroscope 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horoscope"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8388608
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmHoroscope.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Long Description"
      TabPicture(1)   =   "frmHoroscope.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(4)"
      Tab(1).ControlCount=   1
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "HoroscopeDescription"
         DataSource      =   "rsHoroscope"
         Height          =   5685
         Index           =   4
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   480
         Width           =   5295
      End
      Begin VB.Frame Frame1 
         Height          =   5775
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   5295
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "HoroscopeShort"
            DataSource      =   "rsHoroscope"
            Height          =   1245
            Index           =   3
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   1800
            Width           =   4815
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ToDate"
            DataSource      =   "rsHoroscope"
            Height          =   285
            Index           =   2
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   11
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FromDate"
            DataSource      =   "rsHoroscope"
            Height          =   285
            Index           =   1
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   10
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "HoroscopeText"
            DataSource      =   "rsHoroscope"
            Height          =   285
            Index           =   0
            Left            =   2760
            MaxLength       =   25
            TabIndex        =   9
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnDelete 
            Height          =   495
            Index           =   0
            Left            =   4440
            Picture         =   "frmHoroscope.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Delete picture"
            Top             =   4920
            Width           =   495
         End
         Begin VB.CommandButton btnReadFromFile 
            Height          =   495
            Index           =   0
            Left            =   3000
            Picture         =   "frmHoroscope.frx":0182
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Read picture from file"
            Top             =   4920
            Width           =   495
         End
         Begin VB.CommandButton btnPastePicture 
            Height          =   495
            Index           =   0
            Left            =   2520
            Picture         =   "frmHoroscope.frx":0844
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Paste picture"
            Top             =   4920
            Width           =   495
         End
         Begin VB.CommandButton btnCopyPic 
            Height          =   495
            Index           =   0
            Left            =   3480
            Picture         =   "frmHoroscope.frx":0F06
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Copy picture to the Clipboard"
            Top             =   4920
            Width           =   495
         End
         Begin VB.CommandButton btnScan 
            Height          =   495
            Index           =   0
            Left            =   3960
            Picture         =   "frmHoroscope.frx":15C8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Scan a picture"
            Top             =   4920
            Width           =   495
         End
         Begin MSComDlg.CommonDialog Cmd1 
            Left            =   240
            Top             =   3240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sign Picture:"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   19
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Short Description:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "(dd.mm)"
            Height          =   255
            Index           =   5
            Left            =   3720
            TabIndex        =   17
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "(dd.mm)"
            Height          =   255
            Index           =   4
            Left            =   3720
            TabIndex        =   16
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "To Date:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "From Date:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Sign Text:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2535
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "ZodiacSignPicture"
            DataSource      =   "rsHoroscope"
            Height          =   1575
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   2415
         End
      End
   End
   Begin VB.Data rsHoroscope 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Horoscope"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5880
      Left            =   5760
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Signs"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmHoroscope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v1RecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewHoroscope()
    boolNewRecord = True
    rsHoroscope.Recordset.AddNew
    Text1(0).SetFocus
End Sub

Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                For i = 0 To 7
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                If IsNull(.Fields("btnPastePicture")) Then
                    .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
                Else
                    btnPastePicture(0).ToolTipText = .Fields("btnPastePicture")
                End If
                If IsNull(.Fields("btnReadFromFile")) Then
                    .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
                Else
                    btnReadFromFile(0).ToolTipText = .Fields("btnReadFromFile")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete(0).ToolTipText
                Else
                    btnDelete(0).ToolTipText = .Fields("btnDelete")
                End If
                If IsNull(.Fields("btnCopyPic")) Then
                    .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
                Else
                    btnCopyPic(0).ToolTipText = .Fields("btnCopyPic")
                End If
                If IsNull(.Fields("btnScan")) Then
                    .Fields("btnScan") = btnScan(0).ToolTipText
                Else
                    btnScan(0).ToolTipText = .Fields("btnScan")
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
            
        .MoveFirst
        .AddNew
        .Fields("Language") = FileExt
        For i = 0 To 7
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("Form") = Me.Caption
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
        .Fields("btnDelete") = btnDelete(0).ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Public Sub FillList1()
    On Error Resume Next
    List1.Clear
    With rsHoroscope.Recordset
        .MoveLast
        .MoveFirst
        ReDim v1RecordBookmark(.RecordCount)
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                List1.AddItem .Fields("HoroscopeText")
                List1.ItemData(List1.NewIndex) = List1.ListCount - 1
                v1RecordBookmark(List1.ListCount - 1) = .Bookmark
            End If
        .MoveNext
        Loop
    End With
End Sub

Private Sub btnCopyPic_Click(Index As Integer)
    On Error Resume Next
    Clipboard.SetData Image1.Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click(Index As Integer)
    On Error Resume Next
    Set Image1.Picture = LoadPicture()
End Sub

Private Sub btnPastePicture_Click(Index As Integer)
        On Error Resume Next
        Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub btnReadFromFile_Click(Index As Integer)
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

Private Sub btnScan_Click(Index As Integer)
    Dim ret As Long, t As Single
    On Error Resume Next
    ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsHoroscope.Refresh
    ShowText
    FillList1
    List1.ListIndex = 0
    ShowAllButtons
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsHoroscope.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmHoroscope")
    iWhichForm = 38
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmHoroscope: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsHoroscope.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    iWhichForm = 0
    Set frmHoroscope = Nothing
End Sub

Private Sub List1_Click()
    rsHoroscope.Recordset.Bookmark = v1RecordBookmark(List1.ItemData(List1.ListIndex))
End Sub



Private Sub Text1_LostFocus(Index As Integer)
    If boolNewRecord Then
        Select Case Index
        Case 2
        With rsHoroscope.Recordset
            .Fields("Language") = FileExt
            If Len(Text1(1).Text) <> 0 Then
                .Fields("FromDate") = Text1(1).Text
            Else
                .Fields("FromDate") = "01.01"   'can not have null value
            End If
            If Len(Text1(2).Text) <> 0 Then
                .Fields("ToDate") = Text1(2).Text
            Else
                .Fields("ToDate") = "01.01" 'can not have null value
            End If
            .Fields("HoroscopeText") = Trim(Text1(0).Text)
            .Update
            boolNewRecord = False
            .Bookmark = .LastModified
        End With
        Case Else
        End Select
    End If
End Sub


