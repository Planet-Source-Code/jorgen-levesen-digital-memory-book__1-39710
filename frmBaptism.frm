VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBaptism 
   BackColor       =   &H00400040&
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   9870
   Begin TabDlg.SSTab Tab1 
      Height          =   7095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   4194368
      TabCaption(0)   =   "Baptism"
      TabPicture(0)   =   "frmBaptism.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Goodfathers/Goodmothers"
      TabPicture(1)   =   "frmBaptism.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Label2(0)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Attendees"
      TabPicture(2)   =   "frmBaptism.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Gifts"
      TabPicture(3)   =   "frmBaptism.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Notes"
      TabPicture(4)   =   "frmBaptism.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label2(3)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "MyName"
      TabPicture(5)   =   "frmBaptism.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   6375
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   9015
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BaptismTime"
            DataSource      =   "rsBaptism"
            Height          =   285
            Index           =   5
            Left            =   3120
            TabIndex        =   36
            Top             =   840
            Width           =   795
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   255
            Left            =   3120
            TabIndex        =   35
            Top             =   840
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Date1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BaptismDate"
            DataSource      =   "rsBaptism"
            Height          =   315
            Left            =   3120
            TabIndex        =   0
            Top             =   360
            Width           =   1095
         End
         Begin VB.Data rsBaptism 
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
            RecordSource    =   "Baptism"
            Top             =   120
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.CommandButton btnScan 
            Height          =   495
            Index           =   0
            Left            =   4800
            Picture         =   "frmBaptism.frx":00A8
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Scan a picture"
            Top             =   4680
            Width           =   510
         End
         Begin VB.CommandButton btnCopyPic 
            Height          =   495
            Index           =   0
            Left            =   4800
            Picture         =   "frmBaptism.frx":01F2
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Copy picture to the Clipboard"
            Top             =   4200
            Width           =   510
         End
         Begin VB.CommandButton btnPastePicture 
            Height          =   495
            Index           =   0
            Left            =   4800
            Picture         =   "frmBaptism.frx":08B4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Paste picture"
            Top             =   3240
            Width           =   510
         End
         Begin VB.CommandButton btnReadFromFile 
            Height          =   495
            Index           =   0
            Left            =   4800
            Picture         =   "frmBaptism.frx":0F76
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Read picture from file"
            Top             =   3720
            Width           =   510
         End
         Begin VB.CommandButton btnDelete 
            Height          =   495
            Index           =   0
            Left            =   4800
            Picture         =   "frmBaptism.frx":1638
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Delete picture"
            Top             =   5160
            Width           =   510
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BaptismMinisterName"
            DataSource      =   "rsBaptism"
            Height          =   285
            Index           =   1
            Left            =   3165
            TabIndex        =   20
            Top             =   2520
            Width           =   5595
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BaptismWhere"
            DataSource      =   "rsBaptism"
            Height          =   1125
            Index           =   0
            Left            =   3165
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   1320
            Width           =   5595
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   390
            TabIndex        =   18
            Top             =   5640
            Visible         =   0   'False
            Width           =   420
         End
         Begin MSComDlg.CommonDialog Cmd1 
            Left            =   2400
            Top             =   5760
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            Height          =   2415
            Left            =   720
            Picture         =   "frmBaptism.frx":1782
            Stretch         =   -1  'True
            Top             =   3480
            Width           =   1920
         End
         Begin VB.Label Label1 
            Caption         =   "Church picture:"
            Height          =   255
            Index           =   4
            Left            =   5400
            TabIndex        =   30
            Top             =   2880
            Width           =   2085
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Minister Name:"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   29
            Top             =   2520
            Width           =   2085
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Where:"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   28
            Top             =   1320
            Width           =   2085
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Christening time:"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   27
            Top             =   840
            Width           =   2085
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Christening date:"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   26
            Top             =   360
            Width           =   2085
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "ChurchPicture"
            DataSource      =   "rsBaptism"
            Height          =   3015
            Left            =   5400
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   3315
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Names, Address and notes about Goodmothers / Goodfathers:"
         Height          =   6015
         Left            =   -74760
         TabIndex        =   15
         Top             =   840
         Width           =   8775
         Begin VB.Data rsGodmother 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Godmother"
            Top             =   6000
            Visible         =   0   'False
            Width           =   1140
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "GodFatherGodMother"
            DataSource      =   "rsGodmother"
            Height          =   5535
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   8340
            _ExtentX        =   14711
            _ExtentY        =   9763
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmBaptism.frx":18DD4
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Attendees, Names and family relation:"
         Height          =   6015
         Left            =   -74760
         TabIndex        =   13
         Top             =   840
         Width           =   8895
         Begin VB.Data rsBaptismAttendees 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "BaptismAttendees"
            Top             =   5640
            Visible         =   0   'False
            Width           =   1140
         End
         Begin RichTextLib.RichTextBox RichTextBox2 
            DataField       =   "BaptismAttendees"
            DataSource      =   "rsBaptismAttendees"
            Height          =   5535
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   9763
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmBaptism.frx":18EA9
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Baptism gifts and from whom:"
         Height          =   6015
         Left            =   -74760
         TabIndex        =   11
         Top             =   960
         Width           =   8895
         Begin VB.Data rsBaptismGifts 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "BaptismGifts"
            Top             =   5400
            Visible         =   0   'False
            Width           =   1140
         End
         Begin RichTextLib.RichTextBox RichTextBox3 
            DataField       =   "BaptismGifts"
            DataSource      =   "rsBaptismGifts"
            Height          =   5415
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   9551
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmBaptism.frx":18F7E
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Baptism notes:"
         Height          =   5895
         Left            =   -74760
         TabIndex        =   9
         Top             =   960
         Width           =   8895
         Begin VB.Data rsBaptismNote 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "BaptismNote"
            Top             =   5400
            Visible         =   0   'False
            Width           =   1140
         End
         Begin RichTextLib.RichTextBox RichTextBox4 
            DataField       =   "BaptismNotes"
            DataSource      =   "rsBaptismNote"
            Height          =   5415
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   9551
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmBaptism.frx":19053
         End
      End
      Begin VB.Frame Frame6 
         Height          =   6135
         Left            =   -74760
         TabIndex        =   2
         Top             =   720
         Width           =   8775
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "GotTheNameBecause"
            DataSource      =   "rsBaptism"
            Height          =   3285
            Index           =   4
            Left            =   2250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   2640
            Width           =   6285
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "NameChosenBy"
            DataSource      =   "rsBaptism"
            Height          =   1485
            Index           =   3
            Left            =   2280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   960
            Width           =   6285
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "IWasCalled"
            DataSource      =   "rsBaptism"
            Height          =   285
            Index           =   2
            Left            =   2280
            TabIndex        =   3
            Top             =   360
            Width           =   6165
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Got this name because:"
            Height          =   1095
            Index           =   7
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "This name was chosen by:"
            Height          =   1095
            Index           =   6
            Left            =   240
            TabIndex        =   7
            Top             =   960
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "I was named:"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1875
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "cc"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   34
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "cc"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   33
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "cc"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   32
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "cc"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   31
         Top             =   600
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmBaptism"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFirst  As Boolean
Dim rsLanguage As Recordset
Public Sub NewBaptism()
    rsBaptism.Recordset.AddNew
    Date1.SetFocus
    boolNewRecord = True
End Sub

Public Sub PrintBaptism()
Dim iX As Integer
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
    
    cPrint.pPrint Label1(0).Caption, 1, True   'baptism date
    If IsDate(Date1.Text) Then
        cPrint.pPrint Date1.Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(1).Caption, 1, True   'baptism time
    If IsDate(MaskEdBox1.Text) Then
        cPrint.pPrint Format(MaskEdBox1.Text, "hh:mm"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(2).Caption, 1, True   'where ?
    If Len(Text1(0).Text) <> 0 Then
        cPrint.pMultiline Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint Label1(3).Caption, 1, True   'minister name
    If Len(Text1(1).Text) <> 0 Then
        cPrint.pPrint Text1(1).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    cPrint.pPrint Label1(5).Caption, 1, True   'I was named ..
    If Len(Text1(2).Text) <> 0 Then
        cPrint.pPrint Text1(2).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    cPrint.pPrint Label1(6).Caption, 1, True   'my name was chosen by ..
    If Len(Text1(3).Text) <> 0 Then
        cPrint.pMultiline Text1(3).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    cPrint.pPrint Label1(7).Caption, 1, True   'got this name because...
    If Len(Text1(4).Text) <> 0 Then
        cPrint.pMultiline Text1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    
    'now print godfather / godmother
    sHeader = Frame2.Caption
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Frame2.Caption, 1    'goodmother / goodfather and addresses
    cPrint.FontBold = False
    cPrint.pPrint
    If Len(RichTextBox1.Text) <> 0 Then
        cPrint.pMultiline RichTextBox1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 1
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    
    'print attendees
    sHeader = Frame3.Caption
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Frame3.Caption, 1    'attendees and family relations
    cPrint.FontBold = False
    cPrint.pPrint
    If Len(RichTextBox2.Text) <> 0 Then
        cPrint.pMultiline RichTextBox2.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 1
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    
    'print gifts
    sHeader = Frame4.Caption   'gifts and from whom
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Frame4.Caption, 1
    cPrint.FontBold = False
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    If Len(RichTextBox3.Text) <> 0 Then
        cPrint.pMultiline RichTextBox3.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 1
    End If
    
    'print notes
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Frame5.Caption, 1  'baptism notes
    cPrint.FontBold = False
    cPrint.pPrint
    If Len(RichTextBox4.Text) Then
        cPrint.pMultiline RichTextBox4.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    
    'print the church picture
    cPrint.pPrint
    cPrint.pPrint Label1(4).Caption, 1   'church picture
    cPrint.pPrint
    If Not IsNull(rsBaptism.Recordset.Fields("ChurchPicture")) Then
        cPrint.pPrintPicture Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    
    'end print
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    ShowAllButtons
End Sub


Public Sub PrintBaptismWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.Tables(1).AllowAutoFit = False
        .Selection.TypeText Text:=Label1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(Date1.Text, "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(MaskEdBox1.Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        'I was named
        .Selection.TypeText Text:=Label1(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(3).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(4).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(4).Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData Image1.Picture, vbCFBitmap
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'godmothers /godfathers
        .Selection.TypeText Text:=Frame2.Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'attendees
        .Selection.TypeText Text:=Frame3.Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox2.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'gifts
        .Selection.TypeText Text:=Frame4.Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox3.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'notes
        .Selection.TypeText Text:=Frame5.Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox4.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
    End With
    Set wdApp = Nothing
End Sub

Public Sub SelectChild()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM Baptism WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBaptism.RecordSource = Sql
    rsBaptism.Refresh
    
    Sql = "SELECT * FROM Godmother WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsGodmother.RecordSource = Sql
    rsGodmother.Refresh
    
    Sql = "SELECT * FROM BaptismAttendees WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBaptismAttendees.RecordSource = Sql
    rsBaptismAttendees.Refresh
    
    Sql = "SELECT * FROM BaptismGifts WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBaptismGifts.RecordSource = Sql
    rsBaptismGifts.Refresh
    
    Sql = "SELECT * FROM BaptismNote WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBaptismNote.RecordSource = Sql
    rsBaptismNote.Refresh
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
                If IsNull(.Fields("Frame4")) Then
                    .Fields("Frame4") = Frame4.Caption
                Else
                    Frame4.Caption = .Fields("Frame4")
                End If
                If IsNull(.Fields("Frame5")) Then
                    .Fields("Frame5") = Frame5.Caption
                Else
                    Frame5.Caption = .Fields("Frame5")
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
                Tab1.Tab = 4
                If IsNull(.Fields("Tab14")) Then
                    .Fields("Tab14") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab14")
                End If
                Tab1.Tab = 5
                If IsNull(.Fields("Tab15")) Then
                    .Fields("Tab15") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab15")
                End If
                Tab1.Tab = 0
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
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete(0).ToolTipText
                Else
                    btnDelete(0).ToolTipText = .Fields("btnDelete")
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
        .Fields("Form") = Me.Caption
        For i = 0 To 7
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("Frame2") = Frame2.Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("Frame4") = Frame4.Caption
        .Fields("Frame5") = Frame5.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 4
        .Fields("Tab14") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("Tab15") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        .Fields("btnDelete") = btnDelete(0).ToolTipText
        .Fields("FormName") = "Baptism"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Public Sub btnCopyPic_Click(Index As Integer)
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

Private Sub Date1_LostFocus()
    On Error Resume Next
    If boolNewRecord Then
        With rsBaptism.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("BaptismDate") = CDate(Format(Date1.Text, "dd.mm.yyyy"))
            .Update
            .Bookmark = .LastModified
            
            rsGodmother.Recordset.AddNew
            rsGodmother.Recordset.Fields("ChildNo") = glChildNo
            rsGodmother.Recordset.Update
            
            rsBaptismAttendees.Recordset.AddNew
            rsBaptismAttendees.Recordset.Fields("ChildNo") = glChildNo
            rsBaptismAttendees.Recordset.Update
            
            rsBaptismGifts.Recordset.AddNew
            rsBaptismGifts.Recordset.Fields("ChildNo") = glChildNo
            rsBaptismGifts.Recordset.Update
            
            rsBaptismNote.Recordset.AddNew
            rsBaptismNote.Recordset.Fields("ChildNo") = glChildNo
            rsBaptismNote.Recordset.Update
            boolNewRecord = False
        End With
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If bFirst = False Then Exit Sub
    rsBaptism.Refresh
    rsGodmother.Refresh
    rsBaptismAttendees.Refresh
    rsBaptismGifts.Refresh
    rsBaptismNote.Refresh
    ShowKids
    SelectChild
    Label2(0).Caption = gsChildName
    Label2(1).Caption = gsChildName
    Label2(2).Caption = gsChildName
    Label2(3).Caption = gsChildName
    ShowText
    ShowAllButtons
    Me.WindowState = vbMaximized
    bFirst = False
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsBaptism.DatabaseName = dbKidsTxt
    rsGodmother.DatabaseName = dbKidsTxt
    rsBaptismAttendees.DatabaseName = dbKidsTxt
    rsBaptismGifts.DatabaseName = dbKidsTxt
    rsBaptismNote.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmBaptism")
    iWhichForm = 20
    bFirst = True
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmBaptism: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsBaptism.Recordset.Close
    rsGodmother.Recordset.Close
    rsBaptismAttendees.Recordset.Close
    rsBaptismGifts.Recordset.Close
    rsBaptismNote.Recordset.Close
    rsLanguage.Close
    iWhichForm = 0
    HideAllButtons
    HideKids
    Set frmBaptism = Nothing
End Sub

Private Sub MaskEdBox1_LostFocus()
    Text1(5).Text = Format(MaskEdBox1.Text, "hh:mm")
    MaskEdBox1.Visible = False
    Text1(5).Visible = True
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    On Error Resume Next
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox1_SelChange()
    Call RichTextSelChange(frmBaptism.RichTextBox1)
End Sub

Private Sub RichTextBox2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    On Error Resume Next
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox2.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox2_SelChange()
    Call RichTextSelChange(frmBaptism.RichTextBox2)
End Sub

Private Sub RichTextBox3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    On Error Resume Next
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox3.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox3_SelChange()
    Call RichTextSelChange(frmBaptism.RichTextBox3)
End Sub

Private Sub RichTextBox4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    On Error Resume Next
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox4.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox4_SelChange()
    Call RichTextSelChange(frmBaptism.RichTextBox4)
End Sub

Private Sub Text1_Click(Index As Integer)
    Select Case Index
    Case 5
        Text1(5).Visible = False
        MaskEdBox1.Visible = True
        MaskEdBox1.SetFocus
    Case Else
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    onGotFocus
End Sub
