VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBirth 
   BackColor       =   &H00008000&
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   7785
   Begin TabDlg.SSTab Tab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   32768
      TabCaption(0)   =   "The Birth"
      TabPicture(0)   =   "frmBirth.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Birth Notes"
      TabPicture(1)   =   "frmBirth.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "RichTextBox1"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   6855
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   5055
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BirthTime"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   10
            Left            =   1800
            TabIndex        =   40
            Top             =   6120
            Width           =   870
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WentToHospitalTime"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   9
            Left            =   1800
            TabIndex        =   39
            Top             =   3720
            Width           =   870
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "LabourStartAt"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   8
            Left            =   1800
            TabIndex        =   38
            Top             =   720
            Width           =   870
         End
         Begin MSMask.MaskEdBox Time1 
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   1
            Top             =   720
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BirthDate"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   9
            Top             =   5760
            Width           =   1215
         End
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WentToHospitalDate"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   4
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "LabourDate"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.Data rsBirth 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Birth"
            Top             =   6480
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Where"
            DataSource      =   "rsBirth"
            Height          =   975
            Index           =   0
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Symptoms"
            DataSource      =   "rsBirth"
            Height          =   975
            Index           =   1
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WasMetBy"
            DataSource      =   "rsBirth"
            Height          =   975
            Index           =   2
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   4200
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "LaburDuration"
            DataSource      =   "rsBirth"
            Height          =   375
            Index           =   3
            Left            =   1800
            TabIndex        =   7
            Top             =   5280
            Width           =   975
         End
         Begin VB.ComboBox cmbDimension 
            BackColor       =   &H00FFFFC0&
            DataField       =   "LaburDurationDim"
            DataSource      =   "rsBirth"
            Height          =   315
            Index           =   0
            Left            =   3000
            TabIndex        =   8
            Top             =   5280
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Boy"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1845
            TabIndex        =   11
            Top             =   6510
            Width           =   975
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Girl"
            DataField       =   "Girl"
            DataSource      =   "rsBirth"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   2910
            TabIndex        =   12
            Top             =   6510
            Width           =   735
         End
         Begin VB.CommandButton btnHoroscope 
            Caption         =   "&Horoscope"
            Height          =   615
            Left            =   3720
            TabIndex        =   26
            Top             =   6120
            Width           =   1215
         End
         Begin MSMask.MaskEdBox Time1 
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   5
            Top             =   3720
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Time1 
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   10
            Top             =   6120
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Labour:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Time Of Labour:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   35
            Top             =   720
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Where did labour start:"
            ForeColor       =   &H00000000&
            Height          =   705
            Index           =   2
            Left            =   0
            TabIndex        =   34
            Top             =   1185
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Symptons:"
            ForeColor       =   &H00000000&
            Height          =   705
            Index           =   3
            Left            =   0
            TabIndex        =   33
            Top             =   2340
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Went to Hospital date:"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   0
            TabIndex        =   32
            Top             =   3390
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Went to Hospital time:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   31
            Top             =   3720
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Was met by:"
            ForeColor       =   &H00000000&
            Height          =   585
            Index           =   6
            Left            =   0
            TabIndex        =   30
            Top             =   4200
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Labour Duration:"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   7
            Left            =   0
            TabIndex        =   29
            Top             =   5355
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Of Birth:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   28
            Top             =   5760
            Width           =   1755
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Time Of Birth:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   27
            Top             =   6165
            Width           =   1755
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1815
         Left            =   5400
         TabIndex        =   22
         Top             =   1560
         Width           =   1935
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "HairColor"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   1440
            Width           =   1650
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "EyeColor"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   1650
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Hair Color:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   24
            Top             =   1200
            Width           =   1650
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Eye Color:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1650
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Baby Weight"
         Height          =   1215
         Left            =   5400
         TabIndex        =   21
         Top             =   3480
         Width           =   1935
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BabyWeight"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   870
         End
         Begin VB.ComboBox cmbDimension 
            BackColor       =   &H00FFFFC0&
            DataField       =   "BabyWeightDim"
            DataSource      =   "rsBirth"
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   16
            Top             =   600
            Width           =   690
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Baby Height"
         Height          =   1215
         Left            =   5400
         TabIndex        =   20
         Top             =   4800
         Width           =   1935
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BabyLength"
            DataSource      =   "rsBirth"
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   870
         End
         Begin VB.ComboBox cmbDimension 
            BackColor       =   &H00FFFFC0&
            DataField       =   "BabyLengthDim"
            DataSource      =   "rsBirth"
            Height          =   315
            Index           =   2
            Left            =   1080
            TabIndex        =   18
            Top             =   600
            Width           =   690
         End
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         DataField       =   "BirthNotes"
         DataSource      =   "rsBirth"
         Height          =   6735
         Left            =   -74760
         TabIndex        =   37
         Top             =   600
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   11880
         _Version        =   393217
         BackColor       =   16777152
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmBirth.frx":0038
      End
      Begin VB.Image Image1 
         Height          =   1095
         Index           =   0
         Left            =   5280
         Picture         =   "frmBirth.frx":010D
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1995
      End
      Begin VB.Image Image1 
         Height          =   1095
         Index           =   1
         Left            =   5400
         Picture         =   "frmBirth.frx":10ED
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmBirth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTime As Recordset
Dim rsLength As Recordset
Dim rsWeight As Recordset
Dim rsWeightLength As Recordset
Dim rsChildren As Recordset
Dim rsLanguage As Recordset
Public Sub NewBirth()
    rsBirth.Recordset.AddNew
    Date1(0).SetFocus
    boolNewRecord = True
End Sub

Public Sub PrintBirth()
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
    
    cPrint.pPrint Label1(0).Caption, 1, True
    If IsDate(Date1(0).Text) Then
        cPrint.pPrint Format(CDate(Date1(0).Text), "dd.mm.yyyy"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(1).Caption, 1, True
    If IsDate(Text1(8).Text) Then
        cPrint.pPrint Format(Text1(8).Text, "hh:mm"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(2).Caption, 1, True
    If Len(Text1(0).Text) <> 0 Then
        cPrint.pMultiline Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint Label1(3).Caption, 1, True
    If Len(Text1(1).Text) <> 0 Then
        cPrint.pMultiline Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    
    cPrint.pPrint
    cPrint.pPrint Label1(4).Caption, 1, True
    If IsDate(Date1(1).Text) Then
        cPrint.pPrint Format(CDate(Date1(1).Text), "dd.mm.yyyy"), 3.5  'at hospital date
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(5).Caption, 1, True
    If IsDate(Text1(9).Text) Then
        cPrint.pPrint Format(Text1(9).Text, "hh:mm"), 3.5  'at hospital time
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(6).Caption, 1, True
    If Len(Text1(2).Text) <> 0 Then
        cPrint.pMultiline Text1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    
    cPrint.pPrint
    cPrint.pPrint Label1(7).Caption, 1, True
    If Len(Text1(3).Text) <> 0 Then
        cPrint.pPrint Text1(3).Text & "  " & cmbDimension(0).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(8).Caption, 1, True
    If IsDate(Date1(2).Text) Then
        cPrint.pPrint Format(Date1(2).Text, "dd.mm.yyyy"), 3.5
    Else
        cPrint.pPrint "", 3.5
    End If
    cPrint.pPrint Label1(9).Caption, 1, True
    If IsDate(Text1(10).Text) Then
        cPrint.pPrint Format(CDate(Text1(10).Text), "hh:mm"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(10).Caption, 1, True
    If Len(Text1(6).Text) <> 0 Then
        cPrint.pPrint Text1(6).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint Label1(11).Caption, 1, True
    If Len(Text1(7).Text) <> 0 Then
        cPrint.pPrint Text1(7).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Frame3.Caption, 1, True
    If Len(Text1(4).Text) <> 0 Then
        cPrint.pPrint Format(Text1(4).Text, "0.000") & "" & "  " & cmbDimension(1).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Frame4.Caption, 1, True
    If Len(Text1(5).Text) <> 0 Then
        cPrint.pPrint Format(Text1(5).Text, "0.00") & "" & "  " & cmbDimension(2).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint "Note:", 1, True
    If Len(RichTextBox1.Text) <> 0 Then
        cPrint.pMultiline RichTextBox1.Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    MeActivate
End Sub
Public Sub PrintBirthWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.TypeText Text:=Label1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(Date1(0).Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(Text1(8).Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(Date1(1).Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(Text1(9).Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(3).Text & "  " & cmbDimension(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(8).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(Date1(2).Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(9).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(Text1(10).Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(10).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(6).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(11).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(7).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame3.Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(Text1(4).Text, "0.000") & "" & "  " & cmbDimension(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame4.Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(Text1(5).Text, "0.000") & "" & "  " & cmbDimension(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:="Note"
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
    End With
    Set wdApp = Nothing
End Sub


Public Function ShowChild() As Boolean
Dim Sql As String
    On Error GoTo errShowChild
    Sql = "SELECT * FROM Birth WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBirth.RecordSource = Sql
    rsBirth.Refresh
    rsBirth.Recordset.MoveFirst
    ShowChild = True
    Exit Function
    
errShowChild:
    ShowChild = True
    Err.Clear
End Function

Private Sub ShowText()
Dim strHelp As String
    On Error Resume Next
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                For i = 0 To 11
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
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
                If IsNull(.Fields("Check1(0)")) Then
                    .Fields("Check1(0)") = Check1(0).Caption
                Else
                    Check1(0).Caption = .Fields("Check1(0)")
                End If
                If IsNull(.Fields("Check1(1)")) Then
                    .Fields("Check1(1)") = Check1(1).Caption
                Else
                    Check1(1).Caption = .Fields("Check1(1)")
                End If
                If IsNull(.Fields("btnHoroscope")) Then
                    .Fields("btnHoroscope") = btnHoroscope.Caption
                Else
                    btnHoroscope.Caption = .Fields("btnHoroscope")
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
        For i = 0 To 11
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("Frame3") = Frame3.Caption
        .Fields("Frame4") = Frame4.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("Check1(0)") = Check1(0).Caption
        .Fields("Check1(1)") = Check1(1).Caption
        .Fields("btnHoroscope") = btnHoroscope.Caption
        .Fields("FormName") = "The Birth"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Private Sub LoadDimensions()
    With rsTime
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbDimension(0).AddItem .Fields("TimeDim")
            End If
        .MoveNext
        Loop
    End With
   
    With rsLength
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbDimension(2).AddItem .Fields("LengthDim")
            End If
        .MoveNext
        Loop
    End With
    
    With rsWeight
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbDimension(1).AddItem .Fields("WeightDim")
            End If
        .MoveNext
        Loop
    End With
End Sub

Private Sub btnHoroscope_Click()
Dim sString, sYear As String, Date11 As Date, Date21 As Date
    sYear = Right(Date1(2).Text, 4)
    Date11 = CDate(Date1(2).Text)
    Load frmZodiac
    DoEvents
    With frmZodiac.rsHoroscope.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                sString = .Fields("FromDate")
                sString = sString & "." & sYear
                Date21 = CDate(sString)
                If compareDates(Date11, Date21) = 1 Then
                    sString = .Fields("ToDate")
                    sString = sString & "." & sYear
                    Date21 = CDate(sString)
                    If compareDates(Date11, Date21) = 2 Then
                        frmZodiac.Show 1
                        Exit Sub
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    Unload frmZodiac
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
    Select Case Index
    Case 0
        If boolNewRecord Then
            With rsBirth.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("LabourDate") = CDate(Format(Date1(0).Text, "dd.mm.yyyy"))
                .Update
                .Bookmark = .LastModified
            End With
            boolNewRecord = False
        End If
    Case Else
    End Select
End Sub

Private Sub MeActivate()
    On Error Resume Next
    rsBirth.Refresh
    ShowChild
    Frame2.Caption = gsChildName
    LoadDimensions
    ShowText
    ShowAllButtons
    ShowKids
End Sub

Private Sub Form_Activate()
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsBirth.DatabaseName = dbKidsTxt
    Set rsWeightLength = dbKids.OpenRecordset("WeightLength")
    Set rsTime = dbKids.OpenRecordset("DimTime")
    Set rsLength = dbKids.OpenRecordset("DimLength")
    Set rsWeight = dbKids.OpenRecordset("DimWeight")
    Set rsChildren = dbKids.OpenRecordset("Children")
    Set rsLanguage = dbKidLang.OpenRecordset("frmBirth")
    MeActivate
    iWhichForm = 15
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmBirth:  Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim bFound As Boolean
    'update the weight/length records
    On Error Resume Next
    bFound = False
    With rsWeightLength
        .MoveFirst
        Do While Not .EOF
        If CLng(.Fields("ChildNo")) = glChildNo Then
            .Edit
            If IsDate(Date1(2).Text) <> 0 Then
                .Fields("AtBirthDate") = CDate(Format(Date1(2).Text, "dd.mm.yyyy"))
            End If
            If Len(Text1(5).Text) <> 0 Then
                .Fields("AtBirthLength") = Text1(5).Text
            End If
            If Len(Text1(4).Text) <> 0 Then
                .Fields("AtBirthWeight") = Text1(4).Text
            End If
            .Update
            bFound = True
            Exit Do
        End If
        .MoveNext
        Loop
        If bFound Then
            .AddNew
            .Fields("ChildNo") = glChildNo
            If IsDate(Date1(2).Text) <> 0 Then
                .Fields("AtBirthDate") = CDate(Format(Date1(2).Text, "dd.mm.yyyy"))
            End If
            If Len(Text1(5).Text) <> 0 Then
                .Fields("AtBirthLength") = Text1(5).Text
            End If
            If Len(Text1(4).Text) <> 0 Then
                .Fields("AtBirthWeight") = Text1(4).Text
            End If
            .Update
        End If
    End With
    
    'the childrens records
    With rsChildren
        .MoveFirst
        Do While Not .EOF
        If CLng(.Fields("ChildNo")) = glChildNo Then
            .Edit
            If Check1(1).Value = 1 Then
                .Fields("ChildFemale") = True
            End If
            If IsDate(Date1(2).Text) Then
                .Fields("BirthDate") = CDate(Date1(2).Text)
            End If
            If IsDate(Time1(2).Text) Then
                .Fields("BirthTime") = Format(Time1(2).Text, "hh:mm")
            End If
            If IsNumeric(Text1(3).Text) Then
                .Fields("LaburDuration") = Text1(3).Text
            End If
            .Fields("LaburDurationDim") = CStr(cmbDimension(0).Text)
            If IsNumeric(Text1(4).Text) Then
                .Fields("BabyWeight") = Text1(4).Text
            End If
            .Fields("BabyWeightDim") = CStr(cmbDimension(1).Text)
            If IsNumeric(Text1(5).Text) Then
                .Fields("BabyLength") = Text1(5).Text
            End If
            .Fields("BabyLengthDim") = CStr(cmbDimension(2).Text)
            .Update
            Exit Sub
        End If
        .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsBirth.Recordset.Close
    rsWeightLength.Close
    rsLanguage.Close
    rsTime.Close
    rsChildren.Close
    rsLength.Close
    rsWeight.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmBirth = Nothing
End Sub
Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox1_SelChange()
    Call RichTextSelChange(frmBirth.RichTextBox1)
End Sub
Private Sub Text1_Click(Index As Integer)
    Select Case Index
    Case 8
        If Len(Text1(8).Text) <> 0 Then
            Time1(0).Text = Format(Text1(8).Text, "hh:mm")
        End If
        Text1(8).Visible = False
        Time1(0).Visible = True
    Case 9
        If Len(Text1(9).Text) <> 0 Then
            Time1(1).Text = Format(Text1(9).Text, "hh:mm")
        End If
        Text1(9).Visible = False
        Time1(1).Visible = True
    Case 10
        If Len(Text1(10).Text) <> 0 Then
            Time1(2).Text = Format(Text1(10).Text, "hh:mm")
        End If
        Text1(10).Visible = False
        Time1(2).Visible = True
    Case Else
    End Select
End Sub
Private Sub Time1_LostFocus(Index As Integer)
    Select Case Index
    Case 0
        Text1(8).Text = Format(Time1(0).Text, "hh:mm")
        Time1(0).Visible = False
        Text1(8).Visible = True
    Case 1
        Text1(9).Text = Format(Time1(1).Text, "hh:mm")
        Time1(1).Visible = False
        Text1(9).Visible = True
    Case 2
        Text1(10).Text = Format(Time1(2).Text, "hh:mm")
        Time1(2).Visible = False
        Text1(10).Visible = True
    Case Else
    End Select
End Sub
