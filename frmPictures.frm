VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPictures 
   BackColor       =   &H0000C0C0&
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   7890
   Begin TabDlg.SSTab Tab1 
      Height          =   6975
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   49344
      TabCaption(0)   =   "Birth"
      TabPicture(0)   =   "frmPictures.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Cmd1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Picture2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "List2(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnPastePicture(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "btnReadFromFile(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text2(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "btnCopyPic(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "btnDelete(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "btnScan(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "rsBirthPictures"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Infant (0-1 year)"
      TabPicture(1)   =   "frmPictures.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List2(1)"
      Tab(1).Control(1)=   "btnPastePicture(1)"
      Tab(1).Control(2)=   "btnReadFromFile(1)"
      Tab(1).Control(3)=   "Text2(1)"
      Tab(1).Control(4)=   "btnCopyPic(1)"
      Tab(1).Control(5)=   "btnDelete(1)"
      Tab(1).Control(6)=   "btnScan(1)"
      Tab(1).Control(7)=   "rsInfancyPictures"
      Tab(1).Control(8)=   "Label1(1)"
      Tab(1).Control(9)=   "Label2(1)"
      Tab(1).Control(10)=   "Picture1(1)"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Childhood"
      TabPicture(2)   =   "frmPictures.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btnScan(2)"
      Tab(2).Control(1)=   "btnDelete(2)"
      Tab(2).Control(2)=   "btnCopyPic(2)"
      Tab(2).Control(3)=   "Text2(2)"
      Tab(2).Control(4)=   "btnReadFromFile(2)"
      Tab(2).Control(5)=   "btnPastePicture(2)"
      Tab(2).Control(6)=   "List2(2)"
      Tab(2).Control(7)=   "rsChildPictures"
      Tab(2).Control(8)=   "Label2(2)"
      Tab(2).Control(9)=   "Label1(2)"
      Tab(2).Control(10)=   "Picture1(2)"
      Tab(2).ControlCount=   11
      Begin VB.CommandButton btnScan 
         Height          =   450
         Index           =   2
         Left            =   -73245
         Picture         =   "frmPictures.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Scan a picture from Scanner"
         Top             =   2610
         Width           =   375
      End
      Begin VB.CommandButton btnDelete 
         Height          =   450
         Index           =   2
         Left            =   -73245
         Picture         =   "frmPictures.frx":019E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Delete shown picture"
         Top             =   3150
         Width           =   375
      End
      Begin VB.CommandButton btnCopyPic 
         Height          =   450
         Index           =   2
         Left            =   -73245
         Picture         =   "frmPictures.frx":02E8
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Copy picture to the Clipboard"
         Top             =   2070
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PictureCaption"
         DataSource      =   "rsChildPictures"
         Height          =   1410
         Index           =   2
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   5340
         Width           =   5730
      End
      Begin VB.CommandButton btnReadFromFile 
         Height          =   450
         Index           =   2
         Left            =   -73245
         Picture         =   "frmPictures.frx":09AA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Read Picture from a disk file"
         Top             =   1530
         Width           =   375
      End
      Begin VB.CommandButton btnPastePicture 
         Height          =   450
         Index           =   2
         Left            =   -73245
         Picture         =   "frmPictures.frx":106C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Paste Picture from the Clipboard"
         Top             =   990
         Width           =   375
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5880
         Index           =   2
         Left            =   -68835
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   795
         Width           =   810
      End
      Begin VB.Data rsChildPictures 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidPic.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -69825
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ChildPictures"
         Top             =   1230
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5880
         Index           =   1
         Left            =   -68805
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   795
         Width           =   810
      End
      Begin VB.CommandButton btnPastePicture 
         Height          =   450
         Index           =   1
         Left            =   -73125
         Picture         =   "frmPictures.frx":172E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Paste Picture from the Clipboard"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton btnReadFromFile 
         Height          =   450
         Index           =   1
         Left            =   -73125
         Picture         =   "frmPictures.frx":1DF0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Read Picture from a disk file"
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PictureCaption"
         DataSource      =   "rsInfancyPictures"
         Height          =   1410
         Index           =   1
         Left            =   -74505
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   5340
         Width           =   5550
      End
      Begin VB.CommandButton btnCopyPic 
         Height          =   450
         Index           =   1
         Left            =   -73125
         Picture         =   "frmPictures.frx":24B2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Copy picture to the Clipboard"
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton btnDelete 
         Height          =   450
         Index           =   1
         Left            =   -73125
         Picture         =   "frmPictures.frx":2B74
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Delete shown picture"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton btnScan 
         Height          =   450
         Index           =   1
         Left            =   -73125
         Picture         =   "frmPictures.frx":2CBE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Scan a picture from Scanner"
         Top             =   2340
         Width           =   375
      End
      Begin VB.Data rsInfancyPictures 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidPic.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "InfancyPictures"
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data rsBirthPictures 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidPic.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "BirthPictures"
         Top             =   480
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton btnScan 
         Height          =   450
         Index           =   0
         Left            =   1305
         Picture         =   "frmPictures.frx":2E08
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Scan a picture from Scanner"
         Top             =   2415
         Width           =   375
      End
      Begin VB.CommandButton btnDelete 
         Height          =   450
         Index           =   0
         Left            =   1305
         Picture         =   "frmPictures.frx":2F52
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete shown picture"
         Top             =   2955
         Width           =   375
      End
      Begin VB.CommandButton btnCopyPic 
         Height          =   450
         Index           =   0
         Left            =   1305
         Picture         =   "frmPictures.frx":309C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Copy picture to the Clipboard"
         Top             =   1875
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PictureCaption"
         DataSource      =   "rsBirthPictures"
         Height          =   1410
         Index           =   0
         Left            =   555
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   5340
         Width           =   5340
      End
      Begin VB.CommandButton btnReadFromFile 
         Height          =   450
         Index           =   0
         Left            =   1305
         Picture         =   "frmPictures.frx":375E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Read Picture from a disk file"
         Top             =   1335
         Width           =   375
      End
      Begin VB.CommandButton btnPastePicture 
         Height          =   450
         Index           =   0
         Left            =   1305
         Picture         =   "frmPictures.frx":3E20
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Paste Picture from the Clipboard"
         Top             =   795
         Width           =   375
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5880
         Index           =   0
         Left            =   6090
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   795
         Width           =   840
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   660
         ScaleHeight     =   465
         ScaleWidth      =   390
         TabIndex        =   2
         Top             =   4200
         Visible         =   0   'False
         Width           =   420
      End
      Begin MSComDlg.CommonDialog Cmd1 
         Left            =   480
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Caption         =   "Picture Note:"
         Height          =   240
         Index           =   2
         Left            =   -74640
         TabIndex        =   0
         Top             =   5115
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Picture No."
         Height          =   255
         Index           =   2
         Left            =   -68820
         TabIndex        =   28
         Top             =   600
         Width           =   825
      End
      Begin VB.Image Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Picture"
         DataSource      =   "rsChildPictures"
         Height          =   4335
         Index           =   2
         Left            =   -72825
         Stretch         =   -1  'True
         Top             =   960
         Width           =   3720
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Picture No."
         Height          =   255
         Index           =   1
         Left            =   -68805
         TabIndex        =   27
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Picture Note:"
         Height          =   240
         Index           =   1
         Left            =   -74520
         TabIndex        =   26
         Top             =   5115
         Width           =   795
      End
      Begin VB.Image Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Picture"
         DataSource      =   "rsInfancyPictures"
         Height          =   4335
         Index           =   1
         Left            =   -72705
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3720
      End
      Begin VB.Label Label2 
         Caption         =   "Picture Note:"
         Height          =   240
         Index           =   0
         Left            =   555
         TabIndex        =   25
         Top             =   5115
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Picture No."
         Height          =   255
         Index           =   0
         Left            =   6090
         TabIndex        =   24
         Top             =   600
         Width           =   870
      End
      Begin VB.Image Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Picture"
         DataSource      =   "rsBirthPictures"
         Height          =   4335
         Index           =   0
         Left            =   1950
         Stretch         =   -1  'True
         Top             =   840
         Width           =   3945
      End
   End
End
Attribute VB_Name = "frmPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String, bFirstWrite As Boolean
Dim BirthRecordBookmark() As Variant
Dim InfantRecordBookmark() As Variant
Dim ChildRecordBookmark() As Variant
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
            .Tab1.BackColor = &H4040&
        Case Else
        End Select
    End With
End Sub

Public Sub DeletePictures()
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        rsBirthPictures.Recordset.Delete
        If SelectPicBirth Then
            FillList20
            List2(0).ListIndex = 0
        Else
            List2(0).Clear
        End If
    Case 1
        rsInfancyPictures.Recordset.Delete
        If SelectPicInfant Then
            FillList21
            List2(1).ListIndex = 0
        Else
            List2(1).Clear
        End If
    Case 2
        rsChildPictures.Recordset.Delete
        If SelectPicChild Then
            FillList22
            List2(2).ListIndex = 0
        Else
            List2(2).Clear
        End If
    Case Else
    End Select
End Sub


Public Sub NewPictures()
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        rsBirthPictures.Recordset.Move 0
        rsBirthPictures.Recordset.AddNew
        boolNewRecord = True
        Text2(0).SetFocus
    Case 1
        rsInfancyPictures.Recordset.Move 0
        rsInfancyPictures.Recordset.AddNew
        boolNewRecord = True
        Text2(1).SetFocus
    Case 2
        rsChildPictures.Recordset.Move 0
        rsChildPictures.Recordset.AddNew
        boolNewRecord = True
        Text2(2).SetFocus
    Case Else
    End Select
End Sub

Public Function SelectPicChild() As Boolean
    On Error GoTo errSelectPicChild
    Sql = "SELECT * FROM ChildPictures WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsChildPictures.RecordSource = Sql
    rsChildPictures.Refresh
    rsChildPictures.Recordset.MoveFirst
    SelectPicChild = True
    Exit Function
    
errSelectPicChild:
    SelectPicChild = False
    Err.Clear
End Function
Public Function SelectPicInfant() As Boolean
    On Error GoTo errSelectPicInfant
    Sql = "SELECT * FROM InfancyPictures WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsInfancyPictures.RecordSource = Sql
    rsInfancyPictures.Refresh
    rsInfancyPictures.Recordset.MoveFirst
    SelectPicInfant = True
    Exit Function
    
errSelectPicInfant:
    SelectPicInfant = False
    Err.Clear
End Function
Public Sub Write_Print()
Dim rs As Recordset
    On Error Resume Next
    If PrintUseWord Then
        Call WritePicturesWord
    Else
        Call WritePictures
    End If
End Sub
Public Sub WritePictures()
    On Error Resume Next
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    bFirstWrite = True
    cPrint.pStartDoc
    
    Select Case Tab1.Tab
    Case 0  'birth
        sHeader = rsLanguage.Fields("FormName1")
        Call PrintFront
        
        With rsBirthPictures.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("ChildNo")) = CLng(glChildNo) And Not IsNull(.Fields("Picture")) Then
                    'print picture notes
                    If Not bFirstWrite Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    cPrint.pPrint Label2(0).Caption, 1, True
                    If Len(Text2(0).Text) <> 0 Then
                        cPrint.pMultiline Text2(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint
                    cPrint.pPrintPicture Picture1(0).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                    bFirstWrite = False
                End If
            .MoveNext
            Loop
        End With
    Case 1  'infant
        sHeader = rsLanguage.Fields("FormName2")
        Call PrintFront
        
        With rsInfancyPictures.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                    'print picture notes
                    If Not bFirstWrite Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.pPrint Label2(1).Caption, 1, True
                    If Len(Text2(1).Text) <> 0 Then
                        cPrint.pMultiline Text2(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint
                    If Not IsNull(.Fields("Picture")) Then
                        cPrint.pPrintPicture Picture1(1).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                    End If
                    bFirstWrite = False
                End If
            .MoveNext
            Loop
        End With
    Case 2  'childhood
        sHeader = rsLanguage.Fields("FormName3")
        Call PrintFront
        
        With frmPictures.rsChildPictures.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                    'print picture notes
                    If Not bFirstWrite Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.CurrentX = LeftMargin
                    cPrint.pPrint Label2(2).Caption, 1, True
                    If Len(Text2(2).Text) <> 0 Then
                        cPrint.pMultiline Text2(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint
                    If Not IsNull(.Fields("Picture")) Then
                        cPrint.pPrintPicture Picture1(2).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                    End If
                    bFirstWrite = False
                End If
            .MoveNext
            Loop
        End With
    Case Else
    End Select
    
    Screen.MousePointer = vbDefault
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub

Public Sub WritePicturesWord()
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0  'birth
        WriteHeader (rsLanguage.Fields("FormName1"))
        With wdApp
            .Selection.Tables(1).AllowAutoFit = False
            rsBirthPictures.Recordset.MoveFirst
            Do While Not rsBirthPictures.Recordset.EOF
                If CLng(rsBirthPictures.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                    'print the pictures with notes
                    .Selection.TypeText Text:=Format(rsBirthPictures.Recordset.Fields("PictureCaption"))
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    If Not IsNull(rsBirthPictures.Recordset.Fields("Picture")) Then
                        Clipboard.Clear
                        Clipboard.SetData Picture1(0).Picture, vbCFBitmap
                        .Selection.Paste
                    End If
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.InsertBreak Type:=wdPageBreak
                .Selection.MoveDown Unit:=wdLine, Count:=1
            rsBirthPictures.Recordset.MoveNext
            Loop
        End With
    Case 1  'infant
        WriteHeader (rsLanguage.Fields("FormName2"))
        With wdApp
            rsBirthPictures.Recordset.MoveFirst
            Do While Not rsInfancyPictures.Recordset.EOF
                If CLng(rsInfancyPictures.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                    'print the pictures with notes
                    .Selection.TypeText Text:=Format(rsInfancyPictures.Recordset.Fields("PictureCaption"))
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    If Not IsNull(rsInfancyPictures.Recordset.Fields("Picture")) Then
                        Clipboard.Clear
                        Clipboard.SetData Picture1(1).Picture, vbCFBitmap
                        .Selection.Paste
                    End If
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.InsertBreak Type:=wdPageBreak
                .Selection.MoveDown Unit:=wdLine, Count:=1
            rsInfancyPictures.Recordset.MoveNext
            Loop
        End With
    Case 2  'childhood
        WriteHeader (rsLanguage.Fields("FormName2"))
        With wdApp
            rsChildPictures.Recordset.MoveFirst
            Do While Not rsChildPictures.Recordset.EOF
                If CLng(rsChildPictures.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                    'print the pictures with notes
                    .Selection.TypeText Text:=Format(rsChildPictures.Recordset.Fields("PictureCaption"))
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    If Not IsNull(rsChildPictures.Recordset.Fields("Picture")) Then
                        'Picture2.Picture = LoadPicture(rsChildPictures.Recordset.Fields("Picture"))
                        Clipboard.Clear
                        Clipboard.SetData Picture1(2).Picture, vbCFBitmap
                        .Selection.Paste
                    End If
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.InsertBreak Type:=wdPageBreak
                .Selection.MoveDown Unit:=wdLine, Count:=1
            rsChildPictures.Recordset.MoveNext
            Loop
        End With
    Case Else
    End Select
    Set wdApp = Nothing
End Sub

Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1")
                    Label1(1).Caption = .Fields("label1")
                    Label1(2).Caption = .Fields("label1")
                End If
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2(0).Caption
                Else
                    Label2(0).Caption = .Fields("label2")
                    Label2(1).Caption = .Fields("label2")
                    Label2(2).Caption = .Fields("label2")
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
                If IsNull(.Fields("btnPastePicture")) Then
                    .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
                Else
                    btnPastePicture(0).ToolTipText = .Fields("btnPastePicture")
                    btnPastePicture(1).ToolTipText = .Fields("btnPastePicture")
                    btnPastePicture(2).ToolTipText = .Fields("btnPastePicture")
                End If
                If IsNull(.Fields("btnReadFromFile")) Then
                    .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
                Else
                    btnReadFromFile(0).ToolTipText = .Fields("btnReadFromFile")
                    btnReadFromFile(1).ToolTipText = .Fields("btnReadFromFile")
                    btnReadFromFile(2).ToolTipText = .Fields("btnReadFromFile")
                End If
                If IsNull(.Fields("btnCopyPic")) Then
                    .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
                Else
                    btnCopyPic(0).ToolTipText = .Fields("btnCopyPic")
                    btnCopyPic(1).ToolTipText = .Fields("btnCopyPic")
                    btnCopyPic(2).ToolTipText = .Fields("btnCopyPic")
                End If
                If IsNull(.Fields("btnScan")) Then
                    .Fields("btnScan") = btnScan(0).ToolTipText
                Else
                    btnScan(0).ToolTipText = .Fields("btnScan")
                    btnScan(1).ToolTipText = .Fields("btnScan")
                    btnScan(2).ToolTipText = .Fields("btnScan")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete(0).ToolTipText
                Else
                    btnDelete(0).ToolTipText = .Fields("btnDelete")
                    btnDelete(1).ToolTipText = .Fields("btnDelete")
                    btnDelete(2).ToolTipText = .Fields("btnDelete")
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
        .Fields("label1") = Label1(0).Caption
        .Fields("label2") = Label2(0).Caption
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        .Fields("btnDelete") = btnDelete(0).ToolTipText
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("FormName1") = "Birth Pictures"
        .Fields("FormName2") = "Infant Pictures"
        .Fields("FormName3") = "Childhood Pictures"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Public Sub FillList20()
    On Error Resume Next
    List2(0).Clear
    With rsBirthPictures.Recordset
        .MoveLast
        .MoveFirst
        ReDim BirthRecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(0).AddItem .Fields("AutoField")
            List2(0).ItemData(List2(0).NewIndex) = List2(0).ListCount - 1
            BirthRecordBookmark(List2(0).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Public Sub FillList21()
    On Error Resume Next
    List2(1).Clear
    With rsInfancyPictures.Recordset
        .MoveLast
        .MoveFirst
        ReDim InfantRecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(1).AddItem .Fields("AutoField")
            List2(1).ItemData(List2(1).NewIndex) = List2(1).ListCount - 1
            InfantRecordBookmark(List2(1).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Public Sub FillList22()
    On Error Resume Next
    List2(2).Clear
    With rsChildPictures.Recordset
        .MoveLast
        .MoveFirst
        ReDim ChildRecordBookmark(.RecordCount)
        Do While Not .EOF
            List2(2).AddItem .Fields("AutoField")
            List2(2).ItemData(List2(2).NewIndex) = List2(2).ListCount - 1
            ChildRecordBookmark(List2(2).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Function SelectPicBirth() As Boolean
    On Error GoTo errSelectPicBirth
    Sql = "SELECT * FROM BirthPictures WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBirthPictures.RecordSource = Sql
    rsBirthPictures.Refresh
    rsBirthPictures.Recordset.MoveFirst
    SelectPicBirth = True
    Exit Function
    
errSelectPicBirth:
    SelectPicBirth = False
    Err.Clear
End Function
Private Sub btnCopyPic_Click(Index As Integer)
    On Error Resume Next
    Clipboard.SetData Picture1(Index).Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        rsBirthPictures.Recordset.Delete
        FillList20
    Case 1
        rsInfancyPictures.Recordset.Delete
        FillList21
    Case 2
        rsChildPictures.Recordset.Delete
        FillList22
    Case Else
    End Select
End Sub

Private Sub btnPastePicture_Click(Index As Integer)
        On Error Resume Next
        Picture1(Index).Picture = Clipboard.GetData(vbCFDIB)
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
        Set Picture1(Index).Picture = LoadPicture(Cmd1.filename)
End Sub


Private Sub btnScan_Click(Index As Integer)
    Dim ret As Long, t As Single
    On Error Resume Next
    ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Picture1(Index).Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsChildPictures.Refresh
    rsInfancyPictures.Refresh
    rsBirthPictures.Refresh
    ShowText
    ShowAllButtons
    ShowKids
    SelectTab
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsChildPictures.DatabaseName = dbKidPicTxt
    rsInfancyPictures.DatabaseName = dbKidPicTxt
    rsBirthPictures.DatabaseName = dbKidPicTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmPictures")
    iWhichForm = 29
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmPictures: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsChildPictures.Recordset.Close
    rsInfancyPictures.Recordset.Close
    rsBirthPictures.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Erase BirthRecordBookmark
    Erase InfantRecordBookmark
    Erase ChildRecordBookmark
    Set frmPictures = Nothing
End Sub
Private Sub List2_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        rsBirthPictures.Recordset.Bookmark = BirthRecordBookmark(List2(0).ItemData(List2(0).ListIndex))
    Case 1
        rsInfancyPictures.Recordset.Bookmark = InfantRecordBookmark(List2(1).ItemData(List2(1).ListIndex))
    Case 2
        rsChildPictures.Recordset.Bookmark = ChildRecordBookmark(List2(2).ItemData(List2(2).ListIndex))
    Case Else
    End Select
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        If SelectPicBirth Then
            FillList20
            List2(0).ListIndex = 0
        Else
            List2(0).Clear
        End If
        Me.BackColor = &H8000&
        Tab1.BackColor = &H8000&
    Case 1
        If SelectPicInfant Then
            FillList21
            List2(1).ListIndex = 0
        Else
            List2(1).Clear
        End If
        Me.BackColor = &HC0C0&
        Tab1.BackColor = &HC0C0&
    Case 2
        If SelectPicChild Then
            FillList22
            List2(2).ListIndex = 0
        Else
            List2(2).Clear
        End If
        Me.BackColor = &H4040&
        Tab1.BackColor = &H4040&
    Case Else
    End Select
End Sub
Private Sub Text2_LostFocus(Index As Integer)
    On Error Resume Next
    If boolNewRecord Then
        Select Case Index
        Case 0
            With rsBirthPictures.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("BirthPictureCaption") = Text2(0).Text
                .Update
                .Bookmark = .LastModified
                FillList20
            End With
        Case 1
            With rsInfancyPictures.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("InfancyPictureCaption") = Text2(1).Text
                .Update
                .Bookmark = .LastModified
                FillList21
            End With
        Case 2
            With rsChildPictures.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("ChildPictureCaption") = Text2(1).Text
                .Update
                .Bookmark = .LastModified
                FillList22
            End With
        Case Else
        End Select
        boolNewRecord = False
    End If
End Sub
