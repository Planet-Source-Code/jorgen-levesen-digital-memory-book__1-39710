VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmToys 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toys to remember"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Infant (0- 1 year)"
      TabPicture(0)   =   "frmToys.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Cmd1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "List2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Childhood"
      TabPicture(1)   =   "frmToys.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "List3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   6135
         Index           =   1
         Left            =   -73440
         TabIndex        =   21
         Top             =   480
         Width           =   5535
         Begin VB.CommandButton btnDelete 
            Height          =   495
            Index           =   1
            Left            =   4920
            Picture         =   "frmToys.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Delete Picture"
            Top             =   5520
            Width           =   495
         End
         Begin VB.CommandButton btnReadFromFile 
            Height          =   495
            Index           =   1
            Left            =   4920
            Picture         =   "frmToys.frx":0182
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Read Picture from file"
            Top             =   4080
            Width           =   495
         End
         Begin VB.CommandButton btnPastePicture 
            Height          =   495
            Index           =   1
            Left            =   4920
            Picture         =   "frmToys.frx":0844
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Paste Picture"
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchasePrice"
            DataSource      =   "rsToysChild"
            Height          =   285
            Index           =   2
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   26
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ToyName"
            DataSource      =   "rsToysChild"
            Height          =   285
            Index           =   3
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   25
            Top             =   360
            Width           =   3375
         End
         Begin VB.CommandButton btnCopyPic 
            Height          =   495
            Index           =   1
            Left            =   4920
            Picture         =   "frmToys.frx":0F06
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Copy picture to the Clipboard"
            Top             =   4560
            Width           =   495
         End
         Begin VB.CommandButton btnScan 
            Height          =   495
            Index           =   1
            Left            =   4920
            Picture         =   "frmToys.frx":15C8
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Scan a picture"
            Top             =   5040
            Width           =   495
         End
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchaseDate"
            DataSource      =   "rsToysChild"
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
            Left            =   1920
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "ToyNote"
            DataSource      =   "rsToysChild"
            Height          =   1935
            Index           =   1
            Left            =   1920
            TabIndex        =   30
            Top             =   1560
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   3413
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmToys.frx":1712
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Picture"
            DataSource      =   "rsToysChild"
            Height          =   2415
            Index           =   1
            Left            =   2520
            Stretch         =   -1  'True
            Top             =   3600
            Width           =   2295
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Toy Name:"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   36
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchased Date:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchase Price:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   34
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Picture:"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   33
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Image Image2 
            Height          =   1455
            Index           =   1
            Left            =   240
            Picture         =   "frmToys.frx":17CC
            Stretch         =   -1  'True
            Top             =   4320
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Toy note:"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   32
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   31
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6075
         Left            =   -74880
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Height          =   6135
         Index           =   0
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   5535
         Begin VB.Data rsToysInfant 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   4320
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "ToysToRememberInfant"
            Top             =   0
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Data rsToysChild 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   3240
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "ToysToRememberChild"
            Top             =   0
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ToyName"
            DataSource      =   "rsToysInfant"
            Height          =   285
            Index           =   0
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   13
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchasePrice"
            DataSource      =   "rsToysInfant"
            Height          =   285
            Index           =   1
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   12
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton btnDelete 
            Height          =   495
            Index           =   0
            Left            =   4920
            Picture         =   "frmToys.frx":3D2C
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Delete picture"
            Top             =   5520
            Width           =   495
         End
         Begin VB.CommandButton btnReadFromFile 
            Height          =   495
            Index           =   0
            Left            =   4920
            Picture         =   "frmToys.frx":3E76
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Read picture from file"
            Top             =   4080
            Width           =   495
         End
         Begin VB.CommandButton btnPastePicture 
            Height          =   495
            Index           =   0
            Left            =   4920
            Picture         =   "frmToys.frx":4538
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Paste picture"
            Top             =   3600
            Width           =   495
         End
         Begin VB.CommandButton btnScan 
            Height          =   495
            Index           =   4
            Left            =   5520
            Picture         =   "frmToys.frx":4BFA
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Scan a picture"
            Top             =   4440
            Width           =   495
         End
         Begin VB.CommandButton btnCopyPic 
            Height          =   495
            Index           =   0
            Left            =   4920
            Picture         =   "frmToys.frx":4D44
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Copy picture to the Clipboard"
            Top             =   4560
            Width           =   495
         End
         Begin VB.CommandButton btnScan 
            Height          =   495
            Index           =   0
            Left            =   4920
            Picture         =   "frmToys.frx":5406
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Scan a picture"
            Top             =   5040
            Width           =   495
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            ScaleHeight     =   345
            ScaleWidth      =   465
            TabIndex        =   4
            Top             =   3600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchaseDate"
            DataSource      =   "rsToysInfant"
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
            Left            =   1920
            TabIndex        =   3
            Top             =   720
            Width           =   1335
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "ToyNote"
            DataSource      =   "rsToysInfant"
            Height          =   1935
            Index           =   0
            Left            =   1920
            TabIndex        =   5
            Top             =   1560
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   3413
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmToys.frx":5550
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Toy Name:"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchased Date:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchased Price:"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   17
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Picture"
            DataSource      =   "rsToysInfant"
            Height          =   2415
            Index           =   0
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   3600
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Picture:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   3600
            Width           =   2175
         End
         Begin VB.Image Image2 
            Height          =   1455
            Index           =   0
            Left            =   120
            Picture         =   "frmToys.frx":560A
            Stretch         =   -1  'True
            Top             =   4440
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Toy note:"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   15
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   14
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6075
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog Cmd1 
         Left            =   6600
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmToys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ChildRecordBookmark() As Variant
Dim InfantRecordBookmark() As Variant
Dim rsLanguage As Recordset
Dim rsMyRecord As Recordset
Private Sub SelectTab()
    On Error Resume Next
    With Me
        Select Case iTab
        Case 0
            .Tab1.Tab = 0
            .BackColor = &HC0C0&
            .Tab1.BackColor = &HC0C0&
        Case 1
            .Tab1.Tab = 1
            .BackColor = &H4040&
            .Tab1.BackColor = &H4040&
        Case Else
        End Select
    End With
End Sub

Public Sub NewChildToys()
    On Error Resume Next
    Frame1(0).Caption = gsChildName
    Frame1(1).Caption = gsChildName
    SelectToys
    FillList2
    List2.ListIndex = 0
    FillList3
    List3.ListIndex = 0
End Sub

Public Sub NewToys()
    Select Case Tab1.Tab
    Case 0  'infant
        rsToysInfant.Recordset.AddNew
        Text1(0).SetFocus
    Case 1  'child
        rsToysChild.Recordset.AddNew
        Text1(3).SetFocus
    Case Else
    End Select
    boolNewRecord = True
End Sub

Public Sub DeleteToys()
    On Error Resume Next
    Select Case Tab1.Tab
        Case 0
            rsToysInfant.Recordset.Delete
            FillList2
        Case 1
            rsToysChild.Recordset.Delete
            FillList3
    Case Else
    End Select
End Sub

Public Sub WriteToysWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("Form"))
    With wdApp
        .Selection.Tables(1).AllowAutoFit = False
        'infant
        rsToysInfant.Recordset.MoveFirst
        Do While Not rsToysInfant.Recordset.EOF
            .Selection.TypeText Text:=Label1(0).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsToysInfant.Recordset.Fields("ToyName"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(1).Caption
            .Selection.MoveRight Unit:=wdCell
            If IsDate(rsToysInfant.Recordset.Fields("PurchaseDate")) Then
                .Selection.TypeText Text:=Format(CDate(rsToysInfant.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(2).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(CDbl(rsToysInfant.Recordset.Fields("PurchasePrice")), "0.00")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(4).Caption
            .Selection.MoveRight Unit:=wdCell
            Clipboard.Clear
            Clipboard.SetText RichTextBox1(0).TextRTF, vbCFRTF
            .Selection.Paste
            DoEvents
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(3).Caption
            .Selection.MoveRight Unit:=wdCell
            If Not IsNull(rsToysInfant.Recordset.Fields("Picture")) Then
                Picture1.Picture = Image1(0).Picture
                Clipboard.Clear
                Clipboard.SetData Picture1.Picture, vbCFBitmap
                .Selection.Paste
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
        rsToysInfant.Recordset.MoveNext
        Loop
        
        .Selection.MoveRight Unit:=wdCell
        .Selection.InsertBreak Type:=wdPageBreak
        .ActiveWindow.Selection.Font.Name = "Monotype Corsiva"
        .ActiveWindow.Selection.Font.Size = 28
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .Selection.TypeText Text:=rsLanguage.Fields("FormName2")
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .ActiveWindow.Selection.Font.Name = "Times New Roman"
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.MoveDown Unit:=wdLine, Count:=1
        
        'childhood
        rsToysChild.Recordset.MoveFirst
        Do While Not rsToysChild.Recordset.EOF
            .Selection.TypeText Text:=Label1(0).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsToysChild.Recordset.Fields("ToyName"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(1).Caption
            .Selection.MoveRight Unit:=wdCell
            If IsDate(rsToysChild.Recordset.Fields("PurchaseDate")) Then
                .Selection.TypeText Text:=Format(CDate(rsToysChild.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(2).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(CDbl(rsToysChild.Recordset.Fields("PurchasePrice")), "0.00")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(4).Caption
            .Selection.MoveRight Unit:=wdCell
            Clipboard.Clear
            Clipboard.SetText RichTextBox1(1).TextRTF, vbCFRTF
            .Selection.Paste
            DoEvents
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(3).Caption
            .Selection.MoveRight Unit:=wdCell
            If Not IsNull(rsToysChild.Recordset.Fields("Picture")) Then
                Picture1.Picture = Image1(1).Picture
                Clipboard.Clear
                Clipboard.SetData Picture1.Picture, vbCFBitmap
                .Selection.Paste
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
        rsToysChild.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub

Public Sub WriteToys()
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
    cPrint.pStartDoc
    
    Select Case Tab1.Tab
    Case 0
        Call PrintFront
        sHeader = rsLanguage.Fields("Form") & "- " & rsLanguage.Fields("FormName1")
        
        With rsToysInfant.Recordset
            .MoveFirst
            Do While Not .EOF
                cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
                If Not IsNull(.Fields("ToyName")) Then
                    cPrint.pPrint .Fields("ToyName"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
                If IsDate(.Fields("PurchaseDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 1.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
                cPrint.pPrint Format(CDbl(.Fields("PurchasePrice")), "0.00") & "  " & Label4(0).Caption, 3.5
                cPrint.pPrint
                cPrint.pPrint rsLanguage.Fields("label1(4)"), 1
                If Len(RichTextBox1(0).Text) <> 0 Then
                    cPrint.pMultiline RichTextBox1(0).Text, 1, cPrint.GetPaperWidth - 1, , False, True
                Else
                    cPrint.pPrint " ", 1.5
                End If
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint rsLanguage.Fields("label1(3)"), 1
                If Not IsNull(.Fields("Picture")) Then
                    cPrint.pPrintPicture Image1(0).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                End If
                cPrint.pPrint
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
            .MoveNext
            Loop
        End With
        
    Case 1
    
        sHeader = rsLanguage.Fields("Form") & "- " & rsLanguage.Fields("FormName2")
        Call PrintFront
            
        With rsToysChild.Recordset
            .MoveFirst
            Do While Not .EOF
                cPrint.pPrint Label1(0).Caption, 1, True
                If Not IsNull(.Fields("ToyName")) Then
                    cPrint.pPrint .Fields("ToyName"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint Label1(1).Caption, 1, True
                If IsDate(.Fields("PurchaseDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 1.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(2).Caption, 1, True
                cPrint.pPrint Format(CDbl(.Fields("PurchasePrice")), "0.00") & "  " & Label4(0).Caption, 3.5
                cPrint.pPrint
                cPrint.pPrint Label1(4).Caption, 1
                If Len(RichTextBox1(1).Text) <> 0 Then
                    cPrint.pMultiline RichTextBox1(1).Text, 1, cPrint.GetPaperWidth - 1, , False, True
                Else
                    cPrint.pPrint " ", 1.5
                End If
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(3).Caption, 1
                If Not IsNull(.Fields("Picture")) Then
                    cPrint.pPrintPicture Image1(1).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                End If
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint
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

Public Sub FillList3()
    On Error Resume Next
    List3.Clear
    If rsToysChild.Recordset.RecordCount <> 0 Then
        With rsToysChild.Recordset
            .MoveLast
            .MoveFirst
            ReDim ChildRecordBookmark(.RecordCount)
            Do While Not .EOF
                List3.AddItem .Fields("ToyName")
                List3.ItemData(List3.NewIndex) = List3.ListCount - 1
                ChildRecordBookmark(List3.ListCount - 1) = .Bookmark
            .MoveNext
            Loop
        End With
    End If
End Sub

Public Sub FillList2()
    On Error Resume Next
    List2.Clear
    If rsToysInfant.Recordset.RecordCount <> 0 Then
        With rsToysInfant.Recordset
            .MoveLast
            .MoveFirst
            ReDim InfantRecordBookmark(.RecordCount)
            Do While Not .EOF
                List2.AddItem .Fields("ToyName")
                List2.ItemData(List2.NewIndex) = List2.ListCount - 1
                InfantRecordBookmark(List2.ListCount - 1) = .Bookmark
            .MoveNext
            Loop
        End With
    End If
End Sub

Public Sub SelectToys()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM ToysToRememberInfant WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsToysInfant.RecordSource = Sql
    rsToysInfant.Refresh
    
    Sql = "SELECT * FROM ToysToRememberChild WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsToysChild.RecordSource = Sql
    rsToysChild.Refresh
End Sub
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
                For i = 0 To 4
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                        Label2(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("btnPastePicture")) Then
                    .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
                Else
                    btnPastePicture(0).ToolTipText = .Fields("btnPastePicture")
                    btnPastePicture(1).ToolTipText = .Fields("btnPastePicture")
                End If
                If IsNull(.Fields("btnReadFromFile")) Then
                    .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
                Else
                    btnReadFromFile(0).ToolTipText = .Fields("btnReadFromFile")
                    btnReadFromFile(1).ToolTipText = .Fields("btnReadFromFile")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete(0).ToolTipText
                Else
                    btnDelete(0).ToolTipText = .Fields("btnDelete")
                    btnDelete(1).ToolTipText = .Fields("btnDelete")
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
        .Fields("Form") = Me.Caption
        For i = 0 To 4
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnDelete") = btnDelete(0).ToolTipText
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("FormName1") = "Infant"
        .Fields("FormName2") = "Childhood"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Private Sub btnCopyPic_Click(Index As Integer)
    On Error Resume Next
    Clipboard.SetData Image1(Index).Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click(Index As Integer)
    On Error Resume Next
    Set Image1(Index).Picture = LoadPicture()
End Sub

Private Sub btnPastePicture_Click(Index As Integer)
    On Error Resume Next
    Image1(Index).Picture = Clipboard.GetData(vbCFDIB)
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
        Set Image1(Index).Picture = LoadPicture(Cmd1.filename)
End Sub

Private Sub btnScan_Click(Index As Integer)
    Dim ret As Long, t As Single
    On Error Resume Next
    ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Image1(Index).Picture = Clipboard.GetData(vbCFDIB)
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
    rsToysInfant.Refresh
    rsToysChild.Refresh
    Label4(0).Caption = rsMyRecord.Fields("Currency")
    Label4(1).Caption = rsMyRecord.Fields("Currency")
    SelectToys
    FillList2
    List2.ListIndex = 0
    FillList3
    List3.ListIndex = 0
    ShowText
    ShowAllButtons
    ShowKids
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsToysInfant.DatabaseName = dbKidsTxt
    rsToysChild.DatabaseName = dbKidsTxt
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmToys")
    iWhichForm = 22
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmToys: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsToysInfant.Recordset.Close
    rsToysChild.Recordset.Close
    rsMyRecord.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmToys = Nothing
End Sub

Private Sub List2_Click()
    On Error Resume Next
    rsToysInfant.Recordset.Bookmark = InfantRecordBookmark(List2.ItemData(List2.ListIndex))
End Sub

Private Sub List3_Click()
    On Error Resume Next
    rsToysChild.Recordset.Bookmark = ChildRecordBookmark(List3.ItemData(List3.ListIndex))
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

Private Sub RichTextBox1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox1_SelChange(Index As Integer)
    Call RichTextSelChange(frmToys.RichTextBox1(Index))
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab1.Tab
        Case 0
            Me.BackColor = &HC0C0&
            Tab1.BackColor = &HC0C0&
        Case 1
            Me.BackColor = &H4040&
            Tab1.BackColor = &H4040&
    Case Else
    End Select
End Sub
Private Sub Text1_LostFocus(Index As Integer)
   On Error Resume Next
   If boolNewRecord Then
        Select Case Index
        Case 0
            With rsToysInfant.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("ToyName") = Trim(Text1(0).Text)
                .Update
                FillList2
                .Bookmark = .LastModified
            End With
        Case 3
            With rsToysChild.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("ToyName") = Trim(Text1(3).Text)
                .Update
                FillList3
                .Bookmark = .LastModified
            End With
        Case Else
        End Select
    End If
End Sub


