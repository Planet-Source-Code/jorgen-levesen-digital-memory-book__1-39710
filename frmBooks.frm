VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBooks 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Books To Remember"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   49344
      TabCaption(0)   =   "Infant (0 -1 year)"
      TabPicture(0)   =   "frmBooks.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "List2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Child"
      TabPicture(1)   =   "frmBooks.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "List3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   5775
         Left            =   -73200
         TabIndex        =   26
         Top             =   360
         Width           =   5175
         Begin VB.TextBox Date2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchaseDate"
            DataSource      =   "rsBooksChild"
            Height          =   285
            Left            =   2040
            TabIndex        =   11
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BookName"
            DataSource      =   "rsBooksChild"
            Height          =   285
            Index           =   5
            Left            =   2040
            MaxLength       =   100
            TabIndex        =   7
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BookAuthor"
            DataSource      =   "rsBooksChild"
            Height          =   285
            Index           =   7
            Left            =   2040
            MaxLength       =   80
            TabIndex        =   8
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BookIsbn"
            DataSource      =   "rsBooksChild"
            Height          =   285
            Index           =   8
            Left            =   2040
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   3840
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchasePrice"
            DataSource      =   "rsBooksChild"
            Height          =   285
            Index           =   9
            Left            =   2040
            MaxLength       =   80
            TabIndex        =   12
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WherePurchased"
            DataSource      =   "rsBooksChild"
            Height          =   645
            Index           =   6
            Left            =   2040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   5040
            Width           =   3015
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "BookSynopsis"
            DataSource      =   "rsBooksChild"
            Height          =   2655
            Index           =   1
            Left            =   2040
            TabIndex        =   9
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   4683
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmBooks.frx":0038
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Book Name:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Author:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Synopsis:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "ISBN Number:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchase Date:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   30
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchase Price:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   29
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Image Image1 
            Height          =   1455
            Index           =   1
            Left            =   240
            Picture         =   "frmBooks.frx":010D
            Stretch         =   -1  'True
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchased Where:"
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   28
            Top             =   5040
            Width           =   1815
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   27
            Top             =   4680
            Width           =   615
         End
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5685
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   5775
         Left            =   1920
         TabIndex        =   16
         Top             =   360
         Width           =   5175
         Begin VB.TextBox Date1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchaseDate"
            DataSource      =   "rsBooksInfant"
            Height          =   285
            Left            =   2040
            TabIndex        =   4
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Data rsBooksInfant 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   720
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "BooksToRememberInfant"
            Top             =   0
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Data rsBooksChild 
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
            RecordSource    =   "BooksToRememberChild"
            Top             =   0
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BookName"
            DataSource      =   "rsBooksInfant"
            Height          =   285
            Index           =   0
            Left            =   2040
            MaxLength       =   100
            TabIndex        =   0
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BookAuthor"
            DataSource      =   "rsBooksInfant"
            Height          =   285
            Index           =   2
            Left            =   2040
            MaxLength       =   80
            TabIndex        =   1
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BookIsbn"
            DataSource      =   "rsBooksInfant"
            Height          =   285
            Index           =   3
            Left            =   2040
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   3
            Top             =   3840
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PurchasePrice"
            DataSource      =   "rsBooksInfant"
            Height          =   285
            Index           =   4
            Left            =   2040
            MaxLength       =   80
            TabIndex        =   5
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "WherePurchased"
            DataSource      =   "rsBooksInfant"
            Height          =   645
            Index           =   1
            Left            =   2040
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   5040
            Width           =   3015
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "BookSynopsis"
            DataSource      =   "rsBooksInfant"
            Height          =   2655
            Index           =   0
            Left            =   2040
            TabIndex        =   2
            Top             =   1080
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   4683
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmBooks.frx":200C
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Book Name:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Author:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Synopsis:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "ISBN Number:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   21
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchase Date:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   20
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchase Price:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Image Image1 
            Height          =   1455
            Index           =   0
            Left            =   240
            Picture         =   "frmBooks.frx":20E1
            Stretch         =   -1  'True
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchased Where:"
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   18
            Top             =   5040
            Width           =   1815
         End
         Begin VB.Label Label4 
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   17
            Top             =   4680
            Width           =   615
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5685
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v1RecordBookmark() As Variant
Dim InfantRecordBookmark() As Variant
Dim ChildRecordBookmark() As Variant
Dim rsMyRecord As Recordset
Dim rsLanguage As Recordset
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

Public Sub NewChildBooks()
    On Error Resume Next
    Frame1.Caption = gsChildName
    Frame2.Caption = gsChildName
    SelectBooksChild
    Select Case Tab1.Tab
    Case 0
        FillList2
        List2.ListIndex = 0
    Case 1
        FillList3
        List3.ListIndex = 0
    Case Else
    End Select
End Sub

Public Sub NewBooks()
    Select Case Tab1.Tab
    Case 0  'infant
        rsBooksInfant.Recordset.AddNew
        Text1(0).SetFocus
    Case 1  'child
        rsBooksChild.Recordset.AddNew
        Text1(5).SetFocus
    Case Else
    End Select
    boolNewRecord = True
End Sub

Public Sub DeleteBooks()
    On Error Resume Next
    Select Case Tab1.Tab
        Case 0
            rsBooksInfant.Recordset.Delete
            FillList2
        Case 1
            rsBooksChild.Recordset.Delete
            FillList3
    Case Else
    End Select
End Sub

Public Sub FillList2()
    On Error Resume Next
    List2.Clear
    With rsBooksInfant.Recordset
        .MoveLast
        .MoveFirst
        ReDim InfantRecordBookmark(.RecordCount)
        Do While Not .EOF
            List2.AddItem .Fields("BookName")
            List2.ItemData(List2.NewIndex) = List2.ListCount - 1
            InfantRecordBookmark(List2.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub FillList3()
    On Error Resume Next
    List3.Clear
    With rsBooksChild.Recordset
        .MoveLast
        .MoveFirst
        ReDim ChildRecordBookmark(.RecordCount)
        Do While Not .EOF
            List3.AddItem .Fields("BookName")
            List3.ItemData(List3.NewIndex) = List3.ListCount - 1
            ChildRecordBookmark(List3.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Public Sub PrintBooks()
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
    cPrint.pStartDoc
    
    Select Case Tab1.Tab
    Case 0  'infant
        sHeader = rsLanguage.Fields("FormName")
        Call PrintFront
        With rsBooksInfant.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                    cPrint.FontBold = True
                    cPrint.pPrint Label1(0).Caption, 1, True
                    If Len(Text1(0).Text) <> 0 Then
                        cPrint.pPrint Text1(0).Text, 3.5
                    Else
                        cPrint.pPrint "", 3.5
                    End If
                    cPrint.FontBold = False
                    cPrint.pPrint Label1(1).Caption, 1, True
                    If Len(Text1(2).Text) <> 0 Then
                        cPrint.pPrint Text1(2).Text, 3.5
                    Else
                        cPrint.pPrint "", 3.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.pPrint Label1(2).Caption, 1, True
                    If Len(RichTextBox1(0).Text) <> 0 Then
                        cPrint.pMultiline RichTextBox1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint "", 3.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.pPrint Label1(3).Caption, 1, True
                    If Len(Text1(3).Text) <> 0 Then
                        cPrint.pPrint Text1(3).Text, 3.5
                    Else
                        cPrint.pPrint "", 3.5
                    End If
                    cPrint.pPrint Label1(4).Caption, 1, True
                    If IsDate(.Recordset.Fields("PurchaseDate")) Then
                        cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.CurrentX = LeftMargin
                    cPrint.pPrint Label1(5).Caption, 1, True
                    cPrint.pPrint Format(CDbl(Text1(4).Text), "0.00") & "  " & Label4(0).Caption, 3.5
                    cPrint.pPrint Label1(6).Caption, 1, True
                    If Len(Text1(1).Text) <> 0 Then
                        cPrint.pMultiline Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 1.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                End If
            .MoveNext
            Loop
        End With
        
    Case 1  'child
        sHeader = rsLanguage.Fields("FormName") & " - " & rsLanguage.Fields("FormName2")
        Call PrintFront
            
        With rsBooksChild.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                    cPrint.FontBold = True
                    cPrint.pPrint Label1(0).Caption, 1, True
                    If Len(Text1(5).Text) <> 0 Then
                        cPrint.pPrint Text1(5).Text, 3.5
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.FontBold = False
                    cPrint.pPrint Label1(1).Caption, 1, True
                    If Len(Text1(7).Text) <> 0 Then
                        cPrint.pPrint Text1(7).Text, 3.5
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint Label1(2).Caption, 1, True
                    If Len(RichTextBox1(1).Text) <> 0 Then
                        cPrint.pMultiline RichTextBox1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint Label1(3).Caption, 1
                    If Len(Text1(8).Text) <> 0 Then
                        cPrint.pPrint Text1(8).Text, 3.5
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.pPrint Label1(4).Caption, 1, True
                    If IsDate(.Recordset.Fields("PurchaseDate")) Then
                        cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                    cPrint.pPrint Label1(5).Caption, 1, True
                    cPrint.pPrint Format(CDbl(Text1(9).Text), "0.00") & "  " & Label4(0).Caption, 3.5
                    cPrint.pPrint Label1(6).Caption, 1, True
                    If Len(Text1(6).Text) <> 0 Then
                        cPrint.pMultiline Text1(6).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
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

Public Sub PrintBooksWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName1"))
    With wdApp
        rsBooksInfant.Recordset.MoveFirst
        Do While Not rsBooksInfant.Recordset.EOF
            If CLng(rsBooksInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksInfant.Recordset.Fields("BookName"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(1).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksInfant.Recordset.Fields("BookAuthor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(2).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText RichTextBox1(0).TextRTF, vbCFRTF
                .Selection.Paste
                DoEvents
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(3).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksInfant.Recordset.Fields("BookIsbn"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(4).Caption
                .Selection.MoveRight Unit:=wdCell
                If IsDate(rsBooksInfant.Recordset.Fields("PurchaseDate")) Then
                    .Selection.TypeText Text:=Format(CDate(rsBooksInfant.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
                Else
                    .Selection.TypeText Text:=""
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(5).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksInfant.Recordset.Fields("PurchasePrice"), "0.00")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(6).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksInfant.Recordset.Fields("WherePurchased"))
            End If
        rsBooksInfant.Recordset.MoveNext
        Loop

    WriteHeader (rsLanguage.Fields("FormName2"))
        
        rsBooksChild.Recordset.MoveFirst
        Do While Not rsBooksChild.Recordset.EOF
            If CLng(rsBooksChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksChild.Recordset.Fields("BookName"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(1).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksChild.Recordset.Fields("BookAuthor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(2).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText RichTextBox1(1).TextRTF, vbCFRTF
                .Selection.Paste
                DoEvents
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(3).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksChild.Recordset.Fields("BookIsbn"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(4).Caption
                .Selection.MoveRight Unit:=wdCell
                If IsDate(rsBooksChild.Recordset.Fields("PurchaseDate")) Then
                    .Selection.TypeText Text:=Format(CDate(rsBooksChild.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
                Else
                    .Selection.TypeText Text:=""
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(5).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksChild.Recordset.Fields("PurchasePrice"), "0.00")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(6).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsBooksChild.Recordset.Fields("WherePurchased"))
            End If
        rsBooksChild.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub

Public Sub SelectBooksChild()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM BooksToRememberChild WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBooksChild.RecordSource = Sql
    rsBooksChild.Refresh
    
    Sql = "SELECT * FROM BooksToRememberInfant WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBooksInfant.RecordSource = Sql
    rsBooksInfant.Refresh
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
                For i = 0 To 6
                If IsNull(.Fields(i + 2)) Then
                    .Fields(i + 2) = Label1(i).Caption
                Else
                    Label1(i).Caption = .Fields(i + 2)
                    Label2(i).Caption = .Fields(i + 2)
                End If
                Next
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
        .Fields("Form") = Me.Caption
        For i = 0 To 6
            .Fields(i + 2) = Label1(i).Caption
        Next
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("FormName") = Me.Caption
        .Fields("FormName1") = "Infant"
        .Fields("FormName2") = "Childhood"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
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

Private Sub Date2_Click()
Dim UserDate As Date
    If IsDate(Date2.Text) Then
        UserDate = CVDate(Date2.Text)
    Else
        UserDate = Format(Now, "dd.mm.yyyy")
    End If
    If frmCalendar.GetDate(UserDate) Then
        Date2.Text = UserDate
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsBooksInfant.Refresh
    rsBooksChild.Refresh
    Label4(0).Caption = rsMyRecord.Fields("Currency")
    Label4(1).Caption = rsMyRecord.Fields("Currency")
    ShowText
    Frame1.Caption = gsChildName
    Frame2.Caption = gsChildName
    SelectTab
    SelectBooksChild
    FillList2
    List2.ListIndex = 0
    FillList3
    List3.ListIndex = 0
    ShowAllButtons
    ShowKids
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsBooksInfant.DatabaseName = dbKidsTxt
    rsBooksChild.DatabaseName = dbKidsTxt
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmBooks")
    iWhichForm = 21
    Exit Sub
    
errForm_Load:
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmBooks: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsBooksInfant.Recordset.Close
    rsBooksChild.Recordset.Close
    rsMyRecord.Close
    rsLanguage.Close
    iWhichForm = 0
    HideAllButtons
    HideKids
    Set frmBooks = Nothing
End Sub

Public Sub List2_Click()
    On Error Resume Next
    rsBooksInfant.Recordset.Bookmark = InfantRecordBookmark(List2.ItemData(List2.ListIndex))
End Sub


Public Sub List3_Click()
    On Error Resume Next
    rsBooksChild.Recordset.Bookmark = ChildRecordBookmark(List3.ItemData(List3.ListIndex))
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
   On Error Resume Next
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox1_SelChange(Index As Integer)
    On Error Resume Next
    Call RichTextSelChange(frmBooks.RichTextBox1(Index))
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab1.Tab
        Case 0
            FillList2
            List2.ListIndex = 0
            Me.BackColor = &HC0C0&
            Tab1.BackColor = &HC0C0&
        Case 1
            FillList3
            List3.ListIndex = 0
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
            With rsBooksInfant.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("BookName") = Trim(Text1(0).Text)
                .Update
                FillList2
                .Bookmark = .LastModified
            End With
        Case 5
            With rsBooksChild.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("BookName") = Trim(Text1(5).Text)
                .Update
                FillList3
                .Bookmark = .LastModified
            End With
        Case Else
        End Select
        boolNewRecord = False
    End If
End Sub
