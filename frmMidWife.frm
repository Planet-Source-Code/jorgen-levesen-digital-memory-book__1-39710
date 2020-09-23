VERSION 5.00
Begin VB.Form frmMidWife 
   BackColor       =   &H00800000&
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   8880
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Height          =   7335
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   5775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "Note"
         DataSource      =   "rsMidwife"
         Height          =   2325
         Index           =   11
         Left            =   1935
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   4920
         Width           =   3705
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "Email"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   10
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   10
         Top             =   4440
         Width           =   3705
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "FaxNo"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   9
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   9
         Top             =   4080
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "TelephoneNo"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   8
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   8
         Top             =   3720
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "Country"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   7
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3120
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "Town"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   6
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2760
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "ZipCode"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   5
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2400
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "Address2"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   4
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2040
         Width           =   3705
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "Address1"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   3
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1680
         Width           =   3705
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "LastName"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   2
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   0
         Top             =   360
         Width           =   2550
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "FirstName"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   1
         Left            =   1935
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   2550
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         DataField       =   "Proffesion"
         DataSource      =   "rsMidwife"
         Height          =   285
         Index           =   0
         Left            =   1935
         MaxLength       =   60
         TabIndex        =   2
         ToolTipText     =   "Doctor / Midwife / Nurse .."
         Top             =   1200
         Width           =   2550
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   4
         Left            =   4680
         Picture         =   "frmMidWife.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Notes :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   25
         Top             =   4920
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Email :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   24
         Top             =   4440
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Fax No.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   23
         Top             =   4080
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Telephone No.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   22
         Top             =   3720
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Country:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   21
         Top             =   3120
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Town:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   20
         Top             =   2760
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Zip Code:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   2400
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Last Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "First Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Profession:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Midwife(s) / Doctor(s)"
      ForeColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   2775
      Begin VB.Data rsMidwife 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Midwife"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6660
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmMidWife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v1RecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewMidwife()
    rsMidwife.Recordset.AddNew
    Text1(2).SetFocus
    boolNewRecord = True
End Sub


Public Sub WriteMidwifeWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("Frame1"))
    rsMidwife.Recordset.MoveFirst
    With wdApp
        Do While Not rsMidwife.Recordset.EOF
            .Selection.TypeText Text:=Label1(1).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("FirstName"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(2).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("LastName"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(0).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("Proffesion"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(3).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("Address1"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("Address2"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(4).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("ZipCode"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(5).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("Town"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(6).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("Country"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(7).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("TelephoneNo"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(8).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("FaxNo"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(9).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("Email"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label1(10).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(rsMidwife.Recordset.Fields("Note"))
            .Selection.MoveRight Unit:=wdCell
            'new page
            .Selection.InsertBreak Type:=wdPageBreak
            .Selection.MoveDown Unit:=wdLine, Count:=1
        rsMidwife.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub

Public Sub WriteMidwife()
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
    sHeader = rsLanguage.Fields("Frame1")
    
    cPrint.pStartDoc
    
    With rsMidwife.Recordset
        .MoveFirst
        Do While Not .EOF
            Call PrintFront
            cPrint.pPrint
            cPrint.pPrint Label1(1).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("FirstName")), 3.5
            cPrint.pPrint Label1(2).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("LastName")), 3.5
            cPrint.pPrint Label1(0).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("Proffesion")), 3.5
            cPrint.pPrint Label1(3).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("Address1")), 3.5
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("Address2")), 3.5
            cPrint.pPrint Label1(4).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("ZipCode")), 3.5
            cPrint.pPrint Label1(5).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("Town")), 3.5
            cPrint.pPrint Label1(6).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("Country")), 3.5
            cPrint.pPrint Label1(7).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("TelephoneNo")), 3.5
            cPrint.pPrint Label1(8).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("FaxNo")), 3.5
            cPrint.pPrint Label1(9).Caption, 1, True
            cPrint.pPrint Format(rsMidwife.Recordset.Fields("Email")), 3.5
            cPrint.pPrint Label1(10).Caption, 1, True
            cPrint.pMultiline Text1(11).Text, 2, cPrint.GetPaperWidth - 1.2, , False, True
        .MoveNext
        Loop
    End With
    
    Screen.MousePointer = vbDefault
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
End Sub

Private Sub ReadText()
Dim n As Integer, strMemo As String
    'find YOUR rsLanguage text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                For n = 0 To 10
                    If IsNull(.Fields(n + 1)) Then
                        .Fields(n + 1) = Label1(n).Caption
                    Else
                        Label1(n).Caption = .Fields(n + 1)
                    End If
                Next
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
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
        For n = 0 To 10
            .Fields(n + 1) = Label1(n).Caption
        Next
        .Fields("Frame1") = strMemo
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strMemo
        .Update
    End With
    DBEngine.Idle dbFreeLocks
End Sub

Public Sub FillList1()
    On Error Resume Next
    List1.Clear
    With rsMidwife.Recordset
        .MoveLast
        .MoveFirst
        ReDim v1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("FirstName") & "  " & .Fields("LastName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            v1RecordBookmark(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsMidwife.Refresh
    FillList1
    List1.ListIndex = 0
    ReadText
    ShowAllButtons
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsMidwife.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmMidWife")
    iWhichForm = 4
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmMidWife:  Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMidwife.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    iWhichForm = 0
    Set frmMidWife = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsMidwife.Recordset.Bookmark = v1RecordBookmark(List1.ItemData(List1.ListIndex))
End Sub
Private Sub Text1_LostFocus(Index As Integer)
        On Error Resume Next
        Select Case Index
        Case 1
            If boolNewRecord Then
                With rsMidwife.Recordset
                    .Fields("LastName") = Trim(Text1(2).Text)
                    .Fields("FirstName") = Trim(Text1(1).Text)
                    .Update
                    FillList1
                    .Bookmark = .LastModified
                    boolNewRecord = False
                End With
            End If
        Case Else
        End Select
End Sub
