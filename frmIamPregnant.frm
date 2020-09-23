VERSION 5.00
Begin VB.Form frmIamPregnant 
   BackColor       =   &H000080FF&
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   8880
   Begin VB.Data rsIamPregnant 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "IamPregnant"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Later persons to know"
      Height          =   3975
      Index           =   4
      Left            =   4200
      TabIndex        =   15
      Top             =   4320
      Width           =   4575
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LaterPersonsReactions"
         DataSource      =   "rsIamPregnant"
         Height          =   2445
         Index           =   6
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1440
         Width           =   4350
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LaterPersons"
         DataSource      =   "rsIamPregnant"
         Height          =   645
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   480
         Width           =   4350
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Reaction(s):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Person(s) name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "First person to know"
      Height          =   3975
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   3975
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FirstPersonReaction"
         DataSource      =   "rsIamPregnant"
         Height          =   2805
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FirstPersonToKnow"
         DataSource      =   "rsIamPregnant"
         Height          =   285
         Index           =   3
         Left            =   120
         MaxLength       =   60
         TabIndex        =   4
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Reaction:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Person name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "First Signs"
      Height          =   1815
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   8655
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FirstSign"
         DataSource      =   "rsIamPregnant"
         Height          =   1245
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   480
         Width           =   7425
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   7680
         Picture         =   "frmIamPregnant.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Test Taken"
      Height          =   2175
      Index           =   2
      Left            =   5280
      TabIndex        =   10
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TestTaken"
         DataSource      =   "rsIamPregnant"
         Height          =   1770
         Index           =   1
         Left            =   120
         MaxLength       =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Doctor / Nurse Name"
      Height          =   2055
      Index           =   1
      Left            =   2400
      TabIndex        =   9
      Top             =   240
      Width           =   2775
      Begin VB.ComboBox cmbName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "DoctorName"
         DataSource      =   "rsIamPregnant"
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   2460
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Confirmation Date"
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2175
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ConfirmationDate"
         DataSource      =   "rsIamPregnant"
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
         Left            =   480
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   615
         Index           =   1
         Left            =   1320
         Picture         =   "frmIamPregnant.frx":3026A
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "frmIamPregnant.frx":32060
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmIamPregnant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMidwife As Recordset
Dim rsLanguage As Recordset
Public Sub IamPregnantPrintWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.TypeText Text:=Frame1(0).Caption & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=CDate(rsIamPregnant.Recordset.Fields("ConfirmationDate")) & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame1(3).Caption & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsIamPregnant.Recordset.Fields("TestTaken")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame1(2).Caption & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsIamPregnant.Recordset.Fields("DoctorName") & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame1(4).Caption & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsIamPregnant.Recordset.Fields("FirstSign")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame1(5).Caption & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsIamPregnant.Recordset.Fields("FirstPersonToKnow") & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsIamPregnant.Recordset.Fields("FirstPersonReaction") & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame1(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsIamPregnant.Recordset.Fields("LaterPersons") & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsIamPregnant.Recordset.Fields("LaterPersonsReactions") & ""
    End With
    Set wdApp = Nothing
End Sub

Public Sub IamPregnantPrint()
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
    
    With rsIamPregnant.Recordset
        cPrint.pPrint rsLanguage.Fields("Frame1(0)") & ":", 1, True 'confirmation date
        If IsDate(.Fields("ConfirmationDate")) Then
            cPrint.pPrint .Fields("ConfirmationDate"), 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(2)") & ":", 1, True 'test taken
        If Len(Text1(1).Text) <> 0 Then
            cPrint.pPrint Text1(1).Text, 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            Call PrintFront
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(1)") & ":", 1, True 'doctor / nurse name
        If Len(cmbName.Text) <> 0 Then
            cPrint.pPrint cmbName.Text, 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            Call PrintFront
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(3)") & ":", 1, True 'first signs
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
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(5)") & ":", 1, True 'first persons to know
        If Len(Text1(3).Text) <> 0 Then
            cPrint.pPrint Text1(3).Text, 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            Call PrintFront
        End If
        cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True   'reaction
        If Len(Text1(4).Text) <> 0 Then
            cPrint.pMultiline Text1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            Call PrintFront
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(4)") & ":", 1, True 'later persons to know
        If Len(Text1(5).Text) <> 0 Then
            cPrint.pMultiline Text1(5).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            Call PrintFront
        End If
        cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True
        If Len(Text1(6).Text) <> 0 Then
            cPrint.pMultiline Text1(6).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint "", 3.5
        End If
    End With
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub

Public Sub NewPregnancy()
    rsIamPregnant.Recordset.AddNew
    Date1.SetFocus
    boolNewRecord = True
End Sub

Public Function SelectPregnancy() As Boolean
Dim Sql As String
    On Error GoTo errSelectPregnancy
    Sql = "SELECT * FROM  IamPregnant  WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsIamPregnant.RecordSource = Sql
    rsIamPregnant.Refresh
    rsIamPregnant.Recordset.MoveFirst
    SelectPregnancy = True
    Exit Function
    
errSelectPregnancy:
    SelectPregnancy = False
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
                For i = 0 To 3
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
        For i = 0 To 3
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("Frame1(0)") = Frame1(0).Caption
        .Fields("Frame1(2)") = Frame1(2).Caption
        .Fields("Frame1(3)") = Frame1(3).Caption
        .Fields("Frame1(4)") = Frame1(4).Caption
        .Fields("Frame1(5)") = Frame1(5).Caption
        .Fields("FormName") = "I Am Pregnant"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub ShowMidwife()
    On Error Resume Next
    cmbName.Clear
    With rsMidwife
        .MoveFirst
        Do While Not .EOF
            cmbName.AddItem .Fields("FirstName") & "  " & .Fields("LastName")
        .MoveNext
        Loop
    End With
End Sub

Private Sub cmbName_LostFocus()
    On Error Resume Next
    If boolNewRecord Then
        With rsIamPregnant.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("ConfirmationDate") = CDate(Format(Date1.Text, "dd.mm.yyyy"))
            .Update
            .Bookmark = .LastModified
            boolNewRecord = False
            cmbName.SetFocus
        End With
    End If
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

Private Sub Form_Activate()
    On Error Resume Next
    rsIamPregnant.Refresh
    ShowText
    ShowMidwife
    ShowAllButtons
    ShowKids
    SelectPregnancy
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsIamPregnant.DatabaseName = dbKidsTxt
    Set rsMidwife = dbKids.OpenRecordset("Midwife")
    Set rsLanguage = dbKidLang.OpenRecordset("frmIamPregnant")
    iWhichForm = 2
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmIamPregnant: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsIamPregnant.Recordset.Close
    rsMidwife.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmIamPregnant = Nothing
End Sub
