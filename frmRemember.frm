VERSION 5.00
Begin VB.Form frmRemember 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remember to ...."
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.Data rsToRemember 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ToRember"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5880
      Left            =   5160
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   6135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "WhenToRember"
         DataSource      =   "rsToRemember"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Check1"
         DataField       =   "SetReminderAlarm"
         DataSource      =   "rsToRemember"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   5640
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ToRemember"
         DataSource      =   "rsToRemember"
         Height          =   3525
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Purpose"
         DataSource      =   "rsToRemember"
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Set Alarm ?"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   7
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Remember when ?"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   6
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Remember what ?"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Purpose:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmRemember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v1RecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewRemember()
    rsToRemember.Recordset.AddNew
    Text1(0).SetFocus
    boolNewRecord = True
End Sub

Private Sub SelectRemember()
Dim Sql As String
    Sql = "SELECT * FROM ToRember WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsToRemember.RecordSource = Sql
    rsToRemember.Refresh
End Sub

Public Sub WriteRememberWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("Form"))
    With wdApp
        rsToRemember.Recordset.MoveFirst
        Do While Not rsToRemember.Recordset.EOF
            If CLng(rsToRemember.Recordset.Fields("ChildNo")) = glChildNo Then
                .Selection.TypeText Text:=Label1(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Text1(0).Text
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(1).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Text1(1).Text
                .Selection.MoveRight Unit:=wdCell
                If IsDate(rsToRemember.Recordset.Fields("WhenToRember")) Then
                    .Selection.TypeText Text:=Label1(2).Caption
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=Format(CDate(rsToRemember.Recordset.Fields("WhenToRember")), "dd.mm.yyyy")
                Else
                    .Selection.TypeText Text:=Label1(2).Caption
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=""
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        rsToRemember.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub

Public Sub WriteRemember()
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
    sHeader = Me.Caption
    
    cPrint.pStartDoc
    Call PrintFront
    
    With rsToRemember.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = glChildNo Then
                cPrint.pPrint Label1(0).Caption, 1, True
                cPrint.pPrint Text1(0).Text, 3.5
                cPrint.pPrint Label1(1).Caption, 1, True
                cPrint.pMultiline Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                If IsDate(.Fields("WhenToRember")) Then
                    cPrint.pPrint Label1(2).Caption, 1, True
                    cPrint.pPrint Format(CDate(.Fields("WhenToRember")), "dd.mm.yyyy"), 3.5
                End If
                cPrint.pPrint
                cPrint.pPrint
            End If
        .MoveNext
        Loop
    End With
    
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
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
                If IsNull(.Fields("Label1(0)")) Then
                    .Fields("Label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("Label1(0)")
                End If
                If IsNull(.Fields("Label1(1)")) Then
                    .Fields("Label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("Label1(1)")
                End If
                If IsNull(.Fields("Label1(2)")) Then
                    .Fields("Label1(2)") = Label1(2).Caption
                Else
                    Label1(2).Caption = .Fields("Label1(2)")
                End If
                If IsNull(.Fields("Label1(3)")) Then
                    .Fields("Label1(3)") = Label1(3).Caption
                Else
                    Label1(3).Caption = .Fields("Label1(3)")
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
        .Fields("Form") = Me.Caption
        .Fields("Label1(0)") = Label1(0).Caption
        .Fields("Label1(1)") = Label1(1).Caption
        .Fields("Label1(2)") = Label1(2).Caption
        .Fields("Label1(3)") = Label1(3).Caption
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Public Sub FillList1()
    On Error Resume Next
    
    SelectRemember
    
    List1.Clear
    With rsToRemember.Recordset
        .MoveLast
        .MoveFirst
        ReDim v1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("Purpose")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            v1RecordBookmark(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub Check1_LostFocus()
    On Error Resume Next
        If Check1.Value = 1 Then
        frmAlarm.DateAlarm.Text = CDate(Date1.Text)
        frmAlarm.Text1.Text = Trim(Text1(1).Text)
        frmAlarm.Show 1
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
    rsToRemember.Refresh
    FillList1
    List1.ListIndex = 0
    ShowText
    ShowAllButtons
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsToRemember.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmRemember")
    iWhichForm = 26
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmRemember: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsToRemember.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    iWhichForm = 0
    Set frmRemember = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsToRemember.Recordset.Bookmark = v1RecordBookmark(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If boolNewRecord Then
        Select Case Index
        Case 0
            With rsToRemember.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("Purpose") = Trim(Text1(0).Text)
                .Update
                FillList1
                .Bookmark = .LastModified
            End With
        Case Else
        End Select
        boolNewRecord = False
    End If
End Sub


