VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFathersNotesBaptism 
   BackColor       =   &H00400040&
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8340
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Fathers Baptism notes for "
      ForeColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.TextBox NoteDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "NoteDate"
         DataSource      =   "rsBaptismNotes"
         Height          =   285
         Left            =   4920
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.Data rsBaptismNotes 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "FathersNotesBaptism"
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin RichTextLib.RichTextBox RichText1 
         DataField       =   "FathersNote"
         DataSource      =   "rsBaptismNotes"
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   9340
         _Version        =   393217
         BackColor       =   16777152
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmFathersNotesBaptism.frx":0000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400040&
         Caption         =   "Note Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   480
         Width           =   2685
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5880
      Left            =   6720
      TabIndex        =   2
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      Caption         =   "Previous notes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   3
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "frmFathersNotesBaptism"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v2RecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewBaptismNotesFather()
    On Error Resume Next
    rsBaptismNotes.Recordset.Move 0
    rsBaptismNotes.Recordset.AddNew
    NoteDate.SetFocus
    boolNewRecord = True
End Sub


Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
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
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
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
        .Fields("Label1(0)") = Label1(0).Caption
        .Fields("Label1(1)") = Label1(1).Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("FormName") = "Fathers Notes, Babtism"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Public Sub FillList2()
    On Error Resume Next
    List2.Clear
    With rsBaptismNotes.Recordset
        .MoveLast
        .MoveFirst
        ReDim v2RecordBookmark(.RecordCount)
        Do While Not .EOF
            List2.AddItem .Fields("NoteDate")
            List2.ItemData(List2.NewIndex) = List2.ListCount - 1
            v2RecordBookmark(List2.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Function SelectRecords() As Boolean
Dim Sql As String
    On Error GoTo errSelectRecords
    Sql = "SELECT * FROM FathersNotesBaptism WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    Sql = Sql & "ORDER BY NoteDate"
    rsBaptismNotes.RecordSource = Sql
    rsBaptismNotes.Refresh
    SelectRecords = True
    Exit Function
    
errSelectRecords:
    SelectRecords = False
    Err.Clear
End Function
Public Sub WriteBaptismNotes()
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
    
    With rsBaptismNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(.Fields("NoteDate")) Then
                    cPrint.FontBold = True
                    cPrint.pPrint Label1(0).Caption & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1
                    cPrint.FontBold = False
                    If Len(RichText1.Text) <> 0 Then
                        cPrint.pMultiline RichText1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        Call PrintFront
                    End If
                End If
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

Public Sub WriteBaptismNotesWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        rsBaptismNotes.Recordset.MoveFirst
        Do While Not rsBaptismNotes.Recordset.EOF
            If CLng(rsBaptismNotes.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(rsBaptismNotes.Recordset.Fields("NoteDate")) Then
                    DoEvents
                    .Selection.TypeText Text:=Label1(1).Caption
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=Format(CDate(rsBaptismNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    Clipboard.Clear
                    Clipboard.SetText RichText1.TextRTF, vbCFRTF
                    .Selection.Paste
                    DoEvents
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        rsBaptismNotes.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsBaptismNotes.Refresh
    SelectRecords
    FillList2
    List2.ListIndex = 0
    ShowText
    Frame1.Caption = rsLanguage.Fields("Frame1") & "  " & gsChildName
    ShowAllButtons
    ShowKids
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsBaptismNotes.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesBaptism")
    iWhichForm = 12
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmFathersNotesBaptism: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsBaptismNotes.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Erase v2RecordBookmark
    Set frmFathersNotesBaptism = Nothing
End Sub
Private Sub List2_Click()
    With rsBaptismNotes.Recordset
        .Bookmark = v2RecordBookmark(List2.ItemData(List2.ListIndex))
        NoteDate.Text = CDate(.Fields("NoteDate"))
        RichText1.TextRTF = .Fields("FathersNote")
    End With
End Sub

Private Sub NoteDate_Click()
Dim UserDate As Date
    If IsDate(NoteDate.Text) Then
        UserDate = CVDate(NoteDate.Text)
    Else
        UserDate = Format(Now, "dd.mm.yyyy")
    End If
    If frmCalendar.GetDate(UserDate) Then
        NoteDate.Text = UserDate
    End If
End Sub

Private Sub RichText1_GotFocus()
    On Error Resume Next
    If boolNewRecord Then
        With rsBaptismNotes.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("NoteDate") = Format(NoteDate.Text, "dd.mm.yyyy")
            .Update
            FillList2
            boolNewRecord = False
            .Bookmark = .LastModified
            RichText1.SetFocus
        End With
    End If
End Sub

Private Sub RichText1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    On Error Resume Next
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichText1.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichText1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichText1_SelChange()
    Call RichTextSelChange(frmFathersNotesBaptism.RichText1)
End Sub
