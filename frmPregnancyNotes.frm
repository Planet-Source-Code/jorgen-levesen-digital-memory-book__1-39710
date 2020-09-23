VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPregnancyNotes 
   BackColor       =   &H000080FF&
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Pregnancy Notes for "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6975
      Begin VB.TextBox NoteDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "NoteDate"
         DataSource      =   "rsPregnancyNotes"
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
         Left            =   5520
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.Data rsPregnancyNotes 
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
         RecordSource    =   "PregnancyNotes"
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin RichTextLib.RichTextBox RichText1 
         DataField       =   "PregnancyNote"
         DataSource      =   "rsPregnancyNotes"
         Height          =   5295
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   9340
         _Version        =   393217
         BackColor       =   16777152
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmPregnancyNotes.frx":0000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Note Date:"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   2925
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   6075
      Left            =   7200
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Previous notes"
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
      Left            =   7260
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "frmPregnancyNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v2RecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewPregnancyNotes()
    On Error Resume Next
    rsPregnancyNotes.Recordset.Move 0
    rsPregnancyNotes.Recordset.AddNew
    NoteDate.SetFocus
    boolNewRecord = True
End Sub
Public Sub WritePregnancyNotesWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        rsPregnancyNotes.Recordset.MoveFirst
        Do While Not rsPregnancyNotes.Recordset.EOF
            DoEvents
            .Selection.TypeText Text:=Label1(1).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(CDate(rsPregnancyNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            Clipboard.Clear
            Clipboard.SetText RichText1.TextRTF, vbCFRTF
            .Selection.Paste
            DoEvents
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
        rsPregnancyNotes.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub

Public Sub WritePregnancyNotes()
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
    sHeader = rsLanguage.Fields("FormName")
    
    cPrint.pStartDoc
    Call PrintFront
    
    With rsPregnancyNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            cPrint.FontBold = True
            cPrint.pPrint Label1(1).Caption & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1
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

Public Function SelectNotes() As Boolean
Dim Sql As String
    On Error GoTo errSelectNotes
    Sql = "SELECT * FROM PregnancyNotes WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsPregnancyNotes.RecordSource = Sql
    rsPregnancyNotes.Refresh
    SelectNotes = True
    Exit Function
    
errSelectNotes:
    SelectNotes = False
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
        .Fields("Form") = Me.Caption
        .Fields("Label1(0)") = Label1(0).Caption
        .Fields("Label1(1)") = Label1(1).Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("FormName") = "Pregnancy Notes"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Public Sub FillList2()
    On Error Resume Next
    List2.Clear
    If rsPregnancyNotes.Recordset.RecordCount <> 0 Then
        With rsPregnancyNotes.Recordset
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
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsPregnancyNotes.Refresh
    ShowText
    Frame1.Caption = rsLanguage.Fields("Frame1") & " " & gsChildName
    If SelectNotes Then
        FillList2
        List2.ListIndex = 0
    End If
    ShowAllButtons
    ShowKids
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsPregnancyNotes.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmPregnancyNotes")
    iWhichForm = 1
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmPregnancyNotes: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsPregnancyNotes.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Erase v2RecordBookmark
    Set frmPregnancyNotes = Nothing
End Sub
Private Sub List2_Click()
    On Error Resume Next
    rsPregnancyNotes.Recordset.Bookmark = v2RecordBookmark(List2.ItemData(List2.ListIndex))
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
Private Sub RichText1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichText1.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichText1_LostFocus()
    If boolNewRecord Then
        With rsPregnancyNotes.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("NoteDate") = CDate(Format(NoteDate.Text, "dd.mm.yyyy"))
            .Update
            FillList2
            .Bookmark = .LastModified
            boolNewRecord = False
        End With
    End If
End Sub

Private Sub RichText1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichText1_SelChange()
    Call RichTextSelChange(frmPregnancyNotes.RichText1)
End Sub
