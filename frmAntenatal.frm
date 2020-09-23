VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAntenatal 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Antenatal Classes"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   6975
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   1815
      Begin VB.Data rsAntenatal 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   390
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "AntenatalClasses"
         Top             =   6180
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6465
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Date1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "AntenatalDate"
         DataSource      =   "rsAntenatal"
         Height          =   285
         Left            =   3960
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         DataField       =   "AntenatalNotes"
         DataSource      =   "rsAntenatal"
         Height          =   5715
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   10081
         _Version        =   393217
         BackColor       =   16777152
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmAntenatal.frx":0000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Antenatal Date:"
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
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Antenatal Note:"
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
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2865
      End
   End
End
Attribute VB_Name = "frmAntenatal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v1RecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewAntenatal()
    rsAntenatal.Recordset.AddNew
    Date1.SetFocus
    boolNewRecord = True
End Sub


Public Function SelectAntenatalChild() As Boolean
Dim Sql As String
    On Error GoTo errSelectAntenatalChild
    Sql = "SELECT * FROM AntenatalClasses WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    Sql = Sql & "ORDER BY AntenatalDate"
    rsAntenatal.RecordSource = Sql
    rsAntenatal.Refresh
    rsAntenatal.Recordset.MoveFirst
    SelectAntenatalChild = True
    Exit Function
    
errSelectAntenatalChild:
    SelectAntenatalChild = True
    Err.Clear
End Function

Public Sub PrintAntenatal()
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
    
    With rsAntenatal.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = glChildNo Then
                cPrint.FontBold = True
                cPrint.pPrint Label1(0).Caption & "  " & Format(CDate(.Fields("AntenatalDate")), "dd.mm.yyyy"), 1
                cPrint.FontBold = False
                cPrint.pMultiline RichTextBox1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
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
    ShowAllButtons
End Sub
Public Sub PrintAntenatalWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("Form"))
    With wdApp
        rsAntenatal.Recordset.MoveFirst
        Do While Not rsAntenatal.Recordset.EOF
            If CLng(rsAntenatal.Recordset.Fields("ChildNo")) = glChildNo Then
                DoEvents
                .Selection.TypeText Text:=Label1(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(rsAntenatal.Recordset.Fields("AntenatalDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(1).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText RichTextBox1.TextRTF, vbCFRTF
                .Selection.Paste
                DoEvents
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .ActiveWindow.Selection.Font.Name = "Times New Roman"
                .ActiveWindow.Selection.Font.Size = 10
                .Selection.MoveRight Unit:=wdCell
            End If
        rsAntenatal.Recordset.MoveNext
        Loop
    End With
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
        .Fields("sDate") = "Date: "
        .Fields("sPage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Public Sub FillList1()
    On Error Resume Next
    List1.Clear
    With rsAntenatal.Recordset
        .MoveLast
        .MoveFirst
        ReDim v1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("AntenatalDate")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            v1RecordBookmark(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
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

Private Sub Form_Activate()
    On Error Resume Next
    rsAntenatal.Refresh
    If SelectAntenatalChild Then
        FillList1
        List1.ListIndex = 0
    Else
        List1.Clear
    End If
    ShowText
    ShowAllButtons
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsAntenatal.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmAntenatal")
    iWhichForm = 6
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmAntenatal: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsAntenatal.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    iWhichForm = 0
    Set frmAntenatal = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsAntenatal.Recordset.Bookmark = v1RecordBookmark(List1.ItemData(List1.ListIndex))
End Sub

Private Sub RichTextBox1_GotFocus()
    On Error Resume Next
    If boolNewRecord Then
        With rsAntenatal.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("AntenatalDate") = Format(CDate(Date1.Text), "dd.mm.yyyy")
            .Update
            FillList1
            .Bookmark = .LastModified
        End With
    End If
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
    Call RichTextSelChange(frmAntenatal.RichTextBox1)
End Sub

