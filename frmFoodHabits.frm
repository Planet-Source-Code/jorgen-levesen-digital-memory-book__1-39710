VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFoodHabits 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Food Habits"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6495
      Begin VB.Data rsFoodHabits 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "FoodHabits"
         Top             =   0
         Visible         =   0   'False
         Width           =   1140
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         DataField       =   "BestMeal"
         DataSource      =   "rsFoodHabits"
         Height          =   2175
         Left            =   1440
         TabIndex        =   4
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3836
         _Version        =   393217
         BackColor       =   16777152
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmFoodHabits.frx":0000
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "EatVegetablesMonths"
         DataSource      =   "rsFoodHabits"
         Height          =   285
         Index           =   3
         Left            =   3720
         TabIndex        =   3
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "EatPorridgeMonths"
         DataSource      =   "rsFoodHabits"
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   2
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "DrankOfCupMonths"
         DataSource      =   "rsFoodHabits"
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   1
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Breast-feedMonths"
         DataSource      =   "rsFoodHabits"
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   120
         Picture         =   "frmFoodHabits.frx":00D5
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Months old"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Months old"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Months old"
         Height          =   255
         Index           =   1
         Left            =   4440
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000C0C0&
         Caption         =   "Months old"
         Height          =   255
         Index           =   0
         Left            =   4440
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000C0C0&
         Caption         =   "My best meal was:"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   10
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "Eat Wegtables when:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "Eat Porrige when:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "Drank of cup when:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         Caption         =   "Breast feed until:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmFoodHabits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Public Sub NewFoodHabit()
    On Error Resume Next
    rsFoodHabits.Recordset.Move 0
    rsFoodHabits.Recordset.AddNew
    Text1(0).BackColor = &HC0E0FF
    Text1(0).SetFocus
    boolNewRecord = True
End Sub

Public Sub WriteFoodHabits()
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
    
    For i = 0 To 3
        cPrint.pPrint rsLanguage.Fields(i + 2), 1, True
        If Len(Text1(i).Text) <> 0 Then
            cPrint.pPrint Text1(i).Text & "  " & Label2(i).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
    Next
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint
    cPrint.pPrint Label1(4).Caption, 1, True
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
    Call Form_Activate
End Sub

Public Sub WriteFoodHabitsWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        For i = 0 To 3
            .Selection.TypeText Text:=Label1(i).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Text1(i).Text & " " & Label2(i).Caption
            .Selection.MoveRight Unit:=wdCell
        Next
        .Selection.TypeText Text:=Label1(4).Caption
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
    End With
    Set wdApp = Nothing
End Sub

Public Function ReadFoodHabits() As Boolean
Dim Sql As String
    On Error GoTo errReadFoodHabits
    Sql = "SELECT * FROM FoodHabits WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsFoodHabits.RecordSource = Sql
    rsFoodHabits.Refresh
    rsFoodHabits.Recordset.MoveFirst
    ReadFoodHabits = True
    Frame1.Caption = MDIMasterKid.cmbChildren.Text
    Exit Function
    
errReadFoodHabits:
    ReadFoodHabits = False
    Err.Clear
    Frame1.Caption = " "
End Function
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
                    End If
                Next
                If IsNull(.Fields("Label2")) Then
                    .Fields("Label2") = Label2(0).Caption
                Else
                    Label2(0).Caption = .Fields("Label2")
                    Label2(1).Caption = .Fields("Label2")
                    Label2(2).Caption = .Fields("Label2")
                    Label2(3).Caption = .Fields("Label2")
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
            
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Form") = Me.Caption
        For i = 0 To 4
            .Fields("Label1(i)") = Label1(i + 2).Caption
        Next
        .Fields("Label2") = Label2(0).Caption
        .Fields("FormName") = Me.Caption
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsFoodHabits.Refresh
    ReadFoodHabits
    ShowText
    ShowAllButtons
    ShowKids
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsFoodHabits.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmFoodHabits")
    iWhichForm = 25
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmFoodHabits:  Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsFoodHabits.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmFoodHabits = Nothing
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
    Call RichTextSelChange(frmFoodHabits.RichTextBox1)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If boolNewRecord Then
        Select Case Index
        Case 0
            With rsFoodHabits.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("Breast-feedMonths") = Trim(Text1(0).Text)
                .Update
                .Bookmark = .LastModified
            End With
            boolNewRecord = False
            Text1(0).BackColor = &HFFFFC0
        Case Else
        End Select
    End If
End Sub
