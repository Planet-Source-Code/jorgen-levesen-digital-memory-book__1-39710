VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmNames 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Names"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmNames.frx":0000
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frmNames.frx":0016
      TabIndex        =   43
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Data rsNames 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Names"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton btnSeeName 
      Caption         =   "&See name with surname(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   42
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Height          =   4095
      Left            =   3840
      TabIndex        =   11
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "Å"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   28
         Left            =   600
         TabIndex        =   40
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "Ø"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   39
         Top             =   3480
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "Æ"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   26
         Left            =   1080
         TabIndex        =   38
         Top             =   3120
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "Z"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   25
         Left            =   600
         TabIndex        =   37
         Top             =   3120
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "Y"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "X"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   23
         Left            =   1080
         TabIndex        =   35
         Top             =   2760
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "W"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   34
         Top             =   2760
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "V"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   33
         Top             =   2760
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "U"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   20
         Left            =   1080
         TabIndex        =   32
         Top             =   2400
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "T"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   19
         Left            =   600
         TabIndex        =   31
         Top             =   2400
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "S"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "R"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   29
         Top             =   2040
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "Q"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   28
         Top             =   2040
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "P"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "O"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   1080
         TabIndex        =   26
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "N"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   600
         TabIndex        =   25
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "M"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "L"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   1080
         TabIndex        =   23
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "K"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   600
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "J"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "I"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   20
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "H"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   19
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "G"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "F"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "E"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "D"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "C"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "B"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400040&
         Caption         =   "A"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton Option1 
         BackColor       =   &H00400040&
         Caption         =   "Girls Names"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00400040&
         Caption         =   "Boys Names"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   41
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400040&
      Caption         =   "Father/Mother:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400040&
      Caption         =   "Mother/Father:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400040&
      Caption         =   "ll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   8
      Top             =   7320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400040&
      Caption         =   "ll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400040&
      Caption         =   "ll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400040&
      Caption         =   "ll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400040&
      Caption         =   "Father:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400040&
      Caption         =   "Mother:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql As String
Dim rsLanguage As Recordset
Dim rsMyRecord As Recordset
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
                    Me.Caption = Me.Caption & " - " & rsNames.Recordset.RecordCount & " " & .Fields("Names")
                End If
                If IsNull(.Fields("Label2")) Then
                    .Fields("Label2") = Label2.Caption
                Else
                    Label2.Caption = .Fields("Label2")
                End If
                If IsNull(.Fields("Label3")) Then
                    .Fields("Label3") = Label3.Caption
                Else
                    Label3.Caption = .Fields("Label3")
                End If
                If IsNull(.Fields("Label5")) Then
                    .Fields("Label5") = Label5.Caption
                Else
                    Label5.Caption = .Fields("Label5")
                End If
                If IsNull(.Fields("Label6")) Then
                    .Fields("Label6") = Label6.Caption
                Else
                    Label6.Caption = .Fields("Label6")
                End If
                If IsNull(.Fields("Option1(0)")) Then
                    .Fields("Option1(0)") = Option1(0).Caption
                Else
                    Option1(0).Caption = .Fields("Option1(0)")
                End If
                If IsNull(.Fields("Option1(1)")) Then
                    .Fields("Option1(1)") = Option1(1).Caption
                Else
                    Option1(1).Caption = .Fields("Option1(1)")
                End If
                If IsNull(.Fields("Grid1Coln0")) Then
                    .Fields("Grid1Coln0") = Grid1.Columns(0).Caption
                Else
                    Grid1.Columns(0).Caption = .Fields("Grid1Coln0")
                End If
                If IsNull(.Fields("Grid1Coln1")) Then
                    .Fields("Grid1Coln1") = Grid1.Columns(1).Caption
                Else
                    Grid1.Columns(1).Caption = .Fields("Grid1Coln1")
                End If
                If IsNull(.Fields("btnSeeName")) Then
                    .Fields("btnSeeName") = btnSeeName.Caption
                Else
                    btnSeeName.Caption = .Fields("btnSeeName")
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
        .Fields("Label2") = Label2.Caption
        .Fields("Label3") = Label3.Caption
        .Fields("Label5") = Label5.Caption
        .Fields("Label6") = Label6.Caption
        .Fields("Option1(0)") = Option1(0).Caption
        .Fields("Option1(1)") = Option1(1).Caption
        .Fields("Grid1Coln0") = Grid1.Columns(0).Caption
        .Fields("Grid1Coln1") = Grid1.Columns(1).Caption
        .Fields("btnSeeName") = btnSeeName.Caption
        .Fields("Names") = "Names"
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub SelectBoy()
    On Error Resume Next
    Sql = "SELECT * FROM Names WHERE CBool(BoyName) = True ORDER BY FirstName"
    rsNames.RecordSource = Sql
    rsNames.Refresh
    Label1.Caption = rsNames.Recordset.RecordCount & "  " & "Records"
End Sub
Private Sub SelectGirl()
    On Error Resume Next
    Sql = "SELECT * FROM Names WHERE CBool(BoyName) = False ORDER BY FirstName"
    rsNames.RecordSource = Sql
    rsNames.Refresh
    Label1.Caption = rsNames.Recordset.RecordCount & "  " & "Records"
End Sub

Private Sub btnSeeName_Click()
    On Error Resume Next
    Label2.Visible = True
    Label3.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label4(0).Visible = True
    Label4(1).Visible = True
    Label4(2).Visible = True
    Label4(3).Visible = True
    Label4(2).Caption = Grid1.Columns(0).Text & "  " & rsMyRecord.Fields("MotherLastName")
    Label4(3).Caption = Grid1.Columns(0).Text & "  " & rsMyRecord.Fields("FatherLastName")
    If Not IsNull(rsMyRecord.Fields("MotherLastName")) Then
        Label4(0).Caption = Grid1.Columns(0).Text & "  " & rsMyRecord.Fields("MotherLastName")
        Label4(3).Caption = Label4(3).Caption & "  " & rsMyRecord.Fields("MotherLastName")
    Else
        Label4(0).Caption = Grid1.Columns(0).Text & "  " & "?"
    End If
    If Not IsNull(rsMyRecord.Fields("FatherLastName")) Then
        Label4(1).Caption = Grid1.Columns(0).Text & "  " & rsMyRecord.Fields("FatherLastName")
        Label4(2).Caption = Label4(2).Caption & "  " & rsMyRecord.Fields("FatherLastName")
    Else
        Label4(1).Caption = Grid1.Columns(0).Text & "  " & "?"
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsNames.Refresh
    ShowText
    SelectBoy
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsNames.DatabaseName = dbKidsTxt
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmNames")
    iWhichForm = 8
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmNames: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsNames.Recordset.Close
    rsMyRecord.Close
    rsLanguage.Close
    iWhichForm = 0
    HideAllButtons
    Set frmNames = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
    Case 0  'boys names
        SelectBoy
    Case 1  'girls names
        SelectGirl
    Case Else
    End Select
End Sub


Private Sub Option2_Click(Index As Integer)
Dim sLetter As String
    On Error Resume Next
    Select Case Index
    Case 0
        sLetter = "A"
    Case 1
        sLetter = "B"
    Case 2
        sLetter = "C"
    Case 3
        sLetter = "D"
    Case 4
        sLetter = "E"
    Case 5
        sLetter = "F"
    Case 6
        sLetter = "G"
    Case 7
        sLetter = "H"
    Case 8
        sLetter = "I"
    Case 9
        sLetter = "J"
    Case 10
        sLetter = "K"
    Case 11
        sLetter = "L"
    Case 12
        sLetter = "M"
    Case 13
        sLetter = "N"
    Case 14
        sLetter = "O"
    Case 15
        sLetter = "P"
    Case 16
        sLetter = "Q"
    Case 17
        sLetter = "R"
    Case 18
        sLetter = "S"
    Case 19
        sLetter = "T"
    Case 20
        sLetter = "U"
    Case 21
        sLetter = "V"
    Case 22
        sLetter = "W"
    Case 23
        sLetter = "X"
    Case 24
        sLetter = "Y"
    Case 25
        sLetter = "Z"
    Case 26
        sLetter = "Æ"
    Case 27
        sLetter = "Ø"
    Case 28
        sLetter = "Å"
    Case Else
    End Select
    If Option1(0).Value = True Then
        Sql = "SELECT * FROM Names WHERE CBool(BoyName) = True AND Left(FirstName, 1) ="
    Else
        Sql = "SELECT * FROM Names WHERE CBool(BoyName) = False  AND Left(FirstName, 1) ="
    End If
    Sql = Sql & Chr(34) & sLetter & Chr(34)
    Sql = Sql & " ORDER BY FirstName"
    rsNames.RecordSource = Sql
    rsNames.Refresh
    Label1.Caption = rsNames.Recordset.RecordCount & "  " & "Records"
End Sub


