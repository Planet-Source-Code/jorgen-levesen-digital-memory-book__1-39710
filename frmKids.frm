VERSION 5.00
Begin VB.Form frmKids 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Children"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   4515
      Left            =   7440
      TabIndex        =   21
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.Data rsChildren 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Children"
         Top             =   120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ChildCallingName"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   0
         Left            =   2640
         MaxLength       =   60
         TabIndex        =   12
         Top             =   360
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ChildFirstName"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   1
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   11
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ChildLastName"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   2
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00400040&
         Height          =   1095
         Left            =   2640
         TabIndex        =   9
         Top             =   1440
         Width           =   2295
         Begin VB.CheckBox Check2 
            BackColor       =   &H00400040&
            Caption         =   "Boy ?"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   600
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00400040&
            Caption         =   "Girl ?"
            DataField       =   "ChildFemale"
            DataSource      =   "rsChildren"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BirthDate"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   3
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BirthTime"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   4
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LaburDuration"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   5
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   6
         Top             =   3360
         Width           =   975
      End
      Begin VB.ComboBox cmbTimeDim 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LaburDurationDim"
         DataSource      =   "rsChildren"
         Height          =   315
         Index           =   0
         Left            =   3720
         TabIndex        =   5
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BabyWeight"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   6
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   4
         Top             =   3840
         Width           =   975
      End
      Begin VB.ComboBox cmbTimeDim 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BabyWeightDim"
         DataSource      =   "rsChildren"
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   3
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BabyLength"
         DataSource      =   "rsChildren"
         Height          =   285
         Index           =   7
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   2
         Top             =   4200
         Width           =   975
      End
      Begin VB.ComboBox cmbTimeDim 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BabyLengthDim"
         DataSource      =   "rsChildren"
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   1
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Child calling name:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Time:"
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
         Index           =   4
         Left            =   480
         TabIndex        =   16
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Labour Duration:"
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
         Index           =   5
         Left            =   360
         TabIndex        =   15
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Baby weight at birth:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Baby length at birth:"
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
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   4200
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Stored Names"
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
      Index           =   8
      Left            =   7440
      TabIndex        =   22
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmKids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarkKid() As Variant
Dim rsLanguage As Recordset
Dim rsDimTime As Recordset
Dim rsDimLength As Recordset
Dim rsDimWeight As Recordset
Public Sub NewChild()
    rsChildren.Recordset.AddNew
    Text1(0).SetFocus
    boolNewRecord = True
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsChildren.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarkKid(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("ChildCallingName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            bookmarkKid(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub


Public Sub SelectChild()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM Children WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsChildren.RecordSource = Sql
    rsChildren.Refresh
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
                For i = 0 To 8
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("Check1")) Then
                    .Fields("Check1") = Check1.Caption
                Else
                    Check1.Caption = .Fields("Check1")
                End If
                If IsNull(.Fields("Check2")) Then
                    .Fields("Check2") = Check2.Caption
                Else
                    Check2.Caption = .Fields("Check2")
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
        For i = 0 To 8
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("Check1") = Check1.Caption
        .Fields("Check2") = Check2.Caption
        .Fields("FormName") = "Our Children"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub LoadcmbTimeDim()
    On Error Resume Next
    cmbTimeDim(0).Clear 'times
    With rsDimTime
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbTimeDim(0).AddItem .Fields("TimeDim")
            End If
        .MoveNext
        Loop
    End With
    
    cmbTimeDim(1).Clear 'weight
    With rsDimWeight
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbTimeDim(1).AddItem .Fields("WeightDim")
            End If
        .MoveNext
        Loop
    End With
    
    cmbTimeDim(2).Clear 'length
    With rsDimLength
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbTimeDim(2).AddItem .Fields("LengthDim")
            End If
        .MoveNext
        Loop
    End With
End Sub

Public Sub WriteKidsWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.TypeText Text:=Label1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        'boy or girl ?
        If Check1.Value = 1 Then
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("Check1")
        Else
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("Check2")
        End If
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(3).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(4).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(5).Text & "  " & cmbTimeDim(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(6).Text & "  " & cmbTimeDim(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(7).Text & "  " & cmbTimeDim(2).Text
    End With
End Sub

Public Sub WriteKids()
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
    sHeader = rsLanguage.Fields("FormName")
    Call PrintFront
    
    cPrint.CurrentX = LeftMargin
    cPrint.pPrint Label1(0).Caption, 1, True
    cPrint.pPrint Text1(0).Text, 3.5
    cPrint.pPrint Label1(1).Caption, 1, True
    cPrint.pPrint Text1(1).Text, 3.5
    cPrint.pPrint Label1(2).Caption, 1, True
    cPrint.pPrint Text1(2).Text, 3.5
    'boy or girl ?
    If Check1.Value = 1 Then
        cPrint.pPrint "", 1, True
        cPrint.pPrint rsLanguage.Fields("Check1"), 3.5
    Else
        cPrint.pPrint "", 1, True
        cPrint.pPrint rsLanguage.Fields("Check2"), 3.5
    End If
    cPrint.pPrint Label1(3).Caption, 1, True
    cPrint.pPrint Text1(3).Text, 3.5
    cPrint.pPrint Label1(4).Caption, 1, True
    cPrint.pPrint Text1(4).Text, 3.5
    cPrint.pPrint Label1(5).Caption, 1, True
    cPrint.pPrint Text1(5).Text & "  " & cmbTimeDim(0).Text, 3.5
    cPrint.pPrint Label1(6).Caption, 1, True
    cPrint.pPrint Text1(6).Text & "  " & cmbTimeDim(1).Text, 3.5
    cPrint.pPrint Label1(7).Caption, 1, True
    cPrint.pPrint Text1(7).Text & "  " & cmbTimeDim(2).Text, 3.5

    'we are done, release the printer object
    Screen.MousePointer = vbDefault
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    rsChildren.Refresh
    LoadList1
    List1.ListIndex = 0
    LoadcmbTimeDim
    'SelectChild
    ShowText
    ShowAllButtons
    ShowKids
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsChildren.DatabaseName = dbKidsTxt
    Set rsDimTime = dbKids.OpenRecordset("DimTime")
    Set rsDimLength = dbKids.OpenRecordset("DimLength")
    Set rsDimWeight = dbKids.OpenRecordset("DimWeight")
    Set rsLanguage = dbKidLang.OpenRecordset("frmKids")
    iWhichForm = 3
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmKids: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsChildren.Recordset.Close
    rsDimTime.Close
    rsDimLength.Close
    rsDimWeight.Close
    rsLanguage.Close
    iWhichForm = 0
    HideAllButtons
    HideKids
    Erase bookmarkKid
    Set frmKids = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    With rsChildren.Recordset
        .Bookmark = bookmarkKid(List1.ItemData(List1.ListIndex))
        If CBool(.Fields("ChildFemale")) = False Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
    End With
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    If boolNewRecord Then
        Select Case Index
        Case 0
            With rsChildren.Recordset
                .Fields("ChildCallingName") = Trim(Text1(0).Text)
                .Update
                LoadList1
                .Bookmark = .LastModified
                MDIMasterKid.LoadChildren
                boolNewRecord = False
            End With
        Case Else
        End Select
    End If
End Sub


