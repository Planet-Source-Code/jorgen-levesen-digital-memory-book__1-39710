VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   ControlBox      =   0   'False
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00800000&
      Height          =   1575
      Left            =   120
      TabIndex        =   41
      Top             =   4560
      Width           =   4575
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Data rsMyRec 
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
         RecordSource    =   "MyRecord"
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LanguageScreen"
         DataSource      =   "rsMyRec"
         Height          =   315
         Index           =   18
         Left            =   3240
         TabIndex        =   43
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "LanguagePrint"
         DataSource      =   "rsMyRec"
         Height          =   315
         Index           =   19
         Left            =   3240
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Language on Print:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   45
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Language on Screen:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   44
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00800000&
      Height          =   4335
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4575
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "MotherLastName"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   16
         Left            =   2760
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FatherFirstName"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Fax"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   30
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Telefon"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   29
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Country"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   28
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Town"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   27
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Zip"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   26
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Adress3"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   25
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Adress2"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   24
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Adress1"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   23
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "MotherFirstName"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FatherLastName"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   17
         Left            =   2760
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Fathers Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Fax No.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Telephone No.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Country:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Town:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Zip Code:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Mothers Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00800000&
      Height          =   1455
      Left            =   4800
      TabIndex        =   15
      Top             =   4680
      Width           =   3255
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "RegistrationID"
         DataSource      =   "rsMyRec"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "HtmlDirectory"
         DataSource      =   "rsMyRec"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Registration ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   15
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Directory HTML Help"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   16
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   3015
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00800000&
      Height          =   1335
      Left            =   4800
      TabIndex        =   8
      Top             =   3240
      Width           =   3255
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Check2"
         DataField       =   "PrintUsingWord"
         DataSource      =   "rsMyRec"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         DataField       =   "ShowFirst"
         DataSource      =   "rsMyRec"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         DataField       =   "WishToRecieveInfo"
         DataSource      =   "rsMyRec"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Use MS Word for print:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Show Splash Screen:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Send E-mail information:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   17
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Height          =   1455
      Left            =   4800
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Internet"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "EMail"
         DataSource      =   "rsMyRec"
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Internet:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "E-Mail:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   1455
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TermDueDate"
         DataSource      =   "rsMyRec"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   840
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Birth due date"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolFirst As Boolean
Dim rsLanguage As Recordset
Private Sub ReadText()
Dim strMemo As String
    'find YOUR rsLanguage text
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
                For n = 0 To 17
                    If IsNull(.Fields(n + 2)) Then
                        .Fields(n + 2) = Label1(n).Caption
                    Else
                        Label1(n).Caption = .Fields(n + 2)
                    End If
                Next
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
        .Fields("Form") = Me.Caption
        For n = 0 To 17
            .Fields(n + 2) = Label1(n).Caption
        Next
        .Fields("Msg1") = "MS Word is not installed on your computer !"
        .Fields("Help") = strMemo
        .Update
    End With
    DBEngine.Idle dbFreeLocks
End Sub

Private Sub Check2_Click()
    On Error Resume Next
    If Check2.Value = 1 Then
        If IsAppPresent("Word.Document\CurVer", "") = False Then
            MsgBox rsLanguage.Fields("Msg1")
            Check2.Value = 0
            MDIMasterKid.mnuPrintWord.Checked = False
        Else
            MDIMasterKid.mnuPrintWord.Checked = True
        End If
    End If
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    Dither Me
    If boolFirst = False Then Exit Sub
    rsMyRec.Refresh
    Call ReadText
    boolFirst = False
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsMyRec.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmUser")
    boolFirst = True
    iWhichForm = 7
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmUser: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRec.UpdateRecord
    rsMyRec.Recordset.Close
    rsLanguage.Close
    iWhichForm = 0
    Set frmUser = Nothing
End Sub
Private Sub Text1_Click(Index As Integer)
    Select Case Index
    Case 9
        With CommonDialog1
            .filename = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "HTML-file (*.html)|*.html"
            .FilterIndex = 1
            .Action = 1
            Text1(9).Text = .filename
        End With
    Case 13 'path to front picture
        With CommonDialog1
            .filename = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "Pictures (*.bmp; *.pcx;*.jpg;*.jpeg;*.gif)|*.bmp;*.pcx;*.jpg;*.jpeg;*.gif"
            .FilterIndex = 1
            .Action = 1
            Text1(13).Text = .filename
        End With
    Case Else
    End Select
End Sub
