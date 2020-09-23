VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateDatabase 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Update / Back-up"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   Icon            =   "frmUpdateDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   120
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   0
      Top             =   1800
   End
   Begin VB.Image imgOrg 
      Height          =   480
      Left            =   4680
      Picture         =   "frmUpdateDatabase.frx":030A
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   240
      Stretch         =   -1  'True
      Top             =   840
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   240
      Stretch         =   -1  'True
      Top             =   480
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MasterKidPic.mdb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MasterKidLang.mdb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MasterKid.mdb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update finished !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Update finished !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "frmUpdateDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function GetFileSize(strFile As String) As String
Dim fso As New Scripting.FileSystemObject
Dim f As File
Dim lngBytes As Long
Const KB As Long = 1024
Const MB As Long = 1024 * KB
Const GB As Long = 1024 * MB
    Set f = fso.GetFile(fso.GetFile(strFile))
    lngBytes = f.Size
    If lngBytes < KB Then
        GetFileSize = Format(lngBytes) & " bytes"
    ElseIf lngBytes < MB Then
        GetFileSize = Format(lngBytes / KB, "0.00") & " KB"
    ElseIf lngBytes < GB Then
        GetFileSize = Format(lngBytes / MB, "0.00") & " MB"
    Else
        GetFileSize = Format(lngBytes / GB, "0.00") & " GB"
    End If
End Function

Private Sub UpdateDatabase()
Dim strOldName(3) As String, strNewName(3) As String, tmpName As String
Dim iValue As Integer, i As Integer, strLogFile As String
Dim sSizeOld As String, sSizeNew As String
    
    On Error GoTo errUpdate
    tmpName = App.Path & "\temp.mdb"
    strLogFile = App.Path & "\MasterKidUpdate.txt"
    Open strLogFile For Output As #1
    Write #1, "Update/Compact from the: ", Format(Now, "dd.mm.yyyy")
    iValue = 0
    Timer2.Enabled = False
    
    strOldName(0) = App.Path & "\MasterKid.bck"
    strOldName(1) = App.Path & "\MasterKidLang.bck"
    strOldName(2) = App.Path & "\MasterKidPic.bck"
    
    strNewName(0) = App.Path & "\MasterKid.mdb"
    strNewName(1) = App.Path & "\MasterKidLang.mdb"
    strNewName(2) = App.Path & "\MasterKidPic.mdb"
    
    ProgressBar1.Value = iValue
    DBEngine.Idle
    
        For i = 0 To 3
            On Error Resume Next
            Label2(i).ForeColor = &HFF&
            Kill tmpName
            sSizeOld = GetFileSize(strNewName(i))
            DBEngine.CompactDatabase strNewName(i), tmpName
            Kill strOldName(i)
            Name strNewName(i) As strOldName(i)
            Name tmpName As strNewName(i)
            sSizeNew = GetFileSize(strNewName(i))
            Image1(i).Picture = imgOrg.Picture
            iValue = iValue + 20
            ProgressBar1.Value = iValue
            'write the log-file
            Write #1, strNewName(i), " - Compact OK"
            Write #1, "Size Before: ", sSizeOld, "  -  Size After: ", sSizeNew
            DoEvents
        Next
    
    Label1(0).Visible = True
    Label1(1).Visible = True
    Close #1
    DoEvents
    Timer1.Enabled = True
    Exit Sub
    
errUpdate:
    Beep
    MsgBox Err.Description, vbCritical, "Update Database"
    Write #1, Err.Description, strNewName(i)
    Close #1
    Err.Clear
End Sub
Private Sub Form_Activate()
    DoEvents
    Timer2.Enabled = True
End Sub

Private Sub Form_Load()
    Me.Show 1
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set frmUpdateDatabase = Nothing
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub

Private Sub Timer2_Timer()
    UpdateDatabase
End Sub
