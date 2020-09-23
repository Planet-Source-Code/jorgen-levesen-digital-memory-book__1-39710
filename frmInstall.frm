VERSION 5.00
Begin VB.Form frmInstall 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Master Install"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frmInstall.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   2640
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   120
      Picture         =   "frmInstall.frx":030A
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copying HTML-files"
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   10
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copying Templates"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   120
      Picture         =   "frmInstall.frx":0614
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   120
      Picture         =   "frmInstall.frx":091E
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Up MasterKid"
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmInstall.frx":0C28
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Making Directories"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Where is your CD-Rom ?"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub ShellAbout Lib "Shell.DLL" (ByVal hwnd As Integer, ByVal AppName As String, ByVal Copyright As String, ByVal hIcon As Integer)
Dim boolError As Boolean
Dim i As Integer
Dim strAppPath As String
Dim strNextPath As String
Dim strTmpPath As String
Dim RetVal As Long
Private Sub InstallHtml()
    On Error GoTo errInstallHtml
    File1.Visible = True
    strTmpPath = Left(Drive1.Drive, 2) & "\Html"
    strNextPath = strAppPath & "\Html"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\_fpclass"
    strNextPath = strAppPath & "\Html\_fpclass"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\_private"
    strNextPath = strAppPath & "\Html\_private"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\_themes"
    strNextPath = strAppPath & "\Html\_themes"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\_themes\_vti_cnf"
    strNextPath = strAppPath & "\Html\_themes\_vti_cnf"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\_themes\artsy"
    strNextPath = strAppPath & "\Html\_themes\artsy"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\_vti_cnf"
    strNextPath = strAppPath & "\Html\_vti_cnf"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\_vti_pvt"
    strNextPath = strAppPath & "\Html\_vti_pvt"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\images"
    strNextPath = strAppPath & "\Html\images"
    WriteFiles
    strTmpPath = Left(Drive1.Drive, 2) & "\Html\images\_vti_cnf"
    strNextPath = strAppPath & "\Html\images\_vti_cnf"
    WriteFiles
    File1.Visible = False
    Exit Sub
    
errInstallHtml:
    Beep
    MsgBox Err.Description, vbCritical, "Install HTML helpfiles"
    Resume errInstallHtml2:
errInstallHtml2:
    boolError = True
End Sub


Private Sub MakeDir()
    On Error Resume Next
    
    strAppPath = "c:\MasterKid"
    MkDir strAppPath
    
    'make the directories and sub-directories for the html help-file
    strNextPath = strAppPath & "\HTML"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\_fpclass"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\_private"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\_themes"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\_themes\_vti_cnf"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\_themes\artsy"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\_vti_cnf"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\_vti_pvt"
    MkDir strNextPath
    strNextPath = strAppPath & "\HTML\images"
    MkDir strNextPath
    strNextPath = strAppPath & "HTML\images\_vti_cnf"
    MkDir strNextPath
    
    'make the directory for the sound directory
    strNextPath = strAppPath & "\Sound"
    MkDir strNextPath
    
    'make the directory for the video directory
    strNextPath = strAppPath & "\Video"
    MkDir strNextPath
    
    'make the directory for the MS Word template directory
    strNextPath = strAppPath & "\Template"
    MkDir strNextPath
End Sub
Private Sub RunShell(cmdline)
    Dim hProcess As Long
    Dim ProcessId As Long
    Dim exitCode As Long
    ProcessId& = Shell(cmdline, vbNormalFocus)
    hProcess& = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId&)
    Do
        Call GetExitCodeProcess(hProcess&, exitCode&)
        DoEvents
    Loop While exitCode& > 0
End Sub

Private Sub InstallMasterKid()
    On Error GoTo errInstallMasterKid
    strTmpPath = Left(Drive1.Drive, 2) & "\SetUp\Setup.exe"
    RunShell strTmpPath
    Exit Sub
    
errInstallMasterKid:
    Beep
    MsgBox Err.Description, vbCritical, "Install MasterKid"
    Resume errInstallMasterKid2
errInstallMasterKid2:
    boolError = True
End Sub
Private Sub InstallTempPlates()
Dim strTemp1 As String
Dim strTemp2 As String
    On Error GoTo errInstallTempPlates
    File1.Visible = True
    
    strTmpPath = Left(Drive1.Drive, 2) & "\Template\"
    strNextPath = strAppPath & "\Template\"
    ChDir strTmpPath
    File1.Path = strTmpPath
    For i = 0 To File1.ListCount - 1
        strTemp1 = strTmpPath & File1.List(i)
        strTemp2 = strNextPath & File1.List(i)
        FileCopy strTemp1, strTemp2
    Next
    
    File1.Visible = False
    Exit Sub
    
errInstallTempPlates:
    Beep
    MsgBox Err.Description, vbCritical, "Install Tempplates"
    Resume errInstallTempPlates2:
errInstallTempPlates2:
    boolError = True
End Sub
Private Sub WriteFiles()
Dim strTemp1 As String
Dim strTemp2 As String
    ChDir strTmpPath
    File1.Path = strTmpPath
    If File1.ListCount = -1 Then Exit Sub
    For i = 0 To File1.ListCount - 1
        strTemp1 = strTmpPath & File1.List(i)
        strTemp2 = strNextPath & File1.List(i)
        FileCopy strTemp1, strTemp2
        DoEvents
    Next
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnStart_Click()
    strAppPath = "c:\MasterAll"
    boolError = False
    Call MakeDir
    Image1(0).Visible = True
    DoEvents
    
    Call InstallMasterKid
    
    If Not boolError Then
        Image1(1).Visible = True
        DoEvents
        Call InstallTempPlates
        Image1(2).Visible = True
        DoEvents
    Else
        Label3.Caption = "Error install MasterKid"
        DoEvents
    End If
    
    Call InstallHtml
    
    If Not boolError Then
        Image1(3).Visible = True
        DoEvents
        Beep
        Label3.Caption = "Installation Finished !"
        btnExit.SetFocus
    Else
        Label3.Caption = "Error install HTML-files"
        DoEvents
        btnExit.SetFocus
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set frmInstall = Nothing
End Sub
