VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data rsLanguage 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterEmpW\MasterLang.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "frmCountry"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin RichTextLib.RichTextBox Text1 
      DataField       =   "Help"
      DataSource      =   "rsLanguage"
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10186
      _Version        =   393217
      BackColor       =   16777152
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmHelp.frx":030A
   End
   Begin VB.CommandButton btnExit 
      Height          =   615
      Left            =   6360
      Picture         =   "frmHelp.frx":0384
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   4455
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    'Dither Me
    Label1.ForeColor = WHITE
    rsLanguage.RecordSource = Trim(CStr(Label1.Caption))
    rsLanguage.Refresh
    With rsLanguage.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then Exit Do
        .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rsLanguage.DatabaseName = dbKidLangTxt
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Recordset.Close
    Set frmHelp = Nothing
End Sub

Private Sub Label1_Click()
    frmEditHelp.Caption = Label1.Caption
    frmEditHelp.Show
End Sub
