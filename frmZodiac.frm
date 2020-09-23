VERSION 5.00
Begin VB.Form frmZodiac 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Horoscope"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmZodiac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data rsHoroscope 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Horoscope"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      DataField       =   "HoroscopeShort"
      DataSource      =   "rsHoroscope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      DataField       =   "HoroscopeText"
      DataSource      =   "rsHoroscope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image Image1 
      DataField       =   "ZodiacSignPicture"
      DataSource      =   "rsHoroscope"
      Height          =   615
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmZodiac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Private Sub ShowText()
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
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .MoveFirst
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Form") = Me.Caption
        .Update
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ShowText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsHoroscope.DatabaseName = dbKidsTxt
    rsHoroscope.Refresh
    Set rsLanguage = dbKidLang.OpenRecordset("frmZodiac")
    iWhichForm = 39
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmZodiac: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsHoroscope.Recordset.Close
    rsLanguage.Close
    Set frmZodiac = Nothing
End Sub

