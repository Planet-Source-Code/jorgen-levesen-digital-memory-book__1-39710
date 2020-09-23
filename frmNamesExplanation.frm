VERSION 5.00
Begin VB.Form frmNamesExplanation 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Name Explanation"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frmNamesExplanation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data rsNamesExplanation 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "NamesExplanation"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "NameExplanation"
      DataSource      =   "rsNamesExplanation"
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmNamesExplanation"
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
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM NamesExplanation WHERE Trim(FirstName) ="
    Sql = Sql & Chr(34) & Trim(Label1.Caption) & Chr(34)
    Sql = Sql & "AND Language ="
    Sql = Sql & Chr(34) & FileExt & Chr(34)
    rsNamesExplanation.RecordSource = Sql
    rsNamesExplanation.Refresh
    If rsNamesExplanation.Recordset.RecordCount = 0 Then
        'we did not have any explanation record for this name, make one
        rsNamesExplanation.Recordset.AddNew
        rsNamesExplanation.Recordset.Fields("Language") = FileExt
        rsNamesExplanation.Recordset.Fields("FirstName") = Trim(Label1.Caption)
        rsNamesExplanation.Recordset.Update
    End If
    ShowText
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rsNamesExplanation.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmNamesExplanation")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsNamesExplanation.UpdateRecord
    rsNamesExplanation.Recordset.Close
    rsLanguage.Close
    Set frmNamesExplanation = Nothing
End Sub
