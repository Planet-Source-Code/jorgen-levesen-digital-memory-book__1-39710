VERSION 5.00
Begin VB.Form frmFirstTimeUse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "First time user - Program Hint"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Height          =   375
      Left            =   6120
      Picture         =   "frmFirstTimeUse.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   6120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Do not show this screen anymore"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   5415
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmFirstTimeUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Dim rsMyRec As Recordset
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
                If IsNull(.Fields("Label1")) Then
                    .Fields("Label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("Label1")
                End If
                If IsNull(.Fields("Option1")) Then
                    .Fields("Option1") = Option1.Caption
                Else
                    Option1.Caption = .Fields("Option1")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.ToolTipText
                Else
                    btnExit.ToolTipText = .Fields("btnExit")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
            
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Form") = Me.Caption
        .Fields("Label1") = Label1.Caption
        .Fields("Option1") = Option1.Caption
        .Fields("btnExit") = btnExit.ToolTipText
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsMyRec = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmFirstTimeUse")
    ShowText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    With rsMyRec
        If Option1.Value = True Then
            .Edit
            .Fields("ShowFirstScreen") = False
            .Update
        End If
        .Close
    End With
    rsLanguage.Close
    Set frmFirstTimeUse = Nothing
End Sub
