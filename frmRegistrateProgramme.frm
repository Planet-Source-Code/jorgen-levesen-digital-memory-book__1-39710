VERSION 5.00
Begin VB.Form frmRegistrateProgramme 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   120
      MaxLength       =   80
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "frmRegistrateProgramme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMyRecord As Recordset
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
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                If IsNull(.Fields("btnOK")) Then
                    .Fields("btnOK") = btnOk.Caption
                Else
                    btnOk.Caption = .Fields("btnOk")
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
        .Fields("btnExit") = btnExit.Caption
        .Fields("btnOK") = btnOk.Caption
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    With rsMyRecord
        .Edit
        .Fields("RegistrationID") = Trim(Text1.Text)
        .Update
    End With
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Dither Me
    ShowText
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmRegistrateProgramme")
    iWhichForm = 44
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRecord.Close
    rsLanguage.Close
    iWhichForm = 44
    Set frmRegistrateProgramme = Nothing
End Sub
