VERSION 5.00
Object = "{608009F3-E1FB-11D2-9BA1-0040D0002C80}#1.0#0"; "nslock15vb6.ocx"
Begin VB.Form frmRegister 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please, register !"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRegisterMe 
      Caption         =   "&Send registration application"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin nslock15vb6.ActiveLock ActiveLock1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "jorgen"
      SoftwareName    =   "MasterKid"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program now used: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Liberation Key:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Software code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Users should register in order to use the software for a longer period."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This program is fully functional, but it will stop working after 21 days."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub ReadText()
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
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("label1")
                End If
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2.Caption
                Else
                    Label2.Caption = .Fields("label2")
                End If
                If IsNull(.Fields("label3")) Then
                    .Fields("label3") = Label3.Caption
                Else
                    Label3.Caption = .Fields("label3")
                End If
                If IsNull(.Fields("label4")) Then
                    .Fields("label4") = Label4.Caption
                Else
                    Label4.Caption = .Fields("label4")
                End If
                If IsNull(.Fields("label5")) Then
                    .Fields("label5") = Label5.Caption
                Else
                    Label5.Caption = .Fields("label5")
                End If
                If IsNull(.Fields("btnRegisterMe")) Then
                    .Fields("btnRegisterMe") = btnRegisterMe.Caption
                Else
                    btnRegisterMe.Caption = .Fields("btnRegisterMe")
                End If
                If IsNull(.Fields("cmdRegister")) Then
                    .Fields("cmdRegister") = cmdRegister.Caption
                Else
                    cmdRegister.Caption = .Fields("cmdRegister")
                End If
                If IsNull(.Fields("cmdCancel")) Then
                    .Fields("cmdCancel") = cmdCancel.Caption
                Else
                    cmdCancel.Caption = .Fields("cmdCancel")
                End If
                .Update
                DBEngine.Idle dbFreeLocks
                Me.MousePointer = Default
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Form") = Me.Caption
        .Fields("label1") = Label1.Caption
        .Fields("label2") = Label2.Caption
        .Fields("label3") = Label3.Caption
        .Fields("label4") = Label4.Caption
        .Fields("btnRegisterMe") = btnRegisterMe.Caption
        .Fields("cmdRegister") = cmdRegister.Caption
        .Fields("cmdCancel") = cmdCancel.Caption
        .Fields("Days") = "Days"
        .Fields("Msg1") = "Invalid liberation key !"
        .Fields("Msg2") = "Thank you for registering !"
        .Fields("Msg3") = "Your evaluation period has expired."
        .Update
    End With
    DBEngine.Idle dbFreeLocks
End Sub

Private Sub btnRegisterMe_Click()
    Unload Me
    Call MDIMasterKid.CloseActiveForm
    frmRegistration.Show
End Sub

Private Sub cmdCancel_Click()
    If bProgramNotAccesible Then
        MsgBox rsLanguage.Fields("Msg3")
        End
    End If
    Unload Me
End Sub

Private Sub cmdRegister_Click()
  ' Set the LiberationKey:
  ActiveLock1.LiberationKey = Text2
  
  ' Check if it was correct:
  If Not (ActiveLock1.RegisteredUser) Then
    MsgBox rsLanguage.Fields("Msg1")
  Else
    MsgBox rsLanguage.Fields("Msg2")
    With rsUser
        .Edit
        .Fields("RegistrationID") = ActiveLock1.SoftwareCode
        .Fields("RegistrationKey") = Text2
        .Fields("DateInstalled") = CDate(Format(Now, "dd.mm.yyyy"))
        .Update
    End With
    Unload Me
  End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Label5.Caption = Label5.Caption & "  " & ActiveLock1.UsedDays & "  " & rsLanguage.Fields("Days")
    If (ActiveLock1.UsedDays > 21) Then
        bProgramNotAccesible = True
    Else
        bProgramNotAccesible = False
    End If
End Sub

Private Sub Form_Load()
  On Error Resume Next
    Set rsUser = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmRegister")
    Text1 = ActiveLock1.SoftwareCode
    Text2 = ""
    ReadText
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bProgramNotAccesible Then
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsUser.Close
    rsLanguage.Close
    Set frmRegister = Nothing
End Sub
