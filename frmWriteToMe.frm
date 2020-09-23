VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWriteToMe 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail to Programme Developer"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnSend 
      Caption         =   "&Send this Mail"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   5640
      Width           =   8055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   4575
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8070
         _Version        =   393217
         BackColor       =   16777152
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmWriteToMe.frx":0000
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "I would like to have the folowing errors corrected / new facilities added:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmWriteToMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupplier As Recordset
Dim rsLanguage As Recordset
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
                If IsNull(.Fields("Label1")) Then
                    .Fields("Label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("Label1")
                End If
                If IsNull(.Fields("btnSend")) Then
                    .Fields("btnSend") = btnSend.Caption
                Else
                    btnSend.Caption = .Fields("btnSend")
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
        .Fields("Label1") = Label1.Caption
        .Fields("btnStore") = btnSend.Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub btnSend_Click()
    On Error Resume Next
    IsMicrosoftMailRunning
    If Len(RichTextBox1.Text) <> 0 Then
        Call SendOutlookMail("Error/New Message", rsSupplier.Fields("SupplierEmailySysResponse"), RichTextBox1.Text)
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ShowText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    Set rsSupplier = dbKids.OpenRecordset("Supplier")
    Set rsLanguage = dbKidLang.OpenRecordset("frmWriteToMe")
    iWhichForm = 20
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmWriteToMe: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSupplier.Close
    rsLanguage.Close
    iWhichForm = 0
    Set frmWriteToMe = Nothing
End Sub
Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
   End Sub

Private Sub RichTextBox1_SelChange()
    Call RichTextSelChange(frmWriteToMe.RichTextBox1)
End Sub
