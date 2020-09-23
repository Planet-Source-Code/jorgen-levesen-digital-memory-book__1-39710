VERSION 5.00
Begin VB.Form frmRegistration 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programme Registration"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   285
      Index           =   8
      Left            =   840
      TabIndex        =   34
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton btnSendPost 
      Height          =   615
      Left            =   4680
      Picture         =   "frmRegistration.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Send Registration pr. Postal Service"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton btnSendFax 
      Height          =   615
      Left            =   5400
      Picture         =   "frmRegistration.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Send Fax Registration"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton btnSendEmail 
      Height          =   615
      Left            =   6120
      Picture         =   "frmRegistration.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Send Email registration"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Your Data"
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6615
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   9
         Left            =   5280
         TabIndex        =   35
         Top             =   3960
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   10
         Left            =   2880
         MaxLength       =   16
         TabIndex        =   10
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   7
         Left            =   4200
         TabIndex        =   33
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   6
         Left            =   6240
         TabIndex        =   32
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   5
         Left            =   5280
         TabIndex        =   31
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   30
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   29
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   2
         Left            =   4200
         TabIndex        =   28
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   6240
         TabIndex        =   26
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Check1"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   4320
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   9
         Left            =   2880
         TabIndex        =   9
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   8
         Left            =   2880
         TabIndex        =   8
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   7
         Left            =   2880
         TabIndex        =   7
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   6
         Left            =   2880
         TabIndex        =   6
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   5
         Left            =   2880
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   4
         Left            =   2880
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   3
         Left            =   2880
         TabIndex        =   3
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   2
         Left            =   2880
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   0
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Software Code:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   36
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Send Me regular E-mail information:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   25
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Programme Installation Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Email address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Country:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Town:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Full Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Your Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "= have to be filled "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   24
      Top             =   5280
      Width           =   2535
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMyRecord As Recordset
Dim rsSupplier As Recordset
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
                For n = 0 To 10
                    If IsNull(.Fields(n + 2)) Then
                        .Fields(n + 2) = Label1(n).Caption
                    Else
                        Label1(n).Caption = .Fields(n + 2)
                    End If
                Next
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("btnSendEmail")) Then
                    .Fields("btnSendEmail") = btnSendEmail.ToolTipText
                Else
                    btnSendEmail.ToolTipText = .Fields("btnSendEmail")
                End If
                If IsNull(.Fields("btnSendFax")) Then
                    .Fields("btnSendFax") = btnSendFax.ToolTipText
                Else
                    btnSendFax.ToolTipText = .Fields("btnSendFax")
                End If
                If IsNull(.Fields("btnSendPost")) Then
                    .Fields("btnSendPost") = btnSendPost.ToolTipText
                Else
                    btnSendPost.ToolTipText = .Fields("btnSendPost")
                End If
                'If IsNull(.Fields("btnInternet")) Then
                    '.Fields("btnInternet") = btnInternet.ToolTipText
                'Else
                    'btnInternet.ToolTipText = .Fields("btnInternet")
                'End If
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
        For n = 0 To 10
            .Fields(n + 2) = Label1(n).Caption
        Next
        .Fields("Frame1") = Frame1.Caption
        .Fields("btnSendEmail") = btnSendEmail.ToolTipText
        .Fields("btnSendFax") = btnSendFax.ToolTipText
        .Fields("btnSendPost") = btnSendPost.ToolTipText
        '.Fields("btnInternet") = btnInternet.ToolTipText
        .Fields("Msg1") = "Registration Scheme - Childrens Digital Memory Book"
        .Fields("Msg2") = "Unable to connect to the specified host"
        .Fields("Msg3") = "Please Fill In All Information"
        .Fields("Msg4") = "Unable to register, please try again later"
        .Fields("Msg5") = "Thank You!"
        .Fields("Help") = strMemo
        .Update
    End With
    DBEngine.Idle dbFreeLocks
End Sub

Private Sub LoadFields()
    On Error Resume Next
    With rsMyRecord
            If Not IsNull(.Fields("Company")) Then
                Text1(0).Text = .Fields("Company")
            End If
            If Not IsNull(.Fields("Adress1")) Then
                Text1(1).Text = .Fields("Adress1")
            End If
            If Not IsNull(.Fields("Adress2")) Then
                Text1(2).Text = .Fields("Adress2")
            End If
            If Not IsNull(.Fields("Adress3")) Then
                Text1(3).Text = .Fields("Adress3")
            End If
            If Not IsNull(.Fields("Zip")) Then
                Text1(4).Text = .Fields("Zip")
            End If
            If Not IsNull(.Fields("Town")) Then
                Text1(5).Text = .Fields("Town")
            End If
            If Not IsNull(.Fields("Country")) Then
                Text1(6).Text = .Fields("Country")
            End If
            If Not IsNull(.Fields("Telefon")) Then
                Text1(7).Text = .Fields("Telefon")
            End If
            If Not IsNull(.Fields("E-MailAdress")) Then
                Text1(8).Text = .Fields("E-MailAdress")
            End If
            If Not IsNull(.Fields("DateInstalled")) Then
                Text1(9).Text = Format(CDate(.Fields("DateInstalled")), "dd.mm.yyyy")
            Else
                Text1(9).Text = Format(Now, "dd.mm.yyyy")
            End If
            If Not IsNull(.Fields("RegistrationID")) Then
                Text1(10).Text = .Fields("RegistrationID")
            Else
                Text1(10).Text = .Fields("RegistrationID")
            End If
    End With
End Sub
Private Sub btnSendEmail_Click()
Dim Subject As String, Recipient As String, Message As String
    Subject = "MasterKid registration"
    Recipient = rsSupplier.Fields("SupplierEmailySysResponse")
    Message = rsLanguage.Fields("Msg1") & vbCrLf & vbCrLf
    Message = Message & Label1(0).Caption & vbTab & vbTab & Trim(Text1(0).Text) & vbCrLf
    Message = Message & Label1(1).Caption & vbTab & Trim(Text1(1).Text) & vbCrLf
    Message = Message & "                                 " & vbTab & Trim(Text1(2).Text) & vbCrLf
    Message = Message & "                                 " & vbTab & Trim(Text1(3).Text) & vbCrLf
    Message = Message & Label1(2).Caption & vbTab & vbTab & Trim(Text1(4).Text) & vbCrLf
    Message = Message & Label1(3).Caption & vbTab & vbTab & vbTab & Trim(Text1(5).Text) & vbCrLf
    Message = Message & Label1(4).Caption & vbTab & vbTab & vbTab & Trim(Text1(6).Text) & vbCrLf
    Message = Message & Label1(5).Caption & vbTab & Trim(Text1(7).Text) & vbCrLf
    Message = Message & Label1(6).Caption & vbTab & vbTab & Trim(Text1(8).Text) & vbCrLf
    If Len(Text1(9).Text) <> 0 Then
        Message = Message & Label1(7).Caption & vbTab & vbTab & Trim(Text1(9).Text) & vbCrLf & vbCrLf
    Else
        Message = Message & Label1(7).Caption & vbTab & vbTab & Format(Now, "dd.mm.yyyy") & vbCrLf & vbCrLf
    End If
    If Check1.Value = 1 Then
        Message = Message & Label1(9).Caption & vbTab & rsLanguage.Fields("Yes") & vbCrLf
    Else
        Message = Message & Label1(9).Caption & vbTab & rsLanguage.Fields("No") & vbCrLf
    End If
    Message = Message & Label1(10).Caption & vbTab & vbTab & Trim(Text1(10).Text) & vbCrLf
    Call SendOutlookMail(Subject, Recipient, Message)
    If bProgramNotAccesible Then
        End
    End If
End Sub

Private Sub btnSendFax_Click()
    On Error Resume Next
    Set wdApp = New Word.Application
    wdApp.Application.Visible = True
    wdApp.Application.WindowState = wdWindowStateMaximize
    With wdApp
        .Documents.Add DocumentType:=wdNewBlankDocument
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ActiveWindow.Selection.Font.Bold = True
        .ActiveWindow.Selection.Font.Shadow = True
        .ActiveWindow.Selection.Font.Size = 48
        .Selection.TypeText Text:=rsLanguage.Fields("FaxHeading")
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .ActiveWindow.Selection.Font.Bold = False
        .ActiveWindow.Selection.Font.Shadow = False
        .ActiveWindow.Selection.Font.Size = 12
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsLanguage.Fields("ToFaxNo") & vbTab
        .Selection.TypeText Text:=rsSupplier.Fields("SupplierFax")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsLanguage.Fields("ToCompany") & vbTab & vbTab
        .Selection.TypeText Text:=rsSupplier.Fields("SupplierName")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsLanguage.Fields("ToContact") & vbTab & vbTab & vbTab
        .Selection.TypeText Text:=rsSupplier.Fields("SupplierContact")
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ActiveWindow.Selection.Font.Bold = True
        .ActiveWindow.Selection.Font.Size = 16
        .ActiveWindow.Selection.Font.Shadow = True
        .Selection.TypeText Text:=rsLanguage.Fields("Msg1")
        .ActiveWindow.Selection.Font.Bold = False
        .Selection.TypeParagraph
        .ActiveWindow.Selection.Font.Size = 12
        .ActiveWindow.Selection.Font.Shadow = False
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.Tables.Add Range:=.Selection.Range, NumRows:=1, NumColumns:=2
        .Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=168.45, RulerStyle:=wdAdjustNone
        .Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=276.4, RulerStyle:=wdAdjustNone
    End With
    With wdApp.Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With wdApp
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Selection.TypeText Text:=Label1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(0).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(1).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=" "
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(2).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=" "
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(3).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(4).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(5).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(6).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(7).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(8).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Trim(Text1(9).Text)
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(9).Caption
        .Selection.MoveRight Unit:=wdCell
        If Check1.Value = 1 Then
            .Selection.TypeText Text:=rsLanguage.Fields("Yes")
        Else
            .Selection.TypeText Text:=rsLanguage.Fields("No")
        End If
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(10).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(10).Text
    End With
    If bProgramNotAccesible Then
        End
    End If
End Sub

Private Sub btnSendPost_Click()
    On Error Resume Next
    Set wdApp = New Word.Application
    wdApp.Application.Visible = True
    wdApp.Application.WindowState = wdWindowStateMaximize
    With wdApp
        .Documents.Add DocumentType:=wdNewBlankDocument
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .ActiveWindow.Selection.Font.Size = 12
        .Selection.TypeText Text:=rsSupplier.Fields("SupplierName")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsSupplier.Fields("SupplierAddr1") & vbTab & vbTab & vbTab & vbTab & Format(Now, "dd.mm.yyyy")
        If Not IsNull(rsSupplier.Fields("SupplierAddr2")) Then
            .Selection.TypeParagraph
            .Selection.TypeText Text:=rsSupplier.Fields("SupplierAddr2")
        End If
        If Not IsNull(rsSupplier.Fields("SupplierAddr3")) Then
            .Selection.TypeParagraph
            .Selection.TypeText Text:=rsSupplier.Fields("SupplierAddr3")
        End If
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsSupplier.Fields("SupplierZip") & "  " & rsSupplier.Fields("SupplierTown")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsSupplier.Fields("SupplierCountry")
        
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .ActiveWindow.Selection.Font.Bold = True
        .ActiveWindow.Selection.Font.Shadow = True
        .ActiveWindow.Selection.Font.Size = 18
        .Selection.TypeText Text:=rsLanguage.Fields("Msg1")
        .ActiveWindow.Selection.Font.Bold = False
        .ActiveWindow.Selection.Font.Shadow = False
        .ActiveWindow.Selection.Font.Size = 12
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.Tables.Add Range:=.Selection.Range, NumRows:=1, NumColumns:=2
        .Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=168.45, RulerStyle:=wdAdjustNone
        .Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=276.4, RulerStyle:=wdAdjustNone
    End With
    With wdApp.Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With wdApp
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        .Selection.TypeText Text:=Label1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=" "
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=" "
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(3).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(4).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(5).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(6).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(7).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(8).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(9).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(9).Caption
        .Selection.MoveRight Unit:=wdCell
        If Check1.Value = 1 Then
            .Selection.TypeText Text:=rsLanguage.Fields("Yes")
        Else
            .Selection.TypeText Text:=rsLanguage.Fields("No")
        End If
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(10).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(10).Text
    End With
    If bProgramNotAccesible Then
        End
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    LoadFields
    ReadText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsSupplier = dbKids.OpenRecordset("Supplier")
    Set rsLanguage = dbKidLang.OpenRecordset("frmRegistration")
    iWhichForm = 21
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRecord.Close
    rsSupplier.Close
    rsLanguage.Close
    iWhichForm = 0
    Set frmRegistration = Nothing
End Sub
