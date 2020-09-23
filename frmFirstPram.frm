VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFirstPram 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My First Pram"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   2880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data rsFirstPram 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "FirstPram"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PurchaseDate"
         DataSource      =   "rsFirstPram"
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   14
         Top             =   5760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PurchaseWhere"
         DataSource      =   "rsFirstPram"
         Height          =   525
         Index           =   2
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CommandButton btnDelete 
         Height          =   495
         Index           =   0
         Left            =   2040
         Picture         =   "frmFirstPram.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Delete picture"
         Top             =   5640
         Width           =   495
      End
      Begin VB.CommandButton btnReadFromFile 
         Height          =   495
         Index           =   0
         Left            =   2040
         Picture         =   "frmFirstPram.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Read picture from file"
         Top             =   4200
         Width           =   495
      End
      Begin VB.CommandButton btnPastePicture 
         Height          =   495
         Index           =   0
         Left            =   2040
         Picture         =   "frmFirstPram.frx":080C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Paste picture"
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton btnCopyPic 
         Height          =   495
         Index           =   0
         Left            =   2040
         Picture         =   "frmFirstPram.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Copy picture to the Clipboard"
         Top             =   4680
         Width           =   495
      End
      Begin VB.CommandButton btnScan 
         Height          =   495
         Index           =   0
         Left            =   2040
         Picture         =   "frmFirstPram.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Scan a picture"
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Description"
         DataSource      =   "rsFirstPram"
         Height          =   1845
         Index           =   1
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PurchasePrice"
         DataSource      =   "rsFirstPram"
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   1
         Text            =   "0"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   2385
         Left            =   120
         Picture         =   "frmFirstPram.frx":16DA
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1875
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Where Purchased:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Picture"
         DataSource      =   "rsFirstPram"
         Height          =   2175
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Description:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Purchase Price:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Purchase Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmFirstPram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Public Sub NewPram()
    On Error Resume Next
    boolNewRecord = True
    rsFirstPram.Recordset.Move 0
    rsFirstPram.Recordset.AddNew
    Date1.SetFocus
End Sub
Public Sub WriteFirstPramWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.TypeText Text:=Label1(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(Date1.Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label1(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData Image1.Picture, vbCFBitmap
        .Selection.Paste
    End With
    Set wdApp = Nothing
End Sub

Public Sub WriteFirstPram()
    On Error Resume Next
    If Len(MDIMasterKid.cmbChildren.Text) = 0 Then Exit Sub
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    sHeader = rsLanguage.Fields("FormName")
    
    cPrint.pStartDoc
    Call PrintFront
    
    cPrint.pPrint Label1(0).Caption, 1, True
    If IsDate(Date1.Text) Then
        cPrint.pPrint Format(CDate(Date1.Text), "dd.mm.yyyy"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(1).Caption, 1, True
    If Len(Text1(0).Text) <> 0 Then
        cPrint.pPrint Text1(0).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label1(3).Caption, 1, True
    If Len(Text1(2).Text) <> 0 Then
        cPrint.pMultiline Text1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint "", 3.5
    End If
    cPrint.pPrint Label1(2).Caption, 1, True
    If Len(Text1(1).Text) <> 0 Then
        cPrint.pMultiline Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    If Not IsNull(rsFirstPram.Recordset.Fields("Picture")) Then
        cPrint.pPrintPicture Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
    End If
    
    cPrint.pFooter
    Screen.MousePointer = vbDefault
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub

Public Function SelectPram() As Boolean
Dim Sql As String
    On Error GoTo errSelectPram
    Sql = "SELECT * FROM FirstPram WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsFirstPram.RecordSource = Sql
    rsFirstPram.Refresh
    SelectPram = True
    Exit Function
    
errSelectPram:
    SelectPram = False
    Err.Clear
End Function

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
                For i = 0 To 3
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("btnPastePicture")) Then
                    .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
                Else
                    btnPastePicture(0).ToolTipText = .Fields("btnPastePicture")
                End If
                If IsNull(.Fields("btnReadFromFile")) Then
                    .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
                Else
                    btnReadFromFile(0).ToolTipText = .Fields("btnReadFromFile")
                End If
                If IsNull(.Fields("btnCopyPic")) Then
                    .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
                Else
                    btnCopyPic(0).ToolTipText = .Fields("btnCopyPic")
                End If
                If IsNull(.Fields("btnScan")) Then
                    .Fields("btnScan") = btnScan(0).ToolTipText
                Else
                    btnScan(0).ToolTipText = .Fields("btnScan")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete(0).ToolTipText
                Else
                    btnDelete(0).ToolTipText = .Fields("btnDelete")
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
        For i = 0 To 3
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        .Fields("btnDelete") = btnDelete(0).ToolTipText
        .Fields("FormName") = Me.Caption
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Public Function ShowPram() As Boolean
Dim Sql As String
    On Error GoTo errShowPram
    Sql = "SELECT * FROM FirstPram WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsFirstPram.RecordSource = Sql
    rsFirstPram.Refresh
    ShowPram = True
    Exit Function
    
errShowPram:
    ShowPram = False
    Err.Clear
End Function

Private Sub btnCopyPic_Click(Index As Integer)
    On Error Resume Next
    Clipboard.SetData Image1.Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click(Index As Integer)
    On Error Resume Next
    Set Image1.Picture = LoadPicture()
End Sub

Private Sub btnPastePicture_Click(Index As Integer)
        On Error Resume Next
        Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub btnReadFromFile_Click(Index As Integer)
        On Error Resume Next
        With Cmd1
            .filename = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "Pictures (*.bmp; *.pcx;*.jpg;*.jpeg;*.gif)|*.bmp;*.pcx;*.jpg;*.jpeg;*.gif"
            .FilterIndex = 1
            .Action = 1
        End With
        Set Image1.Picture = LoadPicture(Cmd1.filename)
End Sub

Private Sub btnScan_Click(Index As Integer)
    Dim ret As Long, t As Single
    On Error Resume Next
    ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Date1_Click()
Dim UserDate As Date
    If IsDate(Date1.Text) Then
        UserDate = CVDate(Date1.Text)
    Else
        UserDate = Format(Now, "dd.mm.yyyy")
    End If
    If frmCalendar.GetDate(UserDate) Then
        Date1.Text = UserDate
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsFirstPram.Refresh
    If SelectPram Then
    End If
    Frame1.Caption = gsChildName
    ShowText
    ShowAllButtons
    ShowKids
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsFirstPram.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmFirstPram")
    iWhichForm = 32
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmFirstPram: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsFirstPram.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmFirstPram = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Select Case Index
    Case 0
        If boolNewRecord Then
            With rsFirstPram.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("PurchaseDate") = CDate(Format(Date1.Text, "dd.mm.yyyy"))
                .Update
                boolNewRecord = False
                .Bookmark = .LastModified
            End With
        End If
    Case Else
    End Select
End Sub


