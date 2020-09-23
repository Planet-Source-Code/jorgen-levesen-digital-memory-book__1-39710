VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBaptismPictures 
   BackColor       =   &H00400040&
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BabtismPictureDate"
         DataSource      =   "rsBaptismPictures"
         Height          =   285
         Left            =   4800
         TabIndex        =   11
         Top             =   4920
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   5760
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   9
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Data rsBaptismPictures 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKidPic.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "BaptismPictures"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton btnDelete 
         Height          =   495
         Index           =   0
         Left            =   5565
         Picture         =   "frmBaptismPictures.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete picture"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton btnReadFromFile 
         Height          =   495
         Index           =   0
         Left            =   5565
         Picture         =   "frmBaptismPictures.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Read picture from file"
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton btnPastePicture 
         Height          =   495
         Index           =   0
         Left            =   5565
         Picture         =   "frmBaptismPictures.frx":080C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Paste picture"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton btnCopyPic 
         Height          =   495
         Index           =   0
         Left            =   5565
         Picture         =   "frmBaptismPictures.frx":0ECE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Copy picture to the Clipboard"
         Top             =   1320
         Width           =   495
      End
      Begin VB.CommandButton btnScan 
         Height          =   495
         Index           =   0
         Left            =   5565
         Picture         =   "frmBaptismPictures.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Scan a picture"
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "BabtismPictureCaption"
         DataSource      =   "rsBaptismPictures"
         Height          =   2175
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4920
         Width           =   4485
      End
      Begin MSComDlg.CommonDialog Cmd1 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   12
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Note:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "BabtismPicture"
         DataSource      =   "rsBaptismPictures"
         Height          =   4215
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   5220
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400040&
      Caption         =   "Picture No."
      ForeColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmBaptismPictures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFirstWrite As Boolean
Dim PicRecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewBapPic()
    rsBaptismPictures.Recordset.AddNew
    boolNewRecord = True
    Text1.SetFocus
End Sub
Public Sub DeleteRecord()
    On Error Resume Next
    rsBaptismPictures.Recordset.Delete
    If FillList2 Then
        List2.ListIndex = 0
    End If
End Sub
Public Sub PrintBaptismPic()
    On Error Resume Next
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If

    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    bFirstWrite = True
    sHeader = rsLanguage.Fields("FormName")
    cPrint.pStartDoc
    
    Call PrintFront
    
    With rsBaptismPictures.Recordset
        .MoveFirst
        Do While Not .EOF
            If Not bFirstWrite Then
                cPrint.pFooter
                cPrint.pNewPage
                Call PrintFront
            End If
            cPrint.pMultiline Text1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
            cPrint.pPrint
            cPrint.pPrint
            cPrint.pPrintPicture Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
            bFirstWrite = False
        .MoveNext
        Loop
    End With
    'end print
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub
Public Sub PrintBaptismPicWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.TypeText Text:=Format(rsBaptismPictures.Recordset.Fields("BabtismPictureCaption"))
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData Image1.Picture, vbCFBitmap
        .Selection.Paste
    End With
    Set wdApp = Nothing
End Sub
Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("label1(0)")) Then
                    .Fields("label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1(0)")
                End If
                If IsNull(.Fields("label1(1)")) Then
                    .Fields("label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("label1(1)")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
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
        .Fields("label1(0)") = Label1(0).Caption
        .Fields("label1(1)") = Label1(1).Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        .Fields("btnDelete") = btnDelete(0).ToolTipText
        .Fields("FormName") = "Baptism Pictures"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Public Function FillList2() As Boolean
Dim Sql As String
    On Error GoTo errFillList2
    Sql = "SELECT * FROM BaptismPictures WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsBaptismPictures.RecordSource = Sql
    rsBaptismPictures.Refresh
    rsBaptismPictures.Recordset.MoveFirst
    
    List2.Clear
    Set WClone = rsBaptismPictures.Recordset.Clone()
    With WClone
        .MoveLast
        .MoveFirst
        ReDim PicRecordBookmark(.RecordCount)
        Do While Not .EOF
            List2.AddItem .Fields("AutoField")
            List2.ItemData(List2.NewIndex) = List2.ListCount - 1
            PicRecordBookmark(List2.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    Set WClone = Nothing
    FillList2 = True
    Exit Function
        
errFillList2:
    FillList2 = False
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
    rsBaptismPictures.Refresh
    If FillList2 Then
        List2.ListIndex = 0
    End If
    Frame1.Caption = gsChildName
    ShowText
    ShowAllButtons
    ShowKids
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsBaptismPictures.DatabaseName = dbKidPicTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmBaptismPictures")
    iWhichForm = 28
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmBaptismPictures:  Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsBaptismPictures.Recordset.Close
    rsLanguage.Close
    iWhichForm = 0
    HideAllButtons
    HideKids
    Set frmBaptismPictures = Nothing
End Sub
Private Sub List2_Click()
        On Error Resume Next
        rsBaptismPictures.Recordset.Bookmark = PicRecordBookmark(List2.ItemData(List2.ListIndex))
End Sub

Private Sub Text1_LostFocus()
    On Error Resume Next
    If boolNewRecord Then
        With rsBaptismPictures.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("BabtismPictureCaption") = Trim(Text1.Text)
            .Update
            .Bookmark = .LastModified
            boolNewRecord = False
            FillList2
            Date1.SetFocus
        End With
    End If
End Sub
