VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBirthDates 
   BackColor       =   &H00400040&
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   9570
   Begin VB.Frame Frame3 
      Caption         =   "Former Birthdays"
      Height          =   7695
      Left            =   7680
      TabIndex        =   15
      Top             =   0
      Width           =   1695
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   7050
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   1485
      End
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   4194368
      TabCaption(0)   =   "Birthday"
      TabPicture(0)   =   "frmBirthDates.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Pictures"
      TabPicture(1)   =   "frmBirthDates.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(3)=   "Text2"
      Tab(1).Control(4)=   "rsBirthDayPictures"
      Tab(1).Control(5)=   "btnPastePicture"
      Tab(1).Control(6)=   "btnReadFromFile"
      Tab(1).Control(7)=   "btnCopyPic"
      Tab(1).Control(8)=   "btnDelete"
      Tab(1).Control(9)=   "btnScan"
      Tab(1).Control(10)=   "Picture2"
      Tab(1).Control(11)=   "List2"
      Tab(1).Control(12)=   "Cmd1"
      Tab(1).ControlCount=   13
      Begin MSComDlg.CommonDialog Cmd1 
         Left            =   -68280
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame1 
         Caption         =   "Birthday"
         Height          =   1575
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   6855
         Begin VB.TextBox Date1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BirthDayDate"
            DataSource      =   "rsBirthDays"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2760
            TabIndex        =   0
            Top             =   720
            Width           =   1335
         End
         Begin VB.Image Image2 
            Height          =   855
            Index           =   1
            Left            =   240
            Picture         =   "frmBirthDates.frx":0038
            Stretch         =   -1  'True
            Top             =   480
            Width           =   780
         End
         Begin VB.Image Image2 
            Height          =   855
            Index           =   0
            Left            =   5880
            Picture         =   "frmBirthDates.frx":0B71
            Stretch         =   -1  'True
            Top             =   480
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4935
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   6855
         Begin VB.Data rsBirthDays 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "BirthDays"
            Top             =   120
            Visible         =   0   'False
            Width           =   1140
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            DataField       =   "BirthDayNote"
            DataSource      =   "rsBirthDays"
            Height          =   4575
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   6570
            _ExtentX        =   11589
            _ExtentY        =   8070
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmBirthDates.frx":16AA
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6660
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   885
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -68235
         ScaleHeight     =   345
         ScaleWidth      =   300
         TabIndex        =   9
         Top             =   4080
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton btnScan 
         Height          =   495
         Left            =   -68280
         Picture         =   "frmBirthDates.frx":1724
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Scan a picture from Scanner"
         Top             =   2520
         Width           =   420
      End
      Begin VB.CommandButton btnDelete 
         Height          =   495
         Left            =   -68280
         Picture         =   "frmBirthDates.frx":186E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   420
      End
      Begin VB.CommandButton btnCopyPic 
         Height          =   495
         Left            =   -68280
         Picture         =   "frmBirthDates.frx":19B8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Copy picture to the Clipboard"
         Top             =   1920
         Width           =   420
      End
      Begin VB.CommandButton btnReadFromFile 
         Height          =   495
         Left            =   -68280
         Picture         =   "frmBirthDates.frx":207A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   420
      End
      Begin VB.CommandButton btnPastePicture 
         Height          =   495
         Left            =   -68280
         Picture         =   "frmBirthDates.frx":273C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   420
      End
      Begin VB.Data rsBirthDayPictures 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKidPic.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "BirthDayPictures"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "PictureCaption"
         DataSource      =   "rsBirthDayPictures"
         Height          =   1575
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   5760
         Width           =   5790
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Picture No."
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   1005
      End
      Begin VB.Image Picture1 
         DataField       =   "Picture"
         DataSource      =   "rsBirthDayPictures"
         Height          =   4695
         Left            =   -73680
         Stretch         =   -1  'True
         Top             =   720
         Width           =   5220
      End
      Begin VB.Label Label3 
         Caption         =   "Picture Note:"
         Height          =   255
         Left            =   -73680
         TabIndex        =   13
         Top             =   5520
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmBirthDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFirstWrite As Boolean
Dim DaysRecordBookmark() As Variant
Dim PicRecordBookmark() As Variant
Dim rsLanguage As Recordset
Public Sub NewBirthDates()
    Select Case Tab1.Tab
    Case 0
        rsBirthDays.Recordset.AddNew
        Date1.SetFocus
    Case 1  'pictures
        rsBirthDayPictures.Recordset.AddNew
        Text2.SetFocus
    Case Else
    End Select
    boolNewRecord = True
End Sub

Public Sub DeleteBirthDay()
    On Error Resume Next
    rsBirthDayPictures.Recordset.Delete
    FillList2
End Sub

Public Sub FillList2()
    'On Error Resume Next
    List2.Clear
    If rsBirthDayPictures.Recordset.RecordCount <> 0 Then
        Set WClone = rsBirthDayPictures.Recordset.Clone()
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
    End If
End Sub
Public Sub FillList3()
    On Error Resume Next
    List3.Clear
    If rsBirthDays.Recordset.RecordCount <> 0 Then
        With rsBirthDays.Recordset
            .MoveLast
            .MoveFirst
            ReDim DaysRecordBookmark(.RecordCount)
            Do While Not .EOF
                List3.AddItem .Fields("BirthDayDate")
                List3.ItemData(List3.NewIndex) = List3.ListCount - 1
                DaysRecordBookmark(List3.ListCount - 1) = .Bookmark
            .MoveNext
            Loop
        End With
    End If
End Sub

Public Sub NewChildBirthDates()
    On Error Resume Next
    Select Case frmBirthDates.Tab1.Tab
    Case 0
        If SelectDays Then
            FillList3
            List3.ListIndex = 0
        End If
    Case 1
        If SelectDays Then
            FillList3
            List3.ListIndex = 0
            If SelectPictures Then
                FillList2
                List2.ListIndex = 0
            End If
        End If
    Case Else
    End Select
End Sub

Public Sub PrintBirthdays()
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
    bFirstWrite = True
    cPrint.pStartDoc
    Call PrintFront
    
    With rsBirthDays.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("ChildNo") = glChildNo Then
                If Not bFirstWrite Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.FontBold = True
                cPrint.pPrint Frame1.Caption, 1, True
                If IsDate(.Fields("BirthDayDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("BirthDayDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.FontBold = False
                cPrint.pPrint
                cPrint.pPrint "Note:", 1, True
                If Len(RichTextBox1.Text) <> 0 Then
                    cPrint.pMultiline RichTextBox1.Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                
                'now print all the pictures from this birthday
                rsBirthDayPictures.Recordset.MoveFirst
                Do While Not rsBirthDayPictures.Recordset.EOF
                    If rsBirthDayPictures.Recordset.Fields("ChildNo") = glChildNo And CDate(.rsBirthDayPictures.Recordset.Fields("BirthDayDate")) = CDate(.rsBirthDays.Recordset.Fields("Date")) Then
                        If Not bFirstWrite Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                        If Len(Text2.Text) <> 0 Then
                            cPrint.pMultiline Text2.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint " ", 1
                        End If
                        If Not IsNull(rsBirthDayPictures.Recordset.Fields("Picture")) Then
                            cPrint.pPrintPicture Picture1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                            bFirstWrite = False
                        Else
                            bFirstWrite = True
                        End If
                        cPrint.pPrint
                    End If
                rsBirthDayPictures.Recordset.MoveNext
                Loop
                bFirstWrite = False
            End If
        .MoveNext
        Loop
    End With
    
    'end of print
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    
    Set cPrint = Nothing
End Sub

Public Sub PrintBirthdaysWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.TypeText Text:=Frame1.Caption
        .Selection.MoveRight Unit:=wdCell
        If IsDate(rsBirthDays.Recordset.Fields("BirthDayDate")) Then
            .Selection.TypeText Text:=Format(CDate(rsBirthDays.Recordset.Fields("BirthDayDate")), "dd.mm.yyyy")
        Else
            .Selection.TypeText Text:=" "
        End If
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:="Note:"
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        .Selection.MoveRight Unit:=wdCell
    
        'now print all the pictures from this birthday
        rsBirthDayPictures.Recordset.MoveFirst
        Do While Not rsBirthDayPictures.Recordset.EOF
            If rsBirthDayPictures.Recordset.Fields("ChildNo") = glChildNo Then
                If CDate(rsBirthDayPictures.Recordset.Fields("BirthDayDate")) = CDate(List3.List(List3.ListIndex)) Then
                    DoEvents
                    .Selection.TypeText Text:=Text2.Text
                    .Selection.MoveRight Unit:=wdCell
                    Picture2.Picture = Picture1.Picture
                    Clipboard.Clear
                    Clipboard.SetData Picture2.Picture, vbCFBitmap
                    .Selection.Paste
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        rsBirthDayPictures.Recordset.MoveNext
        Loop
    End With
    'end of print
    Set wdApp = Nothing
End Sub
Public Function SelectDays() As Boolean
Dim Sql As String
    On Error GoTo errSelectDays
    Sql = "SELECT * FROM BirthDays WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    Sql = Sql & "ORDER BY BirthDayDate"
    rsBirthDays.RecordSource = Sql
    rsBirthDays.Refresh
    SelectDays = True
    Exit Function
    
errSelectDays:
    Beep
    SelectDays = False
End Function
Public Function SelectPictures() As Boolean
Dim Sql As String
    On Error GoTo errSelectPictures
    If IsDate(Date1.Text) Then
        Sql = "SELECT * FROM BirthDayPictures WHERE CLng(ChildNo) ="
        Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
        Sql = Sql & "AND CDate(Date) ="
        Sql = Sql & Chr(34) & CDate(Date1.Text) & Chr(34)
        Sql = Sql & "ORDER BY Date"
    Else
        Beep
        MsgBox "First select a Birthdate !!"
        SelectPictures = False
        Exit Function
    End If
    rsBirthDayPictures.RecordSource = Sql
    rsBirthDayPictures.Refresh
    SelectPictures = True
    Exit Function
    
errSelectPictures:
    Beep
    SelectPictures = False
End Function
Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("label1")
                End If
                If IsNull(.Fields("label3")) Then
                    .Fields("label3") = Label3.Caption
                Else
                    Label3.Caption = .Fields("label3")
                End If
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 0
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
        .Fields("label1") = Label1.Caption
        .Fields("label3") = Label3.Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("Frame3") = Frame3.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("Msg1") = "Select a Child first !"
        .Fields("Msg2") = "You have to select a Birthday first !"
        .Fields("FormName") = "Birth days"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub btnCopyPic_Click()
    Clipboard.SetData Picture1.Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click()
        On Error Resume Next
    Set Picture1.Picture = LoadPicture()
End Sub

Private Sub btnPastePicture_Click()
        On Error Resume Next
        Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub btnReadFromFile_Click()
        On Error Resume Next
        With Cmd1
            .filename = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "Pictures (*.bmp; *.pcx;*.jpg;*.jpeg;*.gif)|*.bmp;*.pcx;*.jpg;*.jpeg;*.gif"
            .FilterIndex = 1
            .Action = 1
        End With
        Set Picture1.Picture = LoadPicture(Cmd1.filename)
End Sub

Private Sub btnScan_Click()
    Dim ret As Long, t As Single
    On Error Resume Next
    ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Picture1.Picture = Clipboard.GetData(vbCFDIB)
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
    rsBirthDays.Refresh
    rsBirthDayPictures.Refresh
    SelectDays
    FillList3
    List3.ListIndex = 0
    ShowText
    ShowAllButtons
    ShowKids
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsBirthDays.DatabaseName = dbKidsTxt
    rsBirthDayPictures.DatabaseName = dbKidPicTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmBirthDates")
    iWhichForm = 17
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmBirthDates: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsBirthDays.Recordset.Close
    rsBirthDayPictures.Recordset.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    Erase DaysRecordBookmark
    Erase PicRecordBookmark
    iWhichForm = 0
    Set frmBirthDates = Nothing
End Sub
Private Sub List2_Click()
    On Error Resume Next
    rsBirthDayPictures.Recordset.Bookmark = PicRecordBookmark(List2.ItemData(List2.ListIndex))
End Sub

Private Sub List3_Click()
    On Error Resume Next
    rsBirthDays.Recordset.Bookmark = DaysRecordBookmark(List3.ItemData(List3.ListIndex))
    If Tab1.Tab = 1 Then
        SelectPictures
        FillList2
        List2.ListIndex = 0
    End If
End Sub

Private Sub RichTextBox1_GotFocus()
    On Error Resume Next
    With rsBirthDays.Recordset
        If boolNewRecord Then
            .Fields("ChildNo") = glChildNo
            .Fields("BirthDayDate") = Format(Date1.Text, "dd.mm.yyyy")
            .Update
            FillList3
            .Bookmark = .LastModified
            boolNewRecord = False
            RichTextBox1.SetFocus
        End If
    End With
End Sub

Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1.SelText = vbTab
        KeyCode = 0
    End If
End Sub
Private Sub RichTextBox1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub

Private Sub RichTextBox1_SelChange()
    Call RichTextSelChange(frmBirthDates.RichTextBox1)
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        If SelectDays Then
            FillList3
            List3.ListIndex = 0
        End If
    Case 1
        If SelectPictures Then
            FillList2
            List2.ListIndex = 0
        End If
    Case Else
    End Select
End Sub
Private Sub Text2_LostFocus()
    On Error Resume Next
    If boolNewRecord Then
        If Not IsDate((Date1.Text)) Then
           boolNewRecord = False
           rsBirthDayPictures.Recordset.Move 0  'cancel AddNew
           Exit Sub
        End If
        With rsBirthDayPictures.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("Date") = CDate(Date1.Text)
            If Len(Text2.Text) <> 0 Then
                .Fields("PictureCaption") = Trim(Text2.Text)
            End If
            .Update
            FillList2
            .Bookmark = .LastModified
            boolNewRecord = False
        End With
    End If
End Sub
