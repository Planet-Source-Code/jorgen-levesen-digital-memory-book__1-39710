VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFrames 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clip Art"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Clip Arts"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.Frame Frame3 
         Caption         =   "Picture"
         Height          =   5055
         Left            =   2160
         TabIndex        =   2
         Top             =   120
         Width           =   4215
         Begin VB.CommandButton btnDefaultPic 
            Caption         =   "&Set picture as default"
            Height          =   615
            Left            =   120
            TabIndex        =   9
            Top             =   4320
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FrameName"
            DataSource      =   "rsClipArt"
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   7
            Top             =   3960
            Width           =   3255
         End
         Begin VB.CommandButton btnScan 
            Height          =   615
            Index           =   0
            Left            =   3600
            Picture         =   "frmFrames.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Scan a picture"
            Top             =   4320
            Width           =   495
         End
         Begin VB.CommandButton btnCopyPic 
            Height          =   615
            Index           =   0
            Left            =   2640
            Picture         =   "frmFrames.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Copy picture to the Clipboard"
            Top             =   4320
            Width           =   495
         End
         Begin VB.CommandButton btnPastePicture 
            Height          =   615
            Index           =   0
            Left            =   2160
            Picture         =   "frmFrames.frx":080C
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Paste picture"
            Top             =   4320
            Width           =   495
         End
         Begin VB.CommandButton btnReadFromFile 
            Height          =   615
            Index           =   0
            Left            =   3120
            Picture         =   "frmFrames.frx":0ECE
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Read picture from file"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Text:"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3960
            Width           =   615
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "FramePicture"
            DataSource      =   "rsClipArt"
            Height          =   3495
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4710
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.Data rsClipArt 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKidPic.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ClipArt"
         Top             =   4920
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSComDlg.CommonDialog Cmd1 
         Left            =   1200
         Top             =   4800
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBookmark() As Variant
Dim rsMyRecord As Recordset
Dim rsLanguage As Recordset
Public Sub NewFrames()
    boolNewRecord = True
    rsClipArt.Recordset.AddNew
    Text1.SetFocus
End Sub

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
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("label1")
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
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
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
        .Fields("label1") = Label1.Caption
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        .Fields("Frame1") = Frame1.Caption
        .Fields("Msg1") = "Default Print - Section Picture:"
        .Fields("Msg2") = "Picture:"
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Public Sub FillList1()
    On Error Resume Next
    List1.Clear
    With rsClipArt.Recordset
        .MoveLast
        .MoveFirst
        ReDim vBookmark(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("FrameName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vBookmark(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Private Sub btnCopyPic_Click(Index As Integer)
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetData Image1.Picture, vbCFDIB
End Sub
Public Sub DeleteClipArt()
    On Error Resume Next
    rsClipArt.Recordset.Delete
    FillList1
End Sub

Private Sub btnDefaultPic_Click()
    With rsMyRecord
        .Edit
        .Fields("SectionPicID") = CLng(rsClipArt.Recordset.Fields("LineNo"))
        .Update
    End With
    Frame3.Caption = rsLanguage.Fields("Msg1")
End Sub

Private Sub btnPastePicture_Click(Index As Integer)
    On Error Resume Next
    Image1.Picture = Clipboard.GetData(vbCFBitmap)
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

Private Sub Form_Activate()
    On Error Resume Next
    rsClipArt.Refresh
    Dither Me
    ShowText
    FillList1
    ShowAllButtons
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsClipArt.DatabaseName = dbKidPicTxt
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmFrames")
    iWhichForm = 45
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
    rsClipArt.Recordset.Close
    rsMyRecord.Close
    rsLanguage.Close
    HideAllButtons
    iWhichForm = 0
    Erase vBookmark
    Set frmFrames = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsClipArt.Recordset.Bookmark = vBookmark(List1.ItemData(List1.ListIndex))
    If CLng(rsMyRecord.Fields("SectionPicID")) = CLng(rsClipArt.Recordset.Fields("LineNo")) Then
        Frame3.Caption = rsLanguage.Fields("Msg1")
    Else
        Frame3.Caption = rsLanguage.Fields("Msg2")
    End If
End Sub

Private Sub Text1_LostFocus()
   On Error Resume Next
   If boolNewRecord Then
        With rsClipArt.Recordset
            .Fields("FrameName") = Trim(Text1.Text)
            .Update
            FillList1
            .Bookmark = .LastModified
            boolNewRecord = False
        End With
    End If
End Sub
