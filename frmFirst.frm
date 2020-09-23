VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFirst 
   BackColor       =   &H00C00000&
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   9540
   Begin VB.Data rsLanguage 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "frmFirst"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsMyrecord 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MyRecord"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Do not show in future"
      DataField       =   "ShowFirst"
      DataSource      =   "rsMyrecord"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   5880
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      DataField       =   "RichTextBox1"
      DataSource      =   "rsLanguage"
      Height          =   3495
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6165
      _Version        =   393217
      BackColor       =   16761024
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmFirst.frx":0000
   End
   Begin VB.Image Image1 
      Height          =   1695
      Index           =   1
      Left            =   7680
      Picture         =   "frmFirst.frx":00D5
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1695
      Index           =   0
      Left            =   480
      Picture         =   "frmFirst.frx":09EF
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MakeNewLanguage()
    With rsLanguage.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                Check1.Caption = .Fields("Check1")
                Exit Sub
            End If
        .MoveNext
        Loop
    
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Check1") = Check1.Caption
        If Len(RichTextBox1.Text) <> 0 Then
            .Fields("RichTextBox1") = RichTextBox1.Text
        End If
        .Update
        .Bookmark = .LastModified
    End With
End Sub
Public Sub PrintFirst()
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
    cPrint.FontBold = True
    cPrint.FontSize = 28
    cPrint.pCenter rsLanguage.Recordset.Fields("FormName")
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint
    cPrint.FontBold = False
    cPrint.FontSize = 12
    cPrint.pPrint RichTextBox1.Text, 0.3
    cPrint.pPrint
    
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    
    Set cPrint = Nothing
End Sub

Public Sub WriteFirst()
    On Error Resume Next
    Set wdApp = New Word.Application
    wdApp.Application.Visible = True
    wdApp.Application.WindowState = wdWindowStateMaximize
    wdApp.Caption = rsLanguage.Recordset.Fields("FormName")
    wdApp.Documents.Add DocumentType:=wdNewBlankDocument
    
    With wdApp
        .Selection.TypeText Text:=rsLanguage.Recordset.Fields("FormName")
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        Clipboard.Clear
        Clipboard.SetText RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
    End With
    Set wdApp = Nothing
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsMyRecord.Refresh
    rsLanguage.Refresh
    MakeNewLanguage
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rsMyRecord.DatabaseName = dbKidsTxt
    rsLanguage.DatabaseName = dbKidLangTxt
    iWhichForm = 35
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRecord.Recordset.Close
    rsLanguage.Recordset.Close
    HideAllButtons
    iWhichForm = 0
    Set frmFirst = Nothing
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
    Call RichTextSelChange(frmFirst.RichTextBox1)
End Sub
