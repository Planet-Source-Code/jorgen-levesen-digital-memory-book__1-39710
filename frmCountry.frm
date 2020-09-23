VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCountry 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Language"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   Icon            =   "frmCountry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   8388608
      TabCaption(0)   =   "Choose a Language"
      TabPicture(0)   =   "frmCountry.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnOK"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "List1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "List1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Timer1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Change Country Information"
      TabPicture(1)   =   "frmCountry.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CMD1"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "rsCountry"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "bDelete"
      Tab(1).Control(5)=   "bNew"
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(7)=   "List2"
      Tab(1).ControlCount=   8
      Begin VB.Timer Timer1 
         Interval        =   60
         Left            =   6120
         Top             =   5280
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4515
         Index           =   1
         ItemData        =   "frmCountry.frx":047A
         Left            =   4560
         List            =   "frmCountry.frx":047C
         TabIndex        =   33
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   4515
         Index           =   0
         ItemData        =   "frmCountry.frx":047E
         Left            =   840
         List            =   "frmCountry.frx":0480
         TabIndex        =   32
         Top             =   840
         Width           =   3735
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   5295
         ItemData        =   "frmCountry.frx":0482
         Left            =   -74880
         List            =   "frmCountry.frx":0489
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   720
         TabIndex        =   30
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   4800
         TabIndex        =   29
         Top             =   5520
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   -72960
         TabIndex        =   18
         Top             =   360
         Width           =   4455
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CountryPrefix"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   23
            Top             =   1680
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CountryFix"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   22
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "ExchangeRate"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   21
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Currency"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "Country"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   2400
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Telephone Prefix:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Country Short:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Exchange Rate:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Currency:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Country Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton bNew 
         Height          =   615
         Left            =   -69840
         Picture         =   "frmCountry.frx":0494
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "New Country Allowance"
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton bDelete 
         Height          =   615
         Left            =   -69120
         Picture         =   "frmCountry.frx":079E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Delete this Country Allowance"
         Top             =   5280
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Allowance"
         Height          =   1815
         Left            =   -72960
         TabIndex        =   7
         Top             =   2400
         Width           =   4455
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "MoreThanTwelveHours"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   3000
            TabIndex        =   11
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SixToTwelveHours"
            DataSource      =   "rsCountry"
            Height          =   285
            Left            =   3000
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "BoardAllowance"
            DataSource      =   "rsCountry"
            Height          =   285
            Index           =   1
            Left            =   3000
            TabIndex        =   9
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "NightAllowance"
            DataSource      =   "rsCountry"
            Height          =   285
            Index           =   0
            Left            =   3000
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "More Than 12 Hours:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "6 - 12 Hours:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Boarding Allowance:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Night Allowance:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Data rsCountry 
         Appearance      =   0  'Flat
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   315
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Country"
         Top             =   480
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame Frame3 
         Caption         =   "Flag"
         Height          =   1095
         Left            =   -72960
         TabIndex        =   2
         Top             =   4200
         Width           =   4455
         Begin VB.PictureBox Picture1 
            DataField       =   "CountryFlag"
            DataSource      =   "rsCountry"
            Height          =   735
            Left            =   3120
            ScaleHeight     =   675
            ScaleWidth      =   1035
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton btnPaste 
            Height          =   615
            Left            =   2280
            Picture         =   "frmCountry.frx":0AA8
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Copy picture from clipboard"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton btnOpen 
            Height          =   615
            Left            =   1560
            Picture         =   "frmCountry.frx":1112
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Read picture from disk"
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton btnDeletePicture 
            Height          =   615
            Left            =   840
            Picture         =   "frmCountry.frx":177C
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Delete picture"
            Top             =   360
            Width           =   615
         End
      End
      Begin MSComDlg.CommonDialog CMD1 
         Left            =   -70680
         Top             =   5400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Choose one of the folowing Countries:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BookAllowance() As Variant
Dim CountryBookmark() As Variant
Dim boolSelected As Boolean
Dim iNewRecord As Integer
Dim rsMyRec As Recordset
Dim rsLanguage As Recordset
Private Sub LoadList()
    List1(0).Clear
    List1(1).Clear
    List2.Clear
    With rsCountry.Recordset
        .MoveLast
        .MoveFirst
        ReDim CountryBookmark(.RecordCount)
        Do While Not .EOF
            List1(0).AddItem .Fields("Country")
            List1(1).AddItem .Fields("CountryFix")
            List2.AddItem .Fields("Country")
            List2.ItemData(List2.NewIndex) = List2.ListCount - 1
            CountryBookmark(List2.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub
Private Sub ShowText()
Dim strHelp As String
    On Error Resume Next
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
                If IsNull(.Fields("label6")) Then
                    .Fields("label6") = Label6.Caption
                Else
                    Label6.Caption = .Fields("label6")
                End If
                If IsNull(.Fields("label7")) Then
                    .Fields("label7") = Label7.Caption
                Else
                    Label7.Caption = .Fields("label7")
                End If
                If IsNull(.Fields("label8")) Then
                    .Fields("label8") = Label8.Caption
                Else
                    Label8.Caption = .Fields("label8")
                End If
                If IsNull(.Fields("label9")) Then
                    .Fields("label9") = Label9.Caption
                Else
                    Label9.Caption = .Fields("label9")
                End If
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("bNew")) Then
                    .Fields("bNew") = bNew.ToolTipText
                Else
                    bNew.ToolTipText = .Fields("bNew")
                End If
                If IsNull(.Fields("bDelete")) Then
                    .Fields("bDelete") = bDelete.ToolTipText
                Else
                    bDelete.ToolTipText = .Fields("bDelete")
                End If
                If IsNull(.Fields("label10")) Then
                    .Fields("label10") = Label10.Caption
                Else
                    Label10.Caption = .Fields("label10")
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
                If IsNull(.Fields("btnCancel")) Then
                    .Fields("btnCancel") = btnCancel.Caption
                Else
                    btnCancel.Caption = .Fields("btnCancel")
                End If
                If IsNull(.Fields("btnOK")) Then
                    .Fields("btnOK") = btnOK.Caption
                Else
                    btnOK.Caption = .Fields("btnOK")
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
        .Fields("label1") = Label1.Caption
        .Fields("label2") = Label2.Caption
        .Fields("label3") = Label3.Caption
        .Fields("label4") = Label4.Caption
        .Fields("label5") = Label5.Caption
        .Fields("label6") = Label6.Caption
        .Fields("label7") = Label7.Caption
        .Fields("label8") = Label8.Caption
        .Fields("label9") = Label9.Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("bNew") = bNew.ToolTipText
        .Fields("bDelete") = bDelete.ToolTipText
        .Fields("label10") = Label10.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("btnCancel") = btnCancel.Caption
        .Fields("btnOK") = btnOK.Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Private Sub bDelete_Click()
'Dim DgDef, Msg, Response, Title
    'If iNewRecord = 1 Then Exit Sub
    'On Error GoTo ErrbDelete_Click
    'DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    'Title = "Delete Record"
    'Msg = "Do you really want to delete this Country ?"
    'Beep
    'Response = MsgBox(Msg, Title)
    'If Response = IdNo Then
        'Exit Sub
    'End If
    'On Error Resume Next
    'delete this Country
    'rsCountry.Recordset.Delete
    'Beep
    'MsgBox "Country is deleted !!"
    'Exit Sub
'ErrbDelete_Click:
    'Beep
    'MsgBox Error$, 48, "Delete Country"
    'Resume ErrbDelete_Click2
'ErrbDelete_Click2:
End Sub
Private Sub bNew_Click()
    On Error Resume Next
    If iNewRecord = 1 Then Exit Sub
    rsCountry.Recordset.AddNew
    iNewRecord = 1
    Text2.SetFocus
End Sub
Private Sub btnDeletePicture_Click()
    Picture1.Picture = LoadPicture()
End Sub

Private Sub btnOk_Click()
    On Error Resume Next
    If boolSelected = True Then
        With rsMyRec
            .Edit
            .Fields("LanguageScreen") = List1(1).List(List1(1).ListIndex)
            .Update
        End With
        FileExt = List1(1).List(List1(1).ListIndex)
    End If
    Call MDIMasterKid.ReadText
    Call MDIMasterKid.LoadMenu1
    Call MDIMasterKid.ShowMenu
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub btnOpen_Click()
    With CMD1
        .filename = ""
        .DialogTitle = "Load Picture from disk"
        .Filter = "(*.bmp)|*.bmp|(*.pcx)|*.pcx|(*.jpg)|*.jpg"
        .FilterIndex = 1
        .ShowOpen
        Picture1.Picture = LoadPicture(.filename)
    End With
End Sub

Private Sub btnPaste_Click()
    Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Form_Activate()
    'On Error Resume Next
    rsCountry.Refresh
    Call ShowText
    boolSelected = False
    LoadList
    MDIMasterKid.cmbChildren.Visible = False
    MDIMasterKid.Label1.Visible = False
    Me.MousePointer = Default
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsCountry.DatabaseName = dbKidsTxt
    Set rsMyRec = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmCountry")
    iWhichForm = 11
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmCountry: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCountry.Recordset.Close
    rsLanguage.Close
    rsMyRec.Close
    iWhichForm = 0
    MDIMasterKid.cmbChildren.Visible = True
    MDIMasterKid.Label1.Visible = True
    HideAllButtons
    Erase CountryBookmark
    Set frmCountry = Nothing
End Sub

Private Sub List1_Click(Index As Integer)
    On Error Resume Next
    List1(0).ListIndex = List1(Index).ListIndex
    List1(1).ListIndex = List1(Index).ListIndex
    boolSelected = True
End Sub

Private Sub List2_Click()
    On Error Resume Next
    rsCountry.Recordset.Bookmark = CountryBookmark(List2.ItemData(List2.ListIndex))
End Sub
Private Sub Text2_LostFocus()
    If iNewRecord = 1 Then
        On Error GoTo errText2_Click
        With rsCountry.Recordset
            .Fields("Country") = Text2.Text
            .Update
            .Bookmark = rsCountry.Recordset.LastModified
        End With
        iNewRecord = 0
        Text3.SetFocus
    End If
    Exit Sub
    
errText2_Click:
    Beep
    MsgBox Error$, 48, "New Record"
    rsCountry.Recordset.CancelUpdate
    iNewRecord = 0
    Frame2.Visible = False
    Resume errText2_Click2
errText2_Click2:
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    List1(1).TopIndex = List1(0).TopIndex
End Sub


