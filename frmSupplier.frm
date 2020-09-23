VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSupplier 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programme Supplier"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8388608
      TabCaption(0)   =   "Page 1"
      TabPicture(0)   =   "frmSupplier.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Page 2"
      TabPicture(1)   =   "frmSupplier.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4335
         Left            =   -74640
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   7095
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FTPRemoteAddress"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   13
            Left            =   3600
            MaxLength       =   100
            TabIndex        =   35
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FTPUserName"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   14
            Left            =   3600
            MaxLength       =   100
            TabIndex        =   34
            Top             =   720
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FTPPassWord"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   15
            Left            =   3600
            MaxLength       =   100
            TabIndex        =   33
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "FTP Remote Address:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "FTP User Name:"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   37
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "FTP Password:"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   7335
         Begin VB.CommandButton btnPastePicture 
            Height          =   495
            Left            =   240
            Picture         =   "frmSupplier.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Paste picture"
            Top             =   2160
            Width           =   495
         End
         Begin VB.CommandButton btnDelete 
            Height          =   495
            Left            =   1200
            Picture         =   "frmSupplier.frx":06FA
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Delete picture"
            Top             =   2160
            Width           =   495
         End
         Begin VB.CommandButton btnReadFromFile 
            Height          =   495
            Left            =   2040
            Picture         =   "frmSupplier.frx":0844
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Read picture from file"
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton btnCopyPic 
            Height          =   495
            Left            =   720
            Picture         =   "frmSupplier.frx":0F06
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Copy picture to the Clipboard"
            Top             =   2160
            Width           =   495
         End
         Begin VB.CommandButton btnScan 
            Height          =   495
            Left            =   2040
            Picture         =   "frmSupplier.frx":15C8
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Scan a picture"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierProgrammeUppdate"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   12
            Left            =   3960
            MaxLength       =   100
            TabIndex        =   14
            Top             =   4080
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierEmailySysResponse"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   11
            Left            =   3960
            MaxLength       =   100
            TabIndex        =   13
            Top             =   3600
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierEmail"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   10
            Left            =   3960
            MaxLength       =   100
            TabIndex        =   12
            Top             =   2880
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierFax"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   9
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   11
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierPhone"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   8
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   10
            Top             =   2400
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierContact"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   7
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3360
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierCountry"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   6
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   8
            Top             =   1920
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierTown"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   5
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierZip"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   4
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierAddr3"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   3
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierAddr2"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   2
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   4
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierAddr1"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   1
            Left            =   3960
            MaxLength       =   50
            TabIndex        =   3
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "SupplierName"
            DataSource      =   "rsSupplier"
            Height          =   285
            Index           =   0
            Left            =   3960
            MaxLength       =   80
            TabIndex        =   2
            Top             =   240
            Width           =   3255
         End
         Begin VB.Data rsSupplier 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Supplier"
            Top             =   3720
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSComDlg.CommonDialog Cmd1 
            Left            =   240
            Top             =   4080
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Logo:"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   1695
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "SupplierLogo"
            DataSource      =   "rsSupplier"
            Height          =   1575
            Left            =   240
            Stretch         =   -1  'True
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "URL Programme Update:"
            Height          =   255
            Index           =   10
            Left            =   1560
            TabIndex        =   30
            Top             =   4080
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Email:"
            Height          =   255
            Index           =   9
            Left            =   1920
            TabIndex        =   29
            Top             =   3600
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Email:"
            Height          =   255
            Index           =   8
            Left            =   2040
            TabIndex        =   28
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fax No.:"
            Height          =   255
            Index           =   7
            Left            =   2040
            TabIndex        =   27
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Phone No.:"
            Height          =   255
            Index           =   6
            Left            =   2040
            TabIndex        =   26
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Contact Person:"
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   25
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Country:"
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   24
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Town:"
            Height          =   255
            Index           =   3
            Left            =   2040
            TabIndex        =   23
            Top             =   1680
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Zip Code:"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   22
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Address:"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   21
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Company:"
            Height          =   255
            Index           =   0
            Left            =   1920
            TabIndex        =   20
            Top             =   240
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
                For i = 0 To 11
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("btnPastePicture")) Then
                    .Fields("btnPastePicture") = btnPastePicture.ToolTipText
                Else
                    btnPastePicture.ToolTipText = .Fields("btnPastePicture")
                End If
                If IsNull(.Fields("btnCopyPic")) Then
                    .Fields("btnCopyPic") = btnCopyPic.ToolTipText
                Else
                    btnCopyPic.ToolTipText = .Fields("btnCopyPic")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete.ToolTipText
                Else
                    btnDelete.ToolTipText = .Fields("btnDelete")
                End If
                If IsNull(.Fields("btnReadFromFile")) Then
                    .Fields("btnReadFromFile") = btnReadFromFile.ToolTipText
                Else
                    btnReadFromFile.ToolTipText = .Fields("btnReadFromFile")
                End If
                If IsNull(.Fields("btnScan")) Then
                    .Fields("btnScan") = btnScan.ToolTipText
                Else
                    btnScan.ToolTipText = .Fields("btnScan")
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
        For i = 0 To 11
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnPastePicture") = btnPastePicture.ToolTipText
        .Fields("btnCopyPic") = btnCopyPic.ToolTipText
        .Fields("btnDelete") = btnDelete.ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile.ToolTipText
        .Fields("btnScan") = btnScan.ToolTipText
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub btnCopyPic_Click()
    On Error Resume Next
    Clipboard.SetData Image1.Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click()
    On Error Resume Next
    Set Image1.Picture = LoadPicture()
End Sub


Private Sub btnPastePicture_Click()
    On Error Resume Next
    Image1.Picture = Clipboard.GetData(vbCFDIB)
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
        Set Image1.Picture = LoadPicture(Cmd1.filename)
End Sub

Private Sub btnScan_Click()
    Dim Ret As Long, t As Single
    On Error Resume Next
    Ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsSupplier.Refresh
    ShowText
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    DoEvents
    rsSupplier.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmSupplier")
    iWhichForm = 19
    Exit Sub
    
errForm_Load:
    Beep
    Me.MousePointer = Default
    MsgBox Error$, vbCritical, "Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSupplier.Recordset.Close
    rsLanguage.Close
    iWhichForm = 0
    Set frmSupplier = Nothing
End Sub

Private Sub VideoSoftIndexTab1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If NewTab = 1 Then
        frmSecurity.Show 1
        If Password Then
            Password = False
            Frame2.Visible = True
            Exit Sub
        Else
            Cancel = 1
        End If
    Else
        Frame2.Visible = False
    End If
End Sub
