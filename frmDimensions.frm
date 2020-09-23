VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDimensions 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dimensions"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   4194368
      TabCaption(0)   =   "Time"
      TabPicture(0)   =   "frmDimensions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rsCountry"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "rsTime"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DBGrid1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Length"
      TabPicture(1)   =   "frmDimensions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "rsLength"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Weight"
      TabPicture(2)   =   "frmDimensions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DBGrid1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "rsWeight"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Volum"
      TabPicture(3)   =   "frmDimensions.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DBGrid1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "rsVolume"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmDimensions.frx":0070
         Height          =   5055
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "frmDimensions.frx":0085
         TabIndex        =   1
         Top             =   480
         Width           =   4095
      End
      Begin VB.Data rsVolume 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DimVolume"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data rsWeight 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DimWeight"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data rsLength 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -74760
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DimLength"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data rsTime 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DimTime"
         Top             =   240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data rsCountry 
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
         RecordSource    =   "Country"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmDimensions.frx":0A5B
         Height          =   5055
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmDimensions.frx":0A72
         TabIndex        =   2
         Top             =   480
         Width           =   4095
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmDimensions.frx":1448
         Height          =   5055
         Index           =   2
         Left            =   -74880
         OleObjectBlob   =   "frmDimensions.frx":145F
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmDimensions.frx":1E35
         Height          =   5055
         Index           =   3
         Left            =   -74880
         OleObjectBlob   =   "frmDimensions.frx":1E4C
         TabIndex        =   4
         Top             =   480
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmDimensions"
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
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("Tab12") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab12")
                End If
                Tab1.Tab = 3
                If IsNull(.Fields("Tab13")) Then
                    .Fields("Tab13") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab13")
                End If
                Tab1.Tab = 0
                'If IsNull(.Fields("GridColn0")) Then
                    '.Fields("GridColn0") = Grid1.Columns(0).Caption
                'Else
                    'Grid1.Columns(0).Caption = .Fields("GridColn0")
                    'Grid2.Columns(0).Caption = .Fields("GridColn0")
                    'Grid3.Columns(0).Caption = .Fields("GridColn0")
                    'Grid4.Columns(0).Caption = .Fields("GridColn0")
                'End If
                'If IsNull(.Fields("GridColn1")) Then
                    '.Fields("GridColn1") = Grid1.Columns(1).Caption
                'Else
                    'Grid1.Columns(1).Caption = .Fields("GridColn1")
                    'Grid2.Columns(1).Caption = .Fields("GridColn1")
                    'Grid3.Columns(1).Caption = .Fields("GridColn1")
                    'Grid4.Columns(1).Caption = .Fields("GridColn1")
                'End If
                'If IsNull(.Fields("DropDown1Coln0")) Then
                    '.Fields("DropDown1Coln0") = DropDown1.Columns(0).Caption
                'Else
                    'DropDown1.Columns(0).Caption = .Fields("DropDown1Coln0")
                'End If
                'If IsNull(.Fields("DropDown1Coln1")) Then
                    '.Fields("DropDown1Coln1") = DropDown1.Columns(1).Caption
                'Else
                    'DropDown1.Columns(1).Caption = .Fields("DropDown1Coln1")
                'End If
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
        .Fields("Form") = Me.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 0
        '.Fields("GridColn0") = Grid1.Columns(0).Caption
        '.Fields("GridColn1") = Grid1.Columns(1).Caption
        '.Fields("DropDown1Coln0") = DropDown1.Columns(0).Caption
        '.Fields("DropDown1Coln1") = DropDown1.Columns(1).Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
     Dither Me
    rsTime.Refresh
    rsLength.Refresh
    rsWeight.Refresh
    rsCountry.Refresh
    rsVolume.Refresh
    ShowText
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsTime.DatabaseName = dbKidsTxt
    rsLength.DatabaseName = dbKidsTxt
    rsWeight.DatabaseName = dbKidsTxt
    rsCountry.DatabaseName = dbKidsTxt
    rsVolume.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmDimensions")
    iWhichForm = 28
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmDimensions:  Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsTime.Recordset.Close
    rsLength.Recordset.Close
    rsWeight.Recordset.Close
    rsVolume.Recordset.Close
    rsCountry.Recordset.Close
    rsLanguage.Close
    iWhichForm = 0
    Set frmDimensions = Nothing
End Sub


