VERSION 5.00
Begin VB.Form frmAllNameExplanation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Names and Explanations"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   6075
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   3360
      Width           =   4815
      Begin VB.Data rsNamesExplanation2 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "NamesExplanation"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "NameExplanation"
         DataSource      =   "rsNamesExplanation2"
         Height          =   2295
         Index           =   1
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "English"
      Height          =   2895
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Data rsNamesExplanation1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "NamesExplanation"
         Top             =   2280
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "NameExplanation"
         DataSource      =   "rsNamesExplanation1"
         Height          =   2295
         Index           =   0
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Data rsNames 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Names"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "frmAllNameExplanation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v1RecordBookmark() As Variant
Dim rsLanguage As Recordset
Private Sub FillList1()
    On Error Resume Next
    List1.Clear
    With rsNames.Recordset
        .MoveLast
        .MoveFirst
        ReDim v1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("FirstName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            v1RecordBookmark(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
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
                If IsNull(.Fields("Frame1(0)")) Then
                    .Fields("Frame1(0)") = Frame1(0).Caption
                Else
                    Frame1(0).Caption = .Fields("Frame1(0)")
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
        .Fields("Frame1(0)") = Frame1(0).Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub SelectRecords()
Dim Sql As String
    'On Error Resume Next
    Sql = "SELECT * FROM NamesExplanation WHERE Trim(FirstName) ="
    Sql = Sql & Chr(34) & Trim(rsNames.Recordset.Fields("FirstName")) & Chr(34)
    Sql = Sql & "AND Language ="
    Sql = Sql & Chr(34) & "ENG" & Chr(34)
    rsNamesExplanation1.RecordSource = Sql
    rsNamesExplanation1.Refresh
    
    With rsNamesExplanation1.Recordset
        If .RecordCount = 0 Then
            'we did not have any explanation record for this name, make one
            .AddNew
            .Fields("Language") = "ENG"
            .Fields("FirstName") = Trim(rsNames.Recordset.Fields("FirstName"))
            .Update
        End If
    End With
    
    Sql = "SELECT * FROM NamesExplanation WHERE Trim(FirstName) ="
    Sql = Sql & Chr(34) & Trim(rsNames.Recordset.Fields("FirstName")) & Chr(34)
    Sql = Sql & "AND Language ="
    Sql = Sql & Chr(34) & FileExt & Chr(34)
    rsNamesExplanation2.RecordSource = Sql
    rsNamesExplanation2.Refresh
    
    With rsNamesExplanation2.Recordset
        If .RecordCount = 0 Then
            'we did not have any explanation record for this name, make one
            .AddNew
            .Fields("Language") = FileExt
            .Fields("FirstName") = Trim(rsNames.Recordset.Fields("FirstName"))
            .Update
        End If
    End With
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    rsNames.Refresh
    rsNamesExplanation1.Refresh
    rsNamesExplanation2.Refresh
    Frame1(1).Caption = FileExt
    ShowText
    Dither Me
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsNames.DatabaseName = dbKidsTxt
    rsNamesExplanation1.DatabaseName = dbKidsTxt
    rsNamesExplanation2.DatabaseName = dbKidsTxt
    Set rsLanguage = dbKidLang.OpenRecordset("frmAllNameExplanation")
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
    rsNames.Recordset.Close
    rsNamesExplanation1.Recordset.Close
    rsNamesExplanation2.Recordset.Close
    rsLanguage.Close
    Set frmAllNameExplanation = Nothing
End Sub


Private Sub List1_Click()
    On Error Resume Next
    rsNames.Recordset.Bookmark = v1RecordBookmark(List1.ItemData(List1.ListIndex))
    SelectRecords
End Sub


