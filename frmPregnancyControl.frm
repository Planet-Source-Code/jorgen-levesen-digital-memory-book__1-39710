VERSION 5.00
Begin VB.Form frmPregnancyControl 
   BackColor       =   &H000080FF&
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   9825
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Height          =   8415
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7455
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ControlDate"
         DataSource      =   "rsPregnancyControl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbMidwife 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "MidwifeName"
         DataSource      =   "rsPregnancyControl"
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   750
         Width           =   3900
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Results"
         DataSource      =   "rsPregnancyControl"
         Height          =   1455
         Index           =   0
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1140
         Width           =   5040
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "MidwifeComments"
         DataSource      =   "rsPregnancyControl"
         Height          =   1455
         Index           =   1
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2700
         Width           =   5040
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "MyWeight"
         DataSource      =   "rsPregnancyControl"
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   4
         Top             =   4260
         Width           =   825
      End
      Begin VB.ComboBox cmbDim1 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MyWeightDim"
         DataSource      =   "rsPregnancyControl"
         Height          =   315
         Index           =   0
         Left            =   3090
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   4260
         Width           =   1155
      End
      Begin VB.ComboBox cmbDim1 
         BackColor       =   &H00FFFFC0&
         DataField       =   "MyTommyDim"
         DataSource      =   "rsPregnancyControl"
         Height          =   315
         Index           =   1
         Left            =   3090
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   4650
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "MyTommy"
         DataSource      =   "rsPregnancyControl"
         Height          =   315
         Index           =   3
         Left            =   2160
         TabIndex        =   6
         Top             =   4650
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Questions"
         DataSource      =   "rsPregnancyControl"
         Height          =   1440
         Index           =   4
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   5040
         Width           =   5040
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "OwnNotes"
         DataSource      =   "rsPregnancyControl"
         Height          =   1440
         Index           =   5
         Left            =   2160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   6600
         Width           =   5040
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Control Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Midwife Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   750
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Control result:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1140
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Midwifes Comments:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   2700
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Weight at control date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   4245
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "My stomach dimension:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   4665
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Question asked:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   5085
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         Caption         =   "Other Notes this control:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   6600
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Control Dates:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   7680
      TabIndex        =   10
      Top             =   120
      Width           =   2055
      Begin VB.Data rsPregnancyControl 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   690
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PregnancyControl"
         Top             =   7350
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   7830
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmPregnancyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim v1RecordBookmark() As Variant
Dim rsMidwife As Recordset
Dim rsDimLength As Recordset
Dim rsDimWeight As Recordset
Dim rsLanguage As Recordset
Public Sub NewPregnancyControl()
    rsPregnancyControl.Recordset.AddNew
    Date1.SetFocus
    boolNewRecord = True
End Sub
Public Function SelectControl() As Boolean
Dim Sql As String
    On Error GoTo errSelectControl
    Sql = "SELECT * FROM PregnancyControl WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsPregnancyControl.RecordSource = Sql
    rsPregnancyControl.Refresh
    rsPregnancyControl.Recordset.MoveFirst
    SelectControl = True
    Exit Function
    
errSelectControl:
    SelectControl = False
    Err.Clear
End Function

Public Sub WritePregnancyControlWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    rsPregnancyControl.Recordset.MoveFirst
    With wdApp
        Do While Not rsPregnancyControl.Recordset.EOF
            If CLng(rsPregnancyControl.Recordset.Fields("ChildNo")) = glChildNo Then
                .Selection.TypeText Text:=Label1(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(rsPregnancyControl.Recordset.Fields("ControlDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(1).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsPregnancyControl.Recordset.Fields("MidwifeName"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(2).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsPregnancyControl.Recordset.Fields("Results"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(3).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsPregnancyControl.Recordset.Fields("MidwifeComments"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(4).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsPregnancyControl.Recordset.Fields("MyWeight") & " " & Format(rsPregnancyControl.Recordset.Fields("MyWeightDim"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(5).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsPregnancyControl.Recordset.Fields("MyTommy") & " " & Format(rsPregnancyControl.Recordset.Fields("MyTommyDim"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(6).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsPregnancyControl.Recordset.Fields("Questions"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label1(7).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsPregnancyControl.Recordset.Fields("OwnNotes"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        rsPregnancyControl.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub

Public Sub WritePregnancyControl()
    On Error GoTo errWritePregnancyControl
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
    
    With rsPregnancyControl.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = glChildNo Then
                cPrint.FontBold = True
                cPrint.pPrint Label1(0).Caption, 1, True
                If IsDate(.Fields("ControlDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("ControlDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.FontBold = False
                cPrint.pPrint Label1(1).Caption, 1, True
                If Len(cmbMidwife.Text) <> 0 Then
                    cPrint.pPrint cmbMidwife.Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(2).Caption, 1, True
                If Len(Text1(0).Text) <> 0 Then
                    cPrint.pMultiline Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(3).Caption, 1, True
                If Len(Text1(1).Text) <> 0 Then
                    cPrint.pMultiline Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(4).Caption, 1, True
                If Len(Text1(2).Text) <> 0 Then
                    cPrint.pPrint Text1(2).Text & "  " & cmbDim1(0).Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(5).Caption, 1, True
                If Len(Text1(3).Text) <> 0 Then
                    cPrint.pPrint Text1(3).Text & "  " & cmbDim1(1).Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(6).Caption, 1, True
                If Len(Text1(4).Text) <> 0 Then
                    cPrint.pMultiline Text1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
                cPrint.pPrint Label1(7).Caption, 1, True
                If Len(Text1(5).Text) <> 0 Then
                    cPrint.pMultiline Text1(5).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    Call PrintFront
                End If
            End If
        .MoveNext
        Loop
    End With
    
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
    Exit Sub
    
errWritePregnancyControl:
    Beep
    MsgBox Err.Description, vbExclamation, "Print Pregnancy Control"
    Err.Clear
End Sub

Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                For i = 0 To 7
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
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
            
        .AddNew
        .Fields("Language") = FileExt
        For i = 0 To 7
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("Frame1") = Frame1.Caption
        .Fields("FormName") = "Pregnancy Control"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub FillDim()
    On Error Resume Next
    cmbDim1(1).Clear
    With rsDimLength
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbDim1(1).AddItem .Fields("LengthDim")
            End If
        .MoveNext
        Loop
    End With
    
    cmbDim1(0).Clear
    With rsDimWeight
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbDim1(0).AddItem .Fields("WeightDim")
            End If
        .MoveNext
        Loop
    End With
End Sub


Public Sub FillList1()
    On Error Resume Next
    List1.Clear
    With rsPregnancyControl.Recordset
        .MoveLast
        .MoveFirst
        ReDim v1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("ControlDate")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            v1RecordBookmark(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub FillcmbMidwife()
    On Error Resume Next
    cmbMidwife.Clear
    With rsMidwife
        .MoveFirst
        Do While Not .EOF
            cmbMidwife.AddItem .Fields("FirstName") & "  " & .Fields("LastName")
        .MoveNext
        Loop
    End With
End Sub

Private Sub cmbMidwife_LostFocus()
    On Error Resume Next
    If boolNewRecord Then
        With rsPregnancyControl.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("ControlDate") = CDate(Format(Date1.Text, "dd.mm.yyyy"))
            .Update
            FillList1
            .Bookmark = .LastModified
            boolNewRecord = False
            Text1(0).SetFocus
        End With
    End If
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

Private Sub Date1_LostFocus()
    If boolNewRecord Then
        'cmbMidwife.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsPregnancyControl.Refresh
    FillcmbMidwife
    FillDim
    SelectControl
    FillList1
    List1.ListIndex = 0
    ShowAllButtons
    ShowKids
    ShowText
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsPregnancyControl.DatabaseName = dbKidsTxt
    Set rsMidwife = dbKids.OpenRecordset("Midwife")
    Set rsDimLength = dbKids.OpenRecordset("DimLength")
    Set rsDimWeight = dbKids.OpenRecordset("DimWeight")
    Set rsLanguage = dbKidLang.OpenRecordset("frmPregnancyControl")
    iWhichForm = 5
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmPregnancyNotes: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsPregnancyControl.Recordset.Close
    rsLanguage.Close
    rsDimLength.Close
    rsDimWeight.Close
    rsMidwife.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Erase v1RecordBookmark
    Set frmPregnancyControl = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsPregnancyControl.Recordset.Bookmark = v1RecordBookmark(List1.ItemData(List1.ListIndex))
End Sub


