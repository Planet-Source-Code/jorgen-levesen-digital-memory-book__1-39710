VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmTerm 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculate The Approximately Term Week"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar Month1 
      Height          =   3375
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   6255
      _Version        =   524288
      _ExtentX        =   11033
      _ExtentY        =   5953
      _StockProps     =   1
      BackColor       =   33023
      Year            =   2002
      Month           =   9
      Day             =   6
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   16777215
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btnStore 
      Caption         =   "&Store Data"
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
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   6255
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Data rsHoroscope 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Horoscope"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Menstruation cycle "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   4335
      Begin VB.OptionButton Option1 
         BackColor       =   &H000080FF&
         Caption         =   "32 days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000080FF&
         Caption         =   "30 days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000080FF&
         Caption         =   "28 days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      DataField       =   "HoroscopeShort"
      DataSource      =   "rsHoroscope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      DataField       =   "HoroscopeText"
      DataSource      =   "rsHoroscope"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ZodiacSignPicture"
      DataSource      =   "rsHoroscope"
      Height          =   615
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   6000
      Picture         =   "frmTerm.frx":0000
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmTerm.frx":08CA
      Stretch         =   -1  'True
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Your term is calculated to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Click the day of the first day in your last menstruation period"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDate As Variant
Dim rsMyRecord As Recordset
Dim rsLanguage As Recordset
Private Sub CalcHoroscope()
Dim sString, sYear As String, Date1 As Date, Date2 As Date
    sYear = Right(vDate, 4)
    Date1 = CDate(vDate)
    With rsHoroscope.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                sString = .Fields("FromDate")
                sString = sString & "." & sYear
                Date2 = CDate(sString)
                If compareDates(Date1, Date2) = 1 Then
                    sString = .Fields("ToDate")
                    sString = sString & "." & sYear
                    Date2 = CDate(sString)
                    If compareDates(Date1, Date2) = 2 Then
                        Image2.Visible = True
                        Label3(0).Visible = True
                        Label3(1).Visible = True
                        Exit Do
                    End If
                End If
            End If
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
                If IsNull(.Fields("Label1")) Then
                    .Fields("Label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("Label1")
                End If
                If IsNull(.Fields("Label2")) Then
                    .Fields("Label2") = Label2.Caption
                Else
                    Label2.Caption = .Fields("Label2")
                End If
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("btnCalculate")) Then
                    .Fields("btnCalculate") = btnCalculate.Caption
                Else
                    btnCalculate.Caption = .Fields("btnCalculate")
                End If
                If IsNull(.Fields("btnStore")) Then
                    .Fields("btnStore") = btnStore.Caption
                Else
                    btnStore.Caption = .Fields("btnStore")
                End If
                If IsNull(.Fields("Days")) Then
                    .Fields("Days") = "Days"
                Else
                    Option1(0).Caption = "28" & " " & .Fields("Days")
                    Option1(1).Caption = "30" & " " & .Fields("Days")
                    Option1(2).Caption = "32" & " " & .Fields("Days")
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
        .Fields("Label1") = Label1.Caption
        .Fields("Label2") = Label2.Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("btnCalculate") = btnCalculate.Caption
        .Fields("btnStore") = btnStore.Caption
        .Fields("Days") = "Days"
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub btnCalculate_Click()
    vDate = Month1.Value
    vDate = DateAdd("m", 9, vDate)
    If Option1(0).Value = True Then
        vDate = DateAdd("d", 7, vDate)
    ElseIf Option1(1).Value = True Then
        vDate = DateAdd("d", 9, vDate)
    Else
        vDate = DateAdd("d", 11, vDate)
    End If
    Label2.Caption = rsLanguage.Fields("Label2") & "  " & vDate
    Label2.Visible = True
    Image1(0).Visible = True
    Image1(1).Visible = True
    btnStore.Visible = True
    CalcHoroscope
End Sub

Private Sub btnStore_Click()
    With rsMyRecord
        .Edit
        .Fields("TermDueDate") = Format(vDate, "dd.mm.yyyy")
        .Update
    End With
    btnStore.Visible = False
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsHoroscope.Refresh
    ShowText
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsHoroscope.DatabaseName = dbKidsTxt
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsLanguage = dbKidLang.OpenRecordset("frmTerm")
    iWhichForm = 9
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmTerm: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsHoroscope.Recordset.Close
    rsMyRecord.Close
    rsLanguage.Close
    iWhichForm = 0
    Set frmTerm = Nothing
End Sub

Private Sub Month1_Click()
    Image2.Visible = False
    Label3(0).Visible = False
    Label3(1).Visible = False
End Sub

