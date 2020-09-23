VERSION 5.00
Begin VB.Form frmNotice 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alarm Notice"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmNotice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1200
      Width           =   5295
   End
   Begin VB.CommandButton btnSleep 
      Caption         =   "&Sleep"
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
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "&OK"
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
      Left            =   4200
      TabIndex        =   6
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton btnOkSleep 
      Height          =   495
      Left            =   3120
      Picture         =   "frmNotice.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cboTimeAfter 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "You have set the alarm:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Dim rsAlarm As Recordset
Dim rsLanguage As Recordset
Public Sub MakeWindowNotTop(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
       SWP_NOMOVE + SWP_NOSIZE
End Sub
Public Sub MakeWindowAlwaysTop(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
      SWP_NOMOVE + SWP_NOSIZE
End Sub
Private Sub LoadcboTimeAfter()
    With cboTimeAfter
        .AddItem "10 min"
        .AddItem "20 min"
        .AddItem "30 min"
        .AddItem "1 Hour"
        .AddItem "2 Hours"
        .AddItem "3 Hours"
        .AddItem "1 day"
        .Text = .List(.ListIndex = 0)
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
                If IsNull(.Fields("Label1(0)")) Then
                    .Fields("Label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("Label1(0)")
                End If
                If IsNull(.Fields("Label1(1)")) Then
                    .Fields("Label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("Label1(1)")
                End If
                If IsNull(.Fields("btnSleep")) Then
                    .Fields("btnSleep") = btnSleep.Caption
                Else
                    btnSleep.Caption = .Fields("btnSleep")
                End If
                If IsNull(.Fields("btnOK")) Then
                    .Fields("btnOK") = btnOk.Caption
                Else
                    btnOk.Caption = .Fields("btnOK")
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
        .Fields("Label1(0)") = Label1(0).Caption
        .Fields("Label1(1)") = Label1(1).Caption
        .Fields("btnSleep") = btnSleep.Caption
        .Fields("btnOK") = btnOk.Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub btnOk_Click()
    Unload Me
End Sub

Private Sub btnOkSleep_Click()
Dim sTime As String, sDate As String, vCompare As Variant
Dim IntervalType As String

    On Error Resume Next
    sTime = Format(CDate(Text2.Text), "hh:mm")
    sDate = Format(CDate(Text1.Text), "dd.mm.yyyy")
    
    If Len(cboTimeAfter.Text) = 0 Then Exit Sub
    
    'minute
    vCompare = InStr(cboTimeAfter.Text, "min")
    If vCompare <> 0 Then
         IntervalType = "n"
        GoTo startWrite
    End If
    
    'hours
    vCompare = InStr(cboTimeAfter.Text, "Hour")
    If vCompare <> 0 Then
         IntervalType = "h"
        GoTo startWrite
    End If
    
    'days
    vCompare = InStr(cboTimeAfter.Text, "day")
    If vCompare <> 0 Then
         IntervalType = "d"
    End If
    
startWrite:
    sTime = DateAdd(IntervalType, (Val(cboTimeAfter.Text)), sTime)
    With rsAlarm
        .AddNew
        .Fields("AlarmDate") = sDate
        .Fields("AlarmTimeSet") = sTime
        .Fields("RealAlarmDate") = Format(CDate(Text1.Text), "dd.mm.yyyy")
        .Fields("RealAlarmTime") = Format(CDate(Text2.Text), "hh:mm")
        .Fields("AlarmNote") = Text3.Text
        .Update
    End With
    Unload Me
End Sub

Private Sub btnSleep_Click()
    cboTimeAfter.Visible = True
    btnOkSleep.Visible = True
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ShowText
    LoadcboTimeAfter
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsAlarm = dbKids.OpenRecordset("Alarm")
    Set rsLanguage = dbKidLang.OpenRecordset("frmNotice")
    MakeWindowAlwaysTop (Me.hWnd)
   Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmNotice: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    rsAlarm.Close
    MakeWindowNotTop (Me.hWnd)
    Set frmNotice = Nothing
End Sub

Private Sub Timer1_Timer()
Dim RetVal As Long
    RetVal = FlashWindow(Me.hWnd, 1)
End Sub
