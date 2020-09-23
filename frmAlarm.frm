VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmAlarm 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Alarm Event"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Call time before"
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   2895
      Begin VB.ComboBox cboTimeBefore 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Alarm Event"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   2895
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Alarm Date And Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox DateAlarm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   840
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "txtMin"
         BuddyDispid     =   196617
         OrigLeft        =   2040
         OrigTop         =   1200
         OrigRight       =   2280
         OrigBottom      =   1575
         Max             =   60
         Min             =   5
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtHour"
         BuddyDispid     =   196618
         OrigLeft        =   1080
         OrigTop         =   1200
         OrigRight       =   1320
         OrigBottom      =   1575
         Max             =   23
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   1560
         TabIndex        =   2
         Text            =   "00"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtHour 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   600
         TabIndex        =   1
         Text            =   "10"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Minute"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAlarm As Recordset
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
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("Frame2")) Then
                    .Fields("Frame2") = Frame2.Caption
                Else
                    Frame2.Caption = .Fields("Frame2")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
                If IsNull(.Fields("btnOK")) Then
                    .Fields("btnOK") = btnOk.Caption
                Else
                    btnOk.Caption = .Fields("btnOK")
                End If
                If IsNull(.Fields("btnCancel")) Then
                    .Fields("btnCancel") = btnCancel.Caption
                Else
                    btnCancel.Caption = .Fields("btnCancel")
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
        .Fields("Frame1") = Frame1.Caption
        .Fields("Frame2") = Frame2.Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("btnOK") = btnOk.Caption
        .Fields("btnCancel") = btnCancel.Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub LoadcboTimeBefore()
    With cboTimeBefore
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
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
Dim sTime As String, sDate As String, vCompare As Variant
Dim IntervalType As String

    On Error Resume Next
    sTime = CInt(txtHour.Text) & ":" & CInt(txtMin.Text)
    sDate = Format(CDate(DateAlarm.Text), "dd.mm.yyyy")
    
    If Len(cboTimeBefore.Text) = 0 Then GoTo startWrite
    
    vCompare = InStr(cboTimeBefore.Text, "min")
    If vCompare <> 0 Then
         IntervalType = "n"
        sTime = DateAdd(IntervalType, (Val(cboTimeBefore.Text) * -1), sTime)
        GoTo startWrite
    End If
    vCompare = InStr(cboTimeBefore.Text, "Hour")
    If vCompare <> 0 Then
         IntervalType = "h"
        sTime = DateAdd(IntervalType, (Val(cboTimeBefore.Text) * -1), sTime)
        GoTo startWrite
    End If
    vCompare = InStr(cboTimeBefore.Text, "day")
    If vCompare <> 0 Then
         IntervalType = "d"
        sDate = DateAdd(IntervalType, (Val(cboTimeBefore.Text) * -1), sDate)
    End If
    
startWrite:
    With rsAlarm
        .AddNew
        .Fields("AlarmDate") = sDate
        .Fields("AlarmTimeSet") = sTime
        .Fields("RealAlarmDate") = Format(CDate(DateAlarm.Text), "dd.mm.yyyy")
        .Fields("RealAlarmTime") = txtHour.Text & ":" & txtMin.Text
        If Len(Text1.Text) <> 0 Then
            .Fields("AlarmNote") = Text1.Text
        End If
        .Update
    End With
    Unload Me
End Sub

Private Sub DateAlarm_Click()
Dim UserDate As Date
    UserDate = CVDate(DateAlarm.Text)
    If Not IsDate(DateAlarm.Text) Then Exit Sub
    If frmCalendar.GetDate(UserDate) Then
        DateAlarm.Text = UserDate
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    LoadcboTimeBefore
    ShowText
    DateAlarm.Text = Format(Now, "d.mm.yyyy")
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsAlarm = dbKids.OpenRecordset("Alarm")
    Set rsLanguage = dbKidLang.OpenRecordset("frmAlarm")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmAlarm: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsAlarm.Close
    rsLanguage.Close
    HideAllButtons
    Set frmAlarm = Nothing
End Sub
Private Sub UpDown1_DownClick()
On Error Resume Next
    If CInt(txtHour.Text) = 1 Then Exit Sub
    txtHour.Text = CInt(txtHour.Text) - 1
End Sub

Private Sub UpDown1_UpClick()
    On Error Resume Next
    If CInt(txtHour.Text) = 24 Then Exit Sub
    txtHour.Text = CInt(txtHour.Text) + 1
End Sub
Private Sub UpDown2_DownClick()
    If CInt(txtMin.Text) = 5 Then Exit Sub
    txtMin.Text = CInt(txtMin.Text) - 5
End Sub

Private Sub UpDown2_UpClick()
    On Error Resume Next
    If CInt(txtMin.Text) = 60 Then Exit Sub
    txtMin.Text = CInt(txtMin.Text) + 5
End Sub
