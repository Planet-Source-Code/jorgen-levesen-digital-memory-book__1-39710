VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Help Text"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   ControlBox      =   0   'False
   Icon            =   "frmEditHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   7680
      TabIndex        =   28
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   9
      Left            =   8520
      Picture         =   "frmEditHelp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Word spelling"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   9
      Left            =   3960
      Picture         =   "frmEditHelp.frx":060C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Word spelling"
      Top             =   0
      Width           =   495
   End
   Begin VB.Data rsLanguage2 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsLanguage 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cmbFonts 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   1
      Left            =   4680
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   120
      Width           =   3090
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   7890
      TabIndex        =   24
      Text            =   "12"
      Top             =   120
      Width           =   540
   End
   Begin VB.ComboBox cmbFonts 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   0
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   120
      Width           =   3090
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   0
      Left            =   3330
      TabIndex        =   22
      Text            =   "12"
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   8
      Left            =   8520
      Picture         =   "frmEditHelp.frx":07D6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Delete text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   7
      Left            =   8040
      Picture         =   "frmEditHelp.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Paste text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   6
      Left            =   7560
      Picture         =   "frmEditHelp.frx":0FE2
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Copy text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   5
      Left            =   7080
      Picture         =   "frmEditHelp.frx":16A4
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Right justified"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   4
      Left            =   6600
      Picture         =   "frmEditHelp.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Mid justified"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   3
      Left            =   6120
      Picture         =   "frmEditHelp.frx":1938
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Left justified"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   2
      Left            =   5640
      Picture         =   "frmEditHelp.frx":1A82
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Italian text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   1
      Left            =   5160
      Picture         =   "frmEditHelp.frx":1BCC
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Underlined text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Pic2 
      Height          =   495
      Index           =   0
      Left            =   4680
      Picture         =   "frmEditHelp.frx":1D16
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Bold text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmEditHelp.frx":1E60
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Bold text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   1
      Left            =   600
      Picture         =   "frmEditHelp.frx":1FAA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Underlined text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   2
      Left            =   1080
      Picture         =   "frmEditHelp.frx":20F4
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Italian text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   3
      Left            =   1560
      Picture         =   "frmEditHelp.frx":223E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Left justified"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   4
      Left            =   2040
      Picture         =   "frmEditHelp.frx":2388
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Mid justified"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   5
      Left            =   2520
      Picture         =   "frmEditHelp.frx":24D2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Right justified"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   6
      Left            =   3000
      Picture         =   "frmEditHelp.frx":261C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Copy text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   7
      Left            =   3480
      Picture         =   "frmEditHelp.frx":2CDE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Paste text"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnRichText 
      Height          =   495
      Index           =   8
      Left            =   3960
      Picture         =   "frmEditHelp.frx":33A0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Delete text"
      Top             =   480
      Width           =   495
   End
   Begin RichTextLib.RichTextBox Text1 
      DataField       =   "Help"
      DataSource      =   "rsLanguage"
      Height          =   6255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   16777152
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmEditHelp.frx":34EA
   End
   Begin RichTextLib.RichTextBox Text1 
      DataField       =   "Help"
      DataSource      =   "rsLanguage2"
      Height          =   6255
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   16777152
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmEditHelp.frx":35A4
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Language Help Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "English Help Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmEditHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As Control, vBookmarkLang0() As Variant, vBookmarkLang1() As Variant
Dim retText As String
Dim wdApp As Word.Application
Dim dbTemp As Database
Private Sub ShowPictures1()
        On Error Resume Next
        If Text1(0).SelBold = True Then
            iSelBold = True
        Else
            iSelBold = False
        End If
        If Text1(0).SelUnderline = True Then
            iSelUlin = True
        Else
            iSelUlin = False
        End If
        If Text1(0).SelItalic = True Then
            iSelItal = True
        Else
            iSelItal = False
        End If
        If Text1(0).SelAlignment = 0 Then
            iSelLeft = True
        Else
            iSelLeft = False
        End If
        If Text1(0).SelAlignment = 2 Then
            iSelMid = True
        Else
            iSelMid = False
        End If
        If Text1(0).SelAlignment = 1 Then
            iSelRight = True
        Else
            iSelRight = True
        End If
End Sub

Private Sub ShowPictures2()
        On Error Resume Next
        If Text1(1).SelBold = True Then
            iSelBold = True
        Else
            iSelBold = False
        End If
        If Text1(1).SelUnderline = True Then
            iSelUlin = True
        Else
            iSelUlin = False
        End If
        If Text1(1).SelItalic = True Then
            iSelItal = True
        Else
            iSelItal = False
        End If
        If Text1(1).SelAlignment = 0 Then
            iSelLeft = True
        Else
            iSelLeft = False
        End If
        If Text1(1).SelAlignment = 2 Then
            iSelMid = True
        Else
            iSelMid = False
        End If
        If Text1(1).SelAlignment = 1 Then
            iSelRight = True
        Else
            iSelRight = True
        End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnRichText_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0  'font bold
            If iSelBold = False Then
                Text1(0).SelBold = True
                iSelBold = True
            ElseIf iSelBold = True Then
                Text1(0).SelBold = False
                iSelBold = False
            End If
        Case 1  'underlined text
            If iSelUlin = False Then
                Text1(0).SelUnderline = True
                iSelUlin = True
            ElseIf iSelUlin = True Then
                Text1(0).SelUnderline = False
                iSelUlin = False
            End If
    Case 2  'italic text
            If iSelItal = False Then
                Text1(0).SelItalic = True
                iSelItal = True
            ElseIf iSelItal = True Then
                Text1(0).SelItalic = False
                iSelItal = False
            End If
    Case 3  'left justified text
                If iSelLeft = False Then
                    Text1(0).SelAlignment = 0
                    iSelLeft = True
                ElseIf iSelLeft = True Then
                    Text1(0).SelAlignment = 0
                    iSelLeft = False
                End If
    Case 4  'mid justified text
                If iSelMid = False Then
                    Text1(0).SelAlignment = 2
                    iSelMid = True
                ElseIf iSelMid = True Then
                    Text1(0).SelAlignment = 0
                    iSelMid = True
                End If
    Case 5  'Right justified text
                If iSelRight = False Then
                    Text1(0).SelAlignment = 1
                    iSelRight = True
                ElseIf iSelRight = True Then
                    Text1(0).SelAlignment = 0
                    iSelRight = False
                End If
    Case 6  'Copy text to clipboard
                Clipboard.Clear
                Clipboard.SetText Text1(0).SelText
    Case 7  'Paste text/picture from clipboard
                If Clipboard.GetFormat(1) Then
                    Text1(0).SelText = Clipboard.GetText()
                ElseIf Clipboard.GetFormat(2) Then
                    Text1(0).SelText = Clipboard.GetData()
                End If
    Case 8  'delete text
                Text1(0).SelText = ""
    Case 9  'spell check
                Call CheckSpelling(frmEditHelp.Text1(0))
                Me.Show
    Case Else
    End Select
End Sub


Private Sub cmbFonts_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        cmbFonts(0).FontName = cmbFonts(0).Text
        With Text1(0)
            .SelFontName = cmbFonts(0).Text
            .SelFontSize = CInt(Text17(0).Text)
        End With
    Case 1
        cmbFonts(1).FontName = cmbFonts(1).Text
        With Text1(1)
            .SelFontName = cmbFonts(1).Text
            .SelFontSize = CInt(Text17(1).Text)
        End With
    Case Else
    End Select
End Sub

Private Sub Form_Activate()
    'On Error Resume Next
    rsLanguage.Refresh
    rsLanguage2.Refresh
    
    With rsLanguage.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = "ENG" Then Exit Do
        .MoveNext
        Loop
    End With
    
    With rsLanguage2.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then Exit Do
        .MoveNext
        Loop
    End With
    
    cmbFonts(0).Clear
    cmbFonts(1).Clear
    For n = 0 To Screen.FontCount - 1
        cmbFonts(0).AddItem Screen.Fonts(n)
        cmbFonts(1).AddItem Screen.Fonts(n)
    Next
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rsLanguage.DatabaseName = dbKidLangTxt
    rsLanguage2.DatabaseName = dbKidLangTxt
    rsLanguage.RecordSource = Trim(CStr(frmHelp.Label1.Caption))
    rsLanguage2.RecordSource = Trim(CStr(frmHelp.Label1.Caption))
    Unload frmHelp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Recordset.Move 0
    rsLanguage.Recordset.Close
    rsLanguage2.Recordset.Move 0
    rsLanguage2.Recordset.Close
    dbTemp.Close
    Set frmEditHelp = Nothing
End Sub


Private Sub Pic2_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0  'font bold
            If iSelBold = False Then
                Text1(1).SelBold = True
                iSelBold = True
            ElseIf iSelBold = True Then
                Text1(1).SelBold = False
                iSelBold = False
            End If
        Case 1  'underlined text
            If iSelUlin = False Then
                Text1(1).SelUnderline = True
                iSelUlin = True
            ElseIf iSelUlin = True Then
                Text1(1).SelUnderline = False
                iSelUlin = False
            End If
    Case 2  'italic text
            If iSelItal = False Then
                Text1(1).SelItalic = True
                iSelItal = True
            ElseIf iSelItal = True Then
                Text1(1).SelItalic = False
                iSelItal = False
            End If
    Case 3  'left justified text
                If iSelLeft = False Then
                    Text1(1).SelAlignment = 0
                    iSelLeft = True
                ElseIf iSelLeft = True Then
                    Text1(1).SelAlignment = 0
                    iSelLeft = False
                End If
    Case 4  'mid justified text
                If iSelMid = False Then
                    Text1(1).SelAlignment = 2
                    iSelMid = True
                ElseIf iSelMid = True Then
                    Text1(1).SelAlignment = 0
                    iSelMid = True
                End If
    Case 5  'Right justified text
                If iSelRight = False Then
                    Text1(1).SelAlignment = 1
                    iSelRight = True
                ElseIf iSelRight = True Then
                    Text1(1).SelAlignment = 0
                    iSelRight = False
                End If
    Case 6  'Copy text to clipboard
                Clipboard.Clear
                Clipboard.SetText Text1(1).SelText
    Case 7  'Paste text/picture from clipboard
                If Clipboard.GetFormat(1) Then
                    Text1(1).SelText = Clipboard.GetText()
                ElseIf Clipboard.GetFormat(2) Then
                    Text1(1).SelText = Clipboard.GetData()
                End If
    Case 8  'delete text
                Text1(1).SelText = ""
    Case 9  'spell check
                Call CheckSpelling(frmEditHelp.Text1(1))
                Me.Show
    Case Else
    End Select
End Sub


Private Sub Text1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        Text17(0).Text = Text1(0).SelFontSize
        Call ShowPictures1
    Case 1
        Text17(1).Text = Text1(1).SelFontSize
        Call ShowPictures2
    Case Else
    End Select
End Sub


Private Sub Text1_GotFocus(Index As Integer)
' Ignore errors for controls without the TabStop property.
    On Error Resume Next
    ' Switch off the change of focus when pressing TAB.
    For Each Control In Controls
        Control.TabStop = False
    Next Control
End Sub


Private Sub Text1_LostFocus(Index As Integer)
' Ignore errors for controls without the TabStop property.
    On Error Resume Next
    ' Turn on the change of focus when pressing TAB.
    For Each Control In Controls
        Control.TabStop = True
    Next Control
End Sub


Private Sub Text1_SelChange(Index As Integer)
        On Error Resume Next
        Select Case Index
        Case 0
            If Text1(0).SelBold = True Then
                btnRichText(0).Picture = LoadResPicture(42, vbResBitmap)
                iSelBold = True
            ElseIf Text1(0).SelUnderline = True Then
                btnRichText(1).Picture = LoadResPicture(44, vbResBitmap)
                iSelUlin = True
            ElseIf Text1(0).SelItalic = True Then
                btnRichText(2).Picture = LoadResPicture(47, vbResBitmap)
                iSelItal = True
            ElseIf Text1(0).SelAlignment = 0 Then
                btnRichText(3).Picture = LoadResPicture(50, vbResBitmap)
                iSelLeft = True
            ElseIf Text1(0).SelAlignment = 2 Then
                btnRichText(4).Picture = LoadResPicture(52, vbResBitmap)
                iSelMid = True
            ElseIf Text1(0).SelAlignment = 1 Then
                btnRichText(5).Picture = LoadResPicture(54, vbResBitmap)
                iSelRight = True
            End If
            cmbFonts(0).FontName = Text1(0).Font.Name
            Text17(0).Text = Text1(0).SelFontSize
        Case 1
            If Text1(1).SelBold = True Then
                Pic2(0).Picture = LoadResPicture(42, vbResBitmap)
                iSelBold = True
            ElseIf Text1(1).SelUnderline = True Then
                Pic2(1).Picture = LoadResPicture(44, vbResBitmap)
                iSelUlin = True
            ElseIf Text1(1).SelItalic = True Then
                Pic2(2).Picture = LoadResPicture(47, vbResBitmap)
                iSelItal = True
            ElseIf Text1(1).SelAlignment = 0 Then
                Pic2(3).Picture = LoadResPicture(50, vbResBitmap)
                iSelLeft = True
            ElseIf Text1(1).SelAlignment = 2 Then
                Pic2(4).Picture = LoadResPicture(52, vbResBitmap)
                iSelMid = True
            ElseIf Text1(1).SelAlignment = 1 Then
                Pic2(5).Picture = LoadResPicture(54, vbResBitmap)
                iSelRight = True
            End If
            cmbFonts(1).FontName = Text1(1).Font.Name
            Text17(1).Text = Text1(1).SelFontSize
        Case Else
        End Select
End Sub
Private Sub Text17_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        With Text1(0)
            .SelFontName = cmbFonts(0).Text
            .SelFontSize = CInt(Text17(0).Text)
        End With
    Case 1
        With Text1(1)
            .SelFontName = cmbFonts(1).Text
            .SelFontSize = CInt(Text17(1).Text)
        End With
    Case Else
    End Select
End Sub
