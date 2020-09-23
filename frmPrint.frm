VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Memory Book"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   ControlBox      =   0   'False
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   9255
      Begin VB.Frame Frame4 
         Caption         =   "Use section picture:"
         Height          =   2295
         Left            =   4680
         TabIndex        =   15
         Top             =   240
         Width           =   4455
         Begin VB.Data rsClipArt2 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKidPic.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   480
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "ClipArt"
            Top             =   720
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.CommandButton btnPrevious 
            Caption         =   "&Previous Picture"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   17
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton btnNext 
            Caption         =   "&Next Picture"
            Height          =   375
            Index           =   1
            Left            =   2400
            TabIndex        =   16
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            DataField       =   "FrameName"
            DataSource      =   "rsClipArt2"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   19
            Top             =   360
            Width           =   1815
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "FramePicture"
            DataSource      =   "rsClipArt2"
            Height          =   1815
            Index           =   1
            Left            =   120
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Use frontpage picture:"
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4455
         Begin VB.CommandButton btnNext 
            Caption         =   "&Next Picture"
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   14
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnPrevious 
            Caption         =   "&Previous Picture"
            Height          =   375
            Index           =   0
            Left            =   2520
            TabIndex        =   13
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Data rsClipArt 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKidPic.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   480
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "ClipArt"
            Top             =   480
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            DataField       =   "FrameName"
            DataSource      =   "rsClipArt"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   18
            Top             =   360
            Width           =   1815
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            DataField       =   "FramePicture"
            DataSource      =   "rsClipArt"
            Height          =   1815
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton Option1 
         Caption         =   "Print directly using default printer"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   3975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Print using MS Word"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   3975
      End
   End
   Begin VB.CommandButton btnPrint 
      Height          =   375
      Left            =   120
      Picture         =   "frmPrint.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Start printing..."
      Top             =   4920
      Width           =   9255
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox Check5 
         Caption         =   "Print Childhood"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Print Infancy"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Print Babtism"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Print Birth"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Print Pregnancy"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCounter As Integer, lTabPos As Integer
Dim bFirstWrite As Boolean
Dim rsMyRecord As Recordset
Dim rsLanguage As Recordset
Dim rsLanguage2 As Recordset
Private Sub DoNewPagePreview()
    On Error Resume Next
    cPrint.pFooter
    cPrint.pNewPage
    LineCounter = 0
    cPrint.FontBold = True
    cPrint.pCenter MDIMasterKid.cmbChildren.Text
    cPrint.FontBold = False
    cPrint.pBox 0.5, 0.5, cPrint.GetPaperWidth - 1, cPrint.GetPaperHeight - 1, , , vbFSTransparent
    cPrint.FontName = "Ariel"
    cPrint.FontSize = 10
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint
End Sub
Private Sub ShowPictures()
    On Error Resume Next
    If CLng(rsMyRecord.Fields("FrontPicID")) <> 0 Then
        With rsClipArt.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("LineNo")) = CLng(rsMyRecord.Fields("FrontPicID")) Then Exit Do
            .MoveNext
            Loop
        End With
    End If
    If CLng(rsMyRecord.Fields("SectionPicID")) <> 0 Then
        With rsClipArt2.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("LineNo")) = CLng(rsMyRecord.Fields("SectionPicID")) Then Exit Do
            .MoveNext
            Loop
        End With
    End If
End Sub
Private Sub WriteWhenIWasBorn()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmWhenIWasBorn")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame2")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(0)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(1).Text & "  " & frmWhenIWasBorn.Label3(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(1)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(2).Text & "  " & frmWhenIWasBorn.Label3(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(2)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(3).Text & "  " & frmWhenIWasBorn.Label3(2).Caption & " / " & frmWhenIWasBorn.cmbDim.Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(3)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(4).Text & "  " & frmWhenIWasBorn.Label3(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame3")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(4)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(5).Text & "  " & rsLanguage.Fields("label4") & "  " & frmWhenIWasBorn.Label3(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(5)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(6).Text & "  " & rsLanguage.Fields("label4") & "  " & frmWhenIWasBorn.Label3(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(6)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(6).Text & "  " & rsLanguage.Fields("label4") & "  " & frmWhenIWasBorn.Label3(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(7)") & "  " & frmWhenIWasBorn.cmbCurrency(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(8).Text & "  " & rsLanguage.Fields("label4") & "  " & frmWhenIWasBorn.Label3(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(8)") & "  " & frmWhenIWasBorn.cmbCurrency(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(9).Text & "  " & rsLanguage.Fields("label4") & "  " & frmWhenIWasBorn.Label3(8).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2(9)") & "  " & frmWhenIWasBorn.cmbCurrency(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(10).Text & "  " & rsLanguage.Fields("label4") & "  " & frmWhenIWasBorn.Label3(9).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame5")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmWhenIWasBorn.Text1(11).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame4")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData frmWhenIWasBorn.Picture1.Picture, vbCFBitmap
        .Selection.Paste
        .Selection.MoveDown Unit:=wdLine, Count:=1
    End With
    rsLanguage.Close
End Sub

Private Sub PrintWhenIWasBorn()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmWhenIWasBorn")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmWhenIWasBorn
        cPrint.pPrint rsLanguage.Fields("Frame1"), 1, True
        cPrint.pMultiline .Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        cPrint.pPrint
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint rsLanguage.Fields("Frame2"), 1
        cPrint.pPrint
        cPrint.pPrint .Label2(0).Caption, 1, True
        If Len(.Text1(1).Text) Then
            cPrint.pPrint .Text1(1).Text & "  " & .Label3(0).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint rsLanguage.Fields("label2(1)"), 1, True
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pPrint .Text1(2).Text & "  " & .Label3(1).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label2(2)"), 1, True
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pPrint .Text1(2).Text & "  " & .Label3(2).Caption & "  " & .cmbDim.Text, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label2(3)"), 1, True
        If Len(.Text1(3).Text) <> 0 Then
            cPrint.pPrint .Text1(3).Text & "  " & .Label3(3).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame3"), 1    'exchange rates
        cPrint.FontBold = False
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label2(4)"), 1, True
        If Len(.Text1(5).Text) <> 0 Then
            cPrint.pPrint .Text1(5).Text & "  " & rsLanguage.Fields("label4") & "  " & .Label3(4).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint rsLanguage.Fields("label2(5)"), 1, True
        If Len(.Text1(6).Text) <> 0 Then
            cPrint.pPrint .Text1(6).Text & "  " & rsLanguage.Fields("label4") & "  " & .Label3(5).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label2(6)"), 1, True
        If Len(.Text1(7).Text) <> 0 Then
            cPrint.pPrint .Text1(7).Text & "  " & rsLanguage.Fields("label4") & "  " & .Label3(6).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label2(7)") & " " & .cmbCurrency(0).Text, 1, True
        If Len(.Text1(8).Text) <> 0 Then
            cPrint.pPrint .Text1(8).Text & "  " & rsLanguage.Fields("label4") & "  " & .Label3(7).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label2(7)") & " " & .cmbCurrency(1).Text, 1, True
        If Len(.Text1(9).Text) <> 0 Then
            cPrint.pPrint .Text1(9).Text & "  " & rsLanguage.Fields("label4") & "  " & .Label3(8).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label2(7)") & " " & .cmbCurrency(2).Text, 1, True
        If Len(.Text1(10).Text) <> 0 Then
            cPrint.pPrint .Text1(10).Text & "  " & rsLanguage.Fields("label4") & "  " & .Label3(9).Caption, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame5"), 1    'how was the fashion
        cPrint.FontBold = False
        cPrint.pPrint
        If Len(.Text1(11).Text) <> 0 Then
            cPrint.pMultiline .Text1(11).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame4"), 1    'fashion picture
        If Not IsNull(.rsWhenBorn.Recordset.Fields("FashionPic")) Then
            cPrint.pPrintPicture .Picture1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
        End If
    End With
    rsLanguage.Close
End Sub

Private Sub PrepareWord()
    On Error Resume Next
    Set wdApp = New Word.Application
    With wdApp
        .Application.Visible = True
        .Application.WindowState = wdWindowStateMaximize
        .Documents.Add DocumentType:=wdNewBlankDocument
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
        .Selection.TypeText Text:=MDIMasterKid.cmbChildren.Text & "  -  " & Format(CDate(Now), "dd.mm.yyyy")
        .Selection.Sections(1).Headers(1).PageNumbers.Add PageNumberAlignment:= _
            wdAlignPageNumberRight, FirstPage:=False
        .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
        'init styles
        .ActiveDocument.Styles("Heading 1").Font.Name = "Monotype Corsiva"
        .ActiveDocument.Styles("Heading 1").Font.Size = 48
        .ActiveDocument.Styles("Heading 1").Font.Bold = True
        .ActiveDocument.Styles("Heading 1").Font.Italic = True
        .ActiveDocument.Styles("Heading 1").Font.Shadow = True
        .ActiveDocument.Styles("Heading 1").Font.Color = wdColorRed
               
        .ActiveDocument.Styles("Heading 2").Font.Name = "Arial"
        .ActiveDocument.Styles("Heading 2").Font.Size = 18
        .ActiveDocument.Styles("Heading 2").Font.Bold = True
        .ActiveDocument.Styles("Heading 2").Font.Italic = False
        .ActiveDocument.Styles("Heading 2").Font.Color = wdColorBlack
        .ActiveDocument.Styles("Heading 2").Font.Shadow = False
        
        .ActiveDocument.Styles("Normal").Font.Name = "Times New Roman"
        .ActiveDocument.Styles("Normal").Font.Size = 10
        .ActiveDocument.Styles("Normal").Font.Bold = False
        .ActiveDocument.Styles("Normal").Font.Shadow = False
        .ActiveDocument.Styles("Normal").Font.Italic = False
        .ActiveDocument.Styles("Normal").Font.Color = wdColorBlack
  End With
  
    'write 20 empty lines before heading
    For n = 0 To 20
        wdApp.Selection.TypeParagraph
    Next
    
    With wdApp.Selection.Font
        .Name = "Monotype Corsiva"
        .Size = 48
        .Bold = True
        .Italic = True
        .Underline = wdUnderlineNone
        .UnderlineColor = wdColorAutomatic
        .Strikethrough = False
        .DoubleStrikeThrough = False
        .Outline = False
        .Emboss = False
        .Shadow = True
        .Hidden = False
        .SmallCaps = False
        .AllCaps = False
        .Color = wdColorAutomatic
        .Engrave = False
        .Superscript = False
        .Subscript = False
        .Spacing = 0
        .Scaling = 100
        .Position = 0
        .Kerning = 0
        .Animation = wdAnimationNone
    End With
    With wdApp
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeText Text:=rsLanguage2.Fields("PrintFrontPage")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=gsChildName
        .Selection.TypeParagraph
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.Font.Name = "Times New Roman"
        .Selection.Font.Size = 12
        .Selection.Font.Bold = wdToggle
        .Selection.Font.Italic = wdToggle
        'make a new page and prepare for Index
        .Selection.InsertBreak Type:=wdPageBreak
         .ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:="Index"
        .Selection.TypeParagraph
        .Selection.TypeParagraph
    End With
End Sub


Private Sub PrintBooksChild()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBooks")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
        
    With frmBooks.rsBooksChild.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                DoEvents
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
                If Len(frmBooks.Text1(5).Text) <> 0 Then
                    cPrint.pPrint frmBooks.Text1(5).Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
                If Len(frmBooks.Text1(7).Text) <> 0 Then
                    cPrint.pPrint frmBooks.Text1(7).Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
                If Len(frmBooks.RichTextBox1(1).Text) <> 0 Then
                    cPrint.pMultiline frmBooks.RichTextBox1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint rsLanguage.Fields("label1(3)"), 1
                If Len(frmBooks.Text1(8).Text) <> 0 Then
                    cPrint.pPrint frmBooks.Text1(8).Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(4)"), 1, True
                If IsDate(.Recordset.Fields("PurchaseDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(5)"), 1, True
                cPrint.pPrint Format(CDbl(frmBooks.Text1(9).Text), "0.00") & "  " & frmBooks.Label4(0).Caption, 3.5
                cPrint.pPrint rsLanguage.Fields("label1(6)"), 1, True
                If Len(frmBooks.Text1(6).Text) <> 0 Then
                    cPrint.pMultiline frmBooks.Text1(6).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub


'Print header
Public Sub PrintHeaderComplete()
    On Error Resume Next
    cPrint.pFooter
    cPrint.pNewPage
    LineCounter = 0
    cPrint.pBox 0.5, 0.5, cPrint.GetPaperWidth - 1, cPrint.GetPaperHeight - 1, , , vbFSTransparent
    cPrint.FontBold = True
    cPrint.pCenter MDIMasterKid.cmbChildren.Text
    cPrint.FontBold = False
    cPrint.FontName = "Ariel"
    cPrint.FontSize = 16
    cPrint.FontItalic = True
    cPrint.FontBold = True
    cPrint.pCenter sHeader
    cPrint.FontBold = False
    cPrint.FontItalic = False
    cPrint.FontSize = 10
    cPrint.pPrint
    cPrint.pPrint
End Sub

Private Sub PrintBlockName(sString As String)
    On Error Resume Next
    cPrint.pFooter
    cPrint.pNewPage
    cPrint.pPrintPicture Image1(1).Picture, , , , , , False
    cPrint.FontBold = True
    cPrint.pCenter MDIMasterKid.cmbChildren.Text
    cPrint.FontBold = False
    
    'print 15 empty lines
    For n = 1 To 15
        cPrint.pPrint
    Next
    
    cPrint.FontName = "Ariel"
    cPrint.FontSize = 28
    cPrint.FontItalic = True
    cPrint.FontBold = True
    cPrint.ForeColor = Red
    cPrint.pCenter sString, 0.3
    cPrint.FontBold = False
    cPrint.FontItalic = False
    cPrint.ForeColor = BLACK
    cPrint.FontSize = 10
    DoEvents
End Sub


Private Sub PrintFrontPage()
    On Error GoTo errPrintFrontPage
    'print the front page heading
    cPrint.FontName = "Ariel"
    cPrint.FontSize = 48
    cPrint.FontItalic = True
    cPrint.FontBold = True
    cPrint.pPrintPicture Image1(0).Picture, , , , , , False
    Unload frmFrames
    cPrint.CurrentY = (cPrint.GetPaperHeight / 2) - 2
    cPrint.pCenter rsLanguage2.Fields("PrintFrontPage")
    cPrint.pCenter gsChildName
    cPrint.FontBold = False
    cPrint.FontItalic = False
    cPrint.FontSize = 10
    Exit Sub
    
errPrintFrontPage:
    Beep
    MsgBox Err.Description, vbInformation, "Frontpage picture is missing"
    Resume Next
End Sub
Private Sub PrintHealthChild()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmHealth")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName4")
    PrintHeaderComplete
    
    'health childhood
    With frmHealth.rsHealthControlChild.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                DoEvents
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                cPrint.pPrint Format(CDate(.Fields("ControlDate")), "dd.mm.yyyy"), 3.5
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                If Len(frmHealth.RichTextBox1(3).Text) <> 0 Then
                    cPrint.pMultiline frmHealth.RichTextBox1(3).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint "", 3.5
                End If
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
    
    sHeader = rsLanguage.Fields("FormName5")
    PrintHeaderComplete
    
    'vaccinations childhood
    With frmHealth.rsVaccinationChild.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                cPrint.pPrint Format(CDate(.Fields("VaccinationDate")), "dd.mm.yyyy"), 3.5
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label3"), 1, True
                If Not IsNull(.Fields("VaccinationByDoctor")) Then
                    cPrint.pPrint .Fields("VaccinationByDoctor"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label4"), 1, True
                If Not IsNull(.Fields("VaccinationWhere")) Then
                    cPrint.pPrint Format(.Fields("VaccinationWhere")), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label2"), 1, True
                If Len(frmHealth.RichTextBox1(4).Text) <> 0 Then
                    cPrint.pMultiline frmHealth.RichTextBox1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
    
    sHeader = rsLanguage.Fields("FormName6")
    PrintHeaderComplete
    
    'illness childhood
    With frmHealth.rsIllnessChild.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                If IsDate(.Fields("IllnessDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("IllnessDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label5"), 1, True
                If Len(frmHealth.RichTextBox1(5).Text) <> 0 Then
                    cPrint.pMultiline frmHealth.RichTextBox1(5).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                If CBool(.Fields("IllnessDoctor")) = 1 Then
                    cPrint.pPrint rsLanguage.Fields("Check1"), 1, True
                    cPrint.pPrint rsLanguage.Fields("Yes"), 3.5
                    cPrint.pPrint rsLanguage.Fields("label6"), 1, True
                    cPrint.pPrint Format(.Fields("IllnessDoctorName")), 3.5
                Else
                    cPrint.pPrint rsLanguage.Fields("Check1"), 1, True
                    cPrint.pPrint rsLanguage.Fields("No"), 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
End Sub

Private Sub PrintPicturesChild()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPictures")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName3")
    PrintHeaderComplete
    bFirstWrite = True
    
    With frmPictures.rsChildPictures.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                'print picture notes
                If Not bFirstWrite Then
                    DoNewPagePreview
                End If
                cPrint.CurrentX = LeftMargin
                cPrint.pPrint rsLanguage.Fields("label2"), 1, True
                If Len(frmPictures.Text2(2).Text) <> 0 Then
                    cPrint.pMultiline frmPictures.Text2(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint
                If Not IsNull(.Fields("Picture")) Then
                    cPrint.pPrintPicture frmPictures.Picture1(2).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                End If
                bFirstWrite = False
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub PrintPicturesInfant()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPictures")
    ReadLanguage
        
    sHeader = rsLanguage.Fields("FormName2")
    PrintHeaderComplete
    bFirstWrite = True
    
    With frmPictures.rsInfancyPictures.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                'print picture notes
                If Not bFirstWrite Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label2"), 1, True
                If Len(frmPictures.Text2(1).Text) <> 0 Then
                    cPrint.pMultiline frmPictures.Text2(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint
                If Not IsNull(.Fields("Picture")) Then
                    cPrint.pPrintPicture frmPictures.Picture1(1).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                End If
                bFirstWrite = False
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub PrintToysChild()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmToys")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("Form")
    PrintHeaderComplete
        
    With frmToys.rsToysChild.Recordset
        .MoveFirst
        Do While Not .EOF
            cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
            If Not IsNull(.Fields("ToyName")) Then
                cPrint.pPrint .Fields("ToyName"), 3.5
            Else
                cPrint.pPrint " ", 3.5
            End If
            cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
            If IsDate(.Fields("PurchaseDate")) Then
                cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
            Else
                cPrint.pPrint " ", 1.5
            End If
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
            cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
            cPrint.pPrint Format(CDbl(.Fields("PurchasePrice")), "0.00") & "  " & frmToys.Label4(0).Caption, 3.5
            cPrint.pPrint
            cPrint.pPrint rsLanguage.Fields("label1(4)"), 1
            If Len(frmToys.RichTextBox1(1).Text) <> 0 Then
                cPrint.pMultiline frmToys.RichTextBox1(1).Text, 1, cPrint.GetPaperWidth - 1, , False, True
            Else
                cPrint.pPrint " ", 1.5
            End If
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
            cPrint.pPrint rsLanguage.Fields("label1(3)"), 1
            If Not IsNull(.Fields("Picture")) Then
                cPrint.pPrintPicture frmToys.Image1(1).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
            End If
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
            cPrint.pPrint
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Public Sub WriteChildNotesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesChild")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        frmFathersNotesChild.rsChildNotes.Recordset.MoveFirst
        Do While Not frmFathersNotesChild.rsChildNotes.Recordset.EOF
            If CLng(frmFathersNotesChild.rsChildNotes.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(frmFathersNotesChild.rsChildNotes.Recordset.Fields("NoteDate")) Then
                    DoEvents
                    .Selection.TypeText Text:=rsLanguage.Fields("Label1(0)")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=Format(CDate(frmFathersNotesChild.rsChildNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    Clipboard.Clear
                    Clipboard.SetText frmFathersNotesChild.RichText1.TextRTF, vbCFRTF
                    .Selection.Paste
                    DoEvents
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmFathersNotesChild.rsChildNotes.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Public Sub PrintChildNotes()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesChild")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmFathersNotesChild.rsChildNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(.Fields("NoteDate")) Then
                    cPrint.FontBold = True
                    cPrint.pPrint rsLanguage.Fields("label1") & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1
                    cPrint.FontBold = False
                    If Len(frmFathersNotesChild.RichText1.Text) <> 0 Then
                        cPrint.pMultiline frmFathersNotesChild.RichText1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    If cPrint.pEndOfPage Then
                        DoNewPagePreview
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Public Sub PrintBirthdays()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBirthDates")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    bFirstWrite = True
    
    With frmBirthDates
        .rsBirthDays.Recordset.MoveFirst
        Do While Not .rsBirthDays.Recordset.EOF
            If .rsBirthDays.Recordset.Fields("ChildNo") = glChildNo Then
                If Not bFirstWrite Then
                    DoNewPagePreview
                End If
                cPrint.FontBold = True
                cPrint.pPrint .Frame1.Caption, 1, True
                If IsDate(.rsBirthDays.Recordset.Fields("BirthDayDate")) Then
                    cPrint.pPrint Format(CDate(.rsBirthDays.Recordset.Fields("BirthDayDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.FontBold = False
                cPrint.pPrint
                cPrint.pPrint "Note:", 1, True
                If Len(.RichTextBox1.Text) <> 0 Then
                    cPrint.pMultiline .RichTextBox1.Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                
                'now print all the pictures from this birthday
                .rsBirthDayPictures.Recordset.MoveFirst
                Do While Not .rsBirthDayPictures.Recordset.EOF
                    If .rsBirthDayPictures.Recordset.Fields("ChildNo") = glChildNo And CDate(.rsBirthDayPictures.Recordset.Fields("BirthDayDate")) = CDate(.rsBirthDays.Recordset.Fields("Date")) Then
                        If Not bFirstWrite Then
                            DoNewPagePreview
                        End If
                        If Len(.Text2.Text) <> 0 Then
                            cPrint.pMultiline .Text2.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint " ", 1
                        End If
                        If Not IsNull(.rsBirthDayPictures.Recordset.Fields("Picture")) Then
                            cPrint.pPrintPicture .Picture1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                            bFirstWrite = False
                        Else
                            bFirstWrite = True
                        End If
                        cPrint.pPrint
                    End If
                .rsBirthDayPictures.Recordset.MoveNext
                Loop
                bFirstWrite = False
            End If
        .rsBirthDays.Recordset.MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Public Sub WriteBirthdaysWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBirthDates")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.TypeText Text:=frmBirthDates.Frame1.Caption
        .Selection.MoveRight Unit:=wdCell
        If IsDate(frmBirthDates.rsBirthDays.Recordset.Fields("BirthDayDate")) Then
            .Selection.TypeText Text:=Format(CDate(frmBirthDates.rsBirthDays.Recordset.Fields("BirthDayDate")), "dd.mm.yyyy")
        Else
            .Selection.TypeText Text:=" "
        End If
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:="Note:"
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmBirthDates.RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        .Selection.MoveRight Unit:=wdCell
    
        'now print all the pictures from this birthday
        frmBirthDates.rsBirthDayPictures.Recordset.MoveFirst
        Do While Not frmBirthDates.rsBirthDayPictures.Recordset.EOF
            If frmBirthDates.rsBirthDayPictures.Recordset.Fields("ChildNo") = glChildNo Then
                If CDate(frmBirthDates.rsBirthDayPictures.Recordset.Fields("BirthDayDate")) = CDate(frmBirthDates.List3.List(frmBirthDates.List3.ListIndex)) Then
                    DoEvents
                    .Selection.TypeText Text:=frmBirthDates.Text2.Text
                    .Selection.MoveRight Unit:=wdCell
                    frmBirthDates.Picture2.Picture = frmBirthDates.Picture1.Picture
                    Clipboard.Clear
                    Clipboard.SetData frmBirthDates.Picture2.Picture, vbCFBitmap
                    .Selection.Paste
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmBirthDates.rsBirthDayPictures.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub InitStyle()
    With wdApp
        .ActiveDocument.Styles("Heading 1").Font.Name = "Monotype Corsiva"
        .ActiveDocument.Styles("Heading 1").Font.Size = 48
        .ActiveDocument.Styles("Heading 1").Font.Bold = True
        .ActiveDocument.Styles("Heading 1").Font.Italic = True
        .ActiveDocument.Styles("Heading 1").Font.Shadow = True
        .ActiveDocument.Styles("Heading 1").Font.Color = wdColorRed
               
        .ActiveDocument.Styles("Heading 2").Font.Name = "Arial"
        .ActiveDocument.Styles("Heading 2").Font.Size = 18
        .ActiveDocument.Styles("Heading 2").Font.Bold = True
        .ActiveDocument.Styles("Heading 2").Font.Italic = False
        .ActiveDocument.Styles("Heading 2").Font.Color = wdColorBlack
        .ActiveDocument.Styles("Heading 2").Font.Shadow = False
        
        .ActiveDocument.Styles("Normal").Font.Name = "Times New Roman"
        .ActiveDocument.Styles("Normal").Font.Size = 10
        .ActiveDocument.Styles("Normal").Font.Bold = False
        .ActiveDocument.Styles("Normal").Font.Shadow = False
        .ActiveDocument.Styles("Normal").Font.Italic = False
        .ActiveDocument.Styles("Normal").Font.Color = wdColorBlack
    End With
End Sub
Private Sub MakeIndex()
    On Error Resume Next
    With wdApp
        .Selection.GoTo What:=wdGoToBookmark, Name:="Index"
        .ActiveWindow.Selection.Font.Size = 16
        .ActiveWindow.Selection.Font.Bold = True
        .Selection.TypeText Text:=rsLanguage2.Fields("PrintIndex")
        .ActiveWindow.Selection.Font.Bold = False
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.TypeParagraph
    End With
    With wdApp.ActiveDocument
        .TablesOfContents.Add Range:=wdApp.Selection.Range, RightAlignPageNumbers:= _
            True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
            LowerHeadingLevel:=3, IncludePageNumbers:=True, AddedStyles:="", _
            UseHyperlinks:=True, HidePageNumbersInWeb:=True
        .TablesOfContents(1).TabLeader = wdTabLeaderDots
        .TablesOfContents.Format = wdIndexIndent
    End With
End Sub

Private Sub WriteBooksChildWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBooks")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName2"))
        
    With wdApp
        frmBooks.rsBooksChild.Recordset.MoveFirst
        Do While Not frmBooks.rsBooksChild.Recordset.EOF
            If CLng(frmBooks.rsBooksChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksChild.Recordset.Fields("BookName"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksChild.Recordset.Fields("BookAuthor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmBooks.RichTextBox1(1).TextRTF, vbCFRTF
                .Selection.Paste
                DoEvents
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksChild.Recordset.Fields("BookIsbn"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
                .Selection.MoveRight Unit:=wdCell
                If IsDate(frmBooks.rsBooksChild.Recordset.Fields("PurchaseDate")) Then
                    .Selection.TypeText Text:=Format(CDate(frmBooks.rsBooksChild.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
                Else
                    .Selection.TypeText Text:=""
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(5)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksChild.Recordset.Fields("PurchasePrice"), "0.00") & "  " & frmBooks.Label4(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(6)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksChild.Recordset.Fields("WherePurchased"))
                .Selection.MoveRight Unit:=wdCell
            End If
        frmBooks.rsBooksChild.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub SaveMemoryBook()
    wdApp.ActiveDocument.SaveAs filename:=rsLanguage2.Fields("PrintFrontPage") & ".doc", FileFormat:= _
        wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:= _
        True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:= _
        True, SaveNativePictureFormat:=True, SaveFormsData:=False, _
        SaveAsAOCELetter:=False
End Sub

Private Sub WriteBlock(sString As String)
    On Error Resume Next
    With wdApp
        .Selection.InsertBreak Type:=wdPageBreak
        'write 10 empty lines
        For n = 0 To 10
            .Selection.TypeParagraph
        Next
        .Selection.Style = .ActiveDocument.Styles("Heading 1")
        .Selection.TypeText Text:=sString
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.TypeParagraph
        .Selection.InsertBreak Type:=wdPageBreak
    End With
    ProgressBar1.Value = ProgressBar1 + iCounter
End Sub

Private Sub WriteHeading(sString As String)
    On Error Resume Next
    With wdApp
        .Selection.Style = .ActiveDocument.Styles("Heading 2")
        .Selection.TypeText Text:=sString
        .Selection.TypeParagraph
        .Selection.Style = .ActiveDocument.Styles("Normal")
        .ActiveWindow.Selection.Font.Bold = False
        .Selection.TypeParagraph
        .Selection.Tables.Add Range:=.Selection.Range, NumRows:=1, NumColumns:=2
        .Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=168.45, RulerStyle:=wdAdjustNone
        .Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=276.4, RulerStyle:=wdAdjustNone
    End With
    With wdApp.Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    wdApp.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
End Sub
Private Sub WriteHealthChildWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmHealth")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName4"))
        
    With wdApp
        ' Child health
        frmHealth.rsHealthControlChild.Recordset.MoveFirst
        Do While Not frmHealth.rsHealthControlChild.Recordset.EOF
            If CLng(frmHealth.rsHealthControlChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmHealth.rsHealthControlChild.Recordset.Fields("ControlDate"), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label2")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(3).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        frmHealth.rsHealthControlChild.Recordset.MoveNext
        Loop
        
        .Selection.MoveDown Unit:=wdLine, Count:=1
        Call WriteHeading(rsLanguage.Fields("FormName5"))
    
        'vaccinations
        frmHealth.rsVaccinationChild.Recordset.MoveFirst
        Do While Not frmHealth.rsVaccinationChild.Recordset.EOF
            If CLng(frmHealth.rsVaccinationChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmHealth.rsVaccinationChild.Recordset.Fields("VaccinationDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label3")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmHealth.rsVaccinationChild.Recordset.Fields("VaccinationByDoctor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label4")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmHealth.rsVaccinationChild.Recordset.Fields("VaccinationWhere"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label2")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(4).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        frmHealth.rsVaccinationChild.Recordset.MoveNext
        Loop
        
        .Selection.MoveDown Unit:=wdLine, Count:=1
        Call WriteHeading(rsLanguage.Fields("FormName6"))
    
        'illness
        frmHealth.rsIllnessChild.Recordset.MoveFirst
        Do While Not frmHealth.rsIllnessChild.Recordset.EOF
            If CLng(frmHealth.rsIllnessChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmHealth.rsIllnessChild.Recordset.Fields("IllnessDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label5")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(5).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                If CBool(frmHealth.rsIllnessChild.Recordset.Fields("IllnessDoctor")) = 1 Then
                     .Selection.TypeText Text:=rsLanguage.Fields("Check1")
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=rsLanguage.Fields("Yes")
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=rsLanguage.Fields("label6")
                    .Selection.TypeText Text:=Format(frmHealth.rsIllnessChild.Recordset.Fields("IllnessDoctorName"))
                Else
                    .Selection.TypeText Text:=rsLanguage.Fields("Check1")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=rsLanguage.Fields("No")
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmHealth.rsIllnessChild.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub WriteInfancyNotesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesInfancy")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        frmFathersNotesInfancy.rsInfancyNotes.Recordset.MoveFirst
        Do While Not frmFathersNotesInfancy.rsInfancyNotes.Recordset.EOF
            If CLng(frmFathersNotesInfancy.rsInfancyNotes.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(frmFathersNotesInfancy.rsInfancyNotes.Recordset.Fields("NoteDate")) Then
                    DoEvents
                    .Selection.TypeText Text:=rsLanguage.Fields("Label1(0)")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=Format(CDate(frmFathersNotesInfancy.rsInfancyNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    Clipboard.Clear
                    Clipboard.SetText frmFathersNotesInfancy.RichText1.TextRTF, vbCFRTF
                    .Selection.Paste
                    DoEvents
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmFathersNotesInfancy.rsInfancyNotes.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintInfancyNotes()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesInfancy")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmFathersNotesInfancy.rsInfancyNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(.Fields("NoteDate")) Then
                    cPrint.FontBold = True
                    cPrint.pPrint rsLanguage.Fields("Label1(0)") & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1
                    cPrint.FontBold = False
                    cPrint.pPrint rsLanguage.Fields("Label1(1)"), 1, True
                    If Len(frmFathersNotesInfancy.RichText1.Text) <> 0 Then
                        cPrint.pMultiline frmFathersNotesInfancy.RichText1.Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint "", 3.5
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    If cPrint.pEndOfPage Then
                        DoNewPagePreview
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub WritePicturesChildWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPictures")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName3"))
            
    With wdApp
        frmPictures.rsChildPictures.Recordset.MoveFirst
        Do While Not frmPictures.rsChildPictures.Recordset.EOF
            If CLng(frmPictures.rsChildPictures.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                'print the pictures with notes
                .Selection.TypeText Text:=Format(frmPictures.rsChildPictures.Recordset.Fields("PictureCaption"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                If Not IsNull(frmPictures.rsChildPictures.Recordset.Fields("Picture")) Then
                    Clipboard.Clear
                    Clipboard.SetData frmPictures.Picture1(2).Picture, vbCFBitmap
                    .Selection.Paste
                End If
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
        frmPictures.rsChildPictures.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub WritePicturesInfantWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPictures")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName2"))
        
    With wdApp
    frmPictures.rsBirthPictures.Recordset.MoveFirst
    Do While Not frmPictures.rsInfancyPictures.Recordset.EOF
        If CLng(frmPictures.rsInfancyPictures.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
            'print the pictures with notes
            .Selection.TypeText Text:=Format(frmPictures.rsInfancyPictures.Recordset.Fields("PictureCaption"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            If Not IsNull(frmPictures.rsInfancyPictures.Recordset.Fields("Picture")) Then
                Clipboard.Clear
                Clipboard.SetData frmPictures.Picture1(1).Picture, vbCFBitmap
                .Selection.Paste
            End If
        End If
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
    frmPictures.rsInfancyPictures.Recordset.MoveNext
    Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub WriteToysChildWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmToys")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName2"))
        
        'childhood
    With wdApp
        frmToys.rsToysChild.Recordset.MoveFirst
        Do While Not frmToys.rsToysChild.Recordset.EOF
            .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(frmToys.rsToysChild.Recordset.Fields("ToyName"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
            .Selection.MoveRight Unit:=wdCell
            If IsDate(frmToys.rsToysChild.Recordset.Fields("PurchaseDate")) Then
                .Selection.TypeText Text:=Format(CDate(frmToys.rsToysChild.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(CDbl(frmToys.rsToysChild.Recordset.Fields("PurchasePrice")), "0.00") & "  " & frmToys.Label4(0).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
            .Selection.MoveRight Unit:=wdCell
            Clipboard.Clear
            Clipboard.SetText frmToys.RichTextBox1(1).TextRTF, vbCFRTF
            .Selection.Paste
            DoEvents
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
            .Selection.MoveRight Unit:=wdCell
            If Not IsNull(frmToys.rsToysChild.Recordset.Fields("Picture")) Then
                frmToys.Picture1.Picture = frmToys.Image1(1).Picture
                Clipboard.Clear
                Clipboard.SetData frmToys.Picture1.Picture, vbCFBitmap
                .Selection.Paste
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
        frmToys.rsToysChild.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub WriteToysInfantWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmToys")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName1"))
    
    With wdApp
        'infant
        frmToys.rsToysInfant.Recordset.MoveFirst
        Do While Not frmToys.rsToysInfant.Recordset.EOF
            .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(frmToys.rsToysInfant.Recordset.Fields("ToyName"))
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
            .Selection.MoveRight Unit:=wdCell
            If IsDate(frmToys.rsToysInfant.Recordset.Fields("PurchaseDate")) Then
                .Selection.TypeText Text:=Format(CDate(frmToys.rsToysInfant.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(CDbl(frmToys.rsToysInfant.Recordset.Fields("PurchasePrice")), "0.00") & "  " & frmToys.Label4(0).Caption
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
            .Selection.MoveRight Unit:=wdCell
            Clipboard.Clear
            Clipboard.SetText frmToys.RichTextBox1(0).TextRTF, vbCFRTF
            .Selection.Paste
            DoEvents
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
            .Selection.MoveRight Unit:=wdCell
            If Not IsNull(frmToys.rsToysInfant.Recordset.Fields("Picture")) Then
                frmToys.Picture1.Picture = frmToys.Image1(0).Picture
                Clipboard.Clear
                Clipboard.SetData frmToys.Picture1.Picture, vbCFBitmap
                .Selection.Paste
            Else
                .Selection.TypeText Text:=""
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
        frmToys.rsToysInfant.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintToysInfant()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmToys")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("Form")
    Call PrintHeaderComplete
    
    With frmToys.rsToysInfant.Recordset
        .MoveFirst
        Do While Not .EOF
            cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
            If Not IsNull(.Fields("ToyName")) Then
                cPrint.pPrint .Fields("ToyName"), 3.5
            Else
                cPrint.pPrint " ", 3.5
            End If
            cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
            If IsDate(.Fields("PurchaseDate")) Then
                cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
            Else
                cPrint.pPrint " ", 1.5
            End If
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
            cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
            cPrint.pPrint Format(CDbl(.Fields("PurchasePrice")), "0.00") & "  " & frmToys.Label4(0).Caption, 3.5
            cPrint.pPrint
            cPrint.pPrint rsLanguage.Fields("label1(4)"), 1
            If Len(frmToys.RichTextBox1(0).Text) <> 0 Then
                cPrint.pMultiline frmToys.RichTextBox1(0).Text, 1, cPrint.GetPaperWidth - 1, , False, True
            Else
                cPrint.pPrint " ", 1.5
            End If
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
            cPrint.pPrint rsLanguage.Fields("label1(3)"), 1
            If Not IsNull(.Fields("Picture")) Then
                cPrint.pPrintPicture frmToys.Image1(0).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
            End If
            cPrint.pPrint
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
        .MoveNext
        Loop
    End With
    
    rsLanguage.Close
End Sub

Private Sub PrintBooksInfant()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBooks")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmBooks.rsBooksInfant.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
                If Len(frmBooks.Text1(0).Text) <> 0 Then
                    cPrint.pPrint frmBooks.Text1(0).Text, 3.5
                Else
                    cPrint.pPrint "", 3.5
                End If
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
                If Len(frmBooks.Text1(2).Text) <> 0 Then
                    cPrint.pPrint frmBooks.Text1(2).Text, 3.5
                Else
                    cPrint.pPrint "", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
                If Len(frmBooks.RichTextBox1(0).Text) <> 0 Then
                    cPrint.pMultiline frmBooks.RichTextBox1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint "", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True
                If Len(frmBooks.Text1(3).Text) <> 0 Then
                    cPrint.pPrint frmBooks.Text1(3).Text, 3.5
                Else
                    cPrint.pPrint "", 3.5
                End If
                cPrint.pPrint rsLanguage.Fields("label1(4)"), 1, True
                If IsDate(.Recordset.Fields("PurchaseDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("PurchaseDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.CurrentX = LeftMargin
                cPrint.pPrint rsLanguage.Fields("label1(5)"), 1, True
                cPrint.pPrint Format(CDbl(frmBooks.Text1(4).Text), "0.00") & "  " & frmBooks.Label4(0).Caption, 3.5
                cPrint.pPrint rsLanguage.Fields("label1(6)"), 1, True
                If Len(frmBooks.Text1(1).Text) <> 0 Then
                    cPrint.pMultiline frmBooks.Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 1.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub WriteBooksInfantWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBooks")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        frmBooks.rsBooksInfant.Recordset.MoveFirst
        Do While Not frmBooks.rsBooksInfant.Recordset.EOF
            If CLng(frmBooks.rsBooksInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksInfant.Recordset.Fields("BookName"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksInfant.Recordset.Fields("BookAuthor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmBooks.RichTextBox1(0).TextRTF, vbCFRTF
                .Selection.Paste
                DoEvents
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksInfant.Recordset.Fields("BookIsbn"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
                .Selection.MoveRight Unit:=wdCell
                If IsDate(frmBooks.rsBooksInfant.Recordset.Fields("PurchaseDate")) Then
                    .Selection.TypeText Text:=Format(CDate(frmBooks.rsBooksInfant.Recordset.Fields("PurchaseDate")), "dd.mm.yyyy")
                Else
                    .Selection.TypeText Text:=""
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(5)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksInfant.Recordset.Fields("PurchasePrice"), "0.00") & "  " & frmBooks.Label4(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(6)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmBooks.rsBooksInfant.Recordset.Fields("WherePurchased"))
                .Selection.MoveRight Unit:=wdCell
            End If
        frmBooks.rsBooksInfant.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage2
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                If IsNull(.Fields("Check1")) Then
                    .Fields("Check1") = Check1.Caption
                Else
                    Check1.Caption = .Fields("Check1")
                End If
                If IsNull(.Fields("Check2")) Then
                    .Fields("Check2") = Check2.Caption
                Else
                    Check2.Caption = .Fields("Check2")
                End If
                If IsNull(.Fields("Check3")) Then
                    .Fields("Check3") = Check3.Caption
                Else
                    Check3.Caption = .Fields("Check3")
                End If
                If IsNull(.Fields("Check4")) Then
                    .Fields("Check4") = Check4.Caption
                Else
                    Check4.Caption = .Fields("Check4")
                End If
                If IsNull(.Fields("Check5")) Then
                    .Fields("Check5") = Check5.Caption
                Else
                    Check5.Caption = .Fields("Check5")
                End If
                If IsNull(.Fields("Option1(0)")) Then
                    .Fields("Option1(0)") = Option1(0).Caption
                Else
                    Option1(0).Caption = .Fields("Option1(0)")
                End If
                If IsNull(.Fields("Option1(1)")) Then
                    .Fields("Option1(1)") = Option1(1).Caption
                Else
                    Option1(1).Caption = .Fields("Option1(1)")
                End If
                If IsNull(.Fields("btnPrint")) Then
                    .Fields("btnPrint") = btnPrint.ToolTipText
                Else
                    btnPrint.ToolTipText = .Fields("btnPrint")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
                If IsNull(.Fields("Frame4")) Then
                    .Fields("Frame4") = Frame4.Caption
                Else
                    Frame4.Caption = .Fields("Frame4")
                End If
                If IsNull(.Fields("btnNext")) Then
                    .Fields("btnNext") = btnNext(0).Caption
                Else
                    btnNext(0).Caption = .Fields("btnNext")
                    btnNext(1).Caption = .Fields("btnNext")
                End If
                If IsNull(.Fields("btnPrevious")) Then
                    .Fields("btnPrevious") = btnPrevious(0).Caption
                Else
                    btnPrevious(0).Caption = .Fields("btnPrevious")
                    btnPrevious(1).Caption = .Fields("btnPrevious")
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
        .Fields("Check1") = Check1.Caption
        .Fields("Check2") = Check2.Caption
        .Fields("Check3") = Check3.Caption
        .Fields("Check4") = Check4.Caption
        .Fields("Check5") = Check5.Caption
        .Fields("Option1(0)") = Option1(0).Caption
        .Fields("Option1(1)") = Option1(1).Caption
        .Fields("btnPrint") = btnPrint.ToolTipText
        .Fields("Frame3") = Frame3.Caption
        .Fields("Frame4") = Frame4.Caption
        .Fields("btnNext") = btnNext(0).Caption
        .Fields("btnPrevious") = btnPrevious(0).Caption
        .Fields("PrintFrontPage") = "Memory Book For "
        .Fields("PrintPregnancyPage") = "Pregnancy"
        .Fields("PrintBirthPage") = "Birth"
        .Fields("PrintBabtismPage") = "Babtism"
        .Fields("PrintBabyPage") = "Baby"
        .Fields("PrintChildhoodPage") = "Childhood"
        .Fields("PrintIndex") = "Index:"
        .Fields("sDate") = "Date: "
        .Fields("sPage") = "Page: "
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Public Sub WriteHealthInfantWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmHealth")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName1"))
    
    With wdApp
        frmHealth.rsHealthControlInfant.Recordset.MoveFirst
        Do While Not frmHealth.rsHealthControlInfant.Recordset.EOF
            If CLng(frmHealth.rsHealthControlInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmHealth.rsHealthControlInfant.Recordset.Fields("ControlDate"), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label2")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(0).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        frmHealth.rsHealthControlInfant.Recordset.MoveNext
        Loop
        
        .Selection.MoveDown Unit:=wdLine, Count:=1
        Call WriteHeading(rsLanguage.Fields("FormName2"))
        
        'write vaccinations
        frmHealth.rsVaccinationInfant.Recordset.MoveFirst
        Do While Not frmHealth.rsVaccinationInfant.Recordset.EOF
            If CLng(frmHealth.rsVaccinationInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmHealth.rsVaccinationInfant.Recordset.Fields("VaccinationDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label3")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmHealth.rsVaccinationInfant.Recordset.Fields("VaccinationByDoctor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label4")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmHealth.rsVaccinationInfant.Recordset.Fields("VaccinationWhere"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label2")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(1).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        frmHealth.rsVaccinationInfant.Recordset.MoveNext
        Loop
        
        .Selection.MoveDown Unit:=wdLine, Count:=1
        Call WriteHeading(rsLanguage.Fields("FormName3"))
        
        'write illness infant
        frmHealth.rsIllnessInfant.Recordset.MoveFirst
        Do While Not frmHealth.rsIllnessInfant.Recordset.EOF
            If CLng(frmHealth.rsIllnessInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmHealth.rsIllnessInfant.Recordset.Fields("IllnessDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label5")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(2).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                If CBool(frmHealth.rsIllnessInfant.Recordset.Fields("IllnessDoctor")) = 1 Then
                     .Selection.TypeText Text:=rsLanguage.Fields("Check1")
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=rsLanguage.Fields("Yes")
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=rsLanguage.Fields("label6")
                    .Selection.TypeText Text:=Format(frmHealth.rsIllnessInfant.Recordset.Fields("IllnessDoctorName"))
                Else
                    .Selection.TypeText Text:=rsLanguage.Fields("Check1")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=rsLanguage.Fields("No")
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmHealth.rsIllnessInfant.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Public Sub PrintHealthInfant()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmHealth")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName1")
    PrintHeaderComplete
    
    'health infant
    With frmHealth.rsHealthControlInfant.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                cPrint.pPrint Format(CDate(.Fields("ControlDate")), "dd.mm.yyyy"), 3.5
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label2"), 1, True
                If Len(frmHealth.RichTextBox1(0).Text) <> 0 Then
                    cPrint.pMultiline frmHealth.RichTextBox1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint "", 3.5
                End If
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
    
    sHeader = rsLanguage.Fields("FormName2")
    PrintHeaderComplete
    
    'vaccinations infant
    With frmHealth.rsVaccinationInfant.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                cPrint.pPrint Format(CDate(.Fields("VaccinationDate")), "dd.mm.yyyy"), 3.5
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label3"), 1, True
                If Not IsNull(.Fields("VaccinationByDoctor")) Then
                    cPrint.pPrint .Fields("VaccinationByDoctor"), 3.5
                Else
                    cPrint.pPrint "", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label4"), 1, True
                If Not IsNull(.Fields("VaccinationWhere")) Then
                    cPrint.pPrint .Fields("VaccinationWhere"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label2"), 1, True
                If Len(frmHealth.RichTextBox1(1).Text) <> 0 Then
                    cPrint.pMultiline frmHealth.RichTextBox1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
    
    sHeader = rsLanguage.Fields("FormName3")
    PrintHeaderComplete
    
    'illness infant
    With frmHealth.rsIllnessInfant.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                If IsDate(.Fields("IllnessDate")) Then
                    cPrint.pPrint Format(CDate(.Fields("IllnessDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label5"), 1, True
                If Len(frmHealth.RichTextBox1(2).Text) <> 0 Then
                    cPrint.pMultiline frmHealth.RichTextBox1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If CBool(.Fields("IllnessDoctor")) = 1 Then
                    cPrint.pPrint rsLanguage.Fields("Check1"), 1, True
                    cPrint.pPrint rsLanguage.Fields("Yes"), 3.5
                    cPrint.pPrint rsLanguage.Fields("label6"), 1, True
                    cPrint.pPrint Format(.Fields("IllnessDoctorName")), 3.5
                Else
                    cPrint.pPrint rsLanguage.Fields("Check1"), 1, True
                    cPrint.pPrint rsLanguage.Fields("No"), 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .MoveNext
        Loop
    End With
    
    rsLanguage.Close
End Sub

Private Sub PrintFoodHabits()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFoodHabits")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    For i = 0 To 3
        cPrint.pPrint rsLanguage.Fields(i + 2), 1, True
        If Len(frmFoodHabits.Text1(i).Text) <> 0 Then
            cPrint.pPrint frmFoodHabits.Text1(i).Text & "  " & rsLanguage.Fields("label2"), 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
    Next
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    cPrint.pPrint
    cPrint.pPrint rsLanguage.Fields("label1(4)"), 1, True
    If Len(frmFoodHabits.RichTextBox1.Text) <> 0 Then
        cPrint.pMultiline frmFoodHabits.RichTextBox1.Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
End Sub

Private Sub WriteFoodHabitsWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFoodHabits")
    ReadLanguage
    DoEvents
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmFoodHabits.Text1(0).Text & " " & rsLanguage.Fields("label2")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmFoodHabits.Text1(1).Text & " " & rsLanguage.Fields("label2")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmFoodHabits.Text1(2).Text & " " & rsLanguage.Fields("label2")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmFoodHabits.Text1(3).Text & " " & rsLanguage.Fields("label2")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmFoodHabits.RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        .Selection.MoveDown Unit:=wdLine, Count:=1
        DoEvents
    End With
    rsLanguage.Close
End Sub

Private Sub WriteFirstTimesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFirstTimes")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        For i = 0 To 10
            If IsDate(frmFirstTimes.Date1(i).Text) Then
                .Selection.TypeText Text:=rsLanguage.Fields(i + 4)
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("Label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmFirstTimes.Date1(i).Text), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("Label2")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=frmFirstTimes.Text1(i).Text
                .Selection.MoveRight Unit:=wdCell
            End If
        Next
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintFirstTimes()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFirstTimes")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    For i = 0 To 10
        If IsDate(frmFirstTimes.Date1(i).Text) Then
            cPrint.FontBold = True
            cPrint.pPrint rsLanguage.Fields(i + 4), 1
            cPrint.FontBold = False
            cPrint.pPrint rsLanguage.Fields("Label1") & "  " & Format(CDate(frmFirstTimes.Date1(i).Text), "dd.mm.yyyy"), 1
            cPrint.pPrint rsLanguage.Fields("Label2"), 1, True
            If Len(frmFirstTimes.Text1(i).Text) <> 0 Then
                cPrint.pMultiline frmFirstTimes.Text1(i).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
            Else
                cPrint.pPrint " ", 3.5
            End If
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
        End If
    Next
End Sub

Private Sub WriteTeethWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmTeeth")
    ReadLanguage
    
    With wdApp
        .Selection.InsertBreak Type:=wdPageBreak
        .Selection.Style = .ActiveDocument.Styles("Heading 2")
        .Selection.TypeText Text:=rsLanguage.Fields("FormName")
        .Selection.TypeParagraph
        .Selection.Style = .ActiveDocument.Styles("Normal")
    
        Clipboard.Clear
        Clipboard.SetText frmTeeth.RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.TypeParagraph
        .Selection.Style = .ActiveDocument.Styles("Normal")
        .Selection.Tables.Add Range:=.Selection.Range, NumRows:=1, NumColumns:=2
        .Selection.Tables(1).Columns(1).SetWidth ColumnWidth:=168.45, RulerStyle:=wdAdjustNone
        .Selection.Tables(1).Columns(2).SetWidth ColumnWidth:=276.4, RulerStyle:=wdAdjustNone
    End With
    
    With wdApp.Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    
    With wdApp
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label3")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        
        For n = 0 To 9
            If IsDate(frmTeeth.Date1(n).Text) Then
                .Selection.TypeText Text:=rsLanguage.Fields(n + 1)
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmTeeth.Date1(n).Text), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
            End If
        Next
        
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label2")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        
        For n = 10 To 19
            If IsDate(frmTeeth.Date1(n).Text) Then
                .Selection.TypeText Text:=rsLanguage.Fields(n + 1)
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmTeeth.Date1(n).Text), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
            End If
        Next
        
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintTeeth()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmTeeth")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    cPrint.pMultiline frmTeeth.RichTextBox1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint rsLanguage.Fields("label3"), 1
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    
    With frmTeeth
        For n = 0 To 9
            cPrint.CurrentX = LeftMargin
            If IsDate(.Date1(n).Text) Then
                cPrint.pPrint rsLanguage.Fields(n + 1), 1, True
                cPrint.pPrint Format(CDate(.Date1(n).Text), "dd.mm.yyyy"), 3.5
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        Next
    End With
    
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint rsLanguage.Fields("label2"), 1
    cPrint.pPrint
    
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    
    With frmTeeth
        For n = 10 To 19
            cPrint.CurrentX = LeftMargin
            If IsDate(.Date1(n).Text) Then
                cPrint.pPrint rsLanguage.Fields(n + 1), 1, True
                cPrint.pPrint Format(CDate(.Date1(n).Text), "dd.mm.yyyy"), 3.5
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        Next
    End With
    rsLanguage.Close
End Sub

Private Sub PrintWeightLength()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmWeightLength")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("Form")
    PrintHeaderComplete
    
    With frmWeightLength
        For n = 0 To 12
            If IsDate(.Date1(n).Text) Then
                cPrint.pPrint rsLanguage.Fields("label1"), 1, True
                cPrint.pPrint Format(CDate(.Date1(n).Text), "dd.mm.yyyy"), 3.5
                cPrint.pPrint rsLanguage.Fields("label2") & "  " & .Text1(n).Text & "  " & .cmbLength.Text, 1
                cPrint.pPrint rsLanguage.Fields("label3") & "  " & .Text2(n).Text & "  " & .cmbWeight.Text, 1
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        Next
    End With
    rsLanguage.Close
End Sub

Private Sub WriteWeightLengthWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmWeightLength")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("Form"))
    
    With wdApp
        For n = 0 To 12
            If IsDate(frmWeightLength.Date1(n).Text) Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmWeightLength.Date1(n).Text), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label2") & "  " & frmWeightLength.Text1(n).Text & " " & frmWeightLength.cmbLength.Text
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label3") & "  " & frmWeightLength.Text2(n).Text & " " & frmWeightLength.cmbWeight.Text
                .Selection.MoveRight Unit:=wdCell
            End If
        Next
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintBaptismNotes()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesBaptism")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmFathersNotesBaptism.rsBaptismNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(.Fields("NoteDate")) Then
                    cPrint.FontBold = True
                    cPrint.pPrint rsLanguage.Fields("Label1(0)") & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1
                    cPrint.FontBold = False
                    If Len(frmFathersNotesBaptism.RichText1.Text) <> 0 Then
                        cPrint.pMultiline frmFathersNotesBaptism.RichText1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    If cPrint.pEndOfPage Then
                        DoNewPagePreview
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub WriteBaptismNotesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesBaptism")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        frmFathersNotesBaptism.rsBaptismNotes.Recordset.MoveFirst
        Do While Not frmFathersNotesBaptism.rsBaptismNotes.Recordset.EOF
            If CLng(frmFathersNotesBaptism.rsBaptismNotes.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(frmFathersNotesBaptism.rsBaptismNotes.Recordset.Fields("NoteDate")) Then
                    DoEvents
                    .Selection.TypeText Text:=rsLanguage.Fields("Label1(0)")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=Format(CDate(frmFathersNotesBaptism.rsBaptismNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    Clipboard.Clear
                    Clipboard.SetText frmFathersNotesBaptism.RichText1.TextRTF, vbCFRTF
                    .Selection.Paste
                    DoEvents
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmFathersNotesBaptism.rsBaptismNotes.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintBaptismPic()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBaptismPictures")
    ReadLanguage
    bFirstWrite = True
    
    sHeader = rsLanguage.Fields("FormName")
    Call PrintHeaderComplete
    
    With frmBaptismPictures.rsBaptismPictures.Recordset
        .MoveFirst
        Do While Not .EOF
            If Not bFirstWrite Then
                DoNewPagePreview
            End If
            cPrint.pMultiline frmBaptismPictures.Text1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
            cPrint.pPrint
            cPrint.pPrint
            cPrint.pPrintPicture frmBaptismPictures.Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
            bFirstWrite = False
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub
Private Sub WriteBaptismPicWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBaptismPictures")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.TypeText Text:=Format(frmBaptismPictures.rsBaptismPictures.Recordset.Fields("BabtismPictureCaption"))
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData frmBaptismPictures.Image1.Picture, vbCFBitmap
        .Selection.Paste
        .Selection.MoveDown Unit:=wdLine, Count:=1
    End With
    rsLanguage.Close
End Sub

Private Sub PrintBaptism()
Dim iX As Integer
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBaptism")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    Call PrintHeaderComplete
    
    With frmBaptism
        cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True   'baptism date
        If IsDate(.Date1.Text) Then
            cPrint.pPrint .Date1.Text, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True   'baptism time
        If IsDate(.MaskEdBox1.Text) Then
            cPrint.pPrint Format(.MaskEdBox1.Text, "hh:mm"), 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True   'where ?
        If Len(.Text1(0).Text) <> 0 Then
            cPrint.pMultiline .Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True   'minister name
        If Len(.Text1(1).Text) <> 0 Then
            cPrint.pPrint .Text1(1).Text, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label1(5)"), 1, True   'I was named ..
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pPrint .Text1(2).Text, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label1(6)"), 1, True   'my name was chosen by ..
        If Len(.Text1(3).Text) <> 0 Then
            cPrint.pMultiline .Text1(3).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label1(7)"), 1, True   'got this name because...
        If Len(.Text1(4).Text) <> 0 Then
            cPrint.pMultiline .Text1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        
        'now print godfather / godmother
        sHeader = rsLanguage.Fields("Frame2")
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame2"), 1    'goodmother / goodfather and addresses
        cPrint.FontBold = False
        cPrint.pPrint
        If Len(.RichTextBox1.Text) <> 0 Then
            cPrint.pMultiline .RichTextBox1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 1
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        
        'print attendees
        sHeader = rsLanguage.Fields("Frame3")
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame3"), 1    'attendees and family relations
        cPrint.FontBold = False
        cPrint.pPrint
        If Len(.RichTextBox2.Text) <> 0 Then
            cPrint.pMultiline .RichTextBox2.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 1
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        
        'print gifts
        sHeader = rsLanguage.Fields("Frame4")   'gifts and from whom
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame4"), 1
        cPrint.FontBold = False
        cPrint.pPrint
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        If Len(.RichTextBox3.Text) <> 0 Then
            cPrint.pMultiline .RichTextBox3.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 1
        End If
        
        'print notes
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame5"), 1  'baptism notes
        cPrint.FontBold = False
        cPrint.pPrint
        If Len(.RichTextBox4.Text) Then
            cPrint.pMultiline .RichTextBox4.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        
        'print the church picture
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label1(4)"), 1   'church picture
        cPrint.pPrint
        If Not IsNull(.rsBaptism.Recordset.Fields("ChurchPicture")) Then
            cPrint.pPrintPicture .Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
    End With
    rsLanguage.Close
End Sub


Private Sub WriteBaptismWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBaptism")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.Tables(1).AllowAutoFit = False
        .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmBaptism.Date1.Text, "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmBaptism.MaskEdBox1.Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBaptism.Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBaptism.Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        'I was named
        .Selection.TypeText Text:=rsLanguage.Fields("label1(5)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBaptism.Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(6)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBaptism.Text1(3).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(7)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBaptism.Text1(4).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData frmBaptism.Image1.Picture, vbCFBitmap
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'godmothers /godfathers
        .Selection.TypeText Text:=rsLanguage.Fields("Frame2")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmBaptism.RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'attendees
        .Selection.TypeText Text:=rsLanguage.Fields("Frame3")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmBaptism.RichTextBox2.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'gifts
        .Selection.TypeText Text:=rsLanguage.Fields("Frame4")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmBaptism.RichTextBox3.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        'notes
        .Selection.TypeText Text:=rsLanguage.Fields("Frame5")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmBaptism.RichTextBox4.TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveDown Unit:=wdLine, Count:=1
    End With
    rsLanguage.Close
End Sub

Private Sub WriteBirthNotesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesBirth")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        frmFathersNotesBirth.rsBirthNotes.Recordset.MoveFirst
        Do While Not frmFathersNotesBirth.rsBirthNotes.Recordset.EOF
            If CLng(frmFathersNotesBirth.rsBirthNotes.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(frmFathersNotesBirth.rsBirthNotes.Recordset.Fields("NoteDate")) Then
                    DoEvents
                    .Selection.TypeText Text:=rsLanguage.Fields("Label1(1)")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=Format(CDate(frmFathersNotesBirth.rsBirthNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    Clipboard.Clear
                    Clipboard.SetText frmFathersNotesBirth.RichText1.TextRTF, vbCFRTF
                    .Selection.Paste
                    DoEvents
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmFathersNotesBirth.rsBirthNotes.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintBirthNotes()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesBirth")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmFathersNotesBirth.rsBirthNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(.Fields("NoteDate")) Then
                    cPrint.FontBold = True
                    cPrint.pPrint rsLanguage.Fields("Label1(1)") & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1, True
                    cPrint.FontBold = False
                    If Len(frmFathersNotesBirth.RichText1.Text) <> 0 Then
                        cPrint.pMultiline frmFathersNotesBirth.RichText1.Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint "", 3.5
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    If cPrint.pEndOfPage Then
                        DoNewPagePreview
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub WriteFirstPramWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFirstPram")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmFirstPram.Date1.Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmFirstPram.Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmFirstPram.Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmFirstPram.Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData frmFirstPram.Image1.Picture, vbCFBitmap
        .Selection.Paste
        .Selection.MoveDown Unit:=wdLine, Count:=1
    End With
    rsLanguage.Close
End Sub

Private Sub PrintFirstPram()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFirstPram")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmFirstPram
        cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
        If IsDate(.Date1.Text) Then
            cPrint.pPrint Format(CDate(.Date1.Text), "dd.mm.yyyy"), 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
        If Len(.Text1(0).Text) <> 0 Then
            cPrint.pPrint .Text1(0).Text, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pMultiline .Text1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint "", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
        If Len(.Text1(1).Text) <> 0 Then
            cPrint.pMultiline .Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        If Not IsNull(.rsFirstPram.Recordset.Fields("Picture")) Then
            cPrint.pPrintPicture .Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
        End If
    End With
End Sub

Private Sub PrintPicturesBirth()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPictures")
    ReadLanguage
    
        sHeader = rsLanguage.Fields("FormName1")
        PrintHeaderComplete
        bFirstWrite = True
        
        With frmPictures.rsBirthPictures.Recordset
            .MoveFirst
            Do While Not .EOF
                If CLng(.Fields("ChildNo")) = CLng(glChildNo) And Not IsNull(.Fields("Picture")) Then
                    'print picture notes
                    If Not bFirstWrite Then
                        DoNewPagePreview
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    cPrint.pPrint rsLanguage.Fields("label2"), 1, True
                    If Len(frmPictures.Text2(0).Text) <> 0 Then
                        cPrint.pMultiline frmPictures.Text2(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 3.5
                    End If
                    cPrint.pPrint
                    cPrint.pPrintPicture frmPictures.Picture1(0).Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
                    bFirstWrite = False
                End If
            .MoveNext
            Loop
        End With
    rsLanguage.Close
End Sub

Private Sub WritePicturesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPictures")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName1"))
    
        With wdApp
            .Selection.Tables(1).AllowAutoFit = False
            frmPictures.rsBirthPictures.Recordset.MoveFirst
            Do While Not frmPictures.rsBirthPictures.Recordset.EOF
                If CLng(frmPictures.rsBirthPictures.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                    'print the pictures with notes
                    .Selection.TypeText Text:=Format(frmPictures.rsBirthPictures.Recordset.Fields("PictureCaption"))
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    If Not IsNull(frmPictures.rsBirthPictures.Recordset.Fields("Picture")) Then
                        Clipboard.Clear
                        Clipboard.SetData frmPictures.Picture1(0).Picture, vbCFBitmap
                        .Selection.Paste
                    End If
                End If
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            frmPictures.rsBirthPictures.Recordset.MoveNext
            Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintHospital()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmHospital")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName1")
    PrintHeaderComplete
    
    With frmHospital
        'hospital notes
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame1(0)"), 1
        cPrint.FontBold = False
        If Len(.RichTextBox1(0).Text) <> 0 Then
            cPrint.pMultiline .RichTextBox1(0).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        
        'hospital Acquaintance
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("Frame1(1)"), 1
        cPrint.FontBold = False
        If Len(.RichTextBox1(1).Text) <> 0 Then
            cPrint.pMultiline .RichTextBox1(1).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        
        'leaving hospital
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True   'date left hospital
        If IsDate(CDate(.Date1.Text)) Then
            cPrint.pPrint Format(CDate(.Date1.Text), "dd.mm.yyyy"), 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True   'time left hospital
        If Len(.Text1(5).Text) <> 0 Then
            cPrint.pPrint Format(.Text1(5).Text, "hh:mm"), 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True   'who drove us home
        If Len(.RichTextBox1(2).Text) <> 0 Then
            cPrint.pMultiline .RichTextBox1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(4)"), 1, True   'our car was a..
        If Len(.Text1(3).Text) <> 0 Then
            cPrint.pMultiline .Text1(3).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        If Not IsNull(.rsLeaving.Recordset.Fields("OurCarPic")) Then
            cPrint.pPrintPicture .Image1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, .Image1.Picture.Height, False, True
            cPrint.CurrentY = cPrint.CurrentY + .Image1.Picture.Height
        End If
        
        'home from hospital
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label1(6)"), 1, True   'time we came home
        If Len(.Text1(6).Text) <> 0 Then
            cPrint.pPrint Format(.Text1(6).Text, "hh:mm"), 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True   'our address
        If Len(.Text1(0).Text) <> 0 Then
            cPrint.pMultiline .Text1(0).Text, 3.5, , cPrint.GetPaperWidth - 1.2, False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("label1(7)"), 1   'was met at home by...
        cPrint.FontBold = False
        If Len(.Text1(1).Text) <> 0 Then
            cPrint.pMultiline .Text1(1).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("label1(8)"), 1, True   'time slept first night home
        If Len(.Text1(7).Text) <> 0 Then
            cPrint.pPrint .Text1(7).Text, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        cPrint.pPrint rsLanguage.Fields("label1(9)"), 1, True   'woke no.of times
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pPrint .Text1(2).Text, 3.5
        Else
            cPrint.pPrint " ", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.FontBold = True
        cPrint.pPrint rsLanguage.Fields("label1(10)"), 1  'first person to visit me at home
        cPrint.FontBold = False
        If Len(.Text1(4).Text) Then
            cPrint.pMultiline .Text1(4).Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint , 3.5
        End If
    End With
    rsLanguage.Close
End Sub

Private Sub WriteHospitalWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmHospital")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName1"))
    
    With wdApp
        .Selection.Tables(1).AllowAutoFit = False
        'hospital notes
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(0)")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmHospital.RichTextBox1(0).TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        'hospital Acquaintance
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(1)")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmHospital.RichTextBox1(1).TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        'leaving hospital
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("FormName3")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmHospital.Date1.Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmHospital.DTPicker1.Text), "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmHospital.RichTextBox1(2).TextRTF, vbCFRTF
        .Selection.Paste
        DoEvents
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmHospital.Text1(3).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(5)")
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetData frmHospital.Image1.Picture, vbCFBitmap
        .Selection.Paste
        'home from hospital
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("FormName4")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmHospital.Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(6)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmHospital.DTPicker2.Text), "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(7)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmHospital.Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(8)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmHospital.DTPicker3.Text), "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(9)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmHospital.Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(10)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmHospital.Text1(4).Text
        .Selection.MoveDown Unit:=wdLine, Count:=1
    End With
    rsLanguage.Close
End Sub

Private Sub PrintBirth()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBirth")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
    If IsDate(frmBirth.Date1(0).Text) Then
        cPrint.pPrint Format(CDate(frmBirth.Date1(0).Text), "dd.mm.yyyy"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
    If IsDate(frmBirth.Text1(8).Text) Then
        cPrint.pPrint Format(frmBirth.Text1(8).Text, "hh:mm"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
    If Len(frmBirth.Text1(0).Text) <> 0 Then
        cPrint.pMultiline frmBirth.Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True
    If Len(frmBirth.Text1(1).Text) <> 0 Then
        cPrint.pMultiline frmBirth.Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    
    cPrint.pPrint
    cPrint.pPrint rsLanguage.Fields("label1(4)"), 1, True
    If IsDate(frmBirth.Date1(1).Text) Then
        cPrint.pPrint Format(CDate(frmBirth.Date1(1).Text), "dd.mm.yyyy"), 3.5  'at hospital date
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label1(5)"), 1, True
    If IsDate(frmBirth.Text1(9).Text) Then
        cPrint.pPrint Format(frmBirth.Text1(9).Text, "hh:mm"), 3.5  'at hospital time
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label1(6)"), 1, True
    If Len(frmBirth.Text1(2).Text) <> 0 Then
        cPrint.pMultiline frmBirth.Text1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    
    cPrint.pPrint
    cPrint.pPrint rsLanguage.Fields("label1(7)"), 1, True
    If Len(frmBirth.Text1(3).Text) <> 0 Then
        cPrint.pPrint frmBirth.Text1(3).Text & "  " & frmBirth.cmbDimension(0).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label1(8)"), 1, True
    If IsDate(frmBirth.Date1(2).Text) Then
        cPrint.pPrint Format(frmBirth.Date1(2).Text, "dd.mm.yyyy"), 3.5
    Else
        cPrint.pPrint "", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label1(9)"), 1, True
    If IsDate(frmBirth.Text1(10).Text) Then
        cPrint.pPrint Format(CDate(frmBirth.Text1(10).Text), "hh:mm"), 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label1(10)"), 1, True
    If Len(frmBirth.Text1(6).Text) <> 0 Then
        cPrint.pPrint frmBirth.Text1(6).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    cPrint.pPrint rsLanguage.Fields("label1(11)"), 1, True
    If Len(frmBirth.Text1(7).Text) <> 0 Then
        cPrint.pPrint frmBirth.Text1(7).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("Frame3"), 1, True
    If Len(frmBirth.Text1(4).Text) <> 0 Then
        cPrint.pPrint Format(frmBirth.Text1(4).Text, "0.000") & "" & "  " & frmBirth.cmbDimension(1).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("Frame4"), 1, True
    If Len(frmBirth.Text1(5).Text) <> 0 Then
        cPrint.pPrint Format(frmBirth.Text1(5).Text, "0.00") & "" & "  " & frmBirth.cmbDimension(2).Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        DoNewPagePreview
    End If
    cPrint.pPrint "Note:", 1, True
    If Len(frmBirth.RichTextBox1.Text) <> 0 Then
        cPrint.pMultiline frmBirth.RichTextBox1.Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
End Sub


Private Sub WriteBirthWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmBirth")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmBirth.Date1(0).Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmBirth.Text1(8).Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBirth.Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBirth.Text1(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmBirth.Date1(1).Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(5)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmBirth.Text1(9).Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(6)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBirth.Text1(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(7)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBirth.Text1(3).Text & "  " & frmBirth.cmbDimension(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(8)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(CDate(frmBirth.Date1(2).Text), "dd.mm.yyyy")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(9)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmBirth.Text1(10).Text, "hh:mm")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(10)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBirth.Text1(6).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(11)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmBirth.Text1(7).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame3")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmBirth.Text1(4).Text, "0.000") & "  " & frmBirth.cmbDimension(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame4")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmBirth.Text1(5).Text, "0.000") & "  " & frmBirth.cmbDimension(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:="Note"
        .Selection.MoveRight Unit:=wdCell
        Clipboard.Clear
        Clipboard.SetText frmBirth.RichTextBox1.TextRTF, vbCFRTF
        .Selection.Paste
        .Selection.MoveDown Unit:=wdLine, Count:=1
    End With
    
    rsLanguage.Close
End Sub

Private Sub WriteFatherPregnancyNotesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesPregnancy")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        frmFathersNotesPregnancy.rsPregnancyNotes.Recordset.MoveFirst
        Do While Not frmFathersNotesPregnancy.rsPregnancyNotes.Recordset.EOF
            If CLng(frmFathersNotesPregnancy.rsPregnancyNotes.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(frmFathersNotesPregnancy.rsPregnancyNotes.Recordset.Fields("NoteDate")) Then
                    DoEvents
                    .Selection.TypeText Text:=rsLanguage.Fields("Label1(1)")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=Format(CDate(frmFathersNotesPregnancy.rsPregnancyNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    Clipboard.Clear
                    Clipboard.SetText frmFathersNotesPregnancy.RichText1.TextRTF, vbCFRTF
                    .Selection.Paste
                    DoEvents
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        frmFathersNotesPregnancy.rsPregnancyNotes.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintFatherPregnancyNotes()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmFathersNotesPregnancy")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmFathersNotesPregnancy.rsPregnancyNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                If IsDate(.Fields("NoteDate")) Then
                    cPrint.FontBold = True
                    cPrint.pPrint rsLanguage.Fields("Label1(1)") & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1
                    cPrint.FontBold = False
                    If Len(frmFathersNotesPregnancy.RichText1.Text) <> 0 Then
                        cPrint.pMultiline frmFathersNotesPregnancy.RichText1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                    Else
                        cPrint.pPrint " ", 1
                    End If
                    cPrint.pPrint
                    cPrint.pPrint
                    If cPrint.pEndOfPage Then
                        DoNewPagePreview
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub PrintAntenatal()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmAntenatal")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("Form")
    PrintHeaderComplete
    
    With frmAntenatal
        .rsAntenatal.Recordset.MoveFirst
        Do While Not .rsAntenatal.Recordset.EOF
            If CLng(.rsAntenatal.Recordset.Fields("ChildNo")) = glChildNo Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1(0)") & "  " & Format(CDate(.rsAntenatal.Recordset.Fields("AntenatalDate")), "dd.mm.yyyy"), 1
                cPrint.FontBold = False
                cPrint.pMultiline .RichTextBox1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
                cPrint.pPrint
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .rsAntenatal.Recordset.MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub WriteAntenatalWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmAntenatal")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("Form"))
    
    With wdApp
        frmAntenatal.rsAntenatal.Recordset.MoveFirst
        Do While Not frmAntenatal.rsAntenatal.Recordset.EOF
            If CLng(frmAntenatal.rsAntenatal.Recordset.Fields("ChildNo")) = glChildNo Then
                DoEvents
                .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmAntenatal.rsAntenatal.Recordset.Fields("AntenatalDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmAntenatal.RichTextBox1.TextRTF, vbCFRTF
                .Selection.Paste
                DoEvents
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .ActiveWindow.Selection.Font.Name = "Times New Roman"
                .ActiveWindow.Selection.Font.Size = 10
                .Selection.MoveRight Unit:=wdCell
            End If
        frmAntenatal.rsAntenatal.Recordset.MoveNext
        Loop
    End With
    rsLanguage.Close
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
End Sub

Private Sub PrintDefaultPrinter()
    On Error GoTo errPrintDefault
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
        
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
        
    'print the heading
    PrintFrontPage
    DoEvents
        
    If Check1.Value = 1 Then
        'print first block name
        PrintBlockName (rsLanguage2.Fields("PrintPregnancyPage"))   'ok
        DoEvents
        
        '1. print - pregnancy
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmIamPregnant.SelectPregnancy Then
            PrintIamPregnant
            Unload frmIamPregnant
        End If
        '2. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmPregnancyControl.SelectControl Then
            PrintPregnancyControl
            Unload frmPregnancyControl
        End If
        '3. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmPregnancyNotes.SelectNotes Then
            PrintPregnancyNotes 'ok
        End If
        Unload frmPregnancyNotes
        '4. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmAntenatal.SelectAntenatalChild Then
            PrintAntenatal
            Unload frmAntenatal
        End If
        '5. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmFathersNotesPregnancy.SelectRecords Then
            PrintFatherPregnancyNotes
            Unload frmFathersNotesPregnancy
        End If
    End If
        
    If Check2.Value = 1 Then
        'print the Birth Heading
        PrintBlockName (rsLanguage2.Fields("PrintBirthPage"))
        DoEvents
        
        '6. print - birth
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmBirth.ShowChild Then
            PrintBirth
            Unload frmBirth
        End If
        '7. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        With frmHospital
            .SelectHospitalAcquaintance
            .SelectHospitalHome
            .SelectHospitalLeaving
            .SelectHospitalNotes
        End With
        PrintHospital
        Unload frmHospital
        '8. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmPictures.SelectPicBirth Then
            PrintPicturesBirth
            Unload frmPictures
        End If
        '9. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmFirstPram.SelectPram Then
            PrintFirstPram
        Unload frmFirstPram
        End If
        '10. print - when I was born
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmWhenIWasBorn.SelectBorn Then
            PrintWhenIWasBorn
            DoEvents
            Unload frmWhenIWasBorn
        End If
        '11. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmFathersNotesBirth.SelectRecords Then
            PrintBirthNotes
            Unload frmFathersNotesBirth
        End If
    End If
        
    If Check3.Value = 1 Then
        'Baptism
        PrintBlockName (rsLanguage2.Fields("PrintBabtismPage"))
        DoEvents
        
        '12. print - Babtism
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        Call frmBaptism.SelectChild
        PrintBaptism
        Unload frmBaptism
        '13. print the pictures
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmBaptismPictures.FillList2 Then
            PrintBaptismPic
            Unload frmBaptismPictures
        End If
        '14. print the notes
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmFathersNotesBaptism.SelectRecords Then
            PrintBaptismNotes
            Unload frmFathersNotesBaptism
        End If
    End If
        
    If Check4.Value = 1 Then
        'INFANCY
        PrintBlockName (rsLanguage2.Fields("PrintBabyPage"))
        DoEvents
        
        '15. print - infant
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmWeightLength.SelectChild Then
            PrintWeightLength
            Unload frmWeightLength
        End If
        '16. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmTeeth.SelectChild Then
            PrintTeeth
            Unload frmTeeth
        End If
        '17. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmFirstTimes.ShowFirstTime Then
            PrintFirstTimes
            Unload frmFirstTimes
        End If
        '18. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmFoodHabits.ReadFoodHabits Then
            PrintFoodHabits
            Unload frmFoodHabits
        End If
        '19. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        Call frmHealth.SelectHealthChild
        PrintHealthInfant
        Unload frmHealth
        '20. print - books
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        frmBooks.SelectBooksChild
        PrintBooksInfant
        Unload frmBooks
        '21. print - toys
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        frmToys.SelectToys
        PrintToysInfant
        Unload frmToys
        '22. print
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmPictures.SelectPicBirth Then
            PrintPicturesInfant
            Unload frmPictures
        End If
        '23. print - fathers notes infant
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmFathersNotesInfancy.SelectRecords Then
            PrintInfancyNotes
            Unload frmFathersNotesInfancy
        End If
    End If
    
    If Check5.Value = 1 Then
        'CHILDHOOD
        PrintBlockName (rsLanguage2.Fields("PrintChildhoodPage"))
        DoEvents
        '24
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmBirthDates.SelectDays Then
            PrintBirthdays
            Unload frmBirthDates
        End If
        '25
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        Call frmHealth.SelectHealthChild
        PrintHealthChild
        Unload frmHealth
        '26
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        frmBooks.SelectBooksChild
        PrintBooksChild
        Unload frmBooks
        '27
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        frmToys.SelectToys
        PrintToysChild
        Unload frmToys
        '28
        ProgressBar1.Value = ProgressBar1.Value + 3.4
        If frmPictures.SelectPicBirth Then
            PrintPicturesChild
            Unload frmPictures
        End If
        '29
        If ProgressBar1.Value < 93 Then
            ProgressBar1.Value = ProgressBar1.Value + 3.4
        End If
        If frmFathersNotesChild.SelectRecords Then
            PrintChildNotes
            Unload frmFathersNotesChild
        End If
    End If
        
    If ProgressBar1.Value < 96 Then
        ProgressBar1.Value = ProgressBar1.Value + 3.4
    End If
    'we are done, release the printer object
    Screen.MousePointer = vbDefault
    ProgressBar1.Visible = False
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    
    Set cPrint = Nothing
    Exit Sub
        
errPrintDefault:
    Beep
    MsgBox Err.Description, vbInformation, "Print with preview"
    Resume Next
End Sub


Private Sub PrintWithWord()
        On Error Resume Next
        iCounter = 3.57
        bFirstWrite = False
    
        PrepareWord
        ProgressBar1.Value = 0
        
        'PREGNANCY
        If Check1.Value = 1 Then
            WriteBlock (rsLanguage2.Fields("PrintPregnancyPage"))
            Call frmIamPregnant.SelectPregnancy
            WriteIamPregnantWord
            Unload frmIamPregnant
            DoEvents
            '2. print
            frmPregnancyControl.SelectControl
            WritePregnancyControlWord
            Unload frmPregnancyControl
            DoEvents
            '3. print
            Call frmPregnancyNotes.SelectNotes
            WritePregnancyNotesWord
            Unload frmPregnancyNotes
            DoEvents
            '4. print
            Call frmAntenatal.SelectAntenatalChild
            WriteAntenatalWord
            Unload frmAntenatal
            DoEvents
            '5. print
            Call frmFathersNotesPregnancy.SelectRecords
            WriteFatherPregnancyNotesWord
            Unload frmFathersNotesPregnancy
            DoEvents
        End If
        
        ' THE BIRTH
        If Check2.Value = 1 Then
            WriteBlock (rsLanguage2.Fields("PrintBirthPage"))
            Call frmBirth.ShowChild
            WriteBirthWord
            Unload frmBirth
            DoEvents
            '7. print
            With frmHospital
                .SelectHospitalAcquaintance
                .SelectHospitalHome
                .SelectHospitalLeaving
                .SelectHospitalNotes
            End With
            WriteHospitalWord
            Unload frmHospital
            DoEvents
            '8. print
            ProgressBar1.Value = ProgressBar1 + iCounter
            Call frmPictures.SelectPicBirth
            WritePicturesWord
            DoEvents
            Unload frmPictures
            '9. print
            If frmFirstPram.SelectPram Then
                WriteFirstPramWord
            End If
            DoEvents
            Unload frmFirstPram
            '10. print - when I was born
            frmWhenIWasBorn.SelectBorn
            WriteWhenIWasBorn
            DoEvents
            Unload frmWhenIWasBorn
            '11. print
            Call frmFathersNotesBirth.SelectRecords
            WriteBirthNotesWord
            DoEvents
            Unload frmFathersNotesBirth
            DoEvents
        End If
        
        'THE BABTISM
        If Check3.Value = 1 Then
            WriteBlock (rsLanguage2.Fields("PrintBabtismPage"))
            Call frmBaptism.SelectChild
            WriteBaptismWord
            DoEvents
            Unload frmBaptism
            '13. print
            frmBaptismPictures.FillList2
            WriteBaptismPicWord
            DoEvents
            Unload frmBaptismPictures
            '14. print
            frmFathersNotesBaptism.SelectRecords
            WriteBaptismNotesWord
            DoEvents
            Unload frmFathersNotesBaptism
            DoEvents
        End If
        
        'BABY
        If Check4.Value = 1 Then
            WriteBlock (rsLanguage2.Fields("PrintBabyPage"))
            Call frmWeightLength.SelectChild
            WriteWeightLengthWord
            DoEvents
            Unload frmWeightLength
            DoEvents
            '16. print
            Call frmTeeth.SelectChild
            WriteTeethWord
            DoEvents
            Unload frmTeeth
            DoEvents
            '17. print
            Call frmFirstTimes.ShowFirstTime
            WriteFirstTimesWord
            DoEvents
            Unload frmFirstTimes
            DoEvents
            '18. print
            Call frmFoodHabits.ReadFoodHabits
            WriteFoodHabitsWord
            DoEvents
            Unload frmFoodHabits
            DoEvents
            '19. print
            Call frmHealth.SelectHealthChild
            WriteHealthInfantWord
            DoEvents
            Unload frmHealth
            DoEvents
            '20. print - books
            frmBooks.SelectBooksChild
            WriteBooksInfantWord
            Unload frmBooks
            '21. print - toys
            frmToys.SelectToys
            WriteToysInfantWord
            Unload frmToys
            '22. print - baby pictures
            Call frmPictures.SelectPicBirth
            WritePicturesInfantWord
            DoEvents
            Unload frmPictures
            '23. print - fathers notes infant
            frmFathersNotesInfancy.SelectRecords
            WriteInfancyNotesWord
            Unload frmFathersNotesInfancy
        End If
        
        'CHILDHOOD
        If Check5.Value = 1 Then
            WriteBlock (rsLanguage2.Fields("PrintChildhoodPage"))
            frmBirthDates.SelectPictures
            WriteBirthdaysWord
            Unload frmBirthDates
            '25 print - health child
            Call frmHealth.SelectHealthChild
            WriteHealthChildWord
            DoEvents
            Unload frmHealth
            '26. print - books child
            frmBooks.SelectBooksChild
            WriteBooksChildWord
            Unload frmBooks
            '27. print - toys child
            frmToys.SelectToys
            WriteToysChildWord
            Unload frmToys
            '28. print -
            Call frmPictures.SelectPicBirth
            WritePicturesChildWord
            DoEvents
            Unload frmPictures
            '29. print - fathers notes
            frmFathersNotesChild.SelectRecords
            WriteChildNotesWord
            Unload frmFathersNotesChild
        End If
        
        'we are done, write index
        MakeIndex
        SaveMemoryBook
        Set wdApp = Nothing
End Sub


Private Sub WritePregnancyNotesWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPregnancyNotes")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        frmPregnancyNotes.rsPregnancyNotes.Recordset.MoveFirst
        Do While Not frmPregnancyNotes.rsPregnancyNotes.Recordset.EOF
            DoEvents
            .Selection.TypeText Text:=rsLanguage.Fields("Label1(1)")
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Format(CDate(frmPregnancyNotes.rsPregnancyNotes.Recordset.Fields("NoteDate")), "dd.mm.yyyy")
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            Clipboard.Clear
            Clipboard.SetText frmPregnancyNotes.RichText1.TextRTF, vbCFRTF
            .Selection.Paste
            DoEvents
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
            .Selection.MoveRight Unit:=wdCell
        frmPregnancyNotes.rsPregnancyNotes.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintPregnancyNotes()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmPregnancyNotes")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmPregnancyNotes.rsPregnancyNotes.Recordset
        .MoveFirst
        Do While Not .EOF
            cPrint.FontBold = True
            cPrint.pPrint rsLanguage.Fields("Label1(1)") & "  " & Format(CDate(.Fields("NoteDate")), "dd.mm.yyyy"), 1
            cPrint.FontBold = False
            If Len(frmPregnancyNotes.RichText1.Text) <> 0 Then
                cPrint.pMultiline frmPregnancyNotes.RichText1.Text, 1, cPrint.GetPaperWidth - 1.2, , False, True
            Else
                cPrint.pPrint " ", 3.5
            End If
            cPrint.pPrint
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                DoNewPagePreview
            End If
        .MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub WritePregnancyControlWord()
    Set rsLanguage = dbKidLang.OpenRecordset("frmPregnancyControl")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    frmPregnancyControl.rsPregnancyControl.Recordset.MoveFirst
    With wdApp
        Do While Not frmPregnancyControl.rsPregnancyControl.Recordset.EOF
            If CLng(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("ChildNo")) = glChildNo Then
                .Selection.TypeText Text:=rsLanguage.Fields("label1(0)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("ControlDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("MidwifeName"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(2)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("Results"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("MidwifeComments"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(4)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=frmPregnancyControl.rsPregnancyControl.Recordset.Fields("MyWeight") & " " & Format(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("MyWeightDim"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(5)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=frmPregnancyControl.rsPregnancyControl.Recordset.Fields("MyTommy") & " " & Format(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("MyTommyDim"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(6)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("Questions"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=rsLanguage.Fields("label1(7)")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(frmPregnancyControl.rsPregnancyControl.Recordset.Fields("OwnNotes"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        frmPregnancyControl.rsPregnancyControl.Recordset.MoveNext
        Loop
    End With
    wdApp.Selection.MoveDown Unit:=wdLine, Count:=1
    rsLanguage.Close
End Sub

Private Sub PrintPregnancyControl()
    Set rsLanguage = dbKidLang.OpenRecordset("frmPregnancyControl")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmPregnancyControl
        .rsPregnancyControl.Recordset.MoveFirst
        Do While Not .rsPregnancyControl.Recordset.EOF
            If CLng(.rsPregnancyControl.Recordset.Fields("ChildNo")) = glChildNo Then
                cPrint.FontBold = True
                cPrint.pPrint rsLanguage.Fields("label1(0)"), 1, True
                If IsDate(.rsPregnancyControl.Recordset.Fields("ControlDate")) Then
                    cPrint.pPrint Format(CDate(.rsPregnancyControl.Recordset.Fields("ControlDate")), "dd.mm.yyyy"), 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.FontBold = False
                cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True
                If Len(.cmbMidwife.Text) <> 0 Then
                    cPrint.pPrint .cmbMidwife.Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(2)"), 1, True
                If Len(.Text1(0).Text) <> 0 Then
                    cPrint.pMultiline .Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True
                If Len(.Text1(1).Text) <> 0 Then
                    cPrint.pMultiline .Text1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(4)"), 1, True
                If Len(.Text1(2).Text) <> 0 Then
                    cPrint.pPrint .Text1(2).Text & "  " & .cmbDim1(0).Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(5)"), 1, True
                If Len(.Text1(3).Text) <> 0 Then
                    cPrint.pPrint .Text1(3).Text & "  " & .cmbDim1(1).Text, 3.5
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(6)"), 1, True
                If Len(.Text1(4).Text) <> 0 Then
                    cPrint.pMultiline .Text1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
                cPrint.pPrint rsLanguage.Fields("label1(7)"), 1, True
                If Len(.Text1(5).Text) <> 0 Then
                    cPrint.pMultiline .Text1(5).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                Else
                    cPrint.pPrint " ", 3.5
                End If
                cPrint.pPrint
                cPrint.pPrint
                If cPrint.pEndOfPage Then
                    DoNewPagePreview
                End If
            End If
        .rsPregnancyControl.Recordset.MoveNext
        Loop
    End With
    rsLanguage.Close
End Sub

Private Sub WriteIamPregnantWord()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmIamPregnant")
    ReadLanguage
    
    Call WriteHeading(rsLanguage.Fields("FormName"))
    
    With wdApp
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(0)") & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=CDate(frmIamPregnant.rsIamPregnant.Recordset.Fields("ConfirmationDate")) & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(3)") & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmIamPregnant.rsIamPregnant.Recordset.Fields("TestTaken")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(2)") & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmIamPregnant.rsIamPregnant.Recordset.Fields("DoctorName") & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(4)") & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmIamPregnant.rsIamPregnant.Recordset.Fields("FirstSign")
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(5)") & ":"
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmIamPregnant.rsIamPregnant.Recordset.Fields("FirstPersonToKnow") & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(1)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=frmIamPregnant.rsIamPregnant.Recordset.Fields("FirstPersonReaction") & ""
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("Frame1(6)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmIamPregnant.rsIamPregnant.Recordset.Fields("LaterPersons"))
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=rsLanguage.Fields("label1(3)")
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Format(frmIamPregnant.rsIamPregnant.Recordset.Fields("LaterPersonsReactions"))
        .Selection.MoveDown Unit:=wdLine, Count:=1
    End With
    rsLanguage.Close
End Sub

Private Sub PrintIamPregnant()
    On Error Resume Next
    Set rsLanguage = dbKidLang.OpenRecordset("frmIamPregnant")
    ReadLanguage
    
    sHeader = rsLanguage.Fields("FormName")
    PrintHeaderComplete
    
    With frmIamPregnant
        cPrint.pPrint rsLanguage.Fields("Frame1(0)") & ":", 1, True 'confirmation date
        If IsDate(.rsIamPregnant.Recordset.Fields("ConfirmationDate")) Then
            cPrint.pPrint .rsIamPregnant.Recordset.Fields("ConfirmationDate"), 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(3)") & ":", 1, True 'first signs
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pPrint .Text1(2).Text, 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(1)") & ":", 1, True 'doctor / nurse name
        If Len(.cmbName.Text) <> 0 Then
            cPrint.pPrint .cmbName.Text, 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(3)") & ":", 1, True 'first signs
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pPrint .Text1(2).Text, 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(5)") & ":", 1, True 'first persons to know
        If Len(.Text1(3).Text) <> 0 Then
            cPrint.pPrint .Text1(3).Text, 3.5
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint rsLanguage.Fields("label1(1)"), 1, True      'reaction
        If Len(.Text1(4).Text) <> 0 Then
            cPrint.pMultiline .Text1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint
        cPrint.pPrint rsLanguage.Fields("Frame1(4)") & ":", 1, True 'later persons to know
        If Len(.Text1(5).Text) <> 0 Then
            cPrint.pMultiline .Text1(5).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint "", 3.5
        End If
        If cPrint.pEndOfPage Then
            DoNewPagePreview
        End If
        cPrint.pPrint rsLanguage.Fields("label1(3)"), 1, True
        If Len(.Text1(6).Text) <> 0 Then
            cPrint.pMultiline .Text1(6).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
        Else
            cPrint.pPrint "", 3.5
        End If
    End With
    rsLanguage.Close
End Sub

Private Sub ReadLanguage()
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then Exit Do
        .MoveNext
        Loop
    End With
End Sub

Private Sub btnNext_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        With rsClipArt.Recordset
            If Not .EOF Then
                .MoveNext
            End If
        End With
    Case 1
        With rsClipArt2.Recordset
            If Not .EOF Then
                .MoveNext
            End If
        End With
    Case Else
    End Select
End Sub

Private Sub btnPrevious_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        With rsClipArt.Recordset
            If Not .BOF Then
                .MovePrevious
            End If
        End With
    Case 1
        With rsClipArt2.Recordset
            If Not .BOF Then
                .MovePrevious
            End If
        End With
    Case Else
    End Select
End Sub

Private Sub btnPrint_Click()
    'write selected pic-ID's
    On Error Resume Next
    With rsMyRecord
        .Edit
        .Fields("FrontPicID") = CLng(rsClipArt.Recordset.Fields("LineNo"))
        .Fields("SectionPicID") = CLng(rsClipArt2.Recordset.Fields("LineNo"))
        .Update
    End With
    'start the printing
    If Option1(0).Value = True Then
        PrintWithWord
    Else
        PrintDefaultPrinter
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsClipArt.Refresh
    rsClipArt2.Refresh
    ShowText
    If CBool(rsMyRecord.Fields("PrintUsingWord")) Then
        Option1(0).Value = True
    Else
        Option1(1).Value = True
    End If
    
    ShowPictures

    MDIMasterKid.cmbChildren.Enabled = True
    MDIMasterKid.Label1.Enabled = True
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    rsClipArt.DatabaseName = dbKidPicTxt
    rsClipArt2.DatabaseName = dbKidPicTxt
    Set rsLanguage2 = dbKidLang.OpenRecordset("frmPrint")
    iWhichForm = 42
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRecord.Close
    rsClipArt.Recordset.Close
    rsClipArt2.Recordset.Close
    rsLanguage2.Close
    iWhichForm = 0
    MDIMasterKid.cmbChildren.Enabled = False
    Set frmPrint = Nothing
End Sub

