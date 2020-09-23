VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIMasterKid 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Child follow-Up"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10260
   Icon            =   "MDIMasterKid.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMasterKid.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMasterKid.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMasterKid.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMasterKid.frx":08D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMasterKid.frx":0A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMasterKid.frx":0E84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.CtlVerticalMenu Menu1 
      Align           =   3  'Align Left
      Height          =   7515
      Left            =   0
      TabIndex        =   4
      Top             =   495
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   13256
      MenuCaption1    =   "Menu1"
      MenuItemIcon11  =   "MDIMasterKid.frx":0FDE
      BackColor       =   -2147483644
      MenuForeColor   =   0
      MenuItemForeColor=   0
   End
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   2280
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10260
      TabIndex        =   1
      Top             =   0
      Width           =   10260
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Exit"
               Object.ToolTipText     =   "Exit"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   0,8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print.."
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete record"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Mail"
               Object.ToolTipText     =   "Email"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Help"
               Object.ToolTipText     =   "Help"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmbChildren 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   5640
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Child:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8010
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Picture         =   "MDIMasterKid.frx":12F8
            TextSave        =   "30.09.2002"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Picture         =   "MDIMasterKid.frx":1454
            TextSave        =   "12:45"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPregnancy 
      Caption         =   "&Pregnancy"
      Begin VB.Menu mnuIamPregnant 
         Caption         =   "I am Pregnant"
      End
      Begin VB.Menu mnuPregnancyControl 
         Caption         =   "Pregnancy Control"
      End
      Begin VB.Menu mnuPregnancyNotes 
         Caption         =   "Pregnancy Notes"
      End
      Begin VB.Menu mnuAntenatal 
         Caption         =   "Antenatal Classes"
      End
      Begin VB.Menu mnuTerm 
         Caption         =   "Term"
      End
      Begin VB.Menu mnuPregToRemember 
         Caption         =   "To Remember"
      End
      Begin VB.Menu mnuPregFaNotes 
         Caption         =   "Fathers Notes"
      End
   End
   Begin VB.Menu mnuBirth 
      Caption         =   "&Birth"
      Begin VB.Menu mnuBirthRem 
         Caption         =   "Remember .."
      End
      Begin VB.Menu mnuBirthThe 
         Caption         =   "The Birth"
      End
      Begin VB.Menu mnuBirthDairy 
         Caption         =   "Birth Diary"
      End
      Begin VB.Menu mnuBirthAcq 
         Caption         =   "Hospital Acquaintance"
      End
      Begin VB.Menu mnuBirthLeaving 
         Caption         =   "Leaving Hospital"
      End
      Begin VB.Menu mnuHome 
         Caption         =   "Home"
      End
      Begin VB.Menu mnuBirthPic 
         Caption         =   "Pictures"
      End
      Begin VB.Menu mnuBirthSound 
         Caption         =   "Sound"
      End
      Begin VB.Menu mnuBirthVideo 
         Caption         =   "Video"
      End
      Begin VB.Menu mnuBirthPram 
         Caption         =   "My first Pram"
      End
      Begin VB.Menu mnuBirthWhenBorne 
         Caption         =   "When I Was Born"
      End
      Begin VB.Menu mnuBirthFaNotes 
         Caption         =   "Fathers Notes"
      End
   End
   Begin VB.Menu mnuBaptism 
      Caption         =   "&Babtism"
      Begin VB.Menu mnuBapNames 
         Caption         =   "Names"
      End
      Begin VB.Menu mnuBapChrist 
         Caption         =   "The Christening"
      End
      Begin VB.Menu mnuBapGodmother 
         Caption         =   "Godmothers"
      End
      Begin VB.Menu mnuBapPic 
         Caption         =   "Pictures"
      End
      Begin VB.Menu mnuBapVideo 
         Caption         =   "Video"
      End
      Begin VB.Menu mnuBapFaNotes 
         Caption         =   "Fathers Notes"
      End
   End
   Begin VB.Menu mnuInfant 
      Caption         =   "&Infancy "
      Begin VB.Menu mnuInfWeight 
         Caption         =   "Weight / Height"
      End
      Begin VB.Menu mnuTeeth 
         Caption         =   "Teeth"
      End
      Begin VB.Menu mnuInfFirst 
         Caption         =   "First time .."
      End
      Begin VB.Menu mnuInfFood 
         Caption         =   "Food Habits"
      End
      Begin VB.Menu mnuInfHealth 
         Caption         =   "Health"
      End
      Begin VB.Menu mnuInfBooks 
         Caption         =   "Books"
      End
      Begin VB.Menu mnuInfToys 
         Caption         =   "Toys"
      End
      Begin VB.Menu mnuInfPic 
         Caption         =   "Pictures"
      End
      Begin VB.Menu mnuInfSound 
         Caption         =   "Sound"
      End
      Begin VB.Menu mnuInfVideo 
         Caption         =   "Video"
      End
      Begin VB.Menu mnuInfFaNotes 
         Caption         =   "Fathers Notes"
      End
   End
   Begin VB.Menu mnuChildhood 
      Caption         =   "&Childhood "
      Begin VB.Menu mnuChildBirthDays 
         Caption         =   "Birthdays"
      End
      Begin VB.Menu mnuChildHealth 
         Caption         =   "Health"
      End
      Begin VB.Menu mnuChildBooks 
         Caption         =   "Books"
      End
      Begin VB.Menu mnuChildToys 
         Caption         =   "Toys"
      End
      Begin VB.Menu mnuChildPic 
         Caption         =   "Pictures"
      End
      Begin VB.Menu mnuChildSound 
         Caption         =   "Sound"
      End
      Begin VB.Menu mnuChildVideo 
         Caption         =   "Video"
      End
      Begin VB.Menu mnuChildFaNotes 
         Caption         =   "Fathers Notes"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuChildren 
         Caption         =   "Children"
      End
      Begin VB.Menu mnuMidwife 
         Caption         =   "Midwife"
      End
      Begin VB.Menu mnuHoroscope 
         Caption         =   "Horoscope"
      End
      Begin VB.Menu mnuPrintCompleteFrames 
         Caption         =   "Word Frame Pictures"
      End
      Begin VB.Menu space9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDatabaseCompact 
         Caption         =   "Compact / back-Up"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
      Begin VB.Menu mnuPrintComplete 
         Caption         =   "Print Complete Memory Book"
      End
      Begin VB.Menu space10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSetUp 
         Caption         =   "Print Set-Up"
      End
      Begin VB.Menu mnuPrintWord 
         Caption         =   "Use Word"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccessScan 
         Caption         =   "Access Scanner"
      End
   End
   Begin VB.Menu mnuInternet 
      Caption         =   "&Internet"
      Begin VB.Menu mnuInternetConnect 
         Caption         =   "Connect"
         Index           =   0
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInternetBase 
         Caption         =   "Database"
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "&Language"
      Begin VB.Menu mnuNames 
         Caption         =   "Names"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Colors"
      End
      Begin VB.Menu mnuDimension 
         Caption         =   "Dimensions"
      End
      Begin VB.Menu space11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCountry 
         Caption         =   "Change Language"
      End
      Begin VB.Menu mnuScreenText 
         Caption         =   "Screen Text"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbove 
         Caption         =   "About...."
      End
      Begin VB.Menu mnuMailDeveloper 
         Caption         =   "Mail to system responsible"
      End
      Begin VB.Menu mnuRegistration 
         Caption         =   "Registration"
      End
      Begin VB.Menu mnuRegistrationIn 
         Caption         =   "Input Registration No"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Download Update"
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUser1 
         Caption         =   "User"
         Begin VB.Menu mnuUser 
            Caption         =   "User"
         End
         Begin VB.Menu mnuSupplier 
            Caption         =   "Programme Supplier"
         End
      End
      Begin VB.Menu mnuHelpFile 
         Caption         =   "Help"
         Begin VB.Menu mnuHelpFileWord 
            Caption         =   "Word Help File"
         End
         Begin VB.Menu mnuHelpFileHtml 
            Caption         =   "HTML Help File"
         End
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Visible         =   0   'False
      Begin VB.Menu mnuFormatBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuFormatItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuFormatUnderline 
         Caption         =   "Underline"
      End
      Begin VB.Menu space5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuformatStrikeLine 
         Caption         =   "Strike a line"
      End
      Begin VB.Menu space6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatLeft 
         Caption         =   "Left justify"
      End
      Begin VB.Menu mnuFormatMid 
         Caption         =   "Mid Justify"
      End
      Begin VB.Menu mnuFormatRight 
         Caption         =   "Right Justify"
      End
      Begin VB.Menu space7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuFormatPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuFormatCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu space8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatFontS 
         Caption         =   "Fonts"
         Begin VB.Menu mnuFormatFont 
            Caption         =   "xx"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFormatFontSizes 
         Caption         =   "Font Size"
         Begin VB.Menu mnuFormatFontSize 
            Caption         =   "xx"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFormatColors 
         Caption         =   "Colors"
         Begin VB.Menu mnuFormatColor 
            Caption         =   "xx"
            Index           =   0
         End
      End
      Begin VB.Menu space12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatSpell 
         Caption         =   "Spelling"
      End
   End
End
Attribute VB_Name = "MDIMasterKid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFirst As Boolean, v1RecordBookmark() As Variant
Dim hSysMenu As Long      ' This is the system menu's handle.
Dim zMENU As MENUITEMINFO ' This is a structure that is used for the modification of the
                                                            ' system menu.
Dim bookColor() As Variant
Dim rsMyRecord As Recordset
Dim rsInternet As Recordset
Dim rsChildren As Recordset
Dim rsColor As Recordset
Dim rsLanguage As Recordset
Private Sub DeleteRecords()
Dim DgDef, Msg, Response, Title
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    On Error Resume Next
    Title = "DELETE RECORD"
    Msg = rsLanguage.Fields("Msg1")
    Response = MsgBox(Msg, DgDef, Title)
    If Response = IDYES Then
        Select Case iWhichForm
        Case 1  'pregnancynotes
            frmPregnancyNotes.rsPregnancyNotes.Recordset.Delete
            frmPregnancyNotes.FillList2
        Case 2  'I am pregnant
            frmIamPregnant.rsIamPregnant.Recordset.Delete
        Case 3  'children
            frmKids.rsChildren.Recordset.Delete
            LoadChildren
            cmbChildren.ListIndex = 0
            frmKids.SelectChild
        Case 4
            frmMidWife.rsMidwife.Recordset.Delete
            frmMidWife.FillList1
        Case 5  'pregnancy control
            frmPregnancyControl.rsPregnancyControl.Recordset.Delete
            frmPregnancyControl.FillList1
        Case 6  'antenatal
            frmAntenatal.rsAntenatal.Recordset.Delete
            frmAntenatal.FillList1
        Case 10 'fathers pregnancy notes
            frmFathersNotesPregnancy.rsPregnancyNotes.Recordset.Delete
            frmFathersNotesPregnancy.FillList2
        Case 12
            frmFathersNotesBaptism.rsBaptismNotes.Recordset.Delete
            frmFathersNotesBaptism.FillList2
        Case 13 'fathers notes infancy
            frmFathersNotesInfancy.rsInfancyNotes.Recordset.Delete
        Case 14 'fathers noteschild
            frmFathersNotesChild.rsChildNotes.Recordset.Delete
        Case 15 'birth
            frmBirth.rsBirth.Recordset.Delete
            frmBirth.ShowChild
        Case 17 'birthdates
            frmBirthDates.DeleteBirthDay
        Case 18 'hospital
            frmHospital.DeleteHospital
        Case 19 'first times
            frmFirstTimes.rsFirstTime.Recordset.Delete
        Case 20 'baptism
            frmBaptism.rsBaptism.Recordset.Delete
            frmBaptism.SelectChild
        Case 21 'books
            frmBooks.DeleteBooks
        Case 22 'toys
            frmToys.DeleteToys
        Case 23 'health
            frmHealth.DeleteHealth
        Case 25 'food habits
            frmFoodHabits.rsFoodHabits.Recordset.Delete
        Case 26 'to remember
            frmRemember.rsToRemember.Recordset.Delete
            frmRemember.FillList1
        Case 27 'fathers notes birth
            frmFathersNotesBirth.rsBirthNotes.Recordset.Delete
        Case 28 'baptism pictures
            frmBaptismPictures.DeleteRecord
        Case 29
            frmPictures.DeletePictures
        Case 30 'video
            frmVideo.DeleteVideo
        Case 32 'my first pram
            frmFirstPram.rsFirstPram.Recordset.Delete
        Case 33 'sound
            frmSound.DeleteSound
        Case 37 'teeth
            With frmTeeth
                .rsTeeth.Recordset.Delete
                If .SelectChild Then
                    .Label5.Caption = MDIMasterKid.cmbChildren.Text
                Else
                    .Label5.Caption = " "
                End If
            End With
        Case 38 'horoscope
            frmHoroscope.rsHoroscope.Recordset.Delete
            frmHoroscope.FillList1
        Case 41 'weight & length
            frmWeightLength.rsWeightLength.Recordset.Delete
            frmWeightLength.SelectChild
        Case 43 'when I was born
            frmWhenIWasBorn.DeleteRecord
        Case 45 'word frames
            frmFrames.DeleteClipArt
        Case Else
        End Select
    End If
End Sub

Public Sub LoadChildren()
    On Error Resume Next
    cmbChildren.Clear
    With rsChildren
        .MoveLast
        .MoveFirst
        ReDim v1RecordBookmark(.RecordCount)
        Do While Not .EOF
            If Not IsNull(.Fields("ChildFirstName")) Then
                cmbChildren.AddItem .Fields("ChildFirstName")
            Else
                If Not IsNull(.Fields("ChildCallingName")) Then
                    cmbChildren.AddItem .Fields("ChildCallingName")
                Else
                    cmbChildren.AddItem "?"
                End If
            End If
            cmbChildren.ItemData(cmbChildren.NewIndex) = cmbChildren.ListCount - 1
            v1RecordBookmark(cmbChildren.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub LoadColors()
Dim iCount As Integer
    On Error Resume Next
    iCount = 0
    With rsColor
        .MoveLast
        .MoveFirst
        ReDim bookColor(.RecordCount)
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                Me.mnuFormatColor(0).Caption = .Fields("ColorText")
                bookColor(iCount) = .Bookmark
                Exit Do
            End If
        .MoveNext
        Loop
        
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                iCount = iCount + 1
                Load Me.mnuFormatColor(iCount)
                Me.mnuFormatColor(iCount).Caption = .Fields("ColorText")
                bookColor(iCount) = .Bookmark
            End If
        .MoveNext
        Loop
    End With
End Sub

Public Sub LoadMenu1()
    On Error Resume Next
    With Menu1
        .MenusMax = 6
        .MenuCur = 1    'pregnancy
        .MenuItemsMax = 7
        .MenuCaption = rsLanguage.Fields("Preg")
        .MenuItemCur = 1    'I am pregnant
        Set .MenuItemIcon = LoadResPicture(117, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("IAmPreg")
        .MenuItemCur = 2    'pregnancy control
        Set .MenuItemIcon = LoadResPicture(109, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("PregCon")
        .MenuItemCur = 3    'pregnancy note
        Set .MenuItemIcon = LoadResPicture(108, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("PregNote")
        .MenuItemCur = 4    'pregnancy classes
        Set .MenuItemIcon = LoadResPicture(118, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Ante")
        .MenuItemCur = 5   'pregnancy term
        Set .MenuItemIcon = LoadResPicture(114, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Term")
        .MenuItemCur = 6   'to remember
        Set .MenuItemIcon = LoadResPicture(104, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("PregRem")
        .MenuItemCur = 7  'Fathers notes
        Set .MenuItemIcon = LoadResPicture(107, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("PregFathNote")
        
        .MenuCur = 2    'birth
        .MenuCaption = rsLanguage.Fields("Birth")
        .MenuItemsMax = 9
        .MenuItemCur = 1  'to rember
        Set .MenuItemIcon = LoadResPicture(104, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("BirthRem")
        .MenuItemCur = 2  'the Birth
        Set .MenuItemIcon = LoadResPicture(116, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("TheBirth")
        .MenuItemCur = 3 'from the Hospital
        Set .MenuItemIcon = LoadResPicture(103, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Diary")
        .MenuItemCur = 4  'picturs
        Set .MenuItemIcon = LoadResPicture(101, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("BirthPic")
        .MenuItemCur = 5  'sounds
        Set .MenuItemIcon = LoadResPicture(105, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Sound")
        .MenuItemCur = 6  'video
        Set .MenuItemIcon = LoadResPicture(102, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Video")
        .MenuItemCur = 7  'first Pram
        Set .MenuItemIcon = LoadResPicture(127, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Pram")
        .MenuItemCur = 8  'when I was born...
        Set .MenuItemIcon = LoadResPicture(106, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Born")
        .MenuItemCur = 9  'Fathers notes
        Set .MenuItemIcon = LoadResPicture(107, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("BirthNote")
        
        .MenuCur = 3    'baptism
        .MenuCaption = rsLanguage.Fields("Baptism")
        .MenuItemsMax = 5
        .MenuItemCur = 1  'names
        Set .MenuItemIcon = LoadResPicture(110, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Names")
        .MenuItemCur = 2  'baptism
        Set .MenuItemIcon = LoadResPicture(115, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Christ")
        .MenuItemCur = 3  'pictures
        Set .MenuItemIcon = LoadResPicture(101, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("BapPic")
        .MenuItemCur = 4  'video
        Set .MenuItemIcon = LoadResPicture(102, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Video")
        .MenuItemCur = 5 'fathers notes
        Set .MenuItemIcon = LoadResPicture(107, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("BapNote")
        
        .MenuCur = 4    'Infancy
        .MenuCaption = rsLanguage.Fields("Infancy")
        .MenuItemsMax = 11
        .MenuItemCur = 1  'weight/length
        Set .MenuItemIcon = LoadResPicture(121, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("LengthWeight")
        .MenuItemCur = 2  'teeth
        Set .MenuItemIcon = LoadResPicture(120, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Teeth")
        .MenuItemCur = 3  'first time
        Set .MenuItemIcon = LoadResPicture(125, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("First")
        .MenuItemCur = 4  'food
        Set .MenuItemIcon = LoadResPicture(113, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("InfFood")
        .MenuItemCur = 5  'health
        Set .MenuItemIcon = LoadResPicture(103, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("InfHealth")
        .MenuItemCur = 6  'books
        Set .MenuItemIcon = LoadResPicture(111, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("InfBooks")
        .MenuItemCur = 7  'toys
        Set .MenuItemIcon = LoadResPicture(112, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("InfToys")
        .MenuItemCur = 8  'pictures
        Set .MenuItemIcon = LoadResPicture(101, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("InfPic")
        .MenuItemCur = 9  'sound
        Set .MenuItemIcon = LoadResPicture(104, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Sound")
        .MenuItemCur = 10  'video
        Set .MenuItemIcon = LoadResPicture(102, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Video")
        .MenuItemCur = 11 'fathers notes
        Set .MenuItemIcon = LoadResPicture(107, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("InfNote")
        
        .MenuCur = 5    'childhood
        .MenuCaption = rsLanguage.Fields("Childhood")
        .MenuItemsMax = 8
        .MenuItemCur = 1 'birthdays
        Set .MenuItemIcon = LoadResPicture(119, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("ChildBirth")
        .MenuItemCur = 2 'health
        Set .MenuItemIcon = LoadResPicture(103, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("ChildHealth")
        .MenuItemCur = 3 'books
        Set .MenuItemIcon = LoadResPicture(111, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("ChildBooks")
        .MenuItemCur = 4 'toys
        Set .MenuItemIcon = LoadResPicture(112, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("ChildToys")
        .MenuItemCur = 5 'pictures
        Set .MenuItemIcon = LoadResPicture(101, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("ChildPic")
        .MenuItemCur = 6 'sounds
        Set .MenuItemIcon = LoadResPicture(105, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Sound")
        .MenuItemCur = 7 'video
        Set .MenuItemIcon = LoadResPicture(102, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("Video")
        .MenuItemCur = 8 'fathers notes
        Set .MenuItemIcon = LoadResPicture(107, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("ChildNote")
        
        .MenuCur = 6    'database
        .MenuCaption = rsLanguage.Fields("Database")
        .MenuItemsMax = 7
        .MenuItemCur = 1 'children
        Set .MenuItemIcon = LoadResPicture(122, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("DataChild")
        .MenuItemCur = 2 'midwife
        Set .MenuItemIcon = LoadResPicture(117, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("DataMidw")
        .MenuItemCur = 3 'horoscope
        Set .MenuItemIcon = LoadResPicture(121, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("DataHoro")
        .MenuItemCur = 4 'names
        Set .MenuItemIcon = LoadResPicture(110, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("DataNames")
        .MenuItemCur = 5 'dimensions
        Set .MenuItemIcon = LoadResPicture(126, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("DataDimension")
        .MenuItemCur = 6 'colors
        Set .MenuItemIcon = LoadResPicture(123, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("DataColor")
        .MenuItemCur = 7 'internet
        Set .MenuItemIcon = LoadResPicture(124, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("DataInternet")
        
        .MenuCur = 1
    End With
End Sub
Private Sub MakeNewMenu1()
    On Error Resume Next
    With rsLanguage
        .Fields("Preg") = "Pregnancy"
        .Fields("IAmPreg") = "I am Pregnant"
        .Fields("PregCon") = "Pregnancy Control"
        .Fields("PregNote") = "Pregnancy Notes"
        .Fields("Ante") = "Pregnancy Classes"
        .Fields("Term") = "Term"
        .Fields("PregRem") = "To remember"
        .Fields("PregFathNote") = "Fathers Notes"
        .Fields("Birth") = "Birth"
        .Fields("BirthRem") = "To remember"
        .Fields("TheBirth") = "The Birth"
        .Fields("Diary") = "From the Hospital"
        .Fields("BirthPic") = "Pictures"
        .Fields("Sound") = "Sounds"
        .Fields("Video") = "Video"
        .Fields("Pram") = "First Pram"
        .Fields("Born") = "When I was born..."
        .Fields("BirthNote") = "Fathers Notes"
        .Fields("Baptism") = "Baptism"
        .Fields("Names") = "Names"
        .Fields("Christ") = "The Christening"
        .Fields("BapPic") = "Pictures"
        .Fields("Video") = "Video"
        .Fields("BapNote") = "Fathers Notes"
        .Fields("Infancy") = "Infancy"
        .Fields("LengthWeight") = "Length/Weight"
        .Fields("Teeth") = "The Teeth"
        .Fields("First") = "The first time .."
        .Fields("InfFood") = "Food"
        .Fields("InfHealth") = "Health"
        .Fields("InfBooks") = "Books"
        .Fields("InfToys") = "Toys"
        .Fields("InfPic") = "Pictures"
        .Fields("Sound") = "Sounds"
        .Fields("Video") = "Video"
        .Fields("InfNote") = "Fathers Notes"
        .Fields("Childhood") = "Childhood"
        .Fields("ChildBirth") = "Birthdays"
        .Fields("ChildHealth") = "Health"
        .Fields("ChildBooks") = "Books"
        .Fields("ChildToys") = "Toys"
        .Fields("ChildPic") = "Pictures"
        .Fields("Sound") = "Sounds"
        .Fields("Video") = "Video"
        .Fields("ChildNote") = "Fathers Notes"
        .Fields("Database") = "Database"
        .Fields("DataChild") = "Our Children"
        .Fields("DataMidw") = "Midwife"
        .Fields("DataHoro") = "Horoscope"
    End With
End Sub

Private Sub NewRecords()
    Select Case iWhichForm
    Case 1  'pregnancynotes
        frmPregnancyNotes.NewPregnancyNotes
    Case 2  'I am pregnant
        frmIamPregnant.NewPregnancy
    Case 3  'children
        frmKids.NewChild
    Case 4  'MidWife
        frmMidWife.NewMidwife
    Case 5  'pregnancy controls
        frmPregnancyControl.NewPregnancyControl
    Case 6  'antenatal
        frmAntenatal.NewAntenatal
    Case 10 'fathers pregnancy notes
        frmFathersNotesPregnancy.NewPregnancyNotesFather
    Case 12 'fathers birth notes
        frmFathersNotesBaptism.NewBaptismNotesFather
    Case 13 'fathers notes infancy
        frmFathersNotesInfancy.NewInfancy
    Case 14 'fathers notes childhood
        frmFathersNotesChild.NewFathersNotesChild
    Case 15 'birth
        frmBirth.NewBirth
    Case 17 'birthdates
        frmBirthDates.NewBirthDates
    Case 18 'hospital
        frmHospital.NewHospital
    Case 19 'first time...
        frmFirstTimes.NewFirstTime
    Case 20
        frmBaptism.NewBaptism
    Case 21 'books
        frmBooks.NewBooks
    Case 22 'toys
        frmToys.NewToys
    Case 23 'health
        frmHealth.NewHealth
    Case 25 'food habits
        frmFoodHabits.NewFoodHabit
    Case 26 'to remember
        frmRemember.NewRemember
    Case 27
        frmFathersNotesBirth.NewFatherBirth
    Case 28 'baptism pictures
        frmBaptismPictures.NewBapPic
    Case 29
        frmPictures.NewPictures
    Case 30 'new video
        frmVideo.NewVideo
    Case 32 'my first Pram
        frmFirstPram.NewPram
    Case 33 'sound
        frmSound.NewSound
    Case 37 'teeth
        frmTeeth.NewTeeth
    Case 38 'horoscope
        frmHoroscope.NewHoroscope
    Case 41 'weight & length
        frmWeightLength.NewWeightLength
    Case 43 'when I was born
        frmWhenIWasBorn.NewBorn
    Case 45 'word frames
        frmFrames.NewFrames
    Case Else
    End Select
End Sub

Public Sub ShowMenu()
    On Error Resume Next
    'files
    mnuFiles.Caption = rsLanguage.Fields("mnuFiles")
    mnuLanguage.Caption = rsLanguage.Fields("mnuLanguage")
    mnuDimension.Caption = rsLanguage.Fields("mnuDimension")
    mnuNames.Caption = rsLanguage.Fields("mnuNames")
    mnuColors.Caption = rsLanguage.Fields("mnuColors")
    mnuCountry.Caption = rsLanguage.Fields("mnuCountry")
    mnuScreenText.Caption = rsLanguage.Fields("mnuScreenText")
    mnuUser1.Caption = rsLanguage.Fields("mnuUser1")
    mnuUser.Caption = rsLanguage.Fields("mnuUser")
    mnuRegistration.Caption = rsLanguage.Fields("mnuRegistration")
    mnuRegistrationIn.Caption = rsLanguage.Fields("mnuRegistrationIn")
    mnuUpdate.Caption = rsLanguage.Fields("mnuUpdate")
    mnuSupplier.Caption = rsLanguage.Fields("mnuSupplier")
    mnuPrintWord.Caption = rsLanguage.Fields("mnuPrintWord")
    mnuPrint.Caption = rsLanguage.Fields("mnuPrint")
    mnuPrintComplete.Caption = rsLanguage.Fields("mnuPrintComplete")
    mnuPrintCompleteFrames.Caption = rsLanguage.Fields("mnuPrintCompleteFrames")
    mnuPrintSetUp.Caption = rsLanguage.Fields("mnuPrintSetUp")
    mnuAccessScan.Caption = rsLanguage.Fields("mnuAccessScan")
    mnuExit.Caption = rsLanguage.Fields("mnuExit")
    'pregnancy
    mnuPregnancy.Caption = rsLanguage.Fields("Preg")
    mnuIamPregnant.Caption = rsLanguage.Fields("IAmPreg")
    mnuPregnancyControl.Caption = rsLanguage.Fields("PregCon")
    mnuPregnancyNotes.Caption = rsLanguage.Fields("PregNote")
    mnuAntenatal.Caption = rsLanguage.Fields("Ante")
    mnuPregToRemember.Caption = rsLanguage.Fields("PregRem")
    mnuTerm.Caption = rsLanguage.Fields("Term")
    mnuPregFaNotes.Caption = rsLanguage.Fields("PregFathNote")
    'birth
    mnuBirth.Caption = rsLanguage.Fields("Birth")
    mnuBirthThe.Caption = rsLanguage.Fields("TheBirth")
    mnuBirthRem.Caption = rsLanguage.Fields("BirthRem")
    mnuBirthAcq.Caption = rsLanguage.Fields("Acq")
    mnuBirthDairy.Caption = rsLanguage.Fields("Diary")
    mnuBirthLeaving.Caption = rsLanguage.Fields("Leaving")
    mnuHome.Caption = rsLanguage.Fields("Home")
    mnuBirthPram.Caption = rsLanguage.Fields("Pram")
    mnuBirthWhenBorne.Caption = rsLanguage.Fields("Born")
    mnuBirthPic.Caption = rsLanguage.Fields("BirthPic")
    mnuBirthSound.Caption = rsLanguage.Fields("Sound")
    mnuBirthVideo.Caption = rsLanguage.Fields("Video")
    mnuBirthFaNotes.Caption = rsLanguage.Fields("BirthNote")
    'baptism
    mnuBaptism.Caption = rsLanguage.Fields("Baptism")
    mnuBapNames.Caption = rsLanguage.Fields("Names")
    mnuBapChrist.Caption = rsLanguage.Fields("Christ")
    mnuBapGodmother.Caption = rsLanguage.Fields("Godmother")
    mnuBapPic.Caption = rsLanguage.Fields("BapPic")
    mnuBapVideo.Caption = rsLanguage.Fields("Video")
    mnuBapFaNotes.Caption = rsLanguage.Fields("BapNote")
    'infant
    mnuInfant.Caption = rsLanguage.Fields("Infancy")
    mnuInfWeight.Caption = rsLanguage.Fields("LengthWeight")
    mnuTeeth.Caption = rsLanguage.Fields("Teeth")
    mnuInfFirst.Caption = rsLanguage.Fields("First")
    mnuInfHealth.Caption = rsLanguage.Fields("InfHealth")
    mnuInfFood.Caption = rsLanguage.Fields("InfFood")
    mnuInfBooks.Caption = rsLanguage.Fields("InfBooks")
    mnuInfToys.Caption = rsLanguage.Fields("InfToys")
    mnuInfPic.Caption = rsLanguage.Fields("InfPic")
    mnuInfSound.Caption = rsLanguage.Fields("Sound")
    mnuInfVideo.Caption = rsLanguage.Fields("Video")
    mnuInfFaNotes.Caption = rsLanguage.Fields("InfNote")
    'child
    mnuChildhood.Caption = rsLanguage.Fields("Childhood")
    mnuChildBirthDays.Caption = rsLanguage.Fields("ChildBirth")
    mnuChildHealth.Caption = rsLanguage.Fields("ChildHealth")
    mnuChildBooks.Caption = rsLanguage.Fields("ChildBooks")
    mnuChildToys.Caption = rsLanguage.Fields("ChildToys")
    mnuChildPic.Caption = rsLanguage.Fields("ChildPic")
    mnuChildSound.Caption = rsLanguage.Fields("Sound")
    mnuChildVideo.Caption = rsLanguage.Fields("Video")
    mnuBirthWhenBorne = rsLanguage.Fields("Born")
    mnuChildFaNotes.Caption = rsLanguage.Fields("ChildNote")
    'database
    mnuDatabase.Caption = rsLanguage.Fields("Database")
    mnuChildren.Caption = rsLanguage.Fields("DataChild")
    mnuMidwife.Caption = rsLanguage.Fields("DataMidw")
    mnuHoroscope.Caption = rsLanguage.Fields("DataHoro")
    mnuDatabaseCompact.Caption = rsLanguage.Fields("mnuDatabaseCompact")
    'internet
    mnuInternet.Caption = rsLanguage.Fields("mnuInternet")
    mnuInternetBase.Caption = rsLanguage.Fields("mnuInternetBase")
    mnuInternetConnect(0).Caption = rsLanguage.Fields("mnuInternetConnect")
    'help
    mnuHelp.Caption = rsLanguage.Fields("mnuHelp")
    mnuAbove.Caption = rsLanguage.Fields("mnuAbove")
    mnuMailDeveloper.Caption = rsLanguage.Fields("mnuMailDeveloper")
    mnuHelpFile.Caption = rsLanguage.Fields("mnuHelpFile")
    mnuHelpFileWord.Caption = rsLanguage.Fields("mnuHelpFileWord")
    mnuHelpFileHtml.Caption = rsLanguage.Fields("mnuHelpFileHtml")
    'format
    mnuFormatBold.Caption = rsLanguage.Fields("mnuFormatBold")
    mnuFormatItalic.Caption = rsLanguage.Fields("mnuFormatItalic")
    mnuFormatUnderline.Caption = rsLanguage.Fields("mnuFormatUnderline")
    mnuformatStrikeLine.Caption = rsLanguage.Fields("mnuformatStrikeLine")
    mnuFormatLeft.Caption = rsLanguage.Fields("mnuFormatLeft")
    mnuFormatMid.Caption = rsLanguage.Fields("mnuFormatMid")
    mnuFormatRight.Caption = rsLanguage.Fields("mnuFormatRight")
    mnuFormatCopy.Caption = rsLanguage.Fields("mnuFormatCopy")
    mnuFormatPaste.Caption = rsLanguage.Fields("mnuFormatPaste")
    mnuFormatCut.Caption = rsLanguage.Fields("mnuFormatCut")
    mnuFormatFontS.Caption = rsLanguage.Fields("mnuFormatFontS")
    mnuFormatFontSizes.Caption = rsLanguage.Fields("mnuFormatFontSizes")
    mnuFormatColors.Caption = rsLanguage.Fields("mnuFormatColors")
    mnuFormatSpell.Caption = rsLanguage.Fields("mnuFormatSpell")
    'labels
    Label1.Caption = rsLanguage.Fields("Label1")
End Sub

Private Sub WriteNewMenu()
    On Error Resume Next
    With rsLanguage
        'files
        .Fields("mnuFiles") = mnuFiles.Caption
        'language
        .Fields("mnuLanguage") = mnuLanguage.Caption
        .Fields("mnuDimension") = mnuDimension.Caption
        .Fields("mnuNames") = mnuNames.Caption
        .Fields("mnuColors") = mnuColors.Caption
        .Fields("mnuCountry") = mnuCountry.Caption
        .Fields("mnuScreenText") = mnuScreenText.Caption
        'user
        .Fields("mnuUser1") = mnuUser1.Caption
        .Fields("mnuUser") = mnuUser.Caption
        .Fields("mnuUpdate") = mnuUpdate.Caption
        .Fields("mnuRegistration") = mnuRegistration.Caption
        .Fields("mnuRegistrationIn") = mnuRegistrationIn.Caption
        .Fields("mnuSupplier") = mnuSupplier.Caption
        'print
        .Fields("mnuPrint") = mnuPrint.Caption
        .Fields("mnuPrintComplete") = mnuPrintComplete.Caption
        .Fields("mnuPrintCompleteFrames") = mnuPrintCompleteFrames.Caption
        .Fields("mnuPrintSetUp") = mnuPrintSetUp.Caption
        .Fields("mnuPrintWord") = mnuPrintWord.Caption
        .Fields("mnuAccessScan") = mnuAccessScan.Caption
        .Fields("mnuExit") = mnuExit.Caption
        'pregnancy
        .Fields("Preg") = mnuPregnancy.Caption
        .Fields("IAmPreg") = mnuIamPregnant.Caption
        .Fields("PregCon") = mnuPregnancyControl.Caption
        .Fields("PregNote") = mnuPregnancyNotes.Caption
        .Fields("Ante") = mnuAntenatal.Caption
        .Fields("PregRem") = mnuPregToRemember.Caption
        .Fields("Term") = mnuTerm.Caption
        .Fields("PregFathNote") = mnuPregFaNotes.Caption
        'birth
        .Fields("Birth") = mnuBirth.Caption
        .Fields("TheBirth") = mnuBirthThe.Caption
        .Fields("BirthRem") = mnuBirthRem.Caption
        .Fields("Acq") = mnuBirthAcq.Caption
        .Fields("Diary") = mnuBirthDairy.Caption
        .Fields("Leaving") = mnuBirthLeaving.Caption
        .Fields("Home") = mnuHome.Caption
        .Fields("Pram") = mnuBirthPram.Caption
        .Fields("BirthPic") = mnuBirthPic.Caption
        .Fields("Sound") = mnuBirthSound.Caption
        .Fields("Video") = mnuBirthVideo.Caption
        .Fields("BirthNote") = mnuBirthFaNotes.Caption
        'baptism
        .Fields("Baptism") = mnuBaptism.Caption
        .Fields("Names") = mnuBapNames.Caption
        .Fields("Christ") = mnuBapChrist.Caption
        .Fields("Godmother") = mnuBapGodmother.Caption
        .Fields("BapPic") = mnuBapPic.Caption
        .Fields("BapNote") = mnuBapFaNotes.Caption
        'infant
        .Fields("Infancy") = mnuInfant.Caption
        .Fields("LengthWeight") = mnuInfWeight.Caption
        .Fields("Teeth") = mnuTeeth.Caption
        .Fields("First") = mnuInfFirst.Caption
        .Fields("InfHealth") = mnuInfHealth.Caption
        .Fields("InfFood") = mnuInfFood.Caption
        .Fields("InfBooks") = mnuInfBooks.Caption
        .Fields("InfToys") = mnuInfToys.Caption
        .Fields("InfPic") = mnuInfPic.Caption
        .Fields("InfNote") = mnuInfFaNotes.Caption
        'child
        .Fields("Childhood") = mnuChildhood.Caption
        .Fields("ChildBirth") = mnuChildBirthDays.Caption
        .Fields("ChildHealth") = mnuChildHealth.Caption
        .Fields("ChildBooks") = mnuChildBooks.Caption
        .Fields("ChildToys") = mnuChildToys.Caption
        .Fields("ChildPic") = mnuChildPic.Caption
        .Fields("Born") = mnuBirthWhenBorne.Caption
        .Fields("ChildNote") = mnuChildFaNotes.Caption
        'database
        .Fields("Database") = mnuDatabase.Caption
        .Fields("DataChild") = mnuChildren.Caption
        .Fields("DataMidw") = mnuMidwife.Caption
        .Fields("DataHoro") = mnuHoroscope.Caption
        .Fields("mnuDatabaseCompact") = mnuDatabaseCompact.Caption
        'internet
        .Fields("mnuInternet") = mnuInternet.Caption
        .Fields("mnuInternetBase") = mnuInternetBase.Caption
        .Fields("mnuInternetConnect") = mnuInternetConnect(0).Caption
        'help
        .Fields("mnuHelp") = mnuHelp.Caption
        .Fields("mnuAbove") = mnuAbove.Caption
        .Fields("mnuMailDeveloper") = mnuMailDeveloper.Caption
        .Fields("mnuHelpFile") = mnuHelpFile.Caption
        .Fields("mnuHelpFileWord") = mnuHelpFileWord.Caption
        .Fields("mnuHelpFileHtml") = mnuHelpFileHtml.Caption
        'format
        .Fields("mnuFormatBold") = mnuFormatBold.Caption
        .Fields("mnuFormatItalic") = mnuFormatItalic.Caption
        .Fields("mnuFormatUnderline") = mnuFormatUnderline.Caption
        .Fields("mnuformatStrikeLine") = mnuformatStrikeLine.Caption
        .Fields("mnuFormatLeft") = mnuFormatLeft.Caption
        .Fields("mnuFormatMid") = mnuFormatMid.Caption
        .Fields("mnuFormatRight") = mnuFormatRight.Caption
        .Fields("mnuFormatCopy") = mnuFormatCopy.Caption
        .Fields("mnuFormatPaste") = mnuFormatPaste.Caption
        .Fields("mnuFormatCut") = mnuFormatCut.Caption
        .Fields("mnuFormatFontS") = mnuFormatFontS.Caption
        .Fields("mnuFormatFontSizes") = mnuFormatFontSizes.Caption
        .Fields("mnuFormatColors") = mnuFormatColors.Caption
        .Fields("mnuFormatSpell") = mnuFormatSpell.Caption
        'labels
        .Fields("label1") = Label1.Caption
    End With
End Sub

Private Sub ShowEmail()
    On Error Resume Next
    With frmEmail
        Select Case iWhichForm
        Case 4 'midwife mail
            If Len(frmMidWife.Text1(10).Text) <> 0 Then
                .Text1(0).Text = Trim(frmMidWife.Text1(10).Text)
                .Show 1
            End If
        Case 7  'user
            If Len(frmUser.Text1(13).Text) <> 0 Then
                .Text1(0).Text = Trim(frmUser.Text1(13).Text)
                .Show 1
            End If
        Case Else
        End Select
    End With
End Sub

Private Sub ShowHelp()
    On Error Resume Next
    Select Case iWhichForm
    Case 0  'mdi
        frmHelp.Label1.Caption = "MDIMasterKid"
    Case 1  'frmPregnancyNotes
        frmHelp.Label1.Caption = "frmPregnancyNotes"
    Case 2  'frmIamPregnant
        frmHelp.Label1.Caption = "frmIamPregnant"
    Case 3  'frmkids
        frmHelp.Label1.Caption = "frmKids"
    Case 4  'frmMidWife
        frmHelp.Label1.Caption = "frmMidWife"
    Case 5  'frmpregnancyControl
        frmHelp.Label1.Caption = "frmPregnancyControl"
    Case 6  'antenatalClasses
        frmHelp.Label1.Caption = "frmAntenatal"
    Case 7  'user information
        frmHelp.Label1.Caption = "frmUser"
    Case 8  'Names
        frmHelp.Label1.Caption = "frmNames"
    Case 9  'Term
        frmHelp.Label1.Caption = "frmTerm"
    Case 10 'Fathers Notes Pregnancy
        frmHelp.Label1.Caption = "frmFathersNotesPregnancy"
    Case 11 'country
        frmHelp.Label1.Caption = "frmCountry"
    Case 12 'Fathers Notes Babtism
        frmHelp.Label1.Caption = "frmFathersNotesBaptism"
    Case 13 'Fathers Notes Infancy
        frmHelp.Label1.Caption = "frmFathersNotesInfancy"
    Case 14 'Fathers Notes Child
        frmHelp.Label1.Caption = "frmFathersNotesChild"
    Case 15 'Birth
        frmHelp.Label1.Caption = "frmBirth"
    Case 16 'dimensions
        frmHelp.Label1.Caption = "frmDimensions"
    Case 17 'birthdays
        frmHelp.Label1.Caption = "frmBirthDates"
    Case 18 'hospital
        frmHelp.Label1.Caption = "frmHospital"
    Case 19 'first time...
        frmHelp.Label1.Caption = "frmFirstTimes"
    Case 20 'baptism
        frmHelp.Label1.Caption = "frmBaptism"
    Case 21 'books
        frmHelp.Label1.Caption = "frmBooks"
    Case 22 'Toys
        frmHelp.Label1.Caption = "frmToys"
    Case 23
        frmHelp.Label1.Caption = "frmHealth"
    Case 24
        frmHelp.Label1.Caption = "frmEmail"
    Case 25 'food habits
        frmHelp.Label1.Caption = "frmFoodHabits"
    Case 26 'to remember
        frmHelp.Label1.Caption = "frmRemember"
    Case 27 'fathers notes birth
        frmHelp.Label1.Caption = "frmFathersNotesBirth"
    Case 28 'babtism pictures
        frmHelp.Label1.Caption = "frmBaptismPictures"
    Case 29 'pictures
        frmHelp.Label1.Caption = "frmPictures"
    Case 30 'video
        frmHelp.Label1.Caption = "frmVideo"
    Case 31 'internet database
        frmHelp.Label1.Caption = "frmInternet"
    Case 32 'first pram
        frmHelp.Label1.Caption = "frmFirstPram"
    Case 33 'Sounds
        frmHelp.Label1.Caption = "frmSound"
    Case 34 'write to me
        frmHelp.Label1.Caption = "frmWriteToMe"
    'Case 35
        'frmHelp.Label1.Caption = "frmFirst"
    Case 36 'registration
        frmHelp.Label1.Caption = "frmRegistration"
    Case 37 'teeth
        frmHelp.Label1.Caption = "frmTeeth"
    Case 38 'horoscope
        frmHelp.Label1.Caption = "frmHoroscope"
    'Case 39 'Zodiac sign
        'frmHelp.Label1.Caption = "frmZodiac"
    Case 40 'supplier
        frmHelp.Label1.Caption = "frmSupplier"
    Case 41
        frmHelp.Label1.Caption = "frmWeightLength"
    Case 42
        frmHelp.Label1.Caption = "frmPrint"
    Case 43 'when I was born
        frmHelp.Label1.Caption = "frmWhenIWasBorn"
    Case 44 'regsitration
        frmHelp.Label1.Caption = "frmRegistrateProgramme"
    Case 45 'wordframes
        frmHelp.Label1.Caption = "frmFrames"
    Case Else
        Exit Sub
    End Select
    frmHelp.Show
End Sub

Private Sub cmbChildren_Click()
    On Error Resume Next
    rsChildren.Bookmark = v1RecordBookmark(cmbChildren.ItemData(cmbChildren.ListIndex))
    glChildNo = CLng(rsChildren.Fields("ChildNo"))
    gsChildName = cmbChildren.List(cmbChildren.ListIndex)
    Select Case iWhichForm
        Case 1  'pregnancy notes
            With frmPregnancyNotes
                If .SelectNotes Then
                    .FillList2
                    .List2.ListIndex = 0
                Else
                    .List2.Clear
                End If
                .Frame1.Caption = gsChildName
            End With
        Case 2  'I am pregnant
            frmIamPregnant.SelectPregnancy
        Case 3  'children
            frmKids.SelectChild
        Case 5  'pregnancy control
            With frmPregnancyControl
                .SelectControl
                .FillList1
                .List1.ListIndex = 0
            End With
        Case 6  'antenatal
            With frmAntenatal
                .SelectAntenatalChild
                .FillList1
            End With
        Case 10 'fathers notes pregnancy
            With frmFathersNotesPregnancy
                If .SelectRecords Then
                    .FillList2
                    .List2.ListIndex = 0
                Else
                    .List2.Clear
                End If
                .Frame1.Caption = gsChildName
            End With
        Case 12 'Fathers Notes Baptism
            With frmFathersNotesBaptism
                If .SelectRecords Then
                    .FillList2
                    .List2.ListIndex = 0
                    .Frame1.Caption = gsChildName
                Else
                    .List2.Clear
                End If
            End With
        Case 13 'fathers notes infant
            With frmFathersNotesInfancy
                If .SelectRecords Then
                    .FillList2
                    .List2.ListIndex = 0
                Else
                    .List2.Clear
                End If
                .Frame1.Caption = gsChildName
            End With
        Case 14 'fathers nothes child
            With frmFathersNotesChild
                If .SelectRecords Then
                    .FillList2
                    .List2.ListIndex = 0
                Else
                    .List2.Clear
                End If
                .Frame1.Caption = gsChildName
            End With
        Case 15 'birth
            With frmBirth
                .Frame2.Caption = gsChildName
                .ShowChild
            End With
        Case 17 'birth days
            frmBirthDates.NewChildBirthDates
        Case 18 'hosital
            With frmHospital
                If .SelectHospitalNotes Then
                    .SelectHospitalAcquaintance
                    .SelectHospitalLeaving
                    .SelectHospitalHome
                    Toolbar1.Buttons(6).Enabled = False
                Else
                    Toolbar1.Buttons(6).Enabled = True
                End If
            End With
        Case 19 'first times
            frmFirstTimes.ShowFirstTime
        Case 20 'baptism
            With frmBaptism
                .SelectChild
                .Label2(0).Caption = gsChildName
                .Label2(1).Caption = gsChildName
                .Label2(2).Caption = gsChildName
            End With
        Case 21 'books
            frmBooks.NewChildBooks
        Case 22 'toys
            frmToys.NewChildToys
        Case 23 'health
            frmHealth.SelectHealthChild
        Case 25 'food habits
            frmFoodHabits.ReadFoodHabits
        Case 26 'to remember
            With frmRemember
                .FillList1
                .List1.ListIndex = 0
            End With
        Case 27 'fathers nothes birth
            With frmFathersNotesBirth
                If .SelectRecords Then
                    .FillList2
                    .List2.ListIndex = 0
                Else
                    .List2.Clear
                End If
                .Frame1.Caption = gsChildName
            End With
        Case 28 'baptism pictures
            With frmBaptismPictures
                If .FillList2 Then
                    .List2.ListIndex = 0
                End If
                .Frame1.Caption = gsChildName
            End With
        Case 29 'pictures
            With frmPictures
                Select Case .Tab1.Tab
                Case 0
                    If .SelectPicBirth Then
                        .FillList20
                        .List2(0).ListIndex = 0
                    Else
                        .List2(0).Clear
                    End If
                Case 1
                    If .SelectPicInfant Then
                        .FillList21
                        .List2(1).ListIndex = 0
                    Else
                        .List2(1).Clear
                    End If
                Case 2
                    If .SelectPicChild Then
                        .FillList22
                        .List2(2).ListIndex = 0
                    Else
                        .List2(2).Clear
                    End If
                Case Else
                End Select
            End With
        Case 30 'video
            frmVideo.NewChildVideo
        Case 32 'first pram
            With frmFirstPram
                .ShowPram
                .Frame1.Caption = gsChildName
            End With
        Case 33 'sound
            With frmSound
                Select Case .Tab1.Tab
                Case 0
                    If .SelectSoundBirth Then
                        .FillList20
                        .List2(0).ListIndex = 0
                    End If
                Case 1
                    If .SelectSoundBaby Then
                        .FillList21
                        .List2(1).ListIndex = 0
                    End If
                Case 2
                    If .SelectSoundChild Then
                        .FillList22
                        .List2(2).ListIndex = 0
                    End If
                Case Else
                End Select
            End With
        Case 37 'teeth
            If frmTeeth.SelectChild Then
                frmTeeth.Label5.Caption = MDIMasterKid.cmbChildren.Text
            Else
                frmTeeth.Label5.Caption = " "
            End If
        Case 41 'Weight & length
            If frmWeightLength.SelectChild Then
                frmWeightLength.Label6.Caption = MDIMasterKid.cmbChildren.Text
            Else
                frmWeightLength.Label6.Caption = " "
            End If
        Case 43 'when I was born
            frmWhenIWasBorn.SelectBorn
        Case Else
    End Select
End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    Me.Caption = Me.Caption & " - " & "Version.: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Menu1_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    On Error GoTo errMenu1
    CloseActiveForm
    Select Case MenuNumber
    Case 1  'pregnancy
        Select Case MenuItem
            Case 1
                frmIamPregnant.Show
            Case 2
                frmPregnancyControl.Show
            Case 3
                frmPregnancyNotes.Show
            Case 4
                frmAntenatal.Show
            Case 5
                frmTerm.Show
            Case 6
                frmRemember.Show
            Case 7
                frmFathersNotesPregnancy.Show
            Case Else
            End Select
    Case 2  'birth
        Select Case MenuItem
            Case 1
                frmRemember.Show
            Case 2
                frmBirth.Show
            Case 3
                frmHospital.Show
            Case 4
                iTab = 0
                frmPictures.Show
            Case 5
                iTab = 0
                frmSound.Show
            Case 6
                iTab = 0
                frmVideo.Show
            Case 7
                frmFirstPram.Show
            Case 8
                frmWhenIWasBorn.Show
            Case 9
                frmFathersNotesBirth.Show
            Case Else
            End Select
    Case 3  'baptism
        Select Case MenuItem
            Case 1
                frmNames.Show
            Case 2
                frmBaptism.Show
            Case 3
                frmBaptismPictures.Show
            Case 4
                iTab = 1
                frmVideo.Show
            Case 5
                frmFathersNotesBaptism.Show
            Case Else
            End Select
    Case 4  'infancy
        Select Case MenuItem
            Case 1
                frmWeightLength.Show
            Case 2
                frmTeeth.Show
            Case 3
                frmFirstTimes.Show
            Case 4
                frmFoodHabits.Show
            Case 5
                iTab = 0
                frmHealth.Show
            Case 6
                iTab = 0
                frmBooks.Show
            Case 7
                iTab = 0
                frmToys.Show
            Case 8
                iTab = 1
                frmPictures.Show
            Case 9
                iTab = 1
                frmSound.Show
            Case 10
                iTab = 2
                frmVideo.Show
            Case 11
                frmFathersNotesInfancy.Show
            Case Else
            End Select
    Case 5  'childhood
        Select Case MenuItem
            Case 1
                frmBirthDates.Show
            Case 2
                iTab = 1
                frmHealth.Show
            Case 3
                iTab = 1
                frmBooks.Show
            Case 4
                iTab = 1
                frmToys.Show
            Case 5
                iTab = 2
                frmPictures.Show
            Case 6
                frmSound.Show
            Case 7
                iTab = 3
                frmVideo.Show
            Case 8
                frmFathersNotesChild.Show
            Case Else
            End Select
    Case 6  'database
        Select Case MenuItem
            Case 1
                frmKids.Show
            Case 2
                frmMidWife.Show
            Case 3
                frmHoroscope.Show
            Case 4
                frmAllNameExplanation.Show
            Case 5
                frmDimensions.Show
            Case 6
                frmColor.Show
            Case 7
                frmInternet.Show
            Case Else
            End Select
    Case Else
    End Select
    Exit Sub
    
errMenu1:
    Beep
    MsgBox Err.Description, vbCritical, "Menu"
    Err.Clear
End Sub

Private Sub mnuBirthPram_Click()
    CloseActiveForm
    frmFirstPram.Show
End Sub

Private Sub mnuBirthSound_Click()
    On Error Resume Next
    CloseActiveForm
    iTab = 0
    frmSound.Show
End Sub

Private Sub mnuBirthWhenBorne_Click()
    CloseActiveForm
    frmWhenIWasBorn.Show
End Sub

Private Sub mnuChildHealth_Click()
    CloseActiveForm
    CloseActiveForm
    iTab = 1
    frmHealth.Show
End Sub

Private Sub mnuChildSound_Click()
    On Error Resume Next
    CloseActiveForm
    iTab = 2
    frmSound.Show
End Sub

Private Sub mnuColors_Click()
    CloseActiveForm
    frmColor.Show
End Sub

Private Sub mnuDatabaseCompact_Click()
Dim iret As Long
    On Error Resume Next
    rsMyRecord.Close
    rsInternet.Close
    rsChildren.Close
    rsLanguage.Close
    dbKids.Close
    dbKidLang.Close
    dbKidPic.Close
    DoEvents
    iret = ShellExceCute(Me.hWnd, "open", App.Path & "\MasterKidUpdate.exe", vbNullString, vbNullString, SW_SHOWNORMAL)
    Unload Me
End Sub

Private Sub mnuFormatBold_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnLetterClick(0, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnLetterClick(0, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnLetterClick(0, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnLetterClick(0, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnLetterClick(0, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnLetterClick(0, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnLetterClick(0, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnLetterClick(0, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnLetterClick(0, frmHospital.RichTextBox1(0))
        Case 1
            Call btnLetterClick(0, frmHospital.RichTextBox1(1))
        Case 2
            Call btnLetterClick(0, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnLetterClick(0, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnLetterClick(0, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnLetterClick(0, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnLetterClick(0, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnLetterClick(0, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnLetterClick(0, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnLetterClick(0, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnLetterClick(0, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnLetterClick(0, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnLetterClick(0, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnLetterClick(0, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnLetterClick(0, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnLetterClick(0, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnLetterClick(0, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnLetterClick(0, frmEmail.RichText1)
    Case 25 'food habits
        Call btnLetterClick(0, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnLetterClick(0, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnLetterClick(0, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call btnLetterClick(0, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnLetterClick(0, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatColor_Click(Index As Integer)
    With rsColor
        .Bookmark = bookColor(Index)
        lRed = CLng(.Fields("RedValue"))
        lGreen = CLng(.Fields("GreenValue"))
        lBlue = CLng(.Fields("BlueValue"))
    End With
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call formatColor(frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call formatColor(frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call formatColor(frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call formatColor(frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call formatColor(frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call formatColor(frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call formatColor(frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call formatColor(frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call formatColor(frmHospital.RichTextBox1(0))
        Case 1
            Call formatColor(frmHospital.RichTextBox1(1))
        Case 2
            Call formatColor(frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call formatColor(frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call formatColor(frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call formatColor(frmBaptism.RichTextBox3)
        Case 4  'notes
            Call formatColor(frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call formatColor(frmBooks.RichTextBox1(0))
        Case 1  'child
            Call formatColor(frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call formatColor(frmToys.RichTextBox1(0))
        Case 1  'child
            Call formatColor(frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call formatColor(frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call formatColor(frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call formatColor(frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call formatColor(frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call formatColor(frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call formatColor(frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call formatColor(frmEmail.RichText1)
    Case 25 'food habits
        Call formatColor(frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call formatColor(frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call formatColor(frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call formatColor(frmFirst.RichTextBox1)
    Case 37 'teeth
        Call formatColor(frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatCopy_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnClipboardClick(0, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnClipboardClick(0, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnClipboardClick(0, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnClipboardClick(0, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnClipboardClick(0, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnClipboardClick(0, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnClipboardClick(0, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnClipboardClick(0, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnClipboardClick(0, frmHospital.RichTextBox1(0))
        Case 1
            Call btnClipboardClick(0, frmHospital.RichTextBox1(1))
        Case 2
            Call btnClipboardClick(0, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnClipboardClick(0, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnClipboardClick(0, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnClipboardClick(0, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnClipboardClick(0, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnClipboardClick(0, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnClipboardClick(0, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnClipboardClick(0, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnClipboardClick(0, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnClipboardClick(0, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnClipboardClick(0, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnClipboardClick(0, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnClipboardClick(0, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnClipboardClick(0, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnClipboardClick(0, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnClipboardClick(0, frmEmail.RichText1)
    Case 25 'food habits
        Call btnClipboardClick(0, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnClipboardClick(0, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnClipboardClick(0, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call btnClipboardClick(0, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnClipboardClick(0, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatCut_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnClipboardClick(2, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnClipboardClick(2, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnClipboardClick(2, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnClipboardClick(2, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnClipboardClick(2, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnClipboardClick(2, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnClipboardClick(2, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnClipboardClick(2, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnClipboardClick(2, frmHospital.RichTextBox1(0))
        Case 1
            Call btnClipboardClick(2, frmHospital.RichTextBox1(1))
        Case 2
            Call btnClipboardClick(2, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnClipboardClick(2, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnClipboardClick(2, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnClipboardClick(2, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnClipboardClick(2, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnClipboardClick(2, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnClipboardClick(2, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnClipboardClick(2, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnClipboardClick(2, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnClipboardClick(2, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnClipboardClick(2, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnClipboardClick(2, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnClipboardClick(2, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnClipboardClick(2, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnClipboardClick(2, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnClipboardClick(2, frmEmail.RichText1)
    Case 25 'food habits
        Call btnClipboardClick(2, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnClipboardClick(2, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnClipboardClick(2, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call btnClipboardClick(2, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnClipboardClick(2, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatFont_Click(Index As Integer)
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call FontPopUp(mnuFormatFont(Index).Caption, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call FontPopUp(mnuFormatFont(Index).Caption, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call FontPopUp(mnuFormatFont(Index).Caption, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call FontPopUp(mnuFormatFont(Index).Caption, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call FontPopUp(mnuFormatFont(Index).Caption, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call FontPopUp(mnuFormatFont(Index).Caption, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call FontPopUp(mnuFormatFont(Index).Caption, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call FontPopUp(mnuFormatFont(Index).Caption, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call FontPopUp(mnuFormatFont(Index).Caption, frmHospital.RichTextBox1(0))
        Case 1
            Call FontPopUp(mnuFormatFont(Index).Caption, frmHospital.RichTextBox1(1))
        Case 2
            Call FontPopUp(mnuFormatFont(Index).Caption, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call FontPopUp(mnuFormatFont(Index).Caption, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call FontPopUp(mnuFormatFont(Index).Caption, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call FontPopUp(mnuFormatFont(Index).Caption, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call FontPopUp(mnuFormatFont(Index).Caption, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call FontPopUp(mnuFormatFont(Index).Caption, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call FontPopUp(mnuFormatFont(Index).Caption, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call FontPopUp(mnuFormatFont(Index).Caption, frmToys.RichTextBox1(0))
        Case 1  'child
            Call FontPopUp(mnuFormatFont(Index).Caption, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call FontPopUp(mnuFormatFont(Index).Caption, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call FontPopUp(mnuFormatFont(Index).Caption, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call FontPopUp(mnuFormatFont(Index).Caption, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call FontPopUp(mnuFormatFont(Index).Caption, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call FontPopUp(mnuFormatFont(Index).Caption, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call FontPopUp(mnuFormatFont(Index).Caption, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call FontPopUp(mnuFormatFont(Index).Caption, frmEmail.RichText1)
    Case 25 'food habits
        Call FontPopUp(mnuFormatFont(Index).Caption, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call FontPopUp(mnuFormatFont(Index).Caption, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call FontPopUp(mnuFormatFont(Index).Caption, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call FontPopUp(mnuFormatFont(Index).Caption, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call FontPopUp(mnuFormatFont(Index).Caption, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatFontSize_Click(Index As Integer)
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHospital.RichTextBox1(0))
        Case 1
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHospital.RichTextBox1(1))
        Case 2
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBaptism.RichTextBox3)
        Case 4  'notes
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBooks.RichTextBox1(0))
        Case 1  'child
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmToys.RichTextBox1(0))
        Case 1  'child
            Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmEmail.RichText1)
    Case 25 'food habits
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmFirst.RichTextBox1)
    Case 37 'teeth
        Call FontSizePopUp(CInt(mnuFormatFontSize(Index).Caption), frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatItalic_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnLetterClick(1, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnLetterClick(1, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnLetterClick(1, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnLetterClick(1, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnLetterClick(1, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnLetterClick(1, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnLetterClick(1, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnLetterClick(1, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnLetterClick(1, frmHospital.RichTextBox1(0))
        Case 1
            Call btnLetterClick(1, frmHospital.RichTextBox1(1))
        Case 2
            Call btnLetterClick(1, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnLetterClick(1, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnLetterClick(1, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnLetterClick(1, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnLetterClick(1, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnLetterClick(1, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnLetterClick(1, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnLetterClick(1, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnLetterClick(1, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnLetterClick(1, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnLetterClick(1, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnLetterClick(1, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnLetterClick(1, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnLetterClick(1, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnLetterClick(1, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnLetterClick(1, frmEmail.RichText1)
    Case 25 'food habits
        Call btnLetterClick(1, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnLetterClick(1, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnLetterClick(1, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call btnLetterClick(1, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnLetterClick(1, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatLeft_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnJustifyClick(0, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnJustifyClick(0, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnJustifyClick(0, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnJustifyClick(0, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnJustifyClick(0, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnJustifyClick(0, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnJustifyClick(0, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnJustifyClick(0, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnJustifyClick(0, frmHospital.RichTextBox1(0))
        Case 1
            Call btnJustifyClick(0, frmHospital.RichTextBox1(1))
        Case 2
            Call btnJustifyClick(0, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnJustifyClick(0, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnJustifyClick(0, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnJustifyClick(0, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnJustifyClick(0, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnJustifyClick(0, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnJustifyClick(0, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnJustifyClick(0, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnJustifyClick(0, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnJustifyClick(0, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnJustifyClick(0, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnJustifyClick(0, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnJustifyClick(0, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnJustifyClick(0, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnJustifyClick(0, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnJustifyClick(0, frmEmail.RichText1)
    Case 25 'food habits
        Call btnJustifyClick(0, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnJustifyClick(0, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnJustifyClick(0, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call btnJustifyClick(0, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnJustifyClick(0, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatMid_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnJustifyClick(1, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnJustifyClick(1, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnJustifyClick(1, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnJustifyClick(1, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnJustifyClick(1, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnJustifyClick(1, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnJustifyClick(1, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnJustifyClick(1, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnJustifyClick(1, frmHospital.RichTextBox1(0))
        Case 1
            Call btnJustifyClick(1, frmHospital.RichTextBox1(1))
        Case 2
            Call btnJustifyClick(1, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnJustifyClick(1, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnJustifyClick(1, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnJustifyClick(1, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnJustifyClick(1, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnJustifyClick(1, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnJustifyClick(1, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnJustifyClick(1, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnJustifyClick(1, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnJustifyClick(1, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnJustifyClick(1, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnJustifyClick(1, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnJustifyClick(1, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnJustifyClick(1, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnJustifyClick(1, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnJustifyClick(1, frmEmail.RichText1)
    Case 25 'food habits
        Call btnJustifyClick(1, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnJustifyClick(1, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnJustifyClick(1, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call btnJustifyClick(1, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnJustifyClick(1, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatPaste_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnClipboardClick(1, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnClipboardClick(1, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnClipboardClick(1, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnClipboardClick(1, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnClipboardClick(1, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnClipboardClick(1, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnClipboardClick(1, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnClipboardClick(1, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnClipboardClick(1, frmHospital.RichTextBox1(0))
        Case 1
            Call btnClipboardClick(1, frmHospital.RichTextBox1(1))
        Case 2
            Call btnClipboardClick(1, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnClipboardClick(1, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnClipboardClick(1, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnClipboardClick(1, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnClipboardClick(1, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnClipboardClick(1, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnClipboardClick(1, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnClipboardClick(1, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnClipboardClick(1, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnClipboardClick(1, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnClipboardClick(1, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnClipboardClick(1, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnClipboardClick(1, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnClipboardClick(1, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnClipboardClick(1, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnClipboardClick(1, frmEmail.RichText1)
    Case 25 'food habits
        Call btnClipboardClick(1, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnClipboardClick(1, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnClipboardClick(1, frmWriteToMe.RichTextBox1)
    Case 35
        Call btnClipboardClick(1, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnClipboardClick(1, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatRight_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnJustifyClick(2, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnJustifyClick(2, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnJustifyClick(2, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnJustifyClick(2, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnJustifyClick(2, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnJustifyClick(2, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnJustifyClick(2, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnJustifyClick(2, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnJustifyClick(2, frmHospital.RichTextBox1(0))
        Case 1
            Call btnJustifyClick(2, frmHospital.RichTextBox1(1))
        Case 2
            Call btnJustifyClick(2, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnJustifyClick(2, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnJustifyClick(2, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnJustifyClick(2, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnJustifyClick(2, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnJustifyClick(2, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnJustifyClick(2, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnJustifyClick(2, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnJustifyClick(2, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnJustifyClick(2, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnJustifyClick(2, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnJustifyClick(2, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnJustifyClick(2, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnJustifyClick(2, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnJustifyClick(2, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnJustifyClick(2, frmEmail.RichText1)
    Case 25 'food habits
        Call btnJustifyClick(2, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnJustifyClick(2, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnJustifyClick(2, frmWriteToMe.RichTextBox1)
    Case 35
        Call btnJustifyClick(2, frmFirst.RichTextBox1)
    Case 37 'teeth
        Call btnJustifyClick(2, frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuFormatSpell_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call CheckSpelling(frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call CheckSpelling(frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call CheckSpelling(frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call CheckSpelling(frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call CheckSpelling(frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call CheckSpelling(frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call CheckSpelling(frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call CheckSpelling(frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call CheckSpelling(frmHospital.RichTextBox1(0))
        Case 1
            Call CheckSpelling(frmHospital.RichTextBox1(1))
        Case 2
            Call CheckSpelling(frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call CheckSpelling(frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call CheckSpelling(frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call CheckSpelling(frmBaptism.RichTextBox3)
        Case 4  'notes
            Call CheckSpelling(frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call CheckSpelling(frmBooks.RichTextBox1(0))
        Case 1  'child
            Call CheckSpelling(frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call CheckSpelling(frmToys.RichTextBox1(0))
        Case 1  'child
            Call CheckSpelling(frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call CheckSpelling(frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call CheckSpelling(frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call CheckSpelling(frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call CheckSpelling(frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call CheckSpelling(frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call CheckSpelling(frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call CheckSpelling(frmEmail.RichText1)
    Case 25 'food habits
        Call CheckSpelling(frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call CheckSpelling(frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call CheckSpelling(frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call CheckSpelling(frmFirst.RichTextBox1)
    Case 37 'teeth
        Call CheckSpelling(frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuformatStrikeLine_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call StrikeLine(frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call StrikeLine(frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call StrikeLine(frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call StrikeLine(frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call StrikeLine(frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call StrikeLine(frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call StrikeLine(frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call StrikeLine(frmBirthDates.RichTextBox1)
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call StrikeLine(frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call StrikeLine(frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call StrikeLine(frmBaptism.RichTextBox3)
        Case 4  'notes
            Call StrikeLine(frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call StrikeLine(frmBooks.RichTextBox1(0))
        Case 1  'child
            Call StrikeLine(frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call StrikeLine(frmToys.RichTextBox1(0))
        Case 1  'child
            Call StrikeLine(frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 24 'email
        Call StrikeLine(frmEmail.RichText1)
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call StrikeLine(frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call StrikeLine(frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call StrikeLine(frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call StrikeLine(frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call StrikeLine(frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call StrikeLine(frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 25 'food habits
        Call StrikeLine(frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call StrikeLine(frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call StrikeLine(frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call StrikeLine(frmFirst.RichTextBox1)
    Case 37 'teeth
        Call StrikeLine(frmTeeth.RichTextBox1)
    Case Else
    End Select
End Sub
Private Sub mnuFormatUnderline_Click()
    On Error Resume Next
    Select Case iWhichForm
    Case 1  'pregnancy notes
        Call btnLetterClick(2, frmPregnancyNotes.RichText1)
    Case 6  'antenatal classes
        Call btnLetterClick(2, frmAntenatal.RichTextBox1)
    Case 10 'fathers notes pregnancy
        Call btnLetterClick(2, frmFathersNotesPregnancy.RichText1)
    Case 12 'fathers notes baptism
        Call btnLetterClick(2, frmFathersNotesBaptism.RichText1)
    Case 13 'fathers notes infant
        Call btnLetterClick(2, frmFathersNotesInfancy.RichText1)
    Case 14 'fathers notes child
        Call btnLetterClick(2, frmFathersNotesChild.RichText1)
    Case 15 'birth
        Call btnLetterClick(2, frmBirth.RichTextBox1)
    Case 17 'birthdays
        Call btnLetterClick(2, frmBirthDates.RichTextBox1)
    Case 18 'hospital
        Select Case frmHospital.Tab1.Tab
        Case 0
            Call btnLetterClick(2, frmHospital.RichTextBox1(0))
        Case 1
            Call btnLetterClick(2, frmHospital.RichTextBox1(1))
        Case 2
            Call btnLetterClick(2, frmHospital.RichTextBox1(2))
        Case Else
        End Select
    Case 20 'baptism
        Select Case frmBaptism.Tab1.Tab
        Case 1  'goodmothers/-fathers
            Call btnLetterClick(2, frmBaptism.RichTextBox1)
        Case 2  'attendees
            Call btnLetterClick(2, frmBaptism.RichTextBox2)
        Case 3  'gifts
            Call btnLetterClick(2, frmBaptism.RichTextBox3)
        Case 4  'notes
            Call btnLetterClick(2, frmBaptism.RichTextBox4)
        Case Else
        End Select
    Case 21 'books
        Select Case frmBooks.Tab1.Tab
        Case 0  'infants
            Call btnLetterClick(2, frmBooks.RichTextBox1(0))
        Case 1  'child
            Call btnLetterClick(2, frmBooks.RichTextBox1(1))
        Case Else
        End Select
    Case 22 'toys
        Select Case frmToys.Tab1.Tab
        Case 0  'infant
            Call btnLetterClick(2, frmToys.RichTextBox1(0))
        Case 1  'child
            Call btnLetterClick(2, frmToys.RichTextBox1(1))
        Case Else
        End Select
    Case 23 'health
        Select Case frmHealth.Tab1.Tab
            Case 0  'infant
                Select Case frmHealth.Tab2.Tab
                    Case 0  'control
                        Call btnLetterClick(2, frmHealth.RichTextBox1(0))
                    Case 1  'vacination
                        Call btnLetterClick(2, frmHealth.RichTextBox1(1))
                    Case 2  'illness
                        Call btnLetterClick(2, frmHealth.RichTextBox1(2))
                    Case Else
                    End Select
            Case 1  'child
                Select Case frmHealth.Tab3.Tab
                    Case 0  'control
                        Call btnLetterClick(2, frmHealth.RichTextBox1(3))
                    Case 1  'vacination
                        Call btnLetterClick(2, frmHealth.RichTextBox1(4))
                    Case 2   'illness
                        Call btnLetterClick(2, frmHealth.RichTextBox1(5))
                    Case Else
                    End Select
        Case Else
        End Select
    Case 24 'email
        Call btnLetterClick(2, frmEmail.RichText1)
    Case 25 'food habits
        Call btnLetterClick(2, frmFoodHabits.RichTextBox1)
    Case 27 'fathers notes birth
        Call btnLetterClick(2, frmFathersNotesBirth.RichText1)
    Case 34 'write to me
        Call btnLetterClick(2, frmWriteToMe.RichTextBox1)
    Case 35 'first
        Call btnLetterClick(2, frmFirst.RichTextBox1)
    Case Else
    End Select
End Sub

Private Sub mnuHelpFileHtml_Click()
Dim sText As String
    On Error GoTo ErrmnuHelpHtml
    If FileExt = "NOR" Then
        sText = rsMyRecord.Fields("HtmlDirectory") & "/norwegian_index.htm"
    Else
        sText = rsMyRecord.Fields("HtmlDirectory") & "/english_index.htm"
    End If
    Call ShellExceCute(Me.hWnd, "open", sText, vbNullString, vbNullString, SW_SHOWNORMAL)
    Exit Sub
    
ErrmnuHelpHtml:
    Beep
    MsgBox Err.Description, vbExclamation, "Help File"
    Resume ErrmnuHelpHtml2
ErrmnuHelpHtml2:
End Sub

Private Sub mnuHelpFileWord_Click()
Dim sText As String
    On Error GoTo ErrmnuHelpWord
    If FileExt = "ENG" Then
        sText = App.Path & "/MasterKidEng.doc"
    Else
        sText = App.Path & "/MasterKidNor.doc"
    End If
    Call ShellExceCute(Me.hWnd, "open", sText, vbNullString, vbNullString, SW_SHOWNORMAL)
    Exit Sub
    
ErrmnuHelpWord:
    Beep
    MsgBox Err.Description, vbExclamation, "Help File"
    Resume ErrmnuHelpWord2
ErrmnuHelpWord2:
End Sub

Private Sub mnuHome_Click()
    CloseActiveForm
    frmHospital.Tab1.Tab = 3
    frmHospital.Show
End Sub

Private Sub mnuHoroscope_Click()
    CloseActiveForm
    frmHoroscope.Show
End Sub

Private Sub mnuInfHealth_Click()
    On Error Resume Next
    CloseActiveForm
    iTab = 0
    frmHealth.Show
End Sub

Private Sub mnuInfPic_Click()
    On Error Resume Next
    CloseActiveForm
    iTab = 1
    frmPictures.Show
End Sub

Private Sub mnuAbove_Click()
    CloseActiveForm
    frmAbout.Show
End Sub

Private Sub mnuAccessScan_Click()
    On Error Resume Next
    Dim ret As Long
    ret = TWAIN_SelectImageSource(Me.hWnd)
End Sub

Private Sub mnuAntenatal_Click()
    CloseActiveForm
    frmAntenatal.Show
End Sub

Private Sub mnuBapChrist_Click()
    CloseActiveForm
    frmBaptism.Show
End Sub

Private Sub mnuBapFaNotes_Click()
    CloseActiveForm
    frmFathersNotesBaptism.Show
End Sub

Private Sub mnuBapGodmother_Click()
    CloseActiveForm
    frmBaptism.Tab1.Tab = 1
    frmBaptism.Show
End Sub

Private Sub mnuBapNames_Click()
    CloseActiveForm
    frmNames.Show
End Sub

Private Sub mnuBapPic_Click()
    CloseActiveForm
    frmBaptismPictures.Show
End Sub

Private Sub mnuBapVideo_Click()
    On Error Resume Next
    CloseActiveForm
    iTab = 1
    frmVideo.Show
End Sub

Private Sub mnuBirthAcq_Click()
    CloseActiveForm
    frmHospital.Tab1.Tab = 1
    frmHospital.Show
End Sub

Private Sub mnuBirthDairy_Click()
    CloseActiveForm
    frmHospital.Show
End Sub

Private Sub mnuBirthFaNotes_Click()
    CloseActiveForm
    frmFathersNotesBirth.Show
End Sub

Private Sub mnuBirthLeaving_Click()
    CloseActiveForm
    frmHospital.Tab1.Tab = 2
    frmHospital.Show
End Sub

Private Sub mnuBirthPic_Click()
    CloseActiveForm
    iTab = 0
    frmPictures.Show
End Sub

Private Sub mnuBirthRem_Click()
    On Error Resume Next
    CloseActiveForm
    With frmRemember
        .BackColor = &H8000&
        .Frame1.BackColor = &H8000&
        .Check1.BackColor = &H8000&
        .Label1(0).ForeColor = &HFFFFFF
        .Label1(1).ForeColor = &HFFFFFF
        .Label1(2).ForeColor = &HFFFFFF
        .Label1(3).ForeColor = &HFFFFFF
        .Show
    End With
End Sub

Private Sub mnuBirthThe_Click()
    CloseActiveForm
    frmBirth.Show
End Sub

Private Sub mnuBirthVideo_Click()
    CloseActiveForm
    iTab = 0
    frmVideo.Show
End Sub

Private Sub mnuChildBirthDays_Click()
    CloseActiveForm
    frmBirthDates.Show
End Sub

Private Sub mnuChildBooks_Click()
    CloseActiveForm
    iTab = 1
    frmBooks.Show
End Sub

Private Sub mnuChildFaNotes_Click()
    CloseActiveForm
    frmFathersNotesChild.Show
End Sub

Private Sub mnuChildPic_Click()
    CloseActiveForm
    iTab = 2
    frmPictures.Show
End Sub

Private Sub mnuChildToys_Click()
    CloseActiveForm
    iTab = 1
    frmToys.Show
End Sub

Private Sub mnuChildVideo_Click()
    CloseActiveForm
    iTab = 3
    frmVideo.Show
End Sub

Private Sub mnuIamPregnant_Click()
    CloseActiveForm
    frmIamPregnant.Show
End Sub

Private Sub mnuInfBooks_Click()
    CloseActiveForm
    iTab = 0
    frmBooks.Show
End Sub

Private Sub mnuInfFaNotes_Click()
    CloseActiveForm
    frmFathersNotesInfancy.Show
End Sub

Private Sub mnuInfFirst_Click()
    CloseActiveForm
    frmFirstTimes.Show
End Sub

Private Sub mnuInfFood_Click()
    CloseActiveForm
    frmFoodHabits.Show
End Sub

Private Sub mnuInfSound_Click()
    On Error Resume Next
    CloseActiveForm
    iTab = 1
    frmSound.Show
End Sub

Private Sub mnuInfToys_Click()
    CloseActiveForm
    iTab = 0
    frmToys.Show
End Sub

Private Sub mnuInfVideo_Click()
    CloseActiveForm
    iTab = 2
    frmVideo.Show
End Sub

Private Sub mnuInfWeight_Click()
    CloseActiveForm
    frmWeightLength.Show
End Sub

Private Sub mnuInternetConnect_Click(Index As Integer)
Dim iret As Long
    On Error GoTo errmnuInternetConnect_Click
    With rsInternet
        .MoveFirst
        Do While Not .EOF
            If .Fields("LinkName") = Me.mnuInternetConnect(Index).Caption Then
                iret = ShellExceCute(Me.hWnd, _
                    vbNullString, _
                    ("http://" & .Fields("LinkHyper")), vbNullString, "c:\", _
                    SW_SHOWNORMAL)
                Exit Do
            End If
        .MoveNext
        Loop
    End With
    Exit Sub
    
errmnuInternetConnect_Click:
    Beep
    MsgBox Err.Description, vbExclamation, "Internet"
    WriteErrorFile Err.Description, "frmMDI: Internet Connection"
    Resume errmnuInternetConnect_Click2
errmnuInternetConnect_Click2:
End Sub

Public Sub ReadText()
Dim strHelp As String
    'find YOUR rsLanguage text
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
                .Update
                DBEngine.Idle dbFreeLocks
                LoadMenu1
                ShowMenu
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
        .Fields("Form") = Me.Caption
        .Fields("Help") = strHelp
        .Fields("Msg1") = "Do you really want to delete this record ?"
        .Fields("Msg2") = "Sorry, You have to register this programme first !"
        MakeNewMenu1
        WriteNewMenu
        .Update
        .Bookmark = .LastModified
        ShowMenu
        LoadMenu1
    End With
End Sub

Public Sub CloseActiveForm()
  On Error GoTo Cancel
  Do
    Unload Me.ActiveForm 'This will cause an error when all forms are unloaded already
  Loop
Cancel:
    Exit Sub
End Sub

Private Sub LoadFonts()
    On Error Resume Next
    Me.mnuFormatFont(0).Caption = Screen.Fonts(0)
    For n = 1 To Screen.FontCount - 1
        Load Me.mnuFormatFont(n)
        Me.mnuFormatFont(n).Caption = Screen.Fonts(n)
    Next
    
    'load font size
    Me.mnuFormatFontSize(0).Caption = 8
    For n = 9 To 48
        Load Me.mnuFormatFontSize(n)
        Me.mnuFormatFontSize(n).Caption = n
    Next
End Sub


Private Sub ShowPrint()
    On Error GoTo errPrint
    Select Case iWhichForm
    Case 1  'frmPregnancyNotes
        If PrintUseWord Then
            Call frmPregnancyNotes.WritePregnancyNotesWord
        Else
            Call frmPregnancyNotes.WritePregnancyNotes
        End If
    Case 2  'frmIamPregnant
        If PrintUseWord Then
            Call frmIamPregnant.IamPregnantPrintWord
        Else
            Call frmIamPregnant.IamPregnantPrint
        End If
    Case 3  'kids
        If PrintUseWord Then
            Call frmKids.WriteKidsWord
        Else
            Call frmKids.WriteKids
        End If
    Case 4  'MidWife
        If PrintUseWord Then
            Call frmMidWife.WriteMidwifeWord
        Else
            Call frmMidWife.WriteMidwife
        End If
    Case 5  'pregnancy control
        If PrintUseWord Then
            Call frmPregnancyControl.WritePregnancyControlWord
        Else
            Call frmPregnancyControl.WritePregnancyControl
        End If
    Case 6  'antenatal classes
        If PrintUseWord Then
            Call frmAntenatal.PrintAntenatalWord
        Else
            Call frmAntenatal.PrintAntenatal
        End If
    Case 7  'user information
    Case 8  'Names
    Case 9  'Term
    Case 10 'Fathers Notes Pregnancy
        If PrintUseWord Then
            Call frmFathersNotesPregnancy.WriteFatherPregnancyNotesWord
        Else
            Call frmFathersNotesPregnancy.WriteFatherPregnancyNotes
        End If
    Case 11 'country
    Case 12 'Fathers Notes Birth
        If PrintUseWord Then
            Call frmFathersNotesBaptism.WriteBaptismNotesWord
        Else
            Call frmFathersNotesBaptism.WriteBaptismNotes
        End If
    Case 13 'Fathers Notes Infancy
        If PrintUseWord Then
            Call frmFathersNotesInfancy.WriteInfancyNotesWord
        Else
            Call frmFathersNotesInfancy.WriteInfancyNotes
        End If
    Case 14 'Fathers Notes Child
        If PrintUseWord Then
            Call frmFathersNotesChild.WriteChildNotesWord
        Else
            Call frmFathersNotesChild.WriteChildNotes
        End If
    Case 15 'Birth
        If PrintUseWord Then
            Call frmBirth.PrintBirthWord
        Else
            Call frmBirth.PrintBirth
        End If
    Case 16 'dimensions
    Case 17 'birthdays
        Select Case frmBirthDates.Tab1.Tab
            Case 0
                If PrintUseWord Then
                    Call frmBirthDates.PrintBirthdaysWord
                Else
                    Call frmBirthDates.PrintBirthdays
                End If
            Case 1
                Call PrintPictureToFitPage(Printer, frmBirthDates.Picture1)
            Case Else
        End Select
    Case 18 'hospital
        If PrintUseWord Then
            Call frmHospital.WriteHospitalWord
        Else
            Call frmHospital.WriteHospital
        End If
    Case 19 'first time...
        If PrintUseWord Then
            Call frmFirstTimes.WriteFirstTimesWord
        Else
            Call frmFirstTimes.WriteFirstTimes
        End If
    Case 20 'baptism
        If PrintUseWord Then
            Call frmBaptism.PrintBaptismWord
        Else
            Call frmBaptism.PrintBaptism
        End If
    Case 21 'books
        If PrintUseWord Then
            Call frmBooks.PrintBooksWord
        Else
            Call frmBooks.PrintBooks
        End If
    Case 22 'Toys
        If PrintUseWord Then
            Call frmToys.WriteToysWord
        Else
            Call frmToys.WriteToys
        End If
    Case 23 'Health
        If PrintUseWord Then
            Call frmHealth.WriteHealthWord
        Else
            Call frmHealth.WriteHealth
        End If
    Case 25 'food habits
        If PrintUseWord Then
            Call frmFoodHabits.WriteFoodHabitsWord
        Else
            Call frmFoodHabits.WriteFoodHabits
        End If
    Case 26 'rembember to ..
        If PrintUseWord Then
            Call frmRemember.WriteRememberWord
        Else
            Call frmRemember.WriteRemember
        End If
    Case 27 'fathers notes Birth
        If PrintUseWord Then
            Call frmFathersNotesBirth.WriteBirthNotesWord
        Else
            Call frmFathersNotesBirth.WriteBirthNotes
        End If
    Case 28 'baptism pictures
        If PrintUseWord Then
            Call frmBaptismPictures.PrintBaptismPicWord
        Else
            Call frmBaptismPictures.PrintBaptismPic
        End If
    Case 29 'pictures
        Call frmPictures.Write_Print
    Case 32 'first pram
        If PrintUseWord Then
            Call frmFirstPram.WriteFirstPramWord
        Else
            Call frmFirstPram.WriteFirstPram
        End If
    Case 35 'first information
        If PrintUseWord Then
            Call frmFirst.WriteFirst
        Else
            Call frmFirst.PrintFirst
        End If
    Case 37 'the teeths
        If PrintUseWord Then
            Call frmTeeth.WriteTeethWord
        Else
            Call frmTeeth.WriteTeeth
        End If
    Case 41 'weight & Length
        If PrintUseWord Then
            Call frmWeightLength.WriteWeightLengthWord
        Else
            Call frmWeightLength.WriteWeightLength
        End If
    Case 43 'when I was born
        If PrintUseWord Then
            Call frmWhenIWasBorn.WriteWhenIWasBorn
        Else
            Call frmWhenIWasBorn.PrintWhenIWasBorn
        End If
    Case Else
    End Select
    Exit Sub
    
errPrint:
    Beep
    MsgBox "Error:  " & Err.Number & " -  " & Err.Description, vbExclamation, "Print"
    Err.Clear
End Sub

Private Sub MeActivate()
    On Error Resume Next
    With rsMyRecord
        FileExt = .Fields("LanguageScreen")
        If CBool(.Fields("PrintUsingWord")) Then
            PrintUseWord = True
            mnuPrintWord.Checked = True
        Else
            PrintUseWord = False
            mnuPrintWord.Checked = False
        End If
        If Not CBool(.Fields("ShowFirst")) Then
            frmFirst.Show
        End If
        DoEvents
    End With
    
    LoadChildren
    cmbChildren.ListIndex = 0
    glChildNo = CLng(rsChildren.Fields("ChildNo"))
    ReadText
    LoadFonts
    LoadColors
    HideAllButtons
    DoEvents
End Sub
Private Sub MDIForm_Load()
Dim lngOldId As Long, RetVal As Long
    On Error GoTo errForm_Load
    dbKidsTxt = App.Path & "\MasterKid.mdb"
    dbKidLangTxt = App.Path & "\MasterKidLang.mdb"
    dbKidPicTxt = App.Path & "\MasterKidPic.mdb"
    
    Set dbKids = OpenDatabase(dbKidsTxt)
    Set dbKidLang = OpenDatabase(dbKidLangTxt)
    Set dbKidPic = OpenDatabase(dbKidPicTxt)
    
    Set rsMyRecord = dbKids.OpenRecordset("MyRecord")
    Set rsInternet = dbKids.OpenRecordset("Internet")
    Set rsChildren = dbKids.OpenRecordset("Children")
    Set rsLanguage = dbKidLang.OpenRecordset("MDIMasterKid")
    Set rsColor = dbKids.OpenRecordset("Color")
    
    hSysMenu = GetSystemMenu(Me.hWnd, False)
    With zMENU
        .cbSize = Len(zMENU)
        .dwTypeData = String(80, 0)
        .cch = Len(.dwTypeData)
        .fMask = MENU_STATE
        .wid = SC_CLOSE
    End With
    RetVal = GetMenuItemInfo(hSysMenu, zMENU.wid, False, zMENU)
    With zMENU
        lngOldId = .wid         'You need the old wID.
        .wid = xSC_CLOSE        'Change the wID to "no close"
        .fState = MFS_GRAYED    'Make the close methods gray
        .fMask = MENU_ID        'Specifys that the value in wID is a id and not a state
    End With
    RetVal = SetMenuItemInfo(hSysMenu, lngOldId, False, zMENU)
    zMENU.fMask = MENU_STATE
    RetVal = SetMenuItemInfo(hSysMenu, zMENU.wid, False, zMENU)
    MeActivate
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmMDIMasterKid: Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRecord.Close
    rsInternet.Close
    rsChildren.Close
    rsLanguage.Close
    dbKids.Close
    dbKidLang.Close
    dbKidPic.Close
    rsColor.Close
    Erase v1RecordBookmark
    Set MDIMasterKid = Nothing
    End
End Sub
Private Sub mnuCountry_Click()
    CloseActiveForm
    frmCountry.Show
End Sub

Private Sub mnuDimension_Click()
    CloseActiveForm
    frmDimensions.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuInternetBase_Click()
    CloseActiveForm
    frmInternet.Show
End Sub

Private Sub mnuMailDeveloper_Click()
    CloseActiveForm
    frmWriteToMe.Show
End Sub

Private Sub mnuNames_Click()
    CloseActiveForm
    frmAllNameExplanation.Show
End Sub

Private Sub mnuPregFaNotes_Click()
    CloseActiveForm
    frmFathersNotesPregnancy.Show
End Sub

Private Sub mnuPregnancyControl_Click()
    CloseActiveForm
    frmPregnancyControl.Show
End Sub

Private Sub mnuPregnancyNotes_Click()
    CloseActiveForm
    frmPregnancyNotes.Show
End Sub

Private Sub mnuPregToRemember_Click()
    CloseActiveForm
    frmRemember.BackColor = &H80FF&
    frmRemember.Frame1.BackColor = &H80FF&
    frmRemember.Label1(0).BackColor = &H80FF&
    frmRemember.Label1(1).BackColor = &H80FF&
    frmRemember.Label1(2).BackColor = &H80FF&
    frmRemember.Label1(3).BackColor = &H80FF&
    frmRemember.Check1.BackColor = &H80FF&
    frmRemember.Label1(0).ForeColor = &H80000012
    frmRemember.Label1(1).ForeColor = &H80000012
    frmRemember.Label1(2).ForeColor = &H80000012
    frmRemember.Label1(3).ForeColor = &H80000012
    frmRemember.Show
End Sub

Private Sub mnuPrintComplete_Click()
    CloseActiveForm
    frmPrint.Show
End Sub

Private Sub mnuPrintCompleteFrames_Click()
    CloseActiveForm
    frmFrames.Show
End Sub
Private Sub mnuPrintSetUp_Click()
    On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Print"
        .CancelError = True
        .flags = cdlPDPrintSetup
        .ShowPrinter
    End With
End Sub

Private Sub mnuPrintWord_Click()
    On Error Resume Next
    If mnuPrintWord.Checked = True Then
        mnuPrintWord.Checked = False
        PrintUseWord = False
        With rsMyRecord
            .Edit
            .Fields("PrintUsingWord") = False
            .Update
        End With
    Else
        mnuPrintWord.Checked = True
        If IsAppPresent("Word.Document\CurVer", "") Then
            PrintUseWord = True
            With rsMyRecord
                .Edit
                .Fields("PrintUsingWord") = True
                .Update
            End With
        End If
    End If
End Sub

Private Sub mnuRegistration_Click()
    CloseActiveForm
    frmRegistration.Show
End Sub

Private Sub mnuRegistrationIn_Click()
    CloseActiveForm
    frmRegistrateProgramme.Show 1
End Sub

Private Sub mnuScreenText_Click()
    CloseActiveForm
    frmScreenText.Show
End Sub

Private Sub mnuSupplier_Click()
    CloseActiveForm
    frmSupplier.Show
End Sub

Private Sub mnuTeeth_Click()
    CloseActiveForm
    frmTeeth.Show
End Sub

Private Sub mnuTerm_Click()
    CloseActiveForm
    frmTerm.Show
End Sub
Private Sub mnuUpdate_Click()
Dim iret As Long
    
    On Error GoTo errmnuUpdate_Click
    frmLiveUpdate.Show 1
    Exit Sub
    
errmnuUpdate_Click:
    Beep
    MsgBox Err.Description, vbExclamation, "Internet Update"
    WriteErrorFile Err.Description, "MDIMasterKid: Internet Update"
    Err.Clear
End Sub

Private Sub mnuUser_Click()
    CloseActiveForm
    frmUser.Show
End Sub

Private Sub Timer1_Timer()
    CheckAlarm
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
    Case "Exit"
        Unload Me
    Case "Print"
        ShowPrint
    Case "New"
        NewRecords
    Case "Delete"
        DeleteRecords
    Case "Email"
        ShowEmail
    Case "Help"
        ShowHelp
    Case Else
    End Select
End Sub
