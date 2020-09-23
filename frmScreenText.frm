VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmScreenText 
   BackColor       =   &H00800000&
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   7680
   Begin TabDlg.SSTab Tab1 
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   45
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   8388608
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmScreenText.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGrid1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Data1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Antenatal"
      TabPicture(1)   =   "frmScreenText.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1(1)"
      Tab(1).Control(1)=   "Data2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Baptism"
      TabPicture(2)   =   "frmScreenText.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DBGrid1(2)"
      Tab(2).Control(1)=   "Data3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Baptism Pictures"
      TabPicture(3)   =   "frmScreenText.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "DBGrid1(3)"
      Tab(3).Control(1)=   "Data4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Birth"
      TabPicture(4)   =   "frmScreenText.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DBGrid1(4)"
      Tab(4).Control(1)=   "Data5"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Birth Dates"
      TabPicture(5)   =   "frmScreenText.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DBGrid1(5)"
      Tab(5).Control(1)=   "Data6"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Books"
      TabPicture(6)   =   "frmScreenText.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "DBGrid1(6)"
      Tab(6).Control(1)=   "Data7"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Country"
      TabPicture(7)   =   "frmScreenText.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "DBGrid1(7)"
      Tab(7).Control(1)=   "Data8"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Dimension"
      TabPicture(8)   =   "frmScreenText.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "DBGrid1(8)"
      Tab(8).Control(1)=   "Data9"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Email"
      TabPicture(9)   =   "frmScreenText.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "DBGrid1(9)"
      Tab(9).Control(1)=   "Data10"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Baptism Notes"
      TabPicture(10)  =   "frmScreenText.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "DBGrid1(10)"
      Tab(10).Control(1)=   "Data11"
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "Birth Notes"
      TabPicture(11)  =   "frmScreenText.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "DBGrid1(11)"
      Tab(11).Control(1)=   "Data12"
      Tab(11).ControlCount=   2
      TabCaption(12)  =   "Child Notes"
      TabPicture(12)  =   "frmScreenText.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "DBGrid1(12)"
      Tab(12).Control(1)=   "Data13"
      Tab(12).ControlCount=   2
      TabCaption(13)  =   "Infant Notes"
      TabPicture(13)  =   "frmScreenText.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "DBGrid1(13)"
      Tab(13).Control(1)=   "Data14"
      Tab(13).ControlCount=   2
      TabCaption(14)  =   "Pregnancy Notes"
      TabPicture(14)  =   "frmScreenText.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "DBGrid1(14)"
      Tab(14).Control(1)=   "Data15"
      Tab(14).ControlCount=   2
      TabCaption(15)  =   "First Time"
      TabPicture(15)  =   "frmScreenText.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "DBGrid1(15)"
      Tab(15).Control(1)=   "Data16"
      Tab(15).ControlCount=   2
      TabCaption(16)  =   "Food Habits"
      TabPicture(16)  =   "frmScreenText.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "DBGrid1(16)"
      Tab(16).Control(1)=   "Data17"
      Tab(16).ControlCount=   2
      TabCaption(17)  =   "Health"
      TabPicture(17)  =   "frmScreenText.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "DBGrid1(17)"
      Tab(17).Control(1)=   "Data18"
      Tab(17).ControlCount=   2
      TabCaption(18)  =   "Hospital"
      TabPicture(18)  =   "frmScreenText.frx":01F8
      Tab(18).ControlEnabled=   0   'False
      Tab(18).Control(0)=   "DBGrid1(18)"
      Tab(18).Control(1)=   "Data19"
      Tab(18).ControlCount=   2
      TabCaption(19)  =   "I am pregnant"
      TabPicture(19)  =   "frmScreenText.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "DBGrid1(19)"
      Tab(19).Control(1)=   "Data20"
      Tab(19).ControlCount=   2
      TabCaption(20)  =   "Internet"
      TabPicture(20)  =   "frmScreenText.frx":0230
      Tab(20).ControlEnabled=   0   'False
      Tab(20).Control(0)=   "DBGrid1(20)"
      Tab(20).Control(1)=   "Data21"
      Tab(20).ControlCount=   2
      TabCaption(21)  =   "Kids"
      TabPicture(21)  =   "frmScreenText.frx":024C
      Tab(21).ControlEnabled=   0   'False
      Tab(21).Control(0)=   "DBGrid1(21)"
      Tab(21).Control(1)=   "Data22"
      Tab(21).ControlCount=   2
      TabCaption(22)  =   "Midwife"
      TabPicture(22)  =   "frmScreenText.frx":0268
      Tab(22).ControlEnabled=   0   'False
      Tab(22).Control(0)=   "DBGrid1(22)"
      Tab(22).Control(1)=   "Data23"
      Tab(22).ControlCount=   2
      TabCaption(23)  =   "Names"
      TabPicture(23)  =   "frmScreenText.frx":0284
      Tab(23).ControlEnabled=   0   'False
      Tab(23).Control(0)=   "DBGrid1(23)"
      Tab(23).Control(1)=   "Data24"
      Tab(23).ControlCount=   2
      TabCaption(24)  =   "Pictures"
      TabPicture(24)  =   "frmScreenText.frx":02A0
      Tab(24).ControlEnabled=   0   'False
      Tab(24).Control(0)=   "DBGrid1(24)"
      Tab(24).Control(1)=   "Data25"
      Tab(24).ControlCount=   2
      TabCaption(25)  =   "Pregnancy Control"
      TabPicture(25)  =   "frmScreenText.frx":02BC
      Tab(25).ControlEnabled=   0   'False
      Tab(25).Control(0)=   "DBGrid1(25)"
      Tab(25).Control(1)=   "Data26"
      Tab(25).ControlCount=   2
      TabCaption(26)  =   "Pregnancy Notes"
      TabPicture(26)  =   "frmScreenText.frx":02D8
      Tab(26).ControlEnabled=   0   'False
      Tab(26).Control(0)=   "DBGrid1(26)"
      Tab(26).Control(1)=   "Data27"
      Tab(26).ControlCount=   2
      TabCaption(27)  =   "Print"
      TabPicture(27)  =   "frmScreenText.frx":02F4
      Tab(27).ControlEnabled=   0   'False
      Tab(27).Control(0)=   "DBGrid1(27)"
      Tab(27).Control(1)=   "Data28"
      Tab(27).ControlCount=   2
      TabCaption(28)  =   "Remember"
      TabPicture(28)  =   "frmScreenText.frx":0310
      Tab(28).ControlEnabled=   0   'False
      Tab(28).Control(0)=   "DBGrid1(28)"
      Tab(28).Control(1)=   "Data29"
      Tab(28).ControlCount=   2
      TabCaption(29)  =   "Term"
      TabPicture(29)  =   "frmScreenText.frx":032C
      Tab(29).ControlEnabled=   0   'False
      Tab(29).Control(0)=   "DBGrid1(29)"
      Tab(29).Control(1)=   "Data30"
      Tab(29).ControlCount=   2
      TabCaption(30)  =   "Toys"
      TabPicture(30)  =   "frmScreenText.frx":0348
      Tab(30).ControlEnabled=   0   'False
      Tab(30).Control(0)=   "DBGrid1(30)"
      Tab(30).Control(1)=   "Data31"
      Tab(30).ControlCount=   2
      TabCaption(31)  =   "User"
      TabPicture(31)  =   "frmScreenText.frx":0364
      Tab(31).ControlEnabled=   0   'False
      Tab(31).Control(0)=   "DBGrid1(31)"
      Tab(31).Control(1)=   "Data32"
      Tab(31).ControlCount=   2
      TabCaption(32)  =   "Video"
      TabPicture(32)  =   "frmScreenText.frx":0380
      Tab(32).ControlEnabled=   0   'False
      Tab(32).Control(0)=   "DBGrid1(32)"
      Tab(32).Control(1)=   "Data33"
      Tab(32).ControlCount=   2
      TabCaption(33)  =   "Weight/Height"
      TabPicture(33)  =   "frmScreenText.frx":039C
      Tab(33).ControlEnabled=   0   'False
      Tab(33).Control(0)=   "DBGrid1(33)"
      Tab(33).Control(1)=   "Data34"
      Tab(33).ControlCount=   2
      TabCaption(34)  =   "Word"
      TabPicture(34)  =   "frmScreenText.frx":03B8
      Tab(34).ControlEnabled=   0   'False
      Tab(34).Control(0)=   "DBGrid1(34)"
      Tab(34).Control(1)=   "Data35"
      Tab(34).ControlCount=   2
      TabCaption(35)  =   "Horoscope"
      TabPicture(35)  =   "frmScreenText.frx":03D4
      Tab(35).ControlEnabled=   0   'False
      Tab(35).Control(0)=   "DBGrid1(35)"
      Tab(35).Control(1)=   "Data36"
      Tab(35).ControlCount=   2
      TabCaption(36)  =   "Zodiac"
      TabPicture(36)  =   "frmScreenText.frx":03F0
      Tab(36).ControlEnabled=   0   'False
      Tab(36).Control(0)=   "DBGrid1(36)"
      Tab(36).Control(1)=   "Data37"
      Tab(36).ControlCount=   2
      TabCaption(37)  =   "Colors"
      TabPicture(37)  =   "frmScreenText.frx":040C
      Tab(37).ControlEnabled=   0   'False
      Tab(37).Control(0)=   "DBGrid1(37)"
      Tab(37).Control(1)=   "Data38"
      Tab(37).ControlCount=   2
      TabCaption(38)  =   "Name Explanation"
      TabPicture(38)  =   "frmScreenText.frx":0428
      Tab(38).ControlEnabled=   0   'False
      Tab(38).Control(0)=   "DBGrid1(38)"
      Tab(38).Control(1)=   "Data39"
      Tab(38).ControlCount=   2
      TabCaption(39)  =   "Name registration"
      TabPicture(39)  =   "frmScreenText.frx":0444
      Tab(39).ControlEnabled=   0   'False
      Tab(39).Control(0)=   "DBGrid1(39)"
      Tab(39).Control(1)=   "Data40"
      Tab(39).ControlCount=   2
      TabCaption(40)  =   "Print"
      TabPicture(40)  =   "frmScreenText.frx":0460
      Tab(40).ControlEnabled=   0   'False
      Tab(40).Control(0)=   "DBGrid1(40)"
      Tab(40).Control(1)=   "Data41"
      Tab(40).ControlCount=   2
      TabCaption(41)  =   "When I was born"
      TabPicture(41)  =   "frmScreenText.frx":047C
      Tab(41).ControlEnabled=   0   'False
      Tab(41).Control(0)=   "DBGrid1(41)"
      Tab(41).Control(1)=   "Data42"
      Tab(41).ControlCount=   2
      TabCaption(42)  =   "Registration"
      TabPicture(42)  =   "frmScreenText.frx":0498
      Tab(42).ControlEnabled=   0   'False
      Tab(42).Control(0)=   "DBGrid1(42)"
      Tab(42).Control(0).Enabled=   0   'False
      Tab(42).Control(1)=   "Data43"
      Tab(42).Control(1).Enabled=   0   'False
      Tab(42).ControlCount=   2
      TabCaption(43)  =   "Program Update"
      TabPicture(43)  =   "frmScreenText.frx":04B4
      Tab(43).ControlEnabled=   0   'False
      Tab(43).Control(0)=   "DBGrid1(43)"
      Tab(43).Control(0).Enabled=   0   'False
      Tab(43).Control(1)=   "Data44"
      Tab(43).Control(1).Enabled=   0   'False
      Tab(43).ControlCount=   2
      TabCaption(44)  =   "Prog. Information"
      TabPicture(44)  =   "frmScreenText.frx":04D0
      Tab(44).ControlEnabled=   0   'False
      Tab(44).Control(0)=   "DBGrid1(44)"
      Tab(44).Control(0).Enabled=   0   'False
      Tab(44).Control(1)=   "Data45"
      Tab(44).Control(1).Enabled=   0   'False
      Tab(44).Control(2)=   "RichTextBox1"
      Tab(44).Control(2).Enabled=   0   'False
      Tab(44).ControlCount=   3
      Begin RichTextLib.RichTextBox RichTextBox1 
         DataField       =   "RichTextBox1"
         DataSource      =   "Data45"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   46
         Top             =   3840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6165
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmScreenText.frx":04EC
      End
      Begin VB.Data Data45 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFirst"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "MDIMasterKid"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmAntenatal"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data3 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmBaptism"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data4 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmBaptismPictures"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data5 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmBirth"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data6 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmBirthDates"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data7 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmBooks"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data8 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmCountry"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data9 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmDimensions"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data10 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmEmail"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data11 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFathersNotesBaptism"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data12 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFathersNotesBirth"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data13 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFathersNotesChild"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data14 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFathersNotesInfancy"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data15 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFathersNotesPregnancy"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data16 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFirstTimes"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data17 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmFoodHabits"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data18 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmHealth"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data19 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmHospital"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data20 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmIamPregnant"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data21 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmInternet"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data22 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmKids"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data23 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmMidWife"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data24 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmNames"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data25 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPictures"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data26 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPregnancyControl"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data27 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPregnancyNotes"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data28 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPrint"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data29 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmRemember"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data30 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmTerm"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data31 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmToys"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data32 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmUser"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data33 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmVideo"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data34 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmWeightLength"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data35 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LanguageWord"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data36 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmHoroscope"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data37 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmZodiac"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data38 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmColor"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data39 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmNamesExplanation"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data40 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmAllNameExplanation"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data41 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPrint"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data42 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmWhenIWasBorn"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data43 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmRegister"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data44 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\MasterKid\MasterKidLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmLiveUpdate"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":05C1
         Height          =   4725
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "frmScreenText.frx":05D5
         TabIndex        =   1
         Top             =   2640
         Width           =   5610
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":0FAB
         Height          =   4755
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmScreenText.frx":0FBF
         TabIndex        =   2
         Top             =   2640
         Width           =   5730
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":1995
         Height          =   4740
         Index           =   2
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":19A9
         TabIndex        =   3
         Top             =   2640
         Width           =   5595
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":237F
         Height          =   4740
         Index           =   3
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":2393
         TabIndex        =   4
         Top             =   2640
         Width           =   5595
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":2D69
         Height          =   4740
         Index           =   4
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":2D7D
         TabIndex        =   5
         Top             =   2640
         Width           =   5595
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":3753
         Height          =   4725
         Index           =   5
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":3767
         TabIndex        =   6
         Top             =   2640
         Width           =   5610
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":413D
         Height          =   4740
         Index           =   6
         Left            =   -74880
         OleObjectBlob   =   "frmScreenText.frx":4151
         TabIndex        =   7
         Top             =   2640
         Width           =   5715
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":4B27
         Height          =   4740
         Index           =   7
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":4B3B
         TabIndex        =   8
         Top             =   2640
         Width           =   5595
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":5511
         Height          =   4740
         Index           =   8
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":5525
         TabIndex        =   9
         Top             =   2640
         Width           =   5610
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":5EFB
         Height          =   4740
         Index           =   9
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":5F10
         TabIndex        =   10
         Top             =   2640
         Width           =   5595
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":68E6
         Height          =   4695
         Index           =   10
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":68FB
         TabIndex        =   11
         Top             =   2640
         Width           =   5580
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":72D2
         Height          =   4695
         Index           =   11
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":72E7
         TabIndex        =   12
         Top             =   2640
         Width           =   5640
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":7CBE
         Height          =   4710
         Index           =   12
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":7CD3
         TabIndex        =   13
         Top             =   2640
         Width           =   5520
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":86AA
         Height          =   4695
         Index           =   13
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":86BF
         TabIndex        =   14
         Top             =   2640
         Width           =   5640
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":9096
         Height          =   4695
         Index           =   14
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":90AB
         TabIndex        =   15
         Top             =   2640
         Width           =   5640
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":9A82
         Height          =   4695
         Index           =   15
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":9A97
         TabIndex        =   16
         Top             =   2640
         Width           =   5640
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":A46E
         Height          =   4695
         Index           =   16
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":A483
         TabIndex        =   17
         Top             =   2640
         Width           =   5520
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":AE5A
         Height          =   4665
         Index           =   17
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":AE6F
         TabIndex        =   18
         Top             =   2640
         Width           =   5625
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":B846
         Height          =   4770
         Index           =   18
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":B85B
         TabIndex        =   19
         Top             =   2640
         Width           =   5625
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":C232
         Height          =   4770
         Index           =   19
         Left            =   -74880
         OleObjectBlob   =   "frmScreenText.frx":C247
         TabIndex        =   20
         Top             =   2640
         Width           =   5835
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":CC1E
         Height          =   4770
         Index           =   20
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":CC33
         TabIndex        =   21
         Top             =   2640
         Width           =   5625
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":D60A
         Height          =   4650
         Index           =   21
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":D61F
         TabIndex        =   22
         Top             =   2640
         Width           =   5625
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":DFF6
         Height          =   4770
         Index           =   22
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":E00B
         TabIndex        =   23
         Top             =   2640
         Width           =   5625
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":E9E2
         Height          =   4770
         Index           =   23
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":E9F7
         TabIndex        =   24
         Top             =   2640
         Width           =   5625
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":F3CE
         Height          =   4785
         Index           =   24
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":F3E3
         TabIndex        =   25
         Top             =   2640
         Width           =   5625
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":FDBA
         Height          =   4785
         Index           =   25
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":FDCF
         TabIndex        =   26
         Top             =   2640
         Width           =   5580
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":107A6
         Height          =   4815
         Index           =   26
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":107BB
         TabIndex        =   27
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":11192
         Height          =   4695
         Index           =   27
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":111A7
         TabIndex        =   28
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":11B7E
         Height          =   4695
         Index           =   28
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":11B93
         TabIndex        =   29
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":1256A
         Height          =   4695
         Index           =   29
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":1257F
         TabIndex        =   30
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":12F56
         Height          =   4695
         Index           =   30
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":12F6B
         TabIndex        =   31
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":13942
         Height          =   4695
         Index           =   31
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":13957
         TabIndex        =   32
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":1432E
         Height          =   4695
         Index           =   32
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":14343
         TabIndex        =   33
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":14D1A
         Height          =   4815
         Index           =   33
         Left            =   -74880
         OleObjectBlob   =   "frmScreenText.frx":14D2F
         TabIndex        =   34
         Top             =   2640
         Width           =   5805
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":15706
         Height          =   4770
         Index           =   34
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":1571B
         TabIndex        =   35
         Top             =   2640
         Width           =   5610
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":160F2
         Height          =   4770
         Index           =   35
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":16107
         TabIndex        =   36
         Top             =   2640
         Width           =   5580
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":16ADE
         Height          =   4770
         Index           =   36
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":16AF3
         TabIndex        =   37
         Top             =   2640
         Width           =   5580
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":174CA
         Height          =   4725
         Index           =   37
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":174DF
         TabIndex        =   38
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":17EB6
         Height          =   4725
         Index           =   38
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":17ECB
         TabIndex        =   39
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":188A2
         Height          =   4725
         Index           =   39
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":188B7
         TabIndex        =   40
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":1928E
         Height          =   4725
         Index           =   40
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":192A3
         TabIndex        =   41
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":19C7A
         Height          =   4725
         Index           =   41
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":19C8F
         TabIndex        =   42
         Top             =   2640
         Width           =   5565
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":1A666
         Height          =   4725
         Index           =   42
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":1A67B
         TabIndex        =   43
         Top             =   2640
         Width           =   5610
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":1B052
         Height          =   4725
         Index           =   43
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":1B067
         TabIndex        =   44
         Top             =   2640
         Width           =   5610
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmScreenText.frx":1BA3E
         Height          =   1125
         Index           =   44
         Left            =   -74760
         OleObjectBlob   =   "frmScreenText.frx":1BA53
         TabIndex        =   45
         Top             =   2640
         Width           =   5610
      End
   End
End
Attribute VB_Name = "frmScreenText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    On Error Resume Next
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    Data5.Refresh
    Data6.Refresh
    Data7.Refresh
    Data8.Refresh
    Data9.Refresh
    Data10.Refresh
    Data11.Refresh
    Data12.Refresh
    Data13.Refresh
    Data14.Refresh
    Data15.Refresh
    Data16.Refresh
    Data17.Refresh
    Data18.Refresh
    Data19.Refresh
    Data20.Refresh
    Data21.Refresh
    Data22.Refresh
    Data23.Refresh
    Data24.Refresh
    Data25.Refresh
    Data26.Refresh
    Data27.Refresh
    Data28.Refresh
    Data29.Refresh
    Data30.Refresh
    Data31.Refresh
    Data32.Refresh
    Data33.Refresh
    Data34.Refresh
    Data35.Refresh
    Data36.Refresh
    Data37.Refresh
    Data38.Refresh
    Data39.Refresh
    Data40.Refresh
    Data41.Refresh
    Data42.Refresh
    Data43.Refresh
    Data44.Refresh
    Data45.Refresh
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Data1.DatabaseName = dbKidLangTxt
    Data2.DatabaseName = dbKidLangTxt
    Data3.DatabaseName = dbKidLangTxt
    Data4.DatabaseName = dbKidLangTxt
    Data5.DatabaseName = dbKidLangTxt
    Data6.DatabaseName = dbKidLangTxt
    Data7.DatabaseName = dbKidLangTxt
    Data8.DatabaseName = dbKidLangTxt
    Data9.DatabaseName = dbKidLangTxt
    Data10.DatabaseName = dbKidLangTxt
    Data11.DatabaseName = dbKidLangTxt
    Data12.DatabaseName = dbKidLangTxt
    Data13.DatabaseName = dbKidLangTxt
    Data14.DatabaseName = dbKidLangTxt
    Data15.DatabaseName = dbKidLangTxt
    Data16.DatabaseName = dbKidLangTxt
    Data17.DatabaseName = dbKidLangTxt
    Data18.DatabaseName = dbKidLangTxt
    Data19.DatabaseName = dbKidLangTxt
    Data20.DatabaseName = dbKidLangTxt
    Data21.DatabaseName = dbKidLangTxt
    Data22.DatabaseName = dbKidLangTxt
    Data23.DatabaseName = dbKidLangTxt
    Data24.DatabaseName = dbKidLangTxt
    Data25.DatabaseName = dbKidLangTxt
    Data26.DatabaseName = dbKidLangTxt
    Data27.DatabaseName = dbKidLangTxt
    Data28.DatabaseName = dbKidLangTxt
    Data29.DatabaseName = dbKidLangTxt
    Data30.DatabaseName = dbKidLangTxt
    Data31.DatabaseName = dbKidLangTxt
    Data32.DatabaseName = dbKidLangTxt
    Data33.DatabaseName = dbKidLangTxt
    Data34.DatabaseName = dbKidLangTxt
    Data35.DatabaseName = dbKidLangTxt
    Data36.DatabaseName = dbKidLangTxt
    Data37.DatabaseName = dbKidLangTxt
    Data38.DatabaseName = dbKidLangTxt
    Data39.DatabaseName = dbKidLangTxt
    Data40.DatabaseName = dbKidLangTxt
    Data41.DatabaseName = dbKidLangTxt
    Data42.DatabaseName = dbKidLangTxt
    Data43.DatabaseName = dbKidLangTxt
    Data44.DatabaseName = dbKidLangTxt
    Data45.DatabaseName = dbKidLangTxt
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmScreenText: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Data1.Recordset.Close
    Data2.Recordset.Close
    Data3.Recordset.Close
    Data4.Recordset.Close
    Data5.Recordset.Close
    Data6.Recordset.Close
    Data7.Recordset.Close
    Data8.Recordset.Close
    Data9.Recordset.Close
    Data10.Recordset.Close
    Data11.Recordset.Close
    Data12.Recordset.Close
    Data13.Recordset.Close
    Data14.Recordset.Close
    Data15.Recordset.Close
    Data16.Recordset.Close
    Data17.Recordset.Close
    Data18.Recordset.Close
    Data19.Recordset.Close
    Data20.Recordset.Close
    Data21.Recordset.Close
    Data22.Recordset.Close
    Data23.Recordset.Close
    Data24.Recordset.Close
    Data25.Recordset.Close
    Data26.Recordset.Close
    Data27.Recordset.Close
    Data28.Recordset.Close
    Data29.Recordset.Close
    Data30.Recordset.Close
    Data31.Recordset.Close
    Data32.Recordset.Close
    Data33.Recordset.Close
    Data34.Recordset.Close
    Data35.Recordset.Close
    Data36.Recordset.Close
    Data37.Recordset.Close
    Data38.Recordset.Close
    Data39.Recordset.Close
    Data40.Recordset.Close
    Data41.Recordset.Close
    Data42.Recordset.Close
    Data43.Recordset.Close
    Data44.Recordset.Close
    Data45.Recordset.Close
    Set frmScreenText = Nothing
End Sub
