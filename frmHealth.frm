VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHealth 
   BackColor       =   &H0000C0C0&
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   8235
   Begin TabDlg.SSTab Tab1 
      Height          =   6495
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Infant (0 -1year)"
      TabPicture(0)   =   "frmHealth.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Tab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Childhood"
      TabPicture(1)   =   "frmHealth.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Tab3"
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab Tab2 
         Height          =   5655
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   9975
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Health Controls"
         TabPicture(0)   =   "frmHealth.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "List2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "rsHealthControlInfant"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Vaccinations"
         TabPicture(1)   =   "frmHealth.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(1)"
         Tab(1).Control(1)=   "rsVaccinationInfant"
         Tab(1).Control(2)=   "List3"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Illnes"
         TabPicture(2)   =   "frmHealth.frx":0070
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(2)"
         Tab(2).Control(1)=   "List4"
         Tab(2).Control(2)=   "rsIllnessInfant"
         Tab(2).ControlCount=   3
         Begin VB.Frame Frame1 
            Height          =   5055
            Index           =   2
            Left            =   -73680
            TabIndex        =   34
            Top             =   480
            Width           =   5655
            Begin VB.TextBox Date1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "IllnessDate"
               DataSource      =   "rsIllnessInfant"
               Height          =   315
               Index           =   2
               Left            =   4200
               TabIndex        =   6
               Top             =   360
               Width           =   1215
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Called Doctor ?"
               DataField       =   "IllnessDoctor"
               DataSource      =   "rsIllnessInfant"
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   8
               Top             =   3960
               Width           =   1455
            End
            Begin VB.ComboBox cmbDoctor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "IllnessDoctorName"
               DataSource      =   "rsIllnessInfant"
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   960
               TabIndex        =   9
               Top             =   4635
               Width           =   3375
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               DataField       =   "IllnessNote"
               DataSource      =   "rsIllnessInfant"
               Height          =   3045
               Index           =   2
               Left            =   915
               TabIndex        =   7
               Top             =   840
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   5371
               _Version        =   393217
               BackColor       =   16777152
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmHealth.frx":008C
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Date:"
               Height          =   270
               Index           =   2
               Left            =   3000
               TabIndex        =   37
               Top             =   360
               Width           =   1020
            End
            Begin VB.Label Label5 
               Caption         =   "Symptoms:"
               Height          =   270
               Index           =   0
               Left            =   840
               TabIndex        =   36
               Top             =   600
               Width           =   1200
            End
            Begin VB.Label Label6 
               Caption         =   "Doctor Name:"
               Enabled         =   0   'False
               Height          =   270
               Index           =   0
               Left            =   960
               TabIndex        =   35
               Top             =   4440
               Width           =   1200
            End
            Begin VB.Image Image1 
               Height          =   900
               Index           =   4
               Left            =   120
               Picture         =   "frmHealth.frx":0161
               Stretch         =   -1  'True
               Top             =   2115
               Width           =   735
            End
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   4905
            Left            =   -74880
            TabIndex        =   33
            Top             =   600
            Width           =   1035
         End
         Begin VB.Data rsIllnessInfant 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -75000
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "IllnessInfant"
            Top             =   360
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Frame Frame1 
            Height          =   4815
            Index           =   1
            Left            =   -73680
            TabIndex        =   28
            Top             =   600
            Width           =   5655
            Begin VB.TextBox Date1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "VaccinationDate"
               DataSource      =   "rsVaccinationInfant"
               Height          =   315
               Index           =   1
               Left            =   3120
               TabIndex        =   2
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "VaccinationWhere"
               DataSource      =   "rsVaccinationInfant"
               Height          =   300
               Index           =   3
               Left            =   3120
               MaxLength       =   50
               TabIndex        =   4
               Top             =   1080
               Width           =   2310
            End
            Begin VB.ComboBox cmbDoctor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "VaccinationByDoctor"
               DataSource      =   "rsVaccinationInfant"
               Height          =   315
               Index           =   1
               Left            =   3120
               TabIndex        =   3
               Top             =   720
               Width           =   2415
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               DataField       =   "VaccinationNote"
               DataSource      =   "rsVaccinationInfant"
               Height          =   2850
               Index           =   1
               Left            =   90
               TabIndex        =   5
               Top             =   1845
               Width           =   5445
               _ExtentX        =   9604
               _ExtentY        =   5027
               _Version        =   393217
               BackColor       =   16777152
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmHealth.frx":2762
            End
            Begin VB.Label Label2 
               Caption         =   "Note:"
               Height          =   270
               Index           =   1
               Left            =   90
               TabIndex        =   32
               Top             =   1560
               Width           =   660
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Date:"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   31
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Vaccinated by (Doctor/Nurse):"
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   30
               Top             =   720
               Width           =   2235
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Place (Clinic name):"
               Height          =   270
               Index           =   0
               Left            =   1200
               TabIndex        =   29
               Top             =   1080
               Width           =   1770
            End
            Begin VB.Image Image1 
               Height          =   870
               Index           =   3
               Left            =   120
               Picture         =   "frmHealth.frx":2837
               Stretch         =   -1  'True
               Top             =   120
               Width           =   750
            End
         End
         Begin VB.Data rsVaccinationInfant 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -75000
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "VaccinationInfant"
            Top             =   5280
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ListBox List3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   4905
            Left            =   -74880
            TabIndex        =   27
            Top             =   600
            Width           =   1035
         End
         Begin VB.Data rsHealthControlInfant 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "HealthControlInfant"
            Top             =   480
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   4710
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1140
         End
         Begin VB.Frame Frame1 
            Height          =   4935
            Index           =   0
            Left            =   1440
            TabIndex        =   23
            Top             =   600
            Width           =   5655
            Begin VB.TextBox Date1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "ControlDate"
               DataSource      =   "rsHealthControlInfant"
               Height          =   315
               Index           =   0
               Left            =   4200
               TabIndex        =   0
               Top             =   360
               Width           =   1215
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               DataField       =   "ControlNote"
               DataSource      =   "rsHealthControlInfant"
               Height          =   3735
               Index           =   0
               Left            =   120
               TabIndex        =   1
               Top             =   1080
               Width           =   5400
               _ExtentX        =   9525
               _ExtentY        =   6588
               _Version        =   393217
               BackColor       =   16777152
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmHealth.frx":4E38
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Date:"
               Height          =   255
               Index           =   0
               Left            =   3120
               TabIndex        =   26
               Top             =   360
               Width           =   960
            End
            Begin VB.Label Label2 
               Caption         =   "Note:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   840
               Width           =   615
            End
            Begin VB.Image Image1 
               Height          =   855
               Index           =   0
               Left            =   720
               Picture         =   "frmHealth.frx":4F0D
               Stretch         =   -1  'True
               Top             =   120
               Width           =   690
            End
         End
      End
      Begin TabDlg.SSTab Tab3 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   22
         Top             =   600
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   9975
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Health Controls"
         TabPicture(0)   =   "frmHealth.frx":8155
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "rsHealthControlChild"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "List5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame1(3)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Vaccinations"
         TabPicture(1)   =   "frmHealth.frx":8171
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1(4)"
         Tab(1).Control(1)=   "List6"
         Tab(1).Control(2)=   "rsVaccinationChild"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Illnes"
         TabPicture(2)   =   "frmHealth.frx":818D
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame1(5)"
         Tab(2).Control(1)=   "List7"
         Tab(2).Control(2)=   "rsIllnessChild"
         Tab(2).ControlCount=   3
         Begin VB.Frame Frame1 
            Height          =   5055
            Index           =   5
            Left            =   -73920
            TabIndex        =   49
            Top             =   480
            Width           =   5895
            Begin VB.TextBox Date1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "IllnessDate"
               DataSource      =   "rsIllnessChild"
               Height          =   315
               Index           =   5
               Left            =   4560
               TabIndex        =   16
               Top             =   240
               Width           =   1215
            End
            Begin VB.ComboBox cmbDoctor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "IllnessDoctorName"
               DataSource      =   "rsIllnessChild"
               Enabled         =   0   'False
               Height          =   315
               Index           =   3
               Left            =   960
               TabIndex        =   19
               Top             =   4635
               Width           =   3285
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Called Doctor ?"
               DataField       =   "IllnessDoctor"
               DataSource      =   "rsIllnessChild"
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   18
               Top             =   4080
               Width           =   1425
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               DataField       =   "IllnessNote"
               DataSource      =   "rsIllnessChild"
               Height          =   3120
               Index           =   5
               Left            =   900
               TabIndex        =   17
               Top             =   870
               Width           =   4860
               _ExtentX        =   8573
               _ExtentY        =   5503
               _Version        =   393217
               BackColor       =   16777152
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmHealth.frx":81A9
            End
            Begin VB.Label Label6 
               Caption         =   "Doctor Name:"
               Enabled         =   0   'False
               Height          =   270
               Index           =   1
               Left            =   960
               TabIndex        =   52
               Top             =   4440
               Width           =   3210
            End
            Begin VB.Label Label5 
               Caption         =   "Symptoms:"
               Height          =   270
               Index           =   1
               Left            =   840
               TabIndex        =   51
               Top             =   600
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Date:"
               Height          =   270
               Index           =   5
               Left            =   3480
               TabIndex        =   50
               Top             =   240
               Width           =   990
            End
            Begin VB.Image Image1 
               Height          =   900
               Index           =   5
               Left            =   120
               Picture         =   "frmHealth.frx":827E
               Stretch         =   -1  'True
               Top             =   1920
               Width           =   720
            End
         End
         Begin VB.ListBox List7 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   4905
            Left            =   -74880
            TabIndex        =   48
            Top             =   600
            Width           =   855
         End
         Begin VB.Data rsIllnessChild 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -74880
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "IllnessChild"
            Top             =   360
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Frame Frame1 
            Height          =   5055
            Index           =   4
            Left            =   -73800
            TabIndex        =   43
            Top             =   360
            Width           =   5895
            Begin VB.TextBox Date1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "VaccinationDate"
               DataSource      =   "rsVaccinationChild"
               Height          =   315
               Index           =   4
               Left            =   3480
               TabIndex        =   12
               Top             =   360
               Width           =   1215
            End
            Begin VB.ComboBox cmbDoctor 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "VaccinationByDoctor"
               DataSource      =   "rsVaccinationChild"
               Height          =   315
               Index           =   2
               Left            =   3480
               TabIndex        =   13
               Top             =   840
               Width           =   2235
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "VaccinationWhere"
               DataSource      =   "rsVaccinationChild"
               Height          =   285
               Index           =   5
               Left            =   3480
               MaxLength       =   50
               TabIndex        =   14
               Top             =   1320
               Width           =   2160
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               DataField       =   "VaccinationNote"
               DataSource      =   "rsVaccinationChild"
               Height          =   2865
               Index           =   4
               Left            =   195
               TabIndex        =   15
               Top             =   2040
               Width           =   5580
               _ExtentX        =   9843
               _ExtentY        =   5054
               _Version        =   393217
               BackColor       =   16777152
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmHealth.frx":A87F
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Place (Clinic name):"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   47
               Top             =   1320
               Width           =   2955
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Vaccinated by (Doctor/Nurse):"
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   46
               Top             =   840
               Width           =   2535
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Date:"
               Height          =   255
               Index           =   4
               Left            =   2400
               TabIndex        =   45
               Top             =   360
               Width           =   945
            End
            Begin VB.Label Label2 
               Caption         =   "Note:"
               Height          =   255
               Index           =   3
               Left            =   210
               TabIndex        =   44
               Top             =   1800
               Width           =   600
            End
            Begin VB.Image Image1 
               Height          =   855
               Index           =   2
               Left            =   120
               Picture         =   "frmHealth.frx":A954
               Stretch         =   -1  'True
               Top             =   360
               Width           =   690
            End
         End
         Begin VB.ListBox List6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   4905
            Left            =   -74880
            TabIndex        =   42
            Top             =   480
            Width           =   915
         End
         Begin VB.Data rsVaccinationChild 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   -74880
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "VaccinationChild"
            Top             =   360
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Frame Frame1 
            Height          =   4935
            Index           =   3
            Left            =   1200
            TabIndex        =   39
            Top             =   480
            Width           =   5895
            Begin VB.TextBox Date1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "ControlDate"
               DataSource      =   "rsHealthControlChild"
               Height          =   315
               Index           =   3
               Left            =   4440
               TabIndex        =   10
               Top             =   360
               Width           =   1215
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               DataField       =   "ControlNote"
               DataSource      =   "rsHealthControlChild"
               Height          =   3465
               Index           =   3
               Left            =   120
               TabIndex        =   11
               Top             =   1320
               Width           =   5595
               _ExtentX        =   9869
               _ExtentY        =   6112
               _Version        =   393217
               BackColor       =   16777152
               Enabled         =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmHealth.frx":CF55
            End
            Begin VB.Label Label2 
               Caption         =   "Note:"
               Height          =   225
               Index           =   2
               Left            =   120
               TabIndex        =   41
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Date:"
               Height          =   225
               Index           =   3
               Left            =   3480
               TabIndex        =   40
               Top             =   360
               Width           =   825
            End
            Begin VB.Image Image1 
               Height          =   825
               Index           =   1
               Left            =   120
               Picture         =   "frmHealth.frx":D02A
               Stretch         =   -1  'True
               Top             =   120
               Width           =   660
            End
         End
         Begin VB.ListBox List5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   4905
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Width           =   930
         End
         Begin VB.Data rsHealthControlChild 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   120
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "HealthControlChild"
            Top             =   360
            Visible         =   0   'False
            Width           =   1365
         End
      End
   End
End
Attribute VB_Name = "frmHealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim h1RecordBookmark() As Variant, h2RecordBookmark() As Variant
Dim va1RecordBookmark() As Variant, va2RecordBookmark() As Variant
Dim il1RecordBookmark() As Variant, il2RecordBookmark() As Variant
Dim rsMidwife As Recordset
Dim rsLanguage As Recordset
Private Sub SelectTab()
    On Error Resume Next
    With Me
        Select Case iTab
        Case 0
            .Tab1.Tab = 0
            .BackColor = &HC0C0&
            .Tab1.BackColor = &HC0C0&
            FillList2
            List2.ListIndex = 0
        Case 1
            .Tab1.Tab = 1
            .BackColor = &H4040&
            .Tab1.BackColor = &H4040&
            FillList5
            List5.ListIndex = 0
        Case Else
        End Select
    End With
End Sub

Public Sub DeleteHealth()
    On Error Resume Next
    Select Case Tab1.Tab
        Case 0
            Select Case Tab2.Tab
            Case 0
                rsHealthControlInfant.Recordset.Delete
                FillList2
            Case 1
                rsVaccinationInfant.Recordset.Delete
                FillList3
            Case 2
                rsIllnessInfant.Recordset.Delete
                FillList4
            Case Else
            End Select
        Case 1
            Select Case Tab3.Tab
            Case 0
                rsHealthControlChild.Recordset.Delete
                FillList5
            Case 1
                rsVaccinationChild.Recordset.Delete
                FillList6
            Case 2
                rsIllnessChild.Recordset.Delete
                FillList7
            Case Else
            End Select
        Case Else
        End Select
End Sub

Public Sub NewHealth()
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0  'infant
        Select Case Tab2.Tab
        Case 0
            rsHealthControlInfant.Recordset.AddNew
            Date1(0).SetFocus
            boolNewRecord = True
        Case 1
            rsVaccinationInfant.Recordset.AddNew
            Date1(1).SetFocus
            boolNewRecord = True
        Case 2
            rsIllnessInfant.Recordset.AddNew
            Date1(2).SetFocus
            boolNewRecord = True
        Case Else
        End Select
    Case 1  'child
        Select Case Tab3.Tab
        Case 0
            rsHealthControlChild.Recordset.AddNew
            Date1(3).SetFocus
            boolNewRecord = True
        Case 1
            rsVaccinationChild.Recordset.AddNew
            Date1(4).SetFocus
            boolNewRecord = True
        Case 2
            rsIllnessChild.Recordset.AddNew
            Date1(5).SetFocus
            boolNewRecord = True
        Case Else
        End Select
    Case Else
    End Select
End Sub

Public Function FillList2() As Boolean
    On Error GoTo errFillList2
    List2.Clear
    With rsHealthControlInfant.Recordset
        .MoveLast
        .MoveFirst
        ReDim h1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List2.AddItem CDate(.Fields("ControlDate"))
            List2.ItemData(List2.NewIndex) = List2.ListCount - 1
            h1RecordBookmark(List2.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    FillList2 = True
    Exit Function
    
errFillList2:
    FillList2 = False
    Err.Clear
End Function
Public Function FillList3() As Boolean
    On Error GoTo errFillList3
    List3.Clear
    With rsVaccinationInfant.Recordset
        .MoveLast
        .MoveFirst
        ReDim va1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List3.AddItem CDate(.Fields("VaccinationDate"))
            List3.ItemData(List3.NewIndex) = List3.ListCount - 1
            va1RecordBookmark(List3.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    FillList3 = True
    Exit Function
        
errFillList3:
    FillList3 = False
    Err.Clear
End Function

Public Function FillList4() As Boolean
    On Error GoTo errFillList4
    List4.Clear
    With rsIllnessInfant.Recordset
        .MoveLast
        .MoveFirst
        ReDim il1RecordBookmark(.RecordCount)
        Do While Not .EOF
            List4.AddItem CDate(.Fields("IllnessDate"))
            List4.ItemData(List4.NewIndex) = List4.ListCount - 1
            il1RecordBookmark(List4.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    FillList4 = True
    Exit Function
        
errFillList4:
    FillList4 = False
    Err.Clear
End Function

Public Function FillList5() As Boolean
    On Error GoTo errFillList5
    List5.Clear
        With rsHealthControlChild.Recordset
        .MoveLast
        .MoveFirst
        ReDim h2RecordBookmark(.RecordCount)
        Do While Not .EOF
            List5.AddItem CDate(.Fields("ControlDate"))
            List5.ItemData(List5.NewIndex) = List5.ListCount - 1
            h2RecordBookmark(List5.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    FillList5 = True
    Exit Function
        
errFillList5:
    FillList5 = False
    Err.Clear
End Function

Public Function FillList6() As Boolean
    On Error GoTo errFillList6
    List6.Clear
    With rsVaccinationChild.Recordset
        .MoveLast
        .MoveFirst
        ReDim va2RecordBookmark(.RecordCount)
        Do While Not .EOF
            List6.AddItem CDate(.Fields("VaccinationDate"))
            List6.ItemData(List6.NewIndex) = List6.ListCount - 1
            va2RecordBookmark(List6.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    FillList6 = True
    Exit Function
        
errFillList6:
    FillList6 = False
    Err.Clear
End Function

Public Function FillList7() As Boolean
    On Error GoTo errFillList7
    List7.Clear
    With rsIllnessChild.Recordset
        .MoveLast
        .MoveFirst
        ReDim il2RecordBookmark(.RecordCount)
        Do While Not .EOF
            List7.AddItem CDate(.Fields("IllnessDate"))
            List7.ItemData(List7.NewIndex) = List7.ListCount - 1
            il2RecordBookmark(List7.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    FillList7 = True
    Exit Function
        
errFillList7:
    FillList7 = False
    Err.Clear
End Function
Private Sub LoadDoctor()
    cmbDoctor(0).Clear
    cmbDoctor(1).Clear
    cmbDoctor(2).Clear
    cmbDoctor(3).Clear
    With rsMidwife
        .MoveFirst
        Do While Not .EOF
            cmbDoctor(0).AddItem .Fields("FirstName") & "" & "  " & .Fields("LastName") & ""
            cmbDoctor(1).AddItem .Fields("FirstName") & "" & "  " & .Fields("LastName") & ""
            cmbDoctor(2).AddItem .Fields("FirstName") & "" & "  " & .Fields("LastName") & ""
            cmbDoctor(3).AddItem .Fields("FirstName") & "" & "  " & .Fields("LastName") & ""
        .MoveNext
        Loop
    End With
End Sub

Public Sub SelectHealthChild()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM HealthControlChild WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsHealthControlChild.RecordSource = Sql
    rsHealthControlChild.Refresh
    
    Sql = "SELECT * FROM HealthControlInfant WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsHealthControlInfant.RecordSource = Sql
    rsHealthControlInfant.Refresh
    
    Sql = "SELECT * FROM IllnessChild WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsIllnessChild.RecordSource = Sql
    rsIllnessChild.Refresh
    
    Sql = "SELECT * FROM IllnessInfant WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsIllnessInfant.RecordSource = Sql
    rsIllnessInfant.Refresh
    
    Sql = "SELECT * FROM VaccinationChild WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsVaccinationChild.RecordSource = Sql
    rsVaccinationChild.Refresh
    
    Sql = "SELECT * FROM VaccinationInfant WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsVaccinationInfant.RecordSource = Sql
    rsVaccinationInfant.Refresh
End Sub
Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    On Error Resume Next
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1")
                    Label1(1).Caption = .Fields("label1")
                    Label1(2).Caption = .Fields("label1")
                    Label1(3).Caption = .Fields("label1")
                    Label1(4).Caption = .Fields("label1")
                    Label1(5).Caption = .Fields("label1")
                End If
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2(0).Caption
                Else
                    Label2(0).Caption = .Fields("label2")
                    Label2(1).Caption = .Fields("label2")
                    Label2(2).Caption = .Fields("label2")
                    Label2(3).Caption = .Fields("label2")
                End If
                If IsNull(.Fields("label3")) Then
                    .Fields("label3") = Label3(0).Caption
                Else
                    Label3(0).Caption = .Fields("label3")
                    Label3(1).Caption = .Fields("label3")
                End If
                If IsNull(.Fields("label4")) Then
                    .Fields("label4") = Label4(0).Caption
                Else
                    Label4(0).Caption = .Fields("label4")
                    Label4(1).Caption = .Fields("label4")
                End If
                If IsNull(.Fields("label5")) Then
                    .Fields("label5") = Label5(0).Caption
                Else
                    Label5(0).Caption = .Fields("label5")
                    Label5(1).Caption = .Fields("label5")
                End If
                If IsNull(.Fields("label6")) Then
                    .Fields("label6") = Label6(0).Caption
                Else
                    Label6(0).Caption = .Fields("label6")
                    Label6(1).Caption = .Fields("label6")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 0
                Tab2.Tab = 0
                Tab3.Tab = 0
                If IsNull(.Fields("Tab20")) Then
                    .Fields("Tab20") = Tab2.Caption
                Else
                    Tab2.Caption = .Fields("Tab20")
                    Tab3.Caption = .Fields("Tab20")
                End If
                Tab2.Tab = 1
                Tab3.Tab = 1
                If IsNull(.Fields("Tab21")) Then
                    .Fields("Tab21") = Tab2.Caption
                Else
                    Tab2.Caption = .Fields("Tab21")
                    Tab3.Caption = .Fields("Tab21")
                End If
                Tab2.Tab = 2
                Tab3.Tab = 2
                If IsNull(.Fields("Tab22")) Then
                    .Fields("Tab22") = Tab2.Caption
                Else
                    Tab2.Caption = .Fields("Tab22")
                    Tab3.Caption = .Fields("Tab22")
                End If
                Tab2.Tab = 0
                Tab3.Tab = 0
                If IsNull(.Fields("Check1")) Then
                    .Fields("Check1") = Check1(0).Caption
                Else
                    Check1(0).Caption = .Fields("Check1")
                    Check1(1).Caption = .Fields("Check1")
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
        .Fields("label1") = Label1(0).Caption
        .Fields("label2") = Label2(0).Caption
        .Fields("label3") = Label3(0).Caption
        .Fields("label4") = Label4(0).Caption
        .Fields("label5") = Label5(0).Caption
        .Fields("label6") = Label6(0).Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 0
        Tab2.Tab = 0
        .Fields("Tab20") = Tab2.Caption
        Tab2.Tab = 1
        .Fields("Tab21") = Tab2.Caption
        Tab2.Tab = 2
        .Fields("Tab22") = Tab2.Caption
        Tab2.Tab = 0
        .Fields("Check1") = Check1(0).Caption
        .Fields("FormName1") = "Health Baby"
        .Fields("FormName2") = "Vaccination Baby"
        .Fields("FormName3") = "Illness Baby"
        .Fields("FormName4") = "Health Childhood"
        .Fields("FormName5") = "Vaccination Childhood"
        .Fields("FormName6") = "Illness Childhood"
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Yes") = "Yes"
        .Fields("No") = "No"
        .Fields("Help") = strHelp
        .Update
    End With
End Sub
Public Sub WriteHealth()
    On Error Resume Next
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
    
    Select Case Tab1.Tab
    Case 0  'infant
        Select Case Tab2.Tab
        Case 0  'health control
            sHeader = rsLanguage.Fields("FormName")
            Call PrintFront
            'health infant
            With rsHealthControlInfant.Recordset
                .MoveFirst
                Do While Not .EOF
                    If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                        cPrint.FontBold = True
                        cPrint.pPrint Label1(0).Caption, 1, True
                        cPrint.pPrint Format(CDate(.Fields("ControlDate")), "dd.mm.yyyy"), 3.5
                        cPrint.FontBold = False
                        cPrint.pPrint Label2(0).Caption, 1, True
                        If Len(RichTextBox1(0).Text) <> 0 Then
                            cPrint.pMultiline RichTextBox1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint "", 3.5
                        End If
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
        Case 1  ' vaccination
            sHeader = rsLanguage.Fields("FormName2")
            Call PrintFront
            
            'vaccinations infant
            With rsVaccinationInfant.Recordset
                .MoveFirst
                Do While Not .EOF
                    If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                        cPrint.FontBold = True
                        cPrint.pPrint Label1(1).Caption, 1, True
                        cPrint.pPrint Format(CDate(.Fields("VaccinationDate")), "dd.mm.yyyy"), 3.5
                        cPrint.FontBold = False
                        cPrint.pPrint Label3(0).Caption, 1, True
                        If Not IsNull(.Fields("VaccinationByDoctor")) Then
                            cPrint.pPrint .Fields("VaccinationByDoctor"), 3.5
                        Else
                            cPrint.pPrint "", 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                        cPrint.pPrint Label4(0).Caption, 1, True
                        If Not IsNull(.Fields("VaccinationWhere")) Then
                            cPrint.pPrint .Fields("VaccinationWhere"), 3.5
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                        cPrint.pPrint Label2(1).Caption, 1, True
                        If Len(RichTextBox1(1).Text) <> 0 Then
                            cPrint.pMultiline RichTextBox1(1).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
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
        Case 2  'sickness
            sHeader = rsLanguage.Fields("FormName3")
            Call PrintFront
            
            With rsIllnessInfant.Recordset
                .MoveFirst
                Do While Not .EOF
                    If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                        cPrint.FontBold = True
                        cPrint.pPrint Label1(2).Caption, 1, True
                        If IsDate(.Fields("IllnessDate")) Then
                            cPrint.pPrint Format(CDate(.Fields("IllnessDate")), "dd.mm.yyyy"), 3.5
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
                        cPrint.FontBold = False
                        cPrint.pPrint rsLanguage.Fields("label5"), 1, True
                        If Len(RichTextBox1(2).Text) <> 0 Then
                            cPrint.pMultiline RichTextBox1(2).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
                        If CBool(.Fields("IllnessDoctor")) = 1 Then
                            cPrint.pPrint Check1(0).Caption, 1, True
                            cPrint.pPrint rsLanguage.Fields("Yes"), 3.5
                            cPrint.pPrint Label6(0).Caption, 1, True
                            cPrint.pPrint Format(.Fields("IllnessDoctorName")), 3.5
                        Else
                            cPrint.pPrint Check1(0).Caption, 1, True
                            cPrint.pPrint rsLanguage.Fields("No"), 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                    End If
                .MoveNext
                Loop
            End With
        Case Else
        End Select
        
    Case 1  'childhood
        Select Case Tab3.Tab
        Case 0  'health control
            sHeader = rsLanguage.Fields("FormName4")
            Call PrintFront
            
            'health childhood
            With rsHealthControlChild.Recordset
                .MoveFirst
                Do While Not .EOF
                    If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                        DoEvents
                        cPrint.FontBold = True
                        cPrint.pPrint Label1(3).Caption, 1, True
                        cPrint.pPrint Format(CDate(.Fields("ControlDate")), "dd.mm.yyyy"), 3.5
                        cPrint.FontBold = False
                        cPrint.pPrint Label2(2).Caption, 1, True
                        If Len(RichTextBox1(3).Text) <> 0 Then
                            cPrint.pMultiline RichTextBox1(3).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint "", 3.5
                        End If
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
        Case 1  'vaccination
            sHeader = rsLanguage.Fields("FormName5")
            Call PrintFront
            
            'vaccinations childhood
            With rsVaccinationChild.Recordset
                .MoveFirst
                Do While Not .EOF
                    If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                        cPrint.FontBold = True
                        cPrint.pPrint Label1(4).Caption, 1, True
                        cPrint.pPrint Format(CDate(.Fields("VaccinationDate")), "dd.mm.yyyy"), 3.5
                        cPrint.FontBold = False
                        cPrint.pPrint Label3(1).Caption, 1, True
                        If Not IsNull(.Fields("VaccinationByDoctor")) Then
                            cPrint.pPrint .Fields("VaccinationByDoctor"), 3.5
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                        cPrint.pPrint Label4(1).Caption, 1, True
                        If Not IsNull(.Fields("VaccinationWhere")) Then
                            cPrint.pPrint Format(.Fields("VaccinationWhere")), 3.5
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                        cPrint.pPrint Label2(3).Caption, 1, True
                        If Len(RichTextBox1(4).Text) <> 0 Then
                            cPrint.pMultiline RichTextBox1(4).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
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
        Case 2  'sicknes
            sHeader = rsLanguage.Fields("FormName6")
            Call PrintFront
            
            'illness childhood
            With rsIllnessChild.Recordset
                .MoveFirst
                Do While Not .EOF
                    If CLng(.Fields("ChildNo")) = CLng(glChildNo) Then
                        cPrint.FontBold = True
                        cPrint.pPrint Label1(5).Caption, 1, True
                        If IsDate(.Fields("IllnessDate")) Then
                            cPrint.pPrint Format(CDate(.Fields("IllnessDate")), "dd.mm.yyyy"), 3.5
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                        cPrint.FontBold = False
                        cPrint.pPrint Label5(1).Caption, 1, True
                        If Len(RichTextBox1(5).Text) <> 0 Then
                            cPrint.pMultiline RichTextBox1(5).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
                        Else
                            cPrint.pPrint " ", 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                        If CBool(.Fields("IllnessDoctor")) = 1 Then
                            cPrint.pPrint Check1(1).Caption, 1, True
                            cPrint.pPrint rsLanguage.Fields("Yes"), 3.5
                            cPrint.pPrint Label6(1).Caption, 1, True
                            cPrint.pPrint Format(.Fields("IllnessDoctorName")), 3.5
                        Else
                            cPrint.pPrint Check1(1).Caption, 1, True
                            cPrint.pPrint rsLanguage.Fields("No"), 3.5
                        End If
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            Call PrintFront
                        End If
                    End If
                .MoveNext
                Loop
            End With
        Case Else
        End Select
    Case Else
    End Select
    
    Screen.MousePointer = vbDefault
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub

Public Sub WriteHealthWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName1"))
    With wdApp
        'health infant
        rsHealthControlInfant.Recordset.MoveFirst
        Do While Not rsHealthControlInfant.Recordset.EOF
            If CLng(rsHealthControlInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsHealthControlInfant.Recordset.Fields("ControlDate"), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label2(0).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(0).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        rsHealthControlInfant.Recordset.MoveNext
        Loop
        
        .ActiveWindow.Selection.InsertBreak Type:=wdPageBreak
        .ActiveWindow.Selection.Font.Name = "Monotype Corsiva"
        .ActiveWindow.Selection.Font.Size = 28
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .Selection.TypeText Text:=rsLanguage.Fields("FormName2")
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .ActiveWindow.Selection.Font.Name = "Times New Roman"
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.TypeParagraph
        .Selection.MoveDown Unit:=wdLine, Count:=1
    
        'write vaccinations
        rsVaccinationInfant.Recordset.MoveFirst
        Do While Not rsVaccinationInfant.Recordset.EOF
            If CLng(rsVaccinationInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(1).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(rsVaccinationInfant.Recordset.Fields("VaccinationDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label3(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsVaccinationInfant.Recordset.Fields("VaccinationByDoctor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label4(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsVaccinationInfant.Recordset.Fields("VaccinationWhere"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label2(1).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(1).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        rsVaccinationInfant.Recordset.MoveNext
        Loop
        
        .ActiveWindow.Selection.InsertBreak Type:=wdPageBreak
        .ActiveWindow.Selection.Font.Name = "Monotype Corsiva"
        .ActiveWindow.Selection.Font.Size = 28
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .Selection.TypeText Text:=rsLanguage.Fields("FormName3")
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .ActiveWindow.Selection.Font.Name = "Times New Roman"
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.TypeParagraph
        .Selection.MoveDown Unit:=wdLine, Count:=1
        
        'write illness infant
        rsIllnessInfant.Recordset.MoveFirst
        Do While Not rsIllnessInfant.Recordset.EOF
            If CLng(rsIllnessInfant.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(2).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(rsIllnessInfant.Recordset.Fields("IllnessDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label5(0).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(2).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                If CBool(rsIllnessInfant.Recordset.Fields("IllnessDoctor")) = 1 Then
                     .Selection.TypeText Text:=Check1(0).Caption
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=rsLanguage.Fields("Yes")
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=Label6(0).Caption
                    .Selection.TypeText Text:=Format(rsIllnessInfant.Recordset.Fields("IllnessDoctorName"))
                Else
                    .Selection.TypeText Text:=Check1(0).Caption
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=rsLanguage.Fields("No")
                    .Selection.MoveRight Unit:=wdCell
                End If
            End If
        rsIllnessInfant.Recordset.MoveNext
        Loop
        
        .ActiveWindow.Selection.InsertBreak Type:=wdPageBreak
        .ActiveWindow.Selection.Font.Name = "Monotype Corsiva"
        .ActiveWindow.Selection.Font.Size = 28
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .Selection.TypeText Text:="Childhood "
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .ActiveWindow.Selection.Font.Name = "Times New Roman"
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.TypeParagraph
        .Selection.MoveDown Unit:=wdLine, Count:=1
        
        'CHILDHOOD
        
        'health
        rsHealthControlChild.Recordset.MoveFirst
        Do While Not rsHealthControlChild.Recordset.EOF
            If CLng(rsHealthControlChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsHealthControlChild.Recordset.Fields("ControlDate"), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label2(0).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(3).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        rsHealthControlChild.Recordset.MoveNext
        Loop
    
        .ActiveWindow.Selection.InsertBreak Type:=wdPageBreak
        .ActiveWindow.Selection.Font.Name = "Monotype Corsiva"
        .ActiveWindow.Selection.Font.Size = 28
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .Selection.TypeText Text:=rsLanguage.Fields("FormName5")
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .ActiveWindow.Selection.Font.Name = "Times New Roman"
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.TypeParagraph
        .Selection.MoveDown Unit:=wdLine, Count:=1
    
        'vaccinations
        rsVaccinationChild.Recordset.MoveFirst
        Do While Not rsVaccinationChild.Recordset.EOF
            If CLng(rsVaccinationChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(1).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(rsVaccinationChild.Recordset.Fields("VaccinationDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label3(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsVaccinationChild.Recordset.Fields("VaccinationByDoctor"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label4(0).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(rsVaccinationChild.Recordset.Fields("VaccinationWhere"))
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label2(1).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(4).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
                .Selection.MoveRight Unit:=wdCell
            End If
        rsVaccinationChild.Recordset.MoveNext
        Loop
    
        .ActiveWindow.Selection.InsertBreak Type:=wdPageBreak
        .ActiveWindow.Selection.Font.Name = "Monotype Corsiva"
        .ActiveWindow.Selection.Font.Size = 28
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .Selection.TypeText Text:=rsLanguage.Fields("FormName6")
        .ActiveWindow.Selection.Font.Bold = wdToggle
        .ActiveWindow.Selection.Font.Name = "Times New Roman"
        .ActiveWindow.Selection.Font.Size = 10
        .Selection.TypeParagraph
        .Selection.MoveDown Unit:=wdLine, Count:=1
    
        'illness
        rsIllnessChild.Recordset.MoveFirst
        Do While Not rsIllnessChild.Recordset.EOF
            If CLng(rsIllnessChild.Recordset.Fields("ChildNo")) = CLng(glChildNo) Then
                .Selection.TypeText Text:=Label1(2).Caption
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Format(CDate(rsIllnessChild.Recordset.Fields("IllnessDate")), "dd.mm.yyyy")
                .Selection.MoveRight Unit:=wdCell
                .Selection.TypeText Text:=Label5(0).Caption
                .Selection.MoveRight Unit:=wdCell
                Clipboard.Clear
                Clipboard.SetText frmHealth.RichTextBox1(5).TextRTF, vbCFRTF
                .Selection.Paste
                .Selection.MoveRight Unit:=wdCell
                If CBool(rsIllnessChild.Recordset.Fields("IllnessDoctor")) = 1 Then
                     .Selection.TypeText Text:=Check1(0).Caption
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=rsLanguage.Fields("Yes")
                     .Selection.MoveRight Unit:=wdCell
                     .Selection.TypeText Text:=Label6(0).Caption
                    .Selection.TypeText Text:=Format(rsIllnessChild.Recordset.Fields("IllnessDoctorName"))
                Else
                    .Selection.TypeText Text:=Check1(0).Caption
                    .Selection.MoveRight Unit:=wdCell
                    .Selection.TypeText Text:=rsLanguage.Fields("No")
                End If
            End If
        rsIllnessChild.Recordset.MoveNext
        Loop
    End With
    Set wdApp = Nothing
End Sub
Private Sub Check1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        If Check1(0).Value = 1 Then
            Label6(0).Enabled = True
            cmbDoctor(0).Enabled = True
        Else
            Label6(0).Enabled = False
            cmbDoctor(0).Enabled = False
        End If
    Case 1
        If Check1(1).Value = 1 Then
            Label6(1).Enabled = True
            cmbDoctor(3).Enabled = True
        Else
            Label6(1).Enabled = False
            cmbDoctor(3).Enabled = False
        End If
    Case Else
    End Select
End Sub

Private Sub cmbDoctor_GotFocus(Index As Integer)
    On Error Resume Next
    If boolNewRecord Then
        Select Case Index
        Case 1
            With rsVaccinationInfant.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("VaccinationDate") = CDate(Format(Date1(1).Text, "dd.mm.yyyy"))
                .Update
                FillList3
                .Bookmark = .LastModified
                boolNewRecord = False
            End With
        Case 2
                With rsVaccinationChild.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("VaccinationDate") = CDate(Format(Date1(4).Text, "dd.mm.yyyy"))
                .Update
                FillList6
                .Bookmark = .LastModified
                boolNewRecord = False
            End With
        Case Else
        End Select
    End If
End Sub

Private Sub Date1_Click(Index As Integer)
Dim UserDate As Date
    If IsDate(Date1(Index).Text) Then
        UserDate = CVDate(Date1(Index).Text)
    Else
        UserDate = Format(Now, "dd.mm.yyyy")
    End If
    If frmCalendar.GetDate(UserDate) Then
        Date1(Index).Text = UserDate
    End If
End Sub

Private Sub Date1_LostFocus(Index As Integer)
    On Error Resume Next
    If boolNewRecord Then
        Select Case Index
            Case 0  'health infant
            Case 1  'vaccination infant
            Case 2  'Illness infant
            Case 3  'health child
            Case 4  'vaccination child
            Case 5  'Illness Child
            Case Else
            End Select
    End If
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    rsHealthControlInfant.Refresh
    rsVaccinationInfant.Refresh
    rsIllnessInfant.Refresh
    rsHealthControlChild.Refresh
    rsVaccinationChild.Refresh
    rsIllnessChild.Refresh
    LoadDoctor
    ShowText
    ShowAllButtons
    ShowKids
    SelectTab
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsHealthControlInfant.DatabaseName = dbKidsTxt
    rsVaccinationInfant.DatabaseName = dbKidsTxt
    rsIllnessInfant.DatabaseName = dbKidsTxt
    rsHealthControlChild.DatabaseName = dbKidsTxt
    rsVaccinationChild.DatabaseName = dbKidsTxt
    rsIllnessChild.DatabaseName = dbKidsTxt
    Set rsMidwife = dbKids.OpenRecordset("Midwife")
    Set rsLanguage = dbKidLang.OpenRecordset("frmHealth")
    iWhichForm = 23
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmHealth: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsHealthControlInfant.Recordset.Close
    rsVaccinationInfant.Recordset.Close
    rsIllnessInfant.Recordset.Close
    rsHealthControlChild.Recordset.Close
    rsVaccinationChild.Recordset.Close
    rsIllnessChild.Recordset.Close
    rsMidwife.Close
    rsLanguage.Close
    iWhichForm = 0
    HideAllButtons
    HideKids
    Set frmHealth = Nothing
End Sub

Private Sub List2_Click()
    On Error Resume Next
    rsHealthControlInfant.Recordset.Bookmark = h1RecordBookmark(List2.ItemData(List2.ListIndex))
End Sub
Private Sub List3_Click()
    On Error Resume Next
    rsVaccinationInfant.Recordset.Bookmark = va1RecordBookmark(List3.ItemData(List3.ListIndex))
End Sub
Private Sub List4_Click()
    On Error Resume Next
    rsIllnessInfant.Recordset.Bookmark = il1RecordBookmark(List4.ItemData(List4.ListIndex))
End Sub
Private Sub List5_Click()
    On Error Resume Next
    rsHealthControlChild.Recordset.Bookmark = h2RecordBookmark(List5.ItemData(List5.ListIndex))
End Sub
Private Sub List6_Click()
    On Error Resume Next
    rsVaccinationChild.Recordset.Bookmark = va2RecordBookmark(List6.ItemData(List6.ListIndex))
End Sub
Private Sub List7_Click()
    On Error Resume Next
    rsIllnessChild.Recordset.Bookmark = il2RecordBookmark(List7.ItemData(List7.ListIndex))
End Sub

Private Sub RichTextBox1_GotFocus(Index As Integer)
    If boolNewRecord Then
        Select Case Index
        Case 0
            With rsHealthControlInfant.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("ControlDate") = CDate(Format(Date1(0).Text, "dd.mm.yyyy"))
                .Update
                FillList2
                .Bookmark = .LastModified
                boolNewRecord = False
            End With
        Case 2
            With rsIllnessInfant.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("IllnessDate") = CDate(Format(Date1(2).Text, "dd.mm.yyyy"))
                .Update
                FillList4
                .Bookmark = .LastModified
                boolNewRecord = False
            End With
        Case 3
            With rsHealthControlChild.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("ControlDate") = CDate(Format(Date1(3).Text, "dd.mm.yyyy"))
                .Update
                FillList5
                .Bookmark = .LastModified
                boolNewRecord = False
            End With
        Case 5
            With rsIllnessChild.Recordset
                .Fields("ChildNo") = glChildNo
                .Fields("IllnessDate") = CDate(Format(Date1(5).Text, "dd.mm.yyyy"))
                .Update
                FillList7
                .Bookmark = .LastModified
                boolNewRecord = False
            End With
        Case Else
        End Select
    End If
End Sub

Private Sub RichTextBox1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    On Error Resume Next
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1(Index).SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub RichTextBox1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbRightButton Then
      Me.PopupMenu MDIMasterKid.mnuFormat
   End If
End Sub
Private Sub RichTextBox1_SelChange(Index As Integer)
    On Error Resume Next
    Call RichTextSelChange(frmHealth.RichTextBox1(Index))
End Sub
Private Sub Tab1_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab1.Tab
    Case 0
        FillList2
        List2.ListIndex = 0
    Case 1
        FillList5
        List5.ListIndex = 0
    Case Else
    End Select
End Sub

Private Sub Tab2_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab2.Tab
    Case 0
        FillList2
        List2.ListIndex = 0
    Case 1
        FillList3
        List3.ListIndex = 0
    Case 2
        FillList4
        List4.ListIndex = 0
    Case Else
    End Select
End Sub
Private Sub Tab3_Click(PreviousTab As Integer)
    On Error Resume Next
    Select Case Tab3.Tab
    Case 0
        FillList5
        List5.ListIndex = 0
    Case 1
        FillList6
        List6.ListIndex = 0
    Case 2
        FillList7
        List7.ListIndex = 0
    Case Else
    End Select
End Sub
