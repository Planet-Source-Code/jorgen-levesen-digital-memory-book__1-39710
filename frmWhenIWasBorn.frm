VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWhenIWasBorn 
   BackColor       =   &H00008000&
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   8880
   Begin VB.Frame Frame5 
      BackColor       =   &H00008000&
      Caption         =   "What were the fashion at the time"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   4320
      TabIndex        =   52
      Top             =   6240
      Width           =   4335
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "Fashions"
         DataSource      =   "rsWhenBorn"
         Height          =   1860
         Index           =   11
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   240
         Width           =   4110
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00008000&
      Caption         =   "Fashion Picture"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   46
      Top             =   5280
      Width           =   4095
      Begin VB.CommandButton btnScan 
         Height          =   450
         Index           =   0
         Left            =   600
         Picture         =   "frmWhenIWasBorn.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Scan a picture from Scanner"
         Top             =   1860
         Width           =   360
      End
      Begin VB.CommandButton btnDelete 
         Height          =   450
         Index           =   0
         Left            =   600
         Picture         =   "frmWhenIWasBorn.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Delete shown picture"
         Top             =   2400
         Width           =   360
      End
      Begin VB.CommandButton btnCopyPic 
         Height          =   450
         Index           =   0
         Left            =   600
         Picture         =   "frmWhenIWasBorn.frx":0294
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Copy picture to the Clipboard"
         Top             =   1320
         Width           =   360
      End
      Begin VB.CommandButton btnReadFromFile 
         Height          =   450
         Index           =   0
         Left            =   600
         Picture         =   "frmWhenIWasBorn.frx":0956
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Read Picture from a disk file"
         Top             =   780
         Width           =   360
      End
      Begin VB.CommandButton btnPastePicture 
         Height          =   450
         Index           =   0
         Left            =   600
         Picture         =   "frmWhenIWasBorn.frx":1018
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Paste Picture from the Clipboard"
         Top             =   240
         Width           =   360
      End
      Begin VB.Image Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "FashionPic"
         DataSource      =   "rsWhenBorn"
         Height          =   2655
         Left            =   1155
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2640
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Caption         =   "Exchange rates"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   4320
      TabIndex        =   27
      Top             =   3120
      Width           =   4335
      Begin VB.ComboBox cmbCurrency 
         DataField       =   "ExRateCountry3"
         DataSource      =   "rsWhenBorn"
         Height          =   315
         Index           =   2
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox cmbCurrency 
         DataField       =   "ExRateCountry2"
         DataSource      =   "rsWhenBorn"
         Height          =   315
         Index           =   1
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbCurrency 
         DataField       =   "ExRateCountry1"
         DataSource      =   "rsWhenBorn"
         Height          =   315
         Index           =   0
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ExRateCountry3Val"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   10
         Left            =   2430
         TabIndex        =   14
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ExRateCountry2Val"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   9
         Left            =   2430
         TabIndex        =   12
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ExRateCountry1Val"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   8
         Left            =   2430
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ExRateDollar"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   7
         Left            =   2430
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ExRateMark"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   6
         Left            =   2430
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "ExRateYen"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   5
         Left            =   2430
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   9
         Left            =   3795
         TabIndex        =   45
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   8
         Left            =   3795
         TabIndex        =   44
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   7
         Left            =   3795
         TabIndex        =   43
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   6
         Left            =   3795
         TabIndex        =   42
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   5
         Left            =   3795
         TabIndex        =   41
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   4
         Left            =   3795
         TabIndex        =   40
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Against"
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
         Index           =   5
         Left            =   3225
         TabIndex        =   39
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Against"
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
         Index           =   4
         Left            =   3225
         TabIndex        =   38
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Against"
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
         Index           =   3
         Left            =   3225
         TabIndex        =   37
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Against"
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
         Index           =   2
         Left            =   3225
         TabIndex        =   36
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Against"
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
         Left            =   3225
         TabIndex        =   35
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Against"
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
         Left            =   3225
         TabIndex        =   34
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange rate"
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
         Index           =   9
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange rate"
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
         Index           =   8
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange rate"
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
         Index           =   7
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange rate Dollar:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   2085
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange rate D-Mark:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exchange rate Yen:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   2085
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "The Cost Of ..."
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   4320
      TabIndex        =   18
      Top             =   720
      Width           =   4335
      Begin VB.ComboBox cmbDim 
         DataField       =   "CostPetrolim"
         DataSource      =   "rsWhenBorn"
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "NormalWageHour"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   5
         Top             =   1560
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "CostPetrol"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   3
         Top             =   1200
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "CostNewspaper"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   2
         Top             =   840
         Width           =   765
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "CostBread"
         DataSource      =   "rsWhenBorn"
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   3
         Left            =   2760
         TabIndex        =   26
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Index           =   2
         Left            =   2760
         TabIndex        =   25
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Left            =   2760
         TabIndex        =   24
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         DataField       =   "Currency"
         DataSource      =   "rsMyRecord"
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
         Left            =   2760
         TabIndex        =   23
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wage an hour:"
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
         Index           =   3
         Left            =   600
         TabIndex        =   22
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Petrol:"
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
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "A Newspaper:"
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
         Left            =   600
         TabIndex        =   20
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "One Bread:"
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
         Left            =   600
         TabIndex        =   19
         Top             =   480
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "The Newspapers wrote about.."
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   4095
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataField       =   "NewsPapers"
         DataSource      =   "rsWhenBorn"
         Height          =   3900
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   480
         Width           =   3840
      End
   End
   Begin VB.Data rsWhenBorn 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   90
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "WhenBorn"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsMyRecord 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   90
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MyRecord"
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "When I Was Born Then ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   90
      TabIndex        =   16
      Top             =   120
      Width           =   8550
   End
End
Attribute VB_Name = "frmWhenIWasBorn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Dim rsCountry As Recordset
Dim rsVolum As Recordset
Dim boolFirst As Boolean
Public Sub DeleteRecord()
    On Error Resume Next
    rsWhenBorn.Recordset.Delete
    SelectBorn
End Sub
Public Sub NewBorn()
    rsWhenBorn.Recordset.AddNew
    Text1(0).BackColor = &HC0E0FF
    Text1(0).SetFocus
    boolNewRecord = True
End Sub
Public Sub WriteWhenIWasBorn()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("FormName"))
    With wdApp
        .Selection.TypeText Text:=Frame1.Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame2.Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(1).Text & "  " & Label3(0).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(2).Text & "  " & Label3(1).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(2).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(3).Text & "  " & Label3(2).Caption & " " & cmbDim.Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(4).Text & "  " & Label3(3).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame3.Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(5).Text & "  " & Label4(0).Caption & "  " & Label3(4).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(6).Text & "  " & Label4(1).Caption & "  " & Label3(5).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(6).Text & "  " & Label4(2).Caption & "  " & Label3(6).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(7).Caption & "  " & cmbCurrency(0).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(8).Text & "  " & Label4(3).Caption & "  " & Label3(7).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(8).Caption & "  " & cmbCurrency(1).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(9).Text & "  " & Label4(4).Caption & "  " & Label3(8).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Label2(9).Caption & "  " & cmbCurrency(2).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(10).Text & "  " & Label4(5).Caption & "  " & Label3(9).Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame5.Caption
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Text1(11).Text
        .Selection.MoveRight Unit:=wdCell
        .Selection.TypeText Text:=Frame4.Caption
        Clipboard.Clear
        Clipboard.SetData Picture1.Picture, vbCFBitmap
        .Selection.Paste
    End With
    Set wdApp = Nothing
End Sub

Public Sub PrintWhenIWasBorn()
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
    sHeader = rsLanguage.Fields("FormName")
    
    cPrint.pStartDoc
    Call PrintFront
    
    cPrint.pPrint Frame1.Caption, 1, True
    cPrint.pMultiline Text1(0).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint Frame2.Caption, 1
    cPrint.pPrint
    cPrint.pPrint Label2(0).Caption, 1, True
    If Len(Text1(1).Text) Then
        cPrint.pPrint Text1(1).Text & "  " & Label3(0).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint Label2(1).Caption, 1, True
    If Len(Text1(2).Text) <> 0 Then
        cPrint.pPrint Text1(2).Text & "  " & Label3(1).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label2(2).Caption, 1, True
    If Len(Text1(2).Text) <> 0 Then
        cPrint.pPrint Text1(2).Text & "  " & Label3(2).Caption & "  " & cmbDim.Text, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label2(3).Caption, 1, True
    If Len(Text1(3).Text) <> 0 Then
        cPrint.pPrint Text1(3).Text & "  " & Label3(3).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint Frame3.Caption, 1    'exchange rates
    cPrint.FontBold = False
    cPrint.pPrint
    cPrint.pPrint Label2(4).Caption, 1, True
    If Len(Text1(5).Text) <> 0 Then
        cPrint.pPrint Text1(5).Text & "  " & Label4(0).Caption & "  " & Label3(4).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint Label2(5).Caption, 1, True
    If Len(Text1(6).Text) <> 0 Then
        cPrint.pPrint Text1(6).Text & "  " & Label4(1).Caption & "  " & Label3(5).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint rsLanguage.Fields("label2(6)"), 1, True
    If Len(Text1(7).Text) <> 0 Then
        cPrint.pPrint Text1(7).Text & "  " & Label4(2).Caption & "  " & Label3(6).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label2(7).Caption & " " & cmbCurrency(0).Text, 1, True
    If Len(Text1(8).Text) <> 0 Then
        cPrint.pPrint Text1(8).Text & "  " & Label4(3).Caption & "  " & Label3(7).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label2(7).Caption & " " & cmbCurrency(1).Text, 1, True
    If Len(Text1(9).Text) <> 0 Then
        cPrint.pPrint Text1(9).Text & "  " & Label4(4).Caption & "  " & Label3(8).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint Label2(7).Caption & " " & cmbCurrency(2).Text, 1, True
    If Len(Text1(10).Text) <> 0 Then
        cPrint.pPrint Text1(10).Text & "  " & Label4(5).Caption & "  " & Label3(9).Caption, 3.5
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint rsLanguage.Fields("Frame5"), 1    'how was the fashion
    cPrint.FontBold = False
    cPrint.pPrint
    If Len(Text1(11).Text) <> 0 Then
        cPrint.pMultiline Text1(11).Text, 3.5, cPrint.GetPaperWidth - 1.2, , False, True
    Else
        cPrint.pPrint " ", 3.5
    End If
    cPrint.pPrint
    If cPrint.pEndOfPage Then
        cPrint.pFooter
        cPrint.pNewPage
        Call PrintFront
    End If
    cPrint.pPrint
    cPrint.pPrint Frame4.Caption, 1    'fashion picture
    If Not IsNull(rsWhenBorn.Recordset.Fields("FashionPic")) Then
        cPrint.pPrintPicture Picture1.Picture, 1, cPrint.CurrentY, cPrint.GetPaperWidth - 2, cPrint.GetPaperHeight - cPrint.CurrentY - 0.5, False, True
    End If
    
    Screen.MousePointer = vbDefault
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub

Private Sub LoadCurrency()
    With rsCountry
        .MoveFirst
        Do While Not .EOF
            If Not IsNull(.Fields("Currency")) Then
                If Not IsNull(.Fields("Currency")) And Not .Fields("Currency") = "N/A" Then
                    cmbCurrency(0).AddItem .Fields("Currency")
                    cmbCurrency(1).AddItem .Fields("Currency")
                    cmbCurrency(2).AddItem .Fields("Currency")
                End If
            End If
        .MoveNext
        Loop
    End With
End Sub
Private Sub LoadDim()
    cmbDim.Clear
    With rsVolum
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbDim.AddItem .Fields("VolumeDim")
            End If
        .MoveNext
        Loop
    End With
End Sub

Public Function SelectBorn() As Boolean
Dim Sql As String
    On Error GoTo errSelectBorn
    Sql = "SELECT * FROM WhenBorn WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsWhenBorn.RecordSource = Sql
    rsWhenBorn.Refresh
    rsWhenBorn.Recordset.MoveFirst
    SelectBorn = True
    Exit Function
    
errSelectBorn:
    SelectBorn = False
    Err.Clear
End Function

Private Sub ShowText()
Dim strHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                .Edit
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("label1")
                End If
                For n = 0 To 9
                    If IsNull(.Fields(n + 2)) Then
                        .Fields(n + 2) = Label2(n).Caption
                    Else
                        Label2(n).Caption = .Fields(n + 2)
                    End If
                Next
                If IsNull(.Fields("label4")) Then
                    .Fields("label4") = Label4(0).Caption
                Else
                    Label4(0).Caption = .Fields("label4")
                    Label4(1).Caption = .Fields("label4")
                    Label4(2).Caption = .Fields("label4")
                    Label4(3).Caption = .Fields("label4")
                    Label4(4).Caption = .Fields("label4")
                    Label4(5).Caption = .Fields("label4")
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
                If IsNull(.Fields("Frame4")) Then
                    .Fields("Frame4") = Frame4.Caption
                Else
                    Frame4.Caption = .Fields("Frame4")
                End If
                If IsNull(.Fields("Frame5")) Then
                    .Fields("Frame5") = Frame5.Caption
                Else
                    Frame5.Caption = .Fields("Frame5")
                End If
                If IsNull(.Fields("btnPastePicture")) Then
                    .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
                Else
                    btnPastePicture(0).ToolTipText = .Fields("btnPastePicture")
                End If
                If IsNull(.Fields("btnReadFromFile")) Then
                    .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
                Else
                    btnReadFromFile(0).ToolTipText = .Fields("btnReadFromFile")
                End If
                If IsNull(.Fields("btnCopyPic")) Then
                    .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
                Else
                    btnCopyPic(0).ToolTipText = .Fields("btnCopyPic")
                End If
                If IsNull(.Fields("btnScan")) Then
                    .Fields("btnScan") = btnScan(0).ToolTipText
                Else
                    btnScan(0).ToolTipText = .Fields("btnScan")
                End If
                If IsNull(.Fields("btnDelete")) Then
                    .Fields("btnDelete") = btnDelete(0).ToolTipText
                Else
                    btnDelete(0).ToolTipText = .Fields("btnDelete")
                End If
                'If IsNull(.Fields("GridColn0")) Then
                    '.Fields("GridColn0") = cmbCurrency(0).Columns(0).Caption
                'Else
                    'cmbCurrency(0).Columns(0).Caption = .Fields("GridColn0")
                    'cmbCurrency(1).Columns(0).Caption = .Fields("GridColn0")
                    'cmbCurrency(2).Columns(0).Caption = .Fields("GridColn0")
                'End If
                'If IsNull(.Fields("GridColn1")) Then
                    '.Fields("GridColn1") = cmbCurrency(0).Columns(1).Caption
                'Else
                    'cmbCurrency(0).Columns(1).Caption = .Fields("GridColn1")
                    'cmbCurrency(1).Columns(1).Caption = .Fields("GridColn1")
                    'cmbCurrency(2).Columns(1).Caption = .Fields("GridColn1")
                'End If
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
        .Fields("label1") = Label1.Caption
        For n = 0 To 9
            .Fields(n + 2) = Label2(n).Caption
        Next
        .Fields("label4") = Label4(0).Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("Frame2") = Frame2.Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("Frame4") = Frame4.Caption
        .Fields("Frame5") = Frame5.Caption
        .Fields("btnPastePicture") = btnPastePicture(0).ToolTipText
        .Fields("btnReadFromFile") = btnReadFromFile(0).ToolTipText
        .Fields("btnCopyPic") = btnCopyPic(0).ToolTipText
        .Fields("btnScan") = btnScan(0).ToolTipText
        .Fields("btnDelete") = btnDelete(0).ToolTipText
        '.Fields("GridColn0") = cmbCurrency(0).Columns(0).Caption
        '.Fields("GridColn1") = cmbCurrency(0).Columns(1).Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub btnCopyPic_Click(Index As Integer)
    On Error Resume Next
    Clipboard.SetData Picture1.Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click(Index As Integer)
    On Error Resume Next
    Picture1.Picture = LoadPicture()
End Sub

Private Sub btnPastePicture_Click(Index As Integer)
        On Error Resume Next
        Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub btnReadFromFile_Click(Index As Integer)
        On Error Resume Next
        With Cmd1
            .filename = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "Pictures (*.bmp; *.pcx;*.jpg;*.jpeg;*.gif)|*.bmp;*.pcx;*.jpg;*.jpeg;*.gif"
            .FilterIndex = 1
            .Action = 1
        End With
        Set Picture1.Picture = LoadPicture(Cmd1.filename)
End Sub

Private Sub btnScan_Click(Index As Integer)
    Dim ret As Long, t As Single
    On Error Resume Next
    ret = TWAIN_AcquireToClipboard(Me.hWnd, t)
    Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Not boolFirst Then Exit Sub
    rsWhenBorn.Refresh
    rsMyRecord.Refresh
    LoadCurrency
    LoadDim
    LoadCurrency
    ShowText
    ShowAllButtons
    ShowKids
    boolFirst = False
    SelectBorn
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsMyRecord.DatabaseName = dbKidsTxt
    rsWhenBorn.DatabaseName = dbKidsTxt
    Set rsCountry = dbKids.OpenRecordset("Country")
    Set rsVolum = dbKids.OpenRecordset("DimVolume")
    Set rsLanguage = dbKidLang.OpenRecordset("frmWhenIWasBorn")
    iWhichForm = 43
    boolFirst = True
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRecord.Recordset.Close
    rsWhenBorn.UpdateRecord
    rsWhenBorn.Recordset.Close
    rsCountry.Close
    rsVolum.Close
    rsLanguage.Close
    iWhichForm = 0
    HideAllButtons
    HideKids
    Set frmWhenIWasBorn = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    onGotFocus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo errLostFocus
    Select Case Index
    Case 0
    If boolNewRecord Then
        With rsWhenBorn.Recordset
            .Fields("ChildNo") = glChildNo
            .Fields("NewsPapers") = Text1(0).Text
            .Update
            boolNewRecord = False
            .Bookmark = .LastModified
            Text1(0).BackColor = &HFFFFFF
            Text1(1).SetFocus
        End With
    End If
    Case Else
    End Select
    Exit Sub
    
errLostFocus:
    Beep
    MsgBox Err.Description, vbCritical, "New Record"
    Resume errLostFocus2
errLostFocus2:
End Sub


