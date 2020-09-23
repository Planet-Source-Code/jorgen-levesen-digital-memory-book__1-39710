VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmWeightLength 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Weight and Height"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   13
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   49344
      TabCaption(0)   =   "At Birth"
      TabPicture(0)   =   "frmWeightLength.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Date1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "1. Month"
      TabPicture(1)   =   "frmWeightLength.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image1(1)"
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).Control(2)=   "Label2(1)"
      Tab(1).Control(3)=   "Label3(1)"
      Tab(1).Control(4)=   "Text2(1)"
      Tab(1).Control(5)=   "Text1(1)"
      Tab(1).Control(6)=   "Date1(1)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "2. Month"
      TabPicture(2)   =   "frmWeightLength.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1(2)"
      Tab(2).Control(1)=   "Label1(2)"
      Tab(2).Control(2)=   "Label2(2)"
      Tab(2).Control(3)=   "Label3(2)"
      Tab(2).Control(4)=   "Text2(2)"
      Tab(2).Control(5)=   "Text1(2)"
      Tab(2).Control(6)=   "Date1(2)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "3. Month"
      TabPicture(3)   =   "frmWeightLength.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image1(3)"
      Tab(3).Control(1)=   "Label1(3)"
      Tab(3).Control(2)=   "Label2(3)"
      Tab(3).Control(3)=   "Label3(3)"
      Tab(3).Control(4)=   "Text2(3)"
      Tab(3).Control(5)=   "Text1(3)"
      Tab(3).Control(6)=   "Date1(3)"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "4. Month"
      TabPicture(4)   =   "frmWeightLength.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image1(4)"
      Tab(4).Control(1)=   "Label1(4)"
      Tab(4).Control(2)=   "Label2(4)"
      Tab(4).Control(3)=   "Label3(4)"
      Tab(4).Control(4)=   "Text2(4)"
      Tab(4).Control(5)=   "Text1(4)"
      Tab(4).Control(6)=   "Date1(4)"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "5.Month"
      TabPicture(5)   =   "frmWeightLength.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Image1(5)"
      Tab(5).Control(1)=   "Label1(5)"
      Tab(5).Control(2)=   "Label2(5)"
      Tab(5).Control(3)=   "Label3(5)"
      Tab(5).Control(4)=   "Text2(5)"
      Tab(5).Control(5)=   "Text1(5)"
      Tab(5).Control(6)=   "Date1(5)"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "6. Month"
      TabPicture(6)   =   "frmWeightLength.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Image1(6)"
      Tab(6).Control(1)=   "Label1(6)"
      Tab(6).Control(2)=   "Label2(6)"
      Tab(6).Control(3)=   "Label3(6)"
      Tab(6).Control(4)=   "Text2(6)"
      Tab(6).Control(5)=   "Text1(6)"
      Tab(6).Control(6)=   "Date1(6)"
      Tab(6).ControlCount=   7
      TabCaption(7)   =   "7.Month"
      TabPicture(7)   =   "frmWeightLength.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Image1(7)"
      Tab(7).Control(1)=   "Label1(7)"
      Tab(7).Control(2)=   "Label2(7)"
      Tab(7).Control(3)=   "Label3(7)"
      Tab(7).Control(4)=   "Text2(7)"
      Tab(7).Control(5)=   "Text1(7)"
      Tab(7).Control(6)=   "Date1(7)"
      Tab(7).ControlCount=   7
      TabCaption(8)   =   "8. Month"
      TabPicture(8)   =   "frmWeightLength.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Image1(8)"
      Tab(8).Control(1)=   "Label1(8)"
      Tab(8).Control(2)=   "Label2(8)"
      Tab(8).Control(3)=   "Label3(8)"
      Tab(8).Control(4)=   "Text2(8)"
      Tab(8).Control(5)=   "Text1(8)"
      Tab(8).Control(6)=   "Date1(8)"
      Tab(8).ControlCount=   7
      TabCaption(9)   =   "9. Month"
      TabPicture(9)   =   "frmWeightLength.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Image1(9)"
      Tab(9).Control(1)=   "Label1(9)"
      Tab(9).Control(2)=   "Label2(9)"
      Tab(9).Control(3)=   "Label3(9)"
      Tab(9).Control(4)=   "Text2(9)"
      Tab(9).Control(5)=   "Text1(9)"
      Tab(9).Control(6)=   "Date1(9)"
      Tab(9).ControlCount=   7
      TabCaption(10)  =   "10.Month"
      TabPicture(10)  =   "frmWeightLength.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Image1(10)"
      Tab(10).Control(1)=   "Label1(10)"
      Tab(10).Control(2)=   "Label2(10)"
      Tab(10).Control(3)=   "Label3(10)"
      Tab(10).Control(4)=   "Text2(10)"
      Tab(10).Control(5)=   "Text1(10)"
      Tab(10).Control(6)=   "Date1(10)"
      Tab(10).ControlCount=   7
      TabCaption(11)  =   "11.Month"
      TabPicture(11)  =   "frmWeightLength.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Image1(11)"
      Tab(11).Control(1)=   "Label1(11)"
      Tab(11).Control(2)=   "Label2(11)"
      Tab(11).Control(3)=   "Label3(11)"
      Tab(11).Control(4)=   "Text2(11)"
      Tab(11).Control(5)=   "Text1(11)"
      Tab(11).Control(6)=   "Date1(11)"
      Tab(11).ControlCount=   7
      TabCaption(12)  =   "12.Month"
      TabPicture(12)  =   "frmWeightLength.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Image1(12)"
      Tab(12).Control(1)=   "Label1(12)"
      Tab(12).Control(2)=   "Label2(12)"
      Tab(12).Control(3)=   "Label3(12)"
      Tab(12).Control(4)=   "Text2(12)"
      Tab(12).Control(5)=   "Text1(12)"
      Tab(12).Control(6)=   "Date1(12)"
      Tab(12).ControlCount=   7
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TvelveMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   12
         Left            =   -73320
         TabIndex        =   82
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TvelveMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   12
         Left            =   -73320
         TabIndex        =   78
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TvelveMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   12
         Left            =   -73320
         TabIndex        =   77
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ElevenMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   11
         Left            =   -73320
         TabIndex        =   76
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ElevenMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   11
         Left            =   -73320
         TabIndex        =   72
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ElevenMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   11
         Left            =   -73320
         TabIndex        =   71
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TenthMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   10
         Left            =   -73320
         TabIndex        =   70
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TenthMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   10
         Left            =   -73320
         TabIndex        =   66
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TenthMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   10
         Left            =   -73320
         TabIndex        =   65
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "NineMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   9
         Left            =   -73320
         TabIndex        =   64
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "NineMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   9
         Left            =   -73320
         TabIndex        =   60
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "NineMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   9
         Left            =   -73320
         TabIndex        =   59
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "EightMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   8
         Left            =   -73320
         TabIndex        =   58
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "EightMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   8
         Left            =   -73320
         TabIndex        =   54
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "EightMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   8
         Left            =   -73320
         TabIndex        =   53
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "SevenMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   7
         Left            =   -73320
         TabIndex        =   52
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "SevenMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   7
         Left            =   -73320
         TabIndex        =   48
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "SevenMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   7
         Left            =   -73320
         TabIndex        =   47
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "SixMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   6
         Left            =   -73320
         TabIndex        =   46
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "SixMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   6
         Left            =   -73320
         TabIndex        =   42
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "SevenMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   6
         Left            =   -73320
         TabIndex        =   41
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FiveMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   5
         Left            =   -73320
         TabIndex        =   40
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FiveMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   5
         Left            =   -73320
         TabIndex        =   36
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FiveMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   5
         Left            =   -73320
         TabIndex        =   35
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FourMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   4
         Left            =   -73320
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FourMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   4
         Left            =   -73320
         TabIndex        =   30
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "FourMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   4
         Left            =   -73320
         TabIndex        =   29
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ThreeMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   3
         Left            =   -73320
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ThreeMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   3
         Left            =   -73320
         TabIndex        =   24
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "ThreeMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   3
         Left            =   -73320
         TabIndex        =   23
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TwoMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   2
         Left            =   -73320
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TwoMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   2
         Left            =   -73320
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "TwoMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   2
         Left            =   -73320
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "OneMonthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   1
         Left            =   -73320
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "OneMonthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   1
         Left            =   -73320
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "OneMonthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   1
         Left            =   -73320
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Date1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "AtBirthDate"
         DataSource      =   "rsWeightLength"
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
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "AtBirthLength"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "AtBirthWeight"
         DataSource      =   "rsWeightLength"
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   83
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   81
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   80
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   79
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   12
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":016C
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   75
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   74
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   73
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   11
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":1F0B
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   69
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   68
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   10
         Left            =   -74760
         TabIndex        =   67
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   10
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":35F5
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   63
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   62
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   61
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   9
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":4D1C
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   57
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   56
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   55
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   8
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":6567
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   51
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   50
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   49
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   7
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":856C
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   45
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   44
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   6
         Left            =   -74760
         TabIndex        =   43
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   6
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":99F7
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   39
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   38
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   37
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   5
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":B152
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   33
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   32
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   31
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   4
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":C7D1
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   27
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   3
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":E011
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   21
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   20
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   19
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   2
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":F6FB
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   1
         Left            =   -71880
         Picture         =   "frmWeightLength.frx":1110F
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Index           =   0
         Left            =   3120
         Picture         =   "frmWeightLength.frx":12F33
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbLength 
      BackColor       =   &H00FFFFC0&
      DataField       =   "LengthDim"
      DataSource      =   "rsWeightLength"
      Height          =   315
      Left            =   4200
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ComboBox cmbWeight 
      BackColor       =   &H00FFFFC0&
      DataField       =   "WeightDim"
      DataSource      =   "rsWeightLength"
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Data rsWeightLength 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\Master\MasterKid\MasterKid.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "WeightLength"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "Height Dimension:"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Weight Dimension:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "frmWeightLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Dim rsLength As Recordset
Dim rsWeight As Recordset
Public Sub NewWeightLength()
    On Error Resume Next
    boolNewRecord = True
    rsWeightLength.Recordset.AddNew
    Date1(0).SetFocus
End Sub
Public Sub WriteWeightLength()
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
    sHeader = rsLanguage.Fields("Form")
    
    cPrint.pStartDoc
    Call PrintFront
    
    For n = 0 To 12
        If IsDate(Date1(n).Text) Then
            cPrint.pPrint Label1(n).Caption, 1, True
            If Len(Date1(n).Text) <> 0 Then
                cPrint.pPrint Format(CDate(Date1(n).Text), "dd.mm.yyyy"), 3.5
            Else
                cPrint.pPrint " ", 3.5
            End If
            cPrint.pPrint Label2(n).Caption & "  " & Text1(n).Text & "  " & cmbLength.Text, 1
            cPrint.pPrint Label3(n).Caption & "  " & Text2(n).Text & "  " & cmbWeight.Text, 1
            cPrint.pPrint
            If cPrint.pEndOfPage Then
                cPrint.pFooter
                cPrint.pNewPage
                Call PrintFront
            End If
        End If
    Next
    
    Screen.MousePointer = vbDefault
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Call Form_Activate
End Sub

Public Sub WriteWeightLengthWord()
    On Error Resume Next
    WriteHeader (rsLanguage.Fields("Form"))
    With wdApp
        For n = 0 To 12
            .Selection.TypeText Text:=Label1(n).Caption
            .Selection.MoveRight Unit:=wdCell
            If IsDate(Date1(n).Text) Then
                .Selection.TypeText Text:=Format(CDate(Date1(n).Text), "dd.mm.yyyy")
            Else
                .Selection.TypeText Text:=" "
            End If
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label2(0).Caption & "  " & Text1(n).Text & " " & cmbLength.Text
            .Selection.MoveRight Unit:=wdCell
            .Selection.TypeText Text:=Label3(0).Caption & "  " & Text2(n).Text & " " & cmbWeight.Text
            .Selection.MoveRight Unit:=wdCell
        Next
    End With
    Set wdApp = Nothing
End Sub

Public Function SelectChild() As Boolean
Dim Sql As String
    On Error GoTo errSelectChild
    Sql = "SELECT * FROM WeightLength WHERE CLng(ChildNo) ="
    Sql = Sql & Chr(34) & CLng(glChildNo) & Chr(34)
    rsWeightLength.RecordSource = Sql
    rsWeightLength.Refresh
    rsWeightLength.Recordset.MoveFirst
    SelectChild = True
    Exit Function
    
errSelectChild:
    SelectChild = False
    Err.Clear
End Function

Private Sub ReadText()
Dim strMemo As String
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
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1(0).Caption
                Else
                    For i = 0 To 12
                        Label1(i).Caption = .Fields("label1")
                    Next
                End If
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2(0).Caption
                Else
                    For i = 0 To 12
                        Label2(i).Caption = .Fields("label2")
                    Next
                End If
                If IsNull(.Fields("label3")) Then
                    .Fields("label3") = Label3(0).Caption
                Else
                    For i = 0 To 12
                        Label3(i).Caption = .Fields("label3")
                    Next
                End If
                If IsNull(.Fields("label4")) Then
                    .Fields("label4") = Label4.Caption
                Else
                        Label4.Caption = .Fields("label4")
                End If
                If IsNull(.Fields("label5")) Then
                    .Fields("label5") = Label5.Caption
                Else
                        Label5.Caption = .Fields("label5")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab12")
                End If
                Tab1.Tab = 3
                If IsNull(.Fields("Tab13")) Then
                    .Fields("Tab13") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab13")
                End If
                Tab1.Tab = 4
                If IsNull(.Fields("Tab14")) Then
                    .Fields("Tab14") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab14")
                End If
                Tab1.Tab = 5
                If IsNull(.Fields("Tab15")) Then
                    .Fields("Tab15") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab15")
                End If
                Tab1.Tab = 6
                If IsNull(.Fields("Tab16")) Then
                    .Fields("Tab16") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab16")
                End If
                Tab1.Tab = 7
                If IsNull(.Fields("Tab17")) Then
                    .Fields("Tab17") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab17")
                End If
                Tab1.Tab = 8
                If IsNull(.Fields("Tab18")) Then
                    .Fields("Tab18") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab18")
                End If
                Tab1.Tab = 9
                If IsNull(.Fields("Tab19")) Then
                    .Fields("Tab19") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab19")
                End If
                Tab1.Tab = 10
                If IsNull(.Fields("Tab110")) Then
                    .Fields("Tab110") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab110")
                End If
                Tab1.Tab = 11
                If IsNull(.Fields("Tab111")) Then
                    .Fields("Tab111") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab111")
                End If
                Tab1.Tab = 12
                If IsNull(.Fields("Tab112")) Then
                    .Fields("Tab112") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab112")
                End If
                Tab1.Tab = 0
                .Update
                DBEngine.Idle dbFreeLocks
                Me.MousePointer = Default
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = "ENG" Then
                If Not IsNull(.Fields("Help")) Then
                    strMemo = .Fields("Help")
                Else
                    strMemo = " "
                End If
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = FileExt
        .Fields("Form") = Me.Caption
        .Fields("label1") = Label1(0).Caption
        .Fields("label2") = Label2(0).Caption
        .Fields("label3") = Label3(0).Caption
        .Fields("label4") = Label4.Caption
        .Fields("label5") = Label5.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 4
        .Fields("Tab14") = Tab1.Caption
        Tab1.Tab = 5
        .Fields("Tab15") = Tab1.Caption
        Tab1.Tab = 6
        .Fields("Tab16") = Tab1.Caption
        Tab1.Tab = 7
        .Fields("Tab17") = Tab1.Caption
        Tab1.Tab = 8
        .Fields("Tab18") = Tab1.Caption
        Tab1.Tab = 9
        .Fields("Tab19") = Tab1.Caption
        Tab1.Tab = 10
        .Fields("Tab110") = Tab1.Caption
        Tab1.Tab = 11
        .Fields("Tab111") = Tab1.Caption
        Tab1.Tab = 12
        .Fields("Tab112") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("sDate") = "Date: "
        .Fields("spage") = "Page: "
        .Fields("Help") = strMemo
        .Update
    End With
    DBEngine.Idle dbFreeLocks
End Sub
Private Sub LoadLengthWeight()
    On Error Resume Next
    With rsLength
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbLength.AddItem .Fields("LengthDim")
            End If
        .MoveNext
        Loop
    End With
    
    With rsWeight
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = FileExt Then
                cmbWeight.AddItem .Fields("WeightDim")
            End If
        .MoveNext
        Loop
    End With
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

Private Sub Form_Activate()
    On Error Resume Next
    rsWeightLength.Refresh
    LoadLengthWeight
    If SelectChild Then
        Label6.Caption = MDIMasterKid.cmbChildren.Text
    Else
        Label6.Caption = " "
    End If
    ReadText
    ShowAllButtons
    ShowKids
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsWeightLength.DatabaseName = dbKidsTxt
    Set rsLength = dbKids.OpenRecordset("DimLength")
    Set rsWeight = dbKids.OpenRecordset("DimWeight")
    Set rsLanguage = dbKidLang.OpenRecordset("frmWeightLength")
    iWhichForm = 41
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    WriteErrorFile Err.Description, "frmWeightLength: Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsWeightLength.Recordset.Close
    rsLength.Close
    rsWeight.Close
    rsLanguage.Close
    HideAllButtons
    HideKids
    iWhichForm = 0
    Set frmWeightLength = Nothing
End Sub
Private Sub Text1_GotFocus(Index As Integer)
   Select Case Index
   Case 0
    If boolNewRecord Then
         With rsWeightLength.Recordset
             .Fields("ChildNo") = glChildNo
             If IsDate(Date1(0).Text) Then
                 .Fields("AtBirthDate") = Format(Date1(0).Text, "dd.mm.yyyy")
             End If
             .Update
             .Bookmark = .LastModified
         End With
     End If
    Case Else
    End Select
End Sub
