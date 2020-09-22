VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   Caption         =   "Source Code Organizer V.1"
   ClientHeight    =   7515
   ClientLeft      =   1635
   ClientTop       =   1155
   ClientWidth     =   11340
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11340
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   9551
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Contents"
      TabPicture(0)   =   "frmMain.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstTitles"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "frmMain.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "cd1"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(3)=   "Frame1(1)"
      Tab(1).Control(4)=   "Picture1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Add/Del"
      TabPicture(2)   =   "frmMain.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "frameModify"
      Tab(2).Control(2)=   "frameDelete"
      Tab(2).ControlCount=   3
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4815
         Left            =   -74880
         ScaleHeight     =   4815
         ScaleWidth      =   3015
         TabIndex        =   66
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox Textweb 
            Height          =   285
            Left            =   1560
            TabIndex        =   71
            Top             =   4440
            Visible         =   0   'False
            Width           =   1455
         End
         Begin RichTextLib.RichTextBox rtbNotes 
            Height          =   3975
            Left            =   0
            TabIndex        =   67
            TabStop         =   0   'False
            ToolTipText     =   "Write Code info/Author's name or it could be anything."
            Top             =   240
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   7011
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   1.00000e5
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmMain.frx":091E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Source Code Notes/Author Information:"
            Height          =   375
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Width           =   2895
         End
      End
      Begin VB.Frame frameDelete 
         Caption         =   "Delete"
         Height          =   1995
         Left            =   -74880
         TabIndex        =   31
         Top             =   3360
         Width           =   2955
         Begin VB.CommandButton cmdDone 
            Caption         =   "Done"
            Height          =   255
            Left            =   840
            TabIndex        =   34
            Top             =   1680
            Width           =   1095
         End
         Begin VB.ListBox lstDelete 
            Height          =   840
            Left            =   240
            TabIndex        =   32
            Top             =   720
            Width           =   2355
         End
         Begin VB.Label lblDelete 
            Caption         =   "Select the Code Language you wish to delete:"
            Height          =   465
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   2085
         End
      End
      Begin VB.Frame frameModify 
         Caption         =   "Modify"
         Height          =   1785
         Left            =   -74880
         TabIndex        =   28
         Top             =   1560
         Width           =   3015
         Begin VB.ListBox lstModify 
            Height          =   840
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   2475
         End
         Begin VB.Label lblModify 
            Caption         =   "Select the Code Language you wish to modify:"
            Height          =   465
            Left            =   240
            TabIndex        =   30
            Top             =   240
            Width           =   2595
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Add New Language"
         Height          =   975
         Left            =   -74880
         TabIndex        =   25
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton Command3 
            Caption         =   "Add New"
            Height          =   375
            Left            =   1920
            TabIndex        =   27
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "DataBase Search"
         Height          =   1215
         Index           =   1
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton CmdsearchDB 
            Caption         =   "Search"
            Default         =   -1  'True
            Height          =   350
            Left            =   2160
            TabIndex        =   22
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   360
            TabIndex        =   21
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Enter Code Title:"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Internet Search"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   14
         Top             =   1800
         Width           =   3015
         Begin VB.ComboBox cmEngines 
            Height          =   315
            ItemData        =   "frmMain.frx":09A0
            Left            =   240
            List            =   "frmMain.frx":09C8
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Search"
            Height          =   350
            Left            =   2040
            TabIndex        =   16
            Top             =   1320
            Width           =   855
         End
         Begin VB.ComboBox txWhatSearch 
            Height          =   315
            ItemData        =   "frmMain.frx":0A39
            Left            =   240
            List            =   "frmMain.frx":0A5B
            TabIndex        =   15
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "Search Engine:"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lbladvance 
            Caption         =   "Search For:"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   2415
         End
      End
      Begin MSComctlLib.ListView lstTitles 
         Height          =   4815
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   3352
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Laguage"
            Object.Width           =   3176
         EndProperty
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   -75000
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   -74160
         TabIndex        =   68
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog cdoOpenDatabase 
      Left            =   9360
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "mdb"
      DialogTitle     =   "Open Database"
      Filter          =   "Access Database Files|*.mdb|All Files|*.*"
      FilterIndex     =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   9960
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AE7
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BF9
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D0B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E1D
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F2F
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1041
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1153
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1265
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1377
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1691
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2139
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2453
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3607
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EE1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabSnippit 
      Height          =   7095
      Left            =   3480
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   529
      TabMaxWidth     =   3528
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Code Wndow"
      TabPicture(0)   =   "frmMain.frx":47BB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "toolBar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "rtbCodeWindow"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "HTML Editor/Grabber"
      TabPicture(1)   =   "frmMain.frx":47D7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "rtbnet"
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(3)=   "PBar"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "SQL Builder"
      TabPicture(2)   =   "frmMain.frx":47F3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).ControlCount=   2
      Begin VB.PictureBox Picture2 
         Height          =   5175
         Left            =   -74880
         ScaleHeight     =   5115
         ScaleWidth      =   7515
         TabIndex        =   77
         Top             =   1800
         Visible         =   0   'False
         Width           =   7575
         Begin SHDocVwCtl.WebBrowser Web 
            Height          =   5160
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   7575
            ExtentX         =   13361
            ExtentY         =   9102
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
      End
      Begin SourceCodeNotebook.CodeHighlight rtbnet 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   76
         Top             =   1800
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   9128
         Language        =   3
         KeywordColor    =   16576
         OperatorColor   =   12582912
         DelimiterColor  =   32896
         ForeColor       =   0
         FunctionColor   =   12583104
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SourceCodeNotebook.CodeHighlight rtbCodeWindow 
         Height          =   5895
         Left            =   120
         TabIndex        =   75
         Top             =   1080
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   10398
         Language        =   1
         KeywordColor    =   12582912
         OperatorColor   =   12582912
         DelimiterColor  =   32768
         ForeColor       =   0
         FunctionColor   =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame9 
         Caption         =   "HTML Editor / Grabber"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   56
         Top             =   360
         Width           =   7575
         Begin SourceCodeNotebook.chameleonButton chameleonButton2 
            Height          =   255
            Left            =   3960
            TabIndex        =   74
            ToolTipText     =   "View HTML"
            Top             =   960
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "View HTML"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":480F
         End
         Begin SourceCodeNotebook.chameleonButton Command1 
            Height          =   255
            Left            =   3240
            TabIndex        =   73
            ToolTipText     =   "Preview in browser"
            Top             =   960
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Preview"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":482B
         End
         Begin SourceCodeNotebook.chameleonButton chameleonButton1 
            Height          =   255
            Left            =   6480
            TabIndex        =   70
            ToolTipText     =   "Full View Maximized Editor's Window"
            Top             =   960
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Maximized "
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":4847
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Copy"
            Height          =   375
            Left            =   6000
            TabIndex        =   60
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton CmdGet 
            Caption         =   "Get"
            Enabled         =   0   'False
            Height          =   375
            Left            =   5040
            TabIndex        =   59
            ToolTipText     =   "Go get grab it"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox TxtUrl 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   58
            Text            =   "http://"
            ToolTipText     =   "Enter the URL of the website you want to grab"
            Top             =   480
            Width           =   4815
         End
         Begin SourceCodeNotebook.chameleonButton cmdfile 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   57
            ToolTipText     =   "Open/Save/New"
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "File"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":4863
         End
         Begin SourceCodeNotebook.chameleonButton cmdtable 
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   61
            ToolTipText     =   "Table options"
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Table"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":487F
         End
         Begin SourceCodeNotebook.chameleonButton cmdfont 
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   62
            ToolTipText     =   "Insert Scrolling Marquee"
            Top             =   960
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Marquee"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":489B
         End
         Begin SourceCodeNotebook.chameleonButton cmdinsert 
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   63
            ToolTipText     =   "Image/Links/Tag/Background color/Character"
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Insert"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":48B7
         End
         Begin SourceCodeNotebook.chameleonButton cmdother 
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   64
            ToolTipText     =   "Line Break/White Paper etc.."
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Other"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":48D3
         End
         Begin SourceCodeNotebook.chameleonButton cmdprev 
            Height          =   255
            Index           =   0
            Left            =   5040
            TabIndex        =   65
            ToolTipText     =   "Preview In Full Size"
            Top             =   960
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            BTYPE           =   8
            TX              =   "Full Size Preview"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   13160660
            FCOL            =   0
            FCOLO           =   12582912
            MPTR            =   0
            MICON           =   "frmMain.frx":48EF
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "DataBase Scanner"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   49
         Top             =   360
         Width           =   7575
         Begin VB.CommandButton cmdstart 
            Caption         =   "&Scan"
            Height          =   255
            Left            =   1320
            TabIndex        =   54
            ToolTipText     =   "Click to Scan DataBase file"
            Top             =   2040
            Width           =   855
         End
         Begin VB.CheckBox checkscan 
            Caption         =   "Include Subdirectories"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            TabStop         =   0   'False
            ToolTipText     =   "Include Subdirectory"
            Top             =   2040
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.DirListBox Dir1 
            Height          =   1440
            Left            =   120
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   600
            Width           =   2055
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   240
            Width           =   2055
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1815
            Left            =   2280
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "In Folder"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Size"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   3840
            TabIndex        =   80
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label lblscan 
            ForeColor       =   &H00800000&
            Height          =   135
            Left            =   6960
            TabIndex        =   55
            Top             =   240
            Width           =   135
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "SQL Builder"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   35
         Top             =   2760
         Width           =   7575
         Begin VB.CommandButton Command5 
            Caption         =   "Database Viewer"
            Height          =   255
            Left            =   5640
            TabIndex        =   81
            ToolTipText     =   "Find and View Database"
            Top             =   3840
            Width           =   1695
         End
         Begin VB.CommandButton Clearz 
            Caption         =   "Clear"
            Height          =   255
            Left            =   3120
            TabIndex        =   79
            ToolTipText     =   "Insert tag"
            Top             =   3480
            Width           =   1095
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2700
            ItemData        =   "frmMain.frx":490B
            Left            =   5520
            List            =   "frmMain.frx":496C
            TabIndex        =   46
            Top             =   960
            Width           =   1815
         End
         Begin VB.ListBox List2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   645
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Table row"
            Top             =   3120
            Width           =   2895
         End
         Begin VB.CommandButton cmdsq1 
            Caption         =   "Insert SQL"
            Height          =   255
            Left            =   4320
            TabIndex        =   43
            Top             =   3120
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Insert as Auctioneer Tag"
            Height          =   195
            Left            =   2040
            TabIndex        =   42
            Top             =   3840
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Insert as VTrader Tag"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.ListBox List3 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7440
            TabIndex        =   40
            ToolTipText     =   "Coloumn row"
            Top             =   3120
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.CommandButton cmdsq2 
            Caption         =   "Insert BDTB"
            Height          =   255
            Left            =   4320
            TabIndex        =   39
            ToolTipText     =   "Find Insert Database table"
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Insert Tag"
            Height          =   255
            Left            =   3120
            TabIndex        =   38
            ToolTipText     =   "Insert tag"
            Top             =   3120
            Width           =   1095
         End
         Begin VB.CheckBox chksq1 
            Caption         =   "Insert Comma"
            Height          =   255
            Left            =   4200
            TabIndex        =   37
            Top             =   3840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMain.frx":4A33
            Left            =   5520
            List            =   "frmMain.frx":4A40
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   360
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   2895
            Left            =   120
            TabIndex        =   44
            ToolTipText     =   "SQL Builder window"
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   5106
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   3
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"frmMain.frx":4A5A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbllsq3 
            AutoSize        =   -1  'True
            Caption         =   "Scope:"
            Height          =   195
            Left            =   5640
            TabIndex        =   48
            Top             =   120
            Width           =   510
         End
         Begin VB.Label lblsq6 
            AutoSize        =   -1  'True
            Caption         =   "Sql:"
            Height          =   195
            Left            =   5640
            TabIndex        =   47
            Top             =   720
            Width           =   270
         End
      End
      Begin MSComctlLib.Toolbar toolBar 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "New  code snippit"
               Object.Tag             =   "new"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Edit Code "
               Object.Tag             =   "open"
               ImageIndex      =   17
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Save code snippit to the DataBase"
               Object.Tag             =   "save"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copy code snippit"
               Object.Tag             =   "copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "View/write notes source code info"
               Object.Tag             =   "paste"
               ImageIndex      =   18
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "view"
                     Text            =   "View Notes/Author Info.."
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "contents"
                     Text            =   "View Contents"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Delete code snippit"
               Object.Tag             =   "delete"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Search Code Data Base / Internet"
               Object.Tag             =   "find"
               ImageKey        =   "Find"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "DB"
                     Text            =   "Search DataBase"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "SDB"
                     Text            =   "Search Code in the internet"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Print code snippit"
               Object.Tag             =   "print"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modify / Add New / Delete Code"
               Object.Tag             =   "mod"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert Javascript / DHTML Code in FrontPage 2000"
               Object.Tag             =   "front"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert Javascript /sideserver Code in Dreamweaver4"
               Object.Tag             =   "drw"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Insert Code in Visual Basic6"
               Object.Tag             =   "VB"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "Add New Code Language"
         Height          =   1335
         Index           =   0
         Left            =   -73560
         TabIndex        =   2
         Top             =   480
         Width           =   4335
         Begin VB.Label Label3 
            Caption         =   "Enter A new Code Langugae:"
            Height          =   255
            Left            =   720
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   7575
      End
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   135
         Left            =   -74880
         TabIndex        =   72
         Top             =   1680
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3255
      Begin MSComDlg.CommonDialog CmDlg 
         Left            =   2160
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   360
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.ComboBox cmbFilter 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1395
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmMain.frx":4ADC
         Left            =   120
         List            =   "frmMain.frx":4ADE
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1440
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblFilter 
         Caption         =   "Filter By Language:"
         Height          =   225
         Left            =   1680
         TabIndex        =   10
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "Code Language:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Source Code Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   24
      Top             =   7215
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14367
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2548
            MinWidth        =   2548
            TextSave        =   "2/13/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            MinWidth        =   2548
            TextSave        =   "1:27 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnuNew 
         Caption         =   "New code "
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnucont 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnunet 
         Caption         =   "Search Code"
      End
      Begin VB.Menu me1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&Edit "
      Index           =   2
      Begin VB.Menu mnucopy 
         Caption         =   "Copy Code"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste Code"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnusettingz 
         Caption         =   "Undo"
      End
      Begin VB.Menu htmlsetprop 
         Caption         =   "Select All"
      End
      Begin VB.Menu me2 
         Caption         =   "-"
      End
      Begin VB.Menu mnudelete 
         Caption         =   "Delete current code"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save current code"
         Shortcut        =   ^S
      End
      Begin VB.Menu me3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnusetting 
      Caption         =   "Settings"
      Begin VB.Menu mnuusesounds 
         Caption         =   "# Play Typing Sound"
      End
   End
   Begin VB.Menu mnuType 
      Caption         =   "&Add - Language"
      Begin VB.Menu mnuModify 
         Caption         =   "Modify / Add / Delete"
      End
   End
   Begin VB.Menu tool 
      Caption         =   "&Tools"
      Begin VB.Menu mnuicon 
         Caption         =   "Icon Scanner"
      End
      Begin VB.Menu fileview 
         Caption         =   "File Viewer (htm,html,js,css)..."
      End
      Begin VB.Menu iconz 
         Caption         =   "-"
      End
      Begin VB.Menu sql 
         Caption         =   "SQL Builder..."
      End
      Begin VB.Menu html 
         Caption         =   "HTML Editor/Grabber"
      End
   End
   Begin VB.Menu notez 
      Caption         =   "Notes"
      Begin VB.Menu show1 
         Caption         =   "View Write Code Info notes"
      End
      Begin VB.Menu hideme 
         Caption         =   "Hide notes"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpcon 
         Caption         =   "&Contents?"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnucontact 
         Caption         =   "E-mail"
      End
      Begin VB.Menu me5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWebsite 
         Caption         =   "Website"
      End
      Begin VB.Menu mnuPlanetSourceCode 
         Caption         =   "Bug report"
      End
   End
   Begin VB.Menu MnuTables 
      Caption         =   "Tables"
      Visible         =   0   'False
      Begin VB.Menu MnuEnterAsStatement 
         Caption         =   "Enter as Statement "
      End
      Begin VB.Menu mnucol 
         Caption         =   "Columns"
      End
   End
   Begin VB.Menu mnusetz 
      Caption         =   "setting"
      Visible         =   0   'False
      Begin VB.Menu mnucopy2 
         Caption         =   "Copy..."
      End
      Begin VB.Menu mnupaste2 
         Caption         =   "Paste..."
      End
      Begin VB.Menu mnuprint2 
         Caption         =   "Undo"
      End
      Begin VB.Menu space9 
         Caption         =   "Select All"
      End
      Begin VB.Menu spacer007 
         Caption         =   "-"
      End
      Begin VB.Menu mnusetupme 
         Caption         =   "View Code Contents..."
      End
      Begin VB.Menu delete 
         Caption         =   "View Code Notes/Author Info.."
      End
      Begin VB.Menu mnuspace 
         Caption         =   "-"
      End
      Begin VB.Menu editme 
         Caption         =   "Edit with Organizer Pad.."
      End
   End
   Begin VB.Menu menuz 
      Caption         =   "setupme"
      Visible         =   0   'False
      Begin VB.Menu cut 
         Caption         =   "Cut..."
      End
      Begin VB.Menu copy1 
         Caption         =   "Copy.."
      End
      Begin VB.Menu dexter 
         Caption         =   "-"
      End
      Begin VB.Menu paste2 
         Caption         =   "Paste...."
      End
      Begin VB.Menu setup1 
         Caption         =   "Undo"
      End
      Begin VB.Menu print2 
         Caption         =   "Select All"
      End
      Begin VB.Menu char1 
         Caption         =   "-"
      End
      Begin VB.Menu char 
         Caption         =   "Insert Character #@~"
      End
      Begin VB.Menu mnuinsertagz 
         Caption         =   "Insert Tag"
      End
      Begin VB.Menu property 
         Caption         =   "Tag  Properties"
      End
      Begin VB.Menu dex4 
         Caption         =   "-"
      End
      Begin VB.Menu browse 
         Caption         =   "Preview in Browser"
      End
   End
   Begin VB.Menu file 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu mnu1 
         Caption         =   "New"
      End
      Begin VB.Menu mnu2 
         Caption         =   "Open"
      End
      Begin VB.Menu mnu3 
         Caption         =   "Save"
      End
      Begin VB.Menu sound 
         Caption         =   "# Play Typing Sound"
      End
   End
   Begin VB.Menu table 
      Caption         =   "Tables"
      Visible         =   0   'False
      Begin VB.Menu table1 
         Caption         =   "Add the first column"
         Begin VB.Menu back1 
            Caption         =   "Background"
            Begin VB.Menu back2 
               Caption         =   "Black"
            End
            Begin VB.Menu back3 
               Caption         =   "Blue"
            End
            Begin VB.Menu backspace 
               Caption         =   "-"
            End
            Begin VB.Menu back4 
               Caption         =   "Blue Violet"
            End
            Begin VB.Menu back5 
               Caption         =   "Brown"
            End
            Begin VB.Menu back6 
               Caption         =   "Cyan"
            End
            Begin VB.Menu backspace1 
               Caption         =   "-"
            End
            Begin VB.Menu back7 
               Caption         =   "Dark Brown"
            End
            Begin VB.Menu back8 
               Caption         =   "Dark Green"
            End
            Begin VB.Menu back9 
               Caption         =   "Dark Blue"
            End
            Begin VB.Menu backspace4 
               Caption         =   "-"
            End
            Begin VB.Menu back10 
               Caption         =   "Gold"
            End
         End
      End
      Begin VB.Menu addcol1 
         Caption         =   "Add new column"
         Begin VB.Menu ground 
            Caption         =   "Background"
            Begin VB.Menu add1 
               Caption         =   "Black"
            End
            Begin VB.Menu add32 
               Caption         =   "Blue"
            End
            Begin VB.Menu add2 
               Caption         =   "Black Violet"
            End
            Begin VB.Menu add3 
               Caption         =   "Blue Violet"
            End
            Begin VB.Menu add4 
               Caption         =   "Brown"
            End
            Begin VB.Menu add5 
               Caption         =   "Cyan"
            End
            Begin VB.Menu add6 
               Caption         =   "Dark Brown"
            End
            Begin VB.Menu add7 
               Caption         =   "Dark green"
            End
            Begin VB.Menu add8 
               Caption         =   "Dark Blue"
            End
            Begin VB.Menu add9 
               Caption         =   "Gold"
            End
         End
      End
      Begin VB.Menu cell 
         Caption         =   "Add Cells"
      End
      Begin VB.Menu cellspace 
         Caption         =   "-"
      End
      Begin VB.Menu addcolumn1 
         Caption         =   "Add more columns"
         Begin VB.Menu col1 
            Caption         =   "Add One Column"
         End
         Begin VB.Menu col2 
            Caption         =   "Add two Column"
         End
         Begin VB.Menu col3 
            Caption         =   "Add three column"
         End
         Begin VB.Menu col4 
            Caption         =   "Add four column"
         End
         Begin VB.Menu mr1 
            Caption         =   "Add more columns"
         End
         Begin VB.Menu addrow 
            Caption         =   "-"
         End
         Begin VB.Menu addrow1 
            Caption         =   "Add Rows"
         End
      End
   End
   Begin VB.Menu fontme 
      Caption         =   "Marquee"
      Visible         =   0   'False
      Begin VB.Menu fontz 
         Caption         =   "Scrolling Marquee"
         Begin VB.Menu mnumarqsrcolleft 
            Caption         =   "Scroll Left"
         End
         Begin VB.Menu mnumarqscrollright 
            Caption         =   "Scroll Right"
         End
         Begin VB.Menu mnumarqscrollup 
            Caption         =   "Scroll Up"
         End
         Begin VB.Menu mnumarqscrolldown 
            Caption         =   "Scroll Down"
         End
      End
      Begin VB.Menu fn8 
         Caption         =   "Alternate Marquee"
         Begin VB.Menu mnumarqalternateright 
            Caption         =   "Alternate Right"
         End
         Begin VB.Menu mnumarqalternateleft 
            Caption         =   "Alternate Left"
         End
      End
      Begin VB.Menu mnumarquee2 
         Caption         =   "Slide Marquee"
         Begin VB.Menu mnumarqslidelft 
            Caption         =   "Slide Left"
         End
         Begin VB.Menu mnumarqslideright 
            Caption         =   "Slide Right"
         End
      End
   End
   Begin VB.Menu other 
      Caption         =   "Other"
      Visible         =   0   'False
      Begin VB.Menu ot1 
         Caption         =   "Unmumbered List"
      End
      Begin VB.Menu ot2 
         Caption         =   "Numbered List"
      End
      Begin VB.Menu ot3 
         Caption         =   "Definition List"
      End
      Begin VB.Menu ot4 
         Caption         =   "Nested List"
      End
      Begin VB.Menu ot5 
         Caption         =   "Exteneded Quotation"
      End
      Begin VB.Menu ot6 
         Caption         =   "Force Line Breaks"
      End
      Begin VB.Menu ot7 
         Caption         =   "Horizontal Rules"
      End
      Begin VB.Menu ot8 
         Caption         =   "White Space"
      End
   End
   Begin VB.Menu insert 
      Caption         =   "Insert"
      Visible         =   0   'False
      Begin VB.Menu mnucolorpicker 
         Caption         =   "Color Picker"
      End
      Begin VB.Menu time 
         Caption         =   "Time - Date"
      End
      Begin VB.Menu mnutagz 
         Caption         =   "Insert tag"
      End
      Begin VB.Menu mnuchardex 
         Caption         =   "Insert Character"
      End
      Begin VB.Menu mnubackgroundcol 
         Caption         =   "Tag Properties"
      End
   End
   Begin VB.Menu mnufilelist 
      Caption         =   "filez"
      Visible         =   0   'False
      Begin VB.Menu mnudelete01 
         Caption         =   "Delete Code"
      End
      Begin VB.Menu mnuaddmod01 
         Caption         =   "Add/Modify"
      End
      Begin VB.Menu mnusearchcodez 
         Caption         =   "Search Code"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const adSchemaColumns = 4
Dim SName As String
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String
Public Path As String
Public saved As Boolean

Private Sub add1_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#000000>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add2_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add3_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add32_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add4_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add5_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add6_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add7_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub add9_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on " + Chr(13) + Chr(10) + "</TD>" + "<TD>"
End Sub

Private Sub addrow1_Click()
rtbnet.SelText = rtbnet.SelText + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor= >" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the added ROW)" + Chr(13) + Chr(10) + "</TD><TD bgcolor= >" + Chr(13) + Chr(10) + "ADD HERE CELLS. Select and Paste the two lines above: <P>...</TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (Last cell of the added ROW)" + Chr(13) + Chr(10) + "</TD></TR>ADD HERE ROWS"
End Sub

Private Sub back10_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back11_Click()

End Sub

Private Sub back2_Click()
rtbnet.SelText = rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#CD7F32>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + "Here add new cells for the first column" + Chr(13) + Chr(10) + "Here add the second column" + Chr(13) + Chr(10) + "Here add new cells for the second column, and so on " + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back3_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#0000FF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back4_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#9F5F9F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back5_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#A62A2A>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back6_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#00FFFF>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back7_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#5C4033>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back8_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#2F4F2F>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub back9_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + "<P>" + "<TABLE BORDER=Enter a number from 0 and up>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=#871F78>" + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10) + Chr(13) + Chr(10) + "</TABLE>" + "</P>"
End Sub

Private Sub browse_Click()
Picture2.Visible = True
CmdGet.Enabled = False
TxtUrl.Locked = True
Open App.Path & "\Organizerpreview\extreme.html" For Output As #1
Print #1, rtbnet.Text
Close #1
Load Browser
Textweb.Text = App.Path & "\Organizerpreview\extreme.html"
Web.Navigate App.Path & "\Organizerpreview\extreme.html"
End Sub

Private Sub cell_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<P>" + "Enter text, image, and so on" + Chr(13) + Chr(10) + "</TD>" + "<TD>" + Chr(13) + Chr(10)
End Sub

Private Sub chameleonButton1_Click()
If rtbnet.Text <> "" Then
HTMLEditor.Show
HTMLEditor.RichTextBox1 = rtbnet.Text
End If
End Sub

Private Sub chameleonButton2_Click()
Picture2.Visible = False
CmdGet.Enabled = True
TxtUrl.Locked = False
Me.Caption = "Source Code Organizer V.1"
sBar.Panels(1).Text = "Status: Viewing"
End Sub

Private Sub char_Click()
Symbols.Show
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Javascript.com )..."
End Sub

Private Sub Check2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Planet-source-code.com )..."
End Sub

Private Sub Check3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Simplythebest.net )..."
End Sub

Private Sub Check4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Codeguru.com )..."
End Sub

Private Sub Check5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( Planet-source-code.com )..."
End Sub

Private Sub Check6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Instant internet Code search.By checking the box it will take you to..( vyaskn.tripod.com )..."
End Sub
Sub FilesSearch(DrivePath As String, Ext As String)

Dim XDir() As String

Dim TmpDir As String

Dim FFound As String

Dim DirCount As Integer

Dim x As Integer

Dim li As ListItem

DirCount = 0

ReDim XDir(0) As String

XDir(DirCount) = ""

If Right(DrivePath, 1) <> "\" Then

DrivePath = DrivePath & "\"

End If

'Enter here the code for showing the pat
' h being
'search. Example: Form1.label2 = DrivePa
' th
'Search for all directories and store in
' the
'XDir() variable
sBar.Panels(1).Text = DrivePath

DoEvents

TmpDir = Dir(DrivePath, vbDirectory)


Do While TmpDir <> ""


If TmpDir <> "." And TmpDir <> ".." Then


If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then

XDir(DirCount) = DrivePath & TmpDir & "\"

DirCount = DirCount + 1
ReDim Preserve XDir(DirCount) As String
End If
End If
TmpDir = Dir
Loop
'Searches for the files given by extensi
' on Ext
FFound = Dir(DrivePath & Ext)
Do Until FFound = ""
Set li = ListView1.ListItems.Add(, , FFound)
li.ListSubItems.Add , , DrivePath
li.ListSubItems.Add , , FileLen(DrivePath & FFound) & " Bytes"
FFound = Dir
Loop
If checkscan.Value = 1 Then
For x = 0 To (UBound(XDir) - 1)
FilesSearch XDir(x), Ext
Next x
Else
End If
End Sub
Private Sub checkscan_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Check this box to scan sub directory..."
End Sub
Private Sub chksq1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert comma..."
End Sub

Private Sub Clearz_Click()
RichTextBox1.Text = ""
End Sub

Private Sub cmdfile_Click(Index As Integer)
PopupMenu file
End Sub

Private Sub cmdfont_Click(Index As Integer)
PopupMenu fontme
End Sub

Private Sub CmdGet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Get Website HTML..."
End Sub

Private Sub cmdInsert_Click(Index As Integer)
PopupMenu insert
End Sub

Private Sub cmdother_Click(Index As Integer)
PopupMenu other
End Sub

Private Sub cmdprev_Click(Index As Integer)
Open App.Path & "\Organizerpreview\extreme.html" For Output As #1
Print #1, rtbnet.Text
Close #1
Load Browser
Browser.Text1.Text = App.Path & "\Organizerpreview\extreme.html"
Browser.Caption = "You are Browsing:"
Browser.Show
Browser.Web.Navigate App.Path & "\Organizerpreview\extreme.html"
End Sub

Private Sub cmdsq1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert SQL character..."
End Sub

Private Sub cmdsq2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Find Database..."
End Sub
Private Sub cmdStart_Click()
cmdStart.Enabled = False
ListView1.ListItems.Clear
Screen.MousePointer = vbHourglass
FilesSearch Dir1.Path, "*.mdb"
Screen.MousePointer = vbDefault
sBar.Panels(1).Text = ListView1.ListItems.Count & " Database's found!"
Label6.Caption = ListView1.ListItems.Count & " Database's found!"
cmdStart.Enabled = True
End Sub

Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Click to scan database..."
End Sub

Private Sub cmdtable_Click(Index As Integer)
PopupMenu table
End Sub

Private Sub col1_Click()
rtbnet.SelText = rtbnet.SelText + Chr(13) + Chr(10) + "<P><TABLE BORDER=1>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the first cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Write your text in the second cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "Copy and Paste the following code to add more cells" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the code of the cell's background>" + Chr(13) + Chr(10) + "<P>Your your text in the cell" + Chr(13) + Chr(10) + "</TD></TR>" + Chr(13) + Chr(10) + "xxxxxxxxxxxxxxxxxxxxxxxxx" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub col2_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub col3_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub col4_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub
Private Sub Command1_Click()
sBar.Panels(1).Text = "Status: Viewing"
Picture2.Visible = True
CmdGet.Enabled = False
TxtUrl.Locked = True
Open App.Path & "\Organizerpreview\extreme.html" For Output As #1
Print #1, rtbnet.Text
Close #1
Load Browser
Textweb.Text = App.Path & "\Organizerpreview\extreme.html"
Web.Navigate App.Path & "\Organizerpreview\extreme.html"
End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText rtbnet.Text
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Viewing..."
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Copy HTML code..."
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert Tag..."
End Sub
Private Sub Command5_Click()
frmfinddb.Show
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Clear field..."
End Sub
Private Sub copy1_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText rtbCodeWindow.Text
Clipboard.SetText rtbnet.SelText
End Sub

Private Sub cut_Click()
 Clipboard.SetText rtbnet.SelText
 rtbnet.SelText = ""
End Sub

Private Sub delete_Click()
SSTab1.Tab = 1
Picture1.Visible = True
End Sub

Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Select a folder you want to scan..."
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub


Private Sub CmdGet_Click()
saved = False
If IsNetConnectOnline = False Then
MsgBox "You are not connected to the internet!", vbOKOnly, App.ProductName
Unload Me
Else
If TxtUrl.Text = "http://" Then
MsgBox ("This is a malformed url!" & vbCrLf & "Operation has been canceled"), , App.ProductName
Else
Screen.MousePointer = vbHourglass
rtbnet.Text = Inet1.OpenURL(TxtUrl.Text)
Me.Caption = "Source Code Organizer V.1" & "-HTML Code for-" & " [ " & TxtUrl.Text & " ]"
Screen.MousePointer = vbDefault
CmdGet.Enabled = False
End If
End If
End Sub



Private Sub editme_Click()
If rtbCodeWindow.Text <> "" Then
note.Show
note.Rich = rtbCodeWindow.Text
mnuSave.Enabled = False
mnudelete.Enabled = False
mnunew.Enabled = False
mnupaste.Enabled = False
mnuModify.Enabled = False
mnusettingz.Enabled = False
mnuOpen.Enabled = False
End If
End Sub

Private Sub fileview_Click()
frmfind.Show
End Sub

Private Sub Form_Resize()
ResizeForm frmmain
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Code Language Selection..."
End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub Frame8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub hideme_Click()
Picture1.Visible = False
SSTab1.Tab = 0
End Sub

Private Sub html_Click()
tabSnippit.Tab = 1
End Sub

Private Sub htmlsetprop_Click()
 rtbCodeWindow.SelStart = 0
 rtbCodeWindow.SelLength = Len(rtbCodeWindow.Text)
 rtbCodeWindow.SetFocus
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Click an item to insert..."
End Sub

Private Sub List2_Click()
PopupMenu mnuTables
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Scanned Data Base Files..."
End Sub

Private Sub lstTitles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
      PopupMenu mnufilelist
    End If
End Sub

Private Sub mnu1_Click()
If MsgBox("Do you want to save your current project?", vbYesNo, "Save") = vbYes Then
        mnu3_Click
       rtbnet.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//Source Code Organizer HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    rtbnet.Text = rtbnet.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>My Web Page</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
    Else
        rtbnet.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//Source Code Organizer HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    rtbnet.Text = rtbnet.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>My Web Page</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
    End If
End Sub

Private Sub mnu2_Click()
Dim sFile As String
    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "HTML Files (*.html)|*.html|All Files (*.*)|*.*|"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 intFileNum = FreeFile
 Open App.Path & "\recent.dat" For Append As #intFileNum
 Print #intFileNum, sFile
 Close #intFileNum
    
    End With
    rtbnet.LoadFile sFile
End Sub

Private Sub mnu3_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm)|RTF Files (*.rtf)|*.rtf|Text files (*.txt)|*.txt|Ini Files (*.ini)|*.ini|Registry Files (*.log)|*.log|Batch File (*.bat)|*.bat|All files (*.*)|*.*"""
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, rtbnet.Text
    Close #1
End If
End Sub

Private Sub mnuaddmod01_Click()
SSTab1.Tab = 2
End Sub

Private Sub mnubackgroundcol_Click()
frmTagEdit.Show
'frmTool_Rainbow.Show
End Sub

Private Sub mnuchardex_Click()
Symbols.Show
End Sub

Private Sub mnucol_Click()
MsgBox "Not implemted yet..", vbInformation, "Information"
End Sub

Private Sub mnucolorpicker_Click()
frmColor.Show
End Sub

Private Sub mnucont_Click()
SSTab1.Tab = 0
End Sub

Private Sub mnucopy2_Click()
Clipboard.Clear
    Clipboard.SetText rtbCodeWindow.Text
End Sub

Private Sub mnudelete01_Click()
mnuDelete_Click
End Sub

Private Sub mnuhelpcon_Click()
On Error GoTo HandleErrors
 Call Shell("HH.exe help.chm", vbNormalFocus)

 Exit Sub
HandleErrors:
  Dim intresponse As Integer
   Select Case Err.Number
        Case 53, 76
         intresponse = MsgBox("File not found.", vbCritical, "Error")
         End Select
End Sub

Private Sub mnuicon_Click()
frmscan.Show
End Sub

Private Sub mnuinsertagz_Click()
frmTags.Show , frmmain
End Sub

Private Sub mnumarqalternateleft_Click()
rtbnet.SelText = "<Marquee bgcolor= #FFFFFF  Behavior = Alternate  Direction =  Left >" + rtbnet.SelText + "Enter Your Text Here" + "</Marquee>"
End Sub

Private Sub mnumarqalternateright_Click()
rtbnet.SelText = "<Marquee bgcolor= #FFFFFF  Behavior = Alternate  Direction =  Right >" + rtbnet.SelText + "Enter Your Text Here" + "</Marquee>"
End Sub

Private Sub mnumarqscrolldown_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = Down > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnumarqscrollright_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = right > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnumarqscrollup_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = UP > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnumarqslidelft_Click()
RichTextBox1.SelText = "<marquee bgcolor= #FFFFFF & Behavior = Slide & Direction = Left > " + RichTextBox1.SelText + "Hello World" + "</marquee >"
End Sub

Private Sub mnumarqslideright_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Behavior = Slide & Direction = Right > " + rtbnet.SelText + "Hello World" + "</marquee >"
End Sub

Private Sub mnumarqsrcolleft_Click()
rtbnet.SelText = "<marquee bgcolor= #FFFFFF & Direction = Left > " + rtbnet.SelText + "Enter Your Text Here" + "</marquee >"
End Sub

Private Sub mnuModify_Click()
  SSTab1.Tab = 2
End Sub

Private Sub mnupaste2_Click()
rtbCodeWindow.SelText = Clipboard.GetText
rtbCodeWindow.SetFocus
End Sub

Private Sub mnuprint2_Click()
mnusettingz_Click
End Sub

Private Sub mnuset_Click()

 CmDlg.Flags = cdlPDPrintSetup
            CmDlg.ShowPrinter
            DoEvents
End Sub

Private Sub mnusearchcodez_Click()
  SSTab1.Tab = 1
End Sub

Private Sub mnusettingz_Click()
   If gintIndex = 0 Then Exit Sub
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtbCodeWindow.Text = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub
Private Sub mnusetupme_Click()
SSTab1.Tab = 0
End Sub

Private Sub mnutagz_Click()
frmTags.Show
End Sub

Private Sub mnuusesounds_Click()
If mnuusesounds.Checked = False Then
        mnuusesounds.Checked = True
        UseSound = "Yes"
    ElseIf mnuusesounds.Checked = True Then
        mnuusesounds.Checked = False
        UseSound = ""
    End If
End Sub
Private Sub mr1_Click()
rtbnet.SelText = rtbnet.SelText + "<P><TABLE BORDER=1 bgcolor=Write the color-code for all cells>" + Chr(13) + Chr(10) + "<TR>" + Chr(13) + Chr(10) + "<TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (1st cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (2nd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "<P>Write your Text here (3rd cell of the FIRST COLUMN)" + Chr(13) + Chr(10) + "</TD><TD bgcolor=Write the backgroung color of cell or leave it blank>" + Chr(13) + Chr(10) + "ADD HERE COLUMNS. Select and Paste one of the two lines <P></TD><TD>" + Chr(13) + Chr(10) + "<P>Write your Text here (The cell of the LAST COLUMN)" + Chr(13) + Chr(10) + "</TD></TR>ADD ROWS HERE" + Chr(13) + Chr(10) + "</TABLE></P>"
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert Auctioner tag..."
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Insert trader  tag..."
End Sub

Private Sub ot1_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>"
End Sub

Private Sub ot2_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<OL>" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here" + Chr(13) + Chr(10) + "<LI> Type your text here and add more <LI> if necessary" + Chr(13) + Chr(10) + "</OL>"
End Sub

Private Sub ot3_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<DL>" + Chr(13) + Chr(10) + "<DT> Paragraph Title" + Chr(13) + Chr(10) + "<DD> Your Text Here" + Chr(13) + Chr(10) + "<DT> Second Paragraph Title" + Chr(13) + Chr(10) + "<DD> Your Text Here" + Chr(13) + Chr(10) + "</DL>"
End Sub

Private Sub ot4_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Sub-heading" + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here. Add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>" + Chr(13) + Chr(10) + "<LI> Second Sub-heading" + Chr(13) + Chr(10) + "<UL>" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here" + Chr(13) + Chr(10) + "<LI> Your Text Here. Add more <LI> if necessary" + Chr(13) + Chr(10) + "</UL>" + Chr(13) + Chr(10) + "</UL>"
End Sub

Private Sub ot5_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<P>Your text" + Chr(13) + Chr(10) + "<BLOCKQUOTE>" + Chr(13) + Chr(10) + "<P> Write your text here to include lengthy quotations in a separate block on the screen" + Chr(13) + Chr(10) + "</P>" + Chr(13) + Chr(10) + "<P> Add more text here if you want</P>" + Chr(13) + Chr(10) + "</BLOCKQUOTE>"
End Sub

Private Sub ot6_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<BR>"
End Sub

Private Sub ot7_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<HR SIZE= Enter the desired size    WIDTH=" + "Enter a number %>"
End Sub

Private Sub ot8_Click()
rtbnet.SelText = Chr(13) + Chr(10) + rtbnet.SelText + Chr(13) + Chr(10) + "<P>&nbsp;</P>"
End Sub

Private Sub paste2_Click()
 frmmain.rtbnet.SelText = Clipboard.GetText()
End Sub

Private Sub print2_Click()
 rtbnet.SelStart = 0
 rtbnet.SelLength = Len(rtbnet.Text)
 rtbnet.SetFocus
End Sub

Private Sub property_Click()
frmTagEdit.Show
End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "This is the window for SQL coding..."
End Sub

Private Sub rtbCodeWindow_Change()
If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = rtbCodeWindow.Text
    End If
If UseSound = "Yes" Then
        Dim Play As String
        Play = sndPlaySound(App.Path + "\Organizerpreview\Type.wav", SND_ASYNC)
    End If
End Sub

Private Sub rtbCodeWindow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then ' used for popup menu for the list of ftp sites saved for reuse utilizing the writeprivateprofile string
If rtbCodeWindow.Text = "" Then
mnucopy2.Enabled = False
mnusetupme.Enabled = False
mnupaste2.Enabled = True
mnuprint2.Enabled = True
editme.Enabled = False
delete.Enabled = False
space9.Enabled = False
PopupMenu mnusetz
Else
mnucopy2.Enabled = True
mnuprint2.Enabled = True
setup1.Enabled = True
mnupaste2.Enabled = True
editme.Enabled = True
delete.Enabled = True
space9.Enabled = True
PopupMenu mnusetz
End If
End If
End Sub

Private Sub rtbCodeWindow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Source Code..."
End Sub

Private Sub rtbnet_Change()
If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = rtbnet.Text
    End If
If UseSound = "Yes" Then
        Dim Play As String
        Play = sndPlaySound(App.Path + "\Organizerpreview\Type.wav", SND_ASYNC)
    End If
End Sub

Private Sub rtbnet_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then ' used for popup menu for the list of ftp sites saved for reuse utilizing the writeprivateprofile string
If rtbnet.Text = "" Then
copy1.Enabled = False
print2.Enabled = False
setup1.Enabled = False
paste2.Enabled = True
cut.Enabled = False
char.Enabled = False
browse.Enabled = False
mnuinsertagz.Enabled = False
property.Enabled = False
setup1.Enabled = True
PopupMenu menuz
Else
char.Enabled = True
browse.Enabled = True
mnuinsertagz.Enabled = True
property.Enabled = True
copy1.Enabled = True
cut.Enabled = True
print2.Enabled = True
setup1.Enabled = True
paste2.Enabled = True
PopupMenu menuz
End If
End If
End Sub

Private Sub rtbnet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then ' used for popup menu for the list of ftp sites saved for reuse utilizing the writeprivateprofile string
If rtbnet.Text = "" Then
copy1.Enabled = False
print2.Enabled = False
setup1.Enabled = False
paste2.Enabled = True
cut.Enabled = False
PopupMenu menuz
Else
copy1.Enabled = True
cut.Enabled = True
print2.Enabled = True
setup1.Enabled = True
paste2.Enabled = True
PopupMenu menuz

End If
End If

sBar.Panels(1).Text = "This is the window that will display the HTML code..."
End Sub
Private Sub rtbNotes_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Source code notes/author information..."
End Sub

Private Sub setup1_Click()
If gintIndex = 0 Then Exit Sub
    'This is the basic undo stuff.
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    rtbnet.Text = gstrStack(gintIndex)
    gblnIgnoreChange = False
End Sub

Private Sub show1_Click()
SSTab1.Tab = 1
Picture1.Visible = True
End Sub

Private Sub sound_Click()
 If sound.Checked = False Then
       sound.Checked = True
       UseSound = "Yes"
    ElseIf sound.Checked = True Then
        sound.Checked = False
        UseSound = ""
   End If
End Sub
Private Sub space9_Click()
htmlsetprop_Click
End Sub

Private Sub sql_Click()
tabSnippit.Tab = 2
mnucopy.Enabled = False
mnupaste.Enabled = False
mnudelete.Enabled = False
mnuSave.Enabled = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 If SSTab1.Tab = 0 Then
 Picture1.Visible = False
 Me.Caption = "Source Code Organizer V.1"
 End If
 If SSTab1.Tab = 2 Then
 Picture1.Visible = False
 Me.Caption = "Source Code Organizer V.1"
 End If
End Sub
Private Sub tabSnippit_Click(PreviousTab As Integer)
If tabSnippit.Tab = 0 Then
 Picture1.Visible = False
 mnusettingz.Enabled = True
 htmlsetprop.Enabled = True
 mnucopy.Enabled = True
 mnupaste.Enabled = True
 mnudelete.Enabled = True
 mnuSave.Enabled = True
 mnuFind.Enabled = True
 sql.Enabled = True
 mnunew.Enabled = True
 mnuOpen.Enabled = True
 mnuModify.Enabled = True
 hideme.Enabled = True
 mnuusesounds.Enabled = True
 show1.Enabled = True
 html.Enabled = True
 Me.Caption = "Source Code Organizer V.1"
 SSTab1.Tab = 0
 End If
 If tabSnippit.Tab = 1 Then
  Picture1.Visible = False
  mnusettingz.Enabled = False
  mnuFind.Enabled = False
  htmlsetprop.Enabled = True
  mnupaste.Enabled = True
  sql.Enabled = True
  mnudelete.Enabled = False
  mnuSave.Enabled = False
  mnunew.Enabled = False
  mnuOpen.Enabled = False
  mnuModify.Enabled = False
  hideme.Enabled = False
  show1.Enabled = False
  htmlsetprop.Enabled = False
  mnupaste.Enabled = False
  mnucopy.Enabled = False
  html.Enabled = False
  mnuusesounds.Enabled = False
  SSTab1.Tab = 0
  End If
  If tabSnippit.Tab = 2 Then
   Picture1.Visible = False
   mnunew.Enabled = False
   mnusettingz.Enabled = False
   htmlsetprop.Enabled = False
   mnuOpen.Enabled = False
   mnuModify.Enabled = False
   mnuFind.Enabled = False
   sql.Enabled = False
   hideme.Enabled = False
   show1.Enabled = False
   html.Enabled = True
   mnuusesounds.Enabled = False
   mnuSave.Enabled = False
   mnudelete.Enabled = False
   mnupaste.Enabled = False
   mnucopy.Enabled = False
   Me.Caption = "Source Code Organizer V.1"
   SSTab1.Tab = 0
  End If
End Sub

Private Sub tabSnippit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub time_Click()
DateTime.Show
End Sub

Private Sub txtTitle_Click()
SSTab1.Tab = 1
Picture1.Visible = True
End Sub

Private Sub TxtUrl_Change()
CmdGet.Enabled = True
rtbnet.Text = ""
If UseSound = "Yes" Then
        Dim Play As String
        Play = sndPlaySound(App.Path + "\Organizerpreview\Type.wav", SND_ASYNC)
    End If
End Sub

Private Sub TxtUrl_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
CmdGet_Click
End If
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If ListView1.SortKey = ColumnHeader.Index - 1 Then
If ListView1.SortOrder = lvwAscending Then
ListView1.SortOrder = lvwDescending
Else
ListView1.SortOrder = lvwAscending
End If
Else
ListView1.SortOrder = lvwAscending
ListView1.SortKey = ColumnHeader.Index - 1
End If
ListView1.Sorted = True
End Sub

Private Sub cmbFilter_DropDown()
sBar.Panels(1).Text = "Code Language Filter..."
End Sub

Private Sub cmbType_DropDown()
sBar.Panels(1).Text = "Code Language Selection..."
End Sub

Private Sub cmdSearch_Click()
  If txWhatSearch = "" Then Exit Sub
  Search cmEngines.ListIndex, txWhatSearch
  tabSnippit.Tab = 0
End Sub

Private Sub cmbFilter_Click()
    LoadGridBox
End Sub

Private Sub cmdDone_Click()
SSTab1.Tab = 0
tabSnippit.Tab = 0
End Sub

Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Search Code in the internet..."
End Sub

Private Sub CmdsearchDB_Click()
 tabSnippit.Tab = 0
 Dim RetVal As String
    RetVal = Text2.Text
    If RetVal = "" Then
        Exit Sub
    End If
    Find (RetVal)
End Sub

Private Sub CmdsearchDB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Search Code in the DataBase..."
End Sub

Private Sub cmEngines_DropDown()
sBar.Panels(1).Text = "Search Engine Selection..."
End Sub

Private Sub Command3_Click()
On Error GoTo errHandler
    
    Dim RetVal As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    RetVal = Text1.Text
    If RetVal = "" Then
     MsgBox "You must enter a new code language in the Text field", vbInformation, "Information"
       
        Exit Sub
    End If
    
    'connect to the database and add in the new codetype
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "INSERT INTO codetypes (codetype) VALUES ('" & StuffQuotes(RetVal) & "')"
    adoCmd.Execute
    MsgBox "Your new code language is now added to the DataBase.Click done at the bottom to go back to the source code and look for it in the dropdown box", vbInformation, "Added to the DataBase"
    'clean up
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    LoadCodeTypes
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Form_Load()

     If Textweb.Text <> "URL Address" Then Web.Navigate (Textweb.Text)
     rtbnet.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//Source Code Organizer HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    rtbnet.Text = rtbnet.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>My Web Page</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
      cmEngines.ListIndex = 0
  txWhatSearch.ListIndex = 3
    gblConnectString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & AppPath & "sourcebook.mdb"
    gblNewCode = True 'start off with a clean slate
    
    LoadCodeTypes
    rtbNotes.OLEDropMode = 1  'setup for the drag drop in the code windows
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset

    lstModify.Clear
    lstDelete.Clear
    'create the connection and execute the SQL
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "SELECT * FROM codetypes"
    Set adoRS = adoCmd.Execute
            
    Do While Not adoRS.EOF
        'add the recordsetset items into the listboxes
        lstModify.AddItem CStr(adoRS("codetype"))
        lstDelete.AddItem CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    
    'make sure we clean up!
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
        
End Sub
Private Sub cmdsq1_Click()
Dim quote As String
Dim DblQuote As String
quote = """"
DblQuote = quote & quote
If RichTextBox1.Text <> "" Then
If Option1.Value = True Then
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
RichTextBox1.SelText = "<#" & LCase(RichTextBox1.Text) & " #>"
Else
If Combo1.Text <> "" Then
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
RichTextBox1.SelText = "[!Query:" & Combo1.Text & " Name=" & DblQuote & " SQL =" & quote & LCase(RichTextBox1.Text) & quote & "!]"
Else
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = 1
RichTextBox1.SelText = ""
RichTextBox1.SelText = "[!Query:Open" & " Name=" & DblQuote & " SQL =" & quote & LCase(RichTextBox1.Text) & quote & "!]"
End If
End If
Else
Exit Sub
End If
End Sub

Private Sub cmdsq2_Click()
frmmain.List2.Clear
frmmain.List3.Clear
Form16.Show
End Sub

Private Sub Combo1_Click()
If Combo1.List(Combo1.ListIndex) = "Close" Then
If Option1.Value = True Then
frmmain.RichTextBox1.SelText = "<#Query:Close#>"
Unload Me
Else
frmmain.RichTextBox1.SelText = "[!Query:Close!]"
End If
Else
' Do Nothing as the user wants to build a custom sql statement
End If
End Sub


Private Sub Textweb_GotFocus()
    Textweb.SelStart = 0
    Textweb.SelLength = Len(Text1)
End Sub

Private Sub Textweb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Web.Navigate Me.Textweb.Text
    End If
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Textweb.Text = URL
End Sub

Private Sub Web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error GoTo progressERR
    If Progress = -1 Then PBar.Value = 100

    If Progress > 0 And ProgressMax > 0 Then
        PBar.Value = Progress * 100 / ProgressMax
       
    End If
    PBar.Visible = False
    Exit Sub
progressERR:
End Sub

Private Sub Web_StatusTextChange(ByVal Text As String)
    sBar.Panels(1).Text = Text
End Sub

Private Sub Command4_Click()
frmsqltag.Show
End Sub

Private Sub List1_Click()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List1.List(List1.ListIndex))
End Sub

Private Sub List1_DblClick()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List1.List(List1.ListIndex))
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
List1_DblClick
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
If List2.ListCount > 0 Then
PopupMenu mnuTables
End If
End If
End Sub

Private Sub lstTitles_ItemClick(ByVal Item As MSComctlLib.ListItem)
tabSnippit.Tab = 0
End Sub

Private Sub MnuEnterAsStatement_Click()
RichTextBox1.SelText = RichTextBox1.SelText & " " & UCase(List2.List(List2.ListIndex))
End Sub

Private Sub MnuGetColumns_Click()

End Sub



Private Sub LoadGridBox()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim cmdtext As String
    Dim lstItem As ListItem

    lstTitles.ListItems.Clear
    ' here we are building the SQL statement based upon the filter drop down
    If cmbFilter.Text = "No Filter" Then
        cmdtext = "SELECT id, title, codetype FROM source "
    Else
        cmdtext = "SELECT id, title, codetype FROM source WHERE codetype='" & StuffQuotes(cmbFilter.Text) & "' "
    End If
    
    'connect to the database and retrieve the code
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCon.CursorLocation = adUseClientBatch
    adoCmd.CommandText = cmdtext
    Set adoRS = adoCmd.Execute
    
    'loop through the recordset and add each item to the listview
    Do While Not adoRS.EOF
        Set lstItem = lstTitles.ListItems.Add(, , adoRS("title"))
        lstItem.Tag = adoRS("id") 'used for updating and deleting
        lstItem.SubItems(1) = CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    lstTitles_Click 'reset the list
    'make sure to clean up after ourselves
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub

Private Sub lstDelete_Click()
    
    On Error GoTo errHandler
    
    Dim RetVal As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    'only 1 chance to say no!
    RetVal = MsgBox("Are you sure you want to delete this Language?  This will change all snippits that have this code language, to a < 0 > code language", vbYesNo, "Delete Code Language")
    If RetVal = vbNo Then
        Exit Sub
    End If
    
    'connect to the database and delete the code type, the reset the source entries
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "DELETE FROM codetypes WHERE codetype='" & StuffQuotes(lstDelete.Text) & "'"
    adoCmd.Execute
    adoCmd.CommandText = "UPDATE source SET codetype='<blank>' WHERE codetype='" & StuffQuotes(lstDelete.Text) & "'"
    adoCmd.Execute
    
    'make sure we clean up after ourselves
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Form_Load  'reset everything
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub lstModify_Click()

    On Error GoTo errHandler
    
    Dim RetVal As String
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    RetVal = InputBox("Please enter in the new title for the Code langauge", "Modify Code Language", CStr(lstModify.Text))
    If RetVal = "" Then
        Exit Sub
    End If
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "UPDATE codetypes SET codetype='" & RetVal & "' WHERE codetype='" & StuffQuotes(lstModify.Text) & "'"
    adoCmd.Execute
    adoCmd.CommandText = "UPDATE source SET codetype='" & RetVal & "' WHERE codetype='" & StuffQuotes(lstModify.Text) & "'"
    adoCmd.Execute
    
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Form_Load
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub LoadCodeTypes()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    
    cmbType.Clear
    cmbFilter.Clear
    cmbFilter.AddItem "No Filter", 0 'no filter isnt in the db, so add it here so its on top
    
    'connect to the database and retrieve the valid code types
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "SELECT codetype FROM codetypes"
    Set adoRS = adoCmd.Execute
    
    'loop through the recordset and add them to the drop down
    Do While Not adoRS.EOF
        cmbType.AddItem CStr(adoRS("codetype"))
        cmbFilter.AddItem CStr(adoRS("codetype"))
        adoRS.MoveNext
    Loop
    
    'cleaning up the house
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing

    'reset lists to top item
    cmbFilter.ListIndex = 0
    cmbType.ListIndex = 0
    Exit Sub
    
errHandler:
    'just incase anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Function VerifyCode() As Boolean

    VerifyCode = True
    If txtTitle.Text = "" Then
        MsgBox "Please enter a Title for the sourcecode snippit.Then click paste on the toolbarmenu.", vbInformation, "SourceCode Organizer"
        VerifyCode = False
        Exit Function
    End If
    If rtbCodeWindow.Text = "" Then
        MsgBox "You must enter some Sourcecode to save a sourcecode snippit.", vbInformation, "SourceCode Organizer"
        VerifyCode = False
        Exit Function
    End If

End Function

Private Sub lstTitles_Click()
    
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim Index As Long
            
    If lstTitles.ListItems.Count < 1 Then 'if there is nothing in the list yet
        Exit Sub
    End If
    'connect to the database and retrieve the selected items details
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    'here is one place where the tag comes in handy, selecting by title was
    'not the best idea, it would slow things down with large amount of snippits
    adoCmd.CommandText = "SELECT * FROM source WHERE id = " & lstTitles.SelectedItem.Tag
    Set adoRS = adoCmd.Execute
    
    'set up the code and notes windows, etc...
    rtbCodeWindow.Text = adoRS("code")
    txtTitle.Text = adoRS("title")
    rtbNotes.Text = adoRS("notes")
    'find the right code type in the drop down
    For Index = 0 To cmbType.ListCount
        If Trim(cmbType.List(Index)) = Trim(adoRS("codetype")) Then
            cmbType.ListIndex = Index
            Exit For
        End If
    Next Index
    
    'nope this aint a new piece of code
    gblNewCode = False
    
    'cleaning house
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
 
End Sub

Private Sub lstTitles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'some code I foundon PSC to easily sort the listview, kudos to whover posted this
    If lstTitles.SortKey <> ColumnHeader.Index - 1 Then
        lstTitles.SortKey = ColumnHeader.Index - 1
        lstTitles.SortOrder = lvwAscending
    Else
        lstTitles.SortOrder = IIf(lstTitles.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If
    lstTitles.Sorted = True
    
End Sub

Private Sub lstTitles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Source Code that are save in the DataBase..."
sBar.font.Size = 9
End Sub

Private Sub mnuAbout_Click()
  AboutF.Show
End Sub



Private Sub mnucontact_Click()
 Call email("extremedexter_z2001@yahoo.com")
End Sub

Private Sub mnuExit_Click()
    'seems kinda abrupt
    End
End Sub

Private Sub mnuFind_Click()
   SSTab1.Tab = 1
End Sub

Private Sub mnunet_Click()
  SSTab1.Tab = 1
End Sub

Private Sub mnuPaste_Click()
    rtbCodeWindow.SelText = Clipboard.GetText
    'Sets the Focus to rtfText
    rtbCodeWindow.SetFocus
End Sub

Private Sub mnuCopy_Click()
If tabSnippit.Tab = 0 Then
    Clipboard.Clear
    Clipboard.SetText rtbCodeWindow.Text
    End If
    If tabSnippit.Tab = 1 Then
    Clipboard.Clear
    Clipboard.SetText rtbnet.Text
    End If
End Sub

Private Sub mnuDelete_Click()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    
    'connnect to the database and delete the current selected snippit
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = "DELETE FROM source WHERE id=" & lstTitles.SelectedItem.Tag
    adoCmd.Execute
            
    'cleanup the house
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    
    'reset the windows
    txtTitle.Text = ""
    rtbCodeWindow.Text = ""
    cmbType.ListIndex = 0
    LoadGridBox
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description
    
End Sub

Private Sub mnunew_Click()
    SSTab1.Tab = 1
    Picture1.Visible = True
    txtTitle.Text = ""
    cmbType.ListIndex = 0
    rtbCodeWindow.Text = ""
    rtbNotes.TextRTF = ""
    gblNewCode = True
End Sub

Private Sub mnuOpen_Click()

    Dim RetVal As String
    cdoOpenDatabase.ShowOpen
    RetVal = cdoOpenDatabase.FileName
    If RetVal = "" Then
        Exit Sub
    End If
    gblConnectString = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & RetVal
    'this will get saved to the registry in a later version
    LoadGridBox  'reload the title list
    gblNewCode = True
    
End Sub

Private Sub mnuPlanetSourceCode_Click()
    'kudos to PSC
    Dim xRet As Long
    xRet = ShellExecute(0, vbNullString, "http://people.we.mediaone.net/retxed/bugproblemicon.htm", vbNullString, App.Path, 1)
End Sub

Private Sub mnuSave_Click()

    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim RetVal As Boolean
    
    'connect to the database
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    If gblNewCode = False Then      'if we are working on an existing snippit
        RetVal = VerifyCode 'check to make sure the user dotted all i's and crossed all t's
        If RetVal = False Then
            Exit Sub
        End If
        'this really should be a stored procedure, but....
        adoCmd.CommandText = "UPDATE source SET title='" & StuffQuotes(txtTitle) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET code='" & StuffQuotes(rtbCodeWindow.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET codetype='" & StuffQuotes(cmbType.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET [datetime]='" & Now & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
        adoCmd.CommandText = "UPDATE source SET notes='" & StuffQuotes(rtbNotes.Text) & "' WHERE id=" & lstTitles.SelectedItem.Tag
        adoCmd.Execute
    Else  'if its new
        RetVal = VerifyCode 'check to make sure the user dotted all i's and crossed all t's
        If RetVal = False Then
            Exit Sub
        End If
        adoCmd.CommandText = "INSERT INTO source ([datetime],title,codetype,code,notes) VALUES('" & Now & "', '" & StuffQuotes(txtTitle) & "', '" & StuffQuotes(cmbType.Text) & "', '" & StuffQuotes(rtbCodeWindow.Text) & "', '" & StuffQuotes(rtbNotes.TextRTF) & "')"
        adoCmd.Execute
        'we need the new identity created for it
        adoCmd.CommandText = "SELECT id FROM source WHERE title = '" & StuffQuotes(txtTitle) & "'"
        Set adoRS = adoCmd.Execute
        gblNewCode = False  'its no longer new
        adoRS.Close
        Set adoRS = Nothing
    End If
    
    'clean everything up
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    
    LoadGridBox  'reset the list
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub mnuWebsite_Click()
    On Error Resume Next
    Dim xRet As Long
    xRet = ShellExecute(0, vbNullString, "http://clik.to/ret", vbNullString, App.Path, 1)
End Sub

Private Sub rtbNotes_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'if anyone can tell me what the heck an effect of 3 is id appreciate it.
    'msdn documentation tells me there is no such thing
    If Data.GetFormat(vbCFText) Then 'if text
        If Effect = 3 Then
            rtbCodeWindow.Text = Data.GetData(vbCFText) 'set the window to the dragged in text
        Else
            rtbNotes.LoadFile Data.GetData(vbCFText), rtfText 'open the dragged in file
        End If
    End If
    If Data.GetFormat(vbCFFiles) Then 'if files from explorer
        rtbNotes.LoadFile Data.Files(1), rtfText  'open the file dragged from windows
    End If

End Sub


Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = ""
End Sub


Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Enter The Code Title and hit search..."
End Sub

Private Sub toolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    'no sense in recreating the wheel, so just call the menu item procedures
    If Button.Tag = "new" Then
        mnunew_Click
    End If
    If Button.Tag = "delete" Then
        mnuDelete_Click
    End If
    If Button.Tag = "save" Then
        mnuSave_Click
    End If
    If Button.Tag = "paste" Then
        mnuPaste_Click
    End If
    If Button.Tag = "copy" Then
        mnuCopy_Click
    End If
    If Button.Tag = "open" Then
        editme_Click
    End If
    If Button.Tag = "find" Then
      SSTab1.Tab = 1
    End If
    If Button.Tag = "print" Then
       ' mnuPrint_Click
    End If
      If Button.Tag = "mod" Then
      mnuModify_Click
    End If
     If Button.Tag = "front" Then
      On Error GoTo HandleErrors
Call Shell("C:\Program Files\Microsoft Office\Office\FRONTPG.EXE", vbNormalFocus)

 Exit Sub
HandleErrors:
  Dim intresponse As Integer
   Select Case Err.Number
        Case 53, 76
         intresponse = MsgBox("File not found.You may have installed it in a diffirent folder or it has been removed.", vbCritical, "Error")
         End Select
         End If
 If Button.Tag = "drw" Then
   On Error GoTo Err
Call Shell("C:\Program Files\Macromedia\Dreamweaver 4\Dreamweaver.exe", vbNormalFocus)
 Exit Sub
Err:
   MsgBox "File not found.You may have installed it in a diffirent folder or it has been removed.", vbCritical, "Error"
End If
 If Button.Tag = "VB" Then
 On Error GoTo Hell
Call Shell("C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe", vbNormalFocus)

 Exit Sub
Hell:
   MsgBox "File not found.You may have installed it in a diffirent folder or it has been removed.", vbCritical, "Error"
  End If
End Sub

Private Sub rtbCodeWindow_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'if anyone can tell me what the heck an effect of 3 is id appreciate it.
    'msdn documentation tells me there is no such thing
    If Data.GetFormat(vbCFText) Then  'if text
        If Effect = 3 Then
            rtbCodeWindow.Text = Data.GetData(vbCFText) 'set the window to the dragged in text
        Else
            rtbCodeWindow.LoadFile Data.GetData(vbCFText) 'open the dragged in file
        End If
    End If
    If Data.GetFormat(vbCFFiles) Then 'if files from explorer
        rtbCodeWindow.LoadFile Data.Files(1) 'open the file dragged from windows
    End If

End Sub

Private Sub Find(strSearch As String)
    
    On Error GoTo errHandler
    
    Dim adoCon As Connection
    Dim adoCmd As Command
    Dim adoRS As Recordset
    Dim cmdtext As String
    
    'not the best find, but it does the job
    'this can get very slow with large numbers of snippits
    cmdtext = "SELECT title FROM source WHERE title like '%" & strSearch & "%'"
    Set adoCon = CreateObject("ADODB.Connection")
    adoCon.open gblConnectString
    Set adoCmd = CreateObject("ADODB.Command")
    adoCmd.ActiveConnection = adoCon
    adoCmd.CommandText = cmdtext
    Set adoRS = adoCmd.Execute
    
    'only take the first returned result, ignore any others
    If adoRS.EOF = False Then
        'find the item in the listview and select it
        lstTitles.SelectedItem = lstTitles.FindItem(adoRS("title"), , lvwPartial)
        lstTitles_Click 'load it into the code window
    Else
        'nothing matched
        MsgBox "Not Found! No Record that match your keyword.Please try another keyword.", vbInformation, "Search"
    End If
    
    'clean up
    adoRS.Close
    Set adoRS = Nothing
    Set adoCmd = Nothing
    adoCon.Close
    Set adoCon = Nothing
    Exit Sub
    
errHandler:
    'in case anything went wrong
    MsgBox "Error: " & Err.Description

End Sub

Private Sub toolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
 Case "view"
   show1_Click
   Case "contents"
   SSTab1.Tab = 0
Case "DB"
  SSTab1.Tab = 1
  Picture1.Visible = False
  Case "SDB"
  SSTab1.Tab = 1
  Picture1.Visible = False
  Case "html"
   
  End Select
End Sub

Private Sub txtTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Enter The Code Title..."
End Sub

Private Sub TxtUrl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sBar.Panels(1).Text = "Enter The Website URL..."
End Sub

Private Sub txWhatSearch_DropDown()
sBar.Panels(1).Text = "Source Code Language Selection.."
End Sub

Private Sub Web_TitleChange(ByVal Text As String)
    Me.Caption = "Browsing : " & Text
End Sub
