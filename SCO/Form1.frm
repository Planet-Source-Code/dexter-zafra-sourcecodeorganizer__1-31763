VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Properties"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      Tabs            =   18
      Tab             =   16
      TabsPerRow      =   10
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Anchor"
      TabPicture(0)   =   "Form1.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command18(4)"
      Tab(0).Control(1)=   "frm2"
      Tab(0).Control(2)=   "cmdOK"
      Tab(0).Control(3)=   "Picture1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Body"
      TabPicture(1)   =   "Form1.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cdc1"
      Tab(1).Control(1)=   "Picture2(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "Command18(5)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "CheckBox"
      TabPicture(2)   =   "Form1.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1(1)"
      Tab(2).Control(1)=   "Command4"
      Tab(2).Control(2)=   "Frame1(1)"
      Tab(2).Control(3)=   "Command18(6)"
      Tab(2).Control(4)=   "Picture1(12)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Div"
      TabPicture(3)   =   "Form1.frx":019E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command6"
      Tab(3).Control(1)=   "Frame1(2)"
      Tab(3).Control(2)=   "Command18(7)"
      Tab(3).Control(3)=   "Picture1(2)"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Font"
      TabPicture(4)   =   "Form1.frx":01BA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command18(8)"
      Tab(4).Control(1)=   "Frame1(3)"
      Tab(4).Control(2)=   "Command8"
      Tab(4).Control(3)=   "Picture1(3)"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "HR"
      TabPicture(5)   =   "Form1.frx":01D6
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Picture1(4)"
      Tab(5).Control(1)=   "Command18(9)"
      Tab(5).Control(2)=   "Frame1(4)"
      Tab(5).Control(3)=   "Command10"
      Tab(5).ControlCount=   4
      TabCaption(6)   =   "Image"
      TabPicture(6)   =   "Form1.frx":01F2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Picture1(5)"
      Tab(6).Control(1)=   "Command12"
      Tab(6).Control(2)=   "Frame1(5)"
      Tab(6).Control(3)=   "Command18(10)"
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Radio"
      TabPicture(7)   =   "Form1.frx":020E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Picture1(6)"
      Tab(7).Control(1)=   "Command15"
      Tab(7).Control(2)=   "Frame1(6)"
      Tab(7).Control(3)=   "Command18(11)"
      Tab(7).Control(4)=   "Picture1(13)"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "Select"
      TabPicture(8)   =   "Form1.frx":022A
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Picture1(7)"
      Tab(8).Control(1)=   "Command17"
      Tab(8).Control(2)=   "Frame1(7)"
      Tab(8).Control(3)=   "Command18(12)"
      Tab(8).Control(4)=   "Picture1(14)"
      Tab(8).ControlCount=   5
      TabCaption(9)   =   "Submit"
      TabPicture(9)   =   "Form1.frx":0246
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Picture1(8)"
      Tab(9).Control(1)=   "Command18(0)"
      Tab(9).Control(2)=   "Command19"
      Tab(9).Control(3)=   "Frame1(8)"
      Tab(9).Control(4)=   "Picture1(15)"
      Tab(9).ControlCount=   5
      TabCaption(10)  =   "Text Area"
      TabPicture(10)  =   "Form1.frx":0262
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Picture1(9)"
      Tab(10).Control(1)=   "Command21"
      Tab(10).Control(2)=   "Frame1(9)"
      Tab(10).Control(3)=   "Command18(1)"
      Tab(10).Control(4)=   "Picture1(16)"
      Tab(10).ControlCount=   5
      TabCaption(11)  =   "Hidden Text Input"
      TabPicture(11)  =   "Form1.frx":027E
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Picture1(10)"
      Tab(11).Control(1)=   "Command23"
      Tab(11).Control(2)=   "Frame1(10)"
      Tab(11).Control(3)=   "Command18(2)"
      Tab(11).Control(4)=   "Picture1(17)"
      Tab(11).ControlCount=   5
      TabCaption(12)  =   "Visible Text Input"
      TabPicture(12)  =   "Form1.frx":029A
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Picture1(11)"
      Tab(12).Control(1)=   "Command25"
      Tab(12).Control(2)=   "Frame1(11)"
      Tab(12).Control(3)=   "Command18(3)"
      Tab(12).ControlCount=   4
      TabCaption(13)  =   "Table"
      TabPicture(13)  =   "Form1.frx":02B6
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Command18(13)"
      Tab(13).Control(1)=   "Frame1(12)"
      Tab(13).Control(2)=   "Command1"
      Tab(13).Control(3)=   "Picture1(18)"
      Tab(13).ControlCount=   4
      TabCaption(14)  =   "Td"
      TabPicture(14)  =   "Form1.frx":02D2
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Command3"
      Tab(14).Control(1)=   "frm1"
      Tab(14).Control(2)=   "Command18(14)"
      Tab(14).ControlCount=   3
      TabCaption(15)  =   "Tr"
      TabPicture(15)  =   "Form1.frx":02EE
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Frame1(14)"
      Tab(15).Control(1)=   "Command18(15)"
      Tab(15).Control(2)=   "Command5"
      Tab(15).ControlCount=   3
      TabCaption(16)  =   "Text Link"
      TabPicture(16)  =   "Form1.frx":030A
      Tab(16).ControlEnabled=   -1  'True
      Tab(16).Control(0)=   "Frame3"
      Tab(16).Control(0).Enabled=   0   'False
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "Image Link"
      TabPicture(17)  =   "Form1.frx":0326
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Frame4"
      Tab(17).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Text Hyper Link"
         Height          =   3255
         Left            =   360
         TabIndex        =   223
         Top             =   840
         Width           =   6375
         Begin VB.CommandButton ok 
            BackColor       =   &H80000005&
            Caption         =   "OK"
            Height          =   375
            Left            =   2880
            TabIndex        =   227
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox link 
            Height          =   285
            Left            =   1680
            TabIndex        =   226
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox address 
            Height          =   285
            Left            =   1680
            TabIndex        =   225
            Top             =   840
            Width           =   3255
         End
         Begin VB.CommandButton cmdbrowse 
            Height          =   320
            Left            =   5040
            Picture         =   "Form1.frx":0342
            Style           =   1  'Graphical
            TabIndex        =   224
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Label3"
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   4
            Left            =   1200
            TabIndex        =   230
            Top             =   1440
            Width           =   15
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Link text:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   600
            TabIndex        =   229
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "URL:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   228
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Image Hyper Link"
         Height          =   3495
         Left            =   -74760
         TabIndex        =   211
         Top             =   720
         Width           =   6495
         Begin VB.CommandButton cmdPicture 
            Height          =   320
            Left            =   5400
            Picture         =   "Form1.frx":06CC
            Style           =   1  'Graphical
            TabIndex        =   218
            ToolTipText     =   "Browse Image"
            Top             =   600
            Width           =   375
         End
         Begin VB.PictureBox picview 
            BackColor       =   &H00FFFFFF&
            Height          =   1575
            Left            =   600
            ScaleHeight     =   1515
            ScaleWidth      =   5115
            TabIndex        =   217
            Top             =   1320
            Width           =   5175
         End
         Begin VB.TextBox alt 
            Height          =   285
            Left            =   1680
            TabIndex        =   216
            Top             =   960
            Width           =   4095
         End
         Begin VB.TextBox pic 
            Height          =   285
            Left            =   1680
            TabIndex        =   215
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox address2 
            Height          =   285
            Left            =   1680
            TabIndex        =   214
            Top             =   240
            Width           =   3615
         End
         Begin VB.CommandButton cmdbrowsefile 
            Height          =   320
            Left            =   5400
            Picture         =   "Form1.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   213
            ToolTipText     =   "Browse File"
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton okiez 
            BackColor       =   &H80000005&
            Caption         =   "OK"
            Default         =   -1  'True
            Height          =   375
            Left            =   2880
            TabIndex        =   212
            Top             =   3000
            Width           =   1215
         End
         Begin MSComDlg.CommonDialog Dialog 
            Left            =   600
            Top             =   2880
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Image Alt:"
            Height          =   195
            Left            =   600
            TabIndex        =   222
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Label3"
            ForeColor       =   &H80000008&
            Height          =   15
            Index           =   5
            Left            =   1200
            TabIndex        =   221
            Top             =   840
            Width           =   15
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Link to Image:"
            Height          =   195
            Index           =   16
            Left            =   600
            TabIndex        =   220
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "URL:"
            Height          =   195
            Index           =   25
            Left            =   600
            TabIndex        =   219
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   -74580
         Picture         =   "Form1.frx":0DE0
         ScaleHeight     =   285
         ScaleWidth      =   375
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   209
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame frm2 
         Caption         =   "Properties of ANCHOR tag"
         Height          =   2295
         Left            =   -74580
         TabIndex        =   198
         Top             =   960
         Width           =   5415
         Begin VB.TextBox txtIDAnch 
            Height          =   285
            Left            =   1080
            TabIndex        =   203
            Top             =   1800
            Width           =   3975
         End
         Begin VB.TextBox txtTitleAnch 
            Height          =   285
            Left            =   1080
            TabIndex        =   202
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtNameAnch 
            Height          =   285
            Left            =   1080
            TabIndex        =   201
            Top             =   1080
            Width           =   3975
         End
         Begin VB.TextBox txtTargetAnch 
            Height          =   285
            Left            =   1080
            TabIndex        =   200
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox txtHrefAnch 
            Height          =   285
            Left            =   1080
            TabIndex        =   199
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label5 
            Caption         =   "ID:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   208
            Top             =   1830
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Title:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   207
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Name:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   206
            Top             =   1110
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Target:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   205
            Top             =   750
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "HREF:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   204
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   0
         Left            =   -74610
         Picture         =   "Form1.frx":13C6
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   3540
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   196
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Properties of BODY  tag"
         Height          =   2295
         Left            =   -74610
         TabIndex        =   177
         Top             =   1020
         Width           =   5415
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   4680
            Picture         =   "Form1.frx":1908
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   189
            Top             =   1342
            Width           =   270
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   4680
            Picture         =   "Form1.frx":1CCA
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   188
            Top             =   982
            Width           =   270
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   2280
            Picture         =   "Form1.frx":208C
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   187
            Top             =   1702
            Width           =   270
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   2280
            Picture         =   "Form1.frx":244E
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   186
            Top             =   1342
            Width           =   270
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   2280
            Picture         =   "Form1.frx":2810
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   185
            Top             =   982
            Width           =   270
         End
         Begin VB.CommandButton cmdOpen 
            Height          =   320
            Left            =   4680
            Picture         =   "Form1.frx":2BD2
            Style           =   1  'Graphical
            TabIndex        =   184
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtVisLink 
            Height          =   285
            Left            =   3480
            TabIndex        =   183
            Text            =   "Navy"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtActivLink 
            Height          =   285
            Left            =   1080
            TabIndex        =   182
            Text            =   "DodgeBlue"
            Top             =   1680
            Width           =   1095
         End
         Begin VB.TextBox txtLinkColor 
            Height          =   285
            Left            =   1080
            TabIndex        =   181
            Text            =   "Blue"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtTextColor 
            Height          =   285
            Left            =   3480
            TabIndex        =   180
            Text            =   "Black"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtBgColor 
            Height          =   285
            Left            =   1080
            TabIndex        =   179
            Text            =   "White"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtBGImage 
            Height          =   285
            Left            =   1080
            TabIndex        =   178
            Top             =   600
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "Activ Link:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   195
            Top             =   1710
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Visited Link:"
            Height          =   255
            Index           =   0
            Left            =   2640
            TabIndex        =   194
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Link Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   193
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "Text Color:"
            Height          =   255
            Left            =   2640
            TabIndex        =   192
            Top             =   990
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Bg. Color:"
            Height          =   255
            Left            =   240
            TabIndex        =   191
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Bg. Image:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   190
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   -74460
         Picture         =   "Form1.frx":2F5C
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   176
         Top             =   1365
         Width           =   345
      End
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   175
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of CHECKBOX tag"
         Height          =   2175
         Index           =   1
         Left            =   -74580
         TabIndex        =   167
         Top             =   1125
         Width           =   5415
         Begin VB.CheckBox Checked 
            Caption         =   "Checked"
            Height          =   255
            Left            =   1560
            TabIndex        =   171
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtValueChk 
            Height          =   285
            Left            =   1560
            TabIndex        =   170
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox txtCaptionChk 
            Height          =   285
            Left            =   1560
            TabIndex        =   169
            Top             =   1320
            Width           =   3375
         End
         Begin VB.TextBox txtNameChk 
            Height          =   285
            Left            =   1560
            TabIndex        =   168
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Caption"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   174
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Value"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   173
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   172
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   166
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of DIV tag"
         Height          =   2295
         Index           =   2
         Left            =   -74655
         TabIndex        =   163
         Top             =   1065
         Width           =   5415
         Begin VB.ComboBox CoAlignDiv 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Align:"
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   165
            Top             =   990
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   -74610
         Picture         =   "Form1.frx":3586
         ScaleHeight     =   300
         ScaleWidth      =   345
         TabIndex        =   162
         Top             =   3540
         Width           =   345
      End
      Begin VB.CommandButton Command8 
         Caption         =   "OK"
         Height          =   375
         Left            =   -72000
         TabIndex        =   161
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of FONT  tag"
         Height          =   2295
         Index           =   3
         Left            =   -74625
         TabIndex        =   153
         Top             =   1020
         Width           =   5415
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   3240
            Picture         =   "Form1.frx":3B68
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   157
            Top             =   840
            Width           =   270
         End
         Begin VB.ComboBox CoFonts 
            Height          =   315
            ItemData        =   "Form1.frx":3F2A
            Left            =   1680
            List            =   "Form1.frx":40A2
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   360
            Width           =   1830
         End
         Begin VB.TextBox txtSizeFo 
            Height          =   285
            Left            =   1680
            TabIndex        =   155
            Text            =   "4"
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtColorFo 
            Height          =   285
            Left            =   1680
            TabIndex        =   154
            Text            =   "Black"
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Size:"
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   160
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Color:"
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   159
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Fonts Style:"
            Height          =   255
            Left            =   480
            TabIndex        =   158
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   152
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of HR tag"
         Height          =   2295
         Index           =   4
         Left            =   -74595
         TabIndex        =   141
         Top             =   1095
         Width           =   5415
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   3720
            Picture         =   "Form1.frx":4796
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   147
            Top             =   1342
            Width           =   270
         End
         Begin VB.TextBox txtColorHR 
            Height          =   285
            Left            =   1560
            TabIndex        =   146
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CheckBox Che 
            Caption         =   "No Shading"
            Height          =   255
            Left            =   1560
            TabIndex        =   145
            Top             =   1800
            Width           =   2055
         End
         Begin VB.TextBox txtSizeHR 
            Height          =   285
            Left            =   3600
            TabIndex        =   144
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtWHR 
            Height          =   285
            Left            =   1560
            TabIndex        =   143
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox CoAlignHR 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Color:"
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   151
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Size:"
            Height          =   255
            Index           =   7
            Left            =   2760
            TabIndex        =   150
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Width:"
            Height          =   255
            Index           =   8
            Left            =   720
            TabIndex        =   149
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Align:"
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   148
            Top             =   390
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   -74580
         Picture         =   "Form1.frx":4B58
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   140
         Top             =   3570
         Width           =   345
      End
      Begin VB.CommandButton Command12 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   139
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of IMAGE tag"
         Height          =   2295
         Index           =   5
         Left            =   -74580
         TabIndex        =   123
         Top             =   1050
         Width           =   5415
         Begin VB.CommandButton Command13 
            Height          =   320
            Left            =   4680
            Picture         =   "Form1.frx":5182
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   360
            Width           =   375
         End
         Begin VB.ComboBox CoAlignImg 
            Height          =   315
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   130
            Top             =   1800
            Width           =   1935
         End
         Begin VB.TextBox txtBorderImg 
            Height          =   285
            Left            =   1080
            TabIndex        =   129
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txtVsImg 
            Height          =   285
            Left            =   3120
            TabIndex        =   128
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtHsImg 
            Height          =   285
            Left            =   1080
            TabIndex        =   127
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtMapImg 
            Height          =   285
            Left            =   1080
            TabIndex        =   126
            Top             =   1080
            Width           =   3975
         End
         Begin VB.TextBox txtAltImg 
            Height          =   285
            Left            =   1080
            TabIndex        =   125
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox txtSourceImg 
            Height          =   285
            Left            =   1080
            TabIndex        =   124
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label7 
            Caption         =   "Align:"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   138
            Top             =   1830
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Vspace:"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   137
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Border:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   136
            Top             =   1830
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Hspace:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   135
            Top             =   1470
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Use Map:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   134
            Top             =   1110
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Alt Text:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   133
            Top             =   750
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Source:"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   132
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   -74460
         Picture         =   "Form1.frx":550C
         ScaleHeight     =   300
         ScaleWidth      =   330
         TabIndex        =   122
         Top             =   1425
         Width           =   330
      End
      Begin VB.CommandButton Command15 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   121
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of RADIO tag"
         Height          =   2175
         Index           =   6
         Left            =   -74580
         TabIndex        =   113
         Top             =   1185
         Width           =   5415
         Begin VB.CheckBox chRa 
            Caption         =   "Checked"
            Height          =   255
            Left            =   1440
            TabIndex        =   117
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox txtValueRa 
            Height          =   285
            Left            =   1440
            TabIndex        =   116
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox txtCaptionRa 
            Height          =   285
            Left            =   1440
            TabIndex        =   115
            Top             =   1320
            Width           =   3375
         End
         Begin VB.TextBox txtNameRa 
            Height          =   285
            Left            =   1440
            TabIndex        =   114
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Caption"
            Height          =   255
            Index           =   11
            Left            =   600
            TabIndex        =   120
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Value"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   119
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Index           =   12
            Left            =   600
            TabIndex        =   118
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   7
         Left            =   -74475
         Picture         =   "Form1.frx":5A9E
         ScaleHeight     =   330
         ScaleWidth      =   435
         TabIndex        =   112
         Top             =   1320
         Width           =   435
      End
      Begin VB.CommandButton Command17 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   111
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of Select input"
         Height          =   2175
         Index           =   7
         Left            =   -74595
         TabIndex        =   105
         Top             =   1080
         Width           =   5415
         Begin VB.CheckBox Multy 
            Caption         =   "Multiple Selection ?"
            Height          =   255
            Left            =   2280
            TabIndex        =   108
            Top             =   1230
            Width           =   1815
         End
         Begin VB.TextBox txtSizeSel 
            Height          =   285
            Left            =   1440
            TabIndex        =   107
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNameSel 
            Height          =   285
            Left            =   1440
            TabIndex        =   106
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "Size:"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   110
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Index           =   13
            Left            =   600
            TabIndex        =   109
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   -74505
         Picture         =   "Form1.frx":6270
         ScaleHeight     =   225
         ScaleWidth      =   330
         TabIndex        =   104
         Top             =   1395
         Width           =   330
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   0
         Left            =   -70545
         TabIndex        =   103
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command19 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   102
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of Submit Button input"
         Height          =   2175
         Index           =   8
         Left            =   -74625
         TabIndex        =   93
         Top             =   1155
         Width           =   5415
         Begin VB.TextBox txtCaptionSub 
            Height          =   285
            Left            =   1440
            TabIndex        =   97
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox txtHSub 
            Height          =   285
            Left            =   3240
            TabIndex        =   96
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtWSub 
            Height          =   285
            Left            =   1440
            TabIndex        =   95
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtNameSub 
            Height          =   285
            Left            =   1440
            TabIndex        =   94
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Caption:"
            Height          =   255
            Index           =   14
            Left            =   600
            TabIndex        =   101
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Height:"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   100
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Width:"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   99
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Index           =   15
            Left            =   600
            TabIndex        =   98
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   -74460
         Picture         =   "Form1.frx":66AE
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   92
         Top             =   1395
         Width           =   315
      End
      Begin VB.CommandButton Command21 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   91
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of TEXT AREA input"
         Height          =   2175
         Index           =   9
         Left            =   -74580
         TabIndex        =   84
         Top             =   1155
         Width           =   5415
         Begin VB.TextBox txtRowsArea 
            Height          =   285
            Left            =   3240
            TabIndex        =   87
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtColsArea 
            Height          =   285
            Left            =   1440
            TabIndex        =   86
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtNameArea 
            Height          =   285
            Left            =   1440
            TabIndex        =   85
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label3 
            Caption         =   "Rows:"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   90
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cols:"
            Height          =   255
            Index           =   7
            Left            =   600
            TabIndex        =   89
            Top             =   1230
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Index           =   16
            Left            =   600
            TabIndex        =   88
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   -74460
         Picture         =   "Form1.frx":6BF0
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   83
         Top             =   1380
         Width           =   345
      End
      Begin VB.CommandButton Command23 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   82
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of Hiden text input"
         Height          =   2175
         Index           =   10
         Left            =   -74580
         TabIndex        =   77
         Top             =   1140
         Width           =   5415
         Begin VB.TextBox txtNameHid 
            Height          =   285
            Left            =   1185
            TabIndex        =   79
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtValueHid 
            Height          =   285
            Left            =   1200
            TabIndex        =   78
            Top             =   960
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   81
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Value"
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   80
            Top             =   990
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   11
         Left            =   -74580
         Picture         =   "Form1.frx":721A
         ScaleHeight     =   390
         ScaleWidth      =   435
         TabIndex        =   76
         Top             =   3495
         Width           =   435
      End
      Begin VB.CommandButton Command25 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71985
         TabIndex        =   75
         Top             =   3555
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of TEXT input"
         Height          =   2175
         Index           =   11
         Left            =   -74580
         TabIndex        =   66
         Top             =   1095
         Width           =   5415
         Begin VB.TextBox txtMaxLenTxtVis 
            Height          =   285
            Left            =   3120
            TabIndex        =   70
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtValueTxtVis 
            Height          =   285
            Left            =   1080
            TabIndex        =   69
            Top             =   960
            Width           =   3375
         End
         Begin VB.TextBox txtSizetxtVis 
            Height          =   285
            Left            =   1080
            TabIndex        =   68
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtNavetextVis 
            Height          =   285
            Left            =   1080
            TabIndex        =   67
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label1 
            Caption         =   "Max Length:"
            Height          =   255
            Index           =   20
            Left            =   2040
            TabIndex        =   74
            Top             =   1470
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Size:"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   73
            Top             =   1470
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Value"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   72
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   71
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   -70545
         TabIndex        =   65
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   2
         Left            =   -70545
         TabIndex        =   64
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   3
         Left            =   -70530
         TabIndex        =   63
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   4
         Left            =   -70530
         TabIndex        =   62
         Top             =   3570
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   5
         Left            =   -70545
         TabIndex        =   61
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   6
         Left            =   -70515
         TabIndex        =   60
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   7
         Left            =   -70575
         TabIndex        =   59
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   8
         Left            =   -70560
         TabIndex        =   58
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   9
         Left            =   -70530
         TabIndex        =   57
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   10
         Left            =   -70515
         TabIndex        =   56
         Top             =   3555
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   11
         Left            =   -70515
         TabIndex        =   55
         Top             =   3540
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   12
         Left            =   -70530
         TabIndex        =   54
         Top             =   3540
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   -74550
         Picture         =   "Form1.frx":7B4C
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   53
         Top             =   3645
         Width           =   345
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   2
         Left            =   -74655
         Picture         =   "Form1.frx":8176
         ScaleHeight     =   270
         ScaleWidth      =   345
         TabIndex        =   52
         Top             =   3585
         Width           =   345
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   4
         Left            =   -74595
         Picture         =   "Form1.frx":86C8
         ScaleHeight     =   270
         ScaleWidth      =   360
         TabIndex        =   51
         Top             =   3585
         Width           =   360
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   -74610
         Picture         =   "Form1.frx":8C1A
         ScaleHeight     =   300
         ScaleWidth      =   330
         TabIndex        =   50
         Top             =   3555
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   14
         Left            =   -74610
         Picture         =   "Form1.frx":91AC
         ScaleHeight     =   330
         ScaleWidth      =   435
         TabIndex        =   49
         Top             =   3435
         Width           =   435
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   -74625
         Picture         =   "Form1.frx":997E
         ScaleHeight     =   225
         ScaleWidth      =   330
         TabIndex        =   48
         Top             =   3555
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   -74610
         Picture         =   "Form1.frx":9DBC
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   47
         Top             =   3495
         Width           =   315
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   17
         Left            =   -74595
         Picture         =   "Form1.frx":A2FE
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   46
         Top             =   3525
         Width           =   345
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   13
         Left            =   -70500
         TabIndex        =   45
         Top             =   3660
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of TABLE tag"
         Height          =   2295
         Index           =   12
         Left            =   -74565
         TabIndex        =   28
         Top             =   1125
         Width           =   5415
         Begin VB.TextBox txtWTable 
            Height          =   285
            Left            =   1050
            TabIndex        =   37
            Top             =   855
            Width           =   1080
         End
         Begin VB.TextBox txtBordertable 
            Height          =   285
            Left            =   1050
            TabIndex        =   36
            Top             =   1335
            Width           =   1095
         End
         Begin VB.ComboBox CoAlignTable 
            Height          =   315
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txtSpacingTable 
            Height          =   285
            Left            =   3450
            TabIndex        =   34
            Top             =   855
            Width           =   1080
         End
         Begin VB.TextBox txtPaddingTable 
            Height          =   285
            Left            =   3450
            TabIndex        =   33
            Top             =   1335
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   9
            Left            =   2250
            Picture         =   "Form1.frx":A928
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   32
            Top             =   1815
            Width           =   270
         End
         Begin VB.TextBox txtBGTable 
            Height          =   285
            Left            =   1080
            TabIndex        =   31
            Top             =   1785
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   10
            Left            =   4650
            Picture         =   "Form1.frx":ACEA
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   30
            Top             =   1830
            Width           =   270
         End
         Begin VB.TextBox txtBColTable 
            Height          =   285
            Left            =   3465
            TabIndex        =   29
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Align:"
            Height          =   255
            Index           =   18
            Left            =   210
            TabIndex        =   44
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Width:"
            Height          =   255
            Index           =   10
            Left            =   210
            TabIndex        =   43
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Border:"
            Height          =   255
            Index           =   3
            Left            =   210
            TabIndex        =   42
            Top             =   1350
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Cell Spacing:"
            Height          =   255
            Index           =   11
            Left            =   2355
            TabIndex        =   41
            Top             =   870
            Width           =   990
         End
         Begin VB.Label Label4 
            Caption         =   "Cell Padding:"
            Height          =   255
            Index           =   4
            Left            =   2370
            TabIndex        =   40
            Top             =   1350
            Width           =   1065
         End
         Begin VB.Label Label4 
            Caption         =   "BgColor:"
            Height          =   255
            Index           =   8
            Left            =   225
            TabIndex        =   39
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Border Col."
            Height          =   255
            Index           =   9
            Left            =   2610
            TabIndex        =   38
            Top             =   1815
            Width           =   855
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71925
         TabIndex        =   27
         Top             =   3660
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   -74550
         Picture         =   "Form1.frx":B0AC
         ScaleHeight     =   300
         ScaleWidth      =   345
         TabIndex        =   26
         Top             =   3645
         Width           =   345
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71970
         TabIndex        =   25
         Top             =   3615
         Width           =   1335
      End
      Begin VB.Frame frm1 
         Caption         =   "Properties of TD tag"
         Height          =   2595
         Left            =   -74625
         TabIndex        =   9
         Top             =   870
         Width           =   5415
         Begin VB.TextBox txtRSTD 
            Height          =   285
            Left            =   3630
            TabIndex        =   17
            Top             =   1785
            Width           =   1095
         End
         Begin VB.TextBox txtCSTD 
            Height          =   285
            Left            =   3645
            TabIndex        =   16
            Top             =   1305
            Width           =   1080
         End
         Begin VB.ComboBox CoValignTD 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   810
            Width           =   3615
         End
         Begin VB.TextBox txtHTD 
            Height          =   285
            Left            =   1110
            TabIndex        =   14
            Top             =   1785
            Width           =   1095
         End
         Begin VB.TextBox txtWTD 
            Height          =   285
            Left            =   1095
            TabIndex        =   13
            Top             =   1305
            Width           =   1080
         End
         Begin VB.ComboBox CoAlignTD 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   330
            Width           =   3615
         End
         Begin VB.TextBox txtBgTD 
            Height          =   285
            Left            =   1110
            TabIndex        =   11
            Top             =   2190
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   2280
            Picture         =   "Form1.frx":B68E
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   10
            Top             =   2220
            Width           =   270
         End
         Begin VB.Label Label4 
            Caption         =   "Row Span:"
            Height          =   255
            Index           =   5
            Left            =   2430
            TabIndex        =   24
            Top             =   1800
            Width           =   1065
         End
         Begin VB.Label Label2 
            Caption         =   "Col Span:"
            Height          =   255
            Index           =   12
            Left            =   2415
            TabIndex        =   23
            Top             =   1320
            Width           =   990
         End
         Begin VB.Label Label4 
            Caption         =   "Heigth:"
            Height          =   255
            Index           =   6
            Left            =   270
            TabIndex        =   22
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Width:"
            Height          =   255
            Index           =   13
            Left            =   270
            TabIndex        =   21
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Valign:"
            Height          =   255
            Index           =   19
            Left            =   270
            TabIndex        =   20
            Top             =   810
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Align:"
            Height          =   255
            Index           =   23
            Left            =   270
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "BgColor:"
            Height          =   255
            Index           =   7
            Left            =   255
            TabIndex        =   18
            Top             =   2205
            Width           =   855
         End
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   14
         Left            =   -70545
         TabIndex        =   8
         Top             =   3615
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Caption         =   "Properties of TR tag"
         Height          =   2295
         Index           =   14
         Left            =   -74565
         TabIndex        =   3
         Top             =   1140
         Width           =   5415
         Begin VB.ComboBox CoAlignTR 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   600
            Width           =   3615
         End
         Begin VB.ComboBox CoValignTR 
            Height          =   315
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1095
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Align:"
            Height          =   255
            Index           =   24
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Vert. Align:"
            Height          =   255
            Index           =   14
            Left            =   225
            TabIndex        =   6
            Top             =   1140
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   15
         Left            =   -70500
         TabIndex        =   2
         Top             =   3585
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         Height          =   375
         Left            =   -71925
         TabIndex        =   1
         Top             =   3585
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog cdc1 
         Left            =   -69345
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RGBValues(3) As Long 'Red & Green & Blue Values
Dim CurrColor As Long    'Currently Selected Color in Long Value

Private Sub cmdbrowse_Click()
On Error Resume Next
 cdc1.ShowOpen
 address.Text = cdc1.FileName
End Sub

Private Sub cmdbrowsefile_Click()
On Error Resume Next
 cdc1.ShowOpen
 address2.Text = cdc1.FileName
End Sub

'# Anchor
Private Sub cmdOK_Click()
On Error Resume Next
 Dim strChe As String
 Dim insert As String
 insert = "<a href=" & Chr(34) & txtHrefAnch.Text & Chr(34) & " target=" & Chr(34) & txtTargetAnch.Text & Chr(34) & " name=" & Chr(34) & txtNameAnch.Text & Chr(34) & " title=" & Chr(34) & txtTitleAnch.Text & Chr(34) & " id=" & Chr(34) & txtIDAnch.Text & Chr(34) & "></a>"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

Private Sub cmdPicture_Click()
Dialog.DialogTitle = "Browse for Picture..." 'set the dialog title
Dialog.Filter = "GIF|*.gif*|JPG|*.jpg*|BMP|*.bmp"
Dialog.ShowOpen
pic.Text = Dialog.FileName
picview.Picture = LoadPicture(Dialog.FileName)
End Sub

'# Table
Private Sub Command1_Click()
 On Error Resume Next
 Dim insert As String
 insert = "<Table width=" & Chr(34) & txtWTable.Text & Chr(34) & " border=" & Chr(34) & txtBordertable.Text & Chr(34) & " cellspacing=" & Chr(34) & txtSpacingTable.Text & Chr(34) & " cellpadding=" & Chr(34) & txtPaddingTable.Text & Chr(34) & " align=" & Chr(34) & CoAlignTable.Text & Chr(34) & " bgcolor=" & Chr(34) & txtBGTable.Text & Chr(34) & " bordercolor=" & Chr(34) & txtBColTable.Text & Chr(34) & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# HR
Private Sub Command10_Click()
 On Error Resume Next
 Dim strChe As String
 Dim insert As String
  If Che.Value = 1 Then
   strChe = " noshade"
  Else
   strChe = ""
  End If
 insert = "<hr width=" & Chr(34) & txtWHR.Text & Chr(34) & " align=" & Chr(34) & CoAlignHR.Text & Chr(34) & " size=" & Chr(34) & txtSizeHR.Text & Chr(34) & " color=" & Chr(34) & txtColorHR.Text & Chr(34) & " " & strChe & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Image
Private Sub Command12_Click()
 On Error Resume Next
 Dim insert As String
 insert = "<img src=" & Chr(34) & txtSourceImg.Text & Chr(34) & " hspace=" & Chr(34) & txtHsImg.Text & Chr(34) & " vspace=" & Chr(34) & txtVsImg.Text & Chr(34) & " border=" & Chr(34) & txtBorderImg.Text & Chr(34) & " align=" & Chr(34) & CoAlignImg.Text & Chr(34) & " alt=" & Chr(34) & txtAltImg.Text & Chr(34) & " usemap=" & Chr(34) & "#" & txtMapImg.Text & Chr(34) & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Radio
Private Sub Command15_Click()
On Error Resume Next
Dim insert As String, strChe As String
  If chRa.Value = 1 Then
   strChe = "checked"
  Else
   strChe = ""
  End If
 insert = "<input type=" & Chr(34) & "radio" & Chr(34) & " name=" & Chr(34) & txtNameRa.Text & Chr(34) & " value=" & Chr(34) & txtValueRa.Text & Chr(34) & " " & strChe & ">" & txtCaptionRa.Text
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Select
Private Sub Command17_Click()
On Error Resume Next
 Dim strChe As String, insert As String
  If Multy.Value = 1 Then
   strChe = " multiple"
  Else
   strChe = ""
  End If
  insert = "<select name=" & Chr(34) & txtNameSel.Text & Chr(34) & " size=" & Chr(34) & txtSizeSel.Text & Chr(34) & " " & strChe & ">" & vbCrLf & "<option value=1></option>" & vbCrLf & "</select>" & vbCrLf
 HTMLEditor.RichTextBox1.SelText = insert
  Unload Me
End Sub

'# Submit Button
Private Sub Command19_Click()
 Dim insert As String
 insert = "<input type=" & Chr(34) & "submit" & Chr(34) & " name=" & Chr(34) & txtNameSub.Text & Chr(34) & " value=" & Chr(34) & txtCaptionSub.Text & Chr(34) & " width=" & Chr(34) & txtWSub.Text & Chr(34) & " height=" & Chr(34) & txtHSub.Text & Chr(34) & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Body
Private Sub Command2_Click()
On Error Resume Next
Dim strChe As String
Dim insert As String

 If txtBGImage.Text = "" Then
  insert = "<body bgcolor=" & Chr(34) & txtBgColor.Text & Chr(34) & " text=" & Chr(34) & txtTextColor.Text & Chr(34) & " link=" & Chr(34) & txtLinkColor.Text & Chr(34) & " vlink=" & Chr(34) & txtVisLink.Text & Chr(34) & " alink=" & Chr(34) & txtActivLink.Text & Chr(34) & ">"
  HTMLEditor.RichTextBox1.SelText = insert
  Unload Me
  Exit Sub
 End If

 insert = "<body background=" & Chr(34) & txtBGImage.Text & Chr(34) & " bgcolor=" & Chr(34) & txtBgColor.Text & Chr(34) & " text=" & Chr(34) & txtTextColor.Text & Chr(34) & " link=" & Chr(34) & txtLinkColor.Text & Chr(34) & " vlink=" & Chr(34) & txtVisLink.Text & Chr(34) & " alink=" & Chr(34) & txtActivLink.Text & Chr(34) & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Text Area
Private Sub Command21_Click()
On Error Resume Next
 Dim insert As String
 insert = "<textarea cols=" & Chr(34) & txtColsArea.Text & Chr(34) & " rows=" & Chr(34) & txtRowsArea.Text & Chr(34) & " name=" & Chr(34) & txtNameArea.Text & Chr(34) & "></textarea>"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Hidden Text Input
Private Sub Command23_Click()
On Error Resume Next
 Dim strChe As String
 Dim insert As String
 insert = "<input type=" & Chr(34) & "hidden" & Chr(34) & " name=" & Chr(34) & txtNameHid.Text & Chr(34) & " value=" & Chr(34) & txtValueHid.Text & Chr(34) & ">"
HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Visible Text Input
Private Sub Command25_Click()
On Error Resume Next
 Dim insert As String
 insert = "<input type=" & Chr(34) & "text" & Chr(34) & " name=" & Chr(34) & txtNavetextVis.Text & Chr(34) & " value=" & Chr(34) & txtValueTxtVis.Text & Chr(34) & " size=" & Chr(34) & txtSizetxtVis.Text & Chr(34) & " maxlength=" & Chr(34) & txtMaxLenTxtVis.Text & Chr(34) & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Cancel Button
Private Sub Command18_Click(Index As Integer)
On Error Resume Next
 Unload Me
End Sub
'# TD
Private Sub Command3_Click()
On Error Resume Next
 Dim insert As String
 insert = "<Td width=" & Chr(34) & txtWTD.Text & Chr(34) & " height=" & Chr(34) & txtHTD.Text & Chr(34) & " colspan=" & Chr(34) & txtCSTD.Text & Chr(34) & " rowspan=" & Chr(34) & txtRSTD.Text & Chr(34) & " align=" & Chr(34) & CoAlignTD.Text & Chr(34) & " valign=" & Chr(34) & CoValignTD.Text & Chr(34) & " bgcolor=" & Chr(34) & txtBgTD.Text & Chr(34) & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# CheckBox
Private Sub Command4_Click()
On Error Resume Next
 Dim strChe As String
 Dim insert As String
  If Checked.Value = 1 Then
   strChe = "checked"
  Else
   strChe = ""
  End If
  insert = "<input type=" & Chr(34) & "checkbox" & Chr(34) & " name=" & Chr(34) & txtNameChk.Text & Chr(34) & " value=" & Chr(34) & txtValueChk.Text & Chr(34) & " " & strChe & ">" & txtCaptionChk.Text
  HTMLEditor.RichTextBox1.SelText = insert
  Unload Me
End Sub

'#Tr
Private Sub Command5_Click()
On Error Resume Next
 Dim insert As String
 insert = "<tr align=" & Chr(34) & CoAlignTR.Text & Chr(34) & " valign=" & Chr(34) & CoValignTR.Text & Chr(34) & ">"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Div
Private Sub Command6_Click()
On Error Resume Next
 Dim insert As String
 insert = "<div align=" & Chr(34) & CoAlignDiv.Text & Chr(34) & "></div>"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

'# Fonts
Private Sub Command8_Click()
On Error Resume Next
 Dim insert As String
 insert = "<font face=" & Chr(34) & CoFonts.Text & Chr(34) & " color=" & Chr(34) & txtColorFo.Text & Chr(34) & " size=" & Chr(34) & txtSizeFo.Text & Chr(34) & "></font>"
 HTMLEditor.RichTextBox1.SelText = insert
 Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
'Select The whole tag including <>
 HTMLEditor.RichTextBox1.SelLength = HTMLEditor.RichTextBox1.SelLength + 1
 TagLength = Len(HTMLEditor.RichTextBox1.SelText)
 HTMLEditor.RichTextBox1.SelStart = HTMLEditor.RichTextBox1.SelStart - 1
 HTMLEditor.RichTextBox1.SelLength = TagLength + 1

 Dim I As Integer
 CoAlignTD.AddItem "Left"
 CoAlignTD.AddItem "Center"
 CoAlignTD.AddItem "Right"
 CoAlignTD.AddItem "Justify"
 CoAlignTD.ListIndex = 0
 CoValignTD.AddItem "Top"
 CoValignTD.AddItem "Bottom"
 CoValignTD.AddItem "Middle"
 CoValignTD.AddItem "Baseline"
 CoValignTD.ListIndex = 0

 CoAlignDiv.AddItem "Left"
 CoAlignDiv.AddItem "Center"
 CoAlignDiv.AddItem "Right"
 CoAlignDiv.AddItem "Justify"


 CoAlignTR.AddItem "Left"
 CoAlignTR.AddItem "Center"
 CoAlignTR.AddItem "Right"
 CoAlignTR.AddItem "Justify"
 CoAlignTR.AddItem "Char"
 CoValignTR.AddItem "Top"
 CoValignTR.AddItem "Middle"
 CoValignTR.AddItem "Bottom"
 CoValignTR.AddItem "Baseline"

 CoAlignTable.AddItem "left"
 CoAlignTable.AddItem "center"
 CoAlignTable.AddItem "right"

 CoAlignHR.AddItem "Left"
 CoAlignHR.AddItem "Center"
 CoAlignHR.AddItem "Right"

 CoAlignImg.AddItem "Left"
 CoAlignImg.AddItem "Right"
 CoAlignImg.AddItem "Top"
 CoAlignImg.AddItem "Middle"
 CoAlignImg.AddItem "Bottom"
 CoAlignImg.AddItem "Texttop"
 CoAlignImg.AddItem "baseline"

 CoAlignTable.ListIndex = 0
 CoAlignDiv.ListIndex = 0
 CoAlignHR.ListIndex = 0
 CoAlignImg.ListIndex = 0
 CoAlignTR.ListIndex = 0
 CoValignTR.ListIndex = 0


Select Case TabNumber
         Case 0
          GetNameProperty
          txtNameAnch.Text = NameValueInTag
          GetHrefProperty
          txtHrefAnch.Text = HrefValueInTag
          GetTargetProperty
          txtTargetAnch.Text = TargetValueInTag
          GetTitleProperty
          txtTitleAnch.Text = TitleValueInTag
          GetIDProperty
          txtIDAnch.Text = IDValueInTag
         Case 1
          GetBgColorProperty
          txtBgColor.Text = BgColorValueInTag
          GetTextProperty
          txtTextColor.Text = TextValueInTag
          GetLinkProperty
          txtLinkColor.Text = LinkValueInTag
          GetVlinkProperty
          txtVisLink.Text = VlinkValueInTag
          GetAlinkProperty
          txtActivLink.Text = AlinkValueInTag
          GetBGPProperty
          txtBGImage.Text = BgValueInTag
         Case 2
          GetNameProperty
          txtNameChk.Text = NameValueInTag
         Case 3
          GetAlignProperty
          If AlignValueInTag = "CENTER" Then CoAlignDiv.ListIndex = 1
          If AlignValueInTag = "LEFT" Then CoAlignDiv.ListIndex = 0
          If AlignValueInTag = "RIGHT" Then CoAlignDiv.ListIndex = 2
          If AlignValueInTag = "JUSTIFY" Then CoAlignDiv.ListIndex = 3
         Case 4
          GetSizeProperty
          txtSizeFo.Text = SizeValueInTag
          GetColorProperty
          txtColorFo.Text = ColorValueInTag
          GetFaceProperty

If FaceValueInTag = "" Then GoTo Continue:
All = CoFonts.ListCount
For I = 1 To All
 If LCase(CoFonts.List(I)) = FaceValueInTag Then
   CoFonts.ListIndex = I
   GoTo Continue:
 Else
   CoFonts.ListIndex = 1
 End If
Next I
Continue:
         Case 5
          GetWidthProperty
          txtWHR.Text = WidthValueInTag
          GetSizeProperty
          txtSizeHR.Text = SizeValueInTag
          GetColorProperty
          txtColorHR.Text = ColorValueInTag
          GetAlignProperty
          If AlignValueInTag = "CENTER" Then CoAlignHR.ListIndex = 1
          If AlignValueInTag = "LEFT" Then CoAlignHR.ListIndex = 0
          If AlignValueInTag = "RIGHT" Then CoAlignHR.ListIndex = 2
         Case 6
          GetAlignProperty
          If AlignValueInTag = "LEFT" Then CoAlignImg.ListIndex = 0
          If AlignValueInTag = "RIGHT" Then CoAlignImg.ListIndex = 1
          If AlignValueInTag = "TOP" Then CoAlignImg.ListIndex = 2
          If AlignValueInTag = "MIDDLE" Then CoAlignImg.ListIndex = 3
          If AlignValueInTag = "BOTTOM" Then CoAlignImg.ListIndex = 4
          If AlignValueInTag = "TEXTTOP" Then CoAlignImg.ListIndex = 5
          If AlignValueInTag = "BASELINE" Then CoAlignImg.ListIndex = 6
          GetBorderProperty
          txtBorderImg.Text = BorderValueInTag
          GetHspaceProperty
          txtHsImg.Text = HspaceValueInTag
          GetVspaceProperty
          txtVsImg.Text = VspaceValueInTag
          GetUsemapProperty
          txtMapImg.Text = UsemapValueInTag
          GetAltProperty
          txtAltImg.Text = AltValueInTag
          GetSrcProperty
          txtSourceImg.Text = SrcValueInTag
         Case 7
          GetNameProperty
          txtNameRa.Text = NameValueInTag
         Case 8
          GetNameProperty
          txtNameSel.Text = NameValueInTag
          GetSizeProperty
          txtSizeSel.Text = SizeValueInTag
         Case 13
          GetWidthProperty
          txtWTable.Text = WidthValueInTag
          GetBgColorProperty
          txtBGTable.Text = BgColorValueInTag
          GetAlignProperty
          If AlignValueInTag = "CENTER" Then CoAlignTable.ListIndex = 1
          If AlignValueInTag = "LEFT" Then CoAlignTable.ListIndex = 0
          If AlignValueInTag = "RIGHT" Then CoAlignTable.ListIndex = 2
          GetBorderProperty
          txtBordertable.Text = BorderValueInTag
          GetBorderCProperty
          txtBColTable.Text = BcValueInTag
          GetCellSProperty
          txtSpacingTable.Text = CellsValueInTag
          GetCellPProperty
          txtPaddingTable.Text = CellPValueInTag
         Case 14
          GetWidthProperty
          GetHeightProperty
          txtHTD.Text = HeightValueInTag
          txtWTD.Text = WidthValueInTag
          GetBgColorProperty
          txtBgTD.Text = BgColorValueInTag
          GetAlignProperty
          If AlignValueInTag = "CENTER" Then CoAlignTD.ListIndex = 1
          If AlignValueInTag = "LEFT" Then CoAlignTD.ListIndex = 0
          If AlignValueInTag = "RIGHT" Then CoAlignTD.ListIndex = 2
          If AlignValueInTag = "JUSTIFY" Then CoAlignTD.ListIndex = 3
          GetValignProperty
          If ValignValueInTag = "TOP" Then CoValignTD.ListIndex = 0
          If ValignValueInTag = "BOTTOM" Then CoValignTD.ListIndex = 1
          If ValignValueInTag = "MIDDLE" Then CoValignTD.ListIndex = 2
          If ValignValueInTag = "BASELINE" Then CoValignTD.ListIndex = 3
          GetColspanProperty
          txtCSTD.Text = ColspanValueInTag
          GetRowspanProperty
          txtRSTD.Text = RowpanValueInTag
         Case 15
          GetAlignProperty
          If AlignValueInTag = "LEFT" Then CoAlignTR.ListIndex = 0
          If AlignValueInTag = "CENTER" Then CoAlignTR.ListIndex = 1
          If AlignValueInTag = "RIGHT" Then CoAlignTR.ListIndex = 2
          If AlignValueInTag = "JUSTIFY" Then CoAlignTR.ListIndex = 3
          If AlignValueInTag = "CHAR" Then CoAlignTR.ListIndex = 4
          GetValignProperty
          If ValignValueInTag = "TOP" Then CoValignTR.ListIndex = 0
          If ValignValueInTag = "MIDDLE" Then CoValignTR.ListIndex = 1
          If ValignValueInTag = "BOTTOM" Then CoValignTR.ListIndex = 2
          If ValignValueInTag = "BASELINE" Then CoValignTR.ListIndex = 3
         Case 9
          GetWidthProperty
          GetHeightProperty
          GetValueProperty
          GetNameProperty
          txtCaptionSub.Text = ValueValueInTag
          txtHSub.Text = HeightValueInTag
          txtWSub.Text = WidthValueInTag
          txtNameSub.Text = NameValueInTag
         Case 10
          GetNameProperty
          txtNameArea.Text = NameValueInTag
          GetColsProperty
          txtColsArea.Text = ColsValueInTag
          GetRowsProperty
          txtRowsArea.Text = RowsValueInTag
         Case 11
          GetNameProperty
          GetValueProperty
          txtNameHid.Text = NameValueInTag
          txtValueHid.Text = ValueValueInTag
         Case 12
          GetNameProperty
          GetValueProperty
          txtNavetextVis.Text = NameValueInTag
          txtValueTxtVis.Text = ValueValueInTag
          GetSizeProperty
          txtSizetxtVis.Text = SizeValueInTag
          GetMaxLProperty
          txtMaxLenTxtVis.Text = MaxValueInTag
 End Select
  SSTab1.Tab = TabNumber
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
On Error Resume Next
 TabNumber = 0
End Sub

Private Sub ok_Click()
On Error GoTo colorerror
HTMLEditor.RichTextBox1.SelText = HTMLEditor.RichTextBox1.SelText + "<a href=" & Chr(34) & address.Text & Chr(34) & ">" + link.Text + "</a>"
Link1.address.Text = ""
Link1.link.Text = ""
Exit Sub
colorerror:
Unload Me
End Sub

Private Sub okiez_Click()
On Error GoTo colorerror
HTMLEditor.RichTextBox1.SelText = HTMLEditor.RichTextBox1.SelText + "<a href=" & Chr(34) & address2.Text & Chr(34) & ">" + "<img alt=" & Chr(34) & alt.Text & Chr(34) & ", src=" & Chr(34) & pic.Text & Chr(34) & " Border=0>" + "</a>"
Link2.address2.Text = ""
Link2.pic.Text = ""
Link2.alt.Text = ""
Exit Sub
colorerror:
Unload Me
End Sub

Private Sub Picture2_Click(Index As Integer)
On Error Resume Next
If Index = 10 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtBColTable)
End If
If Index = 9 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtBGTable)
End If
If Index = 8 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtBgTD)
End If
If Index = 7 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtColorHR)
End If
If Index = 6 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtColorFo)
End If
If Index = 5 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtBgColor)
End If
If Index = 1 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtLinkColor)
End If
If Index = 2 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtActivLink)
End If
If Index = 3 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtTextColor)
End If
If Index = 4 Then
 cdc1.ShowColor
 CurrColor = cdc1.color
 GetRGB
 Call HexColor(txtVisLink)
End If
End Sub

Function GetRGB()
On Error Resume Next
 RGBValues(3) = CLng(CurrColor)
 RGBValues(0) = RGBValues(3) And 255
 RGBValues(1) = (RGBValues(3) And 65280) \ 256&
 RGBValues(2) = (RGBValues(3) And 16711680) \ 65535
 txtR.Text = RGBValues(0)
 txtG.Text = RGBValues(1)
 txtB.Text = RGBValues(2)
End Function

Function HexColor(txtF As TextBox)
On Error Resume Next
 HexRed = Hex$(txtR.Text)
 If Len(HexRed) = 1 Then HexRed = "0" & HexRed
  HexGreen = Hex$(txtG.Text)
 If Len(HexGreen) = 1 Then HexGreen = "0" & HexGreen
  HexBlue = Hex$(txtB.Text)
 If Len(HexBlue) = 1 Then HexBlue = "0" & HexBlue
  txtF.Text = "#" & HexRed & HexGreen & HexBlue
End Function

' Open Dialogs
Private Sub cmdOpen_Click()
On Error Resume Next
 cdc1.ShowOpen
 txtBGImage.Text = cdc1.FileName
End Sub

Private Sub Command13_Click()
On Error Resume Next
 cdc1.ShowOpen
 txtSourceImg.Text = cdc1.FileName
End Sub




