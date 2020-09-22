VERSION 5.00
Begin VB.Form DateTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Date - Time"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   HelpContextID   =   370
   Icon            =   "Date-Time.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formats"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      Begin VB.ListBox Lsttime 
         Height          =   2790
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "DateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmmain.rtbnet.SelText = Lsttime.Text
Unload Me
End Sub

Private Sub Command2_Click()
DateTime.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Lsttime.AddItem Format(Now, "long time")
Lsttime.AddItem Format(Now, "short time")
Lsttime.AddItem Format(Now, "medium time")
Lsttime.AddItem Format(Now, "general date")
Lsttime.AddItem Format(Now, "long date")
Lsttime.AddItem Format(Now, "medium date")
Lsttime.AddItem Format(Now, "short date")
Lsttime.AddItem (Date)
Lsttime.AddItem Format(Date, "dd - mm - yyyy")
Lsttime.AddItem Format(Date, "dd-mm-yy")
Lsttime.AddItem Format(Date, "dd/mm/yy")
Lsttime.AddItem Format(Date, "dd/mm/yyyy")
Lsttime.AddItem Format(Date, "dd/mm")
Lsttime.AddItem Format(Date, "dd")
Lsttime.AddItem Format(Time, "hh-mm-ss")
Lsttime.AddItem Format(Time, "hh.mm.ss")
Lsttime.AddItem Format(Time, "hh-mm")
End Sub

Private Sub Lsttime_DblClick()
Command1_Click
End Sub
