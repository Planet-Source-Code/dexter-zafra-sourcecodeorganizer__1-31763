VERSION 5.00
Begin VB.Form frmsqltag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tag Chooser"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frmsqltag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List5 
      Height          =   4155
      Left            =   5280
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmsqltag.frx":030A
      Left            =   120
      List            =   "frmsqltag.frx":0425
      TabIndex        =   4
      ToolTipText     =   "Double click the item to add"
      Top             =   240
      Width           =   2895
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmsqltag.frx":07A9
      Left            =   120
      List            =   "frmsqltag.frx":0BBB
      TabIndex        =   3
      ToolTipText     =   "Double click the item to add"
      Top             =   2520
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmsqltag.frx":2671
      Left            =   3120
      List            =   "frmsqltag.frx":269C
      TabIndex        =   1
      ToolTipText     =   "Double click the item to add"
      Top             =   240
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   6000
      Pattern         =   "*.tag"
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Html:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Merkatum:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Opening,Closing & Comment:"
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   2040
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "User Defined Tags:"
      Height          =   195
      Left            =   5280
      TabIndex        =   6
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "frmsqltag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SName As String
Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim I As Integer
Dim T As String

File1.Refresh

List5.Refresh

File1.Path = App.Path

For I = 0 To File1.ListCount - 1

T = Left(File1.List(I), Len(File1.List(I)) - 4)

List5.AddItem (T)

Next I

End Sub

Private Sub List1_DblClick()

frmmain.RichTextBox1.SelText = frmmain.RichTextBox1.SelText & "<" & List1.List(List1.ListIndex) & ">" & "</" & List1.List(List1.ListIndex) & ">"

End Sub
Private Sub List3_DblClick()

frmmain.RichTextBox1.SelText = frmmain.RichTextBox1.SelText & List3.List(List3.ListIndex)

End Sub

Private Sub List4_DblClick()

frmmain.RichTextBox1.SelText = frmmain.RichTextBox1.SelText & List4.List(List4.ListIndex)

End Sub

Private Sub List5_DblClick()

frmmain.RichTextBox1.SelText = SName

End Sub

Private Sub List5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then


End If

End Sub



