VERSION 5.00
Begin VB.Form frmDirCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Directory"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmDirCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   330
      Left            =   1515
      TabIndex        =   1
      Top             =   300
      Width           =   2835
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   150
      Picture         =   "frmDirCreate.frx":000C
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmDirCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ParDir As String

Function CreateDir(CurDirn As String)
txt.Text = "New Folder"
Me.Caption = "Create In " & CurDirn
ParDir = CurDirn
Me.Show vbModal
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Handler
    If Right(ParDir, 1) <> "\" Then
     MkDir ParDir & "\" & txt.Text
    frmDir.dirMain.Path = ParDir & "\" & txt.Text
    Else
     MkDir ParDir & txt.Text
     frmDir.dirMain.Path = ParDir & txt.Text
    End If

    Unload Me
    Exit Sub

Handler:
    MsgBox "Unable to create folder. Please check that whether the name you entered is valid and contains no invalid characters..Try Again", vbInformation + vbOKOnly, "Cannot Create"
End Sub


Private Sub txt_Click()
txt.Text = ""
End Sub

Private Sub txt_GotFocus()
txt.SelStart = 0
txt.SelLength = Len(txt.Text)
End Sub
