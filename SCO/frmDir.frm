VERSION 5.00
Begin VB.Form frmDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directory"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   4320
      Begin VB.DirListBox dirMain 
         Appearance      =   0  'Flat
         Height          =   3015
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   4005
      End
      Begin VB.DriveListBox drvMain 
         Height          =   315
         Left            =   165
         TabIndex        =   1
         Top             =   255
         Width           =   4020
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Directory:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label lblDir 
      Caption         =   "Selected Directory"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4215
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TheDir As String

Private Sub cmdCancel_Click()
TheDir = ""
Unload Me
End Sub

Private Sub cmdCreate_Click()
frmDirCreate.CreateDir dirMain.Path
dirMain.Refresh

End Sub

Private Sub cmdOK_Click()
TheDir = dirMain.Path
Unload Me
End Sub

Private Sub dirMain_Change()
TheDir = dirMain.Path
lblDir = TheDir
End Sub

Private Sub drvMain_Change()
If dirMain.Path = drvMain.Drive Then
 dirMain.Path = drvMain.Drive
 Else
 MsgBox "Ops!Incorrect Entry", vbOKOnly, "Incorrect Entry"
 End If
End Sub

Private Sub Form_Load()
dirMain.Path = CurDir
TheDir = CurDir
lblDir = CurDir
End Sub

Public Function ShowDir(Optional Cap As String = "Select Folder") As String

    Me.Caption = Cap
    Me.Show vbModal
    ShowDir = TheDir
    
End Function
