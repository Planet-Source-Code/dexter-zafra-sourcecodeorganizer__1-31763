VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   3165
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   3225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   3135
      ScaleWidth      =   3165
      TabIndex        =   0
      Top             =   0
      Width           =   3195
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   720
         TabIndex        =   1
         Top             =   3000
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading...."
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2760
         Width           =   1695
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   7440
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   7440
      Top             =   2400
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub
Private Sub Form_Load()
    ProgressBar1.Value = ProgressBar1.Min
End Sub
Private Sub Frame1_Click()
    Unload Me
End Sub
Private Sub Timer1_Timer()
    Unload Me
    frmmain.Show
End Sub
Private Sub Timer2_Timer()
ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value = 50 Then
ProgressBar1.Value = ProgressBar1 + 50
If ProgressBar1.Value >= ProgressBar1.Max Then
Timer2.Enabled = False
End If
End If
End Sub





