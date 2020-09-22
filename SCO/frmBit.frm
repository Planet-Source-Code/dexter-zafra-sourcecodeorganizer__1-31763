VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBit 
   Caption         =   "Enhance View.."
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5220
   Icon            =   "frmBit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   1815
      Left            =   870
      ScaleHeight     =   1755
      ScaleWidth      =   1845
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1905
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image img 
      BorderStyle     =   1  'Fixed Single
      Height          =   3090
      Left            =   30
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4860
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuNext 
         Caption         =   "&Next Icon                   "
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "&Previous Icon            "
      End
      Begin VB.Menu mnuVB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFirst 
         Caption         =   "&First Icon"
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last Icon"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuKeep 
         Caption         =   "&Keep On Top"
         Checked         =   -1  'True
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "frmBit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public I As Integer

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
    mnuNext_Click
End If

If KeyCode = vbKeyBack Then
    mnuPrev_Click
End If

End Sub

Private Sub Form_Load()
Me.Height = GetSetting(App.EXEName, "Bitmapview", "Height", 200)
Me.Width = GetSetting(App.EXEName, "Bitmapview", "Width", 200)
Form_Resize
ToggleState
End Sub

Private Sub Form_Paint()
ToggleState
End Sub

Private Sub Form_Resize()
On Error Resume Next

img.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - SbMain.Height
Me.Caption = "Image Size-[Height=" & img.Height & "-Width=" & img.Width & "]"
End Sub

Private Sub Form_Unload(cancel As Integer)
SaveSetting App.EXEName, "Bitmapview", "Height", Me.Height
SaveSetting App.EXEName, "Bitmapview", "Width", Me.Width
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFirst_Click()
I = 1
img.Picture = frmscan.imgLarge.ListImages(I).ExtractIcon
SetStatus I
End Sub
Private Sub mnuKeep_Click()
mnuKeep.Checked = Not mnuKeep.Checked
ToggleState
End Sub

Sub ToggleState()
        If mnuKeep.Checked Then
           SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
        Else
           SetWindowPos Me.hwnd, HWND_NOTOPMOST, Me.Left / 15, _
                        Me.Top / 15, Me.Width / 15, _
                        Me.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
        End If

End Sub

Private Sub mnuLast_Click()
I = frmscan.imgLarge.ListImages.Count
img.Picture = frmscan.imgLarge.ListImages(I).ExtractIcon
SetStatus I

End Sub

Private Sub mnuNext_Click()

On Error Resume Next
If I = frmscan.imgLarge.ListImages.Count Then I = 1

I = I + 1
img.Picture = frmscan.imgLarge.ListImages(I).ExtractIcon
SetStatus I

End Sub

Private Sub mnuPrev_Click()
On Error Resume Next

If I = 1 Then I = frmscan.imgLarge.ListImages.Count

I = I - 1
img.Picture = frmscan.imgLarge.ListImages(I).ExtractIcon
SetStatus I

End Sub

Private Sub mnuSave_Click()

Dim clOpen As CommonDialog

Set clOpen = frmscan.cdlOpen

On Error GoTo nosave
clOpen.CancelError = True
clOpen.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
clOpen.DialogTitle = "Save Icon"
clOpen.Filter = "Icon File|*.ico"
clOpen.DefaultExt = "ico"
clOpen.ShowSave

SavePicture img.Picture, clOpen.FileTitle

nosave:

End Sub

Public Sub SetStatus(cIndex As Integer)
SbMain.SimpleText = ""
    
    SbMain.SimpleText = frmscan.lvIcons.ListItems(cIndex).Text
    If SbMain.SimpleText <> "" Then SbMain.SimpleText = SbMain.SimpleText & " - "
    SbMain.SimpleText = SbMain.SimpleText & frmscan.FPathFromKey(frmscan.imgLarge.ListImages(cIndex).Key)
    
End Sub
