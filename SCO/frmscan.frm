VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmscan 
   Caption         =   "Scanner V.1"
   ClientHeight    =   6480
   ClientLeft      =   5865
   ClientTop       =   2700
   ClientWidth     =   7920
   Icon            =   "frmscan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Box"
      Height          =   375
      Left            =   6120
      TabIndex        =   26
      ToolTipText     =   "Clear The box"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame frIn 
      Height          =   630
      Left            =   45
      TabIndex        =   18
      Top             =   1740
      Width           =   5835
      Begin MSComCtl2.UpDown udMain 
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         ToolTipText     =   "Go Up/Down"
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtMax"
         BuddyDispid     =   196611
         OrigLeft        =   3480
         OrigTop         =   225
         OrigRight       =   3720
         OrigBottom      =   540
         Max             =   2000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   2640
         TabIndex        =   19
         Text            =   "500"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Automatically stop after listing :"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3D5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox picSide 
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   5970
      ScaleHeight     =   1680
      ScaleWidth      =   1845
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   90
      Width           =   1845
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start Scan"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Scan"
         Top             =   0
         Width           =   1575
      End
      Begin MSComCtl2.Animation anMain 
         Height          =   915
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   570
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1614
         _Version        =   393216
         Center          =   -1  'True
         FullWidth       =   77
         FullHeight      =   61
      End
   End
   Begin VB.PictureBox picLarge 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   6000
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstFoundFiles 
      Height          =   2205
      Left            =   4920
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   2985
   End
   Begin VB.PictureBox picInvisible 
      Height          =   3480
      Left            =   1680
      ScaleHeight     =   3420
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   555
         ScaleHeight     =   2895
         ScaleWidth      =   3855
         TabIndex        =   9
         Top             =   2085
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   135
         ScaleHeight     =   2895
         ScaleWidth      =   3420
         TabIndex        =   10
         Top             =   285
         Width           =   3420
         Begin VB.FileListBox filList 
            Height          =   2040
            Left            =   120
            Pattern         =   "*.exe"
            TabIndex        =   13
            Top             =   480
            Width           =   1815
         End
         Begin VB.DirListBox dirList 
            Height          =   1665
            Left            =   2040
            TabIndex        =   12
            Top             =   960
            Width           =   1575
         End
         Begin VB.DriveListBox drvList 
            Height          =   315
            Left            =   2040
            TabIndex        =   11
            Top             =   480
            Width           =   1575
         End
      End
   End
   Begin MSComctlLib.StatusBar SbMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   6150
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2118
            MinWidth        =   2118
            Picture         =   "frmscan.frx":1CCA
            Object.ToolTipText     =   "Extreme Design 2001"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
      MouseIcon       =   "frmscan.frx":2F2E
   End
   Begin VB.PictureBox picSmall 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   6120
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   3195
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FilterIndex     =   1
   End
   Begin MSComctlLib.ImageList imgLarge 
      Left            =   2280
      Top             =   4035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   1545
      Top             =   4005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.Frame fr 
      Height          =   1740
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   5835
      Begin VB.CommandButton cmdDir 
         Height          =   375
         Left            =   5040
         Picture         =   "frmscan.frx":3248
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Browse "
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox txtLook 
         Height          =   315
         ItemData        =   "frmscan.frx":368A
         Left            =   1320
         List            =   "frmscan.frx":369A
         TabIndex        =   22
         Top             =   240
         Width           =   1000
      End
      Begin VB.CheckBox chkSub 
         Height          =   240
         Left            =   1050
         TabIndex        =   1
         Top             =   1320
         Width           =   285
      End
      Begin VB.TextBox txtFolder 
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Text            =   "C:\Program Files"
         Top             =   735
         Width           =   3600
      End
      Begin VB.Label Label3 
         Caption         =   "Include Subfolder"
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "From Folder:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Scan For:"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   915
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lvIcons 
      Height          =   3705
      Left            =   45
      TabIndex        =   21
      Top             =   2490
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6535
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgLarge"
      SmallIcons      =   "imgSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Path and Extension"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label lblIcons 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblIcon 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSSfolder 
         Caption         =   "&Set Start Folder"
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveSel 
         Caption         =   "Save Selected Files"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSaveAll 
         Caption         =   "Save All Files In List"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuLarge 
         Caption         =   "&Normal View"
      End
      Begin VB.Menu mnuSmall 
         Caption         =   "&List View"
      End
      Begin VB.Menu mnuVB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "&Details View"
      End
      Begin VB.Menu mnuVB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBitmap 
         Caption         =   "&Enhanced View"
         Shortcut        =   ^E
      End
      Begin VB.Menu kardex 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinvmedex 
         Caption         =   "Invert Selection"
      End
   End
   Begin VB.Menu mnuop 
      Caption         =   "&Options"
      Begin VB.Menu mnuShowLabel 
         Caption         =   "&Display Labels"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuOB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIcLabel 
         Caption         =   "&Icon Labels"
         Begin VB.Menu mnuIL1 
            Caption         =   "&File Path Name Extension"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuIL2 
            Caption         =   "&File Name Extension"
         End
      End
   End
   Begin VB.Menu mnuopz 
      Caption         =   "Optionz"
      Visible         =   0   'False
      Begin VB.Menu mnufilez4 
         Caption         =   "Enhanced View"
      End
      Begin VB.Menu mnusavedex 
         Caption         =   "Save Selected File"
      End
      Begin VB.Menu mnusavedexme 
         Caption         =   "Save All Files In List"
      End
      Begin VB.Menu dexter 
         Caption         =   "-"
      End
      Begin VB.Menu mnufilez2 
         Caption         =   "File Name Extension"
      End
      Begin VB.Menu mnufile1 
         Caption         =   "File path name extension"
      End
      Begin VB.Menu mnufilez3 
         Caption         =   "Invert Selection"
      End
      Begin VB.Menu zafire 
         Caption         =   "-"
      End
      Begin VB.Menu mnufilez007 
         Caption         =   "&Normal View"
      End
      Begin VB.Menu mnufilez6 
         Caption         =   "&Details View"
      End
      Begin VB.Menu mnufilez7 
         Caption         =   "&List View"
      End
   End
End
Attribute VB_Name = "frmscan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim glLargeIcons() As Long
Dim glSmallIcons() As Long
Dim lIcons         As Long
Dim sExeName       As String
Dim ExitFlag As Boolean

Dim SearchFlag As Boolean

Const LARGE_ICON As Integer = 32
Const SMALL_ICON As Integer = 16
Const DI_NORMAL = 3

Private Declare Function DrawIconEx Lib "user32" _
    (ByVal hDc As Long, ByVal XLEFT As Long, ByVal yTop As Long, _
    ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, _
    ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function ExtractIconEx Lib "shell32.dll" _
    Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, _
    phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
    
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long

Private Sub AddIcons()

Dim FirstPath As String, DirCount As Integer

Dim I
lvIcons.ListItems.Clear

picSmall.Picture = LoadPicture("")
picLarge.Picture = LoadPicture("")

Set lvIcons.Icons = Nothing
Set lvIcons.SmallIcons = Nothing

With imgLarge
    .ListImages.Clear
    .ImageHeight = LARGE_ICON
    .ImageWidth = LARGE_ICON
End With

With imgSmall
    .ListImages.Clear
    .ImageHeight = SMALL_ICON
    .ImageWidth = SMALL_ICON
End With

SbMain.Panels(2).Text = "Please Wait...Searching " & txtFolder.Text
Label3D5.Caption = "Wait..Searching " & txtFolder.Text

    FirstPath = dirList.Path
    DirCount = dirList.ListCount
    NumFiles = 0
    lstFoundFiles.Clear
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path

Set lvIcons.Icons = imgLarge
Set lvIcons.SmallIcons = imgLarge

For I = 1 To imgLarge.ListImages.Count
SbMain.Panels(2).Text = "Scanning" & I & "th icon to list"
Label3D5.Caption = "Scanning" & I & "th icon to list"
Dim ic As ListItem
    Set ic = lvIcons.ListItems.Add(, , "", I, I)
    ic.SubItems(1) = FPathFromKey(imgLarge.ListImages(I).Key)
Next I

ToggleCaption mnuShowLabel.Checked

SbMain.Panels(2).Text = lvIcons.ListItems.Count & " - Files Found-Scanning Done..."
Label3D5.Caption = lvIcons.ListItems.Count & "Files Found..."

End Sub

Private Sub cmdDir_Click()
txtFolder.Text = frmDir.ShowDir()

On Error Resume Next
dirList.Path = txtFolder.Text

If Trim(txtFolder.Text) = "" Then
    txtFolder.Text = dirList.Path
End If

End Sub

Private Sub cmdDir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3D5.Caption = "Browse for files"
End Sub

Private Sub cmdStart_Click()
On Error Resume Next

If cmdstart.Caption = "&Stop Scan" Then
    ExitFlag = True
    SearchFlag = False
    cmdstart.Caption = "&Start Scan"
    anMain.Stop
    mnuStart.Caption = cmdstart.Caption
    Exit Sub
End If

If cmdstart.Caption = "&Start Scan" Then
    ExitFlag = False
    cmdstart.Caption = "&Stop Scan"
    mnuStart.Caption = cmdstart.Caption
    lblCount.Caption = "0"
    SearchFlag = True
    anMain.Play
    txtFolder_Change
    AddIcons
    mnuStart.Caption = cmdstart.Caption
    txtFolder_Change
    EnableDisable
    cmdstart.Caption = "&Start Scan"
    anMain.Stop
   SbMain.Panels(2).Text = lvIcons.ListItems.Count & " - Files Found-Scanning Done..."
    Label3D5.Caption = lvIcons.ListItems.Count & "Files Found..."
    
    Exit Sub
End If

End Sub

Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3D5.Caption = " Scan Files"
End Sub

Private Sub Command1_Click()

Dim V As Integer
V = MsgBox("Are you sure you that want to remove all files in the list box?", vbYesNo Or vbQuestion, "Organizer Scanner") = vbYes
If V = True Then
 lvIcons.ListItems.Clear
 
   SbMain.Panels(2).Text = "Files Has Been Remove...."
End If

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3D5.Caption = "Remove Files from list"
End Sub
 Private Sub dirList_Change()
filList.Path = dirList.Path
End Sub

Private Sub drvList_Change()
dirList.Path = drvList.Drive
End Sub

Private Sub Form_Load()

ExitFlag = False

picLarge.Height = LARGE_ICON * Screen.TwipsPerPixelY
picLarge.Width = LARGE_ICON * Screen.TwipsPerPixelX
picSmall.Height = SMALL_ICON * Screen.TwipsPerPixelY
picSmall.Width = SMALL_ICON * Screen.TwipsPerPixelX
anMain.Open App.Path & "\Organizerpreview\rem.avi"

dirList.Path = drvList.Drive
filList.Path = dirList.Path
txtFolder.Text = dirList.Path
Form_Resize
CodeUnload = False

LoadPrefs
EnableDisable

End Sub

Public Sub pGetIcons(sExeName As String)
Dim l As Long

lIcons = ExtractIconEx(sExeName, -1, 0, 0, 0)

 If lIcons < 0 Then Exit Sub

ReDim glLargeIcons(lIcons)
ReDim glSmallIcons(lIcons)

Dim lIndex

For lIndex = 0 To lIcons - 1

Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)

With picLarge
    Set .Picture = LoadPicture("")
     .AutoRedraw = True
    Call DrawIconEx(.hDc, 0, 0, glLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
     .Refresh
End With

On Error GoTo stopThis


mykey = sExeName & "(" & lIndex & ")"

If Val(txtMax.Text) = imgLarge.ListImages.Count Then
SearchFlag = False
Else
imgLarge.ListImages.Add , mykey, picLarge.Image
End If

With picSmall
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    Call DrawIconEx(.hDc, 0, 0, glSmallIcons(lIndex), SMALL_ICON, SMALL_ICON, 0, 0, DI_NORMAL)
    .Refresh
End With

nextIcon:
Next lIndex

Exit Sub

stopThis:

txtMax.Text = imgLarge.ListImages.Count

End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer

Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String

Dim RetVal As Integer

    If ExitFlag = True Then
        DirDiver = True
        Exit Function
    End If
    
    
    SearchFlag = True
    DirDiver = False
   
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    
    On Local Error GoTo DirDriverHandler
    
    If chkSub.Value = 1 Then
        DirsToPeek = dirList.ListCount
    Else
        DirsToPeek = 0
    End If
    
    Do While DirsToPeek > 0 And SearchFlag = True
    
        OldPath = dirList.Path
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            
            dirList.Path = dirList.List(DirsToPeek - 1)
            RetVal = DoEvents()
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
     
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then
            ThePath = dirList.Path
        Else
            ThePath = dirList.Path + "\"
        End If
        
        For ind = 0 To filList.ListCount - 1
            entry = ThePath + filList.List(ind)
          
            pGetIcons entry
            SbMain.Panels(2).Text = "Scanning from " & entry
            lblCount.Caption = Str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    If BackUp <> "" Then
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If Err = 7 Then
        DirDiver = True
        MsgBox "You've filled the list box.Stop Scan!"
        Exit Function
    Else
error:
message = MsgBox("Error : " & Err.Number & " : " & Err.Description, vbCritical, "Error")
        End
    End If
    
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3D5.Caption = ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
lvIcons.Move 0, lvIcons.Top, Me.ScaleWidth, (Me.ScaleHeight - lvIcons.Top - SbMain.Height)

lvIcons.ColumnHeaders(2).Width = lvIcons.Width - lvIcons.ColumnHeaders(1).Width * -80

fr.Width = Me.ScaleWidth - (2 * fr.Left + picSide.Width)
frIn.Width = fr.Width
picSide.Left = fr.Width + fr.Left + 100
Command1.Left = fr.Width + fr.Left + 100
If cmdDir.Left < chkSub.Left + chkSub.Width Then
    cmdDir.Visible = True
    Else
    cmdDir.Visible = True
End If

End Sub

Private Sub Form_Unload(cancel As Integer)

On Error Resume Next

SavePrefs

Set lvIcons.Icons = Nothing

lvIcons.ListItems.Clear

imgLarge.ListImages.Clear
imgSmall.ListImages.Clear

Unload Me

End Sub

Private Sub fr_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3D5.Caption = ""
End Sub

Private Sub lvIcons_DblClick()
If lvIcons.ListItems.Count > 0 Then
frmBit.Show
frmBit.img.Picture = imgLarge.ListImages(lvIcons.SelectedItem.Icon).ExtractIcon
frmBit.SetStatus lvIcons.SelectedItem.Icon
End If
End Sub

Private Sub lvIcons_ItemClick(ByVal Item As MSComctlLib.ListItem)
SbMain.Panels(2).Text = Item.Text & ": This files are from " & FPathFromKey(imgLarge.ListImages(Item.Icon).Key)

Me.PopupMenu mnuopz
If BitLoaded Then
    frmBit.I = lvIcons.SelectedItem.Icon
    frmBit.img.Picture = imgLarge.ListImages(Item.Icon).ExtractIcon
End If
End Sub

Private Sub lvIcons_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3D5.Caption = ""
End Sub

Private Sub mnuBitmap_Click()
lvIcons_DblClick
End Sub

Private Sub mnuDetails_Click()
lvIcons.View = lvwReport
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnufile1_Click()
mnuIL2.Checked = Not mnuIL2.Checked
mnuIL1.Checked = Not mnuIL1.Checked
ToggleCaption mnuShowLabel.Checked
End Sub

Private Sub mnufilez007_Click()
lvIcons.View = lvwIcon
End Sub

Private Sub mnufilez2_Click()
mnuIL1_Click
End Sub

Private Sub mnufilez3_Click()
On Error Resume Next
For I = 1 To lvIcons.ListItems.Count
    lvIcons.ListItems(I).Selected = Not lvIcons.ListItems(I).Selected
Next I
End Sub

Private Sub mnufilez4_Click()
lvIcons_DblClick
End Sub

Private Sub mnufilez6_Click()
lvIcons.View = lvwReport
End Sub

Private Sub mnufilez7_Click()
lvIcons.View = lvwList
End Sub

Private Sub mnuIL1_Click()
mnuIL2.Checked = Not mnuIL2.Checked
mnuIL1.Checked = Not mnuIL1.Checked
ToggleCaption mnuShowLabel.Checked
End Sub

Private Sub mnuIL2_Click()
mnuIL1_Click
End Sub

Private Sub mnuinvmedex_Click()
On Error Resume Next
For I = 1 To lvIcons.ListItems.Count
    lvIcons.ListItems(I).Selected = Not lvIcons.ListItems(I).Selected
Next I

End Sub

Private Sub mnuLarge_Click()
lvIcons.View = lvwIcon
End Sub

Private Sub mnuSaveAll_Click()
SaveAllIcons
End Sub

Private Sub mnusavedex_Click()
SaveSelectedIcons
End Sub

Private Sub mnusavedexme_Click()
SaveAllIcons
End Sub

Private Sub mnuSaveSel_Click()
SaveSelectedIcons
End Sub

Private Sub mnuShowLabel_Click()
mnuShowLabel.Checked = Not mnuShowLabel.Checked
ToggleCaption mnuShowLabel.Checked
End Sub

Private Sub mnuSmall_Click()
lvIcons.View = lvwList
End Sub

Private Sub mnuSSfolder_Click()
cmdDir_Click
End Sub

Private Sub mnuStart_Click()
cmdStart_Click
End Sub

Private Sub SbMain_Click()
If SbMain.Panels(1).Enabled = True Then
Call Shell("Start.exe " & "http://clik.to/ret", 0)

End If
End Sub

Private Sub txtFolder_Change()
On Error Resume Next
dirList.Path = txtFolder.Text
End Sub

Private Sub txtFolder_Click()
txtFolder.Text = ""
End Sub

Private Sub txtFolder_GotFocus()
txtFolder.SelStart = 0
txtFolder.SelLength = Len(txtFolder.Text)
End Sub

Private Sub txtFolder_Validate(cancel As Boolean)
On Error GoTo Handle
dirList.Path = txtFolder.Text
Exit Sub

Handle:

mstr = "Ops!The folder you entered is not valid. Please enter a valid folder, or click the button next to it for selecting a folder"

MsgBox mstr, vbOKOnly + vbInformation, "Invalid Directory"
txtFolder.Text = CurDir
dirList.Path = CurDir
cancel = True
End Sub

Private Sub txtLook_Change()
On Error Resume Next
filList.Pattern = txtLook.Text
End Sub

Private Sub txtLook_GotFocus()
txtLook.SelStart = 0
txtLook.SelLength = Len(txtLook.Text)
End Sub

Private Sub txtLook_Validate(cancel As Boolean)
On Error GoTo Handle
filList.Pattern = txtLook.Text
Exit Sub

Handle:
Dim Str As String
mstr = "Ops!The file pattern you entered is invalid. Please enter a valid criteria."
mstr = mstr + vbCrLf + vbCrLf
mstr = mstr + "Examples:" + vbCrLf + vbCrLf + vbCrLf
mstr = mstr + "   1) *.exe       - Searches in all EXE files" + vbCrLf + vbCrLf + vbCrLf
mstr = mstr + "   2) *.exe;*.dll;*.ico;*.cur - Searches in all EXE files ,DLL files,icon and cursor"


MsgBox mstr, vbOKOnly + vbInformation, "Invalid File Type"

filList.Pattern = "*.exe"

End Sub

Sub SaveAllIcons()

    Dim Getpath As String
    Getpath = frmDir.ShowDir("Store To Directory..")
    If Getpath = "" Then Exit Sub
    If Right(Getpath, 1) <> "\" Then Getpath = Getpath & "\"
    Cap = SbMain.Panels(2).Text
    SbMain.Panels(2).Text = "Saving Files.."
    Label3D5.Caption = "Saving Files.."
    For I = 1 To lvIcons.ListItems.Count
    On Error Resume Next
        SavePicture imgLarge.ListImages(lvIcons.ListItems(I).Icon).ExtractIcon, Getpath & "\Icon " & I & ".ico"
    Next I
    SbMain.SimpleText = Cap
End Sub

Sub SaveSelectedIcons()

    Dim Getpath As String
    Getpath = frmDir.ShowDir("Store To Directory..")
    If Getpath = "" Then Exit Sub
    If Right(Getpath, 1) <> "\" Then Getpath = Getpath & "\"
    Cap = SbMain.Panels(2).Text
    SbMain.SimpleText = "Saving Files.."
    Label3D5.Caption = "Saving File.."
    For I = 1 To lvIcons.ListItems.Count
    On Error Resume Next
    If lvIcons.ListItems(I).Selected = True Then _
        SavePicture imgLarge.ListImages(lvIcons.ListItems(I).Icon).ExtractIcon, Getpath & "\Icon " & I & ".ico"
    Next I
    SbMain.Panels(2).Text = Cap
End Sub

Function FNameFromPath(FullFile As String) As String

Dim LastPos
LastPos = -1

For I = 1 To Len(FullFile)
    If Right(VBA.Left(FullFile, I), 1) = "\" Then
        LastPos = I
    End If
Next I
        
If LastPos > 0 Then
        FNameFromPath = Right(FullFile, Len(FullFile) - LastPos)
        Exit Function
End If

End Function

Public Function FPathFromKey(FullFile As String) As String

Dim LastPos
LastPos = -1

For I = 1 To Len(FullFile)
    If Right(VBA.Left(FullFile, I), 1) = "(" Then
        LastPos = I
    End If
Next I
        
If LastPos > 0 Then
        FPathFromKey = Left(FullFile, LastPos - 1)
        Exit Function
End If

End Function

Function ToggleCaption(TState As Boolean)

On Error Resume Next

cucap = SbMain.Panels(2).Text

    If TState = True Then
        For I = 1 To imgLarge.ListImages.Count
            SbMain.Panels(2).Text = "Wait,Setting Captions.."
            If mnuIL1.Checked = True Then
            lvIcons.ListItems(I).Text = "File" & I
            Else
            lvIcons.ListItems(I).Text = FNameFromPath(imgLarge.ListImages(I).Key)
            End If
            
         Next I
    Else
        For I = 1 To imgLarge.ListImages.Count
            SbMain.Panels(2).Text = "Wait,Removing Captions.."
            lvIcons.ListItems(I).Text = ""
         Next I
        
    End If
    
SbMain.SimpleText = cucap
End Function

Sub SavePrefs()
SaveSetting App.EXEName, "Options", "View", lvIcons.View
SaveSetting App.EXEName, "Options", "Lookin", txtLook.Text
SaveSetting App.EXEName, "Options", "Folder", txtFolder.Text
SaveSetting App.EXEName, "Options", "Maximum", txtMax.Text
SaveSetting App.EXEName, "Options", "Sub", chkSub.Value

End Sub

Sub LoadPrefs()
On Error Resume Next

lvIcons.View = GetSetting(App.EXEName, "Options", "View", 0)

txtLook.Text = GetSetting(App.EXEName, "Options", "Lookin", "*.exe")
filList.Pattern = txtLook.Text

txtFolder.Text = GetSetting(App.EXEName, "Options", "SearchFolder", WinDir())
dirList.Path = txtFolder.Text

txtMax.Text = GetSetting(App.EXEName, "Options", "Maximum", "300")

chkSub.Value = GetSetting(App.EXEName, "Options", "Sub", 1)

End Sub

Sub EnableDisable()
If lvIcons.ListItems.Count < 1 Then
    mnuSaveAll.Enabled = False
    mnuSaveSel.Enabled = False
Else
    mnuSaveAll.Enabled = True
    mnuSaveSel.Enabled = True
   
End If
End Sub

Private Sub txtMax_Click()
txtMax.Text = ""
End Sub

Private Sub txtMax_Validate(cancel As Boolean)
If Not IsNumeric(txtMax.Text) Then
MsgBox "Ops!The value you entered for maximum files is invalid. Please enter a valid value", vbInformation + vbOKOnly, "Invalid Entry"
    cancel = 1
    txtMax.Text = udMain.Value
End If
End Sub

Function WinDir() As String
Dim WinPath As String
    WinPath = String(145, Chr(0))
    WinDir = Left(WinPath, GetWindowsDirectory(WinPath, 145))
End Function

