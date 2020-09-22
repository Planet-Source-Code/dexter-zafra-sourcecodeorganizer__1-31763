VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmfinddb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Base Viewer"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4560
   Icon            =   "frmfinddb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Directory List"
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Tag             =   "mov"
         Top             =   240
         Width           =   4095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2520
         Left            =   120
         TabIndex        =   5
         Tag             =   "mov"
         ToolTipText     =   "Double Click or Right Click  the file  to open."
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4445
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         Icons           =   "ImageList4"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date Created"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   7800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":1296
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":16EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":1782
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":1BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":1CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":1EC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":209E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":24F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":257A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2954
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2A1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2B7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":2FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":3080
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":310C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":3224
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmfinddb.frx":3328
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   8280
      TabIndex        =   1
      Tag             =   "mov"
      Top             =   600
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   8160
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image picMenu8 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "frmfinddb.frx":3448
      Top             =   5040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu7 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "frmfinddb.frx":34BC
      Top             =   4680
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu6 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "frmfinddb.frx":352D
      Top             =   4440
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu5 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "frmfinddb.frx":3592
      Top             =   4080
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu4 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "frmfinddb.frx":35FF
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu3 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   9480
      Picture         =   "frmfinddb.frx":368D
      Top             =   3600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu2 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   8640
      Picture         =   "frmfinddb.frx":36DE
      Top             =   2880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   3
      Left            =   8280
      Picture         =   "frmfinddb.frx":3778
      Top             =   2880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   2
      Left            =   9480
      Picture         =   "frmfinddb.frx":37F4
      Top             =   2760
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   1
      Left            =   9480
      Picture         =   "frmfinddb.frx":3883
      Top             =   2520
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image picMenu 
      Appearance      =   0  'Flat
      Height          =   195
      Index           =   0
      Left            =   7920
      Picture         =   "frmfinddb.frx":38FA
      Top             =   1560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7920
      TabIndex        =   3
      Tag             =   "mov"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Caption         =   " Local:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   8040
      TabIndex        =   2
      Tag             =   "mov"
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Menu zLocal 
      Caption         =   "&File"
      Begin VB.Menu zOpenFile 
         Caption         =   "&Open DataBase"
      End
      Begin VB.Menu zSep6 
         Caption         =   "-"
      End
      Begin VB.Menu zLokNF 
         Caption         =   "&Create New Folder"
      End
      Begin VB.Menu zLokDS 
         Caption         =   "&Delete File"
      End
      Begin VB.Menu zFind 
         Caption         =   "&Advance File Search"
         Shortcut        =   ^F
      End
      Begin VB.Menu zProperties 
         Caption         =   "&View File Properties"
      End
      Begin VB.Menu zSep3 
         Caption         =   "&View Window Settings..."
         Begin VB.Menu ztBigIc 
            Caption         =   "&Big Icons"
         End
         Begin VB.Menu ztSmallIc 
            Caption         =   "&Small Icons"
         End
         Begin VB.Menu ztSeznam 
            Caption         =   "&List"
         End
         Begin VB.Menu ztReport 
            Caption         =   "&Report"
         End
      End
   End
End
Attribute VB_Name = "frmfinddb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const SW_SHOWNORMAL = 1
Private Const SW_SHOW = 5
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SHFindFiles Lib "shell32.dll" Alias "#90" (ByVal pidlRoot As Long, ByVal pidlSavedSearches As Long) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (ByRef s As SHELLEXECUTEINFO) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private lsDrag() As ListItem
Private fPat As String
Dim iPos As Integer
Dim strExt As String
Dim tvNode As Node
Dim lsItem As ListItem

Public Sub List()
Dim strFile As String
Dim img As Integer, r As Integer
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Refresh
LoadLocal
End Sub

Private Sub Dir1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
End If
End Sub

Private Sub Drive1_Change()
On Error GoTo Hell
Dir1.Path = Drive1.Drive
Dir1.Refresh
 Exit Sub
Hell:
 MsgBox "Invalid Drive", vbCritical, "Error"
End Sub

Private Sub Form_Load()
File1.Pattern = "*.mdb"
LoadLocal
ListView1.View = 0
Dim hMenu As Long, hSubMenu As Long
Dim RetVal As Long
Dim I As Long
hMenu = GetMenu(Me.hwnd)
hSubMenu = GetSubMenu(hMenu, 0)
For I = 0 To 3
    RetVal = SetMenuItemBitmaps(hSubMenu, I, MF_BYPOSITION, picMenu(I).Picture, picMenu(I).Picture)
Next I
RetVal = SetMenuItemBitmaps(hSubMenu, 5, MF_BYPOSITION, picMenu2.Picture, picMenu2.Picture)
hSubMenu = GetSubMenu(hMenu, 1)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu3.Picture, picMenu3.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, picMenu(0).Picture, picMenu(0).Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 3, MF_BYPOSITION, picMenu4.Picture, picMenu4.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 11, MF_BYPOSITION, picMenu6.Picture, picMenu6.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 12, MF_BYPOSITION, picMenu5.Picture, picMenu5.Picture)
hSubMenu = GetSubMenu(hMenu, 2)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu3.Picture, picMenu3.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 2, MF_BYPOSITION, picMenu4.Picture, picMenu4.Picture)
hSubMenu = GetSubMenu(hMenu, 3)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu7.Picture, picMenu7.Picture)
RetVal = SetMenuItemBitmaps(hSubMenu, 1, MF_BYPOSITION, picMenu(2).Picture, picMenu(2).Picture)
hSubMenu = GetSubMenu(hMenu, 4)
RetVal = SetMenuItemBitmaps(hSubMenu, 0, MF_BYPOSITION, picMenu8.Picture, picMenu8.Picture)
fPat = "*.*"
Dir1.Path = "\"
LoadLocal
End Sub
Private Sub LoadLocal()
Dim x As Integer, img As Integer
Dim y As Long
Drive1.Refresh
Dir1.Refresh
File1.Refresh
ListView1.ListItems.Clear
If Mid(Dir1.Path, Len(Dir1.Path), 1) = "\" Then
       strPath = Dir1.Path
 Else: strPath = Dir1.Path & "\"
End If

For x = 0 To File1.ListCount - 1
 img = ImgNumber(File1.List(x))
 With ListView1.ListItems.Add(, , File1.List(x), img, img)
   .SubItems(1) = Format((FileLen(strPath & File1.List(x)) / 1000), "### ### ###.##") & " Kb"
   .SubItems(2) = FileDateTime(strPath & File1.List(x))
   y = Str(FileLen(strPath & File1.List(x)))
   .SubItems(3) = Str(FileLen(strPath & File1.List(x)))
End With
Next
ListView1.SelectedItem = Nothing

End Sub

Private Sub ListView1_AfterLabelEdit(cancel As Integer, NewString As String)
Dim strEx2 As String, strEx1 As String
Dim Msg As VbMsgBoxResult
On Error GoTo Err
strEx1 = Mid$(ListView1.SelectedItem.Text, InStrRev(ListView1.SelectedItem.Text, ".") + 1)
strEx2 = Mid$(NewString, InStrRev(NewString, ".") + 1)
If strEx1 <> strEx2 Then
    Msg = MsgBox("Are you sure to exchange the file extension from: " & Chr(34) & strEx1 & Chr(34) & " to: " & Chr(34) & strEx2 & Chr(34), vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        cancel = 0
        Name strPath & ListView1.SelectedItem.Text As strPath & NewString
  
    Else: cancel = 1
    End If
Else
    cancel = 0
    Name strPath & ListView1.SelectedItem.Text As strPath & NewString

End If
Err: If Err.Number = 58 Then
MsgBox "More than one file with the same name in one folder? no way!", vbExclamation, App.Title
cancel = 1
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
If ListView1.SortOrder = 0 Then
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = 1
 Else
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = 0
End If
 ListView1.Sorted = True
End Sub

Private Sub ListView1_Click()
Dim I, x As Integer
Dim y, z As Long
x = 0
z = 0
If ListView1.SelectedItem Is Nothing Then Exit Sub
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Selected = True Then
           y = ListView1.ListItems(I).SubItems(3)
           z = z + y
           x = x + 1
        End If
    Next I

zProperties.Enabled = True
End Sub

Private Sub ListView1_DblClick()
zOpenFile_Click
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    zLokDS_Click
ElseIf KeyCode = vbKeyReturn Then
    zOpenFile_Click
End If
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Me.PopupMenu zLocal
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    ListView1.Height = frmmain.ScaleHeight
    Dir1.Height = Dir1.Top
End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    Label3.Width = frmmain.ScaleWidth - Label3.left
End If
End Sub
Private Sub zFind_Click()
SHFindFiles 0, 0
End Sub

Private Sub zLokDS_Click()
Dim Msg As VbMsgBoxResult
Dim I As Integer
If ListView1.SelectedItem Is Nothing Then
    MsgBox "Nothing to delete..You must select a file!", vbExclamation
    Exit Sub
Else
    Msg = MsgBox("Are you sure want to delete this files?" & vbCrLf & "If you are, don't look for them in recycle bin!", vbQuestion + vbYesNo, App.Title)
    If Msg = vbYes Then
        For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Selected = True Then
            Kill strPath & ListView1.ListItems(I).Text
        End If
        Next I
    LoadLocal
    End If
End If
End Sub

Private Sub zLokNF_Click()
On Error GoTo Err
    Dim sRet As String
    sRet = InputBox("Type a name of the new folder here", "Folder Name")
    If sRet <> "" Then
        MkDir strPath & sRet
        Dir1.Refresh
    End If
Err: If Err.Number = 75 Then MsgBox "an error apeared while creating new folder" & vbCrLf & "Make sure folder doesn't exist!", vbExclamation, App.Title
Exit Sub
End Sub
Private Sub zOpenFile_Click()
If Not ListView1.SelectedItem Is Nothing Then
ShellExecute 0, vbNullString, strPath & ListView1.SelectedItem.Text, vbNullString, strPath, SW_SHOWNORMAL
Else: MsgBox "You must select a file below in the view window, then open!", vbExclamation, App.Title
End If
End Sub
Private Function ImgNumber(strFilename As String) As Integer
Dim strExt As String
    strExt = Mid$(strFilename, InStrRev(strFilename, ".") + 1)
    On Error Resume Next
    Select Case LCase(strExt)
       Case "avi", "mpg", "mpeg", "mov"
            ImgNumber = 8
       Case "gif"
            ImgNumber = 4
       Case "jpg", "jpeg", "jpe", "bmp"
            ImgNumber = 1
       Case "htm", "html", "xml", "asp"
            ImgNumber = 2
       Case "js", "css", "cgi"
            ImgNumber = 5
       Case "mp3", "ram", "au", "vaw"
            ImgNumber = 6
       Case "zip", "arj"
            ImgNumber = 7
       Case "exe", "com", "bat"
           ImgNumber = 9
       Case "txt", "log", "doc", "rtf", "ftp", "ini", "dat"
           ImgNumber = 3
       Case Else
            ImgNumber = 10
    End Select
End Function
Private Sub zProperties_Click()
Dim shInfo As SHELLEXECUTEINFO
If ListView1.SelectedItem Is Nothing Then
    MsgBox "Nothing to view..You must select a file"
    Exit Sub
End If
Set lsItem = ListView1.SelectedItem
    With shInfo
        .cbSize = LenB(shInfo)
        .lpFile = strPath & lsItem.Text
        .nShow = SW_SHOW
        .fMask = SEE_MASK_INVOKEIDLIST
        .lpVerb = "properties"
    End With
    ShellExecuteEx shInfo
End Sub

Private Sub ztBigIc_Click()
ListView1.View = 0
End Sub

Private Sub ztReport_Click()
ListView1.View = 3
End Sub

Private Sub ztSeznam_Click()
ListView1.View = 2
End Sub

Private Sub ztSmallIc_Click()
ListView1.View = 1
End Sub
