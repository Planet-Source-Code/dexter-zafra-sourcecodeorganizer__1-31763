VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form note 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Organizer Untitled - Editor"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9495
   Icon            =   "Note.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.txt"
      Filter          =   "All Files (*.*)|*.*|Text Files"" & ""(*.txt)|*.txt|"
      FilterIndex     =   2
      Flags           =   10
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontStrikeThru  =   -1  'True
      FontUnderLine   =   -1  'True
      MaxFileSize     =   1000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":0442
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":0556
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":066A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":0782
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":089A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":09B2
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":0ACA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":0BE2
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Note.frx":0CFA
            Key             =   "Cut"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Rich 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   11880
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      MaxLength       =   20000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Note.frx":0E0E
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
   Begin VB.Menu file 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "Export to Main Code Window"
         Shortcut        =   ^S
      End
      Begin VB.Menu sav 
         Caption         =   "Save As..."
      End
      Begin VB.Menu print 
         Caption         =   "Print......."
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      WindowList      =   -1  'True
      Begin VB.Menu cut 
         Caption         =   "Cut"
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
      End
      Begin VB.Menu Dt 
         Caption         =   "Time/Date"
         Shortcut        =   ^T
      End
      Begin VB.Menu d 
         Caption         =   "-"
      End
      Begin VB.Menu mnuusesounds 
         Caption         =   "#Play Typing Sound"
      End
   End
   Begin VB.Menu format 
      Caption         =   "&Format"
      Begin VB.Menu font 
         Caption         =   "Font"
         Shortcut        =   {F3}
      End
      Begin VB.Menu align 
         Caption         =   "Alignment"
         Begin VB.Menu right1 
            Caption         =   "Right"
         End
         Begin VB.Menu center 
            Caption         =   "Center"
         End
         Begin VB.Menu left 
            Caption         =   "Left"
         End
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu bullet 
         Caption         =   "Bullet"
         Shortcut        =   ^B
      End
      Begin VB.Menu strike 
         Caption         =   "Strikethrough"
         Shortcut        =   ^Y
      End
      Begin VB.Menu under 
         Caption         =   "Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu as 
         Caption         =   "-"
      End
      Begin VB.Menu case 
         Caption         =   "Change Case.."
         Begin VB.Menu up 
            Caption         =   "Upper Case"
         End
         Begin VB.Menu low 
            Caption         =   "Lower Case"
         End
         Begin VB.Menu sen 
            Caption         =   "Sentence Case"
         End
      End
      Begin VB.Menu color 
         Caption         =   "Color"
         Begin VB.Menu fc 
            Caption         =   "Forecolor"
         End
         Begin VB.Menu bc 
            Caption         =   "Backcolor"
         End
      End
   End
End
Attribute VB_Name = "note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim n As Integer
Dim aa As RichTextBox
Dim saved As Boolean
Dim cp As String

Dim striked As Boolean
Dim underlined As Boolean

Dim ques As Integer
Dim tit As String
Dim ab As Integer
Private Sub bc_Click()
    CommonDialog1.ShowColor
    rich.BackColor = CommonDialog1.color
End Sub

Private Sub bullet_Click()
    If rich.SelBullet = False Then
        rich.SelBullet = True
        bullet.Checked = True
    Else
        rich.SelBullet = False
        bullet.Checked = False
    End If
End Sub

Private Sub center_Click()
    If rich.SelLength > 0 Then
        rich.SelAlignment = 2
    End If
End Sub

Private Sub copy_Click()
    Clipboard.Clear
    Clipboard.SetText Me.ActiveControl.SelText
End Sub

Private Sub cut_Click()
    Clipboard.Clear
    Clipboard.SetText note.rich.SelText
    Me.ActiveControl.SelText = ""
End Sub

Private Sub Dt_Click()
    rich.Text = (rich.Text) & Time & Date
End Sub

Private Sub Exit_Click()
    If Len(rich.Text) <> 0 Then
        If saved = False Then
            ques = MsgBox("You like to save the contents?", vbYesNoCancel)
            If ques = 6 Then
                CommonDialog1.ShowSave
                rich.SaveFile CommonDialog1.FileName, rtfText
                Me.Caption = CommonDialog1.FileTitle
                saved = True
            ElseIf ques = "7" Then
                Unload Me
            ElseIf ques = "2" Then
                Me.Refresh
            End If
        Else
        Unload Me
        End If
    Else
      Unload Me
    End If

End Sub

Private Sub fc_Click()
    CommonDialog1.ShowColor
    rich.SelColor = CommonDialog1.color
End Sub
Private Sub font_Click()
    CommonDialog1.ShowFont
    With rich
        .SelFontName = CommonDialog1.FontName
        .SelFontSize = CommonDialog1.FontSize
        .SelBold = CommonDialog1.FontBold
        .SelItalic = CommonDialog1.FontItalic
    End With
End Sub

Private Sub Form_Load()
    rich.AutoVerbMenu = True
    saved = False
    underlined = False
    striked = False
    Clipboard.Clear
End Sub
Private Sub Form_Unload(cancel As Integer)
save_Click
End Sub

Private Sub left_Click()
    If rich.SelLength > 0 Then
        rich.SelAlignment = 0
    End If
End Sub

Private Sub low_Click()
rich.SelText = LCase(rich.SelText)
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

Private Sub new_Click()
    n = 0
    If Len(rich.Text) <> 0 Then
        If saved = False Then
            ques = MsgBox("You like to save the contents?", vbYesNoCancel)
            If ques = "6" Then
                CommonDialog1.CancelError = True
                On Error GoTo errHandler
                CommonDialog1.ShowSave
                rich.SaveFile CommonDialog1.FileName, rtfText
                Me.Caption = "Untitled - Notepad"
                rich = ""
                saved = True
            ElseIf ques = "7" Then
                Me.Caption = "Untitled - Notepad"
                rich = ""
                saved = False
            ElseIf ques = "2" Then
                Me.Refresh
            End If
        Else
            Me.Caption = "Untitled - Notepad"
            rich = ""
            saved = False
        End If
    Else
        Me.Caption = "Untitled - Notepad"
        rich = ""
        saved = False
    End If
errHandler:
    CommonDialog1.CancelError = False
    Exit Sub

End Sub

Private Sub open_Click()

    If Len(rich.Text) <> 0 Then
        If saved = False Then
            ques = MsgBox("You like to save the contents?", vbYesNoCancel)
            If ques = "6" Then
                CommonDialog1.CancelError = True
                On Error GoTo errHandler
                CommonDialog1.ShowSave
                rich.SaveFile CommonDialog1.FileName, rtfText
                On Error GoTo errHandler
                CommonDialog1.ShowOpen
                rich.LoadFile CommonDialog1.FileName, rtfText
                Me.Caption = CommonDialog1.FileTitle
            ElseIf ques = "7" Then
                CommonDialog1.ShowOpen
                rich.LoadFile CommonDialog1.FileName, rtfText
                Me.Caption = CommonDialog1.FileTitle
            ElseIf ques = "2" Then
            End If
        Else
            CommonDialog1.ShowOpen
            rich.LoadFile CommonDialog1.FileName, rtfText
            Me.Caption = CommonDialog1.FileTitle
        End If
    Else
        CommonDialog1.ShowOpen
        rich.LoadFile CommonDialog1.FileName, rtfText
        Me.Caption = CommonDialog1.FileTitle
        saved = False
    End If
errHandler:
    CommonDialog1.CancelError = False
    Exit Sub
    
End Sub

Private Sub paste_Click()
    Me.ActiveControl.SelText = Clipboard.GetText()
End Sub

Private Sub print_Click()

    Dim BeginPage, EndPage, NumCopies, I
    CommonDialog1.CancelError = True
    On Error GoTo errHandler
    CommonDialog1.ShowPrinter
    BeginPage = CommonDialog1.FromPage
    EndPage = CommonDialog1.ToPage
    NumCopies = CommonDialog1.Copies
    For I = 1 To NumCopies
        rich.SelPrint (Printer.hDc)
    Next I
    Exit Sub
errHandler:
    Exit Sub

End Sub

Private Sub replace_Click()
    Dialog1.Show
    Dialog1.Text1.SetFocus
End Sub

Private Sub rich_Change()
    If Len(Clipboard.GetText) > 0 Then
        Paste.Enabled = True
    Else
        Paste.Enabled = False
    End If
    If Save.Enabled = False Then
        Save.Enabled = True
    End If
    If UseSound = "Yes" Then
        Dim Play As String
        Play = sndPlaySound(App.Path + "\Organizerpreview\Type.wav", SND_ASYNC)
    End If
End Sub

Private Sub Rich_GotFocus()
    If Len(rich.Text) = 0 Then
        Copy.Enabled = False
        cut.Enabled = False
    End If
    If Len(Clipboard.GetText) > 0 Then
        Paste.Enabled = True
    Else
        Paste.Enabled = False
    End If
End Sub

Private Sub rich_SelChange()
    If rich.SelLength > 0 Then
        cut.Enabled = True
        Copy.Enabled = True
    Else
        cut.Enabled = False
        Copy.Enabled = False
    End If
    If rich.SelBullet Then
        bullet.Checked = True
    Else
        bullet.Checked = False
    End If
        If rich.SelStrikeThru Then
        strike.Checked = True
    Else
        strike.Checked = False
    End If
    If rich.SelUnderline Then
        under.Checked = True
    Else
        under.Checked = False
    End If
End Sub

Private Sub right1_Click()
If rich.SelLength > 0 Then
        rich.SelAlignment = 1
    End If
End Sub

Private Sub sav_Click()
    CommonDialog1.CancelError = True
    On Error GoTo errHandler
    CommonDialog1.FileName = "c:\my documents\Untitled"
    CommonDialog1.Filter = "RTF Files (*.rtf)|*.rtf|Text files (*.txt)|*.txt|Ini Files (*.ini)|*.ini|Registry Files (*.log)|*.log|Batch File (*.bat)|*.bat|All files (*.*)|*.*"
    CommonDialog1.ShowSave
    rich.SaveFile CommonDialog1.FileName, rtfText
    Me.Caption = CommonDialog1.FileTitle
    saved = True
errHandler:
    CommonDialog1.CancelError = False
    Exit Sub
End Sub

Private Sub save_Click()
If rich.Text <> "" Then
frmmain.Show
frmmain.rtbCodeWindow.Text = rich.Text
frmmain.mnuSave.Enabled = True
frmmain.mnudelete.Enabled = True
frmmain.mnunew.Enabled = True
frmmain.mnupaste.Enabled = True
frmmain.mnuModify.Enabled = True
frmmain.mnusettingz.Enabled = True
frmmain.mnuOpen.Enabled = True
frmmain.mnupaste.Enabled = False
Unload Me
End If
End Sub

Private Sub sen_Click()
Dim TextinS() As String
Dim Letter As String
Dim FinalWord As String
Dim MyItem As Integer
Dim c As Integer
TEXTIN = rich.SelText
TextinS = Split(TEXTIN, ".")
MyItem = UBound(TextinS)
For c = 0 To MyItem
Letter = Left(TextinS(c), 1)
Letter = UCase(Letter)
FinalWord = (Right(Letter, 1)) & Mid(TextinS(c), 2)
TextinS(c) = FinalWord
If UBound(TextinS) = c Then
    upallwords = upallwords & TextinS(c)
Else
    upallwords = upallwords & TextinS(c) & "."
End If
Next c
rich.SelText = upallwords
End Sub

Private Sub strike_Click()
    If rich.SelStrikeThru = False Then
        rich.SelStrikeThru = True
        strike.Checked = True
    Else
        rich.SelStrikeThru = False
        strike.Checked = False
    End If
End Sub

Private Sub under_Click()
    If rich.SelUnderline = False Then
        rich.SelUnderline = True
        underlined = True
    Else
        rich.SelUnderline = False
        underlined = False
    End If
End Sub

Private Sub up_Click()
rich.SelText = UCase(rich.SelText)
End Sub

