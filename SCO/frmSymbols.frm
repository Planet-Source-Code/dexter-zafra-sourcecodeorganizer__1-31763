VERSION 5.00
Begin VB.Form Symbols 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert symbol from "
   ClientHeight    =   5295
   ClientLeft      =   3675
   ClientTop       =   2610
   ClientWidth     =   6375
   ControlBox      =   0   'False
   HelpContextID   =   1290
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInsert 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   160
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1095
   End
   Begin VB.ComboBox cboFonts 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSymbols.frx":0E42
      Left            =   1920
      List            =   "frmSymbols.frx":0E44
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCopy 
      Cancel          =   -1  'True
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   1092
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3732
      Left            =   6000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   252
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1092
   End
   Begin VB.PictureBox picHolder 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   1080
      Width           =   5895
      Begin VB.Label lblBigDisplay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   720
         Left            =   840
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblsymbols 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4920
      Width           =   4815
   End
   Begin VB.Label lblLabel 
      Caption         =   "Insert string:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      Caption         =   "All Symbols contained in:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   630
      Width           =   1815
   End
End
Attribute VB_Name = "Symbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurrentLabel As Integer
Private noperline As Integer
Private linesout As Integer
Private gignore As Boolean
Private minuschars As Integer
Private fntFont As String
Private blnLoadedFonts As Boolean
Private Const BorderWidth As Integer = 100

Private Sub cboFonts_Click()
    lblBigDisplay.Visible = False
    Dim I As Integer
    If lblsymbols(0).FontName <> cboFonts.Text Then
        For I = 0 To lblsymbols.Count - 1
        Next
    End If
    If lblBigDisplay.FontName <> cboFonts.Text Then
    End If
    Me.Caption = "Insert symbol from " & lblBigDisplay.FontName
    txtInsert.font = lblBigDisplay.FontName
    lblBigDisplay.Visible = False
End Sub

Private Sub cboFonts_DropDown()
    Dim I As Integer
    If Not (blnLoadedFonts) Then
        MousePointer = vbArrowHourglass
        cboFonts.Clear
        For I = 0 To Printer.FontCount - 1
            cboFonts.AddItem Screen.Fonts(I)
        Next I
        MousePointer = vbDefault
        blnLoadedFonts = True
        On Error Resume Next
        cboFonts.Text = fntFont
    End If
End Sub
Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub cmdCopy_Click()
    On Error Resume Next
    Clipboard.SetText txtInsert.Text
    picHolder.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    picHolder.SetFocus
    frmmain.rtbnet.SelLength = cboFonts.Text
    frmmain.rtbnet.SelText = ""
    frmmain.rtbnet.SelText = txtInsert.Text
    lblBigDisplay.Visible = False
    txtInsert.Text = ""
    Unload Me
End Sub

Private Sub cmdInsert_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyUp
        If Shift = 0 Then
            picHolder_KeyDown KeyCode, Shift
        End If
        KeyCode = 0
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    blnLoadedFonts = False
    fntFont = frmmain.rtbnet.font
    Me.Caption = "Insert symbol from " & fntFont
    lblMessage = "Symbols contained in: "
    lblBigDisplay.font = fntFont
    noperline = 0
    lblsymbols(0).font = fntFont
    FillSymbols (0)
    gignore = True
    VScroll1.Max = linesout
    VScroll1.Min = 0
    gignore = False
    CurrentLabel = 0
    cboFonts.AddItem (fntFont), 0
    cboFonts.ListIndex = 0
End Sub
Sub FillSymbols(ByVal startnumber As Integer)
    gignore = False
    minuschars = 1
    numberoflines = 1
    lblsymbols(0).Left = -5000
    linesout = 0
    picHolder.Visible = False
    For I = 1 To 223
        Load lblsymbols(I)
        On Error GoTo 0
        currentchar = I + startnumber + 32
        If currentchar > 255 Then Exit For
        lblsymbols(I).Caption = Chr(currentchar)
        NewLeftPos = BorderWidth + ((I) - minuschars) * (lblsymbols(I).Width - 20)
        If NewLeftPos > picHolder.Width - lblsymbols(I).Width Then
            minuschars = lblsymbols.Count - 1
            
            If noperline = 0 Then noperline = lblsymbols.Count - 2
           
            newtop = (numberoflines) * (lblsymbols(I).Height - 20)
            If newtop + lblsymbols(I).Height > picHolder.Height Then
                linesout = linesout + 1
            End If
            NewLeftPos = BorderWidth + (I - minuschars) * (lblsymbols(I).Width - 20)
        End If
        lblsymbols(I).Top = (numberoflines - 0.7) * (lblsymbols(I).Height - 20)
        lblsymbols(I).Left = NewLeftPos
        lblsymbols(I).Visible = True
    Next
    picHolder.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblBigDisplay_DblClick()
    txtInsert.Text = txtInsert.Text & lblsymbols(CurrentLabel).Caption
End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub lblsymbols_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo errHandler
    lblBigDisplay.Left = lblsymbols(Index).Left - ((lblBigDisplay.Width - lblsymbols(Index).Width) / 2)
    lblBigDisplay.Top = lblsymbols(Index).Top - ((lblBigDisplay.Height - lblsymbols(Index).Height) / 2)
    lblBigDisplay.Caption = lblsymbols(Index).Caption
    lblBigDisplay.Visible = True
    CurrentLabel = Index
    fred = lblsymbols(Index).Caption
    lblStatus.Caption = "Special Char " & Asc(fred)
errHandler:

End Sub

Private Sub txtInsert_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblBigDisplay.Visible = False
End Sub

Private Sub VScroll1_Change()
    If Not gignore Then
        MousePointer = vbHourglass
        For Each Label In lblsymbols
            If Not Label.Index = 0 Then
                Unload Label
            End If
        Next
        charstart = VScroll1.Value * noperline
        FillSymbols (charstart)
        MousePointer = vbDefault
    End If
    lblBigDisplay.Visible = False
    picHolder.SetFocus
End Sub
