VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Browser 
   Caption         =   "Oragnizer - HTML Preview "
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   HelpContextID   =   400
   Icon            =   "Browser.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   7845
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Text            =   "URL Address"
      Top             =   5400
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6000
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":06B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":0A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":0E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":11AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":1552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6000
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Browser.frx":18FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7470
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17641
            MinWidth        =   17641
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   3421
            MinWidth        =   3421
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   4080
      Left            =   0
      TabIndex        =   0
      Top             =   105
      Width           =   5895
      ExtentX         =   10398
      ExtentY         =   7197
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ProgressBar PBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   7335
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If Text1.Text <> "URL Address" Then Web.Navigate (Text1.Text)
End Sub

Private Sub Form_Resize()
Web.Height = Browser.Height - 1000
Web.Width = Browser.Width - 100
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Web.Navigate Me.Text1.Text
    End If
End Sub

Private Sub Web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Text1.Text = URL
End Sub

Private Sub Web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    On Error GoTo progressERR
    If Progress = -1 Then PBar.Value = 100

    If Progress > 0 And ProgressMax > 0 Then
        PBar.Value = Progress * 100 / ProgressMax
       
    End If
    PBar.Visible = False
    Exit Sub
progressERR:
End Sub


Private Sub Web_TitleChange(ByVal Text As String)
    Me.Caption = "Browsing :" & Text
End Sub
