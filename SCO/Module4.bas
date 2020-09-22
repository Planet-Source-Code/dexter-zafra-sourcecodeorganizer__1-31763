Attribute VB_Name = "Module4"

Public WidthValueInTag As String
Public HeightValueInTag As String
Public ValueValueInTag As String
Public NameValueInTag As String
Public SizeValueInTag As String
Public ColorValueInTag As String
Public BgColorValueInTag As String
Public AlignValueInTag As String
Public HrefValueInTag As String
Public TargetValueInTag As String
Public TitleValueInTag As String
Public TextValueInTag As String
Public LinkValueInTag As String
Public VlinkValueInTag As String
Public AlinkValueInTag As String
Public BorderValueInTag As String
Public HspaceValueInTag As String
Public VspaceValueInTag As String
Public UsemapValueInTag As String
Public AltValueInTag As String
Public ColsValueInTag As String
Public RowsValueInTag As String
Public MaxValueInTag As String
Public BcValueInTag As String
Public CellsValueInTag As String
Public CellPValueInTag As String
Public ColspanValueInTag As String
Public RowpanValueInTag As String
Public ValignValueInTag As String
Public FaceValueInTag As String
Public IDValueInTag As String
Public BgValueInTag As String
Public SrcValueInTag As String
'@#$

'Api & etc.
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public frmmain As frmmain
Public TabNumber As Integer
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const FOF_ALLOWUNDO = &H40
Public Const WM_USER = &H400
Public Const EM_UNDO = &HC7
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1

Sub Main()
On Error Resume Next

    frmSplash.Show
    frmSplash.Refresh
    Set frmmain = New frmmain
    Load frmmain
    Unload frmSplash
    frmmain.Show
    
End Sub

Public Sub FormDrag(TheForm As Form)
On Error Resume Next
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Function GetWidthProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    WidthP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "width="
                WidthP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "width=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  WidthValue = GetString(2, Tag, "idth=")
  WidthValue = Mid(WidthValue, 6, Len(WidthValue))
  intPos = InStr(WidthValue, Chr(34))
  WidthValue = Mid(WidthValue, 1, intPos - 1)
 Else
  WidthValue = GetString(2, Tag, "idth=")
  WidthValue = Mid(WidthValue, 5, Len(WidthValue))
   If InStr(WidthValue, " ") <= 0 Then
      intPos = InStr(WidthValue, ">")
      WidthValue = Mid(WidthValue, 1, intPos - 1)
   Else
      intPos = InStr(WidthValue, " ")
      WidthValue = Mid(WidthValue, 1, intPos - 1)
   End If
 End If
  WidthValueInTag = WidthValue
End Function

Function GetHeightProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    HeightP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "height="
                HeightP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

 Tag = GetString(2, Tag, "height=")
 Tag = GetString(1, Tag, " ")
 
 If InStr(Tag, Chr(34)) > 0 Then
  HeightValue = GetString(2, Tag, "eight=")
  HeightValue = Mid(HeightValue, 7, Len(HeightValue))
  intPos = InStr(HeightValue, Chr(34))
  HeightValue = Mid(HeightValue, 1, intPos - 1)
 Else
  HeightValue = GetString(2, Tag, "eight=")
  HeightValue = Mid(HeightValue, 6, Len(HeightValue))
   If InStr(HeightValue, " ") <= 0 Then
      intPos = InStr(HeightValue, ">")
      HeightValue = Mid(HeightValue, 1, intPos - 1)
   Else
      intPos = InStr(HeightValue, " ")
      HeightValue = Mid(HeightValue, 1, intPos - 1)
   End If
 End If
  HeightValueInTag = HeightValue
End Function

Function GetValueProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    ValueP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "value="
                ValueP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "value=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  ValueValue = GetString(2, Tag, "alue=")
  ValueValue = Mid(ValueValue, 6, Len(ValueValue))
  intPos = InStr(ValueValue, Chr(34))
  ValueValue = Mid(ValueValue, 1, intPos - 1)
 Else
  ValueValue = GetString(2, Tag, "alue=")
  ValueValue = Mid(ValueValue, 5, Len(ValueValue))
   If InStr(ValueValue, " ") <= 0 Then
      intPos = InStr(ValueValue, ">")
      ValueValue = Mid(ValueValue, 1, intPos - 1)
   Else                                           '
      intPos = InStr(ValueValue, " ")
      ValueValue = Mid(ValueValue, 1, intPos - 1)
   End If
 End If
  ValueValueInTag = ValueValue
End Function

'# Name property of the TAG
Function GetNameProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    NameP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "name="
                NameP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "name=")
Tag = GetString(1, Tag, " ")
 If InStr(Tag, Chr(34)) > 0 Then
  NameValue = GetString(2, Tag, "ame=")
  NameValue = Mid(NameValue, 5, Len(NameValue))
  intPos = InStr(NameValue, Chr(34))
  NameValue = Mid(NameValue, 1, intPos - 1)
 Else
  NameValue = GetString(2, Tag, "me=")
  NameValue = Mid(NameValue, 3, Len(NameValue))
   If InStr(NameValue, " ") <= 0 Then
      intPos = InStr(NameValue, ">")
      NameValue = Mid(NameValue, 1, intPos - 1)
   Else
      intPos = InStr(NameValue, " ")
      NameValue = Mid(NameValue, 1, intPos - 1)
   End If
 End If
  NameValueInTag = UCase(NameValue)
End Function
Function GetSizeProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    SizeP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "size="
                SizeP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "size=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  SizeValue = GetString(2, Tag, "ize=")
  SizeValue = Mid(SizeValue, 5, Len(SizeValue))
  intPos = InStr(SizeValue, Chr(34))
  SizeValue = Mid(SizeValue, 1, intPos - 1)
 Else
  SizeValue = GetString(2, Tag, "ize=")
  SizeValue = Mid(SizeValue, 4, Len(SizeValue))
   If InStr(SizeValue, " ") <= 0 Then
      intPos = InStr(SizeValue, ">")
      SizeValue = Mid(SizeValue, 1, intPos - 1)
   Else
      intPos = InStr(SizeValue, " ")
      SizeValue = Mid(SizeValue, 1, intPos - 1)
   End If
 End If
  SizeValueInTag = SizeValue
End Function

Function GetColorProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    ColorP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "color="
                ColorP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "color=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  ColorValue = GetString(2, Tag, "olor=")
  ColorValue = Mid(ColorValue, 6, Len(ColorValue))
  intPos = InStr(ColorValue, Chr(34))
  ColorValue = Mid(ColorValue, 1, intPos - 1)
 Else
  ColorValue = GetString(2, Tag, "olor=")
  ColorValue = Mid(ColorValue, 5, Len(ColorValue))
   If InStr(ColorValue, " ") <= 0 Then
      intPos = InStr(ColorValue, ">")
      ColorValue = Mid(ColorValue, 1, intPos - 1)
   Else
      intPos = InStr(ColorValue, " ")
      ColorValue = Mid(ColorValue, 1, intPos - 1)
   End If
 End If
  ColorValueInTag = UCase(ColorValue)
End Function

Function GetBgColorProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    BgColorP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "bgcolor="
                BgColorP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "bgcolor=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  bgcolorvalue = GetString(2, Tag, "gcolor=")
  bgcolorvalue = Mid(bgcolorvalue, 8, Len(bgcolorvalue))
  intPos = InStr(bgcolorvalue, Chr(34))
  bgcolorvalue = Mid(bgcolorvalue, 1, intPos - 1)

 Else
  bgcolorvalue = GetString(2, Tag, "gcolor=")
  bgcolorvalue = Mid(bgcolorvalue, 7, Len(bgcolorvalue))
   If InStr(bgcolorvalue, " ") <= 0 Then
      intPos = InStr(bgcolorvalue, ">")
      bgcolorvalue = Mid(bgcolorvalue, 1, intPos - 1)
   Else
      intPos = InStr(bgcolorvalue, " ")
      bgcolorvalue = Mid(bgcolorvalue, 1, intPos - 1)
   End If
 End If
  BgColorValueInTag = UCase(bgcolorvalue)
End Function
Function GetAlignProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    AlignP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "align="
                AlignP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "align=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  AlignValue = GetString(2, Tag, "lign=")
  AlignValue = Mid(AlignValue, 6, Len(AlignValue))
  intPos = InStr(AlignValue, Chr(34))
  AlignValue = Mid(AlignValue, 1, intPos - 1)
 Else
  AlignValue = GetString(2, Tag, "lign=")
  AlignValue = Mid(AlignValue, 5, Len(AlignValue))
   If InStr(AlignValue, " ") <= 0 Then
      intPos = InStr(AlignValue, ">")
      AlignValue = Mid(AlignValue, 1, intPos - 1)
   Else
      intPos = InStr(AlignValue, " ")
      AlignValue = Mid(AlignValue, 1, intPos - 1)
   End If
 End If
  AlignValueInTag = UCase(AlignValue)
End Function

'# HREF property of the TAG
Function GetHrefProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    HrefP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "href="
                HrefP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "href=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  HrefValue = GetString(2, Tag, "ref=")
  HrefValue = Mid(HrefValue, 5, Len(HrefValue))
  intPos = InStr(HrefValue, Chr(34))
  HrefValue = Mid(HrefValue, 1, intPos - 1)
 Else
  HrefValue = GetString(2, Tag, "ef=")
  HrefValue = Mid(HrefValue, 3, Len(HrefValue))
   If InStr(HrefValue, " ") <= 0 Then
      intPos = InStr(HrefValue, ">")
      HrefValue = Mid(HrefValue, 1, intPos - 1)
   Else
      intPos = InStr(HrefValue, " ")
      HrefValue = Mid(HrefValue, 1, intPos - 1)
   End If
 End If
  HrefValueInTag = HrefValue
End Function

Function GetTargetProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    TargetP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "target="
                TargetP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "target=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then                  'If the tag looks like ... width="xx"
  TargetValue = GetString(2, Tag, "arget=")
  TargetValue = Mid(TargetValue, 7, Len(TargetValue))
  intPos = InStr(TargetValue, Chr(34))
  TargetValue = Mid(TargetValue, 1, intPos - 1)
 Else
  TargetValue = GetString(2, Tag, "rget=")
  TargetValue = Mid(TargetValue, 5, Len(TargetValue))
   If InStr(TargetValue, " ") <= 0 Then
      intPos = InStr(TargetValue, ">")
      TargetValue = Mid(TargetValue, 1, intPos - 1)
   Else
      intPos = InStr(TargetValue, " ")
      TargetValue = Mid(TargetValue, 1, intPos - 1)
   End If
 End If
  TargetValueInTag = TargetValue
End Function

Function GetTitleProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    TitleP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "title="
                TitleP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "title=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  TitleValue = GetString(2, Tag, "itle=")
  TitleValue = Mid(TitleValue, 6, Len(TitleValue))
  intPos = InStr(TitleValue, Chr(34))
  TitleValue = Mid(TitleValue, 1, intPos - 1)
 Else
  TitleValue = GetString(2, Tag, "itle=")
  TitleValue = Mid(TitleValue, 5, Len(TitleValue))
   If InStr(TitleValue, " ") <= 0 Then
      intPos = InStr(TitleValue, ">")
      TitleValue = Mid(TitleValue, 1, intPos - 1)
   Else
      intPos = InStr(TitleValue, " ")
      TitleValue = Mid(TitleValue, 1, intPos - 1)
   End If
 End If
  TitleValueInTag = TitleValue
End Function
Function GetTextProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(fMainForm.ActiveForm.rtfText.SelText)

    Dim intPos As Integer
    TextP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "text="
                TextP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "text=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  TextValue = GetString(2, Tag, "ext=")
  TextValue = Mid(TextValue, 5, Len(TextValue))
  intPos = InStr(TextValue, Chr(34))
  TextValue = Mid(TextValue, 1, intPos - 1)
 Else
  TextValue = GetString(2, Tag, "xt=")
  TextValue = Mid(TextValue, 3, Len(TextValue))
   If InStr(TextValue, " ") <= 0 Then
      intPos = InStr(TextValue, ">")
      TextValue = Mid(TextValue, 1, intPos - 1)
   Else
      intPos = InStr(TextValue, " ")
      TextValue = Mid(TextValue, 1, intPos - 1)
   End If
 End If
  TextValueInTag = UCase(TextValue)
End Function

Function GetLinkProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    LinkP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "link="
                LinkP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "link=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  LinkValue = GetString(2, Tag, "ink=")
  LinkValue = Mid(LinkValue, 5, Len(LinkValue))
  intPos = InStr(LinkValue, Chr(34))
  LinkValue = Mid(LinkValue, 1, intPos - 1)
 Else
  LinkValue = GetString(2, Tag, "nk=")
  LinkValue = Mid(LinkValue, 3, Len(LinkValue))
   If InStr(LinkValue, " ") <= 0 Then
      intPos = InStr(LinkValue, ">")
      LinkValue = Mid(LinkValue, 1, intPos - 1)
   Else
      intPos = InStr(LinkValue, " ")
      LinkValue = Mid(LinkValue, 1, intPos - 1)
   End If
 End If
  LinkValueInTag = UCase(LinkValue)
End Function

Function GetVlinkProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    VlinkP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "vlink="
                VlinkP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "vlink=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  VlinkValue = GetString(2, Tag, "link=")
  VlinkValue = Mid(VlinkValue, 6, Len(VlinkValue))
  intPos = InStr(VlinkValue, Chr(34))
  VlinkValue = Mid(VlinkValue, 1, intPos - 1)
 Else
  VlinkValue = GetString(2, Tag, "link=")
  VlinkValue = Mid(VlinkValue, 5, Len(VlinkValue))
   If InStr(VlinkValue, " ") <= 0 Then
      intPos = InStr(VlinkValue, ">")
      VlinkValue = Mid(VlinkValue, 1, intPos - 1)
   Else
      intPos = InStr(VlinkValue, " ")
      VlinkValue = Mid(VlinkValue, 1, intPos - 1)
   End If
 End If
  VlinkValueInTag = UCase(VlinkValue)
End Function


'# Alink property of the TAG
Function GetAlinkProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    AlinkP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "alink="
                AlinkP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "alink=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  AlinkValue = GetString(2, Tag, "link=")
  AlinkValue = Mid(AlinkValue, 6, Len(AlinkValue))
  intPos = InStr(AlinkValue, Chr(34))
  AlinkValue = Mid(AlinkValue, 1, intPos - 1)
 Else
  AlinkValue = GetString(2, Tag, "link=")
  AlinkValue = Mid(AlinkValue, 5, Len(AlinkValue))
   If InStr(AlinkValue, " ") <= 0 Then
      intPos = InStr(AlinkValue, ">")
      AlinkValue = Mid(AlinkValue, 1, intPos - 1)
   Else
      intPos = InStr(AlinkValue, " ")
      AlinkValue = Mid(AlinkValue, 1, intPos - 1)
   End If
 End If
  AlinkValueInTag = UCase(AlinkValue)
End Function

Function GetBorderProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    BorderP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "border="
                BorderP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "border=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  BorderValue = GetString(2, Tag, "order=")
  BorderValue = Mid(BorderValue, 7, Len(BorderValue))
  intPos = InStr(BorderValue, Chr(34))
  BorderValue = Mid(BorderValue, 1, intPos - 1)
 Else
  BorderValue = GetString(2, Tag, "rder=")
  BorderValue = Mid(BorderValue, 5, Len(BorderValue))
   If InStr(BorderValue, " ") <= 0 Then
      intPos = InStr(BorderValue, ">")
      BorderValue = Mid(BorderValue, 1, intPos - 1)
   Else
      intPos = InStr(BorderValue, " ")
      BorderValue = Mid(BorderValue, 1, intPos - 1)
   End If
 End If
  BorderValueInTag = BorderValue
End Function

Function GetHspaceProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    HspaceP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "hspace="
                HspaceP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "hspace=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  HspaceValue = GetString(2, Tag, "space=")
  HspaceValue = Mid(HspaceValue, 7, Len(HspaceValue))
  intPos = InStr(HspaceValue, Chr(34))
  HspaceValue = Mid(HspaceValue, 1, intPos - 1)
 Else
  HspaceValue = GetString(2, Tag, "pace=")
  HspaceValue = Mid(HspaceValue, 5, Len(HspaceValue))
   If InStr(HspaceValue, " ") <= 0 Then
      intPos = InStr(HspaceValue, ">")
      HspaceValue = Mid(HspaceValue, 1, intPos - 1)
   Else
      intPos = InStr(HspaceValue, " ")
      HspaceValue = Mid(HspaceValue, 1, intPos - 1)
   End If
 End If
  HspaceValueInTag = HspaceValue
End Function

Function GetVspaceProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    VspaceP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "vspace="
                VspaceP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "vspace=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  VspaceValue = GetString(2, Tag, "space=")
  VspaceValue = Mid(VspaceValue, 7, Len(VspaceValue))
  intPos = InStr(VspaceValue, Chr(34))
  VspaceValue = Mid(VspaceValue, 1, intPos - 1)
 Else
  VspaceValue = GetString(2, Tag, "pace=")
  VspaceValue = Mid(VspaceValue, 5, Len(VspaceValue))
   If InStr(VspaceValue, " ") <= 0 Then
      intPos = InStr(VspaceValue, ">")
      VspaceValue = Mid(VspaceValue, 1, intPos - 1)
   Else                                           'If the tag looks like ... width="xx">
      intPos = InStr(VspaceValue, " ")
      HspaceValue = Mid(VspaceValue, 1, intPos - 1)
   End If
 End If
  VspaceValueInTag = VspaceValue
End Function

'#Usemap property of the TAG <a>
Function GetUsemapProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    UsemapP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "usemap="
                UsemapP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "usemap=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  UsemapValue = GetString(2, Tag, "semap=")
  UsemapValue = Mid(UsemapValue, 7, Len(UsemapValue))
  intPos = InStr(UsemapValue, Chr(34))
  UsemapValue = Mid(UsemapValue, 1, intPos - 1)
 Else
  UsemapValue = GetString(2, Tag, "emap=")
  UsemapValue = Mid(UsemapValue, 5, Len(UsemapValue))
   If InStr(UsemapValue, " ") <= 0 Then
      intPos = InStr(UsemapValue, ">")
      UsemapValue = Mid(UsemapValue, 1, intPos - 1)
   Else
      intPos = InStr(UsemapValue, " ")
      UsemapValue = Mid(UsemapValue, 1, intPos - 1)
   End If
 End If
  UsemapValueInTag = UsemapValue
End Function

Function GetAltProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    AltP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "alt="
                AltP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "alt=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  AltValue = GetString(2, Tag, "lt=")
  AltValue = Mid(AltValue, 4, Len(AltValue))
  intPos = InStr(AltValue, Chr(34))
  AltValue = Mid(AltValue, 1, intPos - 1)
 Else
  AltValue = GetString(2, Tag, "t=")
  AltValue = Mid(AltValue, 2, Len(AltValue))
   If InStr(AltValue, " ") <= 0 Then
      intPos = InStr(AltValue, ">")
      AltValue = Mid(AltValue, 1, intPos - 1)
   Else
      intPos = InStr(AltValue, " ")
      AltValue = Mid(AltValue, 1, intPos - 1)
   End If
 End If
  AltValueInTag = AltValue
End Function

'# Cols property of the TAG
Function GetColsProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    ColsP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "cols="
                ColsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "cols=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  ColsValue = GetString(2, Tag, "ols=")
  ColsValue = Mid(ColsValue, 5, Len(ColsValue))
  intPos = InStr(ColsValue, Chr(34))
  ColsValue = Mid(ColsValue, 1, intPos - 1)
 Else
  ColsValue = GetString(2, Tag, "ls=")
  ColsValue = Mid(ColsValue, 3, Len(ColsValue))
   If InStr(ColsValue, " ") <= 0 Then
      intPos = InStr(ColsValue, ">")
      ColsValue = Mid(ColsValue, 1, intPos - 1)
   Else
      intPos = InStr(ColsValue, " ")
      ColsValue = Mid(ColsValue, 1, intPos - 1)
   End If
 End If
  ColsValueInTag = ColsValue
End Function

'# Rows property of the TAG
Function GetRowsProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    RowsP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "rows="
                RowsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "rows=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  RowsValue = GetString(2, Tag, "ows=")
  RowsValue = Mid(RowsValue, 5, Len(RowsValue))
  intPos = InStr(RowsValue, Chr(34))
  RowsValue = Mid(RowsValue, 1, intPos - 1)
 Else
  RowsValue = GetString(2, Tag, "ws=")
  RowsValue = Mid(RowsValue, 3, Len(RowsValue))
   If InStr(RowsValue, " ") <= 0 Then
      intPos = InStr(RowsValue, ">")
      RowsValue = Mid(RowsValue, 1, intPos - 1)
   Else
      intPos = InStr(RowsValue, " ")
      RowsValue = Mid(RowsValue, 1, intPos - 1)
   End If
 End If
  RowsValueInTag = RowsValue
End Function

Function GetMaxLProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    MaxP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "maxlength="
                MaxP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "maxlength=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  MaxValue = GetString(2, Tag, "axlength=")
  MaxValue = Mid(MaxValue, 10, Len(MaxValue))
  intPos = InStr(MaxValue, Chr(34))
  MaxValue = Mid(MaxValue, 1, intPos - 1)
 Else
  MaxValue = GetString(2, Tag, "axlength=")
  MaxValue = Mid(MaxValue, 9, Len(MaxValue))
   If InStr(MaxValue, " ") <= 0 Then
      intPos = InStr(MaxValue, ">")
      MaxValue = Mid(MaxValue, 1, intPos - 1)
   Else
      intPos = InStr(MaxValue, " ")
      MaxValue = Mid(MaxValue, 1, intPos - 1)
   End If
 End If
  MaxValueInTag = MaxValue
End Function

Function GetBorderCProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    BcP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "bordercolor="
                BcP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "bordercolor=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  BcValue = GetString(2, Tag, "ordercolor=")
  BcValue = Mid(BcValue, 12, Len(BcValue))
  intPos = InStr(BcValue, Chr(34))
  BcValue = Mid(BcValue, 1, intPos - 1)
 Else
  BcValue = GetString(2, Tag, "ordercolor=")
  BcValue = Mid(BcValue, 11, Len(BcValue))
   If InStr(BcValue, " ") <= 0 Then
      intPos = InStr(BcValue, ">")
      BcValue = Mid(BcValue, 1, intPos - 1)
   Else
      intPos = InStr(BcValue, " ")
      BcValue = Mid(BcValue, 1, intPos - 1)
   End If
 End If
  BcValueInTag = UCase(BcValue)
End Function

Function GetCellSProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    CellsP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "cellspacing="
                CellsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "cellspacing=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  CellsValue = GetString(2, Tag, "ellspacing=")
  CellsValue = Mid(CellsValue, 12, Len(CellsValue))
  intPos = InStr(CellsValue, Chr(34))
  CellsValue = Mid(CellsValue, 1, intPos - 1)
 Else
  CellsValue = GetString(2, Tag, "ellspacing=")
  CellsValue = Mid(CellsValue, 11, Len(CellsValue))
   If InStr(CellsValue, " ") <= 0 Then
      intPos = InStr(CellsValue, ">")
      CellsValue = Mid(CellsValue, 1, intPos - 1)
   Else
      intPos = InStr(CellsValue, " ")
      CellsValue = Mid(CellsValue, 1, intPos - 1)
   End If
 End If
  CellsValueInTag = CellsValue
End Function

Function GetCellPProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    CellP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "cellpadding="
                CellsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "cellpadding=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  CellsValue = GetString(2, Tag, "ellpadding=")
  CellsValue = Mid(CellsValue, 12, Len(CellsValue))
  intPos = InStr(CellsValue, Chr(34))
  CellsValue = Mid(CellsValue, 1, intPos - 1)
 Else
  CellsValue = GetString(2, Tag, "ellpadding=")
  CellsValue = Mid(CellsValue, 11, Len(CellsValue))
   If InStr(CellsValue, " ") <= 0 Then
      intPos = InStr(CellsValue, ">")
      CellsValue = Mid(CellsValue, 1, intPos - 1)
   Else
      CellsValue = Mid(CellsValue, 1, intPos - 1)
   End If
 End If
  CellPValueInTag = CellsValue
End Function

Function GetColspanProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    ColsP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "colspan="
                ColsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "colspan=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  VspaceValue = GetString(2, Tag, "olspan=")
  VspaceValue = Mid(VspaceValue, 8, Len(VspaceValue))
  intPos = InStr(VspaceValue, Chr(34))
  VspaceValue = Mid(VspaceValue, 1, intPos - 1)
 Else
  VspaceValue = GetString(2, Tag, "olspan=")
  VspaceValue = Mid(VspaceValue, 7, Len(VspaceValue))
   If InStr(VspaceValue, " ") <= 0 Then
      intPos = InStr(VspaceValue, ">")
      VspaceValue = Mid(VspaceValue, 1, intPos - 1)
   Else
      intPos = InStr(VspaceValue, " ")
      HspaceValue = Mid(VspaceValue, 1, intPos - 1)
   End If
 End If
  ColspanValueInTag = VspaceValue
End Function

Function GetRowspanProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    ColsP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "rowspan="
                ColsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "rowspan=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  VspaceValue = GetString(2, Tag, "owspan=")
  VspaceValue = Mid(VspaceValue, 8, Len(VspaceValue))
  intPos = InStr(VspaceValue, Chr(34))
  VspaceValue = Mid(VspaceValue, 1, intPos - 1)
 Else
  VspaceValue = GetString(2, Tag, "owspan=")
  VspaceValue = Mid(VspaceValue, 7, Len(VspaceValue))
   If InStr(VspaceValue, " ") <= 0 Then
      intPos = InStr(VspaceValue, ">")
      VspaceValue = Mid(VspaceValue, 1, intPos - 1)
   Else
      intPos = InStr(VspaceValue, " ")
      HspaceValue = Mid(VspaceValue, 1, intPos - 1)
   End If
 End If
  RowpanValueInTag = VspaceValue
End Function

Function GetValignProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    TargetP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "valign="
                TargetP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "valign=")
Tag = GetString(1, Tag, " ")

 If InStr(Tag, Chr(34)) > 0 Then
  TargetValue = GetString(2, Tag, "align=")
  TargetValue = Mid(TargetValue, 7, Len(TargetValue))
  intPos = InStr(TargetValue, Chr(34))
  TargetValue = Mid(TargetValue, 1, intPos - 1)
 Else
  TargetValue = GetString(2, Tag, "lign=")
  TargetValue = Mid(TargetValue, 5, Len(TargetValue))
   If InStr(TargetValue, " ") <= 0 Then
      intPos = InStr(TargetValue, ">")
      TargetValue = Mid(TargetValue, 1, intPos - 1)
   Else
      TargetValue = Mid(TargetValue, 1, intPos - 1)
   End If
 End If
  ValignValueInTag = UCase(TargetValue)
End Function

Function GetFaceProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    RowsP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "face="
                RowsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "face=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  RowsValue = GetString(2, Tag, "ace=")
  RowsValue = Mid(RowsValue, 5, Len(RowsValue))
  intPos = InStr(RowsValue, Chr(34))
  RowsValue = Mid(RowsValue, 1, intPos - 1)
 Else
  RowsValue = GetString(2, Tag, "ce=")
  RowsValue = Mid(RowsValue, 3, Len(RowsValue))
   If InStr(RowsValue, " ") <= 0 Then
      intPos = InStr(RowsValue, ">")
      RowsValue = Mid(RowsValue, 1, intPos - 1)
   Else
      intPos = InStr(RowsValue, " ")
      RowsValue = Mid(RowsValue, 1, intPos - 1)
   End If
 End If
  FaceValueInTag = LCase(RowsValue)
End Function

Function GetIDProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    AltP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "id="
                AltP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "id=")
Tag = GetString(1, Tag, " ")


 If InStr(Tag, Chr(34)) > 0 Then
  AltValue = GetString(2, Tag, "d=")
  AltValue = Mid(AltValue, 3, Len(AltValue))
  intPos = InStr(AltValue, Chr(34))
  AltValue = Mid(AltValue, 1, intPos - 1)
 Else
  AltValue = GetString(2, Tag, "d=")
  AltValue = Mid(AltValue, 2, Len(AltValue))
   If InStr(AltValue, " ") <= 0 Then
      intPos = InStr(AltValue, ">")
      AltValue = Mid(AltValue, 1, intPos - 1)
   Else
      intPos = InStr(AltValue, " ")
      AltValue = Mid(AltValue, 1, intPos - 1)
   End If
 End If
  IDValueInTag = AltValue
End Function

Function GetBGPProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    CellP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "background="
                CellsP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "background=")
Tag = GetString(2, Tag, Chr(34))

 If InStr(Tag, Chr(34)) > 0 Then
  CellsValue = GetString(2, Tag, "ackground=")
  CellsValue = Mid(CellsValue, 12, Len(CellsValue))
  intPos = InStr(CellsValue, Chr(34))
  CellsValue = Mid(CellsValue, 1, intPos - 1)
 Else
  BgValueInTag = Tag
 End If
 If BgValueInTag = "" Then BgValueInTag = CellsValue
End Function

Function GetSrcProperty()
On Error Resume Next
 Dim Tag As String
 Tag = LCase(frmmain.rtbnet.SelText)

    Dim intPos As Integer
    AltP = vbNullString
    intPos = Len(Tag)
    Do While intPos > 0
        Select Case Mid$(Tag, intPos, 6)
            Case "src="
                AltP = Mid$(Tag, intPos + 6, 4)
                Exit Do
            Case Else
        End Select
        intPos = intPos - 1
    Loop

Tag = GetString(2, Tag, "src=")
Tag = GetString(2, Tag, Chr(34))


 If InStr(Tag, Chr(34)) > 0 Then
  AltValue = GetString(2, Tag, "rc=")
  AltValue = Mid(AltValue, 4, Len(AltValue))
  intPos = InStr(AltValue, Chr(34))
  AltValue = Mid(AltValue, 1, intPos - 1)
 End If
  SrcValueInTag = Tag
End Function

Function GetString(ArgNum As Integer, srchstr As String, Delim As String) As String
On Error Resume Next
Dim ArgCount As Integer, LastPos As Integer, Pos As Integer, Arg As String
    Arg = ""
    LastPos = 1
    If ArgNum = 1 Then Arg = srchstr
    Do While InStr(srchstr, Delim) > 0
        Pos = InStr(LastPos, srchstr, Delim)
        If Pos = 0 Then
            If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
            Exit Do
        Else
            ArgCount = ArgCount + 1
            If ArgCount = ArgNum Then
                Arg = Mid(srchstr, LastPos, Pos - LastPos)
                Exit Do
            End If
        End If
        LastPos = Pos + 1
    Loop
    GetString = Arg
End Function




