Attribute VB_Name = "Module1"
Option Explicit
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function Execute(ByVal URL As String) As Long
  Execute = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Public Function Search(Engine As Integer, SearchString As String)
Dim I As Integer
  I = 1
  While I <= Len(SearchString)
    If Mid(SearchString, I, 1) = " " Then Mid(SearchString, I, 1) = "+"
    I = I + 1
  Wend
  Select Case Engine
  Case 0
    Execute "http://www.altavista.com/cgi-bin/query?pg=q&what=web&fmt=.&q=" & SearchString
  Case 1
    Execute "http://astalavista.box.sk/cgi-bin/astalavista/robot?srch=" & SearchString & "&project=robot&gfx=robot"
  Case 2
    Execute "http://www.ask.com/main/askJeeves.asp?ask=" & SearchString & "&origin=&qSource=0&site_name=Jeeves&metasearch=yes"
  Case 3
    Execute "http://www.excite.com/search.gw?search=" & SearchString & "&trace=2"
  Case 4
    Execute "http://www.hotbot.com/?MT=" & SearchString & "&SM=MC&DV=0&LG=any&DC=10&DE=2&_v=2&OPs=MDRTP"
  Case 5
    Execute "http://infoseek.go.com/Titles?qt=" & SearchString & "&col=WW&sv=IS&lk=noframes&svx=home_searchbox"
  Case 6
    Execute "http://search.msn.com/spbasic.htm?MT=" & SearchString
  Case 7
    Execute "http://www.lycos.com/cgi-bin/pursuit?query=" & SearchString & "&cat=dir"
  Case 8
    Execute "http://magellan.excite.com/search.gw?search=" & SearchString & "&c=web&look=magellan"
  Case 9
    Execute "http://www.metacrawler.com/cgi-bin/nph-metaquery.p?general=" & SearchString
  Case 10
    Execute "http://www.webcrawler.com/cgi-bin/WebQuery?" & SearchString
  Case 11
    Execute "http://av.yahoo.com/bin/search?p=" & SearchString
  End Select
End Function
