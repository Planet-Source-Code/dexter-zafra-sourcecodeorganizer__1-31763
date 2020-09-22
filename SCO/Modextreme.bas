Attribute VB_Name = "Modextreme"
Option Explicit
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function StopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function Extreme(Extremedzinez As Single)
Dim Extremedex As Single
Extremedex = GetTickCount
Do While GetTickCount < Extremedex + Extremedzinez * 500
DoEvents
Loop
End Function
Public Sub ExtremesoundFile(ByVal FileName As String, Optional ByVal Wait As Boolean = False)
  If Wait Then
    Call PlaySound(FileName, 0&, &H20000)
  Else
    Call PlaySound(FileName, 0&, &H1 Or &H20000)
  End If
End Sub




