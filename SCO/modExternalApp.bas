Attribute VB_Name = "modExternalApp"
Option Explicit
Public Const SND_ASYNC = &H1
Public UseSound As String
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetTickCount& Lib "kernel32" ()
Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpfilename As String) As Long

Public Sub email(emailAddress As String)
    Dim IE As Object
    Set IE = CreateObject("internetexplorer.application")
    IE.Visible = False
    IE.Navigate "Mailto:" & emailAddress
    IE.Quit
    Set IE = Nothing
End Sub
Public Sub webBrowse(urlAddress As String)
    Dim IE As Object
    Set IE = CreateObject("internetexplorer.application")
    IE.Visible = True
    IE.Navigate urlAddress
End Sub

Public Sub TimeOut(Duration)
Dim Starttime As Long
Starttime = Timer
    Do While Timer - Starttime > Duration
      DoEvents
    Loop
End Sub
