Attribute VB_Name = "Resize"
Public Type ctrObj
       Name As String
       Index As Long
       Parrent As String
       Top As Long
       left As Long
       Height As Long
       Width As Long
       ScaleHeight As Long
       ScaleWidth As Long
End Type

Private FormRecord() As ctrObj
Private ControlRecord() As ctrObj
Private bRunning As Boolean
Private MaxForm As Long
Private MaxControl As Long
Private Function ActualPos(plLeft As Long) As Long
              If plLeft < 0 Then
                     ActualPos = plLeft + 75000
              Else
                     ActualPos = plLeft
              End If
End Function

Private Function FindForm(pfrmIn As Form) As Long
       Dim I As Long
       FindForm = -1

              If MaxForm > 0 Then

                            For I = 0 To (MaxForm - 1)

                                          If FormRecord(I).Name = pfrmIn.Name Then
                                                 FindForm = I
                                                 Exit Function
                                          End If

                            Next I

              End If

End Function


Private Function AddForm(pfrmIn As Form) As Long

       Dim FormControl As Control
       Dim I As Long
       ReDim Preserve FormRecord(MaxForm + 1)

              FormRecord(MaxForm).Name = pfrmIn.Name

                            FormRecord(MaxForm).Top = pfrmIn.Top

                                          FormRecord(MaxForm).left = pfrmIn.left

                                                        FormRecord(MaxForm).Height = pfrmIn.Height

                                                                      FormRecord(MaxForm).Width = pfrmIn.Width
                                                                                    FormRecord(MaxForm).ScaleHeight = pfrmIn.ScaleHeight

                                                                                                  FormRecord(MaxForm).ScaleWidth = pfrmIn.ScaleWidth
                                                                                                         AddForm = MaxForm
                                                                                                         MaxForm = MaxForm + 1

                                                                                                                For Each FormControl In pfrmIn
                                                                                                                       I = FindControl(FormControl, pfrmIn.Name)

                                                                                                                              If I < 0 Then
                                                                                                                                     I = AddControl(FormControl, pfrmIn.Name)
                                                                                                                              End If

                                                                                                                Next FormControl

                                                                                                  End Function


Private Function FindControl(inControl As Control, inName As String) As Long

       Dim I As Long
       FindControl = -1

              For I = 0 To (MaxControl - 1)

                            If ControlRecord(I).Parrent = inName Then
                                          If ControlRecord(I).Name = inControl.Name Then
                                                 On Error Resume Next

                                                        If ControlRecord(I).Index = inControl.Index Then
                                                               FindControl = I
                                                               Exit Function
                                                        End If

                                                 On Error GoTo 0
                                          End If

                            End If

              Next I

End Function


Private Function AddControl(inControl As Control, inName As String) As Long

       ReDim Preserve ControlRecord(MaxControl + 1)
       On Error Resume Next
       ControlRecord(MaxControl).Name = inControl.Name
       ControlRecord(MaxControl).Index = inControl.Index
       ControlRecord(MaxControl).Parrent = inName

              If TypeOf inControl Is line Then
                     ControlRecord(MaxControl).Top = inControl.Y1
                     ControlRecord(MaxControl).left = ActualPos(inControl.X1)
                     ControlRecord(MaxControl).Height = inControl.Y2
                     ControlRecord(MaxControl).Width = ActualPos(inControl.X2)
              Else
                     ControlRecord(MaxControl).Top = inControl.Top
                     ControlRecord(MaxControl).left = ActualPos(inControl.left)
                     ControlRecord(MaxControl).Height = inControl.Height
                     ControlRecord(MaxControl).Width = inControl.Width
              End If

       inControl.IntegralHeight = False
       On Error GoTo 0
       AddControl = MaxControl
       MaxControl = MaxControl + 1
End Function


Private Function PerWidth(pfrmIn As Form) As Long

       Dim I As Long
       I = FindForm(pfrmIn)

              If I < 0 Then
                     I = AddForm(pfrmIn)
              End If

       PerWidth = (pfrmIn.ScaleWidth * 100) \ FormRecord(I).ScaleWidth
End Function


Private Function PerHeight(pfrmIn As Form) As Single

       Dim I As Long
       I = FindForm(pfrmIn)

              If I < 0 Then
                     I = AddForm(pfrmIn)
              End If

       PerHeight = (pfrmIn.ScaleHeight * 100) \ FormRecord(I).ScaleHeight
End Function


Private Sub ResizeControl(inControl As Control, pfrmIn As Form)

       On Error Resume Next
       Dim I As Long
       Dim widthfactor As Single, heightfactor As Single
       Dim minFactor As Single
       Dim yRatio, xRatio, lTop, lLeft, lWidth, lHeight As Long
       yRatio = PerHeight(pfrmIn)
       xRatio = PerWidth(pfrmIn)
       I = FindControl(inControl, pfrmIn.Name)

              If inControl.left < 0 Then
                     lLeft = CLng(((ControlRecord(I).left * xRatio) \ 100) - 75000)
              Else
                     lLeft = CLng((ControlRecord(I).left * xRatio) \ 100)
              End If

       lTop = CLng((ControlRecord(I).Top * yRatio) \ 100)
       lWidth = CLng((ControlRecord(I).Width * xRatio) \ 100)
       lHeight = CLng((ControlRecord(I).Height * yRatio) \ 100)
              If TypeOf inControl Is line Then

                            If inControl.X1 < 0 Then
                                   inControl.X1 = CLng(((ControlRecord(I).left * xRatio) \ 100) - 75000)
                            Else
                                   inControl.X1 = CLng((ControlRecord(I).left * xRatio) \ 100)
                            End If

                     inControl.Y1 = CLng((ControlRecord(I).Top * yRatio) \ 100)

                            If inControl.X2 < 0 Then
                                   inControl.X2 = CLng(((ControlRecord(I).Width * xRatio) \ 100) - 75000)
                            Else
                                   inControl.X2 = CLng((ControlRecord(I).Width * xRatio) \ 100)
                            End If

                     inControl.Y2 = CLng((ControlRecord(I).Height * yRatio) \ 100)
              Else
                     inControl.Move lLeft, lTop, lWidth, lHeight
                     inControl.Move lLeft, lTop, lWidth
                     inControl.Move lLeft, lTop
              End If

End Sub

Public Sub ResizeForm(pfrmIn As Form)

       Dim FormControl As Control
       Dim isVisible As Boolean
       Dim StartX, StartY, MaxX, MaxY As Long
       Dim bNew As Boolean

              If Not bRunning Then
                     bRunning = True

                            If FindForm(pfrmIn) < 0 Then
                                   bNew = True
                            Else
                                   bNew = False
                            End If


                            If pfrmIn.Top < 30000 Then
                                   isVisible = pfrmIn.Visible
                                   On Error Resume Next

                                          If Not pfrmIn.MDIChild Then
                                                 On Error GoTo 0
                                                 '     ' pfrmIn.Visible = False
                                          Else

                                                        If bNew Then
                                                               StartY = pfrmIn.Height
                                                               StartX = pfrmIn.Width
                                                               On Error Resume Next

                                                                      For Each FormControl In pfrmIn

                                                                                    If FormControl.left + FormControl.Width + 200 > MaxX Then
                                                                                           MaxX = FormControl.left + FormControl.Width + 200
                                                                                    End If


                                                                                    If FormControl.Top + FormControl.Height + 500 > MaxY Then
                                                                                           MaxY = FormControl.Top + FormControl.Height + 500
                                                                                    End If


                                                                                    If FormControl.X1 + 200 > MaxX Then
                                                                                           MaxX = FormControl.X1 + 200
                                                                                    End If


                                                                                    If FormControl.Y1 + 500 > MaxY Then
                                                                                           MaxY = FormControl.Y1 + 500
                                                                                    End If

                                                                                    If FormControl.X2 + 200 > MaxX Then
                                                                                           MaxX = FormControl.X2 + 200
                                                                                    End If


                                                                                    If FormControl.Y2 + 500 > MaxY Then
                                                                                           MaxY = FormControl.Y2 + 500
                                                                                    End If

                                                                      Next FormControl

                                                               On Error GoTo 0
                                                               pfrmIn.Height = MaxY
                                                               pfrmIn.Width = MaxX
                                                        End If

                                                 On Error GoTo 0
                                          End If


                                          For Each FormControl In pfrmIn
                                                 ResizeControl FormControl, pfrmIn
                                          Next FormControl

                                   On Error Resume Next

                                          If Not pfrmIn.MDIChild Then
                                                 On Error GoTo 0
                                                 pfrmIn.Visible = isVisible
                                          Else

                                                        If bNew Then
                                                               pfrmIn.Height = StartY
                                                               pfrmIn.Width = StartX

                                                                      For Each FormControl In pfrmIn
                                                                             ResizeControl FormControl, pfrmIn
                                                                      Next FormControl

                                                        End If

                                          End If

                                   On Error GoTo 0
                            End If

                     bRunning = False
              End If

End Sub


Public Sub SaveFormPosition(pfrmIn As Form)

       Dim I As Long

              If MaxForm > 0 Then

                            For I = 0 To (MaxForm - 1)

                                          If FormRecord(I).Name = pfrmIn.Name Then

                                                        FormRecord(I).Top = pfrmIn.Top

                                                                      FormRecord(I).left = pfrmIn.left

                                                                                    FormRecord(I).Height = pfrmIn.Height

                                                                                                  FormRecord(I).Width = pfrmIn.Width
                                                                                                         Exit Sub
                                                                                                  End If

                                                                                    Next I

                                                                             AddForm (pfrmIn)
                                                                      End If

                                                        End Sub


Public Sub RestoreFormPosition(pfrmIn As Form)

       Dim I As Long
              If MaxForm > 0 Then

                            For I = 0 To (MaxForm - 1)

                                          If FormRecord(I).Name = pfrmIn.Name Then

                                                        If FormRecord(I).Top < 0 Then
                                                               pfrmIn.WindowState = 2
                                                        ElseIf FormRecord(I).Top < 30000 Then
                                                               pfrmIn.WindowState = 0
                                                               pfrmIn.Move FormRecord(I).left, FormRecord(I).Top, FormRecord(I).Width, FormRecord(I).Height
                                                        Else
                                                               pfrmIn.WindowState = 1
                                                        End If

                                                 Exit Sub
                                          End If

                            Next I

              End If

End Sub
