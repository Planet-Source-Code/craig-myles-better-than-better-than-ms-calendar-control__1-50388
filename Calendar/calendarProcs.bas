Attribute VB_Name = "calendarProcs"
Option Explicit

Public Function GetOneDate(iTop As Integer, iLeft As Integer) As String

    On Local Error Resume Next
    
    Load frmCalendar
    
    PositionPickListForm frmCalendar, iTop + 60, iLeft + 140
    
    With frmCalendar
        .Caption = "Select a Date"
        Screen.MousePointer = vbDefault
        .Show vbModal
        DoEvents
        Select Case .Tag
            Case "XXX"
                GetOneDate = ""
            Case Else
                GetOneDate = .Tag
        End Select
    End With
    
    Unload frmCalendar
    
End Function

Public Sub PositionPickListForm(X As Form, iTop As Integer, iLeft As Integer)

    On Local Error Resume Next
    
    If iTop <> 0 Or iLeft <> 0 Then
        If iTop <> 0 Then
            If iTop + X.Height > Screen.Height Then
                X.Top = iTop - X.Height
            Else
                X.Top = iTop
            End If
        End If
        
        If iLeft <> 0 Then
            If iLeft + X.Width > Screen.Width Then
                X.Left = iLeft - X.Width
            Else
                X.Left = iLeft
            End If
        End If
    Else
        X.Left = 0.5 * (Screen.Width - X.Width)
        X.Top = 0.5 * (Screen.Height - X.Height)
    End If

End Sub

