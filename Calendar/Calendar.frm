VERSION 5.00
Begin VB.Form frmCalendar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Month Calendar"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   ControlBox      =   0   'False
   HelpContextID   =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4200
   Begin VB.ComboBox cboYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox picMover 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox F 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "2003"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cboMonth 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picCal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2745
      ScaleWidth      =   3930
      TabIndex        =   0
      Top             =   480
      Width           =   3962
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   160
      Width           =   3855
   End
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dX As Integer
Dim Dy As Integer
Dim mfLoaded As Boolean

'// Double click functionality added by Craig Myles
Dim dClickStart As Double
Dim dDifference As Double

Const cnLtGrey = &HC0C0C0
Const cnDkGrey = &H808080
Const cnBlack = &H0&

Dim mnDate As Long


Private Sub cboMonth_Click()

    lblHeader.Caption = cboMonth.Text & " " & cboYear.Text
    picCal.Cls
    DrawMonthHeading
    DrawDays

End Sub

Private Sub cboYear_Click()
    lblHeader.Caption = cboMonth.Text & " " & cboYear.Text
    picCal.Cls
    DrawMonthHeading
    DrawDays

End Sub

Private Sub cmdClose_Click()
    Me.Tag = "XXX"
    Me.Hide
End Sub

Private Sub cmdSelect_Click()

    Me.Tag = Format$(mnDate, "Long Date")
    Me.Hide
    
End Sub

Private Sub F_LostFocus()
    
    picCal.Cls
    DrawMonthHeading
    DrawDays

End Sub

Private Sub Form_Activate()
    If Not mfLoaded Then
        mfLoaded = True
        GetScaleFactor
        DrawMonthHeading
        DrawDays
    End If
End Sub

Private Sub Form_Load()
    Dim iYear As Integer
    
    For iYear = 1900 To 2100
        cboYear.AddItem iYear
    Next
    
    cboYear.ListIndex = Year(Now) - 1900
    
    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "August"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"
    cboMonth.ListIndex = Month(Now) - 1
        
    lblHeader.Caption = Format$(Now, "mmmm yyyy")
    
End Sub
Private Sub GetScaleFactor()

    dX = picCal.Width / 7
    Dy = picCal.Height / 7
    picMover.Height = Dy
    picMover.Width = dX
    
End Sub
Private Sub DrawMonthHeading()

    Dim sText As String
    Dim i As Integer
    Dim iMonth As Integer
    Dim X1 As Integer
    Dim Y1 As Integer
    
    On Local Error Resume Next
    
    iMonth = cboMonth.ListIndex + 1
    
    X1 = picCal.Width / 7
    Y1 = picCal.Height / 7
    
    picCal.Line (0, 0)-(picCal.Width, Dy), cnLtGrey, BF
    
    picCal.ForeColor = cnDkGrey
    
    picCal.Line (X1, Y1)-(X1, picCal.Height)
    picCal.Line (2 * X1, Y1)-(2 * X1, picCal.Height)
    picCal.Line (3 * X1, Y1)-(3 * X1, picCal.Height)
    picCal.Line (4 * X1, Y1)-(4 * X1, picCal.Height)
    picCal.Line (5 * X1, Y1)-(5 * X1, picCal.Height)
    picCal.Line (6 * X1, Y1)-(6 * X1, picCal.Height)
    
    picCal.Line (0, Y1)-(picCal.Width, Y1)
    picCal.Line (0, 2 * Y1)-(picCal.Width, 2 * Y1)
    picCal.Line (0, 3 * Y1)-(picCal.Width, 3 * Y1)
    picCal.Line (0, 4 * Y1)-(picCal.Width, 4 * Y1)
    picCal.Line (0, 5 * Y1)-(picCal.Width, 5 * Y1)
    picCal.Line (0, 6 * Y1)-(picCal.Width, 6 * Y1)
    
    picCal.ForeColor = cnBlack
    
    picCal.FontBold = True
    
    For i = 1 To 7
        sText = Choose(i, "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
        picCal.CurrentY = 0.5 * (Dy - picCal.TextHeight(sText))
        picCal.CurrentX = ((i - 1) * dX) + 0.5 * (dX - picCal.TextWidth(sText))
        picCal.Print sText
    Next
    
    picCal.FontBold = False

End Sub
Private Sub DrawDays()

    Dim nDate As Long
    Dim i As Long
    Dim iLast As Integer
    Dim iRow As Integer
    Dim iMonth As Integer
    Dim sText As String
    
    On Local Error Resume Next
    
    If Not mfLoaded Then Exit Sub
    
    'Get first day in year
    iMonth = cboMonth.ListIndex + 1
    sText = cboYear.Text
    nDate = DateValue("01/" & Format$(iMonth, "0") & "/" & sText)
    
    GetLastDay iMonth, iLast
    iRow = 1
    
    For i = nDate To nDate + iLast - 1
        If Weekday(i) = vbSunday Then
            If i > nDate Then
                iRow = iRow + 1
            End If
        End If
        
        sText = Format$(Day(i), "0")
        picCal.CurrentY = (Dy * iRow) + (0.5 * (Dy - picCal.TextHeight(sText)))
        picCal.CurrentX = dX * (Weekday(i) - 1) + (0.5 * (dX - picCal.TextWidth(sText)))
        picCal.Print sText
    Next
    
End Sub
Private Sub GetLastDay(iMonth, iLast)

    Select Case iMonth
        '// 30 Days have September, April, June and November
        Case 4, 6, 9, 11
            iLast = 30
        '// Everything else has 31, with the exception of February
        Case 1, 3, 5, 7, 8, 10, 12
            iLast = 31
        '// February
        Case 2
         iLast = IIf(IsLeapYear(Val(cboYear.Text)), 29, 28)
  End Select
End Sub

Private Function IsLeapYear(LngY As Long) As Boolean
    IsLeapYear = (LngY Mod 4 = 0 And LngY Mod 100 <> 0) Or (LngY Mod 100 = 0 And LngY Mod 400 = 0)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmCalendar = Nothing
End Sub

Private Sub picCal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim iCol As Integer
    Dim iRow As Integer
    Dim i As Long
    Dim R As Integer
    Dim C As Integer
    Dim iMonth As Integer
    Dim iLast As Integer
    Dim nDate As Long
    Dim sText As String
    
    On Local Error Resume Next
    
    iMonth = cboMonth.ListIndex + 1
    picMover.Visible = False
    GetLastDay iMonth, iLast
    
    iCol = 7 * (X) \ picCal.Width + 1
    iRow = 7 * (Y) \ picCal.Height - 1
    If iRow < 0 Then Exit Sub
    
    nDate = DateValue("01/" & Format$(iMonth, "0") & "/" & cboYear.Text)
    
    R = 0
    
    For i = nDate To nDate + iLast - 1
    
        If Weekday(i) = vbSunday Then
            If i > nDate Then
                R = R + 1
            End If
        End If
        
        C = Weekday(i)
        If R = iRow And C = iCol Then
            mnDate = i
            picMover.Cls
            sText = Day(mnDate)
            
            picMover.Left = (picCal.Left + 20) + ((C - 1) * dX)
            picMover.Top = (picCal.Top + 20) + ((R + 1) * Dy)
            
            If C = 7 Then picMover.Width = dX - 40 Else picMover.Width = dX
            If R = 5 Then picMover.Height = Dy - 20 Else picMover.Height = Dy
            
            picMover.CurrentX = 0.5 * (picMover.Width - picCal.TextWidth(sText))
            picMover.CurrentY = 0.5 * (picMover.Height - picCal.TextHeight(sText))
            
            picMover.Print sText
            
            picMover.Visible = True
            Exit For
        End If
        
    Next
    dClickStart = Now
End Sub

Private Sub picMover_DblClick()
    cmdSelect_Click
End Sub

Private Sub picMover_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dDifference = Now
    dDifference = dDifference - dClickStart
    If dDifference = 0 Then
        cmdSelect_Click
    End If
End Sub
