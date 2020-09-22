Attribute VB_Name = "dateProcs"
Option Explicit
'===========================================================
'Date Validation :: Added 18th November 2003 by Craig Myles
'===========================================================

' /* Function:      dateCheck
'  * Purpose:       Validates a Date.
'  * Inputs:        sUserDate - the Date in String format supplied by the user.
'  * Returns:       The duration between now and the Date supplied.
'  * Side-effects:  sUserDate is passed by reference, and is therefore modified
'  *                to contain the formatted Date.
'  * Author:        Craig Myles.
'  * Date:          18th November 2003
'  */
Public Function dateCheck(sUserDate As String, sCaption As String) As String
    '// Date attributes, held in String format
    Dim sDate As String
    Dim sDay As String
    Dim sMonth As String
    
    '// Date attributes, held in Integer format
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim lYear As Long
    
    On Local Error GoTo DateCheck_Err
        
    '// dateCheck is used mainly on TextBox_LostFocus(). Therefore, if the user
    '// decides to 'pass through' the text field then the error message is suppressed
    If Trim$(sUserDate) = "" Then
        GoTo DateCheck_Err
    End If
    
    '// Assign user value to local value
    sDate = sUserDate
        
    '// Extract the Integer and String representations of the Day and Month
    iDay = grabDay(sDate, sDay)
    iMonth = grabMonth(sDate, sMonth)
        
    '// If the User has supplied a Year, then extract this (the whole Date minus Day and Month).
    '// Otherwise use the current Year
    If Trim$(sDate) <> "" Then
        lYear = Val(sDate)
    Else
        lYear = Year(Now)
    End If
    
    '// Validate the Day according to the supplied Month
    Select Case iMonth
        Case 2 '// February. At most 28 Days, 29 in a leap year
            If IsLeapYear(lYear) Then  'Leap Year
                If iDay > 29 Then
                    MsgBox "Leap Year. " & sMonth & " has at most 29 days", , sCaption
                    GoTo DateCheck_Err
                End If
            ElseIf iDay > 28 Then 'Not a Leap Year
                MsgBox "Non-Leap Year. " & sMonth & " has at most 28 days", , sCaption
                GoTo DateCheck_Err
            End If
            
        Case 4, 6, 9, 11 '// 30 Days have September, April, June and November
            If iDay > 30 Then
                MsgBox sMonth & " has at most 30 days", , sCaption
                GoTo DateCheck_Err
            End If
    
        Case Else '// January, March, May, July, August, October, December
            If iDay > 31 Then
                MsgBox sMonth & " has at most 31 days", , sCaption
                GoTo DateCheck_Err
            End If
    End Select
        
    '// Reform the Date to a defacto standard
    sDate = sDay & "/" & sMonth & "/" & Format(CStr(lYear), "00")
    
    '// Allow Visual Basic to determine the validity of the newly reformed Date
    If Not IsDate(sDate) Then
        MsgBox "Invalid Date", vbInformation, sCaption
        '// The supplied Date is wiped, and the Duration is replaced with a warning message.
        '// In the case of this project the TextBox tooltip is assigned to dateCheck, whereas
        '// the actually Date is replaced by reference of sDate
        sUserDate = ""
        dateCheck = "Invalid Date"
        GoTo dateCheck_Exit
    Else
        '// Allow Visual Basic to format the Date to a 'Long Date'. This must be performed
        '// after customising the Date since the VB 'Format' command regrading Dates isn't
        '// very robust. For example, vb would convert '31/9/03' wrongly.
        '// Return the duration of time between now and the Date
        sUserDate = Format(sDate, "Long Date")
        dateCheck = calcDuration(sDate)
    End If
    
dateCheck_Exit:
    Exit Function
    
DateCheck_Err:
    sUserDate = ""
    dateCheck = ""
    Resume dateCheck_Exit
    
End Function

'// Thanks to Roger Gilchrist for 'IsLeapYear'
Private Function IsLeapYear(LngY As Long) As Boolean
    IsLeapYear = (LngY Mod 4 = 0 And LngY Mod 100 <> 0) Or (LngY Mod 100 = 0 And LngY Mod 400 = 0)
End Function

'// Extract the Day attribute from the supplied Date.
'// sDate = sDate - sDay
Private Function grabDay(sDate As String, sDay As String) As Integer
    Dim iDigit As Integer
    
    '// Immediately exit the function if no argument supplied
    If sDate <> "" Then
        '// Extracts the leading digit from sDate (shortening sDate as we go)
        '// and adds to sDay. This happens until a delimiter is encountered,
        '// or we have extracted at most two digits.
        For iDigit = 1 To 3
            If Len(sDate) > 0 Then
                If delimiterPresent(sDate, 1) Then
                    sDate = Right$(sDate, Len(sDate) - 1)
                    Exit For
                ElseIf iDigit < 3 Then
                    sDay = sDay & Left$(sDate, 1)
                    sDate = Right$(sDate, Len(sDate) - 1)
                End If
            End If
        Next
        grabDay = Val(sDay)
    End If
End Function

'// Extract the month attribute from the supplied Date.
'// sDate = sDate - sMonth
'// When calling this function we assume that grabDay has been called first.
Private Function grabMonth(sDate As String, sMonth As String) As Integer
    Dim iDigit As Integer
    
    '// Immediately exit the function if no argument supplied
    If sDate <> "" Then
        '// User might input the month in long format OR we might be formatting an already
        '// formatted Date! Therefore, we must obtain the text as opposed to digits
        If isLongMonth(sDate) Then
            sMonth = UCase$(Trim$(Left$(sDate, InStr(sDate, " "))))
            sDate = Right$(sDate, Len(sDate) - Len(sMonth))
        Else
            '// Extracts the leading digit from sDate (shortening sDate as we go)
            '// and adds to sMonth. This happens until a delimiter is encountered,
            '// or we have extracted at most two digits.
            For iDigit = 1 To 3
                If Len(sDate) > 0 Then
                    If delimiterPresent(sDate, 1) Then
                        sDate = Right$(sDate, Len(sDate) - 1)
                        Exit For
                    ElseIf iDigit < 3 Then
                        sMonth = sMonth & Left$(sDate, 1)
                        sDate = Right$(sDate, Len(sDate) - 1)
                    End If
                End If
            Next
        End If
        '// Determine the String format for the Month attribute of the suplpied Date.
        grabMonth = cvrtMonth(sMonth)
    End If
    
End Function

'// Determine whether the user has supplied the Month argument in long format
Private Function isLongMonth(ByVal sText As String) As Boolean
    '// Convert to lower case in order to allow a fair comparison.
    '// Ascii key could be used here instead.
    sText = Trim$(LCase(sText))
    If sText <> "" Then
        If InStr(sText, "january") Or InStr(sText, "february") Or InStr(sText, "march") _
         Or InStr(sText, "april") Or InStr(sText, "may") Or InStr(sText, "june") _
         Or InStr(sText, "july") Or InStr(sText, "august") Or InStr(sText, "september") _
         Or InStr(sText, "october") Or InStr(sText, "november") Or InStr(sText, "december") Then
            isLongMonth = True
        End If
    End If
End Function

'// Determine whether a delimter has been supplied at the specified pointer
Private Function delimiterPresent(ByVal sText As String, ByVal iDigit As Integer) As Boolean
    On Local Error Resume Next
    
    If sText <> "" And iDigit > 0 Then
        If Mid$(sText, iDigit, 1) = "/" Or Mid$(sText, iDigit, 1) = "." Or Mid$(sText, iDigit, 1) = "-" Or Mid$(sText, iDigit, 1) = " " Then
            delimiterPresent = True
        End If
    End If
End Function

'// Convert an Integer value representation of a Month to long format
Private Function cvrtMonth(sMonth As String) As Integer
    Dim iMonth As Integer
    
    On Local Error Resume Next
    
    sMonth = Trim$(sMonth)
    
    If sMonth <> "" Then
        iMonth = Val(sMonth)
        sMonth = Choose(iMonth, "January", "February", "March", "April", "May", _
                                "June", "July", "August", "September", "October", _
                                "November", "December")
        cvrtMonth = iMonth
    End If
End Function

'// Calculate the time duration from now until the specified Date.
Private Function calcDuration(ByVal sDate As String) As String
    Dim TotalDays As Integer
    Dim dDate As Date
    Dim sState As String
    
    Dim iDays As Integer
    Dim iMonths As Integer
    Dim iYears As Integer

    On Local Error Resume Next
    
    dDate = sDate
    
    TotalDays = DateValue(Now) - DateValue(dDate)
    
    If TotalDays = 0 Then
        calcDuration = "Present Day"
    Else
    
        If TotalDays < 0 Then
            sState = "in the future"
            iYears = Year(dDate) - Year(Now)
            iMonths = Month(dDate) - Month(Now)
            iDays = Day(dDate) - Day(Now)
        Else
            sState = "in the past"
            iYears = Year(Now) - Year(dDate)
            iMonths = Month(Now) - Month(dDate)
            iDays = Day(Now) - Day(dDate)
        End If
    
        If iDays < 0 Then iMonths = iMonths - 1
        
        If iMonths <= 0 Then
            iMonths = iMonths + 12
            iYears = iYears - 1
        End If
        
        If iMonths = 12 Then
            iMonths = 0
            iYears = iYears + 1
        End If
    
        calcDuration = iYears & " Year(s), " & iMonths & " Month(s), " & iDays & " Day(s) " & sState
    End If
End Function
