' put in sheet code, not in module or workbook
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Address = Me.Range("E20").Address Then
        CalendarForm.Show vbModal
    End If
End Sub
'-------------------------------------------------------------------------------

Option Explicit

Private currentYear  As Integer
Private currentMonth As Integer

Private Sub UserForm_Initialize()
    currentYear = Year(Date)
    currentMonth = Month(Date)
    DrawCalendar
End Sub

Private Sub cmdPrev_Click()
    currentMonth = currentMonth - 1
    If currentMonth = 0 Then
        currentMonth = 12
        currentYear = currentYear - 1
    End If
    DrawCalendar
End Sub

Private Sub cmdNext_Click()
    currentMonth = currentMonth + 1
    If currentMonth = 13 Then
        currentMonth = 1
        currentYear = currentYear + 1
    End If
    DrawCalendar
End Sub

Private Sub DrawCalendar()
    Dim firstOfMonth As Date
    Dim startDay    As Integer, daysInMonth As Integer
    Dim btn As MSForms.CommandButton
    Dim i   As Integer, idx As Integer
    Dim thisDate As Date

    firstOfMonth = DateSerial(currentYear, currentMonth, 1)
    lblMonthYear.Caption = Format(firstOfMonth, "mmmm yyyy")
    startDay = Weekday(firstOfMonth, vbSunday)
    daysInMonth = Day(DateSerial(currentYear, currentMonth + 1, 0))

    ' Clear all day buttons
    For i = 1 To 42
        Set btn = Me.Controls("cmdDay" & i)
        btn.Caption = ""
        btn.Enabled = False
    Next i

    ' Fill in days up to today only
    idx = startDay
    For i = 1 To daysInMonth
        Set btn = Me.Controls("cmdDay" & idx)
        thisDate = DateSerial(currentYear, currentMonth, i)
        If thisDate <= Date Then
            btn.Caption = i
            btn.Enabled = True
        End If
        idx = idx + 1
    Next i

    ' Disable Next if we're at (or beyond) the current month
    With cmdNext
        If currentYear > Year(Date) _
        Or (currentYear = Year(Date) And currentMonth >= Month(Date)) Then
            .Enabled = False
        Else
            .Enabled = True
        End If
    End With

    ' (Optionally) disable Prev if you never want to go before today’s month
    ' With cmdPrev
    '     If currentYear < Year(Date) _
    '     Or (currentYear = Year(Date) And currentMonth <= Month(Date)) Then
    '         .Enabled = False
    '     Else
    '         .Enabled = True
    '     End If
    ' End With
End Sub

' Day-click handlers
Private Sub cmdDay1_Click():  Day_Click 1:  End Sub
Private Sub cmdDay2_Click():  Day_Click 2:  End Sub
Private Sub cmdDay3_Click():  Day_Click 3:  End Sub
Private Sub cmdDay4_Click():  Day_Click 4:  End Sub
Private Sub cmdDay5_Click():  Day_Click 5:  End Sub
Private Sub cmdDay6_Click():  Day_Click 6:  End Sub
Private Sub cmdDay7_Click():  Day_Click 7:  End Sub
Private Sub cmdDay8_Click():  Day_Click 8:  End Sub
Private Sub cmdDay9_Click():  Day_Click 9:  End Sub
Private Sub cmdDay10_Click(): Day_Click 10: End Sub
Private Sub cmdDay11_Click(): Day_Click 11: End Sub
Private Sub cmdDay12_Click(): Day_Click 12: End Sub
Private Sub cmdDay13_Click(): Day_Click 13: End Sub
Private Sub cmdDay14_Click(): Day_Click 14: End Sub
Private Sub cmdDay15_Click(): Day_Click 15: End Sub
Private Sub cmdDay16_Click(): Day_Click 16: End Sub
Private Sub cmdDay17_Click(): Day_Click 17: End Sub
Private Sub cmdDay18_Click(): Day_Click 18: End Sub
Private Sub cmdDay19_Click(): Day_Click 19: End Sub
Private Sub cmdDay20_Click(): Day_Click 20: End Sub
Private Sub cmdDay21_Click(): Day_Click 21: End Sub
Private Sub cmdDay22_Click(): Day_Click 22: End Sub
Private Sub cmdDay23_Click(): Day_Click 23: End Sub
Private Sub cmdDay24_Click(): Day_Click 24: End Sub
Private Sub cmdDay25_Click(): Day_Click 25: End Sub
Private Sub cmdDay26_Click(): Day_Click 26: End Sub
Private Sub cmdDay27_Click(): Day_Click 27: End Sub
Private Sub cmdDay28_Click(): Day_Click 28: End Sub
Private Sub cmdDay29_Click(): Day_Click 29: End Sub
Private Sub cmdDay30_Click(): Day_Click 30: End Sub
Private Sub cmdDay31_Click(): Day_Click 31: End Sub
Private Sub cmdDay32_Click(): Day_Click 32: End Sub
Private Sub cmdDay33_Click(): Day_Click 33: End Sub
Private Sub cmdDay34_Click(): Day_Click 34: End Sub
Private Sub cmdDay35_Click(): Day_Click 35: End Sub
Private Sub cmdDay36_Click(): Day_Click 36: End Sub
Private Sub cmdDay37_Click(): Day_Click 37: End Sub
Private Sub cmdDay38_Click(): Day_Click 38: End Sub
Private Sub cmdDay39_Click(): Day_Click 39: End Sub
Private Sub cmdDay40_Click(): Day_Click 40: End Sub
Private Sub cmdDay41_Click(): Day_Click 41: End Sub
Private Sub cmdDay42_Click(): Day_Click 42: End Sub

Private Sub Day_Click(idx As Integer)
    Dim d As Integer
    d = Val(Me.Controls("cmdDay" & idx).Caption)
    If d > 0 Then
        ThisWokbook.Sheets("VSF").Range("E20").Value = DateSerial(currentYear, currentMonth, d)
        Unload Me
    End If
End Sub

