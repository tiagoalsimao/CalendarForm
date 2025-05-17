Attribute VB_Name = "TestingModule"
Sub TestUserForm()
    UserFormTest.Show
End Sub

Sub BasicCalendar()
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then Range("H16") = dateVariable
End Sub


Sub AdvancedCalendar()
    dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Range("H34").value, _
        FirstDayOfWeek:=Monday, _
        DateFontSize:=12, _
        TodayButton:=True, _
        OkayButton:=True, _
        ShowWeekNumbers:=True, _
        BackgroundColor:=RGB(243, 249, 251), _
        HeaderColor:=RGB(147, 205, 2221), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), _
        DateColor:=RGB(243, 249, 251), _
        DateFontColor:=RGB(31, 78, 120), _
        TrailingMonthFontColor:=RGB(155, 194, 230), _
        DateHoverColor:=RGB(223, 240, 245), _
        DateSelectedColor:=RGB(202, 223, 242), _
        SaturdayFontColor:=RGB(0, 176, 240), _
        SundayFontColor:=RGB(0, 176, 240), _
        TodayFontColor:=RGB(0, 176, 80))
    If dateVariable <> 0 Then Range("H34") = dateVariable
End Sub


Sub AdvancedCalendar2()
    dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Range("H61").value, _
        DateFontSize:=11, _
        TodayButton:=True, _
        BackgroundColor:=RGB(242, 248, 238), _
        HeaderColor:=RGB(84, 130, 53), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(226, 239, 218), _
        SubHeaderFontColor:=RGB(55, 86, 35), _
        DateColor:=RGB(242, 248, 238), _
        DateFontColor:=RGB(55, 86, 35), _
        SaturdayFontColor:=RGB(55, 86, 35), _
        SundayFontColor:=RGB(55, 86, 35), _
        TrailingMonthFontColor:=RGB(106, 163, 67), _
        DateHoverColor:=RGB(198, 224, 180), _
        DateSelectedColor:=RGB(169, 208, 142), _
        TodayFontColor:=RGB(255, 0, 0), _
        DateSpecialEffect:=fmSpecialEffectRaised)
    If dateVariable <> 0 Then Range("H61") = dateVariable
End Sub

Private Sub TestUpdateDate()
    Debug.Print UpdateDate("__/__/____", 1, 1) = "1_/__/____"
    Debug.Print UpdateDate("1_/__/____", 2, 2) = "12/__/____"
    
    Debug.Print UpdateDate("5_/__/____", 5, 1) = "05/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print UpdateDate("3_/__/____", 3, 2) = "31/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print UpdateDate("0_/__/____", 0, 2) = "01/__/____"
    
    Debug.Print UpdateDate("__/__/____", 1, 3) = "__/1_/____"
    Debug.Print UpdateDate("__/__/____", 3, 5) = "__/_3/____"
    Debug.Print UpdateDate("__/_3/____", 3, 5) = "__/_3/____"
    Debug.Print UpdateDate("__/03/____", 3, 5) = "__/03/____"
    Debug.Print UpdateDate("__/0_/____", 0, 5) = "__/01/____"
    Debug.Print UpdateDate("__/33/____", 3, 5) = "__/12/____"
    Debug.Print UpdateDate("__/__/____", 4, 4) = "__/04/____" 'adds leading zero as cannot have months above 12
    
    Debug.Print UpdateDate("__/__/____", 1, 6) = "__/__/1___"
    Debug.Print UpdateDate("__/__/____", 1, 7) = "__/__/1___"
    Debug.Print UpdateDate("__/__/1___", 2, 8) = "__/__/12__"
    Debug.Print UpdateDate("__/__/12__", 3, 9) = "__/__/123_"
    Debug.Print UpdateDate("__/__/123_", 4, 10) = "__/__/1234"
    Debug.Print UpdateDate("__/__/1234", 5, 11) = "__/__/1235"
    
    Debug.Print UpdateDate("31/12/2025", 2, 11) = "31/12/2022"
    
End Sub

Public Function UpdateDate( _
            ByRef CurrentDate As String, _
            ByRef CharPosition As Byte, _
            NewChar As Byte) As String
    
    ' Only allow digits to be inserted
    If NewChar < 0 Or NewChar > 9 Then
        UpdateDate = CurrentDate
        Exit Function
    End If
    
    ' Move one char to the rigth if Char Position falls in Date Separator
    If CharPosition = 3 Or CharPosition = 6 Then CharPosition = CharPosition + 1
    
    ' Change last position
    If CharPosition > 10 Then CharPosition = 10
    
    ' Replace the character at CharPosition
    Dim TempDate As String
    TempDate = CurrentDate
    
    ' Convert Tens to Units if: Day over 4 and Month over 1. Eg. Day = "5_" -> "05"
    If (NewChar > 3 And CharPosition = 1) Or (NewChar > 1 And CharPosition = 4) Then
        Mid(TempDate, CharPosition, 1) = 0
        
        CharPosition = CharPosition + 1
        Mid(TempDate, CharPosition, 1) = CStr(NewChar)
    Else
        Mid(TempDate, CharPosition, 1) = CStr(NewChar)
    End If
    
    ' Fix Day Maximum
    Dim dayStr As String
    dayStr = Mid(TempDate, 1, 2)
    If IsNumeric(dayStr) Then
        dayStr = Format(WorksheetFunction.Max(1, WorksheetFunction.Min(31, dayStr)), "00")
    End If
    
    ' Fix Month Maximum
    Dim monthStr As String
    monthStr = Mid(TempDate, 4, 2)
    If IsNumeric(monthStr) Then
        monthStr = Format(WorksheetFunction.Max(1, WorksheetFunction.Min(12, monthStr)), "00")
    End If
    
    Dim yearStr As String
    yearStr = Mid(TempDate, 7, 4)
    
    UpdateDate = Join(Array(dayStr, monthStr, yearStr), "/")
    
    CurrentDate = UpdateDate
    
End Function

Private Sub TestDateBackSpace()
    Debug.Print DateBackSpace("1_/__/____", 1) = "__/__/____"
    Debug.Print DateBackSpace("12/__/____", 2) = "1_/__/____"
    
    Debug.Print DateBackSpace("05/__/____", 1) = "_5/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateBackSpace("31/__/____", 2) = "3_/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateBackSpace("01/__/____", 2) = "0_/__/____"
    
    Debug.Print DateBackSpace("_1/__/____", 3) = "__/__/____"
    Debug.Print DateBackSpace("__/1_/____", 4) = "__/__/____"
    Debug.Print DateBackSpace("__/_3/____", 5) = "__/__/____"
    Debug.Print DateBackSpace("__/__/____", 5) = "__/__/____"
    Debug.Print DateBackSpace("__/01/____", 5) = "__/0_/____"
    Debug.Print DateBackSpace("__/12/____", 5) = "__/1_/____"
    Debug.Print DateBackSpace("__/34/____", 4) = "__/_4/____" 'adds leading zero as cannot have months above 12
    
    Debug.Print DateBackSpace("__/_1/____", 6) = "__/__/____"
    Debug.Print DateBackSpace("__/__/1___", 7) = "__/__/____"
    Debug.Print DateBackSpace("__/__/12__", 8) = "__/__/1___"
    Debug.Print DateBackSpace("__/__/123_", 9) = "__/__/12__"
    Debug.Print DateBackSpace("__/__/1234", 10) = "__/__/123_"
    
    Debug.Print DateBackSpace("1_/__/1234", 0) = "1_/__/1234"
    
    Debug.Print DateBackSpace("31/12/2025", 1) = "_1/12/2025"
    
End Sub

Public Function DateBackSpace( _
            ByRef CurrentDate As String, _
            CharPosition As Byte) As String
            
    ' Move one char to the rigth if Char Position falls in Date Separator
    If CharPosition = 3 Or CharPosition = 6 Then CharPosition = CharPosition - 1
    
    ' Change last position
    If CharPosition < 1 Then
        DateBackSpace = CurrentDate
        Exit Function
    End If
    
    Dim TempDate As String
    TempDate = CurrentDate
    Mid(TempDate, CharPosition, 1) = "_"
    
    DateBackSpace = TempDate
    
    CurrentDate = DateBackSpace
    
End Function

Private Sub TestDeleteKey()

    Debug.Print DateDeleteRight("1_/__/____", 0) = "__/__/____"
    Debug.Print DateDeleteRight("12/__/____", 1) = "1_/__/____"
    
    Debug.Print DateDeleteRight("05/__/____", 0) = "_5/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateDeleteRight("31/__/____", 1) = "3_/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateDeleteRight("01/__/____", 1) = "0_/__/____"
    
    Debug.Print DateDeleteRight("__/1_/____", 2) = "__/__/____"
    Debug.Print DateDeleteRight("__/1_/____", 3) = "__/__/____"
    Debug.Print DateDeleteRight("__/_3/____", 4) = "__/__/____"
    Debug.Print DateDeleteRight("__/__/____", 4) = "__/__/____"
    Debug.Print DateDeleteRight("__/01/____", 4) = "__/0_/____"
    Debug.Print DateDeleteRight("__/12/____", 4) = "__/1_/____"
    Debug.Print DateDeleteRight("__/34/____", 3) = "__/_4/____" 'adds leading zero as cannot have months above 12
    
    Debug.Print DateDeleteRight("__/__/1___", 5) = "__/__/____"
    Debug.Print DateDeleteRight("__/__/1___", 6) = "__/__/____"
    Debug.Print DateDeleteRight("__/__/12__", 7) = "__/__/1___"
    Debug.Print DateDeleteRight("__/__/123_", 8) = "__/__/12__"
    Debug.Print DateDeleteRight("__/__/1234", 9) = "__/__/123_"
    
    Debug.Print DateDeleteRight("1_/__/1234", 10) = "1_/__/1234"
    
    Debug.Print DateDeleteRight("31/12/2025", 0) = "_1/12/2025"
    
End Sub

' Processes Delete Key
Public Sub DeleteRightTextBox(ByRef txtDate As MSForms.TextBox)
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelStart

    Dim CurrentDate As String
    CurrentDate = txtDate.Text

    DateDeleteRight CurrentDate, TextCursorPosition

    txtDate.Text = CurrentDate
    txtDate.SelStart = TextCursorPosition
    
End Sub

' Processes Delete Key
Public Function DateDeleteRight( _
            ByRef CurrentDate As String, _
            ByRef TextCursorPosition As Byte) As String
    
    ' No Change on Last Position
    If TextCursorPosition > 9 Then GoTo ExitProcedure
    
    ' Move one char to the right if Char Position falls in Date Separator
    If TextCursorPosition = 2 Or TextCursorPosition = 5 Then TextCursorPosition = TextCursorPosition + 1
    
    Dim NewDate As String
    NewDate = CurrentDate
    Mid(NewDate, TextCursorPosition + 1, 1) = "_"
    
    CurrentDate = NewDate
    
ExitProcedure:

    DateDeleteRight = CurrentDate
    
End Function
