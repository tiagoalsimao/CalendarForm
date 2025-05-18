Attribute VB_Name = "TestingModule"
Public Const DATE_MASK As String = "__/__/____"
Public Const DATE_SEPARATOR As String = "/"

Sub TestUserForm()
    DatePicker64bitTemplate.Show
End Sub

Sub BasicCalendar()
    dateVariable = calendarForm.GetDate
    If dateVariable <> 0 Then Range("H16") = dateVariable
End Sub


Sub AdvancedCalendar()
    dateVariable = calendarForm.GetDate( _
        SelectedDate:=Range("H34").Value, _
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
    dateVariable = calendarForm.GetDate( _
        SelectedDate:=Range("H61").Value, _
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

Private Sub TestDateEdit()

    Debug.Print DateEdit("__/__/____", 0, 1) = "1_/__/____"
    Debug.Print DateEdit("1_/__/____", 1, 2) = "12/__/____"
    
    Debug.Print DateEdit("__/__/____", 0, 5) = "05/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateEdit("3_/__/____", 1, 3) = "31/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateEdit("__/__/____", 0, 0) = "0_/__/____"
    Debug.Print DateEdit("0_/__/____", 1, 0) = "01/__/____"
    
    Debug.Print DateEdit("__/__/____", 2, 1) = "__/1_/____"
    Debug.Print DateEdit("__/__/____", 3, 1) = "__/1_/____"
    Debug.Print DateEdit("__/__/____", 3, 3) = "__/03/____"
    Debug.Print DateEdit("__/__/____", 3, 5) = "__/05/____"
    Debug.Print DateEdit("__/_3/____", 3, 5) = "__/05/____"
    Debug.Print DateEdit("__/03/____", 3, 5) = "__/05/____"
    Debug.Print DateEdit("__/0_/____", 4, 0) = "__/01/____"
    Debug.Print DateEdit("__/__/____", 4, 5) = "__/_5/____"
    Debug.Print DateEdit("__/33/____", 4, 5) = "__/12/____"
    Debug.Print DateEdit("__/__/____", 2, 4) = "__/04/____" 'adds leading zero as cannot have months above 12
    Debug.Print DateEdit("__/__/____", 3, 4) = "__/04/____" 'adds leading zero as cannot have months above 12
    
    Debug.Print DateEdit("__/__/____", 5, 1) = "__/__/1___"
    Debug.Print DateEdit("__/__/____", 6, 1) = "__/__/1___"
    Debug.Print DateEdit("__/__/1___", 7, 2) = "__/__/12__"
    Debug.Print DateEdit("__/__/12__", 8, 3) = "__/__/123_"
    Debug.Print DateEdit("__/__/123_", 9, 4) = "__/__/1234"
    Debug.Print DateEdit("__/__/1234", 10, 5) = "__/__/1234"
    
    Debug.Print DateEdit("31/12/2025", 10, 2) = "31/12/2025"
    
End Sub

Public Sub EditDateTextBox( _
            txtDate As MSForms.TextBox, _
            KeyCode As MSForms.ReturnInteger)
    
    Dim TextDate As String
    TextDate = txtDate.Text
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelStart
    
    Dim InputNumber As Byte
    InputNumber = GetDigitFromKeyCode(KeyCode)
    
    DateEdit TextDate, TextCursorPosition, InputNumber
    
    txtDate.Text = TextDate
    txtDate.SelStart = TextCursorPosition
    
End Sub

Function GetDigitFromKeyCode(KeyCode As MSForms.ReturnInteger) As Byte
    Select Case KeyCode
        Case vbKey0 To vbKey9
            GetDigitFromKeyCode = KeyCode - vbKey0
        Case vbKeyNumpad0 To vbKeyNumpad9
            GetDigitFromKeyCode = KeyCode - vbKeyNumpad0
        Case Else
            GetDigitFromKeyCode = -1 ' Not a digit
    End Select
End Function

' Processes any numerical Key
Public Function DateEdit( _
            ByRef CurrentDate As String, _
            ByRef TextCursorPosition As Byte, _
            NewChar As Byte) As String
    
    If TextCursorPosition > 9 Then GoTo ExitProcedure
    
    ' Only allow digits to be inserted
    If NewChar < 0 Or NewChar > 9 Then GoTo ExitProcedure
    
    ' Move one char to the rigth if Char Position falls in Date Separator
    If TextCursorPosition = 2 Or TextCursorPosition = 5 Then TextCursorPosition = TextCursorPosition + 1
    
    ' Replace the character at TextCursorPosition
    Dim NewDate As String
    NewDate = CurrentDate
    
    ' Convert Tens to Units if: Day over 4 and Month over 1. Eg. Day = "5_" -> "05"
    If (NewChar > 3 And TextCursorPosition = 0) Or (NewChar > 1 And TextCursorPosition = 3) Then
        Mid(NewDate, TextCursorPosition + 1, 1) = 0
        
        TextCursorPosition = TextCursorPosition + 1
    End If
    
    Mid(NewDate, TextCursorPosition + 1, 1) = CStr(NewChar)
    
    Dim DateArray As Variant
    DateArray = Split(NewDate, DATE_SEPARATOR)
    
    ' Fix Day Maximum
    Dim dayStr As String
    dayStr = DateArray(0)
    If IsNumeric(dayStr) Then
        dayStr = Format(WorksheetFunction.Max(1, WorksheetFunction.Min(31, dayStr)), "00")
    End If
    
    ' Fix Month Maximum
    Dim monthStr As String
    monthStr = DateArray(1)
    If IsNumeric(monthStr) Then
        monthStr = Format(WorksheetFunction.Max(1, WorksheetFunction.Min(12, monthStr)), "00")
    End If
    
    Dim yearStr As String
    yearStr = DateArray(2)
    
    CurrentDate = Join(Array(dayStr, monthStr, yearStr), "/")
    
    ' Move Text Curser one character right
    TextCursorPosition = TextCursorPosition + 1
    
ExitProcedure:
    
    DateEdit = CurrentDate
    
End Function

Private Sub TestDateDeleteLeft()

    Debug.Print DateDeleteLeft("1_/__/____", 1) = "__/__/____"
    Debug.Print DateDeleteLeft("12/__/____", 2) = "1_/__/____"
    
    Debug.Print DateDeleteLeft("05/__/____", 1) = "_5/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateDeleteLeft("31/__/____", 2) = "3_/__/____" 'adds leading zero as cannot have days above 31
    Debug.Print DateDeleteLeft("01/__/____", 2) = "0_/__/____"
    
    Debug.Print DateDeleteLeft("_1/__/____", 3) = "__/__/____"
    Debug.Print DateDeleteLeft("__/1_/____", 4) = "__/__/____"
    Debug.Print DateDeleteLeft("__/_3/____", 5) = "__/__/____"
    Debug.Print DateDeleteLeft("__/__/____", 5) = "__/__/____"
    Debug.Print DateDeleteLeft("__/01/____", 5) = "__/0_/____"
    Debug.Print DateDeleteLeft("__/12/____", 5) = "__/1_/____"
    Debug.Print DateDeleteLeft("__/34/____", 4) = "__/_4/____" 'adds leading zero as cannot have months above 12
    
    Debug.Print DateDeleteLeft("__/_1/____", 6) = "__/__/____"
    Debug.Print DateDeleteLeft("__/__/1___", 7) = "__/__/____"
    Debug.Print DateDeleteLeft("__/__/12__", 8) = "__/__/1___"
    Debug.Print DateDeleteLeft("__/__/123_", 9) = "__/__/12__"
    Debug.Print DateDeleteLeft("__/__/1234", 10) = "__/__/123_"
    
    Debug.Print DateDeleteLeft("1_/__/1234", 0) = "1_/__/1234"
    
    Debug.Print DateDeleteLeft("31/12/2025", 1) = "_1/12/2025"
    
End Sub

' Processes BackSpace Key
Public Sub DeleteLeftTextBox(ByRef txtDate As MSForms.TextBox)
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelStart
        
    Dim CurrentDate As String
    CurrentDate = txtDate.Text
    
    DateDeleteLeft CurrentDate, TextCursorPosition
    
    txtDate.Text = CurrentDate
    txtDate.SelStart = TextCursorPosition
    
End Sub

' Processes BackSpace Key
Public Function DateDeleteLeft( _
            ByRef CurrentDate As String, _
            ByRef TextCursorPosition As Byte) As String
    
    ' No Change on First Position
    If TextCursorPosition < 1 Then GoTo ExitProcedure
    
    ' Move one char to the right if Char Position falls in Date Separator
    If TextCursorPosition = 3 Or TextCursorPosition = 6 Then TextCursorPosition = TextCursorPosition - 1
    
    Dim EditedDate As String
    EditedDate = CurrentDate
    Mid(EditedDate, TextCursorPosition, 1) = "_"
    
    CurrentDate = EditedDate
    TextCursorPosition = WorksheetFunction.Max(0, TextCursorPosition - 1)
    
ExitProcedure:

    DateDeleteLeft = CurrentDate
    
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
