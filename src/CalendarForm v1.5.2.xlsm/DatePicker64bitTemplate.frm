VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker64bitTemplate 
   Caption         =   "DatePicker64bitTemplate"
   ClientHeight    =   1428
   ClientLeft      =   -399
   ClientTop       =   -1750
   ClientWidth     =   2800
   OleObjectBlob   =   "DatePicker64bitTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePicker64bitTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InitialDate As Date

Private Type ScreenPosition
    Top As Single
    Left As Single
End Type

Private controlHooks As Collection

Private Function GetDateSeparatorPosition(SeparatorRank As Byte) As Byte
    
    Dim DateSeparatorCount As Byte
    DateSeparatorCount = 0
    
    Dim CurrentCharPosition As Byte
    For CurrentCharPosition = 1 To Len(DATE_MASK)
        Dim CurrentChar As String
        CurrentChar = Mid(DATE_MASK, 1, 1)
        
        If CurrentChar = DATE_SEPARATOR Then
            DateSeparatorCount = DateSeparatorCount + 1
            
            If DateSeparatorCount = SeparatorRank Then
                GetDateSeparatorPosition = CurrentCharPosition
                Exit Function
            End If
        End If
    Next
    
    GetDateSeparatorPosition = -1
    
End Function

Private Sub CalendarImageLabel_Click()
    UpdateDateFromLabelClick CalendarImageLabel
End Sub

Sub UpdateDateFromLabelClick(CalendarLabel As MSForms.Label)

    Dim txt As MSForms.TextBox
    Set txt = GetTextBoxUnderLabel(CalendarLabel)

    If txt Is Nothing Then
        MsgBox "No matching TextBox found under Label1."
        Exit Sub
    End If
    
    UpdateDateWithCalendar txt
    
End Sub

Function GetTextBoxUnderLabel(lbl As MSForms.Control) As MSForms.TextBox
    
    Dim lblLeft As Double, lblTop As Double, lblRight As Double, lblBottom As Double
    lblLeft = lbl.Left
    lblTop = lbl.Top
    lblRight = lbl.Left + lbl.Width
    lblBottom = lbl.Top + lbl.Height
    
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
    
        If TypeName(ctrl) = "TextBox" Then
        
            Dim txtLeft As Double, txtTop As Double, txtRight As Double, txtBottom As Double
            
            txtLeft = ctrl.Left
            txtTop = ctrl.Top
            txtRight = ctrl.Left + ctrl.Width
            txtBottom = ctrl.Top + ctrl.Height

            ' Check if the label overlaps the textbox
            If Not (lblRight < txtLeft Or lblLeft > txtRight Or _
                    lblBottom < txtTop Or lblTop > txtBottom) Then
                Set GetTextBoxUnderLabel = ctrl
                Exit Function
            End If
        End If
    Next ctrl

    ' If no match found
    Set GetTextBoxUnderLabel = Nothing
    
End Function

Sub UpdateDateWithCalendar(txtDate As MSForms.TextBox)
    
    Dim InitialTextDate As String
    InitialTextDate = txtDate.Text
    
    Dim CalendarTopLeftPosition As ScreenPosition
    CalendarTopLeftPosition = GetPopupPosition(txtDate, Me)
    
    Dim InitialDate As Date
    If IsDate(InitialTextDate) Then
        InitialDate = CDate(InitialTextDate)
    Else
        InitialDate = Date
    End If
    
    Dim DateSelected As Date
    DateSelected = calendarForm.GetDate(InitialDate, Monday, , , , , True, False, True, _
            FirstFourDays, CalendarTopLeftPosition.Top, CalendarTopLeftPosition.Left, TodayFontColor:=vbRed)
    
    If DateSelected = 0 Then
        Exit Sub
    End If
    
    Dim DateFormat As String
    DateFormat = GetDateFormat()
    
    ' Force output the same as the date format
    txtDate.Text = Format(DateSelected, DateFormat)
    
End Sub

' ParentUserForm must be an object to allow getting the window position .Top and .Left, not available in MSForms.UserForm
Private Function GetPopupPosition( _
            ctrl As MSForms.Control, _
            ParentUserForm As Object) As ScreenPosition
    
    Const Margin As Single = 5
    Const CaptionHeigh As Single = 20

    Dim pos As ScreenPosition
    pos.Left = ParentUserForm.Left + ctrl.Left + ctrl.Width + Margin
    pos.Top = ParentUserForm.Top + CaptionHeigh + ctrl.Top

    GetPopupPosition = pos
End Function

Private Sub DatePicker64bitTextBox_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateDate Me.ActiveControl
End Sub

Private Sub ValidateDate(txtDate As MSForms.TextBox)

    If Not IsDate(txtDate.Text) Then
        MsgBox "The date '" & txtDate.Text & "' is not valid. Reverting to " & txtDate.BoundValue
        txtDate.Value = txtDate.BoundValue
    Else
        txtDate.Value = txtDate.Text
    End If
    
End Sub

Private Sub SpinButton1_SpinDown()
    SpinDate1Day Me.ActiveControl, -1
End Sub

Private Sub SpinButton1_SpinUp()
    SpinDate1Day Me.ActiveControl, 1
End Sub

Private Sub SpinDate1Day( _
            SpinButton As MSForms.SpinButton, _
            DeltaDay As Single)
        
    Dim txt As MSForms.TextBox
    Set txt = GetTextBoxUnderLabel(SpinButton)
    
    If txt Is Nothing Then
        MsgBox "No matching TextBox found under Label1."
        Exit Sub
    End If
    
    txt.Value = CStr(CDate(txt.Value) + DeltaDay)
    
End Sub

Private Sub DatePicker64bitTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TreatDateWithKeyboardEntry Me.ActiveControl, KeyCode
End Sub

Sub TreatDateWithKeyboardEntry( _
            txtDate As MSForms.TextBox, _
            KeyCode As MSForms.ReturnInteger)
    
    ' Keys with allowed standard behavior
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome _
                Or KeyCode = vbKeyControl Then
        Exit Sub
    End If
    
    ' Exit TextBox
    If KeyCode = vbKeyTab Or KeyCode = vbKeyTab Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
            KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        Exit Sub
    End If
    
    If txtDate.SelLength > 1 Then
        DeleteSelectedText txtDate
        
        ' As text already deleted, stop execution
        If KeyCode = vbKeyBack Then
            txtDate.SelStart = 0
            GoTo ExitProcedure
        ElseIf KeyCode = vbKeyDelete Then
            GoTo ExitProcedure
        End If
        
    End If
    
    If KeyCode = vbKeyBack Then
        DeleteLeftTextBox txtDate
    ElseIf KeyCode = vbKeyDelete Then
        DeleteRightTextBox txtDate
    ElseIf (KeyCode >= vbKey0 And KeyCode <= vbKey9) _
            Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) Then
        EditDateTextBox txtDate, KeyCode
    End If
    
ExitProcedure:
    
    ' Cancel any key update
    KeyCode = 0
    
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
    
    TextCursorPosition = TextCursorPosition + 1
    
    Dim NewDate As String
    NewDate = CurrentDate
    Mid(NewDate, TextCursorPosition, 1) = "_"
    
    CurrentDate = NewDate
    
ExitProcedure:

    DateDeleteRight = CurrentDate
    
End Function

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

' Replaces the current selected date with the mask.
' Imitates the deletion of the selected text
Private Sub DeleteSelectedText(txtDate As MSForms.TextBox)
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelStart
    
    With txtDate
        .SelText = Mid(DATE_MASK, .SelStart + 1, .SelLength)
    End With
    
    txtDate.SelStart = TextCursorPosition
    
End Sub

Private Sub DatePicker64bitTextBox_Enter()
    SetDateMask Me.ActiveControl
End Sub

Private Sub SetDateMask(txtDate As MSForms.TextBox)
    
    If Not IsDate(txtDate.Text) Then
        txtDate.Text = Date
    End If
    
    txtDate.SelStart = 0
    
End Sub

Private Sub UserForm_Click()

End Sub
