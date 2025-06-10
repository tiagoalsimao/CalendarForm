VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatePicker64BitClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Hooked controls
Private WithEvents txtDate As MSForms.TextBox
Attribute txtDate.VB_VarHelpID = -1
Private WithEvents spnDate As MSForms.SpinButton
Attribute spnDate.VB_VarHelpID = -1
Private WithEvents lblDate As MSForms.Label
Attribute lblDate.VB_VarHelpID = -1

Private Type ObjectPosition
    Top As Single
    Left As Single
    Bottom As Single
    Right As Single
End Type

Private Const DATE_SEPARATOR As String = "/"

' Initialization method
Public Sub Initialize( _
            ByRef txtTextBox As MSForms.TextBox, _
            ByRef spnSpinButton As MSForms.SpinButton, _
            ByRef lblLabel As MSForms.Label)
    
    Set txtDate = txtTextBox
    Set spnDate = spnSpinButton
    Set lblDate = lblLabel
    
    SetInitialDate txtDate
    
End Sub

Public Property Get Value() As Date
    Value = CDate(txtDate.Value)
End Property

Public Property Let Value(InputValue As Date)
    txtDate.Value = CStr(InputValue)
End Property

Private Function IsValidChange() As Boolean
    IsValidChange = (txtDate.BoundValue = txtDate.Value)
End Function

' === Outside Events ===
Public Sub Enter()
    txtDate.SelStart = 0
End Sub

Public Sub Validate_BeforeUpdate(ByRef Cancel As MSForms.ReturnBoolean)
    ValidateDate txtDate, Cancel
End Sub

' === TextBox Events ===
Private Sub txtDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TreatDateWithKeyboardEntry txtDate, KeyCode
End Sub

' === SpinButton Events ===
Private Sub spnDate_SpinUp()
    ChangeDateBy1 spnDate, 1
End Sub

Private Sub spnDate_SpinDown()
    ChangeDateBy1 spnDate, -1
End Sub

Private Sub spnDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload spnDate.Parent
End Sub

' === Label Events ===
Private Sub lblDate_Click()
    UpdateDateFromLabelClick lblDate
End Sub

' From the Label Click, get TextBox below control and lauch CalendarForm
Private Sub UpdateDateFromLabelClick(CalendarLabel As MSForms.Label)
    
    If CalendarLabel Is Nothing Then
        MsgBox "Label is Nothing."
        Exit Sub
    End If
    
    Dim txt As MSForms.TextBox
    Set txt = GetTextBoxUnderControl(CalendarLabel)

    If txt Is Nothing Then
        MsgBox "No matching TextBox found under Label: " & CalendarLabel.Caption
        Exit Sub
    End If
    
    UpdateDateWithCalendar txt
    
End Sub

Private Function GetObjectPosition(ctrl As MSForms.Control) As ObjectPosition

    With GetObjectPosition
        .Left = ctrl.Left
        .Top = ctrl.Top
        .Right = ctrl.Left + ctrl.Width
        .Bottom = ctrl.Top + ctrl.Height
    End With
    
End Function

Private Function GetTextBoxUnderControl(lbl As MSForms.Control) As MSForms.TextBox
    
    Dim lblPosition As ObjectPosition
    lblPosition = GetObjectPosition(lbl)
    
    Dim ParentUserForm As Object
    Set ParentUserForm = lbl.Parent
    
    Dim ctrl As MSForms.Control
    For Each ctrl In ParentUserForm.Controls
    
        If TypeName(ctrl) = "TextBox" Then
        
            Dim TextBoxPosition As ObjectPosition
            TextBoxPosition = GetObjectPosition(ctrl)

            ' Check if the label overlaps the textbox
            If Not (lblPosition.Right < TextBoxPosition.Left Or lblPosition.Left > TextBoxPosition.Right Or _
                    lblPosition.Bottom < TextBoxPosition.Top Or lblPosition.Top > TextBoxPosition.Bottom) Then
                Set GetTextBoxUnderControl = ctrl
                Exit Function
            End If
        End If
    Next ctrl

    ' If no match found
    Set GetTextBoxUnderControl = Nothing
    
End Function

Private Sub UpdateDateWithCalendar(txtDate As MSForms.TextBox)
    
    Dim InitialTextDate As String
    InitialTextDate = txtDate.Text
    
    Dim CalendarTopLeftPosition As ObjectPosition
    CalendarTopLeftPosition = GetPopupPosition(txtDate)
    
    Dim InitialDate As Date
    If IsDate(InitialTextDate) Then
        InitialDate = CDate(InitialTextDate)
    Else
        InitialDate = Date
    End If
    
    Dim DateSelected As Date
    DateSelected = CalendarForm.GetDate(InitialDate, Monday, , , , , True, False, True, _
            FirstFourDays, CalendarTopLeftPosition.Top, CalendarTopLeftPosition.Left, TodayFontColor:=vbRed)
    
    If DateSelected = 0 Then
        Exit Sub
    End If

    txtDate.Text = Format(DateSelected)
    
End Sub

' ParentUserForm must be an object to allow getting the window position .Top and .Left, not available in MSForms.UserForm
Private Function GetPopupPosition(ctrl As MSForms.Control) As ObjectPosition
            
    Dim ParentUserForm As Object
    Set ParentUserForm = ctrl.Parent
    
    Const Margin As Single = 5
    Const CaptionHeigh As Single = 20

    Dim pos As ObjectPosition
    pos.Left = ParentUserForm.Left + ctrl.Left + ctrl.Width + Margin
    pos.Top = ParentUserForm.Top + CaptionHeigh + ctrl.Top
    
    GetPopupPosition = pos
    
End Function

Public Function HasValidDate() As Boolean
    HasValidDate = IsDate(txtDate.Text)
End Function

Private Sub ValidateDate( _
            txtDate As MSForms.TextBox, _
            ByRef Cancel As MSForms.ReturnBoolean)

    If Not HasValidDate Then
        MsgBox "The date '" & txtDate.Text & "' is not valid. Reverting to " & txtDate.BoundValue
        txtDate.Value = txtDate.BoundValue
        Cancel = True
    Else
        txtDate.Value = CStr(CDate(txtDate.Text))
    End If
    
End Sub

Private Sub ChangeDateBy1( _
            SpinButton As MSForms.SpinButton, _
            DeltaDay As Integer)
    
    Dim txt As MSForms.TextBox
    Set txt = GetTextBoxUnderControl(SpinButton)
    
    If txt Is Nothing Then
        MsgBox "No matching TextBox found under Label1."
        Exit Sub
    End If
    
    With txt
        Dim CurrentDate As Date
        CurrentDate = .Value
        
        Dim TextCursorPosition As Byte
        TextCursorPosition = .SelStart
    End With
    
    Dim NewDate As Date
    Select Case TextCursorPosition
        Case Is < 3
            NewDate = DateAdd("d", DeltaDay, CurrentDate)
        Case Is < 5
            If isEndOfMonth(CurrentDate) Then
                NewDate = DateAdd("m", DeltaDay, CurrentDate + 1) - 1
            Else
                NewDate = DateAdd("m", DeltaDay, CurrentDate)
            End If
        Case Else
            NewDate = DateAdd("yyyy", DeltaDay, CurrentDate)
    End Select
    
    With txt
        .Value = CStr(NewDate)
        .SelStart = TextCursorPosition
    End With
    
End Sub

Private Function isEndOfMonth(InputDate As Date) As Boolean
    isEndOfMonth = (Day(InputDate + 1) = 1)
End Function

Private Sub TreatDateWithKeyboardEntry( _
            txtDate As MSForms.TextBox, _
            KeyCode As MSForms.ReturnInteger)
    
    ' Allow to exit on Escape key
    If KeyCode = vbKeyEscape Then
        ' Avoid error message on escape
        txtDate.Value = txtDate.BoundValue
        Unload txtDate.Parent
        Exit Sub
    End If
    
    ' Keys with allowed standard behavior
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome _
                Or KeyCode = vbKeyControl Then
        Exit Sub
    End If
    
    ' Exit TextBox
    If KeyCode = vbKeyTab Or KeyCode = vbKeyTab Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
            KeyCode = vbKeyReturn Or KeyCode = vbKeyShift Then
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
Private Sub DeleteLeftTextBox(ByRef txtDate As MSForms.TextBox)
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelStart
        
    Dim CurrentDate As String
    CurrentDate = txtDate.Text
    
    DateDeleteLeft CurrentDate, TextCursorPosition
    
    txtDate.Text = CurrentDate
    txtDate.SelStart = TextCursorPosition
    
End Sub

' Processes BackSpace Key
' Public for Tests
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
Private Sub DeleteRightTextBox(ByRef txtDate As MSForms.TextBox)
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelStart

    Dim CurrentDate As String
    CurrentDate = txtDate.Text

    DateDeleteRight CurrentDate, TextCursorPosition

    txtDate.Text = CurrentDate
    txtDate.SelStart = TextCursorPosition
    
End Sub

' Processes Delete Key
' Public for Tests
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

Private Sub EditDateTextBox( _
            txtDate As MSForms.TextBox, _
            KeyCode As MSForms.ReturnInteger)
    
    Dim TextDate As String
    TextDate = txtDate.Text
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelStart
    
    Dim InputNumber As Byte
    InputNumber = GetDigitFromKeyCode(KeyCode)
    
    DateEdit TextDate, TextCursorPosition, InputNumber
    
    ' Disable events to avoid triggering _Change procedure
    Application.EnableEvents = False
        txtDate.Text = TextDate
        txtDate.SelStart = TextCursorPosition
    Application.EnableEvents = True
    
End Sub

Private Function GetDigitFromKeyCode(KeyCode As MSForms.ReturnInteger) As Byte
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
' Public for Tests
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

Private Sub SetInitialDate(txtDate As MSForms.TextBox)
    
    If Not IsDate(txtDate.Value) Then
        txtDate.Value = Date
    End If

    txtDate.SelStart = 0
    
End Sub

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
