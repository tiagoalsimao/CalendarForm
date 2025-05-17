VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTest 
   Caption         =   "UserForm1"
   ClientHeight    =   4459
   ClientLeft      =   -189
   ClientTop       =   -826
   ClientWidth     =   7602
   OleObjectBlob   =   "UserFormTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub CalendarImageLabel_Click()
'    UpdateDateWithCalendar TextBox2
'End Sub
Private controlHooks As Collection

Private Const DATE_MASK As String = "__/__/____"
Private Const DATE_SEPARATOR As String = "/"

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

Function GetTextBoxUnderLabel(lbl As MSForms.Label) As MSForms.TextBox
    Dim ctrl As MSForms.Control
    Dim lblLeft As Double, lblTop As Double, lblRight As Double, lblBottom As Double
    Dim txtLeft As Double, txtTop As Double, txtRight As Double, txtBottom As Double

    lblLeft = lbl.Left
    lblTop = lbl.Top
    lblRight = lbl.Left + lbl.Width
    lblBottom = lbl.Top + lbl.Height

    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
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


Private Sub ComboBoxDate_DropButtonClick()
    UpdateDateWithCalendar ComboBoxDate
End Sub

Sub UpdateDateWithCalendar(txtDate As MSForms.TextBox)
    txtDate.Text = GetCalendarDate(txtDate.Text)
End Sub

Private Sub ComboBoxDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TreatDateWithKeyboardEntry TextBox2, KeyCode
End Sub

Private Sub LabelDate_Click()
    LabelDate.Caption = GetCalendarDate(LabelDate.Caption)
End Sub

Private Sub UserForm_Initialize()
    ComboBoxDate.value = Date
    txtDate.value = Date
    LabelDate.Caption = Date
End Sub

Private Function GetCalendarDate(UserFormObjectValue As String) As String

    Dim CurrentDateVariant As Variant
    CurrentDateVariant = UserFormObjectValue
    
    If IsDate(CurrentDateVariant) Then
        Dim CurrentDate As Date
        CurrentDate = CDate(CurrentDateVariant)
    End If
    
    Dim DateSelected
    DateSelected = CalendarForm.GetDate(CurrentDate, Monday, , , , , True, False, True, FirstFourDays, TodayFontColor:=vbRed)
    
    If DateSelected <> 0 Then
        
        Dim DateFormat As String
        DateFormat = GetDateFormat()
        
        ' Force output the same as the date format
        GetCalendarDate = Format(DateSelected, DateFormat)
    End If
 
End Function

Private Sub txtDate_Enter()
    txtDate.value = GetCalendarDate(txtDate.value)
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TreatDateWithKeyboardEntry TextBox2, KeyCode
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
            KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = vbKeyE Then
        Exit Sub
    End If
    
    If txtDate.SelLength > 1 Then
        DeleteSelectedText txtDate
        
        ' As text already deleted, stop execution
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then GoTo ExitProcedure
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

' Replaces the current selected date with the mask.
' Imitates the deletion of the selected text
Private Sub DeleteSelectedText(txtDate As MSForms.TextBox)
    
    Dim TextCursorPosition As Byte
    TextCursorPosition = txtDate.SelText
    
    With txtDate
        .SelText = Mid(DATE_MASK, .SelStart + 1, .SelLength)
    End With
    
    txtDate.SelText = TextCursorPosition
    
End Sub

Private Sub EditDateTextBox( _
            txtDate As MSForms.TextBox, _
            KeyCode As MSForms.ReturnInteger)
    
    With txtDate
        Dim TextDate As String
        TextDate = .Text
        
        Dim TextCursorPosition As Byte
        TextCursorPosition = .SelStart + 1
        
        Dim InputNumber As Byte
        InputNumber = GetDigitFromKeyCode(KeyCode)
        
        DateEdit TextDate, TextCursorPosition, InputNumber
        
        ' Skip Date Separator
        If TextCursorPosition = 2 Or TextCursorPosition = 5 Then TextCursorPosition = TextCursorPosition + 1
        
        .Text = TextDate
        .SelStart = TextCursorPosition
    End With
    
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

Private Sub TextBox2_Enter()
    SetDateMask TextBox2
End Sub

Private Sub SetDateMask(txtDate As MSForms.TextBox)
    ' Set initial mask if empty
    If txtDate.Text = "" Then
        txtDate.Text = "__/__/____"
        txtDate.SelStart = 0
        txtDate.SelLength = 2
    End If
End Sub

Private Sub TextBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Optional: Adjust selection to prevent selecting slashes
End Sub


