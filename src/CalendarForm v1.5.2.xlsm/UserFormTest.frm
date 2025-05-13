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

Private Sub CalendarImageLabel_Click()
    UpdateDateWithCalendar TextBox2
End Sub

Private Sub ComboBoxDate_DropButtonClick()
    UpdateDateWithCalendar ComboBoxDate
End Sub

Sub UpdateDateWithCalendar(txtDate As MSForms.TextBox)
    txtDate.Text = GetCalendarDate(txtDate.Text)
End Sub

Private Sub ComboBoxDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TreatDateWithKeyboardEntry TextBox2, KeyCode
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub LabelDate_Click()
    LabelDate.Caption = GetCalendarDate(LabelDate.Caption)
End Sub

Private Sub ListBox1_Click()

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
    DateSelected = CalendarForm.GetDate(CurrentDate, Monday, , , , , True, False, True, FirstFourDays, TodayFontColor:=255)
    
    If DateSelected <> 0 Then
        
        ' Force output the same as the date format
        GetCalendarDate = CStr(DateSelected)
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
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight _
           Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Then
        Exit Sub
    End If
    
    Dim KeyPosition As Byte
    KeyPosition = txtDate.SelStart
    
    ' TODO: ' Allow Delete
    ' TODO: ' Allow Selection of part of the date = Replace all selected value with "_"
    If KeyCode = vbKeyBack And KeyPosition > 0 Then
        ' Allow Backspace
        txtDate.Text = DateBackSpace(txtDate.Text, KeyPosition)
        txtDate.SelStart = WorksheetFunction.Max(0, KeyPosition - 1)
    ElseIf (KeyCode >= vbKey0 And KeyCode <= vbKey9) _
            Or (KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) Then
        ' Allow numbers
        UpdateDateTextBox txtDate, KeyCode
    End If
    
    ' Cancel any key update
    KeyCode = 0
    
End Sub

Private Sub UpdateDateTextBox( _
            txtDate As MSForms.TextBox, _
            KeyCode As MSForms.ReturnInteger)
    
    With txtDate
        Dim TextDate As String
        TextDate = .Text
        
        Dim KeyPosition As Byte
        KeyPosition = .SelStart + 1
        
        Dim InputNumber As Byte
        InputNumber = GetDigitFromKeyCode(KeyCode)
        
        UpdateDate TextDate, KeyPosition, InputNumber
        
        ' Skip Date Separator
        If KeyPosition = 2 Or KeyPosition = 5 Then KeyPosition = KeyPosition + 1
        
        .Text = TextDate
        .SelStart = KeyPosition
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

