VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTest 
   Caption         =   "UserForm1"
   ClientHeight    =   4438
   ClientLeft      =   56
   ClientTop       =   252
   ClientWidth     =   10871
   OleObjectBlob   =   "UserFormTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub LabelDate_Click()
    LabelDate.Caption = GetCalendarDate(LabelDate.Caption)
End Sub

Private Sub UserForm_Initialize()
    ComboBoxDate.value = Date
    txtDate.value = Date
    LabelDate.Caption = Date
End Sub

Private Sub ComboBoxDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub ComboBoxDate_Enter()
    ComboBoxDate.value = GetCalendarDate(ComboBoxDate.value)
    TextBox2.SetFocus
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
        Dim DateFormat As String
        DateFormat = DateFormatModule.GetDateFormat()
        
        ' Force output the same as the date format
        UserFormObjectValue = Format(DateSelected, DateFormat)
    End If
 
End Function

Private Sub txtDate_Enter()
    txtDate.value = GetCalendarDate(txtDate.value)
End Sub

Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    Select Case KeyAscii
'        Case 48 To 57 ' Allow digits 0–9
'        Case 8        ' Allow backspace
'        Case Else
            KeyAscii = 0 ' Block all other input
'    End Select
End Sub

'Private Sub txtDate_Change()
'    Dim cleaned As String
'    Dim formatted As String
'
'    Dim TypedDate As String
'    TypedDate = txtDate.Text
'
'
'    If IsDate(TypedDate) Then Exit Sub
'
'    cleaned = Replace(txtDate.Text, "/", "")
'
'    ' Auto-insert slashes as DD/MM/YYYY
'    If Len(cleaned) > 8 Then cleaned = Left(cleaned, 8)
'
'    Select Case Len(cleaned)
'        Case 3 To 4
'            formatted = Left(cleaned, 2) & "/" & Mid(cleaned, 3)
'        Case 5 To 6
'            formatted = Left(cleaned, 2) & "/" & Mid(cleaned, 3, 2) & "/" & Mid(cleaned, 5)
'        Case 7 To 8
'            formatted = Left(cleaned, 2) & "/" & Mid(cleaned, 3, 2) & "/" & Mid(cleaned, 5)
'        Case Else
'            formatted = cleaned
'    End Select
'
'    If formatted <> txtDate.Text Then
'        txtDate.Text = formatted
'        txtDate.SelStart = Len(formatted)
'    End If
'End Sub
'
'Private Sub txtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'    If Not IsDate(txtDate.Text) Then
'        MsgBox "Invalid date. Please enter a valid date in DD/MM/YYYY format.", vbExclamation
'        Cancel = True
'    Else
'        txtDate.Text = Format(CDate(txtDate.Text), "dd/mm/yyyy")
'    End If
'End Sub

