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

Private Sub TextBox1_Enter()
    Dim CurrentDateVariant As Variant
    CurrentDateVariant = TextBox1.value
    
    If IsDate(CurrentDateVariant) Then
        Dim CurrentDate As Date
        CurrentDate = CDate(CurrentDateVariant)
    End If
         
    Dim DateSelected
    DateSelected = CalendarForm.GetDate(CurrentDate, Monday, , , , , True, False, True, FirstFourDays, TodayFontColor:=255)
    
    If DateSelected <> 0 Then TextBox1.value = DateSelected
    
End Sub

Private Sub UserForm_Click()

End Sub
