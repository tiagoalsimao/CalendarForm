VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker64bitTemplate 
   Caption         =   "DatePicker64bitTemplate"
   ClientHeight    =   1946
   ClientLeft      =   -434
   ClientTop       =   -1904
   ClientWidth     =   2842
   OleObjectBlob   =   "DatePicker64bitTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePicker64bitTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DatePicker1 As DatePicker64BitClass
Private DatePicker2 As DatePicker64BitClass

Private Sub UserForm_Initialize()
    
    Set DatePicker1 = New DatePicker64BitClass
    DatePicker1.Initialize DatePicker64bitTextBox1, DatePicker64bitSpinButton1, DatePicker64bitLabel1
    
    Set DatePicker2 = New DatePicker64BitClass
    DatePicker2.Initialize DatePicker64bitTextBox2, DatePicker64bitSpinButton2, DatePicker64bitLabel2
    
    DatePicker1.Value = Date
    DatePicker2.Value = Date + 1
    
End Sub

Private Sub DatePicker64bitTextBox1_Enter()
    DatePicker1.Enter
End Sub

Private Sub DatePicker64bitTextBox2_Enter()
    DatePicker2.Enter
End Sub

' _BeforeUpdate event required to be in userform as it is not available in WithEven from the Class Module.
Private Sub DatePicker64bitTextBox1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
     DatePicker1.Validate_BeforeUpdate Cancel
End Sub

Private Sub DatePicker64bitTextBox2_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
     DatePicker2.Validate_BeforeUpdate Cancel
End Sub

Private Sub DatePicker64bitTextBox1_Change()
    
    If Not DatePicker1.HasValidDate Then Exit Sub
    If DatePicker2 Is Nothing Then Exit Sub
    
    If DatePicker1.Value > DatePicker2.Value Then
        'MsgBox "Date 1 cannot be later that Date 2. setting Date 2 to " & CStr(DatePicker1.Value)
        DatePicker2.Value = DatePicker1.Value
    End If
End Sub

Private Sub DatePicker64bitTextBox2_Change()
    
    If Not DatePicker2.HasValidDate Then Exit Sub
    If DatePicker1 Is Nothing Then Exit Sub
    
    If DatePicker1.Value > DatePicker2.Value Then
        'MsgBox "Date 1 cannot be later that Date 2. setting Date 1 to " & CStr(DatePicker2.Value)
        DatePicker1.Value = DatePicker2.Value
    End If
End Sub
