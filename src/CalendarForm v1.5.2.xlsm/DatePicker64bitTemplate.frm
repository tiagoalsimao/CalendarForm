VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker64bitTemplate 
   Caption         =   "DatePicker64bitTemplate"
   ClientHeight    =   1722
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

Private DatePicker1 As DatePicker64BitClass
Private DatePicker2 As DatePicker64BitClass

Private Sub UserForm_Initialize()
    Set DatePicker1 = New DatePicker64BitClass
    DatePicker1.Initialize DatePicker64bitTextBox1, DatePicker64bitSpinButton1, DatePicker64bitLabel1
                
    Set DatePicker2 = New DatePicker64BitClass
    DatePicker2.Initialize DatePicker64bitTextBox2, DatePicker64bitSpinButton2, DatePicker64bitLabel2
End Sub

' _BeforeUpdate event required to be in userform as it is not available in WithEven from the Class Module.
Private Sub DatePicker64bitTextBox1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
     DatePicker1.Validate Cancel
End Sub

Private Sub DatePicker64bitTextBox2_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
     DatePicker2.Validate Cancel
End Sub
