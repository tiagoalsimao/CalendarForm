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
    DatePicker1.Initialize TextBox:=Me.DatePicker64bitTextBox1, _
                SpinButton:=Me.DatePicker64bitSpinButton1, Label:=Me.DatePicker64bitLabel1
                
    Set DatePicker2 = New DatePicker64BitClass
    DatePicker2.Initialize TextBox:=Me.DatePicker64bitTextBox2, _
                SpinButton:=Me.DatePicker64bitSpinButton2, Label:=Me.DatePicker64bitLabel2
End Sub
