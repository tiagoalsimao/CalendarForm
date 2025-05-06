VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTest 
   Caption         =   "UserForm1"
   ClientHeight    =   2863
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   4298
   OleObjectBlob   =   "UserFormTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim myDate
    myDate = CalendarForm.GetDate
End Sub

