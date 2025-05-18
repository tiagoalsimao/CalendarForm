VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTextBoxClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents DateTextBox As MSForms.TextBox
Attribute DateTextBox.VB_VarHelpID = -1
Public DateTextBox2 As MSForms.TextBox


Private pForm As MSForms.UserForm

Public Sub Initialize(ctrl As MSForms.TextBox, parentForm As Object)
    Set DateTextBox = ctrl
    Set pForm = parentForm
End Sub

