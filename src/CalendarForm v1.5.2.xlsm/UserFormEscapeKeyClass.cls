VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserFormEscapeKeyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents HookedControlTextBox As MSForms.TextBox
Attribute HookedControlTextBox.VB_VarHelpID = -1

Private pForm As MSForms.UserForm

Public Sub Initialize(ctrl As MSForms.TextBox, parentForm As Object)
    Set HookedControlTextBox = ctrl
    Set pForm = parentForm
End Sub

Private Sub HookedControlTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload pForm
        KeyCode = 0
    End If
End Sub
