VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker64bitTemplate 
   Caption         =   "DatePicker64bitTemplate"
   ClientHeight    =   1246
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

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
            ByVal lpClassName As String, _
            ByVal lpWindowName As String) As LongPtr
    
    Private Declare PtrSafe Function ClientToScreen Lib "user32" ( _
            ByVal hwnd As LongPtr, _
            lpPoint As POINTAPI) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
            ByVal lpClassName As String, _
            ByVal lpWindowName As String) As Long
    
    Private Declare Function ClientToScreen Lib "user32" ( _
            ByVal hWnd As Long, _
            lpPoint As POINTAPI) As Long
#End If

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TWIPS_PER_INCH As Long = 1440

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type FormPosition
    Top As Long
    Left As Long
End Type

Private controlHooks As Collection

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
    
    Dim lblLeft As Double, lblTop As Double, lblRight As Double, lblBottom As Double
    lblLeft = lbl.Left
    lblTop = lbl.Top
    lblRight = lbl.Left + lbl.Width
    lblBottom = lbl.Top + lbl.Height
    
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
    
        If TypeName(ctrl) = "TextBox" Then
        
            Dim txtLeft As Double, txtTop As Double, txtRight As Double, txtBottom As Double
            
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

Sub UpdateDateWithCalendar(txtDate As MSForms.TextBox)

    Dim InitialTextDate As String
    InitialTextDate = txtDate.Text
    
    Dim CalendarTopLeftPosition As FormPosition
    CalendarTopLeftPosition = GetPopupPosition(txtDate, calendarForm)
    
    Dim InitialDate As Date
    If IsDate(InitialTextDate) Then
        InitialDate = CDate(InitialTextDate)
    Else
        InitialDate = Date
    End If
    
    Dim DateSelected As Date
    DateSelected = calendarForm.GetDate(InitialDate, Monday, , , , , True, False, True, _
            FirstFourDays, CalendarTopLeftPosition.Top, CalendarTopLeftPosition.Left, TodayFontColor:=vbRed)
    
    If DateSelected = 0 Then
        Exit Sub
    End If
    
    Dim DateFormat As String
    DateFormat = GetDateFormat()
    
    ' Force output the same as the date format
    txtDate.Text = Format(DateSelected, DateFormat)
    
End Sub

Private Function GetPopupPosition(ctrl As MSForms.Control, calendarForm As Object) As FormPosition
    
    Dim hwnd As LongPtr
    hwnd = FindWindow("ThunderDFrame", ctrl.Parent.Caption)
    
    Dim pt As POINTAPI
    ClientToScreen hwnd, pt

    ' Convert to pixels using custom functions
    
    Dim lLeft As Long, lTop As Long
    lLeft = TwipsToPixelsX(ctrl.Left)
    lTop = TwipsToPixelsY(ctrl.Top)
    
    Dim pos As FormPosition
    pos.Left = pt.X + lLeft + TwipsToPixelsX(ctrl.Width)
    pos.Top = pt.Y + lTop

    GetPopupPosition = pos
    
End Function

Public Function TwipsToPixelsX(twips As Single) As Long
    Dim hdc As LongPtr
    Dim dpiX As Long
    hdc = GetDC(0)
    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
    TwipsToPixelsX = twips * dpiX / TWIPS_PER_INCH
End Function

Public Function TwipsToPixelsY(twips As Single) As Long
    Dim hdc As LongPtr
    Dim dpiY As Long
    hdc = GetDC(0)
    dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
    TwipsToPixelsY = twips * dpiY / TWIPS_PER_INCH
End Function

Private Sub DatePicker64bit_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    ValidateDate Me.ActiveControl
End Sub

Private Sub ValidateDate(txtDate As MSForms.TextBox)

    If Not IsDate(txtDate) Then
        MsgBox "The date '" & txtDate.Text & "' is not valid"
        Exit Sub
    End If
    
    txtDate.value = txtDate.Text
    
End Sub

Private Sub UserForm_Initialize()
    HookEscapeKey Me
End Sub

Private Sub HookEscapeKey(pForm As MSForms.UserForm)
    
    Set controlHooks = New Collection
    
    Dim ctrl As Control
    For Each ctrl In pForm.Controls
        Select Case TypeName(ctrl)
            Case "TextBox" ', "ComboBox", "ListBox" ' Only controls that support KeyDown
                
                Dim hook As UserFormEscapeKeyClass
                Set hook = New UserFormEscapeKeyClass
                hook.Initialize ctrl, pForm
                controlHooks.Add hook
        End Select
    Next
End Sub

Private Sub DatePicker64bit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    TreatDateWithKeyboardEntry DatePicker64bit, KeyCode
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
            KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        Exit Sub
    End If
    
    If txtDate.SelLength > 1 Then
        DeleteSelectedText txtDate
        
        ' As text already deleted, stop execution
        If KeyCode = vbKeyBack Then
            txtDate.SelStart = 0
            GoTo ExitProcedure
        ElseIf KeyCode = vbKeyDelete Then
            GoTo ExitProcedure
        End If
        
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
    TextCursorPosition = txtDate.SelStart
    
    With txtDate
        .SelText = Mid(DATE_MASK, .SelStart + 1, .SelLength)
    End With
    
    txtDate.SelStart = TextCursorPosition
    
End Sub

Private Sub DatePicker64bit_Enter()
    SetDateMask DatePicker64bit
End Sub

Private Sub SetDateMask(txtDate As MSForms.TextBox)
    
    If Not IsDate(txtDate.Text) Then
        txtDate.Text = Date
    End If
    
    txtDate.SelStart = 0
    
End Sub
