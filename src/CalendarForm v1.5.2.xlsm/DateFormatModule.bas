Attribute VB_Name = "DateFormatModule"

#If VBA7 Then
    Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
            ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
#Else
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, _
            ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
#End If

Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SSHORTDATE = &H1F ' short date format string
Private Const LOCALE_SLONGDATE = &H20 ' long date format string

' Based on Soucre: https://www.vbforums.com/showthread.php?33984-How-to-GET-the-system-date-format-from-regional-settings
Public Function GetDateFormat() As String
    
    'Get short date format
    Dim strLocale As String
    strLocale = Space(255)
    
    Dim lngRet As Long
    lngRet = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, strLocale, Len(strLocale))
    
    strLocale = Left(strLocale, lngRet - 1)
    
    GetDateFormat = strLocale
    
End Function
