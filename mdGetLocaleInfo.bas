Attribute VB_Name = "mdGetLocaleInfo"
Option Explicit

'--- will Debug.Print what's been asked by dtpicker :-))
#Const LISTEN_TO_LOCALE_INFO = False

Public Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
'Public Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
'Public Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
'Public Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
'Public Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
'Public Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Public Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

'--- customize short day names format
Private Const STR_DAY_FORMAT        As String = "  %1  "
Private Const STR_SHORT_DAY_NAMES   As String = "Ïí|Âò|Ñð|×ò|Ïò|Ñá|Íä"

Public Function MyGetLocaleInfo( _
            ByVal Locale As Long, _
            ByVal LCType As Long, _
            ByVal lpLCData As Long, _
            ByVal cchData As Long) As Long
    Static vSplit   As Variant
    Dim sRet        As String
    
    On Error Resume Next
    '--- check if short names of weekdays
    If LCType >= LOCALE_SABBREVDAYNAME1 _
            And LCType <= LOCALE_SABBREVDAYNAME7 Then
        '--- construct array with the short names of weekdays
        If Not IsArray(vSplit) Then
            vSplit = Split(STR_SHORT_DAY_NAMES, "|")
        End If
        '--- format day name
        sRet = Replace(STR_DAY_FORMAT, "%1", vSplit(LCType - LOCALE_SABBREVDAYNAME1))
        '--- copy only if buffer is large enough
        If cchData >= Len(sRet) + 1 Then
            CopyMemory ByVal lpLCData, ByVal sRet, Len(sRet) + 1
        End If
        '--- return size
        MyGetLocaleInfo = Len(sRet) + 1
    Else
        '--- pass to win32 implementation
        MyGetLocaleInfo = GetLocaleInfo(Locale, LCType, lpLCData, cchData)
    End If
    
#If LISTEN_TO_LOCALE_INFO Then
    If lpLCData <> 0 And cchData > 0 Then
        sRet = String(cchData + 1, 0)
        CopyMemory ByVal sRet, ByVal lpLCData, cchData
        If InStr(1, sRet, Chr(0)) > 0 Then
            sRet = Left(sRet, InStr(1, sRet, Chr(0)))
            Debug.Print Hex(LCType) & ": " & sRet
        End If
    End If
#End If

End Function
