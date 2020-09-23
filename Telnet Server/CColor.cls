VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function GetForegroundColorCode(Color As ColorCode) As String
    If Color > 8 Then
        GetForegroundColorCode = "1;" & CStr(Color + 22)
    Else
        GetForegroundColorCode = CStr(Color + 30)
    End If
End Function


Public Function GetBackgroundColorCode(Color As ColorCode) As String
    If Color > 8 Then
        GetBackgroundColorCode = CStr(Color + 32)
    Else
        GetBackgroundColorCode = CStr(Color + 40)
    End If
End Function


Public Function SetColorCodes(ForegroundColor As ColorCode, BackgroundColor As ColorCode) As String
    SetColorCodes = modANSI.EscapeKey & "[" & GetForegroundColorCode(ForegroundColor) & ";" & GetBackgroundColorCode(BackgroundColor) & "m"
End Function


Public Function SetForegroundColorCode(ForegroundColor As ColorCode) As String
    SetForegroundColorCode = modANSI.EscapeKey & "[" & GetForegroundColorCode(ForegroundColor) & "m"
End Function


Public Function SetBackgroundColorCode(BackgroundColor As ColorCode) As String
    SetBackgroundColorCode = modANSI.EscapeKey & "[" & GetBackgroundColorCode(BackgroundColor) & "m"
End Function


Public Function ResetColorCodes() As String
    ResetColorCodes = modANSI.EscapeKey & "[0m"
End Function


Public Function ParseColors(ByVal strInput As String) As String

    strInput = Replace$(strInput, "<f=black>", SetForegroundColorCode(ColorCode.Black))
    strInput = Replace$(strInput, "<f=blue>", SetForegroundColorCode(ColorCode.Blue))
    strInput = Replace$(strInput, "<f=cyan>", SetForegroundColorCode(ColorCode.Cyan))
    strInput = Replace$(strInput, "<f=darkblue>", SetForegroundColorCode(ColorCode.DarkBlue))
    strInput = Replace$(strInput, "<f=darkcyan>", SetForegroundColorCode(ColorCode.DarkCyan))
    strInput = Replace$(strInput, "<f=darkgray>", SetForegroundColorCode(ColorCode.DarkGray))
    strInput = Replace$(strInput, "<f=darkgreen>", SetForegroundColorCode(ColorCode.DarkGreen))
    strInput = Replace$(strInput, "<f=darkmagenta>", SetForegroundColorCode(ColorCode.DarkMagenta))
    strInput = Replace$(strInput, "<f=darkred>", SetForegroundColorCode(ColorCode.DarkRed))
    strInput = Replace$(strInput, "<f=darkyellow>", SetForegroundColorCode(ColorCode.DarkYellow))
    strInput = Replace$(strInput, "<f=green>", SetForegroundColorCode(ColorCode.Green))
    strInput = Replace$(strInput, "<f=lightgray>", SetForegroundColorCode(ColorCode.LightGray))
    strInput = Replace$(strInput, "<f=magenta>", SetForegroundColorCode(ColorCode.Magenta))
    strInput = Replace$(strInput, "<f=red>", SetForegroundColorCode(ColorCode.Red))
    strInput = Replace$(strInput, "<f=white>", SetForegroundColorCode(ColorCode.White))
    strInput = Replace$(strInput, "<f=yellow>", SetForegroundColorCode(ColorCode.Yellow))

    strInput = Replace$(strInput, "<b=black>", SetBackgroundColorCode(ColorCode.Black))
    strInput = Replace$(strInput, "<b=blue>", SetBackgroundColorCode(ColorCode.Blue))
    strInput = Replace$(strInput, "<b=cyan>", SetBackgroundColorCode(ColorCode.Cyan))
    strInput = Replace$(strInput, "<b=darkblue>", SetBackgroundColorCode(ColorCode.DarkBlue))
    strInput = Replace$(strInput, "<b=darkcyan>", SetBackgroundColorCode(ColorCode.DarkCyan))
    strInput = Replace$(strInput, "<b=darkgray>", SetBackgroundColorCode(ColorCode.DarkGray))
    strInput = Replace$(strInput, "<b=darkgreen>", SetBackgroundColorCode(ColorCode.DarkGreen))
    strInput = Replace$(strInput, "<b=darkmagenta>", SetBackgroundColorCode(ColorCode.DarkMagenta))
    strInput = Replace$(strInput, "<b=darkred>", SetBackgroundColorCode(ColorCode.DarkRed))
    strInput = Replace$(strInput, "<b=darkyellow>", SetBackgroundColorCode(ColorCode.DarkYellow))
    strInput = Replace$(strInput, "<b=green>", SetBackgroundColorCode(ColorCode.Green))
    strInput = Replace$(strInput, "<b=lightgray>", SetBackgroundColorCode(ColorCode.LightGray))
    strInput = Replace$(strInput, "<b=magenta>", SetBackgroundColorCode(ColorCode.Magenta))
    strInput = Replace$(strInput, "<b=red>", SetBackgroundColorCode(ColorCode.Red))
    strInput = Replace$(strInput, "<b=white>", SetBackgroundColorCode(ColorCode.White))
    strInput = Replace$(strInput, "<b=yellow>", SetBackgroundColorCode(ColorCode.Yellow))

    ParseColors = Replace$(strInput, "<none>", ResetColorCodes())
    
End Function
