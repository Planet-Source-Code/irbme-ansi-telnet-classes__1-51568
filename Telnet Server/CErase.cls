VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CErase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Property Get FromCursorToEndOfLine() As String
    FromCursorToEndOfLine = modANSI.EscapeKey & "[K"
End Property


Public Property Get FromBeginningOfLineToCursor() As String
    FromBeginningOfLineToCursor = modANSI.EscapeKey & "[1K"
End Property


Public Property Get EntireLine() As String
    EntireLine = modANSI.EscapeKey & "[2K"
End Property


Public Property Get FromCursorToEndOfScreen() As String
    FromCursorToEndOfScreen = modANSI.EscapeKey & "[J"
End Property


Public Property Get FromBeginningOfScreenToCuror() As String
    FromBeginningOfScreenToCuror = modANSI.EscapeKey & "[1J"
End Property


Public Property Get EntireScreen() As String
    EntireScreen = modANSI.EscapeKey & "[2J"
End Property
