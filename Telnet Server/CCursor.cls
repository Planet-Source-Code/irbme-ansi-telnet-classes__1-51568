VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Property Get MoveDirection(Direction As DirectionConstants, Count As Integer) As String
    MoveDirection = modANSI.EscapeKey & "[" & CStr(Count) & Chr$(Direction)
End Property


Public Property Get MoveTo(Row As Integer, Column As Integer) As String
    MoveTo = modANSI.EscapeKey & "[" & CStr(Row) & ";" + CStr(Column) & "H"
End Property


Public Property Get Save() As String
    Save = modANSI.EscapeKey & "7"
End Property


Public Property Get Restore() As String
    Restore = modANSI.EscapeKey & "8"
End Property
