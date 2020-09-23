Attribute VB_Name = "modANSI"
Option Explicit


Public Const EscapeKey As String = ""


Public Enum DirectionConstants
    Up = 65
    Down = 66
    Right = 67
    Left = 68
End Enum


Public Enum CornerTypeConstants
    TopLeft
    TopRight
    BottomLeft
    BottomRight
End Enum


Public Enum EdgeTypeConstants
    Left
    Right
    Top
    Bottom
End Enum
    

Public Enum EdgeStyleConstants
    Normal
    DoubleLine
    DoubleThick
End Enum
    

Public Enum HatchStyleConstants
    None
    DiagLeft
    DiagRight
    Dotted
    Filled
End Enum


Public Enum ColorCode
    Black
    DarkRed
    DarkGreen
    DarkYellow
    DarkBlue
    DarkMagenta
    DarkCyan
    LightGray
    DarkGray
    Red
    Green
    Yellow
    Blue
    Magenta
    Cyan
    White
End Enum


