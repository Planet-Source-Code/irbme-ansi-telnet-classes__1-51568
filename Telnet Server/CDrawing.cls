VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function DrawEdge(EdgeType As EdgeTypeConstants, EdgeStyle As EdgeStyleConstants) As String

    Select Case EdgeType
        Case EdgeTypeConstants.Top
        
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawEdge = Chr$(196)
                Case EdgeStyleConstants.DoubleLine:  DrawEdge = Chr$(205)
                Case EdgeStyleConstants.DoubleThick: DrawEdge = Chr$(223)
            End Select
            
            Exit Function
        Case EdgeTypeConstants.Bottom
            
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawEdge = Chr$(196)
                Case EdgeStyleConstants.DoubleLine:  DrawEdge = Chr$(205)
                Case EdgeStyleConstants.DoubleThick: DrawEdge = Chr$(220)
            End Select
            
            Exit Function
        Case EdgeTypeConstants.Left, EdgeTypeConstants.Right
        
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawEdge = Chr$(179)
                Case EdgeStyleConstants.DoubleLine:  DrawEdge = Chr$(186)
                Case EdgeStyleConstants.DoubleThick: DrawEdge = Chr$(219)
            End Select
    End Select
    
End Function



Public Function DrawCorner(CornerType As CornerTypeConstants, EdgeStyle As EdgeStyleConstants) As String

    Select Case CornerType
        Case CornerTypeConstants.TopLeft
        
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawCorner = Chr$(218)
                Case EdgeStyleConstants.DoubleLine:  DrawCorner = Chr$(201)
                Case EdgeStyleConstants.DoubleThick: DrawCorner = Chr$(219)
            End Select
            
            Exit Function
        Case CornerTypeConstants.BottomLeft
            
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawCorner = Chr$(192)
                Case EdgeStyleConstants.DoubleLine:  DrawCorner = Chr$(200)
                Case EdgeStyleConstants.DoubleThick: DrawCorner = Chr$(219)
            End Select
                
            Exit Function
        Case CornerTypeConstants.BottomRight
            
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawCorner = Chr$(217)
                Case EdgeStyleConstants.DoubleLine:  DrawCorner = Chr$(188)
                Case EdgeStyleConstants.DoubleThick: DrawCorner = Chr$(219)
            End Select
                
            Exit Function
        Case CornerTypeConstants.TopRight
            
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawCorner = Chr$(191)
                Case EdgeStyleConstants.DoubleLine:  DrawCorner = Chr$(187)
                Case EdgeStyleConstants.DoubleThick: DrawCorner = Chr$(219)
            End Select
    End Select

End Function
    
    
    
Public Function DrawDoubleCorner(EdgeType As EdgeTypeConstants, EdgeStyle As EdgeStyleConstants) As String

    Select Case EdgeType
    
        Case EdgeTypeConstants.Top
                                    
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawDoubleCorner = Chr$(192)
                Case EdgeStyleConstants.DoubleLine:  DrawDoubleCorner = Chr$(203)
                Case EdgeStyleConstants.DoubleThick: DrawDoubleCorner = Chr$(219)
            End Select
                
            Exit Function
        Case EdgeTypeConstants.Bottom
            
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawDoubleCorner = Chr$(193)
                Case EdgeStyleConstants.DoubleLine:  DrawDoubleCorner = Chr$(202)
                Case EdgeStyleConstants.DoubleThick: DrawDoubleCorner = Chr$(219)
            End Select
                
            Exit Function
        Case EdgeTypeConstants.Left
            
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawDoubleCorner = Chr$(195)
                Case EdgeStyleConstants.DoubleLine:  DrawDoubleCorner = Chr$(204)
                Case EdgeStyleConstants.DoubleThick: DrawDoubleCorner = Chr$(219)
            End Select
                
            Exit Function
        Case EdgeTypeConstants.Right
            
            Select Case EdgeStyle
                Case EdgeStyleConstants.Normal:      DrawDoubleCorner = Chr$(180)
                Case EdgeStyleConstants.DoubleLine:  DrawDoubleCorner = Chr$(185)
                Case EdgeStyleConstants.DoubleThick: DrawDoubleCorner = Chr$(219)
            End Select
        End Select
            
End Function
    
    

Public Function DrawCentre(EdgeStyle As EdgeStyleConstants) As String

    Select Case EdgeStyle
        Case EdgeStyleConstants.Normal:      DrawCentre = Chr$(197)
        Case EdgeStyleConstants.DoubleLine:  DrawCentre = Chr$(206)
        Case EdgeStyleConstants.DoubleThick: DrawCentre = Chr$(219)
    End Select
    
End Function
        
    

Public Function DrawBlock(Shading As HatchStyleConstants) As String
    
    Select Case Shading
        Case HatchStyleConstants.None:       DrawBlock = " "
        Case HatchStyleConstants.DiagLeft:   DrawBlock = Chr$(178)
        Case HatchStyleConstants.DiagRight:  DrawBlock = Chr$(176)
        Case HatchStyleConstants.Dotted:     DrawBlock = Chr$(177)
        Case HatchStyleConstants.Filled:     DrawBlock = Chr$(219)
    End Select
    
End Function


Public Function DrawBox(EdgeStyle As EdgeStyleConstants, Width As Byte, Height As Byte)

  Dim x As Integer, y As Integer
  
    DrawBox = DrawBox & DrawCorner(CornerTypeConstants.TopLeft, EdgeStyle)
  
    For x = 2 To Width - 1
        DrawBox = DrawBox & DrawEdge(EdgeTypeConstants.Top, EdgeStyle)
    Next x
  
    DrawBox = DrawBox & DrawCorner(CornerTypeConstants.TopRight, EdgeStyle) & vbCrLf
    
    For y = 2 To Height - 1
        DrawBox = DrawBox & DrawEdge(EdgeTypeConstants.Left, EdgeStyle)
    
        For x = 2 To Width - 1
            DrawBox = DrawBox & " "
        Next x
        
        DrawBox = DrawBox & DrawEdge(EdgeTypeConstants.Right, EdgeStyle) & vbCrLf
    Next y
     
    DrawBox = DrawBox & DrawCorner(CornerTypeConstants.BottomLeft, EdgeStyle)
  
    For x = 2 To Width - 1
        DrawBox = DrawBox & DrawEdge(EdgeTypeConstants.Bottom, EdgeStyle)
    Next x
    
    DrawBox = DrawBox & DrawCorner(CornerTypeConstants.BottomRight, EdgeStyle)
     
End Function
