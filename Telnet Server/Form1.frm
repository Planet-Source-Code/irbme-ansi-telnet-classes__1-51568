VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Telnet Server"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock ws 
      Left            =   1365
      Top             =   1890
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Connect to your IP (on port 23) with telnet. Make sure you have a coloured terminal."
      Height          =   435
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4530
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Drawing As CDrawing
Dim Color   As CColor
Dim Cursor  As CCursor
Dim Eraser  As CErase

Dim Data As String
Dim Temp As String

Private Sub Form_Load()
    Set Drawing = New CDrawing
    Set Color = New CColor
    Set Cursor = New CCursor
    Set Eraser = New CErase

    ws.LocalPort = 23
    ws.Listen
End Sub


Private Sub ws_Close()
    Unload Me
End Sub


Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    ws.Close
    ws.Accept requestID
    Me.Caption = Me.Caption & " - Connected"
    Welcome
End Sub


Private Sub WaitKey()
    Data = vbNullString
    
    While Data = vbNullString Or InStr(Data, "ÿ") > 0 Or InStr(Data, "û") > 0 Or InStr(Data, "") > 0 Or InStr(Data, "") > 0
        DoEvents
    Wend
    
End Sub


Private Sub Welcome()

    Dim Buffer As String
    Dim i As Integer, j As Integer
    
        Buffer = "ÿûÿû"
        ws.SendData Buffer

        Do
            ws.GetData (Data)
        Loop While Data <> vbNullString

        Buffer = Buffer & Color.SetColorCodes(White, Black)
        Buffer = Buffer & vbCrLf & "You are connected from " & ws.RemoteHostIP
        Buffer = Buffer & vbCrLf
        Buffer = Buffer & "[Press any key to begin demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & vbCrLf & "First, let us demonstrate basic colour codes:"
        Buffer = Buffer & vbCrLf & Color.ParseColors("<f=red>Red<f=white>, <f=green>Green<f=white>, ")
        Buffer = Buffer & Color.ParseColors("<f=blue>Blue<f=white>, <f=yellow>Yellow<f=white>, ")
        Buffer = Buffer & Color.ParseColors("<f=magenta>Magenta<f=white>, <f=cyan>Cyan<f=white>, ")
        Buffer = Buffer & Color.ParseColors("<f=darkgray>Gray")
        Buffer = Buffer & vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "[Press any key to continue demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Buffer & vbCrLf
        Buffer = Buffer & vbCrLf & "Now some more advanced color codes:"
        Buffer = Buffer & vbCrLf & Color.ParseColors("<f=white><b=red>Red<b=black>, <b=green>Green<b=black>, ")
        Buffer = Buffer & Color.ParseColors("<b=blue>Blue<b=black>, <b=yellow>Yellow<b=black>, ")
        Buffer = Buffer & Color.ParseColors("<b=magenta>Magenta<b=black>, <b=cyan>Cyan<b=black>, ")
        Buffer = Buffer & Color.ParseColors("<b=lightgray>Gray")
        Buffer = Buffer & vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "[Press any key to continue demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & Color.SetColorCodes(White, Black)
        Buffer = Buffer & vbCrLf & "And now a basic text echo demonstration."
        Buffer = Buffer & Color.ParseColors(vbCrLf & "Enter a password: <b=lightgray>        ")
        Buffer = Buffer & Cursor.MoveDirection(DirectionConstants.Left, 8)
        ws.SendData Buffer
        
        Temp = vbNullString
        
        For i = 1 To 8
            WaitKey
            Temp = Temp & Data
            ws.SendData "*"
        Next i
        
        Buffer = vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "You typed: " & Temp
        Buffer = Buffer & vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "[Press any key to continue demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & vbCrLf & Color.SetColorCodes(Green, Black)
        Buffer = Buffer & vbCrLf & "Here are some graphics:"
        Buffer = Buffer & vbCrLf & vbCrLf & Drawing.DrawBox(Normal, 20, 10)
        Buffer = Buffer & vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "[Press any key to continue demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & vbCrLf & Color.SetColorCodes(Red, Black)
        Buffer = Buffer & vbCrLf & "Here are some more graphics:"
        Buffer = Buffer & vbCrLf & vbCrLf & Drawing.DrawBox(DoubleThick, 20, 10)
        Buffer = Buffer & vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "[Press any key to continue demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & vbCrLf & Color.SetColorCodes(Yellow, Black)
        Buffer = Buffer & vbCrLf & "Here are some more graphics:"
        Buffer = Buffer & vbCrLf & vbCrLf & Drawing.DrawBox(DoubleLine, 20, 10)
        Buffer = Buffer & vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "[Press any key to continue demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & Color.SetColorCodes(White, Black)
        Buffer = Buffer & vbCrLf & "Now some coloured, patterened block demonstrations:" & vbCrLf
        
        For i = 9 To 15
        
            Buffer = Buffer & vbCrLf & EscapeKey & "[" & 1 & ";" & CStr(i + 22) & ";" & CStr(i + 32) & "m"
        
            For j = 1 To 40
                Buffer = Buffer & Drawing.DrawBlock(DiagLeft)
            Next
            
            Buffer = Buffer & vbCrLf
            
            For j = 1 To 40
                Buffer = Buffer & Drawing.DrawBlock(DiagRight)
            Next
            
            Buffer = Buffer & vbCrLf
            
            For j = 1 To 40
                Buffer = Buffer & Drawing.DrawBlock(Dotted)
            Next
            
            Buffer = Buffer & vbCrLf
            
            For j = 1 To 40
                Buffer = Buffer & Drawing.DrawBlock(Filled)
            Next
        Next i
        
        Buffer = Buffer & vbCrLf & vbCrLf & Color.SetColorCodes(White, Black)
        Buffer = Buffer & "[Press any key to continue demonstration]"
        
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & vbCrLf & "Now type 10 characters. This demonstrates the use of backspace"
        Buffer = Buffer & vbCrLf & Drawing.DrawBox(DoubleLine, 12, 3)
        Buffer = Buffer & Cursor.MoveDirection(Up, 1)
        Buffer = Buffer & Cursor.MoveDirection(DirectionConstants.Left, 11)
        Buffer = Buffer & Color.SetColorCodes(DarkBlue, LightGray)
        Buffer = Buffer & "          " & Cursor.MoveDirection(DirectionConstants.Left, 10)
        
        ws.SendData Buffer
        
        Temp = vbNullString
        i = 0: j = 0
        
        While i < 10
            WaitKey
            If Data = Chr$(8) Then
                If i > 0 Then
                    ws.SendData (Cursor.MoveDirection(DirectionConstants.Left, 1) & " " & Cursor.MoveDirection(DirectionConstants.Left, 1))
                    Temp = Left$(Temp, Len(Temp) - 1)
                    i = i - 1
                    j = j + 1
                End If
            Else
                ws.SendData Data
                Temp = Temp & Data
                i = i + 1
            End If
        Wend
        
        Buffer = Cursor.MoveDirection(DirectionConstants.Down, 1)
        Buffer = Buffer & Color.SetColorCodes(White, Black)
        Buffer = Buffer & vbCrLf & "You typed: " & Temp
        Buffer = Buffer & vbCrLf & "You deleted " & j & " characters"
        Buffer = Buffer & vbCrLf & "[Press any key to continue demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        Buffer = Eraser.EntireScreen
        Buffer = Buffer & vbCrLf & "Thank you for connecting. You will now be disconnected"
        Buffer = Buffer & vbCrLf & "[Press any key to end demonstration]"
        ws.SendData Buffer
        
        WaitKey
        
        ws.Close
        Unload Me
End Sub


Private Sub ws_DataArrival(ByVal bytesTotal As Long)

  Dim i As Integer

    ws.GetData Data
End Sub
