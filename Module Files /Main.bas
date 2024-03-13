Attribute VB_Name = "Main"
#If VBA7 Then

'For 64-Bit MS Office
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else
'For 32-Bit MS Office
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If
Public boundaryLeft As Integer
Public boundaryRight As Integer
Public boundaryTop As Integer
Public playerX As Integer
Public score As Integer



Sub GameStart()
    Cells.Clear
    'Set game setup/background
    boundaryLeft = 3
    boundaryRight = 33
    boundaryTop = 1
    Range(Cells(boundaryTop, boundaryLeft), Cells(50, boundaryRight)).Interior.color = RGB(0, 0, 0)
    'Start Player in the middle
    playerX = 15
    Player.DrawPlayer
    xInc = 0
    score = 0
    SetKeys
    'Create Aliens
    InitializeAliens
    DrawAliens
    ' Display the score
    With Range("D3")
        .Value = "Score: " & score
        .Font.Size = 14
        .Font.Bold = True
        .Font.color = RGB(255, 255, 255)
    End With
End Sub


