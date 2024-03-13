Attribute VB_Name = "Player"
Public playerRange As Range
Public Const STARTY As Integer = 47


Sub MovePlayer(xInc As Integer)
    ClearPlayer
    playerX = playerX + xInc
    'Check boundaries
    If playerX < boundaryLeft Then
        playerX = boundaryLeft
    ElseIf playerX + 2 > boundaryRight Then
        playerX = boundaryRight - 2
    End If
    DrawPlayer
End Sub

Sub DrawPlayer()
    Set playerRange = Union(Cells(STARTY, playerX + 1), _
                            Cells(STARTY + 1, playerX), Cells(STARTY + 1, playerX + 1), Cells(STARTY + 1, playerX + 2), _
                            Cells(STARTY + 2, playerX), Cells(STARTY + 2, playerX + 1), Cells(STARTY + 2, playerX + 2))
    playerRange.Interior.color = RGB(255, 255, 255)
End Sub

Sub ClearPlayer()
    If Not playerRange Is Nothing Then
        playerRange.Interior.color = RGB(0, 0, 0)
    End If
End Sub


