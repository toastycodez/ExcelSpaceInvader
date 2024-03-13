Attribute VB_Name = "Bullet"
Public bulletCell As Range
Public bulletY As Integer
Sub MoveBullet()
    bulletY = STARTY
    ' Loop until the bullet hits the top boundary
    Do Until bulletY <= boundaryTop - 1
        ClearBullet
        DrawBullet
        bulletY = bulletY - 1
        Sleep (10)
    Loop
    ClearBullet
End Sub

Sub DisplayYouWinMessage()
    With Range("P30")
        .Value = "You Win!"
        .Font.Size = 24
        .Font.Bold = True
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub DrawBullet()
    If bulletY >= 1 And bulletY <= ActiveSheet.Rows.Count Then
        Set bulletCell = Cells(bulletY, playerX + 1)
        bulletCell.Interior.color = RGB(255, 255, 0)
    End If
End Sub

Sub ClearBullet()
    Dim nextCell As Range
    If bulletY - 1 >= 1 Then
        Set nextCell = Cells(bulletY - 1, playerX + 1)
    End If
    If Not bulletCell Is Nothing Then
        ' Check if is cell contains alien
        If Not nextCell Is Nothing Then
            If nextCell.Interior.color <> RGB(0, 0, 0) Then
                score = score + 1 ' Increase score by 1
                
                ' Update score display
                With Range("D3")
                    .Value = "Score: " & score
                    .Font.Size = 14
                    .Font.Bold = True
                    .Font.color = RGB(255, 255, 255)
                End With
                
                ' Check if the player wins
                If score >= 255 Then
                    DisplayYouWinMessage
                End If
            End If
        End If
        bulletCell.Interior.color = RGB(0, 0, 0)
    End If
End Sub
