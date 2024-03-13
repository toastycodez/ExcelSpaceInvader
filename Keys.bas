Attribute VB_Name = "Keys"
Sub SetKeys()
    Application.OnKey "{LEFT}", "MoveLeft"
    Application.OnKey "{RIGHT}", "MoveRight"
    Application.OnKey "{UP}", "shoot"
End Sub

Sub MoveLeft()
    Player.MovePlayer -1 ' Corrected to call MovePlayer from Player module
End Sub

Sub MoveRight()
    Player.MovePlayer 1 ' Corrected to call MovePlayer from Player module
End Sub

Sub Shoot()
    bullet.MoveBullet
End Sub


Sub ResetKeys()
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    Application.OnKey "{UP}", ""
End Sub
