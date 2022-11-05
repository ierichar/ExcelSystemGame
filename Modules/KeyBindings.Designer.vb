Sub bindKeys()
    Application.OnKey "{LEFT}", "moveLeft"
    Application.OnKey "{RIGHT}", "moveRight"
    Application.OnKey "{UP}", "moveUp"
    Application.OnKey "{DOWN}", "moveDown"
    Application.OnKey "{RETURN}", "interact"
End Sub

Sub moveLeft()
    cinc = -1
    rinc = 0
    MovePlayer
    MoveEnemy
End Sub
Sub moveRight()
    cinc = 1
    rinc = 0
    MovePlayer
    MoveEnemy
End Sub
Sub moveUp()
    cinc = 0
    rinc = -1
    MovePlayer
    MoveEnemy
End Sub
Sub moveDown()
    cinc = 0
    rinc = 1
    MovePlayer
    MoveEnemy
End Sub
Sub freeKeys()
    Application.OnKey "LEFT"
    Application.OnKey "RIGHT"
    Application.OnKey "UP"
    Application.OnKey "DOWN"
    Application.OnKey "RETURN"
End Sub

