Sub bindKeys()
    Application.OnKey "{LEFT}", "moveLeft"
Application.OnKey "{RIGHT}", "moveRight"
Application.OnKey "{UP}", "moveUp"
Application.OnKey "{DOWN}", "moveDown"
End Sub

Sub moveLeft()

    cinc = -1
    rinc = 0
    MovePlayer
End Sub
Sub moveRight()
    cinc = 1
    rinc = 0
    MovePlayer
End Sub
Sub moveUp()
    cinc = 0
    rinc = -1
    MovePlayer
End Sub
Sub moveDown()
    cinc = 0
    rinc = 1
    MovePlayer
End Sub
Sub freeKeys()
    Application.OnKey "LEFT"
Application.OnKey "RIGHT"
Application.OnKey "UP"
Application.OnKey "DOWN"
End Sub
