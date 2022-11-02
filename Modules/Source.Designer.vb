Imports System.Net.Mime.MediaTypeNames

Public rinc As Integer, cinc As Integer
Public vis As Integer
Public steps As Integer

Public isPickedUp As Boolean

Public trap As Integer
Public key As Integer
Public wall As Integer
Public rock As Integer
Public shrub As Integer
Public flower As Integer
Public mushroom As Integer
Public shop As Integer
Public firefly As Integer
Public battery As Integer
Public puddle As Integer
Public escape As Integer
Public gate As Integer


Dim r() As Integer, c() As Integer
Sub StartGame()
    Cells.Clear
    Range("E5:AN32").Interior.Color = vbBlack
    Range("E5:AN32").Font.Size = 18
    ' Non Collidable
    shrub = 1
    rock = 2
    wall = 3
    shop = 4
    flower = 5
    mushroom = 6
    puddle = 7
    firefly = 8
    ' Collidable
    trap = 9
    battery = 10
    key = 11
    escape = 12
    ' Changing states
    gate = 13

    Range("AA18").Value = trap
    Range("AA26:AA28").Value = gate
    Range("J8").Value = mushroom
    Range("O8").Value = puddle
    Range("S8").Value = rock
    Range("X8").Value = shrub
    Range("AB8").Value = flower
    Range("AG8").Value = shop
    Range("AK8").Value = firefly
    Range("AA18").Font.Color = vbRed
    Range("X20").Value = key
    Range("X20").Font.ColorIndex = 26
    Range("H14").Value = battery
    Range("H14").Font.Color = vbGreen
    Range("A1:AR4").Value = wall
    Range("AO5:AR36").Value = wall
    Range("A33:AN36").Value = wall
    Range("A5:D32").Value = wall
    Range("A1:AR4").Interior.Color = vbBlack
    Range("AO5:AR36").Interior.Color = vbBlack
    Range("A33:AN36").Interior.Color = vbBlack
    Range("A5:D32").Interior.Color = vbBlack
    Range("AW3").Font.Size = 20
    ReDim r(1)
    ReDim c(1)
    r(0) = 10
    c(0) = 10
    rinc = 0 : cinc = 0
    vis = 1
    steps = 0
    isPickedUp = False
    bindKeys
    ActionKey()
    ShowVis()
    ShowPlayer()
    Hit()
    interact()

End Sub
Sub ActionKey()
    Application.OnKey "{RETURN}", "interact"
End Sub
Sub ShowPlayer()
    Cells(r(0), c(0)).Interior.Color = vbRed
End Sub
Sub ShowVis()
    Range(Cells(r(0) - vis, c(0) - vis), Cells(r(0) + vis, c(0) + vis)).Interior.ColorIndex = 15
End Sub

Sub MovePlayer()
    If rinc <> 0 Or cinc <> 0 Then
        Cells(r(0), c(0)).Interior.ColorIndex = 15
        If (Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlack Or Cells(r(0) + rinc, c(0) + cinc).Interior.ColorIndex = 15) Then
            If (Cells(r(0) + rinc, c(0) + cinc).Value >= 9 Or Cells(r(0) + rinc, c(0) + cinc).Value = 0) Then
                If (Cells(r(0) + rinc, c(0) + cinc).Value <> gate Or isPickedUp = True) Then
                    r(0) = r(0) + rinc
                    c(0) = c(0) + cinc
                    steps = steps + 1
                    Range("B30").Value = steps

                    If (rinc = 0 And cinc = 1) Then
                        Range(Cells(r(0) - vis + -rinc, c(0) - vis + -cinc), Cells(r(0) + vis, c(0) - (vis * -cinc))).Interior.Color = vbBlack
                    End If
                    If (cinc = 0 And rinc = 1) Then
                        Range(Cells(r(0) - vis + -rinc, c(0) - vis + -cinc), Cells(r(0) - (vis * -rinc), c(0) + vis)).Interior.Color = vbBlack
                    End If
                    If (cinc = 0 And rinc = -1) Then
                        Range(Cells(r(0) + vis + -rinc, c(0) + vis + -cinc), Cells(r(0) + (vis * rinc), c(0) - vis)).Interior.Color = vbBlack
                    End If
                    If (cinc = -1 And rinc = 0) Then
                        Range(Cells(r(0) + vis + -rinc, c(0) + vis + -cinc), Cells(r(0) - vis, c(0) - (vis * -cinc))).Interior.Color = vbBlack
                    End If
                End If
            End If
        End If

        Hit()
        Recharge()
        ShowVis()
        ShowPlayer()

        'Decay
    End If
End Sub
Sub interact()
    If Cells(r(0), c(0)).Value = key Then
        isPickedUp = True
        Cells(r(0), c(0)).Value = Null
        Range("AW3").Value = "Permission Increased"
    End If
    If Cells(r(0), c(0) - 1).Value = rock Or Cells(r(0), c(0) + 1).Value = rock Or Cells(r(0) + 1, c(0)).Value = rock Or Cells(r(0) - 1, c(0)).Value = rock Then
        Range("AW3").Value = "This is a rock"
    End If
    If Cells(r(0), c(0) - 1).Value = shrub Or Cells(r(0), c(0) + 1).Value = shrub Or Cells(r(0) + 1, c(0)).Value = shrub Or Cells(r(0) - 1, c(0)).Value = shrub Then
        Range("AW3").Value = "This is a shrub"
    End If
    If Cells(r(0), c(0) - 1).Value = flower Or Cells(r(0), c(0) + 1).Value = flower Or Cells(r(0) + 1, c(0)).Value = flower Or Cells(r(0) - 1, c(0)).Value = flower Then
        Range("AW3").Value = "This is a flower"
    End If
    If Cells(r(0), c(0) - 1).Value = shop Or Cells(r(0), c(0) + 1).Value = shop Or Cells(r(0) + 1, c(0)).Value = shop Or Cells(r(0) - 1, c(0)).Value = shop Then
        Range("AW3").Value = "This is a shop"
    End If
    If Cells(r(0), c(0) - 1).Value = firefly Or Cells(r(0), c(0) + 1).Value = firefly Or Cells(r(0) + 1, c(0)).Value = firefly Or Cells(r(0) - 1, c(0)).Value = firefly Then
        Range("AW3").Value = "This is a firefly"
    End If
    If Cells(r(0), c(0) - 1).Value = puddle Or Cells(r(0), c(0) + 1).Value = puddle Or Cells(r(0) + 1, c(0)).Value = puddle Or Cells(r(0) - 1, c(0)).Value = puddle Then
        Range("AW3").Value = "This is a puddle"
    End If
    If Cells(r(0), c(0) - 1).Value = mushroom Or Cells(r(0), c(0) + 1).Value = mushroom Or Cells(r(0) + 1, c(0)).Value = mushroom Or Cells(r(0) - 1, c(0)).Value = mushroom Then
        Range("AW3").Value = "This is a mushroom"
    End If
    If Cells(r(0), c(0) - 1).Value = gate Or Cells(r(0), c(0) + 1).Value = gate Or Cells(r(0) + 1, c(0)).Value = gate Or Cells(r(0) - 1, c(0)).Value = gate Then
        Range("AW3").Value = "This is a gate"
    End If
    If Cells(r(0), c(0) - 1).Value = wall Or Cells(r(0), c(0) + 1).Value = wall Or Cells(r(0) + 1, c(0)).Value = wall Or Cells(r(0) - 1, c(0)).Value = wall Then
        Range("AW3").Value = "This is a wall"
    End If
End Sub
Sub Hit()
    If Cells(r(0), c(0)).Value = trap And vis > 0 Then
        vis = vis - 1
        Cells(r(0), c(0)).Value = Null
    End If
End Sub
Sub Recharge()
    If Cells(r(0), c(0)).Value = battery And vis < 2 Then
        vis = 2
        Cells(r(0), c(0)).Value = Null
    End If
End Sub
Sub Decay()
    If (steps Mod 100 = 0 And vis > 0) Then
        vis = vis - 1
    End If
End Sub
