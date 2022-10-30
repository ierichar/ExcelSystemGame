Public rinc As Integer, cinc As Integer
Public vis As Integer
Public steps As Integer
Public isPickedUp As Boolean

Dim r() As Integer, c() As Integer
Sub StartGame()
    Cells.Clear
    Range("C3:AN27").Interior.Color = vbBlack
    Range("AA14").Value = 3
    Range("X16").Value = 5
    Range("E20:AA21").Value = 6
    Range("E21:F27").Value = 6
    Range("I25:AA27").Value = 6
    ReDim r(1)
    ReDim c(1)
    r(0) = 10
    c(0) = 10
    rinc = 0 : cinc = 0
    vis = 2
    steps = 0
    isPickedUp = False
    bindKeys
    ActionKey()
    ShowPlayer()
    ShowVis()
    Hit()
    Hallway()
    HallwaySpawn()
    pickup()

End Sub
Sub ActionKey()
    Application.OnKey "{RETURN}", "pickup"
End Sub
Sub ShowPlayer()
    Cells(r(0), c(0)).Interior.Color = vbRed
End Sub
Sub ShowVis()
    If vis = 2 Then
        If (Cells(r(0) + 2, c(0)).Interior.Color = vbBlack And Cells(r(0) + 2, c(0)).Value < 6) Then
            Cells(r(0) + 2, c(0)).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 2, c(0) + 2).Interior.Color = vbBlack And Cells(r(0) + 2, c(0) + 2).Value < 6) Then
            Cells(r(0) + 2, c(0) + 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0), c(0) + 2).Interior.Color = vbBlack And Cells(r(0), c(0) + 2).Value < 6) Then
            Cells(r(0), c(0) + 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 2, c(0)).Interior.Color = vbBlack And Cells(r(0) - 2, c(0)).Value < 6) Then
            Cells(r(0) - 2, c(0)).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 2, c(0) - 2).Interior.Color = vbBlack And Cells(r(0) - 2, c(0) - 2).Value < 6) Then
            Cells(r(0) - 2, c(0) - 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0), c(0) - 2).Interior.Color = vbBlack And Cells(r(0), c(0) - 2).Value < 6) Then
            Cells(r(0), c(0) - 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 2, c(0) - 2).Interior.Color = vbBlack And Cells(r(0) + 2, c(0) - 2).Value < 6) Then
            Cells(r(0) + 2, c(0) - 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 2, c(0) + 2).Interior.Color = vbBlack And Cells(r(0) - 2, c(0) + 2).Value < 6) Then
            Cells(r(0) - 2, c(0) + 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 1, c(0)).Interior.Color = vbBlack And Cells(r(0) + 1, c(0)).Value < 6) Then
            Cells(r(0) + 1, c(0)).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 1, c(0) + 1).Interior.Color = vbBlack And Cells(r(0) + 1, c(0) + 1).Value < 6) Then
            Cells(r(0) + 1, c(0) + 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0), c(0) + 1).Interior.Color = vbBlack And Cells(r(0), c(0) + 1).Value < 6) Then
            Cells(r(0), c(0) + 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0)).Interior.Color = vbBlack And Cells(r(0) - 1, c(0)).Value < 6) Then
            Cells(r(0) - 1, c(0)).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0) - 1).Interior.Color = vbBlack And Cells(r(0) - 1, c(0) - 1).Value < 6) Then
            Cells(r(0) - 1, c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0), c(0) - 1).Interior.Color = vbBlack And Cells(r(0), c(0) - 1).Value < 6) Then
            Cells(r(0), c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 1, c(0) - 1).Interior.Color = vbBlack And Cells(r(0) + 1, c(0) - 1).Value < 6) Then
            Cells(r(0) + 1, c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0) + 1).Interior.Color = vbBlack And Cells(r(0) - 1, c(0) + 1).Value < 6) Then
            Cells(r(0) - 1, c(0) + 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 2, c(0) + 1).Interior.Color = vbBlack And Cells(r(0) - 2, c(0) + 1).Value < 6) Then
            Cells(r(0) - 2, c(0) + 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0) + 2).Interior.Color = vbBlack And Cells(r(0) - 1, c(0) + 2).Value < 6) Then
            Cells(r(0) - 1, c(0) + 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 2, c(0) + 1).Interior.Color = vbBlack And Cells(r(0) + 2, c(0) + 1).Value < 6) Then
            Cells(r(0) + 2, c(0) + 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 1, c(0) + 2).Interior.Color = vbBlack And Cells(r(0) + 1, c(0) + 2).Value < 6) Then
            Cells(r(0) + 1, c(0) + 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 2, c(0) - 1).Interior.Color = vbBlack And Cells(r(0) - 2, c(0) - 1).Value < 6) Then
            Cells(r(0) - 2, c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0) - 2).Interior.Color = vbBlack And Cells(r(0) - 1, c(0) - 2).Value < 6) Then
            Cells(r(0) - 1, c(0) - 2).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 2, c(0) - 1).Interior.Color = vbBlack And Cells(r(0) + 2, c(0) - 1).Value < 6) Then
            Cells(r(0) + 2, c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 1, c(0) - 2).Interior.Color = vbBlack And Cells(r(0) + 1, c(0) - 2).Value < 6) Then
            Cells(r(0) + 1, c(0) - 2).Interior.Color = vbBlue
        End If
    End If

    If vis = 1 Then
        If (Cells(r(0) + 1, c(0)).Interior.Color = vbBlack And Cells(r(0) + 1, c(0)).Value < 6) Then
            Cells(r(0) + 1, c(0)).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 1, c(0) + 1).Interior.Color = vbBlack And Cells(r(0) + 1, c(0) + 1).Value < 6) Then
            Cells(r(0) + 1, c(0) + 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0), c(0) + 1).Interior.Color = vbBlack And Cells(r(0), c(0) + 1).Value < 6) Then
            Cells(r(0), c(0) + 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0)).Interior.Color = vbBlack And Cells(r(0) - 1, c(0)).Value < 6) Then
            Cells(r(0) - 1, c(0)).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0) - 1).Interior.Color = vbBlack And Cells(r(0) - 1, c(0) - 1).Value < 6) Then
            Cells(r(0) - 1, c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0), c(0) - 1).Interior.Color = vbBlack And Cells(r(0), c(0) - 1).Value < 6) Then
            Cells(r(0), c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) + 1, c(0) - 1).Interior.Color = vbBlack And Cells(r(0) + 1, c(0) - 1).Value < 6) Then
            Cells(r(0) + 1, c(0) - 1).Interior.Color = vbBlue
        End If
        If (Cells(r(0) - 1, c(0) + 1).Interior.Color = vbBlack And Cells(r(0) - 1, c(0) + 1).Value < 6) Then
            Cells(r(0) - 1, c(0) + 1).Interior.Color = vbBlue
        End If

    End If
End Sub

Sub MovePlayer()
    If rinc <> 0 Or cinc <> 0 Then
        Cells(r(0), c(0)).Interior.Color = vbBlue
        If (Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlack Or Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlue) Then
            If Cells(r(0) - 1, c(0)).Value = 5 And isPickedUp = True Then
                Cells(r(0) - 1, c(0)).Value = Null
            End If
            If (Cells(r(0) + rinc, c(0) + cinc).Value < 6) Then
                r(0) = r(0) + rinc
                c(0) = c(0) + cinc
                steps = steps + 1
                Range("B30").Value = steps
            End If
        End If
        If (Not Not isPickedUp = True And Cells(r(0) - 1, c(0)).Value <> 6 And Cells(r(0) - 1, c(0)).Value <> 3 And Cells(r(0) - 1, c(0)).Value <> 2 And Cells(r(0) - 1, c(0)).Value <> 1) Then
            Cells(r(0) - 1, c(0)).Value = 5
        End If
        ShowPlayer()
        ShowVis()
        Hit()
        Hallway()
        RoomSpawn()


    End If
End Sub
Sub pickup()
    If Cells(r(0), c(0)).Value = 5 Then
        isPickedUp = True
        Cells(r(0), c(0)).Value = Null
        Cells(r(0) - 1, c(0)).Value = 5
    End If
End Sub
Sub Hit()
    If Cells(r(0), c(0)).Value = 3 And vis > 0 Then
        vis = vis - 1
    End If
End Sub


Sub Hallway()
    If Cells(r(0), c(0)).Value = 1 Then
        Range(Cells(r(0) - 1, c(0) + 1), Cells(r(0) + 1, c(0) + 20)).Interior.Color = vbBlack
        Cells(r(0), c(0) + 20).Value = 2
    End If
End Sub
Sub RoomSpawn()
    If Cells(r(0), c(0)).Value = 2 Then
        ActiveWindow.ScrollColumn = c(0)
        ActiveWindow.ScrollRow = r(0) / 2
        Range(Cells(r(0) - 10, c(0) + 1), Cells(r(0) + 30, c(0) + 30)).Interior.Color = vbBlack
        For i = 15 To 30
            If (Rnd() * 100 > 70) Then
                Cells(i, c(0) + 30).Value = 1
                Exit For
            End If
        Next i
    End If
End Sub


Sub HallwaySpawn()

    For i = 15 To 27
        If (Rnd() * 100 > 70) Then

            Cells(i, 40).Value = 1
            Exit For
        End If

    Next i

End Sub
