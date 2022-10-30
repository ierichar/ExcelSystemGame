Public rinc As Integer, cinc As Integer
Dim r() As Integer, c() As Integer
Sub StartGame()
    Cells.Clear
    Range("B2:AN27").Interior.Color = vbBlack
    'Range("AN14").Value = 1
    ReDim r(1)
    ReDim c(1)
    r(0) = 10
    c(0) = 10
    rinc = 0 : cinc = 0
    bindKeys
    ShowPlayer()
    Hallway()
    HallwaySpawn()

End Sub

Sub ShowPlayer()
    Cells(r(0), c(0)).Interior.Color = vbRed
    If (Cells(r(0) + 1, c(0)).Interior.Color = vbBlack) Then
        Cells(r(0) + 1, c(0)).Interior.Color = vbBlue
    End If
    If (Cells(r(0) + 1, c(0) + 1).Interior.Color = vbBlack) Then
        Cells(r(0) + 1, c(0) + 1).Interior.Color = vbBlue
    End If
    If (Cells(r(0), c(0) + 1).Interior.Color = vbBlack) Then
        Cells(r(0), c(0) + 1).Interior.Color = vbBlue
    End If
    If (Cells(r(0) - 1, c(0)).Interior.Color = vbBlack) Then
        Cells(r(0) - 1, c(0)).Interior.Color = vbBlue
    End If
    If (Cells(r(0) - 1, c(0) - 1).Interior.Color = vbBlack) Then
        Cells(r(0) - 1, c(0) - 1).Interior.Color = vbBlue
    End If
    If (Cells(r(0), c(0) - 1).Interior.Color = vbBlack) Then
        Cells(r(0), c(0) - 1).Interior.Color = vbBlue
    End If
    If (Cells(r(0) + 1, c(0) - 1).Interior.Color = vbBlack) Then
        Cells(r(0) + 1, c(0) - 1).Interior.Color = vbBlue
    End If
    If (Cells(r(0) - 1, c(0) + 1).Interior.Color = vbBlack) Then
        Cells(r(0) - 1, c(0) + 1).Interior.Color = vbBlue
    End If
End Sub

Sub MovePlayer()
    If rinc <> 0 Or cinc <> 0 Then

        Cells(r(0), c(0)).Interior.Color = vbBlack
        If (Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlack Or Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlue) Then
            r(0) = r(0) + rinc
            c(0) = c(0) + cinc
        End If
        ShowPlayer()
        Hallway()
        RoomSpawn()

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
