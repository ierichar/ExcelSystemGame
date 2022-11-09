Public gameHeight As Integer, gameRow As Integer

Public rinc As Integer, cinc As Integer
Public vis As Integer
Public health As Integer

Public le_isRevealed As Boolean
Public le_isDestroyed As Boolean

Public trap As Integer
Public ptrap As Integer
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
Public usb As Integer
Public footprints As Integer

Public rockSearch As Boolean
Public shrubSearch As Boolean
Public flowerSearch As Boolean
Public mushroomSearch As Boolean
Public puddleSearch As Boolean
Public isHalfway As Boolean

Public rockRevealed As Boolean


Public rockRefresh As Integer
Public shrubRefresh As Integer
Public mushroomRefresh As Integer
Public flowerRefresh As Integer
Public puddleRefresh As Integer

Public rockVisible As Boolean


Public lightData As Integer
Public spaceDiscovered As Integer
Public authorityLevel As Integer
Public level As Integer

Dim r() As Integer, c() As Integer
Dim le_r() As Integer, le_c() As Integer

Dim rockPic As Shape

Sub StartGame()
    'Set width/height of cells
    gameWidth = 6
    gameHeight = 30
    Range("A:BH").ColumnWidth = gameWidth
    Range("1:40").RowHeight = gameHeight

    'Sets Values for enviornment

    ' Non Collidable
    shrub = 1
    rock = 2
    wall = 3
    shop = 4
    flower = 5
    mushroom = 6
    puddle = 7

    ' Collidable
    trap = 9
    firefly = 10
    key = 11
    usb = 12
    ptrap = 16 'May need to use fill tracking

    ' Changing states
    gate = 13
    escape = 14
    footprints = 15

    'loads in the level 1 values
    level = 0
    LoadLevel (0)

    'Player Variables
    ReDim r(1)
    ReDim c(1)
    r(0) = 10
    c(0) = 10
    rinc = 0: cinc = 0
    health = 3
    vis = 0
    authorityLevel = 0

    'Enemy Values
    ReDim le_r(1)
    ReDim le_c(1)
    le_r(0) = 16: le_c(0) = 16
    le_rinc = 0: le_cinc = 0
    le_isRevealed = False
    le_isDestroyed = False

    'Player Trap Values
    ReDim pt_r(1)
    ReDim pt_c(1)
    pt_r(0) = 0: pt_c(0) = 0
    pt_isPlaced = False

    'bind keys and render player
    bindKeys
    ShowVis
    ShowPlayer
    ShowEnemy
    AddUI

End Sub

'----------------------------SHOW PLAYER AND VISION-------------------------------------------------
Sub ShowPlayer()
    Cells(r(0), c(0)).Interior.Color = vbRed
End Sub
Sub ShowEnemy()
    Debug.Print ("revealed = " & le_isRevealed & ", destroyed = " & le_isDestroyed)
    If (le_isRevealed = True And le_isDestroyed = False) Then
        Cells(le_r(0), le_c(0)).Interior.Color = vbGreen
        'ElseIf (le_isRevealed = True And le_isDestroyed = False) Then
        'Cells(le_r(0), le_c(0)).Interior.Color = vbGray
    Else: Cells(le_r(0), le_c(0)).Interior.Color = vbBlack
    End If
End Sub

'=================================ShowVis=====================================
Sub ShowVis()
    Range(Cells(r(0) - vis, c(0) - vis), Cells(r(0) + vis, c(0) + vis)).Interior.ColorIndex = 15
End Sub

'--------------------------UPDATE AND MOVE PLAYER----------------------------------------------------
Sub MovePlayer()
    ' if the player moves then run this
    If rinc <> 0 Or cinc <> 0 Then
        'sets past position to vision color
        Cells(r(0), c(0)).Interior.ColorIndex = 15

        'footprint placed and spaceDiscovered incremented
        If Cells(r(0), c(0)).Value <> footprints Then
            If (Cells(r(0), c(0)).Value <> ptrap) Then
                Cells(r(0), c(0)).Value = footprints
            End If
            spaceDiscovered = spaceDiscovered + 1
            Range("C37").Value = spaceDiscovered
        End If


        'if the cell you are moving to is black or the vision color then run
        If (Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlack Or Cells(r(0) + rinc, c(0) + cinc).Interior.ColorIndex = 15) Then
            'collision  condition
            If (Cells(r(0) + rinc, c(0) + cinc).Value >= 9 Or Cells(r(0) + rinc, c(0) + cinc).Value = 0) Then
                ' permission condition
                If (Cells(r(0) + rinc, c(0) + cinc).Value <> gate Or authorityLevel = 1) Then
                    r(0) = r(0) + rinc
                    c(0) = c(0) + cinc
                    'setting past vision range to black
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

        'updating functions
        Collide
        ShowVis
        ShowPlayer
        ShowEnemy
        MoveEnemy
        UpdateUI
        AuthorityLevelCheck (level)
        SearchRefresh
        RenderImages
        ImgToUI
    End If
End Sub

'==================RevealEnemy===============
'Pre: r(0), c(0), le_r(0), le_r(0), vis
Function RevealEnemy()
    Debug.Print ("Checking reveal...")
    If (Abs(r(0) - le_r(0)) <= vis) And (Abs(c(0) - le_c(0)) <= vis) Then
        le_isRevealed = True
        If (le_isDestroyed = False) Then
            Range("B38").Value = "A light-hungry moth has spotted you! Run away or trap it!"
        End If
        RevealEnemy = True
    Else: le_isRevealed = False
        If (le_isDestroyed = False) Then
            Range("B38").Value = "The moth loses you in the darkness."
        End If
        RevealEnemy = False
    End If
End Function

'==================MoveEnemy=================
'Pre: r(0), c(0), le_r(0), le_r(0)
Sub MoveEnemy()
    Debug.Print ("Moving enemy...")
    If (Cells(le_r(0), le_c(0)).Value = 16) Then
        le_isDestroyed = True
    End If

    If (RevealEnemy() = True And le_isDestroyed = False) Then
        Dim xDiff As Integer, yDiff As Integer

        yDiff = r(0) - le_r(0)
        xDiff = c(0) - le_c(0)
        Debug.Print ("xDiff val: " & xDiff)
        Debug.Print ("yDiff val: " & yDiff)

        If (yDiff >= 0 And xDiff >= 0) Then
            If (yDiff > xDiff) Then
                ' Check if future move is open
                If (Cells(le_r(0) + 1, le_c(0)).Value <> wall) Then
                    le_r(0) = le_r(0) + 1
                End If
            Else
                If (Cells(le_r(0), le_c(0) + 1).Value <> wall) Then
                    le_c(0) = le_c(0) + 1
                End If
            End If
        ElseIf (yDiff >= 0 And xDiff <= 0) Then
            If (Abs(yDiff) > Abs(xDiff)) Then
                If (Cells(le_r(0) + 1, le_c(0)).Value <> wall) Then
                    le_r(0) = le_r(0) + 1
                End If
            Else
                If (Cells(le_r(0), le_c(0) - 1).Value <> wall) Then
                    le_c(0) = le_c(0) - 1
                End If
            End If
        ElseIf (yDiff <= 0 And xDiff >= 0) Then
            If (Abs(yDiff) > Abs(xDiff)) Then
                If (Cells(le_r(0) - 1, le_c(0)).Value <> wall) Then
                    le_r(0) = le_r(0) - 1
                End If
            Else
                If (Cells(le_r(0), le_c(0) + 1).Value <> wall) Then
                    le_c(0) = le_c(0) + 1
                End If
            End If
        ElseIf (yDiff <= 0 And xDiff <= 0) Then
            If (yDiff < xDiff) Then
                If (Cells(le_r(0) - 1, le_c(0)).Value <> Null) Then
                    le_r(0) = le_r(0) - 1
                End If
            Else
                If (Cells(le_r(0), le_c(0) - 1).Value <> Null) Then
                    le_c(0) = le_c(0) - 1
                End If
            End If
        End If
        If (xDiff = 0 And yDiff = 0) Then
            Debug.Print ("reduce player health")
            health = health - 1
            'Testing 1 hit for now
            le_isDestroyed = True
            Range("B38").Value = "The moth strikes! You feel a sharp pain as it disppates."
        End If
    End If

End Sub


'------------------------------INTERACTION CHECKS----------------------------------------------
Sub interact()
    'TODO Make function for each interaction check to make it look prettier

    If Cells(r(0), c(0) - 1).Value = rock Or Cells(r(0), c(0) + 1).Value = rock Or Cells(r(0) + 1, c(0)).Value = rock Or Cells(r(0) - 1, c(0)).Value = rock Then
        If rockSearch = True Then
            Range("B38").Value = "A plain rock"
        End If
        If rockSearch = False Then
            Range("B38").Value = "You find a rock. It looks like its shimmering. You reach out to it and feel an energy transferred to you."
            lightData = lightData + 10
            Range("BB15").Value = lightData
            rockSearch = True
        End If
    End If


    If Cells(r(0), c(0) - 1).Value = shrub Or Cells(r(0), c(0) + 1).Value = shrub Or Cells(r(0) + 1, c(0)).Value = shrub Or Cells(r(0) - 1, c(0)).Value = shrub Then
        If shrubSearch = True Then
            Range("B38").Value = "A shrub. The berries look dull"
        End If
        If shrubSearch = False Then
            Range("B38").Value = "You find a shrub. The colorful berries shine bright giving off a glow of energy"
            lightData = lightData + 10
            Range("BB15").Value = lightData
            shrubSearch = True
        End If
    End If
    If Cells(r(0), c(0) - 1).Value = flower Or Cells(r(0), c(0) + 1).Value = flower Or Cells(r(0) + 1, c(0)).Value = flower Or Cells(r(0) - 1, c(0)).Value = flower Then
        If flowerSearch = True Then
            Range("B38").Value = "The same old flower"
        End If
        If flowerSearch = False Then
            Range("B38").Value = "You find a flower. You lean down to sniff it and feel a burst of energy within you"
            lightData = lightData + 10
            Range("BB15").Value = lightData
            flowerSearch = True
        End If

    End If
    If Cells(r(0), c(0) - 1).Value = shop Or Cells(r(0), c(0) + 1).Value = shop Or Cells(r(0) + 1, c(0)).Value = shop Or Cells(r(0) - 1, c(0)).Value = shop Then
        Range("B38").Value = "You find a shop but there is a painted sign that says OuT fOr LUnCh"
        ' This is Joseph testing the shop for testing purposes, feel free to comment out the line for now
        'UserForm1.Show
    End If

    If Cells(r(0), c(0) - 1).Value = puddle Or Cells(r(0), c(0) + 1).Value = puddle Or Cells(r(0) + 1, c(0)).Value = puddle Or Cells(r(0) - 1, c(0)).Value = puddle Then
        If puddleSearch = True Then
            Range("B38").Value = "The same puddle as before but only your reflection stares back at you."
        End If
        If puddleSearch = False Then
            Range("B38").Value = "You find a puddle. Instead of your reflection it gives off an aura of energy"
            lightData = lightData + 10
            Range("BB15").Value = lightData
            puddleSearch = True
        End If

    End If
    If Cells(r(0), c(0) - 1).Value = mushroom Or Cells(r(0), c(0) + 1).Value = mushroom Or Cells(r(0) + 1, c(0)).Value = mushroom Or Cells(r(0) - 1, c(0)).Value = mushroom Then
        If mushroomSearch = True Then
            Range("B38").Value = "The same mushroom but now much darker tones fill the spots on the cap"
        End If
        If mushroomSearch = False Then
            Range("B38").Value = "You find a mushroom. The spots on the cap seem to be glowing and give off energy"
            lightData = lightData + 10
            Range("BB15").Value = lightData

        End If
    End If
    If Cells(r(0), c(0) - 1).Value = gate Or Cells(r(0), c(0) + 1).Value = gate Or Cells(r(0) + 1, c(0)).Value = gate Or Cells(r(0) - 1, c(0)).Value = gate Then
        If authorityLevel = 0 Then
            Range("B38").Value = "You find what seems like a gate. It has the same glow as the things around you but does not give it off."
        End If
        If authorityLevel = 1 Then
            Range("B38").Value = "The gate has lost its glow and is swung open"
        End If
    End If
    If Cells(r(0), c(0) - 1).Value = wall Or Cells(r(0), c(0) + 1).Value = wall Or Cells(r(0) + 1, c(0)).Value = wall Or Cells(r(0) - 1, c(0)).Value = wall Then
        Range("B38").Value = "A hard sturdy wall. Looks impenetrable"
    End If
End Sub

Sub SearchRefresh()
    If (rockSearch = False) Then
        rockRefresh = spaceDiscovered
    End If
    If (rockSearch = True And spaceDiscovered = rockRefresh + 20) Then
        rockSearch = False
    End If

    If (shrubSearch = False) Then
        shrubRefresh = spaceDiscovered
    End If
    If (shrubSearch = True And spaceDiscovered = shrubRefresh + 20) Then
        shrubSearch = False
    End If

    If (flowerSearch = False) Then
        flowerRefresh = spaceDiscovered
    End If
    If (flowerSearch = True And spaceDiscovered = flowerRefresh + 20) Then
        flowerSearch = False
    End If

    If (mushroomSearch = False) Then
        mushroomRefresh = spaceDiscovered
    End If
    If (mushroomSearch = True And spaceDiscovered = mushroomRefresh + 20) Then
        mushroomSearch = False
    End If

    If (puddleSearch = False) Then
        puddleRefresh = spaceDiscovered
    End If
    If (puddleSearch = True And spaceDiscovered = puddleRefresh + 20) Then
        puddleSearch = False
    End If

End Sub

'-------------------------------Place Item-------------------------------------
'Click on item in inventory, then press p to place it
Sub placeItem()
    Debug.Print ("Placing item")
    If (ActiveCell.Value = ptrap) Then
        Cells(r(0), c(0)).Value = ActiveCell.Value
        ActiveCell.Value = Null
    End If
End Sub

'--------------------------------PLAYER COLLISION IN CELL ----------------------------------------
Sub Collide()
    If Cells(r(0), c(0)).Value = trap And vis > 0 Then
        vis = vis - 1
        Cells(r(0), c(0)).Value = Null
        Range("B38").Value = "YOU STEPPED ON A TRAP: Vision level decreased"
    End If

    If Cells(r(0), c(0)).Value = escape Then
        If level = 0 Then
            Range("B38").Value = "This seems like the way out! Next Level Reached"
            level = level + 1
            LoadLevel (level)
        End If
    End If

    If Cells(r(0), c(0)).Value = firefly And vis < 3 Then
        vis = vis + 1
        Cells(r(0), c(0)).Value = Null
        Range("B38").Value = "YOU CAPTURED A FIREFLY: USB recharged!"
        Cells(r(0), c(0)).Font.Color = vbBlack
    End If
    If Cells(r(0), c(0)).Value = usb And vis = 0 Then
        vis = vis + 2
        Cells(r(0), c(0)).Value = Null
        MsgBox "USB FOUND: Vision capabilites unlocked"
        Cells(r(0), c(0)).Font.Color = vbBlack
        UpdateInventory
    End If
End Sub

Function AuthorityLevelCheck(level As Integer)
    If level = 0 Then
        If lightData >= 20 And spaceDiscovered >= 75 And isHalfway = False Then
            MsgBox "The USB device in your possesion whirls. A bar on the face of the device is half way full"
            isHalfway = True
        End If
        If lightData >= 50 And spaceDiscovered >= 125 And authorityLevel = 0 Then
            MsgBox "The USB device in your possesion whirls again. It flashes with the words AUTHORITY LEVEL INCREASED"
            authorityLevel = 1
        End If
    End If
End Function
Sub RenderImages()

    Dim visRng As Range
    Dim levelRng As Range
    Dim cell As Range
    Dim pic As Picture
    For Each pic In Sheets("Sheet1").Pictures
        pic.Delete
    Next pic

    Set visRng = Range(Cells(r(0) - vis, c(0) - vis), Cells(r(0) + vis, c(0) + vis))
    Set levelRng = Range("A1:AR36")

    For Each cell In levelRng
        If cell.Value = firefly Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\firefly.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If

        If cell.Value = usb Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\usb.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If
    Next cell

    For Each cell In visRng
        If cell.Value = firefly Then
            cell.Font.ColorIndex = gameHeight - 10
        End If

        If cell.Value = rock Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\rock.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If

        If cell.Value = shrub Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\shrub.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If

        If cell.Value = flower Then
            cell.Font.ColorIndex = 15
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\flower.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If

        If cell.Value = mushroom Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\mushroom.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If

        If cell.Value = puddle Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\puddle.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If


    Next cell




End Sub


'------------------------------------------UI ADDING AND UPDATING---------------------------------
Sub AddUI()
    ' Health bar
    Range("AT4", "bd34").Interior.Color = RGB(239, 222, 205)
    Range("AY4").Value = "Health:"
    Range("BA4").Value = health
    ' Batteries
    Range("AY6").Value = "Light Strength:"
    Range("AZ11").Value = 1
    Range("AZ10").Value = 2
    Range("AZ9").Value = 3
    Range("AZ8").Value = 4
    Range("AZ11", "AZ7").Font.Size = (gameHeight - 5)

    ' Item inventory Area
    Range("AU28", "BC33").Interior.Color = RGB(245, 245, 220)
    'Player Traps
    Range("AU28").Value = ptrap

    ' Currency
    Range("AY13").Value = "Bits:"
    Range("AY15").Value = "Light Data: "
    ' Font Size and Center Alignment
    Range("AT4", "bd34").HorizontalAlignment = xlCenter
    Range("AT4", "bd34").Font.Size = (gameHeight - 5)

End Sub
Sub UpdateUI()
    ' Update displayed health
    Range("BA4").Value = health

    ' Update visibility
    If vis = 0 Then
        Range("AY11").Interior.Color = RGB(239, 222, 205)
    End If
    If vis = 1 Then
        Range("AY11").Interior.Color = vbRed
        Range("AY10").Interior.Color = RGB(239, 222, 205)
    End If
    If vis = 2 Then
        Range("AY11", "AY10").Interior.Color = RGB(255, 165, 0)
        Range("AY9").Interior.Color = RGB(239, 222, 205)
    End If
    If vis = 3 Then
        Range("AY11", "AY10").Interior.Color = vbGreen
        Range("AY9").Interior.Color = vbGreen
    End If
End Sub

Sub UpdateInventory()
    ' Adds values to the Inventory, so that RemoveInventory can delete the image and value associated much easier
    Dim i As Integer, j As Integer, count As Integer
    Dim usbRange As Range
    For i = 47 To 55
        If IsEmpty(Cells(28, i)) Then
            ' This adds USB to the inventory
            Cells(28, i).Value = 12
            Exit For
        End If
    Next i
End Sub
Sub ImgToUI()
    ' Adds images to the Inventory associated with the Range, only works for usb ATM
    Dim invRange As Range
    Dim cell As Range
    Set invRange = Range(Cells(28, 47), Cells(28, 55))
    For Each cell In invRange
        If cell.Value = usb Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\usb.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5
        End If
    Next cell
End Sub
Sub RemoveInventory(invValue As Integer)
    'Deletes the Image associated with the given value
    Dim i As Integer
    For i = 47 To 55
        If Cells(28, i).Value = invValue Then
            Cells(28, i).Value = Null
            Exit For
        End If
    Next i
End Sub
'----------------------------------------------LOAIDNG LEVELS-----------------------------------------
Function LoadLevel(level As Integer)
    'Clear All Values
    Cells.Clear
    Dim pic As Picture
    For Each pic In Sheets("Sheet1").Pictures
        pic.Delete
    Next pic



    'Set Bound of Level
    Range("E5:AN32").Interior.Color = vbBlack
    Range("E5:AN32").Font.Size = (gameHeight - 5)
    Range("A1:AR4").Value = wall
    Range("AO5:AR36").Value = wall
    Range("A33:AN36").Value = wall
    Range("A5:D32").Value = wall
    Range("A1:AR4").Interior.Color = vbBlack
    Range("AO5:AR36").Interior.Color = vbBlack
    Range("A33:AN36").Interior.Color = vbBlack
    Range("A5:D32").Interior.Color = vbBlack

    'Envir searching variables
    isHalfway = False
    rockSearch = False
    shrubSearch = False
    flowerSearch = False
    mushroomSearch = False
    puddleSearch = False

    rockVisible = False

    rockRevealed = False

    rockRefresh = 0
    shrubRefresh = 0
    flowerRefresh = 0
    mushroomRefresh = 0
    puddleRefresh = 0


    'Level Progression
    lightData = 0
    spaceDiscovered = 0


    If level = 0 Then



        Range("AA32").Value = escape
        Range("T5:T32").Value = trap
        Range("AA20").Value = firefly
        Range("AA26:AC26").Value = gate
        Range("Z26:Z32").Value = wall
        Range("AD26:AD32").Value = wall
        Range("AB28").Value = trap
        Range("J8").Value = rock
        Range("R15").Value = shrub
        Range("K18").Value = rock
        Range("W11").Value = shrub
        Range("AE24").Value = rock
        Range("AG8").Value = shop
        Range("N27").Value = firefly
        Range("I10").Value = usb
        Range("B38").Font.Size = gameHeight
        Range("BB15").Font.Size = (gameHeight - 10)

        If Range("I10").Value = usb Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\usb.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = Range("I10").Top
            Image.Left = Range("I10").Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5
        End If
        'RenderImages

    End If

    If level = 1 Then

        'Redraw UI
        AddUI
        UpdateUI
        UpdateInventory

        'Player Pos
        r(0) = 10
        c(0) = 10

        'Enemy
        le_r(0) = 16: le_c(0) = 16
        le_isRevealed = False



        'Map Blocking
        Range("AM5").Value = escape
        Range("AJ7").Value = trap
        Range("R13").Value = firefly
        Range("R13").Font.ColorIndex = 6
        'Range("AA26:AC26").Value = gate
        'Range("Z26:Z32").Value = wall
        'Range("AD26:AD32").Value = wall
        Range("AB28").Value = trap
        Range("J8").Value = rock
        Range("AG29").Value = puddle
        Range("F9").Value = mushroom
        Range("P11").Value = shrub
        Range("G20").Value = flower
        Range("AG20").Value = shop
        Range("F8").Value = firefly
        Range("F8").Font.ColorIndex = 6
        Range("B38").Font.Size = gameHeight
        Range("BB15").Font.Size = (gameHeight - 10)

        ShowVis
        ShowPlayer
        ShowEnemy

    End If


End Function






