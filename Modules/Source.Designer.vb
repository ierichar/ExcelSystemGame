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
Public footprints As String
Public potion As Integer
Public potionCount As Integer
Public mothman As Integer
Public maxVis As Integer


Public rockSearch As Boolean
Public shrubSearch As Boolean
Public flowerSearch As Boolean
Public mushroomSearch As Boolean
Public puddleSearch As Boolean
Public isHalfway As Boolean
Public lightDataTut As Boolean
Public mothTut As Boolean
Public doorsTut As Boolean
Public bossMono As Boolean

Public firstHint As Boolean
Public secondHint As Boolean
Public thirdHint As Boolean


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

Public usbFound As Boolean
Public potionBought As Boolean
Public trapCount As Integer

Dim r() As Integer, c() As Integer
Dim le_r() As Integer, le_c() As Integer

Dim rockPic As Shape

Sub StartGame()
    'Set width/height of cells
    Cells.Clear
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
    potion = 17
    mothman = 18
    battery = 19

    ' Changing states
    gate = 13
    escape = 14
    footprints = "."

    lightDataTut = False
    firstHint = False
    secondHint = False
    thirdHint = False



    ReDim pt_r(1)
    ReDim pt_c(1)
    ReDim r(1)
    ReDim c(1)
    ReDim le_r(1)
    ReDim le_c(1)

    'loads in the level 1 values
    level = 0
    LoadLevel (level)

    'Player Variables

    r(0) = 20
    c(0) = 5
    rinc = 0: cinc = 0
    health = 3
    vis = 0
    maxVis = 3
    authorityLevel = 0
    lightData = 0

    'Enemy Values

    'le_r(0) = 16: le_c(0) = 16
    le_rinc = 0: le_cinc = 0
    le_isRevealed = False
    le_isDestroyed = False
    mothTut = False
    doorsTut = False
    bossMono = False

    'Player Trap Values

    pt_r(0) = 0: pt_c(0) = 0
    pt_isPlaced = False

    'bind keys and render player
    bindKeys
    ShowVis
    ShowEnemy
    AddUI
    ShowPlayer

    'Range("AU28").Value = ptrap
    UpdateInventory (ptrap)
    potionCount = 3
    trapCount = 5
    potionBought = False

    'Controls
    Range("BF16:BL27").Font.Size = gameHeight
    Range("BM17").Font.Size = gameHeight
    Range("BM17") = "CONTROLS"
    Range("BF19") = "Arrow Keys: Move"
    Range("BF20") = "Enter: Interact"
    Range("BF21") = "Tab: Lays Trap Down"
    Range("BF22") = "Caps Lock: Uses Health Potion"


End Sub

'----------------------------SHOW PLAYER AND VISION-------------------------------------------------
Sub ShowPlayer()
    If rinc = 0 And cinc = 0 Then
        Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\walkDown1.png"
        Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
        Image.Top = Cells(r(0), c(0)).Top
        Image.Left = Cells(r(0), c(0)).Left
        Image.ShapeRange.Height = gameHeight + 5
        Image.ShapeRange.Width = gameHeight + 5
    End If
    If rinc = 1 Then
        Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\walkDown1.png"
        Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
        Image.Top = Cells(r(0), c(0)).Top
        Image.Left = Cells(r(0), c(0)).Left
        Image.ShapeRange.Height = gameHeight + 5
        Image.ShapeRange.Width = gameHeight + 5
    End If
    If rinc = -1 Then
        Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\walkUp1.png"
        Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
        Image.Top = Cells(r(0), c(0)).Top
        Image.Left = Cells(r(0), c(0)).Left
        Image.ShapeRange.Height = gameHeight + 5
        Image.ShapeRange.Width = gameHeight + 5
    End If
    If cinc = 1 Then
        Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\walkRight.png"
        Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
        Image.Top = Cells(r(0), c(0)).Top
        Image.Left = Cells(r(0), c(0)).Left
        Image.ShapeRange.Height = gameHeight + 5
        Image.ShapeRange.Width = gameHeight + 5
    End If
    If cinc = -1 Then
        Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\walkLeft.png"
        Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
        Image.Top = Cells(r(0), c(0)).Top
        Image.Left = Cells(r(0), c(0)).Left
        Image.ShapeRange.Height = gameHeight + 5
        Image.ShapeRange.Width = gameHeight + 5
    End If


    'Cells(r(0), c(0)).Interior.Color = vbRed
End Sub
Sub ShowEnemy()
    Debug.Print ("revealed = " & le_isRevealed & ", destroyed = " & le_isDestroyed)
    If level > 0 Then
        If (le_isRevealed = True And le_isDestroyed = False) Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\moth.png"
        Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
        Image.Top = Cells(le_r(0), le_c(0)).Top
            Image.Left = Cells(le_r(0), le_c(0)).Left
            Image.ShapeRange.Height = gameHeight + 10
            Image.ShapeRange.Width = gameHeight + 10

            'Cells(le_r(0), le_c(0)).Interior.Color = vbGreen
            'ElseIf (le_isRevealed = True And le_isDestroyed = False) Then
            'Cells(le_r(0), le_c(0)).Interior.Color = vbGray
        End If
        If le_isRevealed = False Or le_isDestroyed Then
            Cells(le_r(0), le_c(0)).Interior.Color = vbBlack
        End If
    End If
End Sub

'=================================ShowVis=====================================
Sub ShowVis()
    Range(Cells(r(0) - vis, c(0) - vis), Cells(r(0) + vis, c(0) + vis)).Interior.ColorIndex = 50
End Sub

'--------------------------UPDATE AND MOVE PLAYER----------------------------------------------------
Sub MovePlayer()
    ' if the player moves then run this
    If rinc <> 0 Or cinc <> 0 Then
        'sets past position to vision color
        'Cells(r(0), c(0)).Interior.ColorIndex = 50

        'footprint placed and spaceDiscovered incremented
        If Cells(r(0), c(0)).Value <> footprints Then
            If (Cells(r(0), c(0)).Value <> ptrap) Then
                Cells(r(0), c(0)).Value = footprints
            End If
            spaceDiscovered = spaceDiscovered + 1
            Range("C37").Value = spaceDiscovered
        End If


        'if the cell you are moving to is black or the vision color then run
        If (Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlack Or Cells(r(0) + rinc, c(0) + cinc).Interior.ColorIndex = 50) Then
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
        Range("A1:AR36").Font.Color = vbBlack
        'updating functions
        Collide
        ShowVis


        UpdateUI
        AuthorityLevelCheck (level)
        SearchRefresh
        RenderImages
        ImgToUI
        ShowEnemy
        MoveEnemy
        ShowPlayer
        GameOverCheck
        BossFight
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
    If level > 0 Then
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
                If mothTut = False Then
                    MsgBox "OUCH!(-1 HP) I need something to stop these moths. That trap in my inventory might help!"
                mothTut = True
                End If
            End If
        End If
    End If
End Sub


'------------------------------INTERACTION CHECKS----------------------------------------------
Sub interact()
    'TODO Make function for each interaction check to make it look prettier
    If Cells(r(0), c(0) - 1).Value = wall Or Cells(r(0), c(0) + 1).Value = wall Or Cells(r(0) + 1, c(0)).Value = wall Or Cells(r(0) - 1, c(0)).Value = wall Then
        Range("B38").Value = "A hard sturdy wall. Looks impenetrable"
    End If

    If lightDataTut = False And level = 1 Then
        MsgBox "Things look different here. They dont feel real. I feel like I can see numbers all around me"
        lightDataTut = True
    End If

    If Cells(r(0), c(0) - 1).Value = rock Or Cells(r(0), c(0) + 1).Value = rock Or Cells(r(0) + 1, c(0)).Value = rock Or Cells(r(0) - 1, c(0)).Value = rock Then
        If rockSearch = True Then
            Range("B38").Value = "A plain rock"
        End If
        If rockSearch = False Then
            If lightDataTut = False Then
                MsgBox "Light Data Collected ... More Light Data Needed ..."
                MsgBox "Wait what!? Light Data? Maybe I should listen to this USB and find some more"
                lightDataTut = True
            End If
            Range("B38").Value = "You find a rock. It looks like its shimmering. You reach out to it and feel an energy transferred to you."
            lightData = lightData + 2
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
            lightData = lightData + 2
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
            lightData = lightData + 5
            Range("BB15").Value = lightData
            flowerSearch = True
        End If

    End If
    If Cells(r(0), c(0) - 1).Value = shop Or Cells(r(0), c(0) + 1).Value = shop Or Cells(r(0) + 1, c(0)).Value = shop Or Cells(r(0) - 1, c(0)).Value = shop Then
        'Range("B38").Value = "You find a shop but there is a painted sign that says OuT fOr LUnCh"

        ' This is Joseph testing the shop for testing purposes, feel free to comment out the line for now
        If level = 3 Then
            MsgBox "Psst. You can do more than you know here. You can change the world"
        End If
        UserForm1.Show
    End If

    If Cells(r(0), c(0) - 1).Value = firefly Or Cells(r(0), c(0) + 1).Value = firefly Or Cells(r(0) + 1, c(0)).Value = firefly Or Cells(r(0) - 1, c(0)).Value = firefly Then
        MsgBox "A firefly! If I get closer I think I can catch it"
    End If

    If Cells(r(0), c(0) - 1).Value = trap Or Cells(r(0), c(0) + 1).Value = trap Or Cells(r(0) + 1, c(0)).Value = trap Or Cells(r(0) - 1, c(0)).Value = trap Then
        MsgBox "This flower looks different than the others. I think I might be able to step over it"
    End If

    If Cells(r(0), c(0) - 1).Value = puddle Or Cells(r(0), c(0) + 1).Value = puddle Or Cells(r(0) + 1, c(0)).Value = puddle Or Cells(r(0) - 1, c(0)).Value = puddle Then
        If puddleSearch = True Then
            Range("B38").Value = "The same puddle as before but only your reflection stares back at you."
        End If
        If puddleSearch = False Then
            Range("B38").Value = "You find a puddle. Instead of your reflection it gives off an aura of energy"
            lightData = lightData + 5
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
            lightData = lightData + 5
            Range("BB15").Value = lightData
            mushroomSearch = True

        End If
    End If
    If Cells(r(0), c(0) - 1).Value = gate Or Cells(r(0), c(0) + 1).Value = gate Or Cells(r(0) + 1, c(0)).Value = gate Or Cells(r(0) - 1, c(0)).Value = gate Then
        If authorityLevel = 0 Then
            If doorsTut = False Then
                MsgBox "I need to open these some how. I can see the exit on the other side"
            End If
            Range("B38").Value = "You find what seems like a gate. It has the same glow as the things around you but does not give it off."
        End If
        If authorityLevel = 1 Then
            Range("B38").Value = "The gate has lost its glow and is swung open"
        End If
    End If

    RenderImages
    ShowPlayer
    
End Sub

Sub SearchRefresh()
    If (rockSearch = False) Then
        rockRefresh = spaceDiscovered
    End If
    If (rockSearch = True And spaceDiscovered > rockRefresh + 4) Then
        rockSearch = False
    End If

    If (shrubSearch = False) Then
        shrubRefresh = spaceDiscovered
    End If
    If (shrubSearch = True And spaceDiscovered > shrubRefresh + 4) Then
        shrubSearch = False
    End If

    If (flowerSearch = False) Then
        flowerRefresh = spaceDiscovered
    End If
    If (flowerSearch = True And spaceDiscovered > flowerRefresh + 4) Then
        flowerSearch = False
    End If

    If (mushroomSearch = False) Then
        mushroomRefresh = spaceDiscovered
    End If
    If (mushroomSearch = True And spaceDiscovered > mushroomRefresh + 4) Then
        mushroomSearch = False
    End If

    If (puddleSearch = False) Then
        puddleRefresh = spaceDiscovered
    End If
    If (puddleSearch = True And spaceDiscovered > puddleRefresh + 4) Then
        puddleSearch = False
    End If

End Sub

'-------------------------------Place Item-------------------------------------
'Click on item in inventory, then press p to place it
Sub placeItem()
    Dim invRange As Range
    Dim cell As Range
    Set invRange = Range(Cells(28, 47), Cells(33, 55))
    For Each cell In invRange
        If cell.Value = ptrap Then
            Cells(r(0), c(0)).Value = ptrap
        End If
    Next cell
    RemoveInventory (ptrap)
    '    Debug.Print ("Placing item")
    '    If (ActiveCell.Value = ptrap) Then
    '        Cells(r(0), c(0)).Value = ActiveCell.Value
    '        ActiveCell.Value = Null
    '    End If
End Sub

Sub usePotion()
    'Debug.Print ("Placing item")
    Dim invRange As Range
    Dim cell As Range
        'Set invRange = Range(Cells(28, 47), Cells(33, 55))
        'For Each cell In invRange
            'If cell.Value = potion Then
    If health < 3 Then
        health = health + 1
        RemoveInventory (potion)
        'ActiveCell.Value = Null
    ElseIf health = 3 Then
        MsgBox ("My health is full at the moment")
    End If
            'End If
        'Next cell

End Sub

'--------------------------------PLAYER COLLISION IN CELL ----------------------------------------
Sub Collide()
    If Cells(r(0), c(0)).Value = trap And vis > 0 Then
        vis = vis - 1
        Cells(r(0), c(0)).Value = Null
        MsgBox "The flower saps your USB energy. Vision Decreased"
    End If



    If Cells(r(0), c(0)).Value = escape Then
        If level = 3 Then
            MsgBox "I DID IT"
            StartGame
        End If
        If level = 2 Then
            MsgBox "This has to be it"
            level = level + 1
            LoadLevel (level)
        End If

        If level = 1 Then
            MsgBox "I have to be getting close to the exit"
            level = level + 1
            LoadLevel (level)
        End If

        If level = 0 Then
            MsgBox "This seems like the way out!"
            level = level + 1
            LoadLevel (level)
        End If

    End If

    If Cells(r(0), c(0)).Value = firefly And vis < maxVis Then
        vis = vis + 1
        Cells(r(0), c(0)).Value = Null
        MsgBox "Oh my god I caught a Firefly! I can see more now!"
        Cells(r(0), c(0)).Font.Color = vbBlack
    End If
    If Cells(r(0), c(0)).Value = firefly And vis = 3 Then
        Range("B38").Value = "USB Light Program Full. Find Upgrade For More Vision"
    End If
    If Cells(r(0), c(0)).Value = usb And vis = 0 Then
        vis = vis + 2
        Cells(r(0), c(0)).Value = Null
        MsgBox "Reading USB ... Light Program Installed ..."
        MsgBox "Huh?! I can see! From a USB? How is that possible"
        MsgBox "Are those gates?? What is going on here?"
        MsgBox "Woah those rocks over there are glowing maybe I should check them out"


        Cells(r(0), c(0)).Font.Color = vbBlack
        usbFound = True
        UpdateInventory (12)
    End If
End Sub

Function AuthorityLevelCheck(level As Integer)
    If level = 0 Then
        If lightData >= 15 And spaceDiscovered >= 30 And isHalfway = False Then
            MsgBox "The USB device in your possesion whirls. A bar on the face of the device is half way full. More Light Data Needed"
            isHalfway = True
        End If
        If lightData >= 25 And spaceDiscovered >= 75 And authorityLevel = 0 Then
            MsgBox "The USB device in your possesion whirls again. It flashes with the words AUTHORITY LEVEL INCREASED"
            MsgBox "I think I heard a gate open up somewhere"
            authorityLevel = 1
        End If
    End If

    If level = 1 Then
        If lightData >= 50 And spaceDiscovered >= 60 And isHalfway = False Then
            MsgBox "The USB device in your possesion whirls. A bar on the face of the device is half way full. More Light Data Needed"
            isHalfway = True
        End If
        If lightData >= 100 And spaceDiscovered >= 100 And authorityLevel = 0 Then
            MsgBox "The USB device in your possesion whirls again. It flashes with the words AUTHORITY LEVEL INCREASED"
            MsgBox "I think I heard a gate open up somewhere"
            authorityLevel = 1
        End If
    End If
    If level = 2 Then
        If lightData >= 125 And spaceDiscovered >= 50 And isHalfway = False Then
            MsgBox "The USB device in your possesion whirls. A bar on the face of the device is half way full. More Light Data Needed"
            isHalfway = True
        End If
        If lightData >= 150 And spaceDiscovered >= 100 And authorityLevel = 0 Then
            MsgBox "The USB device in your possesion whirls again. It flashes with the words AUTHORITY LEVEL INCREASED"
            MsgBox "I think I heard a gate open up somewhere"
            authorityLevel = 1
        End If
    End If
    If level = 3 And bossMono = True Then
        If spaceDiscovered >= 50 And isHalfway = False And Range("U10").Value = mothman Then
            MsgBox "There has to be some way to beat him. If I could just "
            isHalfway = True
        End If
        If spaceDiscovered >= 75 And authorityLevel = 0 And Range("U10").Value = mothman And firstHint = False Then
            MsgBox "Maybe there is something in the shop that could help me"
            firstHint = True
        End If
        If spaceDiscovered >= 100 And authorityLevel = 0 And Range("U10").Value = mothman And secondHint = False Then
            MsgBox "I have the power. I can change anything in this world."
            secondHint = True
        End If
        If spaceDiscovered >= 150 And authorityLevel = 0 And Range("U10").Value = mothman And thirdHint = False Then
            MsgBox "Wait this is excel I can change any of the values. Even if they are infinity"
            thirdHint = True
        End If
    End If
End Function
Sub GameOverCheck()
    If vis = 0 And level > 0 Then
        MsgBox "USB Depleted: GAME OVER"
        StartGame
    End If
    If health = 0 Then
        MsgBox "YOU DIED: GAME OVER"
        StartGame
    End If
End Sub
Sub BossFight()
    If level = 3 And Range("X18") = 0 And bossMono = True And Range("U10").Value = mothman Then
        authorityLevel = 1
        Range("U10") = Null
        MsgBox "YOU DEFEATED THE MOTH MAN! Authority Level Increased"
    End If

    If level = 3 And r(0) < 23 And authorityLevel = 0 Then
        If bossMono = False Then
            MsgBox "Your journey ends here! I am immortal! You will be stuck here forever!"
            bossMono = True
        End If
        Range("X18").Font.Color = vbRed
        Range("U18").Font.Color = vbRed
        Range("U18") = "Health: "
        Range("X18") = "Infinity"
    End If

End Sub
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
        If cell.Value = mothman Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\mothman.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight + 200
            Image.ShapeRange.Width = gameHeight + 200

        End If
        
        If cell.Value = shop Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\lovelandfrog.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                        
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
    
            End If


        If cell.Value = firefly And level = 0 Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\firefly.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If
        If cell.Value = gate Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\door.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If
        If cell.Value = firefly Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\firefly2.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                        
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\firefly3.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                        
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            

        End If

        If cell.Value = usb Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\usb.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If

        If cell.Value = footprints Then
            cell.Font.Color = vbBlack
        End If


    Next cell

    For Each cell In visRng
        If cell.Value = firefly Then
            cell.Font.ColorIndex = 50
        End If
        If cell.Value = ptrap Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\setTrap.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If

        '        If cell.Value = footprints Then
        '            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\fpDown.png"
        '            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
        '            cell.Font.ColorIndex = 50
        '            Image.Top = cell.Top
        '            Image.Left = cell.Left
        '            Image.ShapeRange.Height = gameHeight
        '            Image.ShapeRange.Width = gameHeight + 5
        '
        '        End If

        If cell.Value = trap Then
            If level = 0 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\trap.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                        
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
            End If
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\trap2.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                        
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\trap3.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                        
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
            End If

        End If

        If cell.Value = wall Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\wall.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            cell.Font.ColorIndex = 50
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If

        If cell.Value = ptrap Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\setTrap.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            cell.Font.ColorIndex = 50
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If

        If cell.Value = rock And rockSearch = False And level = 0 Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\rockLight.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            cell.Font.ColorIndex = 50
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If
        If cell.Value = rock And rockSearch = False Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\rock2Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\rock2Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
            End If

        End If

        If cell.Value = rock And rockSearch = True And level = 0 Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\rock.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
            
            cell.Font.ColorIndex = 50
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If
        If cell.Value = rock And rockSearch = True Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\rock2.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\rock3.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight + 5
            End If

        End If

        If cell.Value = shrub And shrubSearch = False And level = 0 Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\shrubLight.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
            
            cell.Font.ColorIndex = 50
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If
        If cell.Value = shrub And shrubSearch = False Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\shrub2Light.PNG"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\shrub3Light.PNG"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If

        End If

        If cell.Value = shrub And shrubSearch = True And level = 0 Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\shrub.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
            
            cell.Font.ColorIndex = 50
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight

        End If
        If cell.Value = shrub And shrubSearch = True Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\shrub2.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\shrub3.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If

        End If


        If cell.Value = flower And flowerSearch = False Then
            If level = 1 Or level = 2 Then
                'cell.Font.ColorIndex = 15
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\flower2Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                'cell.Font.ColorIndex = 15
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\flower3Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If

        End If
        If cell.Value = flower And flowerSearch = True Then
            If level = 1 Or level = 2 Then
                cell.Font.ColorIndex = 15
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\flower2.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                cell.Font.ColorIndex = 15
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\flower3.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If

        End If

        If cell.Value = mushroom And mushroomSearch = False Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\mushroom2Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\mushroom3Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If


        End If
        If cell.Value = mushroom And mushroomSearch = True Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\mushroom2.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\mushroom3.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If

        End If


        If cell.Value = puddle And puddleSearch = False Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\puddle2Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\puddle3Light.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If


        End If
        If cell.Value = puddle And puddleSearch = True Then
            If level = 1 Or level = 2 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 2\puddle2.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If
            If level = 3 Then
                Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\stage 3\puddle3.png"
                Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                
                cell.Font.ColorIndex = 50
                Image.Top = cell.Top
                Image.Left = cell.Left
                Image.ShapeRange.Height = gameHeight
                Image.ShapeRange.Width = gameHeight
            End If

        End If


    Next cell




End Sub



'------------------------------------------UI ADDING AND UPDATING---------------------------------
Sub AddUI()
    ' Health bar
    Range("AT4", "bd34").Interior.Color = RGB(11, 218, 81)
    Range("AY4").Value = "Health:"
    Range("BA4").Value = health
    ' Batteries
    Range("AY6").Value = "Light Strength:"
    Range("AZ11").Value = 1
    Range("AZ10").Value = 2
    Range("AZ9").Value = 3
    'Range("AZ8").Value = 4
    Range("AZ11", "AZ7").Font.Size = (gameHeight - 5)

    ' Item inventory Area
    Range("AU28", "BC33").Interior.Color = RGB(193, 225, 193)
    'Player Traps


    ' Currency
    'Range("AY13").Value = "Bits:"
    Range("AY15").Value = "Light Data: "
    ' Font Size and Center Alignment
    Range("AT4", "bd34").HorizontalAlignment = xlCenter
    'Range("AT4", "bd34").Font.Size = (gameHeight - 5)
    Range("AT4", "bd27").Font.Size = (gameHeight - 5)

End Sub
Sub UpdateUI()
    ' Update displayed health
    Range("BA4").Value = health
    Dim cell As Range
    Set invRange = Range(Cells(28, 47), Cells(28, 55))

    ' Update visibility
    If vis = 0 Then
        Range("AY11").Interior.Color = RGB(11, 218, 81)
    End If
    If vis = 1 Then
        Range("AY11").Interior.Color = vbRed
        Range("AY10").Interior.Color = RGB(11, 218, 81)
    End If
    If vis = 2 Then
        Range("AY11", "AY10").Interior.Color = RGB(255, 165, 0)
        Range("AY9").Interior.Color = RGB(11, 218, 81)
    End If
    If vis = 3 Then
        Range("AY11", "AY10").Interior.Color = RGB(255, 255, 0)
        Range("AY9").Interior.Color = RGB(255, 255, 0)
    End If
    If vis = 4 Then
        Range("AY11", "AY9").Interior.Color = RGB(0, 255, 0)
        Range("AY8").Interior.Color = RGB(0, 255, 0)
    End If

End Sub

Sub UpdateInventory(invValue As Integer)
    ' Adds values to the Inventory, so that RemoveInventory can delete the image and value associated much easier
    Dim cnt As Range
    Dim length As Range
    Set length = Range(Cells(28, 47), Cells(33, 55))
    For Each cnt In length
        If IsEmpty(cnt) Then
            cnt.Value = invValue
            Exit For
        End If
    Next cnt
End Sub
Sub ImgToUI()
    ' Adds images to the Inventory associated with the Range, only works for usb ATM
    Dim invRange As Range
    Dim cell As Range
    Set invRange = Range(Cells(28, 47), Cells(33, 55))
    For Each cell In invRange
        If cell.Value = usb Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\usb.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5
        End If
        If cell.Value = battery Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\battery.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If
        If cell.Value = potion Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\maxhpUP.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
            Image.Top = cell.Top
            Image.Left = cell.Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5

        End If
        If cell.Value = ptrap Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\setTrap.png"
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

    Range("A1:AR36").Clear

    Dim pic As Picture
    For Each pic In Sheets("Sheet1").Pictures
        pic.Delete
    Next pic


    If level < 3 Then
        'Set Bound of Level
        Range("E5:AN32").Interior.Color = vbBlack
        Range("A1:AR4").Value = wall
        Range("AO5:AR36").Value = wall
        Range("A33:AN36").Value = wall
        Range("A5:D32").Value = wall
        Range("A1:AR4").Interior.Color = vbBlack
        Range("AO5:AR36").Interior.Color = vbBlack
        Range("A33:AN36").Interior.Color = vbBlack
        Range("A5:D32").Interior.Color = vbBlack
    End If
    Range("A1:AN39").Font.Size = (gameHeight - 5)

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
    spaceDiscovered = 0
    authorityLevel = 0


    If level = 0 Then
        MsgBox "Where am I? What am I doing here? Why can’t I see anything?"
        MsgBox "I should walk around and see what I can find."
        MsgBox "I think there is something on the ground in front of me, let me head over and pick it up"




        Range("H28:I28").Value = escape

        Range("H27:I27").Value = gate
        Range("F18:F24").Value = wall
        Range("F17:H17").Value = wall
        Range("E15:H15").Value = wall
        Range("G27:G32").Value = wall
        Range("J27:J32").Value = wall
        Range("K27:M32").Value = wall
        Range("N26:P32").Value = wall
        Range("V25:Y25").Value = wall
        Range("AA1:AG32").Value = wall
        Range("AD26:AD32").Value = wall
        Range("Y26:Z27").Value = wall
        Range("Y21:Y24").Value = wall
        Range("X19:X21").Value = wall
        Range("V15:Z15").Value = wall
        Range("Z5:Z11").Value = wall
        Range("Q26").Value = wall
        Range("S25").Value = wall
        Range("V5").Value = wall
        Range("V8").Value = wall
        Range("V11:V10").Value = wall
        Range("R12:R15").Value = wall
        Range("M11:Q11").Value = wall
        Range("I15:L15").Value = wall
        Range("L12:L14").Value = wall
        Range("O14:O15").Value = wall
        Range("H8:L8").Value = wall
        Range("H8:L8").Value = wall
        Range("H10:J11").Value = wall
        Range("F7:F11").Value = wall
        Range("G7").Value = wall
        Range("G11").Value = wall

        Range("X10").Value = firefly

        Range("W15").Value = rock
        Range("V9").Value = rock
        Range("X11").Value = rock
        Range("T25").Value = rock
        Range("Y18").Value = rock
        Range("X21").Value = rock
        Range("Z26").Value = rock
        Range("P29").Value = rock
        Range("R26").Value = rock
        Range("H16").Value = rock
        Range("AE24").Value = rock
        Range("G8").Value = rock
        Range("G10").Value = rock
        Range("E27:F27").Value = rock

        Range("V6:V7").Value = shrub
        Range("W11").Value = shrub
        Range("L9:L11").Value = shrub
        Range("G9").Value = shrub
        Range("V29").Value = shrub
        Range("O13").Value = shrub

        Range("G26").Value = usb
        Range("B38").Font.Size = gameHeight
        Range("BB15").Font.Size = (gameHeight - 10)

        If Range("G26").Value = usb Then
            Image_Location = Application.ActiveWorkbook.Path + "\ExcelArtAssets\usb.png"
            Set Image = Sheets("Sheet1").Pictures.Insert(Image_Location)
                    
            Image.Top = Range("G26").Top
            Image.Left = Range("G26").Left
            Image.ShapeRange.Height = gameHeight
            Image.ShapeRange.Width = gameHeight + 5
        End If
        'RenderImages

    End If

    If level = 1 Then


        'Redraw UI
        AddUI
        UpdateUI
        'UpdateInventory

        'Player Pos
        r(0) = 31
        c(0) = 9

        'Enemy
        le_r(0) = 29: le_c(0) = 16
        '        le_r(0) = 21: le_c(0) = 17
        '        le_r(0) = 21: le_c(0) = 17
        '        le_r(0) = 10: le_c(0) = 25
        '        le_r(0) = 21: le_c(0) = 17
        '        le_r(0) = 10: le_c(0) = 5
        '        le_r(0) = 19: le_c(0) = 7
        le_isRevealed = False

        lightDataTut = False

        'Map Blocking
        Range("K31") = wall
        Range("J29") = wall
        Range("L32") = wall
        Range("J30") = wall
        Range("H28:I28") = wall
        Range("E25:G25") = wall
        Range("F28") = wall
        Range("K23") = wall
        Range("L26") = wall
        Range("M23:M25") = wall
        Range("N21:N23") = wall
        Range("I24:J24") = wall
        Range("K21") = wall
        Range("I20") = wall
        Range("F23") = wall
        Range("K17") = wall
        Range("J18:J19") = wall
        Range("F23") = wall
        Range("F19") = wall
        Range("L16:M16") = wall
        Range("N17:N18") = wall
        Range("O22:P22") = wall
        Range("Q18:Q20") = wall
        Range("R21") = wall
        Range("G17") = wall
        Range("H16") = wall
        Range("S19:S20") = wall
        Range("O14:P14") = wall
        Range("E15:E16") = wall
        Range("E5:AN9") = wall
        Range("J10:J11") = wall
        Range("M10:M11") = wall
        Range("E12:F12") = wall
        Range("R17:S17") = wall
        Range("R16") = wall
        Range("X14") = wall
        Range("X16") = wall
        Range("Y18") = wall
        Range("V13:W13") = wall
        Range("X19:X20") = wall
        Range("Z10:AF32") = wall
        Range("W21") = wall
        Range("U22:V22") = wall
        Range("T23") = wall
        Range("X24") = wall
        Range("X23") = wall
        Range("P25:Q25") = wall
        Range("L28:M28") = wall
        Range("K28") = wall
        Range("N29") = wall
        Range("M30") = wall
        Range("S26") = wall
        Range("U26") = wall
        Range("U29") = wall
        Range("P28") = wall
        Range("O28") = wall
        Range("Y26") = wall
        Range("X27:X28") = wall

        Range("K11:L11") = gate
        Range("K10:L10") = escape


        Range("G32") = rock
        Range("H20") = rock
        Range("G21") = rock
        Range("E19") = rock
        Range("N14") = rock
        Range("G18") = rock
        Range("Q26") = rock

        Range("F32") = shrub
        Range("F22") = shrub
        Range("T14") = shrub
        Range("F22") = shrub
        Range("U13") = shrub
        Range("W16") = shrub
        Range("U27") = shrub
        Range("R30") = shrub

        Range("E32") = puddle
        Range("L14") = puddle
        Range("G13") = puddle
        Range("X15") = puddle
        Range("T26") = puddle
        Range("T29") = puddle
        Range("S30") = puddle


        Range("E28") = flower
        Range("H25") = flower
        Range("G20") = flower
        Range("K14") = flower
        Range("T17") = flower
        Range("U28") = flower

        Range("X13") = mushroom
        Range("W14") = mushroom
        Range("W17") = mushroom
        Range("Q29") = mushroom
        Range("Y27") = mushroom

        Range("G28") = trap
        Range("I15") = trap
        Range("E17") = trap
        Range("X21") = trap
        Range("R24") = trap
        Range("W32") = trap

        Range("V19") = shop

        Range("R20") = firefly
        Range("E18") = firefly
        Range("Y19") = firefly
        Range("K29") = firefly


        Range("B38").Font.Size = gameHeight
        Range("BB15").Font.Size = (gameHeight - 10)

        ShowVis
        ShowPlayer
        ShowEnemy

    End If

    If level = 2 Then
        'Player pos
        r(0) = 8: c(0) = 12

        'Enemy pos and state
        le_r(0) = 8: le_c(0) = 8
        le_isRevealed = False
        le_isDestoryed = False

        Range("G5:AB7") = wall   'top wall
        Range("D8:G28") = wall  'left wall
        Range("AB7:AE28") = wall 'right wall
        Range("H28:X31") = wall 'bottom wall
        Range("AA27") = wall

        Range("I7:K8") = wall
        Range("O8:O10") = wall
        Range("P9") = wall
        Range("X9") = wall
        Range("R10") = wall
        Range("U10") = wall
        Range("Z10") = wall
        Range("I11") = wall
        Range("L11") = wall
        Range("N11") = wall
        Range("R11") = wall
        Range("M12") = wall
        Range("S12:T12") = wall
        Range("Y12") = wall
        Range("N13") = wall
        Range("R13") = wall
        Range("H14") = wall
        Range("J14") = wall
        Range("S14:T14") = wall
        Range("W14") = wall
        Range("T15") = wall
        Range("Y15:Y16") = wall
        Range("Q16:R16") = wall
        Range("U16") = wall
        Range("Y16") = wall
        Range("K17") = wall
        Range("M17:N17") = wall
        Range("Z17") = wall
        Range("S18") = wall
        Range("J19") = wall
        Range("Q19:R19") = wall
        Range("I20") = wall
        Range("M20:N20") = wall
        Range("S20:S22") = wall
        Range("X20:Y20") = wall
        Range("L21") = wall
        Range("R20") = wall
        Range("X20:Y20") = wall
        Range("V21:V22") = wall
        Range("Y22:Z22") = wall
        Range("P23") = wall
        Range("T23:U23") = wall
        Range("J24") = wall
        Range("Q24:Q25") = wall
        Range("S24:S25") = wall
        Range("L25") = wall
        Range("W25:W26") = wall
        Range("Z25") = wall
        Range("R26") = wall
        Range("U26") = wall
        Range("AA26") = wall
        Range("J27") = wall

        'Fireflies
        Range("I9") = firefly
        Range("P8") = firefly
        Range("Y17") = firefly
        Range("L20") = firefly
        Range("R25") = firefly

        'Shrubs
        Range("J9") = shrub
        Range("I10") = shrub
        Range("X25:Y25") = shrub
        Range("V26") = shrub
        Range("AA27") = shrub

        'Rocks
        Range("K15") = rock
        Range("H16") = rock
        Range("L16") = rock
        Range("J18") = rock
        Range("W9") = rock
        Range("U11") = rock
        Range("Y11") = rock
        Range("X14") = rock

        'Puddles
        Range("Q9") = puddle
        Range("V9") = puddle
        Range("V12") = puddle
        Range("Y13") = puddle
        Range("I21") = puddle
        Range("L24") = puddle
        Range("K25") = puddle
        Range("I26") = puddle

        'Mushrooms
        Range("W18") = mushroom
        Range("U19") = mushroom
        Range("X19") = mushroom
        Range("V20") = mushroom
        Range("O21") = mushroom
        Range("Y21") = mushroom
        Range("L22") = mushroom
        Range("P22") = mushroom
        Range("O24") = mushroom
        Range("M26") = mushroom

        'Flower
        Range("Z9") = flower
        Range("O16") = flower
        Range("W19") = flower
        Range("M21") = flower
        Range("V23") = flower

        'Traps
        Range("P11:Q11") = trap
        Range("O12:P12") = trap
        Range("R12") = trap
        Range("I13") = trap
        Range("T16") = trap
        Range("V19") = trap
        Range("U20") = trap
        Range("N21") = trap
        Range("X21") = trap
        Range("O22") = trap
        Range("M24") = trap
        Range("R24") = trap
        Range("H27") = trap

        'Shop
        Range("S10") = shop

        Range("Y28:AA28") = gate
        Range("Y29:AA29") = escape


        ShowVis
        ShowPlayer
        ShowEnemy
    End If

    'Boss Level!
    If level = 3 Then
        Range("A1:AR37").Interior.Color = vbBlack
        'Player pos
        r(0) = 33: c(0) = 38     'AJ35

        'Enemies (6)
        le_r(0) = 24: le_c(0) = 20     'T24
        le_isRevealed = False
        le_isDestroyed = False
        'Y29, Q33, H35, Z34, AM35

        'Shop
        Range("AC25") = shop
        Range("W7:X7") = gate
        Range("W6:X6") = escape

        Range("G7:V7") = wall   'top wall
        Range("Y7:AN7") = wall
        Range("G8:G36") = wall  'left wall
        Range("AN8:AN36") = wall 'right wall
        Range("H36:AN36") = wall 'bottom wall
        Range("AI30:AM30") = wall
        Range("AI27:AI29") = wall
        Range("AF27:AH27") = wall
        Range("AF23:Af26") = wall
        Range("Y23:AE23") = wall
        Range("O23:V23") = wall
        Range("O24:O27") = wall
        Range("L27:N27") = wall
        Range("L28:L30") = wall
        Range("H30:K30") = wall
        Range("T20:AA20") = wall
        Range("S9:S20") = wall
        Range("T9:AA9") = wall
        Range("AB9:AB20") = wall







        'Traps
        Range("P25:R25") = trap
        Range("Y25") = trap
        Range("AB25:AB27") = trap
        Range("AC27:AC29") = trap
        Range("AD27") = trap
        Range("P26") = trap
        Range("T26:T28") = trap
        Range("U26") = trap
        Range("X26") = trap
        Range("Y27:Z27") = trap
        Range("Q28") = trap
        Range("Y28") = trap
        Range("O29:Q29") = trap
        Range("M30:O30") = trap
        Range("M31:M34") = trap
        Range("K34:L34") = trap
        Range("I32:I33") = trap
        Range("J33") = trap
        Range("K32") = trap
        Range("I35") = trap
        Range("R28") = trap
        Range("U29") = trap
        Range("Z29:Z30") = trap
        Range("AA30") = trap
        Range("AB31") = trap
        Range("AC32") = trap
        Range("AD31") = trap
        Range("AE29:AE30") = trap
        Range("AF31:AG31") = trap
        Range("AH29:AH30") = trap
        Range("W30") = trap
        Range("X31") = trap
        Range("V30:V33") = trap
        Range("R31") = trap
        Range("S30") = trap
        Range("T31") = trap
        Range("U32") = trap
        Range("S33:Z33") = trap
        Range("AA34:AB34") = trap
        Range("Q34:S34") = trap
        Range("Q35") = trap
        Range("AK31:AM31") = trap
        Range("AM32") = trap
        Range("AI32") = trap
        Range("AF33:AJ33") = trap
        Range("AE34") = trap
        Range("AD35") = trap
        Range("AL35") = trap

        'Mushrooms
        Range("P24") = mushroom
        Range("U25") = mushroom
        Range("AE24") = mushroom
        Range("R35") = mushroom
        Range("Z32") = mushroom

        'Rocks
        Range("Q26") = rock

        'Puddles
        Range("V26") = puddle
        Range("T34") = puddle
        Range("U34") = puddle

        'Shrubs
        Range("P32") = shrub
        Range("V34") = shrub
        Range("AK33") = shrub

        'Fireflies
        Range("M29") = firefly
        Range("K33") = firefly
        Range("W31") = firefly
        Range("AG30") = firefly

        Range("U10") = mothman

        ShowVis
        ShowPlayer
        ShowEnemy
    End If


End Function











