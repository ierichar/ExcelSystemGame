Imports System.Net.Mime.MediaTypeNames

Public rinc As Integer, cinc As Integer
Public vis As Integer
Public health As Integer
Public steps As Integer
Public level As Integer

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
Public usb As Integer

Public rockSearch As Boolean
Public shrubSearch As Boolean
Public flowerSearch As Boolean
Public mushroomSearch As Boolean
Public fireflySearch As Boolean
Public puddleSearch As Boolean
Public isHalfway As Boolean

Public lightData As Integer

Dim r() As Integer, c() As Integer
Dim le_r() As Integer, le_c() As Integer
Sub StartGame()

    'Clear All Values
    Cells.Clear

    'Set Bound of Level
    Range("E5:AN32").Interior.Color = vbBlack
    Range("E5:AN32").Font.Size = 18

    'Sets Values for enviornment

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
    usb = 12

    ' Changing states
    gate = 13
    escape = 14

    Range("AA32").Value = escape
    Range("AA18").Value = trap
    Range("AA20").Value = battery
    Range("AA20").Font.ColorIndex = 6
    Range("AA26:AC26").Value = gate
    Range("Z26:Z32").Value = wall
    Range("AD26:AD32").Value = wall
    Range("AB28").Value = trap
    Range("J8").Value = rock
    Range("R15").Value = puddle
    Range("K18").Value = mushroom
    Range("W11").Value = shrub
    Range("AE24").Value = flower
    Range("AG8").Value = shop
    Range("N27").Value = firefly
    Range("I10").Value = usb
    Range("I10").Font.Color = vbGreen
    Range("A1:AR4").Value = wall
    Range("AO5:AR36").Value = wall
    Range("A33:AN36").Value = wall
    Range("A5:D32").Value = wall
    Range("A1:AR4").Interior.Color = vbBlack
    Range("AO5:AR36").Interior.Color = vbBlack
    Range("A33:AN36").Interior.Color = vbBlack
    Range("A5:D32").Interior.Color = vbBlack
    Range("AW3").Font.Size = 26
    Range("AW11").Font.Size = 26
    Range("BB11").Font.Size = 15
    Range("AW11").Value = "Light Data "
    Range("AW10").Font.Size = 26
    Range("AW10").Value = "Health"

    'sets level back to 0
    level = 0

    'loads in the level 1 values
    LoadLevel

    'Player Variables
    ReDim r(1)
    ReDim c(1)
    r(0) = 10
    c(0) = 10
    rinc = 0 : cinc = 0
    vis = 0
    lightData = 0

    'Envir searching variables
    isPickedUp = False
    rockSearch = False
    flowerSearch = False
    shrubSearch = False
    fireflySearch = False
    puddleSearch = False
    mushroomSearch = False
    isHalfway = False

    'Enemy Values
    ReDim le_r(1)
    ReDim le_c(1)
    le_r(0) = 16 : le_c(0) = 16
    le_rinc = 0 : le_cinc = 0
    le_isRevealed = False

    'bind keys and render player
    bindKeys
    ShowVis
    ShowPlayer
    AddUI

End Sub

'----------------------------SHOW PLAYER AND VISION-------------------------------------------------
Sub ShowPlayer()
    Cells(r(0), c(0)).Interior.Color = vbRed
End Sub
Sub ShowEnemy()
    Cells(le_r(0), le_c(0)).Interior.Color = vbGreen
End Sub

'=================================ShowVis=====================================
Sub ShowVis()
    Range(Cells(r(0) - vis, c(0) - vis), Cells(r(0) + vis, c(0) + vis)).Interior.ColorIndex = 15
    ShowEnemy
End Sub

'--------------------------UPDATE AND MOVE PLAYER----------------------------------------------------
Sub MovePlayer()
    ' if the player moves then run this
    If rinc <> 0 Or cinc <> 0 Then
        'sets past position to vision color
        Cells(r(0), c(0)).Interior.ColorIndex = 15
        'if the cell you are moving to is black or the vision color then run
        If (Cells(r(0) + rinc, c(0) + cinc).Interior.Color = vbBlack Or Cells(r(0) + rinc, c(0) + cinc).Interior.ColorIndex = 15) Then
            'collision  condition
            If (Cells(r(0) + rinc, c(0) + cinc).Value >= 9 Or Cells(r(0) + rinc, c(0) + cinc).Value = 0) Then
                ' permission condition
                If (Cells(r(0) + rinc, c(0) + cinc).Value <> gate Or isPickedUp = True) Then
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
        Hit
        Recharge
        ShowVis
        ShowPlayer
        UpdateUI
        
        'TODO: FUNCTION FOR THIS checks light data and informs player maybe change to UI boolean check to permission levels
        If lightData = 30 And isHalfway = False Then
            MsgBox "The USB device in your possesion whirls. A bar on the face of the device is half way full"
        isHalfway = True
        End If
        If lightData = 50 And isPickedUp = False Then
            MsgBox "The USB device in your possesion whirls again. It flashes with the words PERMISSION INCREASED"
        isPickedUp = True
        End If

    End If
End Sub

'==================RevealEnemy===============
'Pre: r(0), c(0), le_r(0), le_r(0), vis
Function RevealEnemy()
    Debug.Print("Checking reveal...")
    If Abs(r(0) - le_r(0)) <= vis And Abs(c(0) - le_r(0)) <= vis Then
        le_isRevealed = True
        Debug.Print("enemy found player!")
        RevealEnemy = True
    Else : le_isRevealed = False
    End If
End Function

'==================MoveEnemy=================
'Pre: r(0), c(0), le_r(0), le_r(0)
Sub MoveEnemy()
    Debug.Print("Moving enemy...")
    If (RevealEnemy() = True) Then
        Dim xDiff As Integer, yDiff As Integer

        yDiff = r(0) - le_r(0)
        xDiff = c(0) - le_c(0)
        Debug.Print("xDiff val: " & xDiff)
        Debug.Print("yDiff val: " & yDiff)

        If (yDiff >= 0 And xDiff >= 0) Then
            If (yDiff > xDiff) Then
                le_r(0) = le_r(0) + 1
            Else : le_c(0) = le_c(0) + 1
            End If
        ElseIf (yDiff >= 0 And xDiff <= 0) Then
            If (Abs(yDiff) > Abs(xDiff)) Then
                le_r(0) = le_r(0) + 1
            Else : le_c(0) = le_c(0) - 1
            End If
        ElseIf (yDiff <= 0 And xDiff >= 0) Then
            If (Abs(yDiff) > Abs(xDiff)) Then
                le_r(0) = le_r(0) - 1
            Else : le_c(0) = le_c(0) + 1
            End If
        ElseIf (yDiff <= 0 And xDiff <= 0) Then
            If (yDiff < xDiff) Then
                le_r(0) = le_r(0) - 1
            Else : le_c(0) = le_c(0) - 1
            End If
        End If
        If (CheckCollision(le_r(0), le_c(0)) = True) Then
            health = 3
        End If
    End If

End Sub

'=====================CheckCollision===================
Function CheckCollision(x1 As Integer, y1 As Integer) As Boolean
    'check cell of direction vector
    If (Cells(x1, y1) = Cells(c(0), r(0))) Then
        CheckCollision = True
    Else : CheckCollision = False
    End If
End Function

'------------------------------INTERACTION CHECKS----------------------------------------------
Sub interact()
    'TODO Make function for each interaction check to make it look prettier

    If Cells(r(0), c(0) - 1).Value = rock Or Cells(r(0), c(0) + 1).Value = rock Or Cells(r(0) + 1, c(0)).Value = rock Or Cells(r(0) - 1, c(0)).Value = rock Then
        If rockSearch = True Then
            Range("AW3").Value = "A plain rock"
        End If
        If rockSearch = False Then
            Range("AW3").Value = "You find a rock. It looks like its shimmering. You reach out to it and feel an energy transferred to you."
            lightData = lightData + 10
            Range("BB15").Value = lightData
            rockSearch = True
        End If

    End If
    If Cells(r(0), c(0) - 1).Value = shrub Or Cells(r(0), c(0) + 1).Value = shrub Or Cells(r(0) + 1, c(0)).Value = shrub Or Cells(r(0) - 1, c(0)).Value = shrub Then
        If shrubSearch = True Then
            Range("AW3").Value = "A shrub. The berries look dull"
        End If
        If shrubSearch = False Then
            Range("AW3").Value = "You find a shrub. The colorful berries shine bright giving off a glow of energy"
            lightData = lightData + 10
            Range("BB15").Value = lightData
            shrubSearch = True
        End If
    End If
    If Cells(r(0), c(0) - 1).Value = flower Or Cells(r(0), c(0) + 1).Value = flower Or Cells(r(0) + 1, c(0)).Value = flower Or Cells(r(0) - 1, c(0)).Value = flower Then
        If flowerSearch = True Then
            Range("AW3").Value = "The same old flower"
        End If
        If flowerSearch = False Then
            Range("AW3").Value = "You find a flower. You lean down to sniff it and feel a burst of energy within you"
            lightData = lightData + 10
            Range("BB15").Value = lightData
            flowerSearch = True
        End If

    End If
    If Cells(r(0), c(0) - 1).Value = shop Or Cells(r(0), c(0) + 1).Value = shop Or Cells(r(0) + 1, c(0)).Value = shop Or Cells(r(0) - 1, c(0)).Value = shop Then
        Range("AW3").Value = "You find a shop but there is a painted sign that says OuT fOr LUnCh"
    End If
    If Cells(r(0), c(0) - 1).Value = firefly Or Cells(r(0), c(0) + 1).Value = firefly Or Cells(r(0) + 1, c(0)).Value = firefly Or Cells(r(0) - 1, c(0)).Value = firefly Then
        If fireflySearch = True Then
            Range("AW3").Value = "The same fireflies as before but more dull"
        End If
        If fireflySearch = False Then
            Range("AW3").Value = "You find a couple of fireflies that circle around you emiting some sort of energy"
            lightData = lightData + 10
            Range("BB15").Value = lightData
            fireflySearch = True
        End If
    End If
    If Cells(r(0), c(0) - 1).Value = puddle Or Cells(r(0), c(0) + 1).Value = puddle Or Cells(r(0) + 1, c(0)).Value = puddle Or Cells(r(0) - 1, c(0)).Value = puddle Then
        If puddleSearch = True Then
            Range("AW3").Value = "The same puddle as before but only your reflection stares back at you."
        End If
        If puddleSearch = False Then
            Range("AW3").Value = "You find a puddle. Instead of your reflection it gives off an aura of energy"
            lightData = lightData + 10
            Range("BB11").Value = lightData
            puddleSearch = True
        End If

    End If
    If Cells(r(0), c(0) - 1).Value = mushroom Or Cells(r(0), c(0) + 1).Value = mushroom Or Cells(r(0) + 1, c(0)).Value = mushroom Or Cells(r(0) - 1, c(0)).Value = mushroom Then
        If mushroomSearch = True Then
            Range("AW3").Value = "The same mushroom but now much darker tones fill the spots on the cap"
        End If
        If mushroomSearch = False Then
            Range("AW3").Value = "You find a mushroom. The spots on the cap seem to be glowing and give off energy"
            lightData = lightData + 10
            Range("BB15").Value = lightData
            mushroomSearch = True
        End If
    End If
    If Cells(r(0), c(0) - 1).Value = gate Or Cells(r(0), c(0) + 1).Value = gate Or Cells(r(0) + 1, c(0)).Value = gate Or Cells(r(0) - 1, c(0)).Value = gate Then
        If isPickedUp = False Then
            Range("AW3").Value = "You find what seems like a gate. It has the same glow as the things around you but does not give it off."
        End If
        If isPickedUp = True Then
            Range("AW3").Value = "The gate has lost its glow and is swung open"
        End If
    End If
    If Cells(r(0), c(0) - 1).Value = wall Or Cells(r(0), c(0) + 1).Value = wall Or Cells(r(0) + 1, c(0)).Value = wall Or Cells(r(0) - 1, c(0)).Value = wall Then
        Range("AW3").Value = "A hard sturdy wall. Looks impenetrable"

    End If
End Sub

'--------------------------------PLAYER ON VALUE IN CELL ----------------------------------------
'TODO combine RECHARGE and HIT function
Sub Hit()
    If Cells(r(0), c(0)).Value = trap And vis > 0 Then
        vis = vis - 1
        Cells(r(0), c(0)).Value = Null
        Range("AW3").Value = "YOU STEPPED ON A TRAP: Vision level decreased"
    End If

    If Cells(r(0), c(0)).Value = escape Then
        Range("AW3").Value = "This seems like the way out! Next Level Reached"
    End If
End Sub

'===========================Recharge==============================
Sub Recharge()
    If Cells(r(0), c(0)).Value = battery And vis <= 3 Then
        vis = vis + 1
        Cells(r(0), c(0)).Value = Null
        Range("AW3").Value = "BATTERY RECOVERED: 1 vision level restored"
    End If
    If Cells(r(0), c(0)).Value = 12 And vis = 0 Then
        vis = vis + 2
        Cells(r(0), c(0)).Value = Null
        MsgBox "USB FOUND: Vision capabilites unlocked"
End If

End Sub
'------------------------------------------UI ADDING AND UPDATING---------------------------------
Sub AddUI()
    ' Health bar
    Range("AT4", "bd34").Interior.Color = RGB(239, 222, 205)
    Range("AY4").Value = "Health:"
    ' Batteries
    Range("AY6").Value = "Light Strength:"
    Range("AY11").Interior.Color = vbRed
    Range("AZ11").Value = 1
    Range("AZ10").Value = 2
    Range("AZ9").Value = 3
    Range("AZ8").Value = 4
    Range("AZ11", "AZ7").Font.Size = 18
    ' Item inventory Area
    Range("AU28", "BC33").Interior.Color = RGB(245, 245, 220)
    ' Currency
    Range("AY13").Value = "Bits:"
    Range("AY15").Value = "Light Data: "
    ' Font Size and Center Alignment
    Range("AT4", "bd34").HorizontalAlignment = xlCenter
    Range("AT4", "bd34").Font.Size = 18

End Sub
Sub UpdateUI()
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
'----------------------------------------------LOAIDNG LEVELS-----------------------------------------
Sub LoadLevel()
    If level = 0 Then
        Range("AA32").Value = escape
        Range("AA18").Value = trap
        Range("AA20").Value = battery
        Range("AA20").Font.ColorIndex = 6
        Range("AA26:AC26").Value = gate
        Range("Z26:Z32").Value = wall
        Range("AD26:AD32").Value = wall
        Range("AB28").Value = trap
        Range("J8").Value = rock
        Range("R15").Value = puddle
        Range("K18").Value = mushroom
        Range("W11").Value = shrub
        Range("AE24").Value = flower
        Range("AG8").Value = shop
        Range("N27").Value = firefly
        Range("I10").Value = usb
        Range("I10").Font.Color = vbGreen
        Range("A1:AR4").Value = wall
        Range("AO5:AR36").Value = wall
        Range("A33:AN36").Value = wall
        Range("A5:D32").Value = wall
        Range("A1:AR4").Interior.Color = vbBlack
        Range("AO5:AR36").Interior.Color = vbBlack
        Range("A33:AN36").Interior.Color = vbBlack
        Range("A5:D32").Interior.Color = vbBlack
        Range("AW3").Font.Size = 26
        Range("BB15").Font.Size = 15
    End If

End Sub

