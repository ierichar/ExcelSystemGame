VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public potionBought As Boolean


Private Sub CommandButton1_Click()
    ' This Button Buys the battery
    If lightData >= 40 Then
        If Range("AZ8").Value <> 4 Then
            If MsgBox("Are you sure?", vbYesNo) = vbNo Then
                'If bits = 1 Then
                'MsgBox ("one battery bought")
                'End If
            'Exit Sub
            Else
                MsgBox ("Shopkeeper: Hehe thank yee")
                Range("BB15").Value = lightData
                lightData = lightData - 40
                Range("AZ8").Value = 4
            End If
        Else
            MsgBox ("Shopkeeper: I dunt got anymur! Ye bot the lust one!")
        End If
    ElseIf lightData < 40 Then
        MsgBox ("Shopkeeper: Thas noot enoof leetdata yanno!")
    End If
End Sub

Private Sub CommandButton2_Click()
' This Button Buys the potion
    'If
    'If MsgBox("Are you sure?", vbYesNo) = vbNo Then
     '   Else
      '      MsgBox ("One potion bought")
       ' End If
    'End If
    If lightData >= 20 Then
        If potionCount > 0 Then
            If MsgBox("Are you sure?", vbYesNo) = vbNo Then
            Else
                potionBought = True
                'Debug.Print (potionBought)
                MsgBox ("Shopkeeper: Hehe thank yee")
                lightData = lightData - 20
                Range("BB15").Value = lightData
                potionCount = potionCount - 1
                UpdateInventory (17)
            End If
        Else
            MsgBox ("Shopkeeper: I dunt got anymur! Ye bot the lust one!")
        End If
        
    ElseIf potionCount = 0 Then
        MsgBox ("Shopkeeper: I dunt got anymur! Ye bot the lust one!")
    ElseIf lightData < 20 Then
        MsgBox ("Shopkeeper: Thas noot enoof leetdata yanno!")
    End If
End Sub

Private Sub CommandButton3_Click()
' This Button Buys the placeholder item
    If MsgBox("Are you sure?", vbYesNo) = vbNo Then
    Else
        MsgBox ("One item bought")
    End If
End Sub
