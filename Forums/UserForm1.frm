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

Private Sub CommandButton1_Click()
' This Button Buys the battery
    If Range("AZ7").Value <> 5 Then
        If MsgBox("Are you sure?", vbYesNo) = vbNo Then
            'If bits = 1 Then
            'MsgBox ("one battery bought")
            'End If
        'Exit Sub
        Else
            MsgBox ("One battery bought")
            Range("AZ7").Value = 5
        End If
    Else
        MsgBox ("Max Battery Capacity!")
    End If

End Sub

Private Sub CommandButton2_Click()
' This Button Buys the potion
    If MsgBox("Are you sure?", vbYesNo) = vbNo Then
    Else
        MsgBox ("One potion bought")
    End If
End Sub

Private Sub CommandButton3_Click()
' This Button Buys the placeholder item
    If MsgBox("Are you sure?", vbYesNo) = vbNo Then
    Else
        MsgBox ("One item bought")
    End If
End Sub
