Attribute VB_Name = "KeyControls"
Public Up As Boolean
Public Down As Boolean
Public Left As Boolean
Public Right As Boolean

Public Sub KeyDown(KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyUp
            Up = True
        Case vbKeyDown
            Down = True
        Case vbKeyLeft
            Left = True
        Case vbKeyRight
            Right = True
        Case vbKeySpace
            Form1.Restart
    End Select
End Sub

Public Sub KeyUp(KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyUp
            Up = False
        Case vbKeyDown
            Down = False
        Case vbKeyLeft
            Left = False
        Case vbKeyRight
            Right = False
    End Select
End Sub

Public Sub CheckKeys()
    If Left = True Then Ang = Ang + 2 * RadInv
    If Right = True Then Ang = Ang - 2 * RadInv
    If Up = True Then Vel = Vel + 1 Else Vel = Vel * SlowDown
    If Down = True Then Vel = Vel - 1
    
    CheckVelocities
End Sub

Public Sub CheckVelocities()
    If Vel > MaxSpeed Then Vel = MaxSpeed
    If Vel < MinSpeed Then Vel = MinSpeed
End Sub

